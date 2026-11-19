import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { ConversationReference, TurnContext } from 'botbuilder';

import {
  CHAT_READ_ALL_RSC_PERMISSION,
  createTeamsConversationSendAdapter,
  createTeamsMeetingSource,
  createTeamsReadinessProvider,
  createTeamsTargetedSendAdapter,
  POLL_ANSWER_VALUE_TYPE,
  POLL_ANSWER_VERB,
  TeamsConversationReferenceCache,
  TeamsGraphResolver,
  TeamsMeetingSessions,
} from '@omadia/channel-teams';
import type { TeamsSessionEvent } from '@omadia/channel-teams';
import type { TeamsConversationRefPersistence } from '@omadia/channel-teams';

import {
  ANNA,
  BERT,
  botNotInstalledError,
  CARLA,
  cardExecuteActivity,
  CHANNEL_ID,
  channelThreadId,
  editActivity,
  GROUP_CHAT_ID,
  JONAS,
  MEETING_CHAT_ID,
  MEETING_ID,
  meetingLifecycleActivity,
  meetingParticipantActivity,
  membersAddedActivity,
  messageActivity,
  MIRA,
  OTHER_GROUP_CHAT_ID,
  SimulatedConnector,
  SimulatedGraph,
  SimulatedKernel,
  simulatedBot,
  TEAM_AAD_GROUP_ID,
  TEMPLATE_RSC,
  TENANT_ID,
  type SimBot,
  type SimUser,
} from './support/teamsSimulation.js';

/** `TurnOrigin.channelConversationId` (channel-sdk 0.4.0): the conversation
 *  id as the channel names it. */
function channelConversationIdOf(origin: unknown): string | undefined {
  const value = (origin as { channelConversationId?: unknown } | undefined)?.channelConversationId;
  return typeof value === 'string' ? value : undefined;
}

/**
 * The contract scenarios of byte5ai/omadia#2043 §6.2 and issues #43 / #44,
 * end to end through the real bot, adapters and providers.
 *
 * THERE IS NO TEAMS TEST TENANT. Bot Framework, the Bot Connector and Graph
 * are simulated with payloads in the documented shape (see
 * tests/support/teamsSimulation.ts). These tests prove what the plugin does
 * with such payloads. Whether Teams behaves that way is checked by hand:
 * docs/facilitator-live-verification.md has one section per scenario below.
 */

/** Conversation references as the kernel's Postgres table keeps them. */
class SimulatedRefTable implements TeamsConversationRefPersistence {
  private readonly rows = new Map<string, { ref: Partial<ConversationReference>; teamsType?: string; botAppId?: string }>();

  async save(conversationId: string, ref: Partial<ConversationReference>, teamsType?: string, botAppId?: string): Promise<void> {
    this.rows.set(`${botAppId ?? ''}|${conversationId}`, {
      ref,
      ...(teamsType ? { teamsType } : {}),
      ...(botAppId ? { botAppId } : {}),
    });
  }

  async load(conversationId: string, botAppId?: string): Promise<{ ref: Partial<ConversationReference>; teamsType?: string; botAppId?: string } | undefined> {
    return this.rows.get(`${botAppId ?? ''}|${conversationId}`);
  }
}

interface Instance {
  kernel: SimulatedKernel;
  connector: SimulatedConnector;
  graph: SimulatedGraph;
  refs: TeamsConversationReferenceCache;
  membership: Record<string, unknown>[];
  sessionEvents: TeamsSessionEvent[];
  bot: ReturnType<typeof simulatedBot>;
  post: ReturnType<typeof createTeamsConversationSendAdapter>['sendToConversation'];
  dm: ReturnType<typeof createTeamsTargetedSendAdapter>['sendToUser'];
  readiness: ReturnType<typeof createTeamsReadinessProvider>;
  meetings: ReturnType<typeof createTeamsMeetingSource>;
}

const BOTS: SimBot[] = [MIRA, JONAS];
const ROSTER = [ANNA, BERT, CARLA].map((user) => ({
  channelUserId: user.id,
  aadObjectId: user.aadObjectId,
  displayName: user.name,
  email: null,
  userPrincipalName: null,
}));

/** One running middleware. Pass the previous instance's `kernel`, `graph`,
 *  `connector` and `table` to model a restart: Teams and the database stay,
 *  every in-process cache is new. */
function startInstance(
  carry: Partial<Pick<Instance, 'kernel' | 'graph' | 'connector'>> & { table?: SimulatedRefTable; ownIdentity?: boolean } = {},
): Instance & { table: SimulatedRefTable } {
  const kernel = carry.kernel ?? new SimulatedKernel();
  const graph = carry.graph ?? new SimulatedGraph();
  const connector = carry.connector ?? new SimulatedConnector();
  const table = carry.table ?? new SimulatedRefTable();
  const refs = new TeamsConversationReferenceCache();
  refs.attachPersistence(table);
  const membership: Record<string, unknown>[] = [];
  const sessionEvents: TeamsSessionEvent[] = [];
  const meetingSessions = new TeamsMeetingSessions({
    emit: (event) => {
      sessionEvents.push(event);
    },
  });
  const bot = simulatedBot(kernel, {
    captureConversationReference: (context) => refs.capture(context),
    emitMembershipEvent: (event) => {
      membership.push(event);
    },
    meetingSessions,
    ...(carry.ownIdentity ? { ownIdentity: true } : {}),
  });
  const postAdapter = createTeamsConversationSendAdapter({
    refs,
    sendProactive: connector.sendProactive,
    roster: { list: async () => ROSTER },
  });
  const dmAdapter = createTeamsTargetedSendAdapter({
    sendProactive: connector.sendProactive,
    direct: {
      refs,
      openDirectConversation: connector.openDirectConversation,
      botLabel: (botAppId) => BOTS.find((candidate) => candidate.appId === botAppId)?.name,
    },
  });
  const readiness = createTeamsReadinessProvider({
    botLabel: (botAppId) => BOTS.find((candidate) => candidate.appId === botAppId)?.name,
    // The instance's registration, not a facilitator's.
    graph: new TeamsGraphResolver({ tenantId: TENANT_ID, clientId: 'instance-app', clientSecret: 'simulated', log: () => {}, fetchImpl: graph.fetch }),
    refs,
  });
  return {
    kernel,
    connector,
    graph,
    refs,
    table,
    membership,
    sessionEvents,
    bot,
    post: postAdapter.sendToConversation.bind(postAdapter),
    dm: dmAdapter.sendToUser.bind(dmAdapter),
    readiness,
    meetings: createTeamsMeetingSource({
      refs,
      sessions: meetingSessions,
      grantedPermissions: (conversationId, asChannelKey) => readiness.grantedPermissions(conversationId, { asChannelKey }),
    }),
  };
}

function dmTarget(user: SimUser): Parameters<Instance['dm']>[0] {
  return { principalId: `user:${user.aadObjectId}`, userRef: { kind: 'teams-aad', id: user.aadObjectId } };
}

/** Let fire-and-forget persistence writes settle. */
async function settle(): Promise<void> {
  await new Promise((resolve) => setImmediate(resolve));
}

describe('scenario 0 — the mention that opens a facilitation', () => {
  // Nobody is subscribed yet: the mention is an ordinary turn, and the kernel
  // decides from its origin whether a facilitator may act (group) or has to
  // refuse (private request). 0.36.0 to 0.36.2 sent no `conversationType`,
  // so the kernel refused every such mention.
  const isGroupForKernel = (origin: unknown): boolean => {
    const o = origin as { scope?: { kind?: string }; conversationType?: string } | undefined;
    if (!o) return false;
    if (o.scope?.kind === 'group') return true;
    return o.scope?.kind === 'conversation' && o.conversationType === 'group';
  };

  it('a mention in a group chat, a team channel and a meeting chat reaches the turn with a group origin', async () => {
    const activities = [
      messageActivity({ bot: MIRA, from: ANNA, text: 'starte eine Abstimmung', mention: true }),
      messageActivity({ bot: MIRA, from: ANNA, text: 'starte eine Abstimmung', mention: true, conversationId: channelThreadId('1795023950656'), conversationType: 'channel' }),
      messageActivity({ bot: MIRA, from: ANNA, text: 'starte eine Abstimmung', mention: true, conversationId: MEETING_CHAT_ID }),
    ];
    for (const activity of activities) {
      const i = startInstance();

      await i.bot.deliver(activity);

      assert.equal(i.bot.turns.length, 1, 'the mention is a turn');
      const origin = i.bot.origins[0] as { channelType?: string; conversationType?: string; scope?: { kind?: string } };
      assert.equal(origin.channelType, 'teams');
      assert.equal(origin.scope?.kind, 'conversation');
      assert.equal(origin.conversationType, 'group', String(activity.conversation?.id));
      assert.equal(isGroupForKernel(origin), true, 'the kernel lets a facilitator act here');
    }
  });

  it('the turn names its sender with Entra object id and display name', async () => {
    // The kernel builds `callerUserRef` of the tool call context from this, so
    // the facilitator can name who started a facilitation. A card click that
    // continues detached builds its origin from the original activity with
    // the same function (`teamsTurnOrigin`); that path needs the proactive
    // adapter and is not run here.
    for (const conversationId of [GROUP_CHAT_ID, channelThreadId('1795023950656'), MEETING_CHAT_ID]) {
      const i = startInstance({ ownIdentity: true });

      await i.bot.deliver(
        messageActivity({
          bot: MIRA,
          from: ANNA,
          text: 'starte eine Abstimmung',
          mention: true,
          conversationId,
          ...(conversationId.includes('@thread.tacv2') ? { conversationType: 'channel' as const } : {}),
        }),
      );

      assert.deepEqual(
        i.bot.origins.map((origin) => (origin as { speaker?: unknown }).speaker),
        [{ kind: 'teams-aad', id: ANNA.aadObjectId, displayName: ANNA.name }],
        conversationId,
      );
    }
  });

  it('a message in the 1:1 with the bot reaches the turn as a private request', async () => {
    const i = startInstance();

    await i.bot.deliver(
      messageActivity({ bot: MIRA, from: ANNA, text: 'hilf mir', conversationId: 'a:1dRsHCobZ1AxURzY05Dc_personal', conversationType: 'personal' }),
    );

    assert.equal(i.bot.turns.length, 1);
    const origin = i.bot.origins[0] as { conversationType?: string; scope?: { kind?: string } };
    assert.equal(origin.scope?.kind, 'personal');
    assert.equal(origin.conversationType, 'direct');
    assert.equal(isGroupForKernel(origin), false, 'a facilitator refuses a private request');
  });
});

describe('scenario 10 — a facilitation in a channel thread, facilitator with its own bot', () => {
  // Ids as captured live on 2026-11-18. A root post in a channel is its own
  // conversation (`<channel>;messageid=<root>`); the facilitation belongs to
  // the thread it was started in. The kernel and the engine get that id from
  // the turn origin and hand it back unchanged.
  const ROOT = '1795023950656';
  const THREAD = channelThreadId(ROOT);
  const TEAM_RSC = TEMPLATE_RSC.filter((permission) => permission.endsWith('.Group'));

  function threadMessage(from: SimUser, text: string, opts: { mention?: boolean; id?: string; root?: string } = {}): ReturnType<typeof messageActivity> {
    return messageActivity({
      bot: MIRA,
      from,
      text,
      conversationId: channelThreadId(opts.root ?? ROOT),
      conversationType: 'channel',
      ...(opts.mention ? { mention: true } : {}),
      ...(opts.id ? { id: opts.id } : {}),
    });
  }

  it('the opening mention states the thread id; the session scope is bot-qualified and is not that id', async () => {
    const i = startInstance({ ownIdentity: true });

    await i.bot.deliver(threadMessage(ANNA, 'starte eine Abstimmung', { mention: true, id: ROOT }));

    assert.deepEqual(i.bot.scopes, [`teams-${MIRA.key}-${THREAD}`]);
    const stated = channelConversationIdOf(i.bot.origins[0]);
    assert.equal(stated, THREAD);

    // The kernel subscribes with the id as given; the message events of the
    // same conversation carry the same string.
    i.kernel.subscribe(stated ?? '', MIRA);
    await i.bot.deliver({ ...threadMessage(BERT, 'B > A > C'), replyToId: ROOT });
    assert.deepEqual(i.kernel.events.map((e) => e.conversationId), [stated]);
  });

  it('readiness for the thread asks Graph for the TEAM, never with the thread id or a bot key', async () => {
    const i = startInstance({ ownIdentity: true });
    i.graph.install(TEAM_AAD_GROUP_ID, MIRA, { consented: TEAM_RSC });
    await i.bot.deliver(threadMessage(ANNA, 'starte eine Abstimmung', { mention: true, id: ROOT }));

    const report = await i.readiness.check(THREAD, { asChannelKey: MIRA.key });

    assert.deepEqual(report, { installed: true, canReadAllMessages: true, canPost: true, ready: true, missing: [] });
    const graphCalls = i.graph.requests.filter((r) => r.includes('graph.microsoft.com'));
    assert.ok(graphCalls.length >= 2);
    for (const call of graphCalls) {
      assert.ok(call.includes(`/teams/${TEAM_AAD_GROUP_ID}/`), call);
      assert.ok(!/messageid|28%3A|28:|thread\.tacv2/i.test(call), `not for Graph: ${call}`);
    }
  });

  it('a reply in the thread is recorded under the thread id; a new root post in the channel is not covered', async () => {
    const i = startInstance({ ownIdentity: true });
    await i.bot.deliver(threadMessage(ANNA, 'starte eine Abstimmung', { mention: true, id: ROOT }));
    i.kernel.subscribe(THREAD, MIRA);

    // Documented, not captured live: a reply carries the thread's id and
    // `replyToId` = the root.
    await i.bot.deliver({ ...threadMessage(BERT, 'B > A > C'), replyToId: ROOT });
    await i.bot.deliver({ ...threadMessage(CARLA, 'C > A > B'), replyToId: ROOT });
    // Another root post: another conversation.
    await i.bot.deliver(threadMessage(BERT, 'ganz anderes Thema', { root: '1795024337919', id: '1795024337919' }));

    assert.deepEqual(
      i.kernel.events.map((e) => [e.conversationId, e.author.displayName, e.text]),
      [
        [THREAD, BERT.name, 'B > A > C'],
        [THREAD, CARLA.name, 'C > A > B'],
      ],
    );
    assert.equal(i.bot.turns.length, 1, 'only the opening mention was a turn');
  });

  it('the facilitator posts INTO the thread, also into one it holds no reference of its own for', async () => {
    const i = startInstance({ ownIdentity: true });
    await i.bot.deliver(threadMessage(ANNA, 'starte eine Abstimmung', { mention: true, id: ROOT }));
    await settle();

    const posted = await i.post(THREAD, { text: 'Die Abstimmung läuft.' }, { asChannelKey: MIRA.key });

    assert.equal(posted.outcome, 'delivered');
    assert.equal(i.connector.sent.at(-1)?.conversationId, THREAD, 'a reply in the thread, not a new root post');
    assert.equal(i.connector.sent.at(-1)?.asBot, MIRA.key);

    // After a restart, with only the bare channel row (written when the bot
    // was added to the team) the thread is still reachable.
    const table = new SimulatedRefTable();
    const fresh = startInstance({ ownIdentity: true, table });
    await fresh.bot.deliver(
      messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true, conversationId: CHANNEL_ID, conversationType: 'channel' }),
    );
    await settle();
    const restarted = startInstance({ ownIdentity: true, table, connector: fresh.connector });
    const other = channelThreadId('1795023471899');
    const intoOther = await restarted.post(other, { text: 'Erinnerung.' }, { asChannelKey: MIRA.key });

    assert.equal(intoOther.outcome, 'delivered');
    assert.equal(restarted.connector.sent.at(-1)?.conversationId, other);
    assert.equal(restarted.refs.teamIdFor(other), TEAM_AAD_GROUP_ID, 'the team of every thread is the channel\'s team');
  });

  it('refuses a conversation id cut out of the session scope, by name and without asking Graph', async () => {
    const i = startInstance({ ownIdentity: true });
    i.graph.install(TEAM_AAD_GROUP_ID, MIRA, { consented: TEAM_RSC });
    await i.bot.deliver(threadMessage(ANNA, 'starte eine Abstimmung', { mention: true, id: ROOT }));
    // What the kernel derived before the origin stated the id (D13).
    const cutFromScope = (i.bot.scopes[0] ?? '').slice('teams-'.length);
    assert.equal(cutFromScope, `${MIRA.key}-${THREAD}`);

    const report = await i.readiness.check(cutFromScope, { asChannelKey: MIRA.key });
    const repair = await i.readiness.repairConsent(cutFromScope, { asChannelKey: MIRA.key });
    const posted = await i.post(cutFromScope, { text: 'x' }, { asChannelKey: MIRA.key });

    assert.deepEqual(report.missing.map((m) => m.code), ['unsupported_conversation']);
    assert.match(report.missing[0]?.message ?? '', /is not a Teams conversation id: it starts with a bot key/);
    assert.ok(report.missing[0]?.message.includes(cutFromScope), 'the id is named');
    assert.equal(repair.outcome === 'failed' ? repair.code : '', 'unsupported_conversation');
    assert.equal(posted.outcome, 'unreachable');
    assert.match(posted.outcome === 'unreachable' ? posted.message : '', /starts with a bot key/);
    assert.deepEqual(i.graph.requests.filter((r) => r.includes('graph.microsoft.com')), []);
    assert.deepEqual(i.connector.sent, []);
  });

  it('repairs the consent of a team install through the team, with the team-scoped permissions', async () => {
    const i = startInstance({ ownIdentity: true });
    i.graph.install(TEAM_AAD_GROUP_ID, MIRA, { consented: ['TeamsActivity.Send.Group'], version: '1.3.0' });
    await i.bot.deliver(threadMessage(ANNA, 'starte eine Abstimmung', { mention: true, id: ROOT }));
    assert.deepEqual((await i.readiness.check(THREAD, { asChannelKey: MIRA.key })).missing.map((m) => m.code), ['consent_missing']);

    const outcome = await i.readiness.repairConsent(THREAD, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'upgraded');
    assert.deepEqual(i.graph.upgrades.map((u) => u.scope), [TEAM_AAD_GROUP_ID]);
    assert.ok(i.graph.requests.some((r) => r.startsWith(`POST https://graph.microsoft.com/v1.0/teams/${TEAM_AAD_GROUP_ID}/installedApps/`) && r.endsWith('/upgrade')));
    assert.equal((await i.readiness.check(THREAD, { asChannelKey: MIRA.key })).ready, true);
  });
});

describe('scenario 11 — a facilitation in a meeting chat, facilitator with its own bot', () => {
  it('the opening mention states the chat id, and readiness asks Graph for exactly that chat', async () => {
    const i = startInstance({ ownIdentity: true });
    i.graph.install(MEETING_CHAT_ID, MIRA);

    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'starte eine Abstimmung', mention: true, conversationId: MEETING_CHAT_ID }));

    assert.deepEqual(i.bot.scopes, [`teams-${MIRA.key}-${MEETING_CHAT_ID}`]);
    const stated = channelConversationIdOf(i.bot.origins[0]);
    assert.equal(stated, MEETING_CHAT_ID);

    const report = await i.readiness.check(stated ?? '', { asChannelKey: MIRA.key });
    assert.equal(report.ready, true, JSON.stringify(report.missing));
    const chat = encodeURIComponent(MEETING_CHAT_ID);
    const graphCalls = i.graph.requests.filter((r) => r.includes('graph.microsoft.com'));
    assert.ok(graphCalls.length >= 2);
    assert.ok(graphCalls.every((call) => call.includes(`/chats/${chat}/`)), graphCalls.join('\n'));
  });

  it('records every message under the chat id and posts as the facilitator', async () => {
    const i = startInstance({ ownIdentity: true });
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'starte eine Abstimmung', mention: true, conversationId: MEETING_CHAT_ID }));
    i.kernel.subscribe(MEETING_CHAT_ID, MIRA);

    await i.bot.deliver(messageActivity({ bot: MIRA, from: BERT, text: 'B', conversationId: MEETING_CHAT_ID }));
    const posted = await i.post(MEETING_CHAT_ID, { text: 'Danke, Bert.' }, { asChannelKey: MIRA.key });

    assert.deepEqual(i.kernel.events.map((e) => [e.conversationId, e.text]), [[MEETING_CHAT_ID, 'B']]);
    assert.equal(posted.outcome, 'delivered');
    assert.equal(i.connector.sent.at(-1)?.conversationId, MEETING_CHAT_ID);
  });

  it('refuses the bot-key-prefixed id the kernel used before (the live failure of 2026-11-18)', async () => {
    const i = startInstance({ ownIdentity: true });
    i.graph.install(MEETING_CHAT_ID, MIRA);
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'starte eine Abstimmung', mention: true, conversationId: MEETING_CHAT_ID }));

    const report = await i.readiness.check(`${MIRA.key}-${MEETING_CHAT_ID}`, { asChannelKey: MIRA.key });

    assert.deepEqual(report.missing.map((m) => m.code), ['unsupported_conversation'], 'one named reason, not graph_error plus cannot_post');
    assert.deepEqual(i.graph.requests.filter((r) => r.includes('graph.microsoft.com')), []);
  });
});

describe('scenario 1 — message without mention', () => {
  it('four people vote, two without mentioning the bot: every vote is recorded with its author', async () => {
    const i = startInstance();
    i.kernel.subscribe(GROUP_CHAT_ID, MIRA);

    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A > B > C', mention: true }));
    await i.bot.deliver(messageActivity({ bot: MIRA, from: BERT, text: 'B > A > C' }));
    await i.bot.deliver(messageActivity({ bot: MIRA, from: CARLA, text: 'C > A > B' }));
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'doch lieber B > A > C', mention: true }));

    assert.deepEqual(
      i.kernel.events.map((e) => [e.kind, e.author.displayName, e.text]),
      [
        ['created', ANNA.name, 'Mira A > B > C'],
        ['created', BERT.name, 'B > A > C'],
        ['created', CARLA.name, 'C > A > B'],
        ['created', ANNA.name, 'Mira doch lieber B > A > C'],
      ],
    );
    assert.deepEqual(i.bot.turns, [], 'the engine drives every turn; the adapter starts none, also not for a mention');
    assert.deepEqual(i.bot.sent, []);
  });

  it('in a chat without facilitation the mention gate still applies', async () => {
    const i = startInstance();
    i.kernel.subscribe(GROUP_CHAT_ID, MIRA);

    await i.bot.deliver(messageActivity({ bot: MIRA, from: BERT, text: 'B > A > C', conversationId: OTHER_GROUP_CHAT_ID }));

    assert.deepEqual(i.kernel.events, []);
    assert.deepEqual(i.bot.turns, []);
  });
});

describe('scenario 2 — missing consent', () => {
  it('the readiness check says exactly what is missing and the facilitation must not start', async () => {
    const i = startInstance();
    // The app was installed from a package without the RSC permission:
    // Teams then delivers mentioned messages only.
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));
    i.graph.install(GROUP_CHAT_ID, MIRA, { consented: ['TeamsActivity.Send.Chat'], version: '1.0.1' });

    const report = await i.readiness.check(GROUP_CHAT_ID, { asChannelKey: MIRA.key });

    assert.equal(report.ready, false);
    assert.deepEqual(
      { installed: report.installed, canPost: report.canPost, canReadAllMessages: report.canReadAllMessages },
      { installed: true, canPost: true, canReadAllMessages: false },
    );
    assert.deepEqual(report.missing.map((m) => m.code), ['consent_missing']);
    assert.ok(report.missing[0]?.message.includes(CHAT_READ_ALL_RSC_PERMISSION));
  });

  it('the Graph upgrade grants the permission and the same check is ready', async () => {
    const i = startInstance();
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));
    i.graph.install(GROUP_CHAT_ID, MIRA, { consented: ['TeamsActivity.Send.Chat'], version: '1.0.1' });

    const repair = await i.readiness.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });

    assert.equal(repair.outcome, 'upgraded');
    assert.equal((await i.readiness.check(GROUP_CHAT_ID, { asChannelKey: MIRA.key })).ready, true);
  });
});

describe('scenario 3 — first-contact DM', () => {
  it('a private reminder reaches a person who never wrote to the bot, from the facilitator\'s own bot', async () => {
    const i = startInstance();

    const outcome = await i.dm(dmTarget(CARLA), { text: 'Deine Stimme fehlt noch.' }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'delivered');
    assert.ok(outcome.outcome === 'delivered' && outcome.messageId);
    assert.deepEqual(i.connector.opened, [{ asBot: MIRA.key, userId: CARLA.aadObjectId }]);
    assert.deepEqual(i.connector.sent.map((s) => s.asBot), [MIRA.key]);
  });

  it('a person without the app gets no DM; the reminder goes neutrally to the group, as the same bot', async () => {
    const i = startInstance();
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));
    i.connector.failDirect(MIRA, CARLA, botNotInstalledError());

    const dm = await i.dm(dmTarget(CARLA), { text: 'Deine Stimme fehlt noch.' }, { asChannelKey: MIRA.key });
    assert.equal(dm.outcome === 'unreachable' ? dm.code : '', 'app_not_installed');

    const group = await i.post(GROUP_CHAT_ID, { text: 'Es fehlen noch Stimmen.' }, { asChannelKey: MIRA.key });
    assert.equal(group.outcome, 'delivered');
    assert.deepEqual(i.connector.sent.map((s) => [s.asBot, s.conversationId]), [[MIRA.key, GROUP_CHAT_ID]]);
  });
});

describe('scenario 4 — blocked DM', () => {
  it('a person who blocked the bot yields blocked, and no other bot writes to them', async () => {
    const i = startInstance();
    i.connector.block(MIRA, BERT);

    const outcome = await i.dm(dmTarget(BERT), { text: 'Deine Stimme fehlt noch.' }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'blocked');
    assert.deepEqual(i.connector.sent, []);
    assert.deepEqual(i.connector.opened.map((o) => o.asBot), [MIRA.key]);
  });
});

describe('scenario 5 — restart', () => {
  it('after a restart the next vote without mention is still recorded and the bot still posts as itself', async () => {
    const before = startInstance();
    before.kernel.subscribe(GROUP_CHAT_ID, MIRA);
    await before.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A' }));
    await before.bot.deliver(messageActivity({ bot: MIRA, from: BERT, text: 'B' }));
    await settle();

    const after = startInstance({ kernel: before.kernel, graph: before.graph, connector: before.connector, table: before.table });
    // No inbound activity yet in the new process: the reference comes from the table.
    const posted = await after.post(GROUP_CHAT_ID, { text: 'Zwischenstand: 2 von 4.' }, { asChannelKey: MIRA.key });
    await after.bot.deliver(messageActivity({ bot: MIRA, from: CARLA, text: 'C' }));

    assert.equal(posted.outcome, 'delivered');
    assert.equal(after.connector.sent.at(-1)?.asBot, MIRA.key);
    assert.deepEqual(after.kernel.events.map((e) => e.text), ['A', 'B', 'C']);
    assert.deepEqual(after.bot.turns, []);
  });

  it('after a restart nothing waits for a new bot_added: readiness, roster, post, DM and message events work at once', async () => {
    const before = startInstance();
    before.kernel.subscribe(GROUP_CHAT_ID, MIRA);
    before.graph.install(GROUP_CHAT_ID, MIRA);
    await before.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A' }));
    await settle();

    // Teams, Graph, the kernel's subscriptions and the reference table stay;
    // every map of the plugin is new and no membership event is delivered.
    const after = startInstance({ kernel: before.kernel, graph: before.graph, connector: before.connector, table: before.table });
    const sentBefore = after.connector.sent.length;

    const report = await after.readiness.check(GROUP_CHAT_ID, { asChannelKey: MIRA.key });
    const posted = await after.post(
      GROUP_CHAT_ID,
      { text: '@Bert, deine Stimme fehlt.', mentions: [{ user: { kind: 'teams-aad', id: BERT.aadObjectId }, text: '@Bert' }] },
      { asChannelKey: MIRA.key },
    );
    const reminded = await after.dm(dmTarget(BERT), { text: 'Deine Stimme fehlt noch.' }, { asChannelKey: MIRA.key });
    await after.bot.deliver(messageActivity({ bot: MIRA, from: BERT, text: 'B' }));

    assert.equal(report.ready, true, JSON.stringify(report.missing));
    assert.equal(posted.outcome, 'delivered');
    assert.equal(reminded.outcome, 'delivered');
    assert.ok(
      JSON.stringify(after.connector.sent.slice(sentBefore)).includes('"type":"mention"'),
      'the roster was fetched on demand and the mention resolved',
    );
    assert.ok(after.connector.sent.slice(sentBefore).every((sent) => sent.asBot === MIRA.key));
    assert.deepEqual(after.kernel.events.map((e) => e.text), ['A', 'B']);
    assert.deepEqual(after.bot.turns, []);
    assert.deepEqual(after.membership, [], 'no membership event was needed');
  });

  it('after a restart an edit still carries the message id and its platform time', async () => {
    const before = startInstance();
    before.kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const ballot = messageActivity({ bot: MIRA, from: ANNA, text: 'A' });
    const id = ballot.id as string;
    await before.bot.deliver(ballot);
    await before.bot.deliver(editActivity({ bot: MIRA, from: ANNA, id, text: 'B', timestamp: '2026-11-18T09:20:00.000Z' }));

    const after = startInstance({ kernel: before.kernel });
    await after.bot.deliver(editActivity({ bot: MIRA, from: ANNA, id, text: 'C', timestamp: '2026-11-18T09:40:00.000Z' }));

    // The kernel assigns the stored revision (highest stored + 1); the
    // adapter's number is a hint and is deliberately not asserted here. What
    // the kernel needs from the adapter is the id, the kind, the text and
    // the platform time of every edit.
    assert.deepEqual(
      after.kernel.events.map((e) => [e.kind, e.messageId, e.text, e.occurredAt]),
      [
        ['created', id, 'A', '2026-11-18T09:15:00.000Z'],
        ['edited', id, 'B', '2026-11-18T09:20:00.000Z'],
        ['edited', id, 'C', '2026-11-18T09:40:00.000Z'],
      ],
    );
    assert.ok(after.kernel.events.slice(1).every((e) => e.revision >= 1), 'an edit is never revision 0');
  });
});

describe('scenario 6 — two facilitator bots in one chat', () => {
  it('each bot records the message for itself and posts under its own identity', async () => {
    const i = startInstance();
    i.kernel.subscribe(GROUP_CHAT_ID, MIRA);
    i.kernel.subscribe(GROUP_CHAT_ID, JONAS);
    const id = '1731930009001';

    // Teams delivers one message once per bot installed in the chat.
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A', id }));
    await i.bot.deliver(messageActivity({ bot: JONAS, from: ANNA, text: 'A', id }));
    await i.post(GROUP_CHAT_ID, { text: 'Für dich notiert: A' }, { asChannelKey: MIRA.key });
    await i.post(GROUP_CHAT_ID, { text: 'Danke, Anna.' }, { asChannelKey: JONAS.key });

    assert.deepEqual(i.kernel.eventsFor(MIRA).map((e) => e.messageId), [id]);
    assert.deepEqual(i.kernel.eventsFor(JONAS).map((e) => e.messageId), [id]);
    assert.deepEqual(i.connector.sent.map((s) => [s.asBot, s.activity['text']]), [
      [MIRA.key, 'Für dich notiert: A'],
      [JONAS.key, 'Danke, Anna.'],
    ]);
    assert.deepEqual(i.bot.turns, []);
  });

  it('a facilitator and an ordinary agent bot in one chat: only the facilitator hears everything', async () => {
    const i = startInstance();
    i.kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const id = '1731930009002';

    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A', id }));
    await i.bot.deliver(messageActivity({ bot: JONAS, from: ANNA, text: 'A', id }));

    assert.deepEqual(i.kernel.events.map((e) => e.receivedByChannelKey), [MIRA.key]);
  });

  it('one facilitator in two chats keeps them apart', async () => {
    const i = startInstance();
    i.kernel.subscribe(GROUP_CHAT_ID, MIRA);
    i.kernel.subscribe(OTHER_GROUP_CHAT_ID, MIRA);

    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A' }));
    await i.bot.deliver(messageActivity({ bot: MIRA, from: BERT, text: 'X', conversationId: OTHER_GROUP_CHAT_ID }));
    await i.post(OTHER_GROUP_CHAT_ID, { text: 'nur hier' }, { asChannelKey: MIRA.key });

    assert.deepEqual(i.kernel.events.map((e) => [e.conversationId, e.text]), [
      [GROUP_CHAT_ID, 'A'],
      [OTHER_GROUP_CHAT_ID, 'X'],
    ]);
    assert.deepEqual(i.connector.sent.map((s) => s.conversationId), [OTHER_GROUP_CHAT_ID]);
  });
});

describe('scenario 7a — the snapshot says `online` only while a meeting is known to run', () => {
  // The engine reads the snapshot at set-up and takes `online` to mean "a
  // meeting is running". Rollout defect D26: a meeting chat reported
  // `online` at any time, just for being a meeting chat.
  const modalityOf = async (i: ReturnType<typeof startInstance>): Promise<string | undefined> =>
    (await i.meetings.snapshot(MEETING_CHAT_ID, { asChannelKey: MIRA.key }))?.modality;

  it('chat before the start, online during the meeting, chat after its end', async () => {
    const i = startInstance();
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'Thema: Budget', conversationId: MEETING_CHAT_ID }));

    const before = await i.meetings.snapshot(MEETING_CHAT_ID, { asChannelKey: MIRA.key });
    await i.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));
    const during = await i.meetings.snapshot(MEETING_CHAT_ID, { asChannelKey: MIRA.key });
    await i.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'end', at: '2026-11-18T11:00:00.000Z' }));
    const after = await i.meetings.snapshot(MEETING_CHAT_ID, { asChannelKey: MIRA.key });

    assert.deepEqual([before?.modality, during?.modality, after?.modality], ['chat', 'online', 'chat']);
    // What the chat offers does not change with the meeting.
    assert.deepEqual(before?.inputs, ['text', 'message_edits', 'card_answers']);
    assert.deepEqual(before?.outputs, ['text', 'mentions', 'cards', 'polls', 'private_messages']);
    assert.deepEqual(during?.outputs, before?.outputs);
    assert.equal('sessionRef' in (before ?? {}), false, 'no session while nothing runs');
    assert.equal(during?.sessionRef, MEETING_ID);
    assert.equal('sessionRef' in (after ?? {}), false);
  });

  it('after a restart during a meeting the state is unknown, and unknown is reported as chat', async () => {
    const before = startInstance();
    await before.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hi', conversationId: MEETING_CHAT_ID }));
    await before.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));
    assert.equal(await modalityOf(before), 'online');
    await settle();

    // The process restarts while the meeting runs. Microsoft documents no way
    // to ask whether a meeting is running, so nothing is claimed.
    const after = startInstance({ kernel: before.kernel, graph: before.graph, connector: before.connector, table: before.table });
    assert.equal(await modalityOf(after), 'chat');

    // The next evidence brings the state back: somebody joins.
    await after.bot.deliver(meetingParticipantActivity({ bot: MIRA, from: ANNA, phase: 'join', people: [BERT] }));
    assert.equal(await modalityOf(after), 'online');
    await after.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'end', at: '2026-11-18T11:00:00.000Z' }));
    assert.equal(await modalityOf(after), 'chat');
  });

  it('a leave event never turns a chat into a running meeting', async () => {
    const i = startInstance();
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hi', conversationId: MEETING_CHAT_ID }));
    await i.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));
    await i.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'end', at: '2026-11-18T11:00:00.000Z' }));

    // A leave that arrives after the end event (or after a restart).
    await i.bot.deliver(meetingParticipantActivity({ bot: MIRA, from: ANNA, phase: 'leave', people: [BERT], at: '2026-11-18T11:00:05.000Z' }));

    assert.equal(await modalityOf(i), 'chat');
  });

  it('an ordinary group chat is a chat, and a call started in it is online until it ends', async () => {
    const i = startInstance();
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hi' }));
    const modality = async (): Promise<string | undefined> =>
      (await i.meetings.snapshot(GROUP_CHAT_ID, { asChannelKey: MIRA.key }))?.modality;

    assert.equal(await modality(), 'chat');
    await i.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start', meetingType: 'GroupCall', conversationId: GROUP_CHAT_ID }));
    assert.equal(await modality(), 'online');
    await i.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'end', meetingType: 'GroupCall', conversationId: GROUP_CHAT_ID, at: '2026-11-18T10:30:00.000Z' }));
    assert.equal(await modality(), 'chat');
  });
});

describe('scenario 7 — scheduled meeting chat', () => {
  it('the facilitator reads the meeting chat without mentions and declares what it perceives', async () => {
    const i = startInstance();
    i.kernel.subscribe(MEETING_CHAT_ID, MIRA);

    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'Thema: Budget', conversationId: MEETING_CHAT_ID }));
    const posted = await i.post(MEETING_CHAT_ID, { text: 'Agenda: 1. Budget' }, { asChannelKey: MIRA.key });
    const snapshot = await i.meetings.snapshot(MEETING_CHAT_ID, { asChannelKey: MIRA.key, sessionRef: 'session-1' });

    assert.deepEqual(i.kernel.events.map((e) => [e.conversationId, e.text]), [[MEETING_CHAT_ID, 'Thema: Budget']]);
    assert.equal(posted.outcome, 'delivered');
    assert.equal(snapshot?.modality, 'chat', 'no meeting is running: the chat of a meeting is a chat');
    assert.equal(snapshot?.sessionRef, 'session-1');
    assert.equal((snapshot?.inputs as readonly string[]).includes('transcribed_speech'), false, 'it does not hear the call');
    assert.equal((snapshot?.inputs as readonly string[]).includes('presence'), false);
  });
});

describe('scenario 8 — poll answers', () => {
  it('a live poll is posted as a card and each click arrives with its platform-stated answerer', async () => {
    const i = startInstance();
    i.kernel.subscribe(MEETING_CHAT_ID, MIRA);
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'los', conversationId: MEETING_CHAT_ID, mention: true }));
    const before = i.kernel.events.length;

    const poll = { pollId: 'poll-budget', question: 'Budget freigeben?', options: [{ id: 'yes', label: 'Ja' }, { id: 'no', label: 'Nein' }] };
    const posted = await i.post(MEETING_CHAT_ID, { text: 'Abstimmung', poll }, { asChannelKey: MIRA.key });
    assert.ok(posted.outcome === 'delivered' && posted.messageId && posted.cardMessageId);
    assert.notEqual(posted.outcome === 'delivered' ? posted.messageId : '', posted.outcome === 'delivered' ? posted.cardMessageId : '');
    const cardMessageId = posted.outcome === 'delivered' ? (posted.cardMessageId as string) : '';
    const click = (from: SimUser, optionId: string, anonymous = false): Promise<void> =>
      i.bot.deliver(
        cardExecuteActivity({
          bot: MIRA,
          from,
          anonymous,
          conversationId: MEETING_CHAT_ID,
          cardMessageId,
          verb: POLL_ANSWER_VERB,
          data: { type: POLL_ANSWER_VALUE_TYPE, pollId: 'poll-budget', optionId },
        }),
      );

    await click(ANNA, 'yes');
    await click(BERT, 'no');
    await click(BERT, 'yes');
    await click(CARLA, 'no', true);

    assert.deepEqual(
      i.kernel.events.slice(before).map((e) => [e.kind, e.messageId, e.author.id, e.cardAnswer]),
      [
        ['card_answer', cardMessageId, ANNA.aadObjectId, { cardId: 'poll-budget', value: { optionId: 'yes' }, answererVerified: true }],
        ['card_answer', cardMessageId, BERT.aadObjectId, { cardId: 'poll-budget', value: { optionId: 'no' }, answererVerified: true }],
        ['card_answer', cardMessageId, BERT.aadObjectId, { cardId: 'poll-budget', value: { optionId: 'yes' }, answererVerified: true }],
        // An anonymous participant: no Entra object id, flagged as unverified.
        ['card_answer', cardMessageId, '8:teamsvisitor:7d3a4e5f', { cardId: 'poll-budget', value: { optionId: 'no' }, answererVerified: false }],
      ],
      'a changed answer is a second event; the engine decides which one counts',
    );
    assert.deepEqual(i.bot.turns, [], 'no channel-started turn in an engine-driven conversation');
  });
});

describe('scenario 9 — a participant joining late', () => {
  it('a person added to the chat mid-way is announced to the kernel and their vote counts as theirs', async () => {
    const i = startInstance();
    i.kernel.subscribe(MEETING_CHAT_ID, MIRA);
    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A', conversationId: MEETING_CHAT_ID }));

    await i.bot.deliver(membersAddedActivity({ bot: MIRA, from: ANNA, conversationId: MEETING_CHAT_ID, added: [CARLA] }));
    await i.bot.deliver(messageActivity({ bot: MIRA, from: CARLA, text: 'B', conversationId: MEETING_CHAT_ID }));

    assert.equal(i.membership.length, 1);
    assert.equal(i.membership[0]?.['kind'], 'members_added');
    assert.equal(i.membership[0]?.['conversationId'], MEETING_CHAT_ID);
    assert.deepEqual(i.membership[0]?.['members'], [{ kind: 'teams-aad', id: CARLA.aadObjectId, displayName: CARLA.name }]);
    assert.deepEqual(i.kernel.events.map((e) => [e.author.id, e.text]), [
      [ANNA.aadObjectId, 'A'],
      [CARLA.aadObjectId, 'B'],
    ]);
  });
});

describe('scenario 7b — scheduled meeting with lifecycle and presence', () => {
  it('start, people joining, a late joiner and the end reach the kernel; presence is declared only while proven', async () => {
    const i = startInstance();
    i.kernel.subscribe(MEETING_CHAT_ID, MIRA);
    i.graph.install(MEETING_CHAT_ID, MIRA);
    const inputs = async (): Promise<readonly string[]> =>
      (await i.meetings.snapshot(MEETING_CHAT_ID, { asChannelKey: MIRA.key }))?.inputs ?? [];

    await i.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start', at: '2026-11-18T10:00:00.000Z' }));
    assert.equal((await inputs()).includes('presence'), false, 'no participant event seen yet');
    await i.bot.deliver(meetingParticipantActivity({ bot: MIRA, from: ANNA, phase: 'join', people: [ANNA, BERT], at: '2026-11-18T10:00:20.000Z' }));
    assert.equal((await inputs()).includes('presence'), true);
    await i.bot.deliver(meetingParticipantActivity({ bot: MIRA, from: CARLA, phase: 'join', people: [CARLA], at: '2026-11-18T10:25:00.000Z' }));
    await i.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'end', at: '2026-11-18T10:45:00.000Z' }));

    assert.deepEqual(
      i.sessionEvents.map((e) => [e.kind, e.participants?.map((p) => p.displayName), e.occurredAt]),
      [
        ['session_started', undefined, '2026-11-18T10:00:00.000Z'],
        ['participant_joined', [ANNA.name, BERT.name], '2026-11-18T10:00:20.000Z'],
        ['participant_joined', [CARLA.name], '2026-11-18T10:25:00.000Z'],
        ['session_ended', undefined, '2026-11-18T10:45:00.000Z'],
      ],
    );
    assert.equal((await inputs()).includes('presence'), false, 'withdrawn when the meeting ends');
    assert.deepEqual(i.bot.turns, []);
  });

  it('an install from the earlier package gets no presence until its consent is upgraded', async () => {
    const i = startInstance();
    i.graph.install(MEETING_CHAT_ID, MIRA, { consented: ['ChatMessage.Read.Chat', 'ChatMember.Read.Chat', 'TeamsActivity.Send.Chat'] });
    const hasPresence = async (): Promise<boolean> =>
      ((await i.meetings.snapshot(MEETING_CHAT_ID, { asChannelKey: MIRA.key }))?.inputs ?? []).includes('presence');
    // Simulated only: in Teams an install without the permissions would not
    // receive these events at all.
    await i.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));
    await i.bot.deliver(meetingParticipantActivity({ bot: MIRA, from: ANNA, phase: 'join', people: [ANNA] }));
    assert.equal(await hasPresence(), false, 'the grant is read from Graph, not assumed');

    const repair = await i.readiness.repairConsent(MEETING_CHAT_ID, { asChannelKey: MIRA.key });

    assert.equal(repair.outcome, 'upgraded');
    assert.equal(await hasPresence(), true);
  });

  it('a Meet-now call keeps chat and cards only', async () => {
    const i = startInstance();
    i.graph.install(GROUP_CHAT_ID, MIRA);

    await i.bot.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start', meetingType: 'MeetNow', conversationId: GROUP_CHAT_ID }));

    const snapshot = await i.meetings.snapshot(GROUP_CHAT_ID, { asChannelKey: MIRA.key });
    assert.deepEqual(snapshot?.inputs, ['text', 'message_edits', 'card_answers']);
  });
});


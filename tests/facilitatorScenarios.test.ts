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
  TeamsConversationReferenceCache,
  TeamsGraphResolver,
} from '@omadia/channel-teams';
import type { TeamsConversationRefPersistence } from '@omadia/channel-teams';

import {
  ANNA,
  BERT,
  botNotInstalledError,
  CARLA,
  cardSubmitActivity,
  editActivity,
  GROUP_CHAT_ID,
  JONAS,
  MEETING_CHAT_ID,
  membersAddedActivity,
  messageActivity,
  messageWritesBlockedError,
  MIRA,
  OTHER_GROUP_CHAT_ID,
  SimulatedConnector,
  SimulatedGraph,
  SimulatedKernel,
  simulatedBot,
  TENANT_ID,
  type SimBot,
  type SimUser,
} from './support/teamsSimulation.js';

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
  carry: Partial<Pick<Instance, 'kernel' | 'graph' | 'connector'>> & { table?: SimulatedRefTable } = {},
): Instance & { table: SimulatedRefTable } {
  const kernel = carry.kernel ?? new SimulatedKernel();
  const graph = carry.graph ?? new SimulatedGraph();
  const connector = carry.connector ?? new SimulatedConnector();
  const table = carry.table ?? new SimulatedRefTable();
  const refs = new TeamsConversationReferenceCache();
  refs.attachPersistence(table);
  const membership: Record<string, unknown>[] = [];
  const bot = simulatedBot(kernel, {
    captureConversationReference: (context) => refs.capture(context),
    emitMembershipEvent: (event) => {
      membership.push(event);
    },
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
    bot,
    post: postAdapter.sendToConversation.bind(postAdapter),
    dm: dmAdapter.sendToUser.bind(dmAdapter),
    readiness,
    meetings: createTeamsMeetingSource({ refs }),
  };
}

function dmTarget(user: SimUser): Parameters<Instance['dm']>[0] {
  return { principalId: `user:${user.aadObjectId}`, userRef: { kind: 'teams-aad', id: user.aadObjectId } };
}

/** Let fire-and-forget persistence writes settle. */
async function settle(): Promise<void> {
  await new Promise((resolve) => setImmediate(resolve));
}

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
    assert.equal(i.bot.turns.length, 2, 'only the two mentions start a turn');
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
    i.connector.failDirect(MIRA, BERT, messageWritesBlockedError());

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

  it('known limit: the revision counter starts again after a restart', async () => {
    const before = startInstance();
    before.kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const ballot = messageActivity({ bot: MIRA, from: ANNA, text: 'A' });
    const id = ballot.id as string;
    await before.bot.deliver(ballot);
    await before.bot.deliver(editActivity({ bot: MIRA, from: ANNA, id, text: 'B', timestamp: '2026-11-18T09:20:00.000Z' }));

    const after = startInstance({ kernel: before.kernel });
    await after.bot.deliver(editActivity({ bot: MIRA, from: ANNA, id, text: 'C', timestamp: '2026-11-18T09:40:00.000Z' }));

    // Two edits, both numbered 1. `occurredAt` still orders them; the kernel
    // must not rely on the adapter's revision across a restart.
    assert.deepEqual(
      after.kernel.events.map((e) => [e.kind, e.revision, e.occurredAt]),
      [
        ['created', 0, '2026-11-18T09:15:00.000Z'],
        ['edited', 1, '2026-11-18T09:20:00.000Z'],
        ['edited', 1, '2026-11-18T09:40:00.000Z'],
      ],
    );
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

describe('scenario 7 — scheduled meeting chat', () => {
  it('the facilitator reads the meeting chat without mentions and declares what it perceives', async () => {
    const i = startInstance();
    i.kernel.subscribe(MEETING_CHAT_ID, MIRA);

    await i.bot.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'Thema: Budget', conversationId: MEETING_CHAT_ID }));
    const posted = await i.post(MEETING_CHAT_ID, { text: 'Agenda: 1. Budget' }, { asChannelKey: MIRA.key });
    const snapshot = await i.meetings.snapshot(MEETING_CHAT_ID, { asChannelKey: MIRA.key, sessionRef: 'session-1' });

    assert.deepEqual(i.kernel.events.map((e) => [e.conversationId, e.text]), [[MEETING_CHAT_ID, 'Thema: Budget']]);
    assert.equal(posted.outcome, 'delivered');
    assert.equal(snapshot?.modality, 'online');
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
    const posted = await i.post(MEETING_CHAT_ID, { text: 'Abstimmung', poll } as never, { asChannelKey: MIRA.key });
    assert.ok(posted.outcome === 'delivered' && posted.messageId);
    const cardMessageId = posted.outcome === 'delivered' ? (posted.messageId as string) : '';
    const click = (from: SimUser, optionId: string): Promise<void> =>
      i.bot.deliver(
        cardSubmitActivity({
          bot: MIRA,
          from,
          conversationId: MEETING_CHAT_ID,
          cardMessageId,
          value: { type: POLL_ANSWER_VALUE_TYPE, pollId: 'poll-budget', optionId },
        }),
      );

    await click(ANNA, 'yes');
    await click(BERT, 'no');
    await click(BERT, 'yes');

    assert.deepEqual(
      i.kernel.events.slice(before).map((e) => [e.kind, e.messageId, e.author.id, e.cardAnswer]),
      [
        ['card_answer', cardMessageId, ANNA.aadObjectId, { cardId: 'poll-budget', value: { optionId: 'yes' } }],
        ['card_answer', cardMessageId, BERT.aadObjectId, { cardId: 'poll-budget', value: { optionId: 'no' } }],
        ['card_answer', cardMessageId, BERT.aadObjectId, { cardId: 'poll-budget', value: { optionId: 'yes' } }],
      ],
      'a changed answer is a second event; the engine decides which one counts',
    );
    assert.equal(i.bot.turns.length, 1, 'only the opening mention was a turn');
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

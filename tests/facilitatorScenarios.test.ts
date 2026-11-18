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
  editActivity,
  GROUP_CHAT_ID,
  JONAS,
  MEETING_CHAT_ID,
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
  carry: Partial<Pick<Instance, 'kernel' | 'graph' | 'connector'>> & { table?: SimulatedRefTable } = {},
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


import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  CHAT_READ_ALL_RSC_PERMISSION,
  createTeamsReadinessProvider,
  createTeamsTargetedSendAdapter,
  POLL_ANSWER_VALUE_TYPE,
  TeamsConversationReferenceCache,
  TeamsGraphResolver,
} from '@omadia/channel-teams';

import {
  ANNA,
  BERT,
  cardSubmitActivity,
  channelThreadId,
  deleteActivity,
  directConversationId,
  editActivity,
  GROUP_CHAT_ID,
  JONAS,
  MEETING_CHAT_ID,
  messageActivity,
  MIRA,
  quotedReplyActivity,
  reactionActivity,
  SimulatedConnector,
  SimulatedGraph,
  SimulatedKernel,
  simulatedBot,
  TEMPLATE_RSC,
  TENANT_ID,
} from './support/teamsSimulation.js';

/**
 * byte5ai/omadia#2043 / #2089 (K4.3, channel-sdk 0.3.5) — DELIVERY MODES.
 *
 * The kernel answers per bot and conversation what it is handed: `all`,
 * `addressed` or `none`. A facilitation in follow-up holds `addressed`: the
 * mention gate is back, and what addresses the facilitator still reaches the
 * process as an event.
 *
 * The simulated kernel here COLLECTS what the adapter hands over. It does not
 * model the kernel's own routing (a reply is kept only when it answers a
 * message bound to a process). NOT VERIFIED BY THESE TESTS: what a Teams
 * client sends for a reply in a 1:1 chat or a channel; the quote shape is the
 * one of `quotedReplyActivity`, which was not captured live.
 */

const REMINDER_ID = '1731930009001';

function followUp(): { kernel: SimulatedKernel; sim: ReturnType<typeof simulatedBot> } {
  const kernel = new SimulatedKernel();
  kernel.subscribe(GROUP_CHAT_ID, MIRA, 'addressed');
  return { kernel, sim: simulatedBot(kernel) };
}

const quoteOfReminder = (text: string, extra: { mention?: boolean; conversationId?: string; conversationType?: 'personal' | 'channel' } = {}) =>
  quotedReplyActivity({
    bot: MIRA,
    from: ANNA,
    text,
    quotedMessageId: REMINDER_ID,
    quotedAuthor: { id: `28:${MIRA.appId}`, aadObjectId: '', name: MIRA.name },
    quotedText: 'Erinnerung: Angebot einholen bis 28.11.',
    ...extra,
  });

describe('#2089 addressed mode in a group chat (a facilitation in follow-up)', () => {
  it('a message that addresses nobody is not handed over and starts no turn: the mention gate is back', async () => {
    const { kernel, sim } = followUp();

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'Mittagessen um 12?' }));

    assert.deepEqual(kernel.events, []);
    assert.deepEqual(sim.turns, []);
    assert.deepEqual(sim.sent, []);
  });

  it('a mention is an event with addressedToBot AND the ordinary turn; the adapter posts the one turn answer', async () => {
    const { kernel, sim } = followUp();
    const mention = messageActivity({ bot: MIRA, from: ANNA, text: 'das Angebot ist da', mention: true });

    await sim.deliver(mention);

    assert.deepEqual(
      kernel.events.map((e) => [e.kind, e.messageId, e.addressedToBot, e.author.id, e.receivedByChannelKey]),
      [['created', mention.id, true, ANNA.aadObjectId, MIRA.key]],
    );
    assert.equal(sim.turns.length, 1, 'the ordinary mention turn still runs');
    // Exactly one post for this message comes from the adapter: the turn's
    // answer. Handing over the event posts nothing.
    assert.equal(sim.sent.filter((activity) => (activity as { type?: string }).type !== 'typing').length, 1);
  });

  it('a mention whose turn answer is empty (the engine answers) posts nothing at all', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA, 'addressed');
    const sim = simulatedBot(kernel, {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      answerText: '',
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'erledigt', mention: true }));

    assert.equal(kernel.events.length, 1);
    assert.equal(sim.turns.length, 1);
    assert.deepEqual(
      sim.sent.filter((activity) => (activity as { type?: string }).type === 'message'),
      [],
      'no second voice next to the engine',
    );
  });

  it('a quoted reply without a mention is an event with its reply reference and no turn', async () => {
    const { kernel, sim } = followUp();
    const reply = quoteOfReminder('ist erledigt');

    await sim.deliver(reply);

    assert.deepEqual(
      kernel.events.map((e) => [e.kind, e.messageId, e.replyToMessageId, e.quoted, e.addressedToBot, e.text]),
      // Not `addressedToBot`: whether the quoted message is one of a
      // process's own is the kernel's record, not the adapter's claim.
      [['created', reply.id, REMINDER_ID, true, undefined, 'ist erledigt']],
    );
    assert.deepEqual(sim.turns, [], 'past the mention gate as an event only');
    assert.deepEqual(sim.sent, []);
  });

  it('a quote stated as a quotedReply entity (the shape Microsoft describes, not seen live) is read the same way', async () => {
    const { kernel, sim } = followUp();

    await sim.deliver({
      ...messageActivity({ bot: MIRA, from: ANNA, text: 'ist erledigt' }),
      entities: [{ type: 'quotedReply', quotedReply: { messageId: REMINDER_ID, senderName: MIRA.name, preview: 'Erinnerung' } }],
    });

    assert.deepEqual(kernel.events.map((e) => [e.replyToMessageId, e.quoted]), [[REMINDER_ID, true]]);
    assert.deepEqual(sim.turns, []);
  });

  it('a quoted reply WITH a mention carries both and runs the turn', async () => {
    const { kernel, sim } = followUp();

    await sim.deliver(quoteOfReminder('ist erledigt', { mention: true }));

    assert.deepEqual(kernel.events.map((e) => [e.replyToMessageId, e.addressedToBot]), [[REMINDER_ID, true]]);
    assert.equal(sim.turns.length, 1);
  });

  it('reactions are handed over; an edit only when the message addresses the bot; a deletion never', async () => {
    const { kernel, sim } = followUp();

    await sim.deliver(reactionActivity({ bot: MIRA, from: BERT, targetMessageId: REMINDER_ID, added: ['like'] }));
    await sim.deliver(editActivity({ bot: MIRA, from: ANNA, id: '1731930009100', text: 'nur ein Tippfehler' }));
    const editedMention = {
      ...editActivity({ bot: MIRA, from: ANNA, id: '1731930009101', text: `<at>${MIRA.name}</at> doch erst morgen` }),
      entities: [{ type: 'mention', mentioned: { id: `28:${MIRA.appId}`, name: MIRA.name }, text: `<at>${MIRA.name}</at>` }],
    };
    await sim.deliver(editedMention);
    await sim.deliver(deleteActivity({ bot: MIRA, from: ANNA, id: '1731930009101' }));

    assert.deepEqual(
      kernel.events.map((e) => [e.kind, e.messageId, e.addressedToBot]),
      [
        ['reaction', REMINDER_ID, undefined],
        ['edited', '1731930009101', true],
      ],
    );
    assert.deepEqual(sim.turns, []);
  });

  it('a card click is handed over: a poll click ends there, without a turn', async () => {
    const { kernel, sim } = followUp();

    await sim.deliver(
      cardSubmitActivity({
        bot: MIRA,
        from: BERT,
        cardMessageId: '1731930009200',
        value: { type: POLL_ANSWER_VALUE_TYPE, pollId: 'poll-9', optionId: 'a' },
      }),
    );

    assert.deepEqual(kernel.events.map((e) => [e.kind, e.messageId, e.cardAnswer?.cardId, e.author.id]), [
      ['card_answer', '1731930009200', 'poll-9', BERT.aadObjectId],
    ]);
    assert.deepEqual(sim.turns, []);
  });

  it('is per bot: another bot in the same chat hands over nothing and keeps its gate', async () => {
    const { kernel, sim } = followUp();

    await sim.deliver(messageActivity({ bot: JONAS, from: ANNA, text: 'hallo', mention: true }));
    // The same quoted reply as Jonas receives it: Jonas holds no delivery.
    await sim.deliver({ ...quoteOfReminder('erledigt'), recipient: { id: `28:${JONAS.appId}`, name: JONAS.name } });

    assert.deepEqual(kernel.events, []);
    assert.equal(sim.turns.length, 1, 'only the mention of Jonas is a turn of Jonas');
  });
});

describe('#2089 addressed mode in a channel thread', () => {
  const THREAD = channelThreadId('1731930001000');

  it('replyToId that is only the thread root is no reply to anybody; a quote inside the thread is', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(THREAD, MIRA, 'addressed');
    const sim = simulatedBot(kernel);
    const inThread = { conversationId: THREAD, conversationType: 'channel' as const };

    // Every message of a thread carries the root as `replyToId`.
    await sim.deliver({ ...messageActivity({ bot: MIRA, from: ANNA, text: 'weiter im Thread', ...inThread }), replyToId: '1731930001000' });
    const quote = { ...quoteOfReminder('erledigt', inThread), replyToId: '1731930001000' };
    await sim.deliver(quote);

    assert.deepEqual(kernel.events.map((e) => [e.messageId, e.replyToMessageId, e.conversationId]), [[quote.id, REMINDER_ID, THREAD]]);
    assert.deepEqual(sim.turns, []);
  });
});

describe('#2089 all, none, and a kernel without delivery modes', () => {
  it('all: as before (recorded, no turn); a mention is marked as addressed for other processes in follow-up', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA, 'all');
    const sim = simulatedBot(kernel);

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A' }));
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'stopp', mention: true }));
    await sim.deliver(deleteActivity({ bot: MIRA, from: ANNA, id: '1731930009300' }));

    assert.deepEqual(kernel.events.map((e) => [e.kind, e.addressedToBot]), [
      ['created', undefined],
      ['created', true],
      ['deleted', undefined],
    ]);
    assert.deepEqual(sim.turns, []);
  });

  it('none: nothing is handed over, not even a mention or a reply; the mention is an ordinary turn', async () => {
    const kernel = new SimulatedKernel();
    const sim = simulatedBot(kernel);

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));
    await sim.deliver(quoteOfReminder('erledigt'));
    await sim.deliver(reactionActivity({ bot: MIRA, from: BERT, targetMessageId: REMINDER_ID, added: ['like'] }));

    assert.deepEqual(kernel.events, []);
    assert.equal(sim.turns.length, 1);
  });

  it('a kernel before channel-sdk 0.3.5 knows all and none only: follow-up gets nothing, moderation works as before', async () => {
    const kernel = new SimulatedKernel();
    kernel.offersDeliveryModes = false;
    kernel.subscribe(GROUP_CHAT_ID, MIRA, 'addressed'); // such a kernel says "not subscribed" for it
    const sim = simulatedBot(kernel);

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));
    assert.deepEqual(kernel.events, []);
    assert.equal(sim.turns.length, 1);

    kernel.subscribe(GROUP_CHAT_ID, MIRA, 'all');
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A > B' }));
    assert.equal(kernel.events.length, 1);
    assert.equal(sim.turns.length, 1);
  });
});

describe('#2089 / #2118 private messages to a facilitator bot', () => {
  const DM = directConversationId(`${MIRA.appId}|${ANNA.aadObjectId}`);
  const privateMessage = (text: string) =>
    messageActivity({ bot: MIRA, from: ANNA, text, conversationId: DM, conversationType: 'personal' });

  it('every private message is a turn whose origin states the author as Teams does: the private status path', async () => {
    const kernel = new SimulatedKernel();
    const sim = simulatedBot(kernel, { captureConversationReference: () => {}, emitMembershipEvent: () => {}, ownIdentity: true });

    await sim.deliver(privateMessage('Wie ist der Stand?'));

    assert.deepEqual(sim.turns, ['Wie ist der Stand?']);
    const origin = sim.origins[0] as Record<string, unknown>;
    assert.equal(origin['channelType'], 'teams');
    assert.equal(origin['conversationType'], 'direct');
    assert.equal(origin['channelConversationId'], DM);
    // `principal.userId` is what the kernel's private status path compares
    // with the people a running process lists: the Entra object id.
    assert.deepEqual(origin['principal'], { kind: 'user', userId: ANNA.aadObjectId });
    assert.deepEqual(origin['speaker'], { kind: 'teams-aad', id: ANNA.aadObjectId, displayName: ANNA.name });
    assert.deepEqual(kernel.events, [], 'no event while the kernel asks for nothing of this 1:1');
  });

  it('when the kernel asks for the 1:1 (addressed), a private message is ALSO an event: addressed, with the author, and the turn still runs', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(DM, MIRA, 'addressed');
    const sim = simulatedBot(kernel);
    const message = privateMessage('Ich habe das Angebot eingeholt.');

    await sim.deliver(message);

    assert.deepEqual(
      kernel.events.map((e) => [e.kind, e.messageId, e.conversationId, e.conversationType, e.addressedToBot, e.replyToMessageId, e.author.id, e.tenantId]),
      // A plain message in a 1:1 names no message it answers.
      [['created', message.id, DM, 'direct', true, undefined, ANNA.aadObjectId, TENANT_ID]],
    );
    assert.deepEqual(sim.turns, ['Ich habe das Angebot eingeholt.'], 'the kernel decides what the turn answers');
  });

  it('a private reply that quotes a reminder names the reminder: conversation and message id of the direct send', async () => {
    const connector = new SimulatedConnector();
    const refs = new TeamsConversationReferenceCache();
    const direct = createTeamsTargetedSendAdapter({
      sendProactive: connector.sendProactive,
      direct: { refs, openDirectConversation: connector.openDirectConversation, botLabel: () => MIRA.name },
    });
    const reminder = await direct.sendToUser(
      { principalId: `user:${ANNA.aadObjectId}`, userRef: { kind: 'teams-aad', id: ANNA.aadObjectId } },
      { text: 'Erinnerung: Angebot einholen bis 28.11.' },
      { asChannelKey: MIRA.key },
    );
    assert.equal(reminder.outcome, 'delivered');
    const sentTo = reminder.outcome === 'delivered' ? reminder : undefined;
    assert.equal(sentTo?.conversationId, DM, 'the direct send states the 1:1 it went into');
    assert.ok(sentTo?.messageId);

    const kernel = new SimulatedKernel();
    kernel.subscribe(DM, MIRA, 'addressed');
    const sim = simulatedBot(kernel);
    await sim.deliver(
      quotedReplyActivity({
        bot: MIRA,
        from: ANNA,
        text: 'erledigt',
        conversationId: DM,
        conversationType: 'personal',
        quotedMessageId: sentTo?.messageId ?? '',
        quotedAuthor: { id: `28:${MIRA.appId}`, aadObjectId: '', name: MIRA.name },
        quotedText: 'Erinnerung: Angebot einholen bis 28.11.',
      }),
    );

    assert.deepEqual(kernel.events.map((e) => [e.conversationId, e.replyToMessageId, e.quoted, e.addressedToBot, e.text]), [
      [sentTo?.conversationId, sentTo?.messageId, true, true, 'erledigt'],
    ]);
  });
});

describe('#2043 criterion 60: readiness of a linked meeting chat is its own check', () => {
  function harness() {
    const graph = new SimulatedGraph();
    const refs = new TeamsConversationReferenceCache();
    const provider = createTeamsReadinessProvider({
      botLabel: (botAppId) => (botAppId === MIRA.appId ? MIRA.name : undefined),
      graph: new TeamsGraphResolver({ tenantId: TENANT_ID, clientId: '737c6ddd-0000-4000-8000-000000000001', clientSecret: 'simulated', log: () => {}, fetchImpl: graph.fetch }),
      refs,
      template: { revision: 3, permissions: TEMPLATE_RSC },
    });
    const seen = (conversationId: string): void =>
      refs.capture({ activity: messageActivity({ bot: MIRA, from: ANNA, text: 'hi', mention: true, conversationId }) } as unknown as TurnContext);
    return { graph, provider, seen, check: (conversationId: string) => provider.check(conversationId, { asChannelKey: MIRA.key }) };
  }

  it('the home conversation is ready; the meeting chat hears mentions only: not ready, with the reason, and the home is not affected', async () => {
    const h = harness();
    h.seen(GROUP_CHAT_ID);
    h.graph.install(GROUP_CHAT_ID, MIRA);
    // Linked from inside the meeting chat by a mention: the bot is installed
    // there and holds a reference, but all-messages consent was not given.
    h.seen(MEETING_CHAT_ID);
    h.graph.install(MEETING_CHAT_ID, MIRA, { consented: ['TeamsActivity.Send.Chat'] });

    const [home, meeting] = [await h.check(GROUP_CHAT_ID), await h.check(MEETING_CHAT_ID)];

    assert.equal(home.ready, true);
    assert.deepEqual(
      { ready: meeting.ready, installed: meeting.installed, canReadAllMessages: meeting.canReadAllMessages, canPost: meeting.canPost },
      { ready: false, installed: true, canReadAllMessages: false, canPost: true },
    );
    assert.deepEqual(meeting.missing.map((issue) => issue.code), ['consent_missing']);
    assert.match(meeting.missing[0]?.message ?? '', new RegExp(CHAT_READ_ALL_RSC_PERMISSION.replace(/\./g, '\\.')));
    const meetingChat = encodeURIComponent(MEETING_CHAT_ID);
    assert.ok(h.graph.requests.some((r) => r.includes(`/chats/${meetingChat}/permissionGrants`)), 'asked for the meeting chat itself');
  });

  it('after the permission is granted there, the next check (before the session goes live) says ready: nothing is cached', async () => {
    const h = harness();
    h.seen(MEETING_CHAT_ID);
    h.graph.install(MEETING_CHAT_ID, MIRA, { consented: ['TeamsActivity.Send.Chat'] });
    assert.equal((await h.check(MEETING_CHAT_ID)).ready, false);

    h.graph.install(MEETING_CHAT_ID, MIRA);

    assert.deepEqual(await h.check(MEETING_CHAT_ID), { installed: true, canReadAllMessages: true, canPost: true, ready: true, missing: [] });
  });

  it('a meeting chat the bot was never addressed in and is not installed in: not installed, cannot post', async () => {
    const h = harness();
    // The chat exists for Graph (another app is installed there); Mira's is not.
    h.graph.install(MEETING_CHAT_ID, JONAS);

    const report = await h.check(MEETING_CHAT_ID);

    assert.equal(report.ready, false);
    assert.deepEqual(report.missing.map((issue) => issue.code), ['app_not_installed', 'cannot_post']);
  });
});

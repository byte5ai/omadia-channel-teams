import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  buildPollCard,
  parsePollAnswerValue,
  plainMessageText,
  platformAuthor,
  POLL_ANSWER_VALUE_TYPE,
  readQuoteInfo,
  validatePollInput,
} from '@omadia/channel-teams';

import {
  ANNA,
  BERT,
  cardSubmitActivity,
  deleteActivity,
  editActivity,
  GROUP_CHAT_ID,
  JONAS,
  messageActivity,
  MIRA,
  OTHER_GROUP_CHAT_ID,
  quotedReplyActivity,
  reactionActivity,
  SimulatedKernel,
  simulatedBot,
  TENANT_ID,
  undeleteActivity,
} from './support/teamsSimulation.js';

// byte5ai/omadia#2043 §5.3 / §6.2, issue #43 — message events and the
// mention gate. Simulated transport: see tests/support/teamsSimulation.ts.

describe('#43 mention gate under a message subscription', () => {
  it('records a message without mention and starts no turn', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const sim = simulatedBot(kernel);
    const vote = messageActivity({ bot: MIRA, from: ANNA, text: 'A > C > B' });

    await sim.deliver(vote);

    assert.equal(kernel.events.length, 1);
    const event = kernel.events[0];
    assert.ok(event);
    assert.equal(event.kind, 'created');
    assert.equal(event.revision, 0);
    assert.equal(event.messageId, vote.id);
    assert.equal(event.text, 'A > C > B');
    assert.equal(event.conversationId, GROUP_CHAT_ID);
    assert.equal(event.conversationType, 'group');
    assert.equal(event.receivedByChannelKey, MIRA.key);
    assert.equal(event.channelType, 'teams');
    assert.equal(event.tenantId, TENANT_ID);
    assert.deepEqual(event.author, { kind: 'teams-aad', id: ANNA.aadObjectId, displayName: ANNA.name });
    assert.equal(event.occurredAt, '2026-11-18T09:15:00.000Z');
    assert.deepEqual(sim.turns, [], 'a message event is not a turn');
    assert.deepEqual(sim.sent, [], 'the bot says nothing');
  });

  it('records a mentioned message AND runs the normal turn', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const sim = simulatedBot(kernel);

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'wie ist der Stand?', mention: true }));

    assert.equal(kernel.events.length, 1);
    assert.equal(kernel.events[0]?.kind, 'created');
    assert.equal(kernel.events[0]?.text, 'Mira wie ist der Stand?');
    assert.equal(sim.turns.length, 1, 'the mention still reaches the agent');
  });

  it('changes nothing in a chat without subscription', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(OTHER_GROUP_CHAT_ID, MIRA);
    const sim = simulatedBot(kernel);

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A > C > B' }));
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    assert.deepEqual(kernel.events, [], 'nothing is emitted for a chat nobody moderates');
    assert.equal(sim.turns.length, 1, 'only the mention reaches the agent, as before');
  });

  it('handles a two-person chat with an agent app like any moderated chat', async () => {
    // Teams delivers a 1:1 between two people as `groupChat`; nothing here
    // may conclude a head count from the conversation type.
    const twoPersonChat = '19:5b1e2c3d-4f5a-4b6c-8d7e-9f0a1b2c3d4e_6c2f3d4e-5a6b-4c7d-9e8f-0a1b2c3d4e5f@unq.gbl.spaces';
    const kernel = new SimulatedKernel();
    kernel.subscribe(twoPersonChat, MIRA);
    const sim = simulatedBot(kernel);

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A', conversationId: twoPersonChat }));

    assert.equal(kernel.events.length, 1);
    assert.equal(kernel.events[0]?.conversationType, 'group');
    assert.deepEqual(sim.turns, []);
  });

  it('keeps the gate for another bot in the same chat', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const sim = simulatedBot(kernel);
    const id = '1731930000999';

    // Teams delivers the same message once per bot in the chat.
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'B', id }));
    await sim.deliver(messageActivity({ bot: JONAS, from: ANNA, text: 'B', id }));

    assert.equal(kernel.events.length, 1);
    assert.equal(kernel.events[0]?.receivedByChannelKey, MIRA.key);
    assert.deepEqual(sim.turns, []);
  });

  it('emits nothing for a message sent by a bot', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const sim = simulatedBot(kernel);
    const fromBot = messageActivity({ bot: MIRA, from: ANNA, text: 'Ergebnis: A' });

    await sim.deliver({ ...fromBot, from: { id: JONAS.key, name: JONAS.name } });

    assert.deepEqual(kernel.events, []);
  });

  it('stops recording when the subscription ends', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const sim = simulatedBot(kernel);
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'A' }));
    kernel.unsubscribe(GROUP_CHAT_ID, MIRA);
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'B' }));

    assert.equal(kernel.events.length, 1);
  });
});

describe('#43 edits, deletions, reactions and card answers', () => {
  it('increments the revision per edit and keeps the message id', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const sim = simulatedBot(kernel);
    const original = messageActivity({ bot: MIRA, from: ANNA, text: 'A > B > C' });
    const id = original.id as string;

    await sim.deliver(original);
    await sim.deliver(editActivity({ bot: MIRA, from: ANNA, id, text: 'B > A > C' }));
    await sim.deliver(editActivity({ bot: MIRA, from: ANNA, id, text: 'C > A > B' }));

    assert.deepEqual(
      kernel.events.map((e) => [e.kind, e.messageId, e.revision, e.text]),
      [
        ['created', id, 0, 'A > B > C'],
        ['edited', id, 1, 'B > A > C'],
        ['edited', id, 2, 'C > A > B'],
      ],
    );
    assert.equal(kernel.events[2]?.author.id, ANNA.aadObjectId);
    assert.deepEqual(sim.turns, []);
  });

  it('counts revisions per receiving bot', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA);
    kernel.subscribe(GROUP_CHAT_ID, JONAS);
    const sim = simulatedBot(kernel);
    const id = '1731930001234';

    await sim.deliver(editActivity({ bot: MIRA, from: ANNA, id, text: 'x' }));
    await sim.deliver(editActivity({ bot: JONAS, from: ANNA, id, text: 'x' }));

    assert.deepEqual(kernel.eventsFor(MIRA).map((e) => e.revision), [1]);
    assert.deepEqual(kernel.eventsFor(JONAS).map((e) => e.revision), [1]);
  });

  it('emits a deletion without text and an undelete as a new revision', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const sim = simulatedBot(kernel);
    const id = '1731930002345';

    await sim.deliver(deleteActivity({ bot: MIRA, from: ANNA, id }));
    await sim.deliver(undeleteActivity({ bot: MIRA, from: ANNA, id, text: 'A' }));

    assert.equal(kernel.events[0]?.kind, 'deleted');
    assert.equal(kernel.events[0]?.revision, 1);
    assert.equal('text' in (kernel.events[0] ?? {}), false);
    assert.equal(kernel.events[1]?.kind, 'edited');
    assert.equal(kernel.events[1]?.revision, 2);
    assert.equal(kernel.events[1]?.text, 'A');
  });

  it('emits one event per reaction, added and removed', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const sim = simulatedBot(kernel);

    await sim.deliver(reactionActivity({ bot: MIRA, from: BERT, targetMessageId: '1731930003456', added: ['like'] }));
    await sim.deliver(
      reactionActivity({ bot: MIRA, from: BERT, targetMessageId: '1731930003456', added: ['heart'], removed: ['like'] }),
    );

    assert.deepEqual(
      kernel.events.map((e) => [e.kind, e.messageId, e.reaction, e.author.id]),
      [
        ['reaction', '1731930003456', { name: 'like', removed: false }, BERT.aadObjectId],
        ['reaction', '1731930003456', { name: 'heart', removed: false }, BERT.aadObjectId],
        ['reaction', '1731930003456', { name: 'like', removed: true }, BERT.aadObjectId],
      ],
    );
  });

  it('emits a poll click as card_answer with the clicking person as author and no turn', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const sim = simulatedBot(kernel);

    await sim.deliver(
      cardSubmitActivity({
        bot: MIRA,
        from: BERT,
        cardMessageId: '1731930004567',
        // A forged name inside the payload must not matter.
        value: { type: POLL_ANSWER_VALUE_TYPE, pollId: 'poll-7', optionId: 'b', voter: ANNA.name },
      }),
    );

    assert.equal(kernel.events.length, 1);
    const event = kernel.events[0];
    assert.ok(event);
    assert.equal(event.kind, 'card_answer');
    assert.equal(event.messageId, '1731930004567');
    assert.deepEqual(event.cardAnswer, { cardId: 'poll-7', value: { optionId: 'b' } });
    assert.equal(event.author.id, BERT.aadObjectId);
    assert.deepEqual(sim.turns, []);
    assert.deepEqual(sim.sent, []);
  });

  it('never turns a poll click into a chat turn, subscribed or not', async () => {
    const kernel = new SimulatedKernel();
    const sim = simulatedBot(kernel);

    await sim.deliver(
      cardSubmitActivity({
        bot: MIRA,
        from: BERT,
        cardMessageId: '1731930004568',
        value: JSON.stringify({ type: POLL_ANSWER_VALUE_TYPE, pollId: 'poll-7', optionId: 'a' }),
      }),
    );

    assert.deepEqual(kernel.events, []);
    assert.deepEqual(sim.turns, []);
    assert.deepEqual(sim.sent, []);
  });

  it('emits nothing for edits and reactions in a chat without subscription', async () => {
    const kernel = new SimulatedKernel();
    const sim = simulatedBot(kernel);

    await sim.deliver(editActivity({ bot: MIRA, from: ANNA, id: '1', text: 'x' }));
    await sim.deliver(deleteActivity({ bot: MIRA, from: ANNA, id: '1' }));
    await sim.deliver(reactionActivity({ bot: MIRA, from: ANNA, targetMessageId: '1', added: ['like'] }));

    assert.deepEqual(kernel.events, []);
  });
});

describe('#43 authorship and text come from the platform activity', () => {
  it('marks a quoted reply and names the quoted message', async () => {
    const kernel = new SimulatedKernel();
    kernel.subscribe(GROUP_CHAT_ID, MIRA);
    const sim = simulatedBot(kernel);

    await sim.deliver(
      quotedReplyActivity({
        bot: MIRA,
        from: BERT,
        text: 'Anna stimmt auch zu',
        quotedMessageId: '1731930005678',
        quotedAuthor: ANNA,
        quotedText: 'A > C > B',
      }),
    );

    const event = kernel.events[0];
    assert.ok(event);
    assert.equal(event.quoted, true);
    assert.equal(event.replyToMessageId, '1731930005678');
    assert.equal(event.author.id, BERT.aadObjectId, 'the quoting person is the author, not the quoted one');
  });

  it('marks a forwarded message', () => {
    const info = readQuoteInfo({ attachments: [{ contentType: 'forwardedMessageReference', content: {} }] });
    assert.deepEqual(info, { quoted: false, forwarded: true });
  });

  it('takes the author from `from`, falling back to the channel id without an AAD id', () => {
    assert.deepEqual(platformAuthor({ from: { id: '29:abc', name: 'Gast' } }), {
      kind: 'teams-aad',
      id: '29:abc',
      displayName: 'Gast',
    });
    assert.equal(platformAuthor({ from: { id: '28:app', name: 'Bot' } }), undefined);
    assert.equal(platformAuthor({ from: { id: '29:x', name: 'B', role: 'bot' } }), undefined);
    assert.equal(platformAuthor({}), undefined);
  });

  it('strips mention tokens and markup from the text', () => {
    assert.equal(plainMessageText('<at>Mira</at> <b>A</b> &gt; B &amp; C&nbsp;'), 'Mira A > B & C');
    assert.equal(plainMessageText(undefined), '');
  });
});

describe('#44 poll card', () => {
  it('renders one submit per option and round-trips the answer', () => {
    const card = buildPollCard({
      pollId: 'poll-7',
      question: 'Welcher Termin?',
      hint: 'Eine Stimme pro Person',
      options: [
        { id: 'a', label: 'Montag' },
        { id: 'b', label: 'Dienstag' },
      ],
    });
    assert.equal(card.contentType, 'application/vnd.microsoft.card.adaptive');
    const content = card.content as { body: { text: string }[]; actions: { type: string; title: string; data: unknown }[] };
    assert.deepEqual(content.body.map((b) => b.text), ['Welcher Termin?', 'Eine Stimme pro Person']);
    assert.deepEqual(content.actions.map((a) => [a.type, a.title]), [
      ['Action.Submit', 'Montag'],
      ['Action.Submit', 'Dienstag'],
    ]);
    assert.deepEqual(parsePollAnswerValue(content.actions[1]?.data), {
      type: POLL_ANSWER_VALUE_TYPE,
      pollId: 'poll-7',
      optionId: 'b',
    });
  });

  it('refuses an unusable poll with the reason', () => {
    assert.match(validatePollInput({ pollId: 'p', question: 'q', options: [{ id: 'a', label: 'A' }] }) ?? '', /at least two/);
    assert.match(
      validatePollInput({ pollId: 'p', question: 'q', options: [{ id: 'a', label: 'A' }, { id: 'a', label: 'B' }] }) ?? '',
      /occurs twice/,
    );
    assert.match(validatePollInput({ pollId: '', question: 'q', options: [] }) ?? '', /pollId/);
    assert.throws(() => buildPollCard({ pollId: 'p', question: ' ', options: [] }), /invalid poll/);
  });

  it('rejects payloads that are not a poll answer', () => {
    assert.equal(parsePollAnswerValue({ type: 'conductor.approval', decision: 'approve' }), undefined);
    assert.equal(parsePollAnswerValue({ type: POLL_ANSWER_VALUE_TYPE, pollId: 'p' }), undefined);
    assert.equal(parsePollAnswerValue('not json'), undefined);
    assert.equal(parsePollAnswerValue(undefined), undefined);
  });
});

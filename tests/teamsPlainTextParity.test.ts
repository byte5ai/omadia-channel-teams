import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import { parseCardReply, renderCardAsText } from '@omadia/channel-sdk';
import {
  createTeamsConversationSendAdapter,
  QUESTION_CARD_ANSWER_VERB,
  TeamsConversationReferenceCache,
  TeamsQuestionCards,
} from '@omadia/channel-teams';

import type { CardQuestion, QuestionCard } from '../src/kernel-types.js';
import {
  ANNA,
  BERT,
  cardExecuteActivity,
  channelThreadId,
  GROUP_CHAT_ID,
  messageActivity,
  MIRA,
  quotedReplyActivity,
  SimulatedConnector,
  SimulatedKernel,
  simulatedBot,
} from './support/teamsSimulation.js';

/**
 * Spec 2042 §5.3 / §5.14, criterion 30 (channel parity), for Teams: giving a
 * contribution in a facilitation and answering a poll also work through plain
 * text, next to the Adaptive Card Teams renders where it can. Choosing an agent
 * in plain text is driven in `teamsManagementLayer.test.ts`.
 *
 * Since kernel #2086 the KERNEL does the plain-text work: `sendToConversation`
 * posts a card or poll the channel does not show natively as numbered text
 * (`renderCardAsText`) and keeps its question open, and a matching typed
 * reply becomes the typed `card_answer`. This plugin renders and parses no
 * ballot itself; `parseCardReply` below stands in for the kernel's reading.
 * What the channel owes that path:
 *   - declare honestly which cards it renders, so the kernel sends the rest
 *     as text (and refuse, with nothing posted, a card it was handed anyway);
 *   - deliver the text exactly as written, and report the Teams activity id
 *     as its `messageId`;
 *   - hand a typed message of the chat to the engine as it was typed, with
 *     `replyToMessageId` only for a message the person really answered and
 *     `addressedToBot` for a mention (see the last block).
 */

const OPTIONS = [
  { id: 'a', label: 'Dienstag' },
  { id: 'b', label: 'Mittwoch' },
  { id: 'c', label: 'Freitag' },
];

const SINGLE: QuestionCard = {
  cardId: 'c-single',
  proposalId: 'p-4',
  proposalVersion: 2,
  title: 'Welchen Termin nehmen wir?',
  questions: [{ id: 'termin', label: 'Termin', kind: 'single_choice', required: true, options: OPTIONS }],
};

/** A question Teams does not render as a card (kind `number`): the text path is the only one. */
const COUNT: QuestionCard = {
  cardId: 'c-count',
  proposalId: 'p-5',
  proposalVersion: 1,
  title: 'Wie viele Plätze brauchen wir?',
  questions: [{ id: 'plaetze', label: 'Plätze', kind: 'number', required: true }],
};

function harness() {
  const kernel = new SimulatedKernel();
  kernel.subscribe(GROUP_CHAT_ID, MIRA);
  const refs = new TeamsConversationReferenceCache();
  const sim = simulatedBot(kernel, { captureConversationReference: (context) => refs.capture(context), emitMembershipEvent: () => {} });
  const cards = new TeamsQuestionCards({ now: () => Date.parse('2026-11-19T10:00:00.000Z') });
  sim.bot.attachQuestionCards(cards);
  const connector = new SimulatedConnector();
  refs.capture({ activity: messageActivity({ bot: MIRA, from: ANNA, text: 'hi' }) } as unknown as TurnContext);
  const provider = createTeamsConversationSendAdapter({ refs, sendProactive: connector.sendProactive, cards });
  return { kernel, sim, connector, provider };
}

/** The typed answer a person gives: a message in the moderated chat. */
async function type(h: ReturnType<typeof harness>, text: string): Promise<string> {
  const before = h.kernel.events.length;
  await h.sim.deliver(messageActivity({ bot: MIRA, from: BERT, text }));
  const event = h.kernel.events[before];
  assert.ok(event, 'the typed message reached the engine as an event');
  assert.equal(event.kind, 'created');
  assert.equal(event.author.id, BERT.aadObjectId, 'the author is who Teams says typed it');
  assert.deepEqual(h.sim.turns, [], 'a message in a moderated chat is never a chat turn');
  return event.text ?? '';
}

const questionOf = (card: QuestionCard): CardQuestion => {
  const question = card.questions[0];
  assert.ok(question);
  return question;
};

describe('criterion 30: a poll is answered through plain text in Teams', () => {
  it('declares only what it renders: a card the kernel must send as text is refused, nothing posted', async () => {
    const h = harness();
    assert.ok(!h.provider.supports.card?.kinds.includes('number'));

    const outcome = await h.provider.sendToConversation(GROUP_CHAT_ID, { text: 'x', card: COUNT } as never, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'unreachable');
    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'channel_error');
    assert.deepEqual(h.connector.sent, [], 'nothing was posted, so the engine posts the text form once');
  });

  it('delivers the rendered text exactly as written: numbered options, one message, the AI label', async () => {
    const h = harness();
    const text = renderCardAsText(SINGLE);

    const outcome = await h.provider.sendToConversation(GROUP_CHAT_ID, { text }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'delivered');
    assert.equal(h.connector.sent.length, 1);
    const activity = h.connector.sent[0]?.activity as { text?: string; attachments?: unknown[]; entities?: { additionalType?: string[] }[] };
    assert.equal(activity.text, text);
    assert.ok(text.includes('1) Dienstag') && text.includes('3) Freitag'));
    assert.equal(activity.attachments, undefined, 'no card: it is the plain-text path');
    assert.equal(activity.entities?.filter((e) => e.additionalType?.includes('AIGeneratedContent')).length, 1);
  });

  it('a typed "2" is the same typed answer as a click on the card', async () => {
    const h = harness();
    const posted = await h.provider.sendToConversation(GROUP_CHAT_ID, { text: '', card: SINGLE } as never, { asChannelKey: MIRA.key });
    assert.equal(posted.outcome, 'delivered');
    const cardMessageId = (posted as { cardMessageId?: string }).cardMessageId ?? '';
    const content = (h.connector.sent.at(-1)?.activity['attachments'] as { content: { actions: { data: Record<string, unknown> }[] } }[])[0]?.content;
    assert.ok(content);
    await h.sim.deliver(
      cardExecuteActivity({
        bot: MIRA,
        from: ANNA,
        conversationId: GROUP_CHAT_ID,
        cardMessageId,
        verb: QUESTION_CARD_ANSWER_VERB,
        data: { ...(content.actions.at(-1)?.data ?? {}), q0: 'b' },
      }),
    );
    const clicked = h.kernel.events.at(-1)?.cardAnswer?.answers?.[0];

    const typed = await type(h, '2');
    const parsed = parseCardReply(questionOf(SINGLE), typed);

    assert.deepEqual(clicked, { questionId: 'termin', kind: 'single_choice', optionId: 'b' });
    assert.ok(parsed.ok);
    assert.deepEqual(parsed.ok ? parsed.value : undefined, clicked);
  });

  const typedForms: [string, CardQuestion, string, unknown][] = [
    ['multiple choice "1,3"', { id: 'm', label: 'Wann?', kind: 'multiple_choice', options: OPTIONS }, '1,3', { questionId: 'm', kind: 'multiple_choice', optionIds: ['a', 'c'] }],
    ['ranking "3, 1, 2"', { id: 'r', label: 'Reihenfolge', kind: 'ranking', options: OPTIONS }, '3, 1, 2', { questionId: 'r', kind: 'ranking', ranking: ['c', 'a', 'b'] }],
    ['an option by its label', { id: 's', label: 'Termin', kind: 'single_choice', options: OPTIONS }, 'Freitag', { questionId: 's', kind: 'single_choice', optionId: 'c' }],
    ['consent: agreeing', { id: 'k', label: 'Einverstanden?', kind: 'consent' }, 'Ja', { questionId: 'k', kind: 'consent', decision: 'agree' }],
    ['consent: an objection with its reason', { id: 'k', label: 'Einverstanden?', kind: 'consent' }, 'Einwand: zu teuer', { questionId: 'k', kind: 'consent', decision: 'object', reason: 'zu teuer' }],
  ];
  for (const [name, question, reply, expected] of typedForms) {
    it(`${name} reaches the engine as typed and parses to the typed answer`, async () => {
      const h = harness();

      const parsed = parseCardReply(question, await type(h, reply));

      assert.ok(parsed.ok, JSON.stringify(parsed));
      assert.deepEqual(parsed.ok ? parsed.value : undefined, expected);
    });
  }

  it('a reply that does not fit is refused by the parser, not guessed (the engine asks again in words)', async () => {
    const h = harness();

    const parsed = parseCardReply(questionOf(SINGLE), await type(h, 'vielleicht'));

    assert.equal(parsed.ok, false);
  });
});

describe('criterion 30: a contribution to a facilitation is plain text in Teams', () => {
  it('a typed contribution is an event with the author Teams names and the text as typed, never a turn', async () => {
    const h = harness();

    const text = await type(h, 'Ich schlage vor, den Termin auf Mittwoch zu legen – dann sind alle da.');

    assert.equal(text, 'Ich schlage vor, den Termin auf Mittwoch zu legen – dann sind alle da.');
  });
});

/**
 * The kernel turns a typed reply into the `card_answer` of an open question it
 * posted as text (kernel #2086): by `replyToMessageId` when the event names the
 * message it answers, else, for a message that addresses the bot, when exactly
 * one question is open. What the channel owes that: the id it reports for a
 * sent text is the Teams activity id, `replyToMessageId` names a message only
 * when the person really answered one, and `addressedToBot` is set for a
 * mention. A wrong `replyToMessageId` is worse than none: it switches the
 * fallback off.
 */
describe('the kernel reads a typed reply to a text ballot (#2086)', () => {
  const sendBallot = async (h: ReturnType<typeof harness>, conversationId = GROUP_CHAT_ID) => {
    const outcome = await h.provider.sendToConversation(conversationId, { text: renderCardAsText(COUNT) }, { asChannelKey: MIRA.key });
    assert.equal(outcome.outcome, 'delivered');
    return outcome as { messageId?: string };
  };

  it('the id reported for a sent text is the id of the Teams activity', async () => {
    const h = harness();

    const outcome = await sendBallot(h);

    assert.ok(outcome.messageId);
    assert.equal(outcome.messageId, h.connector.sent.at(-1)?.id);
  });

  it('a quoted reply to the ballot names that message, and the person is the author', async () => {
    const h = harness();
    const { messageId } = await sendBallot(h);

    await h.sim.deliver(
      quotedReplyActivity({ bot: MIRA, from: BERT, text: '12', quotedMessageId: messageId ?? '', quotedAuthor: { id: MIRA.key, aadObjectId: '', name: MIRA.name }, quotedText: 'Wie viele Plätze brauchen wir?' }),
    );

    const event = h.kernel.events.at(-1);
    assert.equal(event?.replyToMessageId, messageId);
    assert.equal(event?.quoted, true);
    assert.equal(event?.text, '12');
    assert.equal(event?.author.id, BERT.aadObjectId);
  });

  it('a mention is addressedToBot; a plain message in the chat is not', async () => {
    const h = harness();

    await h.sim.deliver(messageActivity({ bot: MIRA, from: BERT, text: '12', mention: true }));
    await h.sim.deliver(messageActivity({ bot: MIRA, from: BERT, text: '13' }));

    assert.deepEqual(h.kernel.events.map((e) => e.addressedToBot === true), [true, false]);
  });

  it('a message in a channel thread carries no replyToMessageId just because every message of the thread has the root as replyToId', async () => {
    const h = harness();
    const rootId = '1731930000100';
    const conversationId = channelThreadId(rootId);
    h.kernel.subscribe(conversationId, MIRA);

    await h.sim.deliver({
      ...messageActivity({ bot: MIRA, from: BERT, text: '12', mention: true, conversationType: 'channel', conversationId }),
      replyToId: rootId,
    });

    const event = h.kernel.events.at(-1);
    assert.equal(event?.replyToMessageId, undefined, 'the root of the thread is not the question: the kernel would not try its one-open-question fallback');
    assert.equal(event?.addressedToBot, true);
  });

  it('a real reply in a channel thread (a quote of another message) still names that message', async () => {
    const h = harness();
    const rootId = '1731930000100';
    const conversationId = channelThreadId(rootId);
    h.kernel.subscribe(conversationId, MIRA);

    await h.sim.deliver({
      ...quotedReplyActivity({ bot: MIRA, from: BERT, text: '12', conversationType: 'channel', conversationId, quotedMessageId: '1731930000200', quotedAuthor: { id: MIRA.key, aadObjectId: '', name: MIRA.name }, quotedText: 'Frage' }),
      replyToId: rootId,
    });

    assert.equal(h.kernel.events.at(-1)?.replyToMessageId, '1731930000200');
  });
});

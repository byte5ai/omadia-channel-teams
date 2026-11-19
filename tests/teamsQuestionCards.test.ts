import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  createTeamsConversationSendAdapter,
  QUESTION_CARD_ANSWER_VERB,
  QUESTION_CARD_CHANGE_VERB,
  QUESTION_CARD_VIEW_VERB,
  TEAMS_CARD_CAPABILITY,
  TeamsConversationReferenceCache,
  TeamsQuestionCards,
} from '@omadia/channel-teams';

import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import type { QuestionCard } from '../src/kernel-types.js';
import { readCardAnswerFields } from './support/kernelCardAnswers.js';

import {
  ANNA,
  BERT,
  CARLA,
  cardExecuteActivity,
  GROUP_CHAT_ID,
  JONAS,
  messageActivity,
  MIRA,
  OTHER_GROUP_CHAT_ID,
  SimulatedConnector,
  SimulatedKernel,
  simulatedBot,
  type SimBot,
  type SimUser,
} from './support/teamsSimulation.js';

/**
 * byte5ai/omadia#2043 §6.1, #2089 — QUESTION CARDS.
 *
 * Votes and structured input as Adaptive Cards: one card per step, the
 * person's own answer shown only to them, one update when the step closes.
 * The real bot, the real send adapter and the real card registry run here
 * over the simulated Bot Framework transport.
 *
 * NOT VERIFIED BY THESE TESTS: anything a Teams client does. A click is
 * built here as the documented `adaptiveCard/action` invoke with the inputs
 * merged into the action's data. Which checks a client runs before it sends
 * (`isRequired`, `maxLength`), how it shows a returned card, and what a guest
 * or anonymous person's click carries is on the live checklist
 * (`docs/facilitator-live-verification.md`, "Question cards"). That is why
 * every rule is enforced on the server and tested with clicks no well-behaved
 * client would send.
 */

type Json = Record<string, unknown>;

const NOW = Date.parse('2026-11-19T10:00:00.000Z');

function harness(opts: { subscribed?: boolean } = {}) {
  const clock = { now: NOW };
  const kernel = new SimulatedKernel();
  if (opts.subscribed !== false) kernel.subscribe(GROUP_CHAT_ID, MIRA);
  const refs = new TeamsConversationReferenceCache();
  const sim = simulatedBot(kernel, {
    captureConversationReference: (context) => refs.capture(context),
    emitMembershipEvent: () => {},
  });
  const cards = new TeamsQuestionCards({ now: () => clock.now });
  sim.bot.attachQuestionCards(cards);
  const connector = new SimulatedConnector();
  for (const bot of [MIRA, JONAS]) {
    refs.capture({ activity: messageActivity({ bot, from: ANNA, text: 'hi' }) } as unknown as TurnContext);
  }
  const provider = createTeamsConversationSendAdapter({ refs, sendProactive: connector.sendProactive, cards });

  /** Post a card as Mira; returns the posted Adaptive Card and its message id. */
  async function post(card: QuestionCard, text = ''): Promise<{ content: Json; messageId: string }> {
    const outcome = await provider.sendToConversation(GROUP_CHAT_ID, { text, card } as never, { asChannelKey: MIRA.key });
    assert.equal(outcome.outcome, 'delivered', JSON.stringify(outcome));
    const sent = connector.sent.at(-1);
    const attachment = (sent?.activity['attachments'] as { content: Json }[] | undefined)?.[0];
    assert.ok(attachment, 'a card was posted');
    return { content: attachment.content, messageId: (outcome as { cardMessageId?: string }).cardMessageId ?? '' };
  }

  /** The data of an action of a card, as a client sends it back. */
  function actionData(content: Json, title?: string): Json {
    const actions = (content['actions'] as { title?: string; data: Json }[] | undefined) ?? [];
    const action = title ? actions.find((candidate) => candidate.title === title) : actions.at(-1);
    assert.ok(action, `action ${title ?? '(last)'} not on the card`);
    return action.data;
  }

  /** One click. Returns the invoke response body. */
  async function click(
    from: SimUser,
    messageId: string,
    data: Json,
    extra: { verb?: string; conversationId?: string; anonymous?: boolean; bot?: SimBot } = {},
  ): Promise<{ type: string; value: unknown }> {
    const before = sim.sent.length;
    await sim.deliver(
      cardExecuteActivity({
        bot: extra.bot ?? MIRA,
        from,
        conversationId: extra.conversationId ?? GROUP_CHAT_ID,
        cardMessageId: messageId,
        verb: extra.verb ?? QUESTION_CARD_ANSWER_VERB,
        data,
        ...(extra.anonymous ? { anonymous: true } : {}),
      }),
    );
    const response = (sim.sent.slice(before) as Json[]).find((activity) => activity['type'] === 'invokeResponse');
    assert.ok(response, 'the click was answered');
    return (response['value'] as { body: { type: string; value: unknown } }).body;
  }

  return { clock, kernel, sim, cards, connector, provider, refs, post, actionData, click };
}

const answersOf = (kernel: SimulatedKernel): unknown[] => kernel.events.map((event) => event.cardAnswer?.answers);
const isCard = (body: { type: string }): boolean => body.type === 'application/vnd.microsoft.card.adaptive';
const textOf = (body: { value: unknown }): string => JSON.stringify(body.value);

const SINGLE: QuestionCard = {
  cardId: 'c-single',
  proposalId: 'p-4',
  proposalVersion: 2,
  title: 'Welchen Termin nehmen wir?',
  closesAt: '2026-11-20T16:00:00.000Z',
  questions: [
    {
      id: 'termin',
      label: 'Termin',
      kind: 'single_choice',
      required: true,
      options: [
        { id: 'a', label: 'Dienstag' },
        { id: 'b', label: 'Mittwoch' },
        { id: 'c', label: 'Freitag' },
      ],
    },
  ],
};

const OPTIONS = [
  { id: 'a', label: 'Eule' },
  { id: 'b', label: 'Fuchs' },
  { id: 'c', label: 'Otter' },
];

describe('#2089 question cards: what is posted and declared', () => {
  it('posts an Adaptive Card 1.5 with Action.Execute, a refresh without userIds and one AI label', async () => {
    const h = harness();

    const { content, messageId } = await h.post(SINGLE, 'Bitte stimmt ab.');

    assert.equal(content['version'], '1.5');
    assert.equal((content['actions'] as Json[]).every((action) => action['type'] === 'Action.Execute'), true);
    const refresh = content['refresh'] as { action: Json; userIds?: unknown };
    assert.equal(refresh.action['verb'], QUESTION_CARD_VIEW_VERB);
    assert.equal('userIds' in refresh, false, 'Teams refreshes automatically only up to 60 members');
    assert.ok(messageId.length > 0, 'the card message id is returned for the later update');
    // Text and card are two messages, each with exactly one AI label.
    assert.equal(h.connector.sent.length, 2);
    for (const sent of h.connector.sent) {
      const labels = (sent.activity['entities'] as { additionalType?: string[] }[]).filter((entity) =>
        entity.additionalType?.includes('AIGeneratedContent'),
      );
      assert.equal(labels.length, 1);
    }
    assert.ok(JSON.stringify(content).includes('Offen bis'), 'the closing time is on the card');
  });

  it('declares exactly what it renders, and nothing without the card registry', () => {
    const h = harness();
    assert.deepEqual(h.provider.supports.card, TEAMS_CARD_CAPABILITY);
    assert.deepEqual(h.provider.supports.card?.kinds, ['single_choice', 'multiple_choice', 'ranking', 'consent', 'assignment', 'text']);
    assert.equal(h.provider.supports.card?.maxQuestions, 5);
    assert.equal(h.provider.supports.card?.update, true);

    const plain = createTeamsConversationSendAdapter({ refs: h.refs, sendProactive: h.connector.sendProactive });
    assert.equal(plain.supports.card, undefined);
  });

  it('refuses a card it cannot render, before anything is posted', async () => {
    const h = harness();
    const send = (card: unknown, extra: Json = {}) =>
      h.provider.sendToConversation(GROUP_CHAT_ID, { text: 'x', card, ...extra } as never, { asChannelKey: MIRA.key });
    const refused: [string, unknown, Json?][] = [
      ['a kind that is not built', { ...SINGLE, questions: [{ id: 'n', label: 'Wie viele?', kind: 'number' }] }],
      ['six questions', { ...SINGLE, questions: Array.from({ length: 6 }, (_, n) => ({ ...SINGLE.questions[0], id: `q${String(n)}` })) }],
      ['one option', { ...SINGLE, questions: [{ ...SINGLE.questions[0], options: [OPTIONS[0]] }] }],
      ['an option id used twice', { ...SINGLE, questions: [{ ...SINGLE.questions[0], options: [OPTIONS[0], OPTIONS[0]] }] }],
      ['no card id', { ...SINGLE, cardId: '' }],
      ['an id with a character the contract does not allow', { ...SINGLE, questions: [{ ...SINGLE.questions[0], options: [{ id: 'a,b', label: 'A' }, OPTIONS[1]] }] }],
      ['a dotted question id', { ...SINGLE, questions: [{ ...SINGLE.questions[0], id: 'q.rank1' }] }],
      ['a poll and a card', SINGLE, { poll: { pollId: 'p', question: 'q', options: OPTIONS } }],
    ];
    for (const [name, card, extra] of refused) {
      const outcome = await send(card, extra);
      assert.equal(outcome.outcome, 'unreachable', name);
      assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'channel_error', name);
    }
    assert.deepEqual(h.connector.sent, [], 'nothing was posted, not even the text');
  });

  it('refuses a card when the registry is not wired', async () => {
    const h = harness();
    const plain = createTeamsConversationSendAdapter({ refs: h.refs, sendProactive: h.connector.sendProactive });
    const outcome = await plain.sendToConversation(GROUP_CHAT_ID, { text: 'x', card: SINGLE } as never, { asChannelKey: MIRA.key });
    assert.equal(outcome.outcome, 'unreachable');
    assert.deepEqual(h.connector.sent, []);
  });
});

describe('#2089 question cards: the six forms', () => {
  it('single choice: a typed answer with proposal id and version; only the clicker sees their answer', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);

    const body = await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'b' });

    assert.deepEqual(
      h.kernel.events.map((e) => [e.kind, e.messageId, e.author.id, e.cardAnswer?.cardId, e.cardAnswer?.proposalId, e.cardAnswer?.proposalVersion, e.cardAnswer?.answererVerified]),
      [['card_answer', messageId, ANNA.aadObjectId, 'c-single', 'p-4', 2, true]],
    );
    assert.deepEqual(answersOf(h.kernel), [[{ questionId: 'termin', kind: 'single_choice', optionId: 'b' }]]);
    assert.deepEqual(h.kernel.events[0]?.cardAnswer?.value, { answers: [{ questionId: 'termin', kind: 'single_choice', optionId: 'b' }] });
    // The response to the click is a card for Anna alone.
    assert.ok(isCard(body));
    assert.ok(textOf(body).includes('"text":"Deine Antwort"') && textOf(body).includes('Mittwoch'));
    assert.ok(textOf(body).includes('"title":"Ändern"'));
    // The posted message was not touched: the group sees the question card,
    // no names, no tally.
    assert.deepEqual(h.connector.updated, []);
    assert.equal(h.connector.sent.length, 1);
    assert.deepEqual(h.sim.turns, [], 'a click is never a turn');
  });

  it('multiple choice: several options; the minimum and maximum are checked on the server', async () => {
    const h = harness();
    const card: QuestionCard = {
      cardId: 'c-multi',
      title: 'Welche Themen?',
      questions: [{ id: 'themen', label: 'Themen', kind: 'multiple_choice', options: OPTIONS, min: 1, max: 2 }],
    };
    const { content, messageId } = await h.post(card);
    const data = h.actionData(content);

    const tooMany = await h.click(ANNA, messageId, { ...data, q0: 'a,b,c' });
    const none = await h.click(ANNA, messageId, { ...data });
    assert.ok(textOf(tooMany).includes('höchstens 2'));
    assert.ok(textOf(none).includes('Bitte eine Option wählen'));
    assert.deepEqual(h.kernel.events, [], 'a refused submission emits nothing');

    await h.click(ANNA, messageId, { ...data, q0: 'a,c' });
    assert.deepEqual(answersOf(h.kernel), [[{ questionId: 'themen', kind: 'multiple_choice', optionIds: ['a', 'c'] }]]);
  });

  it('ranked vote: one dropdown per rank; a rank used twice comes back with the values kept and the reason', async () => {
    const h = harness();
    const card: QuestionCard = {
      cardId: 'c-rank',
      title: 'Reihenfolge der drei Vorschläge',
      questions: [{ id: 'rang', label: 'Rangfolge', kind: 'ranking', required: true, options: OPTIONS, rankLabels: ['1. Platz (3 Punkte)', '2. Platz (2 Punkte)', '3. Platz (1 Punkt)'] }],
    };
    const { content, messageId } = await h.post(card);
    const data = h.actionData(content);
    const dropdowns = (content['body'] as Json[]).filter((element) => element['type'] === 'Input.ChoiceSet');
    assert.deepEqual(dropdowns.map((d) => [d['id'], d['label'], d['style']]), [
      ['q0.r1', '1. Platz (3 Punkte)', 'compact'],
      ['q0.r2', '2. Platz (2 Punkte)', 'compact'],
      ['q0.r3', '3. Platz (1 Punkt)', 'compact'],
    ]);

    const duplicate = await h.click(ANNA, messageId, { ...data, 'q0.r1': 'a', 'q0.r2': 'c', 'q0.r3': 'a' });

    assert.ok(isCard(duplicate), 'the same card again, for Anna only');
    assert.ok(textOf(duplicate).includes('Jede Option darf nur einmal vorkommen: „Eule“ steht auf Platz 1 und 3.'));
    const kept = ((duplicate.value as Json)['body'] as Json[]).filter((element) => element['type'] === 'Input.ChoiceSet');
    assert.deepEqual(kept.map((d) => d['value']), ['a', 'c', 'a'], 'what she entered is still there');
    assert.deepEqual(h.kernel.events, []);

    const incomplete = await h.click(ANNA, messageId, { ...data, 'q0.r1': 'a' });
    assert.ok(textOf(incomplete).includes('Bitte jeden Platz vergeben'));
    assert.deepEqual(h.kernel.events, []);

    const ok = await h.click(ANNA, messageId, { ...data, 'q0.r1': 'a', 'q0.r2': 'c', 'q0.r3': 'b' });
    assert.deepEqual(answersOf(h.kernel), [[{ questionId: 'rang', kind: 'ranking', ranking: ['a', 'c', 'b'] }]]);
    assert.ok(textOf(ok).includes('1. Eule, 2. Otter, 3. Fuchs'));
  });

  it('consent: agreeing needs nothing more; an objection needs a reason', async () => {
    const h = harness();
    const card: QuestionCard = {
      cardId: 'c-consent',
      proposalId: 'p-9',
      proposalVersion: 1,
      title: 'Vorschlag: Wir starten am 1.12. mit Variante B.',
      questions: [{ id: 'konsent', label: 'Deine Haltung', kind: 'consent', required: true }],
    };
    const { content, messageId } = await h.post(card);
    const data = h.actionData(content);

    const noReason = await h.click(BERT, messageId, { ...data, q0: 'object' });
    assert.ok(textOf(noReason).includes('Ein Einwand braucht eine Begründung.'));
    assert.deepEqual(h.kernel.events, []);

    await h.click(ANNA, messageId, { ...data, q0: 'agree' });
    await h.click(BERT, messageId, { ...data, q0: 'object', 'q0.reason': '  Der Termin kollidiert mit dem Release.  ' });

    assert.deepEqual(answersOf(h.kernel), [
      [{ questionId: 'konsent', kind: 'consent', decision: 'agree' }],
      [{ questionId: 'konsent', kind: 'consent', decision: 'object', reason: 'Der Termin kollidiert mit dem Release.' }],
    ]);
    assert.deepEqual(h.kernel.events.map((e) => e.author.displayName), [ANNA.name, BERT.name]);
  });

  it('assignment: "Ich übernehme das" is the clicker; the people picker names a person by Entra object id', async () => {
    const h = harness();
    const card: QuestionCard = {
      cardId: 'c-assign',
      title: 'Wer übernimmt „Angebot einholen“ bis 28.11.?',
      questions: [{ id: 'wer', label: 'Person', kind: 'assignment', candidates: 'conversation_members' }],
    };
    const { content, messageId } = await h.post(card);
    const picker = (content['body'] as Json[]).find((element) => element['type'] === 'Input.ChoiceSet');
    assert.deepEqual(picker?.['choices.data'], { type: 'Data.Query', dataset: 'graph.microsoft.com/users?scope=currentContext' });

    const volunteer = await h.click(CARLA, messageId, h.actionData(content, 'Ich übernehme das'));
    await h.click(ANNA, messageId, { ...h.actionData(content, 'Vorschlag senden'), q0: BERT.aadObjectId.toUpperCase() });
    const garbage = await h.click(ANNA, messageId, { ...h.actionData(content, 'Vorschlag senden'), q0: '../users/all' });

    assert.deepEqual(answersOf(h.kernel), [
      [{ questionId: 'wer', kind: 'assignment', assignee: 'self' }],
      [{ questionId: 'wer', kind: 'assignment', assignee: { kind: 'teams-aad', id: BERT.aadObjectId } }],
    ]);
    assert.equal(h.kernel.events[0]?.author.id, CARLA.aadObjectId, 'the volunteer is who Teams says clicked');
    assert.ok(textOf(volunteer).includes('ich selbst'));
    assert.ok(textOf(garbage).includes('Diese Person konnte nicht zugeordnet werden.'));
  });

  it('assignment from a given list: only a listed candidate', async () => {
    const h = harness();
    const card: QuestionCard = {
      cardId: 'c-assign-list',
      title: 'Wer moderiert das nächste Treffen?',
      questions: [{ id: 'wer', label: 'Person', kind: 'assignment', allowVolunteer: false, candidates: [{ id: 'u1', label: 'Anna Köhler' }, { id: 'u2', label: 'Bert Maier' }, { id: BERT.aadObjectId, label: 'Bert Maier' }] }],
    };
    const { content, messageId } = await h.post(card);
    assert.equal((content['actions'] as Json[]).length, 1, 'no volunteer button when the card does not allow it');

    await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'u2' });
    // In the contract the ids of a fixed list are channel user ids: an Entra
    // object id is passed on as that person.
    await h.click(ANNA, messageId, { ...h.actionData(content), q0: BERT.aadObjectId });
    const forgedVolunteer = await h.click(ANNA, messageId, { ...h.actionData(content), preset: { q0: 'self' } });

    assert.deepEqual(answersOf(h.kernel), [
      [{ questionId: 'wer', kind: 'assignment', assignee: { kind: 'option', id: 'u2', displayName: 'Bert Maier' } }],
      [{ questionId: 'wer', kind: 'assignment', assignee: { kind: 'teams-aad', id: BERT.aadObjectId, displayName: 'Bert Maier' } }],
    ]);
    assert.ok(isCard(forgedVolunteer) && h.kernel.events.length === 2, 'a volunteer click the card never offered is refused');
  });

  it('free text, append: every submission is one more contribution, the original text', async () => {
    const h = harness();
    const card: QuestionCard = {
      cardId: 'c-ideas',
      title: 'Welche Themen sollen auf die Agenda?',
      questions: [{ id: 'idee', label: 'Deine Idee (eine pro Absenden)', kind: 'text', required: true, mode: 'append', maxLength: 200 }],
    };
    const { content, messageId } = await h.post(card);
    const data = h.actionData(content);

    await h.click(ANNA, messageId, { ...data, q0: 'Budget 2027 <b>fett</b>' });
    const second = await h.click(ANNA, messageId, { ...data, q0: 'Onboarding' });
    const empty = await h.click(ANNA, messageId, { ...data, q0: '   ' });
    // `maxLength` on the input is a client-side hint; the server enforces it.
    const tooLong = await h.click(ANNA, messageId, { ...data, q0: 'x'.repeat(201) });

    assert.deepEqual(answersOf(h.kernel), [
      [{ questionId: 'idee', kind: 'text', text: 'Budget 2027 <b>fett</b>' }],
      [{ questionId: 'idee', kind: 'text', text: 'Onboarding' }],
    ]);
    assert.ok(textOf(second).includes('Deine Beiträge (2)') && textOf(second).includes('Weiteren Beitrag senden'));
    assert.ok(textOf(empty).includes('Bitte ausfüllen.'));
    assert.ok(textOf(tooLong).includes('höchstens 200 Zeichen'));
    const again = await h.click(ANNA, messageId, data, { verb: QUESTION_CARD_CHANGE_VERB });
    const field = ((again.value as Json)['body'] as Json[]).find((element) => element['type'] === 'Input.Text');
    assert.equal(field?.['value'], undefined, 'the field is empty for the next idea');
  });

  it('free text, replace: the newest statement stands for the person\'s view', async () => {
    const h = harness();
    const card: QuestionCard = {
      cardId: 'c-statement',
      title: 'Was spricht dagegen?',
      questions: [{ id: 'grund', label: 'Dein Einwand', kind: 'text', required: true }],
    };
    const { content, messageId } = await h.post(card);
    const data = h.actionData(content);

    await h.click(ANNA, messageId, { ...data, q0: 'Zu teuer.' });
    const second = await h.click(ANNA, messageId, { ...data, q0: 'Zu teuer und zu spät.' });

    assert.equal(h.kernel.events.length, 2, 'each submission is an event; the engine lets the newest win');
    assert.ok(textOf(second).includes('Zu teuer und zu spät.') && !textOf(second).includes('Deine Beiträge'));
  });

  it('questionnaire: all questions on one card, one submission is one event', async () => {
    const h = harness();
    const card: QuestionCard = {
      cardId: 'c-retro',
      title: 'Kurzer Rückblick',
      intro: 'Drei Fragen, eine Minute.',
      questions: [
        { id: 'stimmung', label: 'Wie war die Woche?', kind: 'single_choice', required: true, options: OPTIONS },
        { id: 'themen', label: 'Was hat geholfen?', kind: 'multiple_choice', required: false, options: OPTIONS },
        { id: 'wunsch', label: 'Was wünschst du dir?', kind: 'text', required: false, maxLength: 500 },
      ],
    };
    const { content, messageId } = await h.post(card);
    const data = h.actionData(content);

    // The first question is missing: the card comes back with the other
    // answers kept.
    const missing = await h.click(ANNA, messageId, { ...data, q1: 'a,b', q2: 'Mehr Fokuszeit' });
    assert.ok(textOf(missing).includes('Bitte prüfe die markierten Angaben.'));
    const keptText = ((missing.value as Json)['body'] as Json[]).find((element) => element['id'] === 'q2');
    assert.equal(keptText?.['value'], 'Mehr Fokuszeit');
    assert.deepEqual(h.kernel.events, []);

    const ok = await h.click(ANNA, messageId, { ...data, q0: 'c', q1: 'a,b', q2: 'Mehr Fokuszeit' });

    assert.deepEqual(answersOf(h.kernel), [
      [
        { questionId: 'stimmung', kind: 'single_choice', optionId: 'c' },
        { questionId: 'themen', kind: 'multiple_choice', optionIds: ['a', 'b'] },
        { questionId: 'wunsch', kind: 'text', text: 'Mehr Fokuszeit' },
      ],
    ]);
    assert.ok(textOf(ok).includes('Wie war die Woche?: Otter') && textOf(ok).includes('Was wünschst du dir?: Mehr Fokuszeit'));

    // An optional question left empty is simply not in the answer.
    await h.click(BERT, messageId, { ...data, q0: 'a' });
    assert.deepEqual(answersOf(h.kernel)[1], [{ questionId: 'stimmung', kind: 'single_choice', optionId: 'a' }]);
  });
});

describe('#2089 question cards: the person\'s own view', () => {
  it('"Ändern" brings the inputs back with the answer filled in; the refresh shows each person their own state', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);
    const data = h.actionData(content);
    await h.click(ANNA, messageId, { ...data, q0: 'b' });

    const change = await h.click(ANNA, messageId, data, { verb: QUESTION_CARD_CHANGE_VERB });
    const annaRefresh = await h.click(ANNA, messageId, data, { verb: QUESTION_CARD_VIEW_VERB });
    const bertRefresh = await h.click(BERT, messageId, data, { verb: QUESTION_CARD_VIEW_VERB });

    const input = ((change.value as Json)['body'] as Json[]).find((element) => element['type'] === 'Input.ChoiceSet');
    assert.equal(input?.['value'], 'b');
    const OWN_HEADING = '"text":"Deine Antwort"';
    assert.ok(textOf(annaRefresh).includes(OWN_HEADING) && textOf(annaRefresh).includes('Mittwoch'));
    assert.ok(!textOf(bertRefresh).includes(OWN_HEADING) && textOf(bertRefresh).includes('Input.ChoiceSet'), 'Bert has not answered: he sees the question');
    assert.ok(!textOf(bertRefresh).includes('Anna'), 'and nothing about Anna');
    assert.equal(h.kernel.events.length, 1, 'neither "Ändern" nor a refresh is an answer');
  });

  it('a card that is not changeable offers no "Ändern"', async () => {
    const h = harness();
    const { content, messageId } = await h.post({ ...SINGLE, cardId: 'c-final', changeable: false });
    const body = await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'a' });
    assert.ok(isCard(body) && !textOf(body).includes('"title":"Ändern"'));
    assert.ok(!JSON.stringify(content).includes('bis zum Ende ändern'));
  });
});

describe('#2089 question cards: clicks a client could forge', () => {
  it('an option that is not on the card', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);
    const body = await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'z' });
    assert.ok(textOf(body).includes('Diese Option gibt es auf dieser Karte nicht.'));
    assert.deepEqual(h.kernel.events, []);
  });

  it('a card in the click that differs from the one that was posted: the posted one counts', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);
    const forgedCard = { ...SINGLE, questions: [{ ...SINGLE.questions[0], options: [...OPTIONS, { id: 'z', label: 'Erfunden' }] }] };

    const body = await h.click(ANNA, messageId, { ...h.actionData(content), card: forgedCard, q0: 'z' });

    assert.ok(textOf(body).includes('Diese Option gibt es auf dieser Karte nicht.'));
    assert.deepEqual(h.kernel.events, []);
  });

  it('an answer to an earlier version of the proposal: the current card comes back, nothing is recorded', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);
    const oldData = h.actionData(content);
    // The engine corrects the card: version 3, another option text.
    const corrected: QuestionCard = { ...SINGLE, proposalVersion: 3, title: 'Welchen Termin nehmen wir? (korrigiert)' };
    assert.deepEqual(await h.provider.updateCard(GROUP_CHAT_ID, messageId, { state: 'open', card: corrected }, { asChannelKey: MIRA.key }), { outcome: 'updated' });

    const stale = await h.click(ANNA, messageId, { ...oldData, q0: 'a' });

    assert.ok(isCard(stale));
    assert.ok(textOf(stale).includes('Der Vorschlag wurde inzwischen geändert') && textOf(stale).includes('(korrigiert)'));
    assert.deepEqual(h.kernel.events, []);
    // An answer to the current version counts and carries version 3.
    const current = (h.connector.updated.at(-1)?.activity['attachments'] as { content: Json }[])[0]?.content as Json;
    await h.click(ANNA, messageId, { ...h.actionData(current), q0: 'a' });
    assert.deepEqual(h.kernel.events.map((e) => e.cardAnswer?.proposalVersion), [3]);
  });

  it('a click after the card was closed: the closed card, nothing is recorded', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);
    await h.provider.updateCard(
      GROUP_CHAT_ID,
      messageId,
      { state: 'closed', title: SINGLE.title, resultLines: ['Ergebnis: Mittwoch (3 von 4)'], closedAt: '2026-11-19T10:05:00.000Z' },
      { asChannelKey: MIRA.key },
    );

    const body = await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'a' });

    assert.ok(isCard(body) && textOf(body).includes('Ergebnis: Mittwoch (3 von 4)') && !textOf(body).includes('Action.Execute'));
    assert.deepEqual(h.kernel.events, []);
  });

  it('a click after the closing time, also when the close update has not arrived', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);
    h.clock.now = Date.parse('2026-11-20T16:00:01.000Z');

    const body = await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'a' });

    assert.equal(body.value, 'Diese Abstimmung ist nicht mehr aktiv.');
    assert.deepEqual(h.kernel.events, []);
  });

  it('a card clicked in another conversation than it was posted into', async () => {
    const h = harness();
    h.kernel.subscribe(OTHER_GROUP_CHAT_ID, MIRA);
    const { content, messageId } = await h.post(SINGLE);

    const body = await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'a' }, { conversationId: OTHER_GROUP_CHAT_ID });

    assert.match(String(body.value), /gehört nicht in diese Unterhaltung/);
    assert.deepEqual(h.kernel.events, []);
  });

  it('a click sent to another bot than the one that posted the card', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);

    // Jonas is in the chat too but holds no subscription and posted no card.
    const body = await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'a' }, { bot: JONAS });

    assert.equal(body.value, 'Diese Abstimmung ist nicht mehr aktiv.');
    assert.deepEqual(h.kernel.events, []);
  });

  it('a click when no facilitation holds the conversation', async () => {
    const h = harness({ subscribed: false });
    const { content, messageId } = await h.post(SINGLE);
    const body = await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'a' });
    assert.equal(body.value, 'Diese Abstimmung ist nicht mehr aktiv.');
    assert.deepEqual(h.kernel.events, []);
  });

  it('a click of somebody Teams does not identify is passed on as unverified, and told so', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);

    const body = await h.click(CARLA, messageId, { ...h.actionData(content), q0: 'a' }, { anonymous: true });

    assert.deepEqual(h.kernel.events.map((e) => e.cardAnswer?.answererVerified), [false]);
    assert.match(String(body.value), /nicht mitgeteilt, wer du bist/);
  });

  it('an author named in the click data is ignored: the author is who Teams says clicked', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);

    await h.click(BERT, messageId, { ...h.actionData(content), q0: 'a', from: { aadObjectId: ANNA.aadObjectId }, userId: ANNA.aadObjectId, author: ANNA.name });

    assert.deepEqual(h.kernel.events.map((e) => e.author.id), [BERT.aadObjectId]);
  });

  it('data that is not a question card, or unreadable', async () => {
    const h = harness();
    const { messageId } = await h.post(SINGLE);
    const broken = await h.click(ANNA, messageId, { type: 'facilitation.card', cardId: 'nope' });
    assert.match(String(broken.value), /konnte nicht gelesen werden/);
    assert.deepEqual(h.kernel.events, []);
  });

  it('after a restart the click is read against the card it carries; the engine checks it against what it sent', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);
    // A new process: the registry is empty, the subscription is still there.
    const fresh = new TeamsQuestionCards({ now: () => h.clock.now });
    h.sim.bot.attachQuestionCards(fresh);

    const ok = await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'b' });
    const unknown = await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'z' });

    assert.ok(isCard(ok) && textOf(ok).includes('Mittwoch'));
    assert.ok(textOf(unknown).includes('Diese Option gibt es auf dieser Karte nicht.'));
    assert.deepEqual(h.kernel.events.map((e) => [e.cardAnswer?.cardId, e.cardAnswer?.proposalVersion]), [['c-single', 2]]);
  });
});

describe('#2089 question cards: updateCard', () => {
  it('closed: replaces the posted message with the engine\'s result, no inputs, one AI label', async () => {
    const h = harness();
    const { content, messageId } = await h.post(SINGLE);
    await h.click(ANNA, messageId, { ...h.actionData(content), q0: 'b' });

    const outcome = await h.provider.updateCard(
      GROUP_CHAT_ID,
      messageId,
      { state: 'closed', title: SINGLE.title, resultLines: ['Ergebnis: Mittwoch', 'Dienstag 1, Mittwoch 3, Freitag 0'], closedAt: '2026-11-19T10:05:00.000Z' },
      { asChannelKey: MIRA.key },
    );

    assert.deepEqual(outcome, { outcome: 'updated' });
    assert.equal(h.connector.updated.length, 1);
    const update = h.connector.updated[0];
    assert.equal(update?.activity['id'], messageId, 'the stored activity id');
    assert.equal(update?.asBot, MIRA.key);
    const card = JSON.stringify(update?.activity['attachments']);
    assert.ok(card.includes('Ergebnis: Mittwoch') && card.includes('Geschlossen am'));
    assert.ok(!card.includes('Input.') && !card.includes('Action.Execute') && !card.includes('refresh'));
    assert.equal((update?.activity['entities'] as { additionalType?: string[] }[]).filter((e) => e.additionalType?.includes('AIGeneratedContent')).length, 1);
    // Anna's own view is gone: a refresh shows the closed card.
    const refresh = await h.click(ANNA, messageId, h.actionData(content), { verb: QUESTION_CARD_VIEW_VERB });
    assert.ok(textOf(refresh).includes('Ergebnis: Mittwoch') && !textOf(refresh).includes('"text":"Deine Antwort"'));
  });

  it('never updates as another bot, and reports a failed update', async () => {
    const h = harness();
    const { messageId } = await h.post(SINGLE);
    const closed = { state: 'closed' as const, title: 't', resultLines: ['r'], closedAt: '2026-11-19T10:05:00.000Z' };

    const notABot = await h.provider.updateCard(GROUP_CHAT_ID, messageId, closed, { asChannelKey: 'mira' });
    const elsewhere = await h.provider.updateCard('19:unknown@thread.v2', messageId, closed, { asChannelKey: MIRA.key });
    h.connector.updateFailure = Object.assign(new Error('The message was deleted.'), { statusCode: 404 });
    const failed = await h.provider.updateCard(GROUP_CHAT_ID, messageId, closed, { asChannelKey: MIRA.key });

    assert.deepEqual([notABot.outcome, elsewhere.outcome, failed.outcome], ['unreachable', 'unreachable', 'unreachable']);
    assert.equal(notABot.outcome === 'unreachable' ? notABot.code : '', 'not_permitted');
    assert.equal(elsewhere.outcome === 'unreachable' ? elsewhere.code : '', 'not_permitted', 'the card was posted into another conversation');
    assert.match(failed.outcome === 'unreachable' ? failed.message : '', /The message was deleted\./);
    assert.deepEqual(h.connector.updated, []);
  });

  it('updates only a question card that the same bot posted: not another bot\'s card, not a text message', async () => {
    // The kernel keeps no record of posted cards; it checks identity and
    // subscription. Which message is a card of which bot is known only here.
    const h = harness();
    const closed = { state: 'closed' as const, title: 't', resultLines: ['r'], closedAt: '2026-11-19T10:05:00.000Z' };
    const open = { state: 'open' as const, card: { ...SINGLE, proposalVersion: 3 } };

    // Mira posts a text and a card; Jonas posts a card of his own.
    await h.post(SINGLE, 'Bitte stimmt ab.');
    const miraText = h.connector.sent[0]?.id ?? '';
    const miraCard = h.connector.sent[1]?.id ?? '';
    const jonas = await h.provider.sendToConversation(GROUP_CHAT_ID, { text: '', card: { ...SINGLE, cardId: 'c-jonas' } } as never, { asChannelKey: JONAS.key });
    const jonasCard = jonas.outcome === 'delivered' ? (jonas.cardMessageId ?? '') : '';
    assert.ok(miraText.length > 0 && miraCard.length > 0 && jonasCard.length > 0);

    const refused = [
      await h.provider.updateCard(GROUP_CHAT_ID, jonasCard, closed, { asChannelKey: MIRA.key }),
      await h.provider.updateCard(GROUP_CHAT_ID, jonasCard, open, { asChannelKey: MIRA.key }),
      await h.provider.updateCard(GROUP_CHAT_ID, miraCard, closed, { asChannelKey: JONAS.key }),
      await h.provider.updateCard(GROUP_CHAT_ID, miraText, closed, { asChannelKey: MIRA.key }),
      await h.provider.updateCard(GROUP_CHAT_ID, miraText, open, { asChannelKey: MIRA.key }),
      await h.provider.updateCard(GROUP_CHAT_ID, '1700000000000', closed, { asChannelKey: MIRA.key }),
    ];

    for (const outcome of refused) {
      assert.equal(outcome.outcome, 'unreachable');
      assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'not_permitted');
      assert.match(outcome.outcome === 'unreachable' ? outcome.message : '', /is not a question card that bot/);
    }
    assert.deepEqual(h.connector.updated, [], 'no message was touched');

    // Each bot can still update its own card.
    assert.deepEqual(await h.provider.updateCard(GROUP_CHAT_ID, miraCard, closed, { asChannelKey: MIRA.key }), { outcome: 'updated' });
    assert.deepEqual(await h.provider.updateCard(GROUP_CHAT_ID, jonasCard, closed, { asChannelKey: JONAS.key }), { outcome: 'updated' });
    assert.deepEqual(h.connector.updated.map((update) => [update.asBot, update.id]), [[MIRA.key, miraCard], [JONAS.key, jonasCard]]);
  });

  it('a corrected card keeps its id', async () => {
    const h = harness();
    const { messageId } = await h.post(SINGLE);
    const outcome = await h.provider.updateCard(GROUP_CHAT_ID, messageId, { state: 'open', card: { ...SINGLE, cardId: 'c-other' } }, { asChannelKey: MIRA.key });
    assert.equal(outcome.outcome, 'unreachable');
    assert.match(outcome.outcome === 'unreachable' ? outcome.message : '', /a corrected card keeps its id/);
    assert.deepEqual(h.connector.updated, []);
  });

  it('after a restart the registry is empty: an update is refused and the engine posts its result as text', async () => {
    const h = harness();
    const { messageId } = await h.post(SINGLE);
    const restarted = createTeamsConversationSendAdapter({ refs: h.refs, sendProactive: h.connector.sendProactive, cards: new TeamsQuestionCards() });

    const outcome = await restarted.updateCard(
      GROUP_CHAT_ID,
      messageId,
      { state: 'closed', title: 't', resultLines: ['r'], closedAt: '2026-11-19T10:05:00.000Z' },
      { asChannelKey: MIRA.key },
    );

    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'not_permitted');
    assert.deepEqual(h.connector.updated, []);
  });

  it('a corrected open card is checked like a new one', async () => {
    const h = harness();
    const { messageId } = await h.post(SINGLE);
    const outcome = await h.provider.updateCard(GROUP_CHAT_ID, messageId, { state: 'open', card: { ...SINGLE, questions: [] } }, { asChannelKey: MIRA.key });
    assert.equal(outcome.outcome, 'unreachable');
    assert.deepEqual(h.connector.updated, []);
  });
});

describe('#2089 question cards: what is emitted passes the kernel\'s check', () => {
  // The kernel refuses the WHOLE event when `answers` is not in the
  // contract's form. `readCardAnswerFields` is the kernel's function, copied
  // (tests/support/kernelCardAnswers.ts).
  const forms: { name: string; card: QuestionCard; data: Json; from?: SimUser; action?: string }[] = [
    { name: 'single choice', card: SINGLE, data: { q0: 'b' } },
    {
      name: 'multiple choice',
      card: { cardId: 'k-multi', title: 't', questions: [{ id: 'q', label: 'l', kind: 'multiple_choice', options: OPTIONS }] },
      data: { q0: 'a,c' },
    },
    {
      name: 'ranking',
      card: { cardId: 'k-rank', title: 't', questions: [{ id: 'q', label: 'l', kind: 'ranking', options: OPTIONS }] },
      data: { 'q0.r1': 'c', 'q0.r2': 'a', 'q0.r3': 'b' },
    },
    {
      name: 'ranking of the top two',
      card: { cardId: 'k-rank2', title: 't', questions: [{ id: 'q', label: 'l', kind: 'ranking', options: OPTIONS, ranks: 2 }] },
      data: { 'q0.r1': 'c', 'q0.r2': 'a' },
    },
    { name: 'consent, agree', card: { cardId: 'k-agree', title: 't', questions: [{ id: 'q', label: 'l', kind: 'consent' }] }, data: { q0: 'agree' } },
    {
      name: 'consent, objection with a reason',
      card: { cardId: 'k-object', proposalId: 'p-1', proposalVersion: 0, title: 't', questions: [{ id: 'q', label: 'l', kind: 'consent' }] },
      data: { q0: 'object', 'q0.reason': 'Zu früh.' },
    },
    {
      name: 'assignment, volunteer',
      card: { cardId: 'k-self', title: 't', questions: [{ id: 'q', label: 'l', kind: 'assignment', candidates: 'conversation_members' }] },
      data: {},
      action: 'Ich übernehme das',
    },
    {
      name: 'assignment, people picker',
      card: { cardId: 'k-pick', title: 't', questions: [{ id: 'q', label: 'l', kind: 'assignment', candidates: 'conversation_members' }] },
      data: { q0: BERT.aadObjectId },
    },
    {
      name: 'assignment, fixed list',
      card: { cardId: 'k-list', title: 't', questions: [{ id: 'q', label: 'l', kind: 'assignment', candidates: [{ id: 'u1', label: 'Anna' }, { id: BERT.aadObjectId, label: 'Bert' }] }] },
      data: { q0: 'u1' },
    },
    {
      name: 'free text, the longest the adapter accepts',
      card: { cardId: 'k-text', title: 't', questions: [{ id: 'q', label: 'l', kind: 'text', maxLength: 4000 }] },
      data: { q0: 'ä'.repeat(4000) },
    },
    {
      name: 'questionnaire with five questions',
      card: {
        cardId: 'k-five',
        proposalId: 'p-5',
        proposalVersion: 7,
        title: 't',
        questions: [
          { id: 'q1', label: 'l', kind: 'single_choice', options: OPTIONS },
          { id: 'q2', label: 'l', kind: 'multiple_choice', options: OPTIONS },
          { id: 'q3', label: 'l', kind: 'ranking', options: OPTIONS },
          { id: 'q4', label: 'l', kind: 'consent' },
          { id: 'q5', label: 'l', kind: 'text' },
        ],
      },
      data: { q0: 'a', q1: 'a,b,c', 'q2.r1': 'a', 'q2.r2': 'b', 'q2.r3': 'c', q3: 'object', 'q3.reason': 'x', q4: 'Text' },
    },
  ];

  for (const form of forms) {
    it(`${form.name}`, async () => {
      const h = harness();
      const { content, messageId } = await h.post(form.card);

      await h.click(form.from ?? ANNA, messageId, { ...h.actionData(content, form.action), ...form.data });

      assert.equal(h.kernel.events.length, 1, 'the submission was accepted and emitted');
      const emitted = h.kernel.events[0]?.cardAnswer;
      assert.ok(emitted);
      const read = readCardAnswerFields(emitted);
      assert.ok('value' in read, `the kernel refuses the event: ${'error' in read ? read.error : ''}`);
      // Nothing is lost or reshaped on the way through the kernel's reader.
      assert.deepEqual(read.value.answers, emitted.answers);
      assert.equal(read.value.proposalId, form.card.proposalId);
      assert.equal(read.value.proposalVersion, form.card.proposalVersion);
      // One answer per question, in question order.
      const ids = (emitted.answers ?? []).map((answer) => answer.questionId);
      assert.equal(new Set(ids).size, ids.length);
      assert.equal(typeof emitted.cardId === 'string' && emitted.cardId.length > 0, true);
    });
  }

  it('the copied check refuses what the contract forbids', () => {
    const bad: unknown[][] = [
      [{ questionId: 'q', kind: 'rating', value: 5 }],
      [{ questionId: 'q', kind: 'single_choice' }],
      [{ questionId: 'q', kind: 'single_choice', optionId: 'a' }, { questionId: 'q', kind: 'single_choice', optionId: 'b' }],
      [{ questionId: 'q.1', kind: 'single_choice', optionId: 'a' }],
      [{ questionId: 'q', kind: 'ranking', ranking: ['a', 'a'] }],
      [{ questionId: 'q', kind: 'consent', decision: 'maybe' }],
    ];
    for (const answers of bad) {
      assert.deepEqual(readCardAnswerFields({ answers }), { error: 'cardAnswer.answers' }, JSON.stringify(answers));
    }
  });

  it('the copy is the kernel\'s code, where a kernel checkout with the contract is present', () => {
    // CI has no kernel checkout; then the copy stands as written.
    const here = dirname(fileURLToPath(import.meta.url));
    const kernelFile = join(here, '..', '..', 'odoo-bot', 'middleware', 'src', 'channels', 'questionCard.ts');
    const copyFile = join(here, '..', 'tests', 'support', 'kernelCardAnswers.ts');
    if (!existsSync(kernelFile) || !existsSync(copyFile)) return;
    const kernel = readFileSync(kernelFile, 'utf8');
    const copy = readFileSync(copyFile, 'utf8');
    const source = (text: string, name: string): string | undefined => {
      const start = text.search(new RegExp(`^(export )?function ${name}\\(`, 'm'));
      if (start === -1) return undefined;
      return text.slice(start, text.indexOf('\n}\n', start) + 3).replace(/^export /, '');
    };
    for (const name of ['readAnswer', 'readCardAnswerFields']) {
      const theirs = source(kernel, name);
      if (theirs === undefined) return; // a kernel checkout from before #2105
      assert.equal(source(copy, name), theirs, `${name} differs from the kernel: copy it again and re-check the emitted shapes`);
    }
  });
});

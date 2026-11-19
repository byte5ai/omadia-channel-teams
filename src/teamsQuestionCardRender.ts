// Question cards as Adaptive Cards (#2043 §6.1, #2089; design note
// "Design: questionnaire cards in Teams").
//
// Schema 1.5 (Universal Actions need it in Teams), one column, full width:
// a meeting chat is the narrowest place a card is shown.
//
// THREE CARDS exist for one question card:
//   the question card   posted into the conversation; everybody sees it
//   the own-answer card returned to the person who clicked; only they see it
//   the closed card     replaces the posted message when the step closes
//
// CLIENT-SIDE CHECKS ARE A CONVENIENCE. `isRequired`, `errorMessage` and
// `maxLength` are set so a client that honours them tells the person early.
// Nothing relies on them: every submission is read again on the server
// (`teamsQuestionCardAnswers.ts`), and a refused one comes back as this same
// card with the person's values kept and the reason shown.

import { CardFactory, type Attachment } from 'botbuilder';

import type { CardAnswerValue, CardOption, CardQuestion, QuestionCard } from './kernel-types.js';
import { describeAnswer, inputId, VOLUNTEER_PRESET, type CardInputValues } from './teamsQuestionCardAnswers.js';
import {
  DEFAULT_TEXT_MAX_LENGTH,
  EXPANDED_CHOICE_LIMIT,
  isRequired,
  rankCount,
  textLimit,
} from './teamsQuestionCardSpec.js';
import { formatCardTime, questionCardTexts, type QuestionCardTexts } from './teamsQuestionCardTexts.js';

export const QUESTION_CARD_DATA_TYPE = 'facilitation.card';
/** Submit the answers. */
export const QUESTION_CARD_ANSWER_VERB = 'facilitation.card.answer';
/** The per-person refresh: "what does this person see". */
export const QUESTION_CARD_VIEW_VERB = 'facilitation.card.view';
/** "Ändern": show the inputs again, filled with the person's answer. */
export const QUESTION_CARD_CHANGE_VERB = 'facilitation.card.change';
/** People of the conversation the card is in (Teams people picker). */
const CONVERSATION_MEMBERS_DATASET = 'graph.microsoft.com/users?scope=currentContext';

type Element = Record<string, unknown>;

/**
 * What every action of a card carries back: which card, and which version of
 * the proposal the person was looking at. NOT the card itself: a click comes
 * from a client and may be forged, so the questions and options are always
 * taken from the record of what was posted.
 */
export interface QuestionCardActionData {
  type: typeof QUESTION_CARD_DATA_TYPE;
  cardId: string;
  proposalId?: string;
  proposalVersion?: number;
}

export function questionCardActionData(card: QuestionCard): QuestionCardActionData {
  return {
    type: QUESTION_CARD_DATA_TYPE,
    cardId: card.cardId,
    ...(card.proposalId !== undefined ? { proposalId: card.proposalId } : {}),
    ...(card.proposalVersion !== undefined ? { proposalVersion: card.proposalVersion } : {}),
  };
}

function choices(options: readonly CardOption[]): Element[] {
  return options.map((option) => ({ title: option.label, value: option.id }));
}

function requiredProps(question: CardQuestion, message: string): Element {
  return isRequired(question) ? { isRequired: true, errorMessage: message } : {};
}

function questionElements(
  question: CardQuestion,
  index: number,
  values: CardInputValues,
  error: string | undefined,
  texts: QuestionCardTexts,
): Element[] {
  const elements: Element[] = [
    { type: 'TextBlock', text: question.label, weight: 'Bolder', wrap: true, spacing: 'Medium' },
  ];
  if (question.help && question.help.trim().length > 0) {
    elements.push({ type: 'TextBlock', text: question.help, isSubtle: true, size: 'Small', spacing: 'None', wrap: true });
  }
  if (error) {
    elements.push({ type: 'TextBlock', text: error, color: 'Attention', wrap: true, spacing: 'Small' });
  }
  const id = inputId(index);
  const value = values[id];
  const preset = value !== undefined ? { value } : {};
  switch (question.kind) {
    case 'single_choice':
      elements.push({
        type: 'Input.ChoiceSet',
        id,
        style: question.options.length <= EXPANDED_CHOICE_LIMIT ? 'expanded' : 'compact',
        wrap: true,
        choices: choices(question.options),
        ...requiredProps(question, texts.chooseOne),
        ...preset,
      });
      break;
    case 'multiple_choice':
      elements.push({
        type: 'Input.ChoiceSet',
        id,
        style: 'expanded',
        isMultiSelect: true,
        wrap: true,
        choices: choices(question.options),
        ...requiredProps(question, texts.chooseOne),
        ...preset,
      });
      break;
    case 'ranking':
      for (let place = 1; place <= rankCount(question); place += 1) {
        const rankId = inputId(index, `r${String(place)}`);
        const rankValue = values[rankId];
        elements.push({
          type: 'Input.ChoiceSet',
          id: rankId,
          label: question.rankLabels?.[place - 1] ?? texts.rankLabel(place),
          style: 'compact',
          choices: choices(question.options),
          ...requiredProps(question, texts.chooseOne),
          ...(rankValue !== undefined ? { value: rankValue } : {}),
        });
      }
      break;
    case 'consent': {
      const reasonId = inputId(index, 'reason');
      const reason = values[reasonId];
      elements.push(
        {
          type: 'Input.ChoiceSet',
          id,
          style: 'expanded',
          choices: [
            { title: question.agreeLabel ?? texts.agree, value: 'agree' },
            { title: question.objectLabel ?? texts.object, value: 'object' },
          ],
          ...requiredProps(question, texts.chooseOne),
          ...preset,
        },
        {
          type: 'Input.Text',
          id: reasonId,
          label: texts.reasonLabel,
          isMultiline: true,
          maxLength: DEFAULT_TEXT_MAX_LENGTH,
          ...(reason !== undefined ? { value: reason } : {}),
        },
      );
      break;
    }
    case 'assignment':
      elements.push(
        question.candidates === 'conversation_members'
          ? {
              type: 'Input.ChoiceSet',
              id,
              style: 'filtered',
              choices: [],
              'choices.data': { type: 'Data.Query', dataset: CONVERSATION_MEMBERS_DATASET },
              ...requiredProps(question, texts.pickSomeone),
              ...preset,
            }
          : {
              type: 'Input.ChoiceSet',
              id,
              style: 'compact',
              choices: choices(question.candidates),
              ...requiredProps(question, texts.pickSomeone),
              ...preset,
            },
      );
      break;
    case 'text':
      elements.push({
        type: 'Input.Text',
        id,
        isMultiline: question.multiline !== false,
        maxLength: textLimit(question),
        ...requiredProps(question, texts.required),
        ...preset,
      });
      break;
    default:
      break;
  }
  return elements;
}

function header(card: QuestionCard, texts: QuestionCardTexts): Element[] {
  const elements: Element[] = [{ type: 'TextBlock', text: card.title, weight: 'Bolder', size: 'Medium', wrap: true }];
  if (card.intro && card.intro.trim().length > 0) {
    elements.push({ type: 'TextBlock', text: card.intro, wrap: true, spacing: 'Small' });
  }
  const facts = [
    ...(card.closesAt ? [texts.openUntil(formatCardTime(card.closesAt, card.locale))] : []),
    ...(card.changeable !== false ? [texts.changeableUntilClose] : []),
  ];
  if (facts.length > 0) {
    elements.push({ type: 'TextBlock', text: facts.join(' '), isSubtle: true, size: 'Small', spacing: 'Small', wrap: true });
  }
  return elements;
}

/** The per-person refresh. Without `userIds` Teams refreshes automatically
 *  in a conversation of up to 60 members and offers "Refresh Card" in the
 *  message menu above that; the own-answer card returned to a click does not
 *  depend on it. */
function refresh(card: QuestionCard): Element {
  return { action: { type: 'Action.Execute', verb: QUESTION_CARD_VIEW_VERB, data: questionCardActionData(card) } };
}

function adaptiveCard(body: Element[], actions: Element[], extra: Element = {}): Record<string, unknown> {
  return {
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    ...extra,
    body,
    ...(actions.length > 0 ? { actions } : {}),
  };
}

function isSingleAppendText(card: QuestionCard): boolean {
  const only = card.questions.length === 1 ? card.questions[0] : undefined;
  return only?.kind === 'text' && only.mode === 'append';
}

export interface QuestionCardView {
  /** What the person entered, to show the card again with it. */
  values?: CardInputValues;
  /** Question index → what is wrong. */
  errors?: Readonly<Record<number, string>>;
  /** A line above the questions (for example "the proposal changed"). */
  notice?: string;
}

/** The card with its inputs. */
export function buildQuestionCardContent(card: QuestionCard, view: QuestionCardView = {}): Record<string, unknown> {
  const texts = questionCardTexts(card.locale);
  const values = view.values ?? {};
  const errors = view.errors ?? {};
  const body: Element[] = [...header(card, texts)];
  const notice = view.notice ?? (Object.keys(errors).length > 0 ? texts.fixErrors : undefined);
  if (notice) body.push({ type: 'TextBlock', text: notice, color: 'Attention', weight: 'Bolder', wrap: true });
  for (const [index, question] of card.questions.entries()) {
    body.push(...questionElements(question, index, values, errors[index], texts));
  }
  const data = questionCardActionData(card);
  const only = card.questions.length === 1 ? card.questions[0] : undefined;
  const actions: Element[] = [];
  // "Ich übernehme das" answers for the person who clicks, with no input.
  // Only on a card whose single question is the assignment: a button that
  // sends no inputs would drop the answers to the other questions.
  if (only?.kind === 'assignment' && only.allowVolunteer !== false) {
    actions.push({
      type: 'Action.Execute',
      title: texts.volunteer,
      verb: QUESTION_CARD_ANSWER_VERB,
      associatedInputs: 'none',
      data: { ...data, preset: { [inputId(0)]: VOLUNTEER_PRESET } },
    });
  }
  actions.push({
    type: 'Action.Execute',
    title: only?.kind === 'assignment' ? texts.proposePerson : isSingleAppendText(card) ? texts.submitIdea : texts.submit,
    verb: QUESTION_CARD_ANSWER_VERB,
    data,
  });
  return adaptiveCard(body, actions, { refresh: refresh(card) });
}

export interface OwnAnswerView {
  answers: readonly CardAnswerValue[];
  /** ISO 8601, when the answer was sent. */
  answeredAt: string;
  /** Every text of the person so far, for an idea collection (`append`). */
  contributions?: readonly string[];
}

/** What the person who answered sees instead of the inputs. */
export function buildOwnAnswerCardContent(card: QuestionCard, own: OwnAnswerView): Record<string, unknown> {
  const texts = questionCardTexts(card.locale);
  const body: Element[] = [{ type: 'TextBlock', text: card.title, weight: 'Bolder', size: 'Medium', wrap: true }];
  const isAppend = isSingleAppendText(card);
  if (isAppend && own.contributions) {
    body.push({ type: 'TextBlock', text: texts.yourContributions(own.contributions.length), weight: 'Bolder', wrap: true });
    for (const contribution of own.contributions) {
      body.push({ type: 'TextBlock', text: `• ${contribution}`, wrap: true, spacing: 'Small' });
    }
  } else {
    body.push({ type: 'TextBlock', text: texts.yourAnswer, weight: 'Bolder', wrap: true });
    for (const answer of own.answers) {
      const question = card.questions.find((candidate) => candidate.id === answer.questionId);
      if (!question) continue;
      const described = describeAnswer(question, answer, texts);
      body.push({
        type: 'TextBlock',
        text: card.questions.length > 1 ? `${question.label}: ${described}` : described,
        wrap: true,
        spacing: 'Small',
      });
    }
  }
  const facts = [
    texts.answeredAt(formatCardTime(own.answeredAt, card.locale)),
    ...(card.changeable !== false && !isAppend ? [texts.changeableUntilClose] : []),
  ];
  body.push({ type: 'TextBlock', text: facts.join(' '), isSubtle: true, size: 'Small', wrap: true });
  const actions: Element[] =
    isAppend || card.changeable !== false
      ? [
          {
            type: 'Action.Execute',
            title: isAppend ? texts.addAnother : texts.change,
            verb: QUESTION_CARD_CHANGE_VERB,
            associatedInputs: 'none',
            data: questionCardActionData(card),
          },
        ]
      : [];
  return adaptiveCard(body, actions, { refresh: refresh(card) });
}

export interface ClosedCardView {
  title: string;
  /** The engine's result, line by line. The adapter computes nothing. */
  resultLines: readonly string[];
  /** ISO 8601. */
  closedAt: string;
  locale?: string;
}

/** The card after the step closed: the result, no inputs, no refresh. */
export function buildClosedCardContent(closed: ClosedCardView): Record<string, unknown> {
  const texts = questionCardTexts(closed.locale);
  return adaptiveCard(
    [
      { type: 'TextBlock', text: closed.title, weight: 'Bolder', size: 'Medium', wrap: true },
      { type: 'TextBlock', text: texts.closedAt(formatCardTime(closed.closedAt, closed.locale)), isSubtle: true, size: 'Small', spacing: 'Small', wrap: true },
      ...closed.resultLines.map((line) => ({ type: 'TextBlock', text: line, wrap: true, spacing: 'Small' })),
    ],
    [],
  );
}

export function toCardAttachment(content: Record<string, unknown>): Attachment {
  return CardFactory.adaptiveCard(content);
}

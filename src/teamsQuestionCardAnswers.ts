// Reading what a person submitted on a question card (#2043 §6.1, #2089).
//
// EVERYTHING IS CHECKED HERE, on the server. Which checks a Teams client runs
// before it sends a click (`isRequired`, `maxLength`) is not documented for
// `Action.Execute`, and a click can be crafted by hand. So a submission is
// read against the card's questions: unknown options, a rank used twice, a
// missing reason, too long a text are all refused here, whatever the client
// let through. The engine checks the answer once more against the card it
// sent; this module is the first line and the person's feedback.

import type { CardAnswerValue, CardOption, CardQuestion, QuestionCard } from './kernel-types.js';
import { DEFAULT_TEXT_MAX_LENGTH, isRequired as isQuestionRequired, rankCount, textLimit } from './teamsQuestionCardSpec.js';
import type { QuestionCardTexts } from './teamsQuestionCardTexts.js';

/** Value of the "I will take this" button for an assignment question. */
export const VOLUNTEER_PRESET = 'self';
const ENTRA_OBJECT_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** The id of a question's input on the card. By POSITION, not by the
 *  engine's question id: ids on the wire stay ours and cannot collide. */
export function inputId(questionIndex: number, part?: string): string {
  return part ? `q${String(questionIndex)}.${part}` : `q${String(questionIndex)}`;
}

/** The raw input values of one submission, by input id. Kept so a card can
 *  be shown again with what the person entered. */
export type CardInputValues = Readonly<Record<string, string>>;

export type CardSubmission =
  | { ok: true; answers: CardAnswerValue[]; values: CardInputValues }
  /** `errors`: question index → what is wrong, in the card's language. */
  | { ok: false; errors: Readonly<Record<number, string>>; values: CardInputValues };

function text(raw: unknown): string {
  return typeof raw === 'string' ? raw.trim() : typeof raw === 'number' || typeof raw === 'boolean' ? String(raw) : '';
}

function optionOf(options: readonly CardOption[], id: string): CardOption | undefined {
  return options.find((option) => option.id === id);
}

type Read = { answer?: CardAnswerValue; error?: string };

function readQuestion(
  question: CardQuestion,
  index: number,
  values: Record<string, string>,
  presets: Readonly<Record<string, unknown>>,
  texts: QuestionCardTexts,
): Read {
  const questionId = question.id;
  const value = values[inputId(index)] ?? '';
  const isRequired = isQuestionRequired(question);
  switch (question.kind) {
    case 'single_choice': {
      if (value.length === 0) return isRequired ? { error: texts.chooseOne } : {};
      return optionOf(question.options, value)
        ? { answer: { questionId, kind: 'single_choice', optionId: value } }
        : { error: texts.unknownOption };
    }
    case 'multiple_choice': {
      const chosen = [...new Set(value.split(',').map((part) => part.trim()).filter((part) => part.length > 0))];
      if (chosen.some((id) => !optionOf(question.options, id))) return { error: texts.unknownOption };
      const min = question.min ?? (isRequired ? 1 : 0);
      if (chosen.length < min) return { error: chosen.length === 0 && min === 1 ? texts.chooseOne : texts.tooFew(min) };
      if (question.max !== undefined && chosen.length > question.max) return { error: texts.tooMany(question.max) };
      // Nothing chosen on a question that asks for nothing is no answer.
      return chosen.length === 0 ? {} : { answer: { questionId, kind: 'multiple_choice', optionIds: chosen } };
    }
    case 'ranking': {
      const places = rankCount(question);
      const ranking = Array.from({ length: places }, (_, place) => values[inputId(index, `r${String(place + 1)}`)] ?? '');
      const filled = ranking.filter((id) => id.length > 0);
      if (filled.length === 0 && !isRequired) return {};
      if (filled.length < places) return { error: texts.rankingIncomplete };
      if (ranking.some((id) => !optionOf(question.options, id))) return { error: texts.unknownOption };
      const duplicate = ranking.find((id, place) => ranking.indexOf(id) !== place);
      if (duplicate !== undefined) {
        const where = ranking.flatMap((id, place) => (id === duplicate ? [place + 1] : []));
        return { error: texts.rankingDuplicate(optionOf(question.options, duplicate)?.label ?? duplicate, where) };
      }
      return { answer: { questionId, kind: 'ranking', ranking } };
    }
    case 'consent': {
      if (value.length === 0) return isRequired ? { error: texts.chooseOne } : {};
      if (value !== 'agree' && value !== 'object') return { error: texts.unknownOption };
      const reason = values[inputId(index, 'reason')] ?? '';
      if (reason.length > DEFAULT_TEXT_MAX_LENGTH) return { error: texts.textTooLong(DEFAULT_TEXT_MAX_LENGTH) };
      if (value === 'object' && reason.length === 0 && question.reasonRequiredOnObject !== false) {
        return { error: texts.reasonRequired };
      }
      return { answer: { questionId, kind: 'consent', decision: value, ...(reason.length > 0 ? { reason } : {}) } };
    }
    case 'assignment': {
      // The volunteer button: the assignee is whoever Teams says clicked.
      if (presets[inputId(index)] === VOLUNTEER_PRESET) {
        return question.allowVolunteer === false
          ? { error: texts.unknownOption }
          : { answer: { questionId, kind: 'assignment', assignee: 'self' } };
      }
      if (value.length === 0) return isRequired ? { error: texts.pickSomeone } : {};
      if (question.candidates === 'conversation_members') {
        // The people picker returns the Entra object id of the person.
        return ENTRA_OBJECT_ID.test(value)
          ? { answer: { questionId, kind: 'assignment', assignee: { kind: 'teams-aad', id: value.toLowerCase() } } }
          : { error: texts.unknownPerson };
      }
      // A fixed list: the option ids are channel user ids (Entra object ids
      // in Teams); anything else is passed on as the option it is.
      const candidate = optionOf(question.candidates, value);
      if (!candidate) return { error: texts.unknownOption };
      const assignee = ENTRA_OBJECT_ID.test(candidate.id)
        ? { kind: 'teams-aad', id: candidate.id.toLowerCase(), displayName: candidate.label }
        : { kind: 'option', id: candidate.id, displayName: candidate.label };
      return { answer: { questionId, kind: 'assignment', assignee } };
    }
    case 'text': {
      if (value.length === 0) return isRequired ? { error: texts.required } : {};
      const limit = textLimit(question);
      return value.length > limit ? { error: texts.textTooLong(limit) } : { answer: { questionId, kind: 'text', text: value } };
    }
    default:
      return {};
  }
}

/** Every input id a card can submit. */
function inputIdsOf(card: QuestionCard): string[] {
  return card.questions.flatMap((question, index) => {
    if (question.kind === 'ranking') {
      return Array.from({ length: rankCount(question) }, (_, place) => inputId(index, `r${String(place + 1)}`));
    }
    return question.kind === 'consent' ? [inputId(index), inputId(index, 'reason')] : [inputId(index)];
  });
}

/**
 * Read one submission against the card's questions. `data` is the
 * `Action.Execute` data as it came from the client: the input values by
 * input id, plus `preset` of the button that was clicked.
 *
 * A submission with no answer at all (nothing required, nothing entered) is
 * an error on the first question: an empty click is never an answer.
 */
export function readCardSubmission(card: QuestionCard, data: unknown, texts: QuestionCardTexts): CardSubmission {
  const record = data && typeof data === 'object' ? (data as Record<string, unknown>) : {};
  const presets =
    record['preset'] && typeof record['preset'] === 'object' ? (record['preset'] as Record<string, unknown>) : {};
  const values: Record<string, string> = {};
  for (const id of inputIdsOf(card)) {
    const value = text(record[id]);
    if (value.length > 0) values[id] = value;
  }
  const answers: CardAnswerValue[] = [];
  const errors: Record<number, string> = {};
  for (const [index, question] of card.questions.entries()) {
    const read = readQuestion(question, index, values, presets, texts);
    if (read.error) errors[index] = read.error;
    else if (read.answer) answers.push(read.answer);
  }
  if (Object.keys(errors).length === 0 && answers.length === 0) errors[0] = texts.required;
  return Object.keys(errors).length > 0 ? { ok: false, errors, values } : { ok: true, answers, values };
}

/** One line per answer for the person's own view, in the card's words. */
export function describeAnswer(question: CardQuestion, answer: CardAnswerValue, texts: QuestionCardTexts): string {
  const labelOf = (options: readonly CardOption[], id: string): string => optionOf(options, id)?.label ?? id;
  switch (answer.kind) {
    case 'single_choice':
      return question.kind === 'single_choice' ? labelOf(question.options, answer.optionId) : answer.optionId;
    case 'multiple_choice':
      return question.kind === 'multiple_choice'
        ? answer.optionIds.map((id) => labelOf(question.options, id)).join(', ')
        : answer.optionIds.join(', ');
    case 'ranking':
      return answer.ranking
        .map((id, place) => `${String(place + 1)}. ${question.kind === 'ranking' ? labelOf(question.options, id) : id}`)
        .join(', ');
    case 'consent': {
      const decision =
        answer.decision === 'agree'
          ? (question.kind === 'consent' && question.agreeLabel) || texts.agree
          : (question.kind === 'consent' && question.objectLabel) || texts.object;
      return answer.reason ? `${decision}: ${answer.reason}` : decision;
    }
    case 'assignment':
      return answer.assignee === 'self' ? texts.self : (answer.assignee.displayName ?? answer.assignee.id);
    case 'text':
      return answer.text;
    default:
      return '';
  }
}

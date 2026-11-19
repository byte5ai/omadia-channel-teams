// Question cards: what this adapter accepts (byte5ai/omadia#2043 §6.1, #2089).
//
// The engine sends a `QuestionCard`; this module says whether it can be
// rendered and holds the limits. Nothing here talks to Teams.

import type { CardSupport, CardOption, CardQuestion, CardQuestionKind, QuestionCard } from './kernel-types.js';

/** The question kinds this adapter renders. `number` and `date` are in the
 *  contract and not built; a card that contains one is refused. */
export const RENDERED_CARD_KINDS: readonly CardQuestionKind[] = [
  'single_choice',
  'multiple_choice',
  'ranking',
  'consent',
  'assignment',
  'text',
];

export const MAX_CARD_QUESTIONS = 5;
/** Above this a ranking is one dropdown too many on a phone; the chat ballot
 *  is the better input. */
export const MAX_RANKING_OPTIONS = 6;
export const MAX_CHOICE_OPTIONS = 20;
/** Radio buttons up to this many options, a dropdown above. */
export const EXPANDED_CHOICE_LIMIT = 6;
export const DEFAULT_TEXT_MAX_LENGTH = 2000;
/** Hard cap for any text answer, whatever the card asks for. */
export const HARD_TEXT_MAX_LENGTH = 4000;
/** The kernel's `CARD_ID_PATTERN`: no `.` and no `,`, because several chosen
 *  options come back as one comma-separated string. */
const CARD_ID_PATTERN = /^[A-Za-z0-9_-]{1,64}$/;
const ID_RULE = '1 to 64 characters of A-Z a-z 0-9 _ -';
const MAX_LABEL_LENGTH = 300;

/** Is the question required? The contract's default is yes. */
export function isRequired(question: Pick<CardQuestion, 'required'>): boolean {
  return question.required !== false;
}

/** `supports.card` of the conversation send provider. */
export const TEAMS_CARD_CAPABILITY: CardSupport = {
  kinds: RENDERED_CARD_KINDS,
  maxQuestions: MAX_CARD_QUESTIONS,
  update: true,
};

function isId(raw: unknown): raw is string {
  return typeof raw === 'string' && CARD_ID_PATTERN.test(raw);
}

function isLabel(raw: unknown): raw is string {
  return typeof raw === 'string' && raw.trim().length > 0 && raw.length <= MAX_LABEL_LENGTH;
}

function optionsProblem(options: unknown, where: string, max: number): string | undefined {
  if (!Array.isArray(options)) return `${where}: options are missing`;
  if (options.length < 2) return `${where}: at least two options are needed`;
  if (options.length > max) return `${where}: at most ${String(max)} options`;
  const seen = new Set<string>();
  for (const option of options as CardOption[]) {
    if (!option || !isId(option.id)) return `${where}: every option needs an id of ${ID_RULE}`;
    if (!isLabel(option.label)) return `${where}: option '${option.id}' has no usable label`;
    if (seen.has(option.id)) return `${where}: option id '${option.id}' occurs twice`;
    seen.add(option.id);
  }
  return undefined;
}

function questionProblem(question: CardQuestion, index: number): string | undefined {
  const where = `question ${String(index + 1)}`;
  if (!question || typeof question !== 'object') return `${where} is not an object`;
  if (!isId(question.id)) return `${where}: the id must be ${ID_RULE}`;
  if (!isLabel(question.label)) return `${where}: the label is empty or too long`;
  if (!RENDERED_CARD_KINDS.includes(question.kind)) return `${where}: kind '${String(question.kind)}' is not rendered by the Teams channel`;
  switch (question.kind) {
    case 'single_choice':
      return optionsProblem(question.options, where, MAX_CHOICE_OPTIONS);
    case 'multiple_choice': {
      const problem = optionsProblem(question.options, where, MAX_CHOICE_OPTIONS);
      if (problem) return problem;
      const { min, max } = question;
      if (min !== undefined && (!Number.isInteger(min) || min < 0 || min > question.options.length)) return `${where}: min is out of range`;
      if (max !== undefined && (!Number.isInteger(max) || max < 1 || max > question.options.length)) return `${where}: max is out of range`;
      if (min !== undefined && max !== undefined && min > max) return `${where}: min is above max`;
      return undefined;
    }
    case 'ranking': {
      const problem = optionsProblem(question.options, where, MAX_RANKING_OPTIONS);
      if (problem) return problem;
      const { ranks } = question;
      if (ranks !== undefined && (!Number.isInteger(ranks) || ranks < 1 || ranks > question.options.length)) {
        return `${where}: ranks must be between 1 and the number of options`;
      }
      return undefined;
    }
    case 'assignment':
      return question.candidates === 'conversation_members'
        ? undefined
        : optionsProblem(question.candidates, where, MAX_CHOICE_OPTIONS);
    case 'text': {
      const { maxLength } = question;
      if (maxLength !== undefined && (!Number.isInteger(maxLength) || maxLength < 1 || maxLength > HARD_TEXT_MAX_LENGTH)) {
        return `${where}: maxLength must be between 1 and ${String(HARD_TEXT_MAX_LENGTH)}`;
      }
      return undefined;
    }
    default:
      return undefined;
  }
}

/** Why a card cannot be rendered; `undefined` when it can. */
export function questionCardProblem(card: unknown): string | undefined {
  if (!card || typeof card !== 'object') return 'the card is not an object';
  const c = card as QuestionCard;
  if (!isId(c.cardId)) return `cardId must be ${ID_RULE}`;
  if (!isLabel(c.title)) return 'the title is empty or too long';
  if (c.proposalId !== undefined && (typeof c.proposalId !== 'string' || c.proposalId.length === 0 || c.proposalId.length > 128)) {
    return 'proposalId must have 1 to 128 characters';
  }
  if (c.proposalVersion !== undefined && (!Number.isInteger(c.proposalVersion) || c.proposalVersion < 0)) {
    return 'proposalVersion must be a non-negative integer';
  }
  if (c.closesAt !== undefined && (typeof c.closesAt !== 'string' || Number.isNaN(Date.parse(c.closesAt)))) {
    return 'closesAt is not an ISO 8601 time';
  }
  if (!Array.isArray(c.questions) || c.questions.length === 0) return 'a card needs at least one question';
  if (c.questions.length > MAX_CARD_QUESTIONS) return `a card carries at most ${String(MAX_CARD_QUESTIONS)} questions`;
  const ids = new Set<string>();
  for (const [index, question] of c.questions.entries()) {
    const problem = questionProblem(question, index);
    if (problem) return problem;
    if (ids.has(question.id)) return `question id '${question.id}' occurs twice`;
    ids.add(question.id);
  }
  return undefined;
}

/** How many places a ranking question ranks. */
export function rankCount(question: Extract<CardQuestion, { kind: 'ranking' }>): number {
  return question.ranks ?? question.options.length;
}

/** The longest text a text question accepts. */
export function textLimit(question: Extract<CardQuestion, { kind: 'text' }>): number {
  return Math.min(question.maxLength ?? DEFAULT_TEXT_MAX_LENGTH, HARD_TEXT_MAX_LENGTH);
}

/** `true` once the card's closing time has passed. */
export function isPastClose(card: Pick<QuestionCard, 'closesAt'>, now: number): boolean {
  return card.closesAt !== undefined && now >= Date.parse(card.closesAt);
}

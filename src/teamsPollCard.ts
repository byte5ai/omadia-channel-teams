// Live poll card for native facilitators (byte5ai/omadia#2043 §6.3, #44).
//
// A poll is an Adaptive Card with one `Action.Execute` per option. A click
// arrives as an `adaptiveCard/action` invoke whose `value.action.data` is a
// PollAnswerValue and whose `from` is set by Teams: nobody can answer for
// someone else. The answerer is `from.aadObjectId`; a click without one (an
// anonymous meeting participant) is passed on flagged as unverified. The bot
// does not count or store the click; it emits a `card_answer` message event
// and the engine records it.
//
// `Action.Submit` payloads of the same shape (a message activity with
// `value`) are still read, for cards posted by an earlier build.

import { CardFactory, type Attachment } from 'botbuilder';

import { assertTeamsDecoration, TEAMS_DECORATION_IDS } from './teamsDecorations.js';

export const POLL_ANSWER_VALUE_TYPE = 'facilitation.poll';
/** `verb` of the poll's `Action.Execute`. */
export const POLL_ANSWER_VERB = 'facilitation.poll.answer';

/** Adaptive Card limit for one ActionSet in Teams. */
const MAX_POLL_OPTIONS = 6;
const MAX_ID_LENGTH = 128;

export interface PollOption {
  /** Stable option id the engine evaluates (not the label). */
  id: string;
  label: string;
}

export interface BuildPollCardInput {
  /** The engine's id for this poll; round-tripped with every answer. */
  pollId: string;
  question: string;
  options: readonly PollOption[];
  /** Optional line under the question (e.g. the rule or the deadline). */
  hint?: string;
}

export interface PollAnswerValue {
  type: typeof POLL_ANSWER_VALUE_TYPE;
  pollId: string;
  optionId: string;
}

/** Why a poll request cannot be rendered; `undefined` when it can. */
export function validatePollInput(input: BuildPollCardInput): string | undefined {
  if (!isBoundedId(input.pollId)) return 'pollId must be a non-empty string of at most 128 characters';
  if (input.question.trim().length === 0) return 'the poll question is empty';
  if (input.options.length < 2) return 'a poll needs at least two options';
  if (input.options.length > MAX_POLL_OPTIONS) {
    return `a poll card carries at most ${String(MAX_POLL_OPTIONS)} options`;
  }
  const ids = new Set<string>();
  for (const option of input.options) {
    if (!isBoundedId(option.id)) return 'every option needs a non-empty id of at most 128 characters';
    if (option.label.trim().length === 0) return `option '${option.id}' has an empty label`;
    if (ids.has(option.id)) return `option id '${option.id}' occurs twice`;
    ids.add(option.id);
  }
  return undefined;
}

export function buildPollCard(input: BuildPollCardInput): Attachment {
  assertTeamsDecoration(TEAMS_DECORATION_IDS.facilitationPoll);
  const problem = validatePollInput(input);
  if (problem) throw new Error(`invalid poll: ${problem}`);
  const body: Record<string, unknown>[] = [
    { type: 'TextBlock', text: input.question, weight: 'Bolder', size: 'Medium', wrap: true },
  ];
  if (input.hint && input.hint.trim().length > 0) {
    body.push({ type: 'TextBlock', text: input.hint, size: 'Small', isSubtle: true, spacing: 'None', wrap: true });
  }
  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body,
    actions: input.options.map((option) => ({
      type: 'Action.Execute',
      verb: POLL_ANSWER_VERB,
      title: option.label,
      data: {
        type: POLL_ANSWER_VALUE_TYPE,
        pollId: input.pollId,
        optionId: option.id,
      } satisfies PollAnswerValue,
    })),
  });
}

/** The card payload of an activity: `value.action.data` of an
 *  `adaptiveCard/action` invoke, else the `value` of an Action.Submit. */
export function cardActionData(value: unknown): unknown {
  const action = (value as { action?: { type?: unknown; data?: unknown } } | null)?.action;
  return action && typeof action === 'object' && action.type === 'Action.Execute' ? action.data : value;
}

/** Runtime guard for a poll click, from either action type. Some Teams
 *  clients deliver the `data` as a JSON string; both forms are read. */
export function parsePollAnswerValue(value: unknown): PollAnswerValue | undefined {
  let candidate: unknown = cardActionData(value);
  if (typeof candidate === 'string') {
    try {
      candidate = JSON.parse(candidate);
    } catch {
      return undefined;
    }
  }
  if (!candidate || typeof candidate !== 'object') return undefined;
  const v = candidate as Record<string, unknown>;
  if (v['type'] !== POLL_ANSWER_VALUE_TYPE) return undefined;
  const pollId = v['pollId'];
  const optionId = v['optionId'];
  if (!isBoundedId(pollId) || !isBoundedId(optionId)) return undefined;
  return { type: POLL_ANSWER_VALUE_TYPE, pollId, optionId };
}

function isBoundedId(raw: unknown): raw is string {
  return typeof raw === 'string' && raw.length > 0 && raw.length <= MAX_ID_LENGTH;
}

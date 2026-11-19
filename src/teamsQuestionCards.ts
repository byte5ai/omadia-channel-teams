// Question cards at runtime (#2043 §6.1, #2089): which cards this plugin
// posted, and what a click on one of them does.
//
// A CLICK IS ANSWERED AT ONCE, from what the adapter knows. Teams waits only
// a short time for the answer to an `adaptiveCard/action` invoke (how long is
// not documented), so the adapter never waits for the engine: it reads the
// submission, emits the `card_answer` event and returns the person's own
// view. Whether the answer counts is the engine's decision.
//
// WHAT A CLICK IS CHECKED AGAINST. A click carries the card back from the
// client, and a client can send anything. While this process remembers the
// card it posted (the registry below), the click is checked against THAT:
// the conversation it was posted into, its proposal version, whether it was
// closed. After a restart the registry is empty; the click is then read
// against the card it carries, and the engine, which holds the card it sent,
// drops what does not match (unknown card id, wrong version, unknown option).

import type { Activity, AdaptiveCardInvokeResponse } from 'botbuilder';

import type { CardAnswerValue, QuestionCard } from './kernel-types.js';
import { parseTeamsBotKey } from './teamsBotIdentity.js';
import { bareConversationId } from './teamsConversationIds.js';
import { readCardSubmission, type CardInputValues } from './teamsQuestionCardAnswers.js';
import {
  buildClosedCardContent,
  buildOwnAnswerCardContent,
  buildQuestionCardContent,
  QUESTION_CARD_ANSWER_VERB,
  QUESTION_CARD_CHANGE_VERB,
  QUESTION_CARD_DATA_TYPE,
  QUESTION_CARD_VIEW_VERB,
  type ClosedCardView,
} from './teamsQuestionCardRender.js';
import { isPastClose, questionCardProblem } from './teamsQuestionCardSpec.js';
import { questionCardTexts } from './teamsQuestionCardTexts.js';

const MAX_REGISTERED_CARDS = 2_000;
const MAX_OWN_ANSWERS = 10_000;
/** The response type for a card. Microsoft's samples spell it two ways; this
 *  is the one of the Bot Framework SDK and of the Universal Action Model. */
const CARD_RESPONSE_TYPE = 'application/vnd.microsoft.card.adaptive';
const MESSAGE_RESPONSE_TYPE = 'application/vnd.microsoft.activity.message';

interface RegisteredCard {
  card: QuestionCard;
  conversationId: string;
  botAppId: string;
  messageId: string | undefined;
  closed?: ClosedCardView;
}

interface OwnAnswer {
  answers: readonly CardAnswerValue[];
  values: CardInputValues;
  answeredAt: string;
  contributions: readonly string[];
}

/** What the handler hands to the kernel for one accepted submission. */
export interface QuestionCardAnswer {
  cardId: string;
  proposalId?: string;
  proposalVersion?: number;
  answers: readonly CardAnswerValue[];
}

type InvokeActivity = Partial<Activity>;

function boundedSet<V>(map: Map<string, V>, key: string, value: V, max: number): void {
  map.delete(key);
  map.set(key, value);
  if (map.size > max) {
    const oldest = map.keys().next().value;
    if (oldest !== undefined) map.delete(oldest);
  }
}

function actionOf(value: unknown): { verb: string; data: Record<string, unknown> } | undefined {
  const action = (value as { action?: { type?: unknown; verb?: unknown; data?: unknown } } | null)?.action;
  if (!action || typeof action !== 'object' || typeof action.verb !== 'string') return undefined;
  let data: unknown = action.data;
  // Some Teams clients deliver `data` as a JSON string.
  if (typeof data === 'string') {
    try {
      data = JSON.parse(data);
    } catch {
      return undefined;
    }
  }
  if (!data || typeof data !== 'object') return undefined;
  const record = data as Record<string, unknown>;
  return record['type'] === QUESTION_CARD_DATA_TYPE ? { verb: action.verb, data: record } : undefined;
}

function cardResponse(content: Record<string, unknown>): AdaptiveCardInvokeResponse {
  return { statusCode: 200, type: CARD_RESPONSE_TYPE, value: content };
}

function messageResponse(text: string): AdaptiveCardInvokeResponse {
  return { statusCode: 200, type: MESSAGE_RESPONSE_TYPE, value: text as unknown as Record<string, unknown> };
}

export class TeamsQuestionCards {
  /** `${botAppId}|${cardId}` → the card this bot posted. */
  private readonly cards = new Map<string, RegisteredCard>();
  /** `${botAppId}|${cardId}|${person}` → what that person last sent. Only
   *  for the person's own view; the engine holds what counts. */
  private readonly ownAnswers = new Map<string, OwnAnswer>();
  private readonly now: () => number;

  constructor(options: { now?: () => number } = {}) {
    this.now = options.now ?? Date.now;
  }

  private static cardKey(botAppId: string, cardId: string): string {
    return `${botAppId.toLowerCase()}|${cardId}`;
  }

  /** Remember a card this bot posted (or re-posted corrected). */
  register(entry: { card: QuestionCard; conversationId: string; botAppId: string; messageId: string | undefined }): void {
    const key = TeamsQuestionCards.cardKey(entry.botAppId, entry.card.cardId);
    const previous = this.cards.get(key);
    // A corrected card is another version of the question: what people saw
    // as "your answer" belonged to the earlier one.
    if (previous && previous.card.proposalVersion !== entry.card.proposalVersion) this.forgetOwnAnswers(key);
    boundedSet(this.cards, key, { ...entry }, MAX_REGISTERED_CARDS);
  }

  /** The card a posted message carries, for `updateCard`. */
  findByMessage(botAppId: string, messageId: string): { card: QuestionCard; conversationId: string } | undefined {
    const prefix = `${botAppId.toLowerCase()}|`;
    for (const [key, entry] of this.cards) {
      if (key.startsWith(prefix) && entry.messageId === messageId) return { card: entry.card, conversationId: entry.conversationId };
    }
    return undefined;
  }

  /** The step closed: later clicks get the closed card. */
  markClosed(botAppId: string, cardId: string, closed: ClosedCardView): void {
    const key = TeamsQuestionCards.cardKey(botAppId, cardId);
    const entry = this.cards.get(key);
    if (entry) this.cards.set(key, { ...entry, closed });
    this.forgetOwnAnswers(key);
  }

  private forgetOwnAnswers(cardKey: string): void {
    const prefix = `${cardKey}|`;
    for (const key of [...this.ownAnswers.keys()]) {
      if (key.startsWith(prefix)) this.ownAnswers.delete(key);
    }
  }

  /** Is this invoke a click on a question card (answer, change, refresh)? */
  isQuestionCardInvoke(value: unknown): boolean {
    const action = actionOf(value);
    return (
      action !== undefined &&
      [QUESTION_CARD_ANSWER_VERB, QUESTION_CARD_CHANGE_VERB, QUESTION_CARD_VIEW_VERB].includes(action.verb)
    );
  }

  /**
   * Answer one click. `emit` hands an accepted submission to the kernel and
   * says whether anybody is listening (a facilitation holds the
   * conversation); it is called at most once.
   */
  handleInvoke(activity: InvokeActivity, emit: (answer: QuestionCardAnswer) => boolean): AdaptiveCardInvokeResponse {
    const action = actionOf(activity.value);
    const fallbackTexts = questionCardTexts(undefined);
    if (!action) return messageResponse(fallbackTexts.couldNotRead);
    const cardId = action.data['cardId'];
    const botAppId = parseTeamsBotKey(activity.recipient?.id ?? '');
    if (typeof cardId !== 'string' || cardId.length === 0 || !botAppId) return messageResponse(fallbackTexts.couldNotRead);

    const cardKey = TeamsQuestionCards.cardKey(botAppId, cardId);
    const registered = this.cards.get(cardKey);
    const carried = action.data['card'];
    // The card this process posted wins over the one the click carries.
    const card: QuestionCard | undefined =
      registered?.card ??
      (questionCardProblem(carried) === undefined && (carried as QuestionCard).cardId === cardId
        ? (carried as QuestionCard)
        : undefined);
    if (!card) return messageResponse(fallbackTexts.couldNotRead);
    const texts = questionCardTexts(card.locale);

    if (registered) {
      // A card answers only in the conversation it was posted into (for a
      // channel: the channel of that thread).
      const here = bareConversationId(activity.conversation?.id ?? '');
      if (here !== bareConversationId(registered.conversationId)) {
        console.error(`[teams] question card '${cardId}' clicked in another conversation than it was posted into; ignored`);
        return messageResponse(texts.notForThisConversation);
      }
      if (registered.closed) return cardResponse(buildClosedCardContent(registered.closed));
    }
    if (isPastClose(card, this.now())) return messageResponse(texts.noLongerActive);

    const from = activity.from;
    const person = (from?.aadObjectId ?? from?.id ?? '').toLowerCase();
    const ownKey = `${cardKey}|${person}`;
    const own = person.length > 0 ? this.ownAnswers.get(ownKey) : undefined;

    // A click on an earlier version of the question: show the current one.
    const clickedVersion = action.data['proposalVersion'];
    const clickedProposal = action.data['proposalId'];
    if (
      registered &&
      (clickedVersion !== card.proposalVersion || clickedProposal !== card.proposalId)
    ) {
      return cardResponse(buildQuestionCardContent(card, { notice: texts.proposalChanged }));
    }

    if (action.verb === QUESTION_CARD_VIEW_VERB) {
      return cardResponse(own ? buildOwnAnswerCardContent(card, own) : buildQuestionCardContent(card));
    }
    if (action.verb === QUESTION_CARD_CHANGE_VERB) {
      const isAppend = card.questions.length === 1 && card.questions[0]?.kind === 'text' && card.questions[0].mode === 'append';
      return cardResponse(buildQuestionCardContent(card, own && !isAppend ? { values: own.values } : {}));
    }

    const submission = readCardSubmission(card, action.data, texts);
    if (!submission.ok) {
      // The same card, the person's values kept, the reason shown. Only this
      // person sees it; nothing is emitted.
      return cardResponse(buildQuestionCardContent(card, { values: submission.values, errors: submission.errors }));
    }
    const isListening = emit({
      cardId,
      ...(card.proposalId !== undefined ? { proposalId: card.proposalId } : {}),
      ...(card.proposalVersion !== undefined ? { proposalVersion: card.proposalVersion } : {}),
      answers: submission.answers,
    });
    if (!isListening) return messageResponse(texts.noLongerActive);
    // Teams did not say who clicked (an anonymous meeting participant): the
    // answer went out flagged as unverified; no personal view is kept.
    if (typeof from?.aadObjectId !== 'string' || from.aadObjectId.length === 0) {
      return messageResponse(texts.unverifiedAnswerer);
    }
    const newTexts = submission.answers.flatMap((answer) => (answer.kind === 'text' ? [answer.text] : []));
    const ownNow: OwnAnswer = {
      answers: submission.answers,
      values: submission.values,
      answeredAt: new Date(this.now()).toISOString(),
      contributions: [...(own?.contributions ?? []), ...newTexts],
    };
    boundedSet(this.ownAnswers, ownKey, ownNow, MAX_OWN_ANSWERS);
    return cardResponse(buildOwnAnswerCardContent(card, ownNow));
  }
}

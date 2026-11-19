// Question cards at runtime (#2043 §6.1, #2089): which cards this plugin
// posted, and what a click on one of them does.
//
// A CLICK IS ANSWERED AT ONCE, from what the adapter knows. Teams waits only
// a short time for the answer to an `adaptiveCard/action` invoke (how long is
// not documented), so the adapter never waits for the engine: it reads the
// submission, emits the `card_answer` event and returns the person's own
// view. Whether the answer counts is the engine's decision.
//
// WHAT A CLICK IS CHECKED AGAINST: the card this plugin POSTED, never
// anything the click carries. A click names a card id; the card itself comes
// from the registry below, which is a cache in front of a table
// (`teamsQuestionCardStore.ts`) and therefore still knows the card after a
// restart: the conversation it was posted into, its questions and options,
// its proposal version, whether it was closed. A card id nobody posted, or
// one whose record is gone, gets "no longer active" and emits nothing.

import type { Activity, AdaptiveCardInvokeResponse } from 'botbuilder';

import type { CardAnswerValue } from './kernel-types.js';
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
import { isPastClose } from './teamsQuestionCardSpec.js';
import type { StoredQuestionCard, TeamsQuestionCardPersistence } from './teamsQuestionCardStore.js';
import { questionCardTexts } from './teamsQuestionCardTexts.js';

const MAX_REGISTERED_CARDS = 2_000;
const MAX_OWN_ANSWERS = 10_000;
const MAX_REFUSED_SUBMISSIONS = 2_000;
/** How long a refused submission is shown again when the client asks for the
 *  person's view. Long enough for the refresh a client sends after it got a
 *  card back, short enough that nobody meets an old error the next day. */
const REFUSED_VIEW_MS = 10 * 60 * 1000;
/** The response type for a card. Microsoft's samples spell it two ways; this
 *  is the one of the Bot Framework SDK and of the Universal Action Model. */
const CARD_RESPONSE_TYPE = 'application/vnd.microsoft.card.adaptive';
const MESSAGE_RESPONSE_TYPE = 'application/vnd.microsoft.activity.message';
/** A card nobody closed is closed by the channel this long after its closing
 *  time: the engine's own update, with the result, normally comes first. */
export const OVERDUE_GRACE_MS = 5 * 60 * 1000;
/** At most this many overdue cards are closed in one sweep. */
const MAX_OVERDUE_PER_SWEEP = 50;

type RegisteredCard = StoredQuestionCard;

interface OwnAnswer {
  answers: readonly CardAnswerValue[];
  values: CardInputValues;
  answeredAt: string;
  contributions: readonly string[];
}

/** A submission that was refused: what the person entered and why not. */
interface RefusedSubmission {
  values: CardInputValues;
  errors: Readonly<Record<number, string>>;
  at: number;
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
  /** Same key → that person's last submission, if it was refused. A client
   *  may ask for the person's view (`refresh`) right after it got the card
   *  with the error back; the answer to that must be the same card, or the
   *  error line is gone before anybody read it. */
  private readonly refused = new Map<string, RefusedSubmission>();
  private readonly now: () => number;
  private readonly persistence: TeamsQuestionCardPersistence | undefined;

  /**
   * `persistence`: the table behind the registry. Without it the registry is
   * memory only and a restart forgets every posted card.
   */
  constructor(options: { now?: () => number; persistence?: TeamsQuestionCardPersistence } = {}) {
    this.now = options.now ?? Date.now;
    this.persistence = options.persistence;
  }

  /** Write a record through to the table. A failed write is logged and the
   *  card still works from memory until the next restart. */
  private async persist(record: RegisteredCard): Promise<void> {
    if (!this.persistence) return;
    try {
      await this.persistence.save(record);
    } catch (err) {
      console.error(
        `[teams] question card '${record.card.cardId}' could not be stored and is known until the next restart only:`,
        err instanceof Error ? err.message : err,
      );
    }
  }

  /** The posted card: from memory, else from the table (after a restart). */
  private async lookup(botAppId: string, cardId: string): Promise<RegisteredCard | undefined> {
    const key = TeamsQuestionCards.cardKey(botAppId, cardId);
    const cached = this.cards.get(key);
    if (cached || !this.persistence) return cached;
    const stored = await this.persistence.loadByCard(botAppId, cardId).catch(() => undefined);
    if (stored) boundedSet(this.cards, key, stored, MAX_REGISTERED_CARDS);
    return stored;
  }

  private static cardKey(botAppId: string, cardId: string): string {
    return `${botAppId.toLowerCase()}|${cardId}`;
  }

  /** Remember a card this bot posted (or re-posted corrected), in memory and
   *  in the table. */
  async register(entry: Omit<RegisteredCard, 'closed'>): Promise<void> {
    const record: RegisteredCard = { ...entry, botAppId: entry.botAppId.toLowerCase() };
    const key = TeamsQuestionCards.cardKey(record.botAppId, record.card.cardId);
    const previous = this.cards.get(key);
    // A corrected card is another version of the question: what people saw
    // as "your answer" belonged to the earlier one.
    if (previous && previous.card.proposalVersion !== record.card.proposalVersion) this.forgetOwnAnswers(key);
    boundedSet(this.cards, key, record, MAX_REGISTERED_CARDS);
    await this.persist(record);
  }

  /** The card a posted message carries, for `updateCard`: only a message
   *  this bot posted as a question card is found. */
  async findByMessage(botAppId: string, messageId: string): Promise<RegisteredCard | undefined> {
    const prefix = `${botAppId.toLowerCase()}|`;
    for (const [key, entry] of this.cards) {
      if (key.startsWith(prefix) && entry.messageId === messageId) return entry;
    }
    if (!this.persistence) return undefined;
    const stored = await this.persistence.loadByMessage(botAppId, messageId).catch(() => undefined);
    if (stored) boundedSet(this.cards, TeamsQuestionCards.cardKey(stored.botAppId, stored.card.cardId), stored, MAX_REGISTERED_CARDS);
    return stored;
  }

  /** The step closed: later clicks get the closed card. */
  async markClosed(botAppId: string, cardId: string, closed: ClosedCardView): Promise<void> {
    const key = TeamsQuestionCards.cardKey(botAppId, cardId);
    const entry = await this.lookup(botAppId, cardId);
    this.forgetOwnAnswers(key);
    if (!entry) return;
    const record: RegisteredCard = { ...entry, closed };
    this.cards.set(key, record);
    await this.persist(record);
  }

  /** Open cards whose closing time passed more than `graceMs` ago: what the
   *  registry holds in memory, and what the table still lists as open. */
  private async overdueOpen(graceMs: number): Promise<readonly RegisteredCard[]> {
    const before = this.now() - graceMs;
    const isOverdue = (entry: RegisteredCard): boolean => !entry.closed && isPastClose(entry.card, before);
    const stored = (await this.persistence?.listOverdueOpen?.(new Date(before), MAX_OVERDUE_PER_SWEEP)) ?? [];
    const found = new Map<string, RegisteredCard>();
    for (const entry of [...this.cards.values(), ...stored]) {
      const key = TeamsQuestionCards.cardKey(entry.botAppId, entry.card.cardId);
      // Memory is the newer truth: a card it knows as closed or corrected
      // is not taken from an older row.
      const current = this.cards.get(key) ?? entry;
      if (isOverdue(current) && found.size < MAX_OVERDUE_PER_SWEEP) found.set(key, current);
    }
    return [...found.values()];
  }

  /**
   * Close every card that is past its closing time and that nobody closed
   * (the engine's `updateCard` never arrived). `replace` puts the closed form
   * in place of the posted message; the card is closed here whether or not
   * that worked, because nothing can be answered on it any more. Returns how
   * many cards were closed.
   */
  async closeOverdue(
    replace: (posted: RegisteredCard, closed: ClosedCardView) => Promise<void>,
    log: (message: string) => void = (message) => console.error(message),
    graceMs: number = OVERDUE_GRACE_MS,
  ): Promise<number> {
    const overdue = await this.overdueOpen(graceMs);
    for (const posted of overdue) {
      const card = posted.card;
      const closed: ClosedCardView = {
        title: card.title,
        resultLines: [questionCardTexts(card.locale).noLongerActive],
        closedAt: card.closesAt ?? new Date(this.now()).toISOString(),
        ...(card.locale ? { locale: card.locale } : {}),
      };
      let replaced = 'the posted card was replaced by its closed form';
      try {
        await replace(posted, closed);
      } catch (err) {
        replaced = `the posted card could not be replaced (${err instanceof Error ? err.message : String(err)})`;
      }
      await this.markClosed(posted.botAppId, card.cardId, closed);
      log(`[teams] question card '${card.cardId}' was still open past its closing time (${String(card.closesAt)}) and is closed now; ${replaced}`);
    }
    return overdue.length;
  }

  private forgetOwnAnswers(cardKey: string): void {
    const prefix = `${cardKey}|`;
    for (const key of [...this.ownAnswers.keys()]) {
      if (key.startsWith(prefix)) this.ownAnswers.delete(key);
    }
    for (const key of [...this.refused.keys()]) {
      if (key.startsWith(prefix)) this.refused.delete(key);
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
  async handleInvoke(
    activity: InvokeActivity,
    emit: (answer: QuestionCardAnswer) => boolean,
  ): Promise<AdaptiveCardInvokeResponse> {
    const action = actionOf(activity.value);
    const fallbackTexts = questionCardTexts(undefined);
    if (!action) return messageResponse(fallbackTexts.couldNotRead);
    const cardId = action.data['cardId'];
    const botAppId = parseTeamsBotKey(activity.recipient?.id ?? '');
    if (typeof cardId !== 'string' || cardId.length === 0 || !botAppId) return messageResponse(fallbackTexts.couldNotRead);

    const cardKey = TeamsQuestionCards.cardKey(botAppId, cardId);
    // The card THIS BOT POSTED: from memory, or from the table after a
    // restart. Whatever card a click may carry is never read.
    const registered = await this.lookup(botAppId, cardId);
    if (!registered) {
      console.error(`[teams] click on question card '${cardId}', which this bot has no record of posting; ignored`);
      return messageResponse(fallbackTexts.noLongerActive);
    }
    const card = registered.card;
    const texts = questionCardTexts(card.locale);

    // A card answers only in the conversation it was posted into (for a
    // channel: the channel of that thread).
    const here = bareConversationId(activity.conversation?.id ?? '');
    if (here !== bareConversationId(registered.conversationId)) {
      console.error(`[teams] question card '${cardId}' clicked in another conversation than it was posted into; ignored`);
      return messageResponse(texts.notForThisConversation);
    }
    if (registered.closed) return cardResponse(buildClosedCardContent(registered.closed));
    if (isPastClose(card, this.now())) return messageResponse(texts.noLongerActive);

    const from = activity.from;
    const person = (from?.aadObjectId ?? from?.id ?? '').toLowerCase();
    const ownKey = `${cardKey}|${person}`;
    const own = person.length > 0 ? this.ownAnswers.get(ownKey) : undefined;

    // A click on an earlier version of the question: show the current one.
    const clickedVersion = action.data['proposalVersion'];
    const clickedProposal = action.data['proposalId'];
    if (clickedVersion !== card.proposalVersion || clickedProposal !== card.proposalId) {
      return cardResponse(buildQuestionCardContent(card, { notice: texts.proposalChanged }));
    }

    if (action.verb === QUESTION_CARD_VIEW_VERB) {
      const refused = person.length > 0 ? this.refused.get(ownKey) : undefined;
      if (refused && this.now() - refused.at <= REFUSED_VIEW_MS) {
        return cardResponse(buildQuestionCardContent(card, { values: refused.values, errors: refused.errors }));
      }
      return cardResponse(own ? buildOwnAnswerCardContent(card, own) : buildQuestionCardContent(card));
    }
    this.refused.delete(ownKey);
    if (action.verb === QUESTION_CARD_CHANGE_VERB) {
      const isAppend = card.questions.length === 1 && card.questions[0]?.kind === 'text' && card.questions[0].mode === 'append';
      return cardResponse(buildQuestionCardContent(card, own && !isAppend ? { values: own.values } : {}));
    }

    const submission = readCardSubmission(card, action.data, texts);
    if (!submission.ok) {
      // The same card, the person's values kept, the reason shown. Only this
      // person sees it; nothing is emitted.
      if (person.length > 0) {
        boundedSet(
          this.refused,
          ownKey,
          { values: submission.values, errors: submission.errors, at: this.now() },
          MAX_REFUSED_SUBMISSIONS,
        );
      }
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

// Inbound message events for native facilitators (byte5ai/omadia#2043 §5.3,
// §6.2; issue #43).
//
// In a conversation an agent moderates, the kernel needs EVERY human message
// with its platform-stated author: also the ones that do not mention the bot,
// and edits, deletions, reactions and card answers. This module turns Bot
// Framework activities into `ConversationMessageEvent`s.
//
// HOW MUCH the kernel is handed is the kernel's answer per bot and
// conversation (`conversationDelivery`, #2089, channel-sdk 0.3.5):
//   all        every human message; the mention gate is lifted and the engine
//              drives every turn (a facilitation in `configuring` / `active`)
//   addressed  only what addresses the bot: a mention of it, a reply to a
//              message, an answer to a card, a reaction; and every message of
//              the bot's own 1:1 chat. The mention gate stays, the ordinary
//              turn still runs (a facilitation in `following_up`)
//   none       nothing
// A kernel without `conversationDelivery` knows `all` and `none` only.
//
// WHAT "A REPLY" IS HERE. The adapter does not keep a list of the bot's own
// messages that survives a restart, so it does not decide whose message a
// reply answers. In `addressed` mode it hands over every message that names
// another message as the one it answers (a quoted reply; `replyToId` where
// that is not just the root of a channel thread), and the KERNEL keeps it
// only when that message is one of a process's own (its record of bound
// messages) and drops it otherwise, unstored.
//
// Three rules hold everywhere below:
//   1. Delivery is per bot AND conversation. Bot B in the same chat as a
//      subscribed bot A emits nothing and keeps its mention gate.
//   2. The author is `activity.from`, set by Teams. Nothing is read from the
//      message text, a card payload or a quoted block.
//   3. An event is not a turn. Emitting starts no model call, and a failure
//      to emit never breaks the activity that carried it. In a subscribed
//      conversation the adapter starts no turn at all; the engine does.

import type { Activity } from 'botbuilder';

import type { CardAnswerValue, ConversationMessageEvent, TeamsUserRef } from './kernel-types.js';
import { parseTeamsBotKey, teamsBotKey, TEAMS_BOT_KEY_PREFIX } from './teamsBotIdentity.js';
import { threadRootMessageId } from './teamsConversationIds.js';
import { toSdkConversationType } from './teamsGroupPrimitives.js';
import { stripMentionTokens } from './teamsMentions.js';
import { cardActionData, parsePollAnswerValue } from './teamsPollCard.js';

const CHANNEL_TYPE = 'teams';
/** Replaced with the catalog id by plugin.ts; the bot cannot know it. */
const CHANNEL_ID_PLACEHOLDER = 'de.byte5.channel.teams';
const MAX_TRACKED_REVISIONS = 5_000;
/** Teams marks a quoted reply with this schema.org item type. */
const QUOTED_REPLY_MARKER = 'schema.skype.com/Reply';
const FORWARDED_ATTACHMENT_TYPE = 'forwardedMessageReference';

/** What the kernel is handed of a conversation, for one bot. */
export type TeamsDeliveryMode = 'all' | 'addressed' | 'none';

/** The kernel calls this module needs, already feature-detected. */
export interface TeamsConversationMessageSink {
  isAllMessagesSubscribed(conversationId: string, channelKey: string): boolean;
  /** #2089 — absent on a kernel before channel-sdk 0.3.5: then the mode is
   *  `all` where `isAllMessagesSubscribed` says yes, and `none` otherwise. */
  conversationDelivery?(conversationId: string, channelKey: string): TeamsDeliveryMode;
  emit(event: ConversationMessageEvent): void;
}

/** Why a message counts as addressed to the bot, for the log line. */
export type TeamsAddressedReason = 'mention' | 'direct' | 'reply' | 'card' | 'reaction';

type InboundActivity = Partial<Activity>;

/** The normalized `28:<appId>` key of the bot an activity was addressed to. */
export function receivingBotKey(activity: InboundActivity): string | undefined {
  const recipientId = activity.recipient?.id;
  if (typeof recipientId !== 'string') return undefined;
  const appId = parseTeamsBotKey(recipientId);
  return appId ? teamsBotKey(appId) : undefined;
}

/** Does the activity carry a mention entity of the bot it was sent to? Set
 *  by the Teams client when a person picks the bot from the @ list; text
 *  that merely contains the bot's name is no mention. */
export function isBotMentioned(activity: InboundActivity): boolean {
  const recipientId = activity.recipient?.id;
  if (!recipientId) return false;
  return (activity.entities ?? []).some(
    (entity) => entity.type === 'mention' && (entity as { mentioned?: { id?: string } }).mentioned?.id === recipientId,
  );
}

/**
 * The id a message activity is known by to the kernel: `messageId` of the
 * event this module emits for it, and `messageId` of the origin of the turn
 * it starts (`teamsTurnOrigin`). ONE function so the two cannot differ.
 * A message: its own activity id (in a channel thread the reply's id, never
 * the thread root). An `Action.Submit` click (a message activity with a
 * `value` object): the card message it was clicked on.
 */
export function inboundMessageId(activity: {
  readonly id?: string | undefined;
  readonly replyToId?: string | undefined;
  readonly value?: unknown;
}): string | undefined {
  const isCardClick = typeof activity.value === 'object' && activity.value !== null;
  const id = isCardClick ? (activity.replyToId ?? activity.id) : activity.id;
  return typeof id === 'string' && id.length > 0 ? id : undefined;
}

/** `activity.from` as the kernel's author, or `undefined` for a bot / absent
 *  sender. Only people are authors of contributions. */
export function platformAuthor(activity: InboundActivity): TeamsUserRef | undefined {
  const from = activity.from;
  if (!from || typeof from.id !== 'string' || from.id.length === 0) return undefined;
  if (from.id.startsWith(TEAMS_BOT_KEY_PREFIX) || from.role === 'bot') return undefined;
  return {
    kind: 'teams-aad',
    id: from.aadObjectId ?? from.id,
    ...(from.name ? { displayName: from.name } : {}),
  };
}

/** Message text as plain text: `<at>Name</at>` becomes `Name`, remaining tags
 *  are dropped and the basic entities decoded. */
export function plainMessageText(raw: string | undefined): string {
  if (!raw) return '';
  return stripMentionTokens(raw)
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&')
    .trim();
}

interface QuoteInfo {
  quoted: boolean;
  forwarded: boolean;
  replyToMessageId?: string;
}

/** Entity type of a quoted reply. Microsoft's page on conversation messages
 *  says a quoted reply's metadata "is available on the inbound activity" as
 *  quoted reply entities with a `messageId`; it shows SDK code, not the wire
 *  format. Read tolerantly; NOT seen live. */
const QUOTED_REPLY_ENTITY_TYPE = 'quotedReply';

function quotedReplyEntityMessageId(activity: InboundActivity): string | undefined {
  for (const entity of activity.entities ?? []) {
    if (entity.type !== QUOTED_REPLY_ENTITY_TYPE) continue;
    const messageId = (entity as { quotedReply?: { messageId?: unknown } }).quotedReply?.messageId;
    if (typeof messageId === 'string' && messageId.length > 0) return messageId;
  }
  return undefined;
}

/** Quote / forward markers of a message. A quoted reply carries the quoted
 *  message as an HTML attachment with a `schema.skype.com/Reply` blockquote
 *  whose `itemid` is the quoted message id, or (newer) as a `quotedReply`
 *  entity. Neither shape is documented as a wire format. */
export function readQuoteInfo(activity: InboundActivity): QuoteInfo {
  let quoted = false;
  let forwarded = false;
  let replyToMessageId: string | undefined =
    typeof activity.replyToId === 'string' && activity.replyToId.length > 0 ? activity.replyToId : undefined;
  for (const attachment of activity.attachments ?? []) {
    if (attachment.contentType === FORWARDED_ATTACHMENT_TYPE) forwarded = true;
    if (attachment.contentType !== 'text/html' || typeof attachment.content !== 'string') continue;
    if (!attachment.content.includes(QUOTED_REPLY_MARKER)) continue;
    quoted = true;
    const itemId = /itemid="([^"]+)"/.exec(attachment.content)?.[1];
    if (itemId) replyToMessageId = itemId;
  }
  if (!quoted) {
    const entityMessageId = quotedReplyEntityMessageId(activity);
    if (entityMessageId) {
      quoted = true;
      replyToMessageId = entityMessageId;
    }
  }
  return { quoted, forwarded, ...(replyToMessageId ? { replyToMessageId } : {}) };
}

function occurredAt(activity: InboundActivity): string {
  const timestamp = activity.timestamp;
  if (timestamp instanceof Date) return timestamp.toISOString();
  if (typeof timestamp === 'string' && !Number.isNaN(Date.parse(timestamp))) {
    return new Date(timestamp).toISOString();
  }
  return new Date().toISOString();
}

function tenantOf(activity: InboundActivity): string | undefined {
  const fromConversation = activity.conversation?.tenantId;
  if (typeof fromConversation === 'string' && fromConversation.length > 0) return fromConversation;
  const channelData = activity.channelData as { tenant?: { id?: unknown } } | undefined;
  const fromChannelData = channelData?.tenant?.id;
  return typeof fromChannelData === 'string' && fromChannelData.length > 0 ? fromChannelData : undefined;
}

/** A card click that is not a poll: the card kind as its id. */
function genericCardId(value: unknown): string {
  if (value && typeof value === 'object') {
    const record = value as Record<string, unknown>;
    for (const key of ['type', 'kind', 'action']) {
      const candidate = record[key];
      if (typeof candidate === 'string' && candidate.length > 0) return candidate;
    }
  }
  return 'card';
}

/**
 * Emits message events for activities a subscribed bot receives. One instance
 * serves every bot of the plugin; all state is keyed by the receiving bot.
 */
export class TeamsConversationMessages {
  /** `${botKey}|${conversationId}|${messageId}` → last revision emitted.
   *  A HINT only: the kernel assigns the stored revision itself (created = 0,
   *  each edit or deletion = highest stored + 1), so it does not matter that
   *  this in-memory count starts again after a restart. */
  private readonly revisions = new Map<string, number>();

  constructor(private readonly sink: TeamsConversationMessageSink) {}

  /** What the kernel is handed of this conversation for the bot the
   *  activity was sent to. `none` for anything not positively stated. */
  delivery(activity: InboundActivity): TeamsDeliveryMode {
    const conversationId = activity.conversation?.id;
    const botKey = receivingBotKey(activity);
    if (!conversationId || !botKey) return 'none';
    try {
      if (this.sink.conversationDelivery) {
        const mode = this.sink.conversationDelivery(conversationId, botKey);
        return mode === 'all' || mode === 'addressed' ? mode : 'none';
      }
      return this.sink.isAllMessagesSubscribed(conversationId, botKey) === true ? 'all' : 'none';
    } catch (err) {
      console.error('[teams] subscription lookup failed:', err instanceof Error ? err.message : err);
      return 'none';
    }
  }

  /** Does the bot this activity was addressed to receive every message of
   *  this conversation? `false` for anything not positively subscribed. */
  isSubscribed(activity: InboundActivity): boolean {
    return this.delivery(activity) === 'all';
  }

  /**
   * Why this message activity is handed over in `addressed` mode, or
   * `undefined` when it is not. A mention and a message in the bot's own 1:1
   * chat address the bot by themselves (`addressedToBot`); a reply does only
   * if the kernel finds the answered message among a process's own.
   */
  addressedReason(activity: InboundActivity): TeamsAddressedReason | undefined {
    if (activity.conversation?.conversationType === 'personal') return 'direct';
    if (isBotMentioned(activity)) return 'mention';
    return this.replyReference(activity) !== undefined ? 'reply' : undefined;
  }

  /** The message this one answers, when that says something: a quoted
   *  message, or `replyToId` unless it is only the root of the channel
   *  thread the conversation IS (every message of a thread carries that). */
  private replyReference(activity: InboundActivity): string | undefined {
    const quote = readQuoteInfo(activity);
    if (!quote.replyToMessageId) return undefined;
    if (quote.quoted) return quote.replyToMessageId;
    const threadRoot = threadRootMessageId(activity.conversation?.id ?? '');
    return quote.replyToMessageId === threadRoot ? undefined : quote.replyToMessageId;
  }

  /** A new human message (mentioned or not). */
  emitCreated(activity: InboundActivity): boolean {
    const quote = readQuoteInfo(activity);
    return this.emit(activity, {
      kind: 'created',
      messageId: inboundMessageId(activity),
      revision: 0,
      text: plainMessageText(activity.text),
      ...(quote.replyToMessageId ? { replyToMessageId: quote.replyToMessageId } : {}),
      ...(quote.quoted ? { quoted: true } : {}),
      ...(quote.forwarded ? { forwarded: true } : {}),
    });
  }

  /** A platform edit (`messageUpdate`, `eventType: editMessage`); also an
   *  undelete, which restores the text under a new revision. */
  emitEdited(activity: InboundActivity): boolean {
    const quote = readQuoteInfo(activity);
    return this.emit(activity, {
      kind: 'edited',
      messageId: activity.id,
      revision: this.nextRevision(activity),
      text: plainMessageText(activity.text),
      ...(quote.replyToMessageId ? { replyToMessageId: quote.replyToMessageId } : {}),
      ...(quote.quoted ? { quoted: true } : {}),
      ...(quote.forwarded ? { forwarded: true } : {}),
    });
  }

  /** A platform deletion (`messageDelete`, `eventType: softDeleteMessage`). */
  emitDeleted(activity: InboundActivity): boolean {
    return this.emit(activity, {
      kind: 'deleted',
      messageId: activity.id,
      revision: this.nextRevision(activity),
    });
  }

  /** `messageReaction`: one event per reaction added or removed. The message
   *  reacted to is `replyToId`. Returns how many events went out. */
  emitReactions(activity: InboundActivity): number {
    const changes = [
      ...(activity.reactionsAdded ?? []).map((reaction) => ({ name: reaction.type, removed: false })),
      ...(activity.reactionsRemoved ?? []).map((reaction) => ({ name: reaction.type, removed: true })),
    ];
    let emitted = 0;
    for (const reaction of changes) {
      if (typeof reaction.name !== 'string' || reaction.name.length === 0) continue;
      if (this.emit(activity, { kind: 'reaction', messageId: activity.replyToId, revision: 0, reaction })) {
        emitted += 1;
      }
    }
    return emitted;
  }

  /**
   * An Adaptive Card click (`Action.Execute` invoke or `Action.Submit`
   * message); `replyToId` is the card. For a poll the card id is the poll id
   * and the value `{ optionId }`. `cardAnswer.answererVerified` says whether
   * Teams stated an Entra object id for `from`; an anonymous meeting
   * participant has none and is passed on with `false`, so the engine can
   * refuse to count the answer.
   */
  emitCardAnswer(activity: InboundActivity): boolean {
    const poll = parsePollAnswerValue(activity.value);
    const data = cardActionData(activity.value);
    const answererVerified = typeof activity.from?.aadObjectId === 'string' && activity.from.aadObjectId.length > 0;
    return this.emit(activity, {
      kind: 'card_answer',
      messageId: inboundMessageId(activity),
      revision: 0,
      cardAnswer: poll
        ? { cardId: poll.pollId, value: { optionId: poll.optionId }, answererVerified }
        : { cardId: genericCardId(data), value: data, answererVerified },
    });
  }

  /**
   * One accepted submission of a question card (#2089): typed values per
   * question, with the proposal id and version the card was posted for. The
   * author is `activity.from` of the click, set by Teams, never anything the
   * card carried. `value` repeats the answers for readers that only know it.
   */
  emitQuestionCardAnswer(
    activity: InboundActivity,
    answer: { cardId: string; proposalId?: string; proposalVersion?: number; answers: readonly CardAnswerValue[] },
  ): boolean {
    const answererVerified = typeof activity.from?.aadObjectId === 'string' && activity.from.aadObjectId.length > 0;
    return this.emit(activity, {
      kind: 'card_answer',
      messageId: activity.replyToId ?? activity.id,
      revision: 0,
      cardAnswer: {
        cardId: answer.cardId,
        value: { answers: answer.answers },
        answererVerified,
        ...(answer.proposalId !== undefined ? { proposalId: answer.proposalId } : {}),
        ...(answer.proposalVersion !== undefined ? { proposalVersion: answer.proposalVersion } : {}),
        answers: answer.answers,
      },
    });
  }

  private nextRevision(activity: InboundActivity): number {
    const key = `${receivingBotKey(activity) ?? ''}|${activity.conversation?.id ?? ''}|${activity.id ?? ''}`;
    const next = (this.revisions.get(key) ?? 0) + 1;
    // delete + set keeps insertion order = recency, so the cap evicts the
    // message edited longest ago.
    this.revisions.delete(key);
    this.revisions.set(key, next);
    if (this.revisions.size > MAX_TRACKED_REVISIONS) {
      const oldest = this.revisions.keys().next().value;
      if (oldest !== undefined) this.revisions.delete(oldest);
    }
    return next;
  }

  private emit(
    activity: InboundActivity,
    part: Pick<ConversationMessageEvent, 'kind' | 'revision'> &
      Partial<Pick<ConversationMessageEvent, 'text' | 'replyToMessageId' | 'quoted' | 'forwarded' | 'reaction' | 'cardAnswer'>> & {
        messageId: string | undefined;
      },
  ): boolean {
    try {
      const mode = this.delivery(activity);
      if (mode === 'none') return false;
      const isMessage = part.kind === 'created' || part.kind === 'edited';
      const reason: TeamsAddressedReason | undefined = isMessage
        ? this.addressedReason(activity)
        : part.kind === 'card_answer'
          ? 'card'
          : part.kind === 'reaction'
            ? 'reaction'
            : undefined;
      // `addressed`: a deletion names no addressee and is not handed over.
      if (mode === 'addressed' && reason === undefined) return false;
      const addressedToBot = isMessage && (reason === 'mention' || reason === 'direct');
      const conversationId = activity.conversation?.id;
      const botKey = receivingBotKey(activity);
      const author = platformAuthor(activity);
      const { messageId, ...rest } = part;
      if (!conversationId || !botKey || !author || !messageId) return false;
      const conversationType = toSdkConversationType(activity.conversation?.conversationType);
      const tenantId = tenantOf(activity);
      this.sink.emit({
        channelId: CHANNEL_ID_PLACEHOLDER,
        channelType: CHANNEL_TYPE,
        conversationId,
        ...(conversationType ? { conversationType } : {}),
        receivedByChannelKey: botKey,
        messageId,
        author,
        ...(tenantId ? { tenantId } : {}),
        ...rest,
        ...(addressedToBot ? { addressedToBot: true } : {}),
        occurredAt: occurredAt(activity),
        rawEvent: activity,
      });
      return true;
    } catch (err) {
      console.error('[teams] message event emit failed:', err instanceof Error ? err.message : err);
      return false;
    }
  }
}

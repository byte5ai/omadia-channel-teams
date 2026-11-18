// Inbound message events for native facilitators (byte5ai/omadia#2043 §5.3,
// §6.2; issue #43).
//
// In a conversation an agent moderates, the kernel needs EVERY human message
// with its platform-stated author: also the ones that do not mention the bot,
// and edits, deletions, reactions and card answers. This module turns Bot
// Framework activities into `ConversationMessageEvent`s.
//
// Three rules hold everywhere below:
//   1. Subscription is per bot AND conversation. Bot B in the same chat as a
//      subscribed bot A emits nothing and keeps its mention gate.
//   2. The author is `activity.from`, set by Teams. Nothing is read from the
//      message text, a card payload or a quoted block.
//   3. An event is not a turn. Emitting starts no model call, and a failure
//      to emit never breaks the activity that carried it. In a subscribed
//      conversation the adapter starts no turn at all; the engine does.

import type { Activity } from 'botbuilder';

import type { ConversationMessageEvent, TeamsUserRef } from './kernel-types.js';
import { parseTeamsBotKey, teamsBotKey, TEAMS_BOT_KEY_PREFIX } from './teamsBotIdentity.js';
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

/** The two kernel calls this module needs, already feature-detected. */
export interface TeamsConversationMessageSink {
  isAllMessagesSubscribed(conversationId: string, channelKey: string): boolean;
  emit(event: ConversationMessageEvent): void;
}

type InboundActivity = Partial<Activity>;

/** The normalized `28:<appId>` key of the bot an activity was addressed to. */
export function receivingBotKey(activity: InboundActivity): string | undefined {
  const recipientId = activity.recipient?.id;
  if (typeof recipientId !== 'string') return undefined;
  const appId = parseTeamsBotKey(recipientId);
  return appId ? teamsBotKey(appId) : undefined;
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

/** Quote / forward markers of a message. A quoted reply carries the quoted
 *  message as an HTML attachment with a `schema.skype.com/Reply` blockquote
 *  whose `itemid` is the quoted message id. */
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

  /** Does the bot this activity was addressed to receive every message of
   *  this conversation? `false` for anything not positively subscribed. */
  isSubscribed(activity: InboundActivity): boolean {
    const conversationId = activity.conversation?.id;
    const botKey = receivingBotKey(activity);
    if (!conversationId || !botKey) return false;
    try {
      return this.sink.isAllMessagesSubscribed(conversationId, botKey) === true;
    } catch (err) {
      console.error('[teams] subscription lookup failed:', err instanceof Error ? err.message : err);
      return false;
    }
  }

  /** A new human message (mentioned or not). */
  emitCreated(activity: InboundActivity): boolean {
    const quote = readQuoteInfo(activity);
    return this.emit(activity, {
      kind: 'created',
      messageId: activity.id,
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
   * and the value `{ optionId, answererVerified }`. The answerer is verified
   * when Teams states an Entra object id for `from`; an anonymous meeting
   * participant has none and is passed on with `answererVerified: false`, so
   * the engine can refuse to count the answer.
   */
  emitCardAnswer(activity: InboundActivity): boolean {
    const poll = parsePollAnswerValue(activity.value);
    const data = cardActionData(activity.value);
    const answererVerified = typeof activity.from?.aadObjectId === 'string' && activity.from.aadObjectId.length > 0;
    return this.emit(activity, {
      kind: 'card_answer',
      messageId: activity.replyToId ?? activity.id,
      revision: 0,
      cardAnswer: poll
        ? { cardId: poll.pollId, value: { optionId: poll.optionId, answererVerified } }
        : { cardId: genericCardId(data), value: data },
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
      if (!this.isSubscribed(activity)) return false;
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

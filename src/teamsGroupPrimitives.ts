// #330 B2 — Teams adapters for the channel-SDK group-conversation primitives
// (roster provider, targeted send, conversation-reference capture). All three
// kernel seams are OPTIONAL CoreApi methods (feature-detected in plugin.ts),
// so this plugin keeps activating unchanged on kernels that predate them.

import { TurnContext } from 'botbuilder';
import type { Activity, ConversationReference } from 'botbuilder';
import type {
  ConversationParticipant,
  ConversationRoster,
  ConversationRosterProvider,
  ConversationSendProvider,
  ConversationType,
  TargetedMessage,
  TargetedSendProvider,
} from '@omadia/channel-sdk';

import type { TeamsOpenDirectConversation, TeamsProactiveSend } from './messagesRouter.js';
import {
  buildOutgoingActivity,
  buildPollActivity,
  classifyDirectSendError,
  isEntraObjectId,
  PROACTIVE_POSTS_CARRY_AI_LABEL,
  retryTransient,
  type RetryOptions,
  readPollRequest,
  renderMentions,
} from './teamsOutboundMessage.js';
import { parseTeamsBotKey, TEAMS_BOT_KEY_PREFIX } from './teamsBotIdentity.js';
import {
  bareConversationId,
  botKeyPrefixedConversationMessage,
  isBotKeyPrefixedConversationId,
  reportBotKeyPrefixedConversationId,
} from './teamsConversationIds.js';
import { normalizeTeamsBotAppId } from './teamsConversationRefStore.js';
import type { TeamsConversationRefPersistence } from './teamsConversationRefStore.js';
import { aiLabelEntity } from './teamsCard.js';
import { buildClosedCardContent, buildQuestionCardContent, toCardAttachment } from './teamsQuestionCardRender.js';
import { questionCardProblem, TEAMS_CARD_CAPABILITY } from './teamsQuestionCardSpec.js';
import type { TeamsQuestionCards } from './teamsQuestionCards.js';
import type { TeamsRosterProvider, TeamsRosterRead } from './teamsRoster.js';
import type {
  ChatParticipant,
  QuestionCard,
  CardUpdate,
  CardUpdateOutcome,
  RosterRequestOptions,
  TeamsConversationRoster,
  FacilitatorDeliveryOutcome,
  SendProviderCapabilities,
  TargetedMention,
  TargetedSendOptions,
} from './kernel-types.js';

const MAX_CACHED_REFERENCES = 500;

/** #2043 — key under which a channel's team (AAD group id) rides in the
 *  PERSISTED reference. The reference column is the only state of this
 *  plugin that survives a restart, and the readiness check of a channel
 *  needs the team before any new activity arrives. Stripped again on load:
 *  the in-memory reference stays a plain Bot Framework reference. */
const PERSISTED_TEAM_ID_KEY = 'omadiaTeamAadGroupId';
const AAD_GROUP_ID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Splits a persisted reference into the plain reference and the team id it
 *  may carry. A value that is not an AAD group id is dropped, never used. */
function splitPersistedReference(persisted: Partial<ConversationReference>): {
  ref: Partial<ConversationReference>;
  teamId?: string;
} {
  const { [PERSISTED_TEAM_ID_KEY]: stored, ...ref } = persisted as Record<string, unknown>;
  const teamId = typeof stored === 'string' && AAD_GROUP_ID_PATTERN.test(stored) ? stored : undefined;
  return { ref: ref as Partial<ConversationReference>, ...(teamId ? { teamId } : {}) };
}

/** Teams' native `conversation.conversationType` → the SDK's binary notion.
 *  'personal' is the bot's own 1:1 with a person; 'groupChat' and 'channel'
 *  are groups; unknown stays undefined (consumers treat unknown as direct —
 *  no group semantics without a positive statement).
 *
 *  'group' says NOTHING about how many people are in the chat. A 1:1 between
 *  two people with an agent app installed (`19:…@unq.gbl.spaces`, Graph
 *  `chatType: oneOnOne`) arrives as `groupChat` (seen live, 2026-11-18). Who
 *  is in a chat comes from the roster, never from this value. */
export function toSdkConversationType(teamsType: string | undefined): ConversationType | undefined {
  if (teamsType === 'personal') return 'direct';
  if (teamsType === 'groupChat' || teamsType === 'channel') return 'group';
  return undefined;
}

/**
 * #330 field report — speaker attribution for group turns. A group
 * conversation funnels every participant into ONE conversation-scoped agent
 * session; without attribution the agent cannot tell speakers apart (it
 * literally had to ask "wer von euch bist du?"). The prefix is built from the
 * Bot-Framework activity's verified sender — never from message content — so
 * participants cannot impersonate each other any more convincingly than free
 * text always allowed. 1:1 turns stay untouched.
 */
export function attributeGroupMessage(
  text: string,
  opts: { isGroup: boolean; senderName?: string | undefined },
): string {
  const name = opts.senderName?.trim();
  if (!opts.isGroup || !name) return text;
  return `[${name}]: ${text}`;
}

/**
 * Per-conversation `ConversationReference` cache, fed from every inbound
 * activity (messages AND membership updates). The roster provider needs it to
 * open a proactive TurnContext for a conversation OUTSIDE an active turn —
 * `TeamsInfo.getPagedMembers` only works inside one. Bounded FIFO: at the cap,
 * the oldest conversation is evicted (its roster degrades to "unknown" until
 * the next inbound activity re-captures it).
 */
export class TeamsConversationReferenceCache {
  /** #860 W0a — keyed `${botAppId}|${conversationId}` (bot part normalized
   *  lowercase, `''` when unattributed): one bot's reference for a
   *  conversation must never be served — or evicted — as another bot's. */
  private readonly refs = new Map<string, { ref: Partial<ConversationReference>; teamsType?: string }>();
  /** Last serialization written through per (bot, conversation) — capture()
   *  fires on EVERY inbound activity, the store only needs changes. */
  private readonly lastWritten = new Map<string, string>();
  private persistence?: TeamsConversationRefPersistence;
  private defaultBotAppId?: string;
  /** #2043 — `${botAppId}|${userId}` → the 1:1 conversation of that bot with
   *  that user (both the AAD object id and the Teams id are indexed). */
  private readonly directConversations = new Map<string, string>();
  /** #2043 — the regional connector endpoint each bot was last reached on. */
  private readonly serviceUrls = new Map<string, string>();
  /** #2043 — BARE channel id → AAD group id of its team
   *  (`channelData.team.aadGroupId`); Graph addresses a team by that id.
   *  Keyed without the `;messageid=` thread part: every thread of a channel
   *  belongs to the same team. */
  private readonly teamIds = new Map<string, string>();
  /** #2043 — `${botAppId}|${conversationId}` → did the last proactive post
   *  of that bot into that conversation succeed? */
  private readonly postResults = new Map<string, { ok: boolean; detail?: string }>();

  /** #330 field report — optional write-through store (kernel graph table
   *  `teams_conversation_refs`) so references survive restarts. Writes are
   *  fire-and-forget; a missing table degrades to cache-only behaviour. */
  attachPersistence(persistence: TeamsConversationRefPersistence): void {
    this.persistence = persistence;
  }

  /** #860 W0a — which bot a lookup WITHOUT an explicit `botAppId` means
   *  (the legacy adapters' call shape). Wired by the config layer with
   *  `teams_bots[0].appId`; unset, lookups fall back to the unique-holder
   *  rule below, which is exactly the single-bot behaviour. */
  setDefaultBotAppId(botAppId: string): void {
    this.defaultBotAppId = normalizeTeamsBotAppId(botAppId);
  }

  private static key(conversationId: string, botKeyPart: string): string {
    return `${botKeyPart}|${conversationId}`;
  }

  /** Bot part of the composite key for a lookup: explicit bot, else the
   *  configured default, else — single-bot compatibility — the one bot that
   *  holds a reference for this conversation, IF it is unique. The
   *  unattributed sentinel `''` is a WILDCARD, superseded by exactly one
   *  real app id (a restart-load may seed `''` before the next inbound
   *  activity re-captures under the real bot key — that is one bot, not
   *  two). Two DIFFERENT real bots holding the same conversation with no
   *  explicit choice is ambiguous: no cross-bot guessing, the lookup
   *  misses. */
  private resolveBotKeyPart(conversationId: string, botAppId: string | undefined): string | undefined {
    const explicit = normalizeTeamsBotAppId(botAppId);
    if (explicit !== undefined) return explicit;
    if (this.defaultBotAppId !== undefined) return this.defaultBotAppId;
    const suffix = `|${conversationId}`;
    let sole: string | undefined;
    for (const key of this.refs.keys()) {
      if (!key.endsWith(suffix)) continue;
      const botPart = key.slice(0, key.length - suffix.length);
      // Sentinel entries never make the lookup ambiguous — they are the
      // "bot unknown" placeholder, not a second bot.
      if (botPart === '') continue;
      if (sole !== undefined && sole !== botPart) return undefined;
      sole = botPart;
    }
    return sole ?? '';
  }

  private evictOverflow(): void {
    if (this.refs.size <= MAX_CACHED_REFERENCES) return;
    // Oldest composite entry only — one bot's eviction never drops another
    // bot's reference for the same conversation.
    const oldest = this.refs.keys().next().value;
    if (oldest !== undefined) {
      this.refs.delete(oldest);
      this.lastWritten.delete(oldest);
    }
  }

  capture(context: TurnContext): void {
    const conversationId = context.activity.conversation?.id;
    if (!conversationId) return;
    // The bot the activity was addressed to (`28:<appId>`) — with multiple
    // bots in one tenant this is the identity the reference belongs to.
    const botAppId = normalizeTeamsBotAppId(parseTeamsBotKey(context.activity.recipient?.id ?? ''));
    const key = TeamsConversationReferenceCache.key(conversationId, botAppId ?? '');
    const ref = TurnContext.getConversationReference(context.activity);
    // A real bot attribution supersedes the unattributed sentinel twin a
    // restart-load may have seeded for the same conversation — one bot must
    // occupy one cache slot, not two (and never trip the ambiguity rule).
    if (botAppId !== undefined) {
      const sentinelKey = TeamsConversationReferenceCache.key(conversationId, '');
      this.refs.delete(sentinelKey);
      this.lastWritten.delete(sentinelKey);
    }
    // Map.delete+set keeps insertion order = recency, making the FIFO an LRU.
    this.refs.delete(key);
    const teamsType = context.activity.conversation?.conversationType;
    const teamAadGroupId = (context.activity.channelData as { team?: { aadGroupId?: unknown } } | undefined)?.team
      ?.aadGroupId;
    if (typeof teamAadGroupId === 'string' && teamAadGroupId.length > 0) {
      this.teamIds.set(bareConversationId(conversationId), teamAadGroupId);
    }
    if (botAppId !== undefined) {
      if (typeof ref.serviceUrl === 'string') this.serviceUrls.set(botAppId, ref.serviceUrl);
      if (teamsType === 'personal') {
        const from = context.activity.from;
        this.rememberDirectConversation(botAppId, [from?.aadObjectId, from?.id], conversationId);
      }
    }
    this.refs.set(key, { ref, ...(teamsType ? { teamsType } : {}) });
    if (this.persistence) {
      // The team of a channel is written with the reference so that it is
      // known again after a restart (see PERSISTED_TEAM_ID_KEY).
      const teamId = this.teamIds.get(bareConversationId(conversationId));
      const persistedRef = teamId ? { ...ref, [PERSISTED_TEAM_ID_KEY]: teamId } : ref;
      const serialized = JSON.stringify({ ref: persistedRef, teamsType: teamsType ?? null });
      if (this.lastWritten.get(key) !== serialized) {
        this.lastWritten.set(key, serialized);
        void this.persistence.save(conversationId, persistedRef, teamsType, botAppId).catch(() => {
          // Next capture retries: forgetting the marker keeps save attempts alive.
          this.lastWritten.delete(key);
        });
      }
    }
    this.evictOverflow();
  }

  /** #2043 — remember which conversation is the 1:1 of `botAppId` with a
   *  user. Bounded like the reference cache. */
  rememberDirectConversation(
    botAppId: string,
    userIds: readonly (string | undefined)[],
    conversationId: string,
  ): void {
    const bot = normalizeTeamsBotAppId(botAppId);
    if (bot === undefined) return;
    for (const userId of userIds) {
      if (!userId) continue;
      const key = `${bot}|${userId.toLowerCase()}`;
      this.directConversations.delete(key);
      this.directConversations.set(key, conversationId);
    }
    while (this.directConversations.size > MAX_CACHED_REFERENCES * 2) {
      const oldest = this.directConversations.keys().next().value;
      if (oldest === undefined) break;
      this.directConversations.delete(oldest);
    }
  }

  /** #2043 — the 1:1 reference of exactly this bot with this user, or
   *  `undefined`. Never another bot's conversation with the same user. */
  getDirect(botAppId: string, userId: string): { ref: Partial<ConversationReference>; teamsType?: string } | undefined {
    const bot = normalizeTeamsBotAppId(botAppId);
    if (bot === undefined) return undefined;
    const conversationId = this.directConversations.get(`${bot}|${userId.toLowerCase()}`);
    if (!conversationId) return undefined;
    return this.refs.get(TeamsConversationReferenceCache.key(conversationId, bot));
  }

  /** #2043 — AAD group id of the team a channel conversation belongs to.
   *  Known once an activity from that channel was seen in this process, or
   *  once `getOrLoad` read the persisted reference of that channel (after a
   *  restart no activity has arrived yet). */
  teamIdFor(conversationId: string): string | undefined {
    return this.teamIds.get(bareConversationId(conversationId));
  }

  /** #2043 — the team of a channel or of one of its threads: from memory,
   *  else from the persisted reference of that thread, else from the one of
   *  the bare channel (written when the bot was added to the team). */
  async resolveTeamId(conversationId: string, botAppId?: string): Promise<string | undefined> {
    const known = this.teamIdFor(conversationId);
    if (known) return known;
    await this.getOrLoad(conversationId, botAppId).catch(() => undefined);
    return this.teamIdFor(conversationId);
  }

  /** #2043 — record how a proactive post of `botAppId` into a conversation
   *  ended. "Can this bot post here" has no read-only answer in Teams; the
   *  last real send is the evidence. */
  recordPostResult(conversationId: string, botAppId: string | undefined, ok: boolean, detail?: string): void {
    const bot = normalizeTeamsBotAppId(botAppId) ?? '';
    const key = `${bot}|${conversationId}`;
    this.postResults.delete(key);
    this.postResults.set(key, { ok, ...(detail ? { detail } : {}) });
    if (this.postResults.size > MAX_CACHED_REFERENCES) {
      const oldest = this.postResults.keys().next().value;
      if (oldest !== undefined) this.postResults.delete(oldest);
    }
  }

  lastPostResult(conversationId: string, botAppId: string): { ok: boolean; detail?: string } | undefined {
    return this.postResults.get(`${normalizeTeamsBotAppId(botAppId) ?? ''}|${conversationId}`);
  }

  /** #2043 — the connector endpoint this bot was last reached on. */
  serviceUrlFor(botAppId: string): string | undefined {
    const bot = normalizeTeamsBotAppId(botAppId);
    return bot === undefined ? undefined : this.serviceUrls.get(bot);
  }

  get(conversationId: string, botAppId?: string): { ref: Partial<ConversationReference>; teamsType?: string } | undefined {
    const botKeyPart = this.resolveBotKeyPart(conversationId, botAppId);
    if (botKeyPart === undefined) return undefined;
    return this.refs.get(TeamsConversationReferenceCache.key(conversationId, botKeyPart));
  }

  /** Cache-or-store lookup for the proactive adapters: a miss after a restart
   *  falls back to the persisted reference and re-seeds the cache. */
  async getOrLoad(conversationId: string, botAppId?: string): Promise<{ ref: Partial<ConversationReference>; teamsType?: string } | undefined> {
    const own = await this.getOrLoadExact(conversationId, botAppId);
    if (own) return own;
    // #2043 — a channel THREAD this bot has no reference of its own for yet
    // (nobody addressed it there in this bot's lifetime): the reference of
    // the bare channel, written when the bot was added to the team, leads
    // into the thread once it names the thread as its conversation. Teams
    // posts a message for `<channel>;messageid=<root>` as a reply to <root>.
    const bareId = bareConversationId(conversationId);
    if (bareId === conversationId) return undefined;
    const channel = await this.getOrLoadExact(bareId, botAppId);
    if (!channel) return undefined;
    const { activityId: _channelActivity, ...channelRef } = channel.ref;
    return {
      ...channel,
      ref: { ...channelRef, conversation: { ...channelRef.conversation, id: conversationId } as ConversationReference['conversation'] },
    };
  }

  private async getOrLoadExact(conversationId: string, botAppId?: string): Promise<{ ref: Partial<ConversationReference>; teamsType?: string } | undefined> {
    const cached = this.get(conversationId, botAppId);
    if (cached) return cached;
    if (!this.persistence) return undefined;
    // The store applies the same default-bot semantics on its side (incl.
    // pre-backfill legacy rows), so the raw intent is passed through — not
    // the resolved key part.
    const effective = normalizeTeamsBotAppId(botAppId) ?? this.defaultBotAppId;
    const persisted = await this.persistence.load(conversationId, effective);
    if (!persisted) return undefined;
    // Re-seed under the key capture() would use for this conversation: the
    // OWNING bot of the served row when the store knows it, else the
    // caller's intent, else the unattributed sentinel (which a later
    // capture() supersedes — see resolveBotKeyPart). Seeding under a key
    // capture() never writes would split one bot across two cache slots.
    const owningBotAppId = normalizeTeamsBotAppId(persisted.botAppId) ?? effective;
    const { ref, teamId } = splitPersistedReference(persisted.ref);
    const teamKey = bareConversationId(conversationId);
    if (teamId && !this.teamIds.has(teamKey)) this.teamIds.set(teamKey, teamId);
    const entry = {
      ref,
      ...(persisted.teamsType ? { teamsType: persisted.teamsType } : {}),
    };
    const key = TeamsConversationReferenceCache.key(conversationId, owningBotAppId ?? '');
    this.refs.delete(key);
    this.refs.set(key, entry);
    this.evictOverflow();
    return entry;
  }
}

function toParticipant(member: ChatParticipant): ConversationParticipant {
  return {
    userRef: {
      kind: 'teams-aad',
      id: member.aadObjectId ?? member.channelUserId,
      ...(member.displayName ? { displayName: member.displayName } : {}),
      ...(member.email ? { email: member.email } : {}),
    },
    // Bot/app identities carry the `28:` Bot-Framework prefix on their
    // channel-native id; human members never do.
    isBot: member.channelUserId.startsWith('28:'),
    externalId: member.channelUserId,
    userPrincipalName: member.userPrincipalName,
  };
}

/** The roster provider with the `asChannelKey` option of channel-sdk 0.3.2. */
export interface TeamsRosterAdapter extends ConversationRosterProvider {
  getRoster(conversationId: string, opts?: RosterRequestOptions): Promise<TeamsConversationRoster | undefined>;
}

/** Conversation ids are long; the log shows enough to tell chats and
 *  threads apart (the thread part is at the end). */
function conversationForLog(conversationId: string): string {
  return conversationId.length <= 72 ? conversationId : `${conversationId.slice(0, 40)}…${conversationId.slice(-28)}`;
}

/**
 * SDK roster provider over `TeamsRosterProvider`. Opens a proactive turn on
 * the cached conversation reference; no reference = `undefined` ("cannot
 * resolve", which the kernel treats as unknown), never an invented roster.
 *
 * HOW COMPLETE (kernel spec B §5.5, #2091). Every answer states `status`
 * (`complete`, `partial`, `unknown`), `readAt` and, when not complete, why.
 * A read that failed is `unknown` with no participants, or `partial` with the
 * ones read before the failure; it is never an empty list presented as the
 * roster. `opts.fresh` bypasses the cache; a cached answer keeps the time of
 * its read. `partial` (the older flag) is true unless the status is complete.
 *
 * WHOSE VIEW (#2043). A roster is what ONE bot sees. With
 * `opts.asChannelKey` the reference of exactly that bot is used and the
 * members are read with that bot's credentials; a bot without a reference in
 * the conversation gets `undefined`, never another bot's view. Without it the
 * lookup is the one of #330: the default bot's reference. A facilitator's
 * chat usually does not contain the default bot, which is why its handshake
 * listed nobody before this option existed.
 *
 * Every path logs one line, so an empty or missing roster can be explained
 * from the log.
 */
export function createTeamsRosterAdapter(deps: {
  refs: TeamsConversationReferenceCache;
  roster: Pick<TeamsRosterProvider, 'list'> & Partial<Pick<TeamsRosterProvider, 'read'>>;
  sendProactive: TeamsProactiveSend;
}): TeamsRosterAdapter {
  const log = (conversationId: string, as: string, outcome: string): void => {
    console.error(`[teams-roster] roster conv=${conversationForLog(conversationId)} as=${as} ${outcome}`);
  };
  return {
    channelType: 'teams',
    async getRoster(conversationId: string, opts?: RosterRequestOptions): Promise<TeamsConversationRoster | undefined> {
      const asChannelKey = opts?.asChannelKey;
      const botAppId = asChannelKey !== undefined ? parseTeamsBotKey(asChannelKey) : undefined;
      const as = asChannelKey === undefined ? 'default-bot' : (botAppId ? `${TEAMS_BOT_KEY_PREFIX}${botAppId.slice(0, 8)}…` : 'invalid');
      if (asChannelKey !== undefined && !botAppId) {
        log(conversationId, as, `refused: '${asChannelKey}' is not a Teams bot identity key`);
        return undefined;
      }
      if (isBotKeyPrefixedConversationId(conversationId)) {
        reportBotKeyPrefixedConversationId(conversationId, 'roster');
        log(conversationId, as, 'refused: not a Teams conversation id (bot-key-prefixed)');
        return undefined;
      }
      const cached = await deps.refs.getOrLoad(conversationId, botAppId);
      // Asked as one bot: only that bot's own reference counts.
      const ownerAppId = cached?.ref.bot?.id ? parseTeamsBotKey(cached.ref.bot.id) : undefined;
      if (!cached || (botAppId && ownerAppId && ownerAppId !== botAppId)) {
        log(conversationId, as, 'no reference: this bot has not received an activity from the conversation');
        return undefined;
      }

      // The members of a channel are asked for the CHANNEL; a thread id
      // (`…;messageid=<root>`) is this plugin's and the kernel's key only.
      const bareId = bareConversationId(conversationId);
      const rosterRef: Partial<ConversationReference> =
        bareId === conversationId
          ? cached.ref
          : { ...cached.ref, conversation: { ...cached.ref.conversation, id: bareId } as ConversationReference['conversation'] };

      let read: Pick<TeamsRosterRead, 'status' | 'participants' | 'readAt' | 'httpStatus' | 'detail' | 'fromCache'> | undefined;
      let turnFailure: string | undefined;
      try {
        await deps.sendProactive(
          rosterRef,
          async (turnContext) => {
            if (deps.roster.read) {
              read = await deps.roster.read(turnContext, { fresh: opts?.fresh === true });
              return;
            }
            // A roster source that cannot say how complete its answer is:
            // a non-empty list is taken as given, an empty one is unknown.
            const members = await deps.roster.list(turnContext);
            read = { status: members.length > 0 ? 'complete' : 'unknown', participants: members, readAt: Date.now(), fromCache: false };
          },
          ...(botAppId ? ([{ botAppId }] as const) : ([] as const)),
        );
      } catch (err) {
        // The proactive turn itself failed (credentials, transport). Asked as
        // one bot that is "unknown", not a crash of the caller. The legacy
        // path keeps throwing, as it always did.
        turnFailure = err instanceof Error ? err.message : String(err);
        if (!botAppId) {
          log(conversationId, as, `status=unknown detail=${turnFailure}`);
          throw err;
        }
      }
      const result = read ?? {
        status: 'unknown' as const,
        participants: [],
        readAt: Date.now(),
        fromCache: false,
        detail: turnFailure ?? 'the member list was not read',
      };
      log(
        conversationId,
        as,
        `status=${result.status} count=${String(result.participants.length)}` +
          (result.httpStatus !== undefined ? ` http=${String(result.httpStatus)}` : '') +
          (result.detail ? ` detail=${result.detail}` : '') +
          ` readAt=${new Date(result.readAt).toISOString()}${result.fromCache ? ' (cached)' : ''}${opts?.fresh === true ? ' fresh=true' : ''}`,
      );

      const bot = cached.ref.bot;
      const self: ConversationParticipant | undefined = bot
        ? {
            userRef: { kind: 'teams-aad', id: bot.id, ...(bot.name ? { displayName: bot.name } : {}) },
            isBot: true,
            externalId: bot.id,
            userPrincipalName: null,
          }
        : undefined;

      const statusDetail =
        result.status === 'complete'
          ? undefined
          : `${result.httpStatus !== undefined ? `HTTP ${String(result.httpStatus)}: ` : ''}${result.detail ?? 'the member list could not be read'}`;
      return {
        conversationType: toSdkConversationType(cached.teamsType) ?? 'direct',
        participants: result.participants.map(toParticipant),
        ...(self ? { self } : {}),
        // `partial` is the flag kernels before channel-sdk 0.4 read: nothing
        // may be concluded from an absence unless the read is complete.
        partial: result.status !== 'complete',
        status: result.status,
        readAt: new Date(result.readAt).toISOString(),
        ...(statusDetail ? { statusDetail } : {}),
      };
    },
  };
}

/** The posted message of a question card: the card, and the AI label every
 *  message under an agent's name carries (one marking per message). */
function buildQuestionCardActivity(card: QuestionCard): Partial<Activity> {
  return {
    type: 'message',
    attachments: [toCardAttachment(buildQuestionCardContent(card))],
    entities: [aiLabelEntity()],
  };
}

/**
 * The declared capabilities, built by functions and not as object literals at
 * the provider: the plugin is compiled against whatever `@omadia/channel-sdk`
 * is installed, and a literal with a key that SDK does not know yet
 * (`aiLabel` before 0.3.2) is an excess-property error there. 0.36.6 did not
 * typecheck against the kernel's main branch for exactly that reason. The
 * values are this plugin's mirror type, which is a superset of every SDK's.
 */
function targetedCapabilities(hasDirectPath: boolean): SendProviderCapabilities {
  return { asChannelKey: hasDirectPath, aiLabel: PROACTIVE_POSTS_CARRY_AI_LABEL };
}

function conversationCapabilities(hasRoster: boolean, hasCards: boolean): SendProviderCapabilities {
  return {
    asChannelKey: true,
    mentions: hasRoster,
    poll: true,
    aiLabel: PROACTIVE_POSTS_CARRY_AI_LABEL,
    // Question cards: the kinds really rendered, and `updateCard`.
    ...(hasCards ? { card: TEAMS_CARD_CAPABILITY } : {}),
  };
}

/** What a DM "as bot X" needs beyond the legacy path (#2043). */
export interface TeamsDirectSendDeps {
  refs: Pick<TeamsConversationReferenceCache, 'getDirect' | 'rememberDirectConversation' | 'serviceUrlFor'>;
  openDirectConversation: TeamsOpenDirectConversation;
  /** Display name of a CONFIGURED bot; `undefined` for an unknown app id. */
  botLabel(botAppId: string): string | undefined;
}

/**
 * SDK targeted-send provider (#330 B3, #2043).
 *
 * Without `asChannelKey` the behaviour is the one of #330: the kernel hands
 * over the cached per-user conversation reference, and no reference means
 * `no_binding`.
 *
 * With `asChannelKey` (a facilitator's private reminder) the DM comes from
 * exactly that bot or not at all:
 *   1. the kernel's reference, if it belongs to that bot;
 *   2. the 1:1 this bot already has with the user (captured on the user's
 *      own messages to it);
 *   3. first contact: the bot opens the 1:1 itself, which Teams allows when
 *      its app is installed for the user. This is also how the 1:1 is found
 *      again after a restart.
 * `app_not_installed` and `blocked` are returned as outcomes so the caller
 * can fall back to a neutral reminder in the group. Another bot's reference
 * with the same user is never used.
 */
export function createTeamsTargetedSendAdapter(deps: {
  sendProactive: TeamsProactiveSend;
  direct?: TeamsDirectSendDeps;
  retry?: RetryOptions;
}): TargetedSendProvider {
  type Target = Parameters<TargetedSendProvider['sendToUser']>[0];

  const sendOnReference = async (
    reference: Partial<ConversationReference>,
    text: string,
    botAppId: string | undefined,
  ): Promise<string | undefined> => {
    let messageId: string | undefined;
    await deps.sendProactive(
      reference,
      async (turnContext) => {
        const sent = await retryTransient(
          () => turnContext.sendActivity(buildOutgoingActivity({ text })),
          deps.retry,
        );
        messageId = sent?.id;
      },
      ...(botAppId ? ([{ botAppId }] as const) : ([] as const)),
    );
    return messageId;
  };

  const delivered = (messageId: string | undefined): FacilitatorDeliveryOutcome => ({
    outcome: 'delivered',
    ...(messageId ? { messageId } : {}),
  });

  const sendAsBot = async (
    target: Target,
    text: string,
    asChannelKey: string,
  ): Promise<FacilitatorDeliveryOutcome> => {
    const botAppId = parseTeamsBotKey(asChannelKey);
    const botLabel = botAppId ? deps.direct?.botLabel(botAppId) : undefined;
    if (!botAppId || !deps.direct || botLabel === undefined) {
      return {
        outcome: 'unreachable',
        code: 'not_permitted',
        message: botAppId
          ? 'the requested sender is not a bot configured in the Teams channel - refusing to send the direct message as another bot'
          : `'${asChannelKey}' is not a Teams bot identity key (expected '${TEAMS_BOT_KEY_PREFIX}<appId>') - refusing to send under an unresolved sender`,
      };
    }
    const direct = deps.direct;
    // The Entra object id only: a Teams user id (`29:…`) is per bot and
    // would address nobody for this one.
    const userId =
      target.userRef?.kind === 'teams-aad' && isEntraObjectId(target.userRef.id) ? target.userRef.id : undefined;

    const kernelRef = target.conversationRef as Partial<ConversationReference> | null | undefined;
    const kernelRefBot = kernelRef?.bot?.id ? parseTeamsBotKey(kernelRef.bot.id) : undefined;
    const known =
      (kernelRef && kernelRefBot === botAppId ? kernelRef : undefined) ??
      (userId ? direct.refs.getDirect(botAppId, userId)?.ref : undefined);
    try {
      if (known) return delivered(await sendOnReference(known, text, botAppId));
      if (!userId) {
        return {
          outcome: 'unreachable',
          code: 'no_binding',
          message: `no Entra object id is known for '${target.principalId}', so '${botLabel}' cannot open a direct conversation`,
        };
      }
      let messageId: string | undefined;
      let isOpened = false;
      const serviceUrl = direct.refs.serviceUrlFor(botAppId);
      const open = (): Promise<void> =>
        direct.openDirectConversation(
          botAppId,
          { userId, ...(serviceUrl ? { serviceUrl } : {}) },
          async (turnContext) => {
            isOpened = true;
            const conversationId = turnContext.activity.conversation?.id;
            if (conversationId) direct.refs.rememberDirectConversation(botAppId, [userId], conversationId);
            const sent = await retryTransient(
              () => turnContext.sendActivity(buildOutgoingActivity({ text })),
              deps.retry,
            );
            messageId = sent?.id;
          },
        );
      // Creating the conversation is retried while nothing was sent; once the
      // conversation is open only the send itself is (above), so a retry can
      // never post the message twice.
      await retryTransient(open, deps.retry, () => !isOpened);
      return delivered(messageId);
    } catch (err) {
      const outcome = classifyDirectSendError(err, botLabel);
      // Log-safe: the bot's display name and the outcome code, no ids.
      console.error(`[teams] direct message as '${botLabel}' not delivered: ${outcome.code}`);
      return outcome;
    }
  };

  return {
    channelType: 'teams',
    // Declared, because the kernel only hands an own-identity DM to a
    // provider that says it honours `asChannelKey`. True only when the per-bot
    // path is wired: without it a named sender is refused, never substituted,
    // but there would be nothing to deliver with either.
    //   aiLabel       every activity of this provider is built by
    //                 `buildOutgoingActivity`, which attaches the label
    supports: targetedCapabilities(deps.direct !== undefined),
    // Rest tuple for the same reason as `sendToConversation` below: the third
    // argument exists from channel-sdk 0.3.0 on.
    async sendToUser(
      ...args: [target: Target, message: TargetedMessage, opts?: TargetedSendOptions]
    ): Promise<FacilitatorDeliveryOutcome> {
      const [target, message, opts] = args;
      if (opts?.asChannelKey !== undefined) return sendAsBot(target, message.text, opts.asChannelKey);
      if (target.conversationRef === undefined || target.conversationRef === null) {
        return {
          outcome: 'unreachable',
          code: 'no_binding',
          message: `no Teams conversation reference known for '${target.principalId}' — the user has not talked to this bot yet`,
        };
      }
      try {
        return delivered(
          await sendOnReference(target.conversationRef as Partial<ConversationReference>, message.text, undefined),
        );
      } catch (err) {
        return {
          outcome: 'unreachable',
          code: 'channel_error',
          message: err instanceof Error ? err.message : String(err),
        };
      }
    },
  };
}


/**
 * The contract's delivery outcome for a group post. `cardMessageId` is the
 * card's own message, the id poll answers name. One addition is not in the
 * contract: an `unreachable` outcome carries `messageId` when the text was
 * posted and only the card failed, so the caller does not post the text twice.
 */
export type ConversationPostOutcome =
  | Extract<FacilitatorDeliveryOutcome, { outcome: 'delivered' }>
  | (Extract<FacilitatorDeliveryOutcome, { outcome: 'unreachable' }> & { messageId?: string });

/**
 * SDK conversation-send provider (#330 C3b - the Facilitator's group nudges).
 * Delivers INTO a conversation via the proactive path on the cached
 * per-conversation reference. The kernel has already scope-checked the caller
 * (own ephemeral attachment only); this adapter only knows how to deliver.
 * No cached reference = named unreachable outcome, never a throw.
 */
export function createTeamsConversationSendAdapter(deps: {
  refs: TeamsConversationReferenceCache;
  sendProactive: TeamsProactiveSend;
  /** #2043 — resolves a mentioned person to their member id in the chat.
   *  Without it mentions stay plain text. */
  roster?: Pick<TeamsRosterProvider, 'list'>;
  /** #2089 — the cards this plugin posted. With it the provider renders
   *  question cards and declares `supports.card`; without it a card is
   *  refused. The bot is given the same instance to answer the clicks. */
  cards?: TeamsQuestionCards;
  retry?: RetryOptions;
  // The declared return type is a SUPERSET of the SDK interface, for the same
  // reason the send method takes a rest tuple: this plugin ships independently
  // and is compiled against whatever `@omadia/channel-sdk` is installed, which
  // may predate `sendTyping`. Declaring the extra method here keeps the object
  // assignable to the narrower contract without a cast and without a version
  // floor; a kernel that does not know about it simply never calls it.
}): ConversationSendProvider & {
  supports: SendProviderCapabilities;
  sendTyping(conversationId: string, opts?: { asChannelKey?: string }): Promise<void>;
  updateCard(
    conversationId: string,
    cardMessageId: string,
    update: CardUpdate,
    opts: { asChannelKey: string },
  ): Promise<CardUpdateOutcome>;
} {
  return {
    channelType: 'teams',
    // Declared, because the kernel only hands an own-identity post, mentions
    // or a poll to a provider that says it implements them:
    //   asChannelKey  the post goes out as exactly that bot or not at all
    //   mentions      resolved against the chat roster, so only with one
    //   poll          posted as a card; an invalid poll is refused
    // Text and poll card are built by `buildOutgoingActivity` /
    // `buildPollActivity`, which attach the AI label.
    supports: conversationCapabilities(deps.roster !== undefined, deps.cards !== undefined),
    // Rest tuple, not three named parameters, ON PURPOSE: the plugin ships
    // independently of the kernel and is compiled against whatever
    // `@omadia/channel-sdk` is installed. On an SDK that predates
    // `ConversationSendOptions` a three-parameter method is not assignable to
    // the two-parameter contract, while this shape is — so the plugin
    // implements the superset without a cast and without a version floor. On
    // an older kernel the third argument simply never arrives and delivery
    // behaves exactly as before.
    async sendToConversation(
      ...args: [conversationId: string, message: TargetedMessage, opts?: { asChannelKey?: string }]
    ): Promise<ConversationPostOutcome> {
      const [conversationId, message, opts] = args;
      // WHICH BOT SAYS IT.
      //
      // Several provisioned bots share one group chat. Without an explicit
      // identity this falls to the reference's owning bot, and for the kernel's
      // agent dialogue that would put one agent's words under another bot's
      // name and avatar — indistinguishable, in the chat, from that agent
      // having said them. `asChannelKey` is the caller stating the sender, and
      // an unresolvable one is REFUSED rather than silently substituted: a
      // wrong sender is worse than no message.
      const asChannelKey = opts?.asChannelKey;
      let botAppId: string | undefined;
      if (asChannelKey !== undefined) {
        botAppId = parseTeamsBotKey(asChannelKey);
        if (!botAppId) {
          return {
            outcome: 'unreachable',
            code: 'not_permitted',
            message: `'${asChannelKey}' is not a Teams bot identity key (expected '${TEAMS_BOT_KEY_PREFIX}<appId>') - refusing to send under an unresolved sender`,
          };
        }
      }

      // #2043 — a session-scope fragment instead of a conversation id is a
      // caller bug; say so instead of "the bot was not added to that chat".
      if (isBotKeyPrefixedConversationId(conversationId)) {
        reportBotKeyPrefixedConversationId(conversationId, 'conversation send');
        return { outcome: 'unreachable', code: 'no_binding', message: botKeyPrefixedConversationMessage(conversationId) };
      }
      const cached = await deps.refs.getOrLoad(conversationId, botAppId);
      if (!cached) {
        return {
          outcome: 'unreachable',
          code: 'no_binding',
          message: botAppId
            ? `bot '${botAppId}' has no Teams conversation reference for '${conversationId}' - it has not been added to that chat yet`
            : `no Teams conversation reference cached for '${conversationId}' - no inbound activity seen yet`,
        };
      }

      // A reference served for another bot must not be sent through the
      // requested one: the Bot-Framework would deliver it under whichever
      // identity the reference carries, which is exactly the substitution the
      // caller asked us to prevent.
      const refBotAppId = cached.ref.bot?.id ? parseTeamsBotKey(cached.ref.bot.id) : undefined;
      if (botAppId && refBotAppId && refBotAppId !== botAppId) {
        return {
          outcome: 'unreachable',
          code: 'not_permitted',
          message: `the cached reference for '${conversationId}' belongs to bot '${refBotAppId}', not '${botAppId}' - refusing to speak under another bot's identity`,
        };
      }

      // #2043 — a poll that cannot be rendered is refused before anything is
      // posted: half a poll in the chat is worse than a named failure.
      const pollRequest = readPollRequest(message.poll);
      if (pollRequest && 'problem' in pollRequest) {
        return { outcome: 'unreachable', code: 'channel_error', message: `invalid poll: ${pollRequest.problem}` };
      }
      // #2089 — a question card. `message.card` is read structurally: the
      // installed SDK may not know the field yet. Refused before anything is
      // posted when it cannot be rendered, like a poll.
      const requestedCard = (message as { card?: unknown }).card;
      let questionCard: QuestionCard | undefined;
      if (requestedCard !== undefined && requestedCard !== null) {
        if (!deps.cards) {
          return { outcome: 'unreachable', code: 'channel_error', message: 'question cards are not wired in this Teams channel' };
        }
        if (pollRequest) {
          return { outcome: 'unreachable', code: 'channel_error', message: 'a message carries a poll or a question card, not both' };
        }
        const cardProblem = questionCardProblem(requestedCard);
        if (cardProblem) {
          return { outcome: 'unreachable', code: 'channel_error', message: `invalid question card: ${cardProblem}` };
        }
        questionCard = requestedCard as QuestionCard;
      }
      const mentions: readonly TargetedMention[] | undefined = message.mentions;

      // Text and card go out as SEPARATE activities and each returns its own
      // id (a combined activity may be split by Teams and then returns none).
      // `messageId` is the text message, or the card when there is no text;
      // `cardMessageId` is the card, the id poll answers refer to.
      let textMessageId: string | undefined;
      let cardMessageId: string | undefined;
      let isTextPosted = false;
      const hasText = message.text.trim().length > 0;
      const senderAppId = botAppId ?? refBotAppId;
      const proactiveOpts = botAppId ? ([{ botAppId }] as const) : ([] as const);
      const wantsRoster = hasText && mentions !== undefined && mentions.length > 0 && deps.roster !== undefined;
      // In a channel THREAD the members are those of the channel: asked in a
      // turn of its own on the bare channel id, before anything is posted.
      // A roster that cannot be read leaves the mentions as plain text.
      let channelMembers: ChatParticipant[] | undefined;
      const bareId = bareConversationId(conversationId);
      if (wantsRoster && bareId !== conversationId) {
        const roster = deps.roster;
        channelMembers = [];
        await deps
          .sendProactive(
            { ...cached.ref, conversation: { ...cached.ref.conversation, id: bareId } as ConversationReference['conversation'] },
            async (turnContext) => {
              channelMembers = roster ? await roster.list(turnContext) : [];
            },
            ...proactiveOpts,
          )
          .catch((err: unknown) => {
            console.error(
              `[teams-roster] mention roster conv=${conversationForLog(conversationId)} refused detail=${err instanceof Error ? err.message : String(err)}`,
            );
          });
      }
      try {
        await deps.sendProactive(
          cached.ref,
          async (turnContext) => {
            if (hasText) {
              const members =
                channelMembers ?? (wantsRoster && deps.roster ? await deps.roster.list(turnContext) : []);
              // Resolved from the roster by Entra object id: only a member of
              // this chat can be mentioned, and the entity text equals the
              // `<at>…</at>` token in the body exactly.
              const rendered = renderMentions(message.text, mentions, (userId) => {
                const wanted = userId.toLowerCase();
                return members.find((member) => member.aadObjectId?.toLowerCase() === wanted)?.channelUserId;
              });
              const sent = await retryTransient(
                () => turnContext.sendActivity(buildOutgoingActivity({ text: rendered.text, entities: rendered.entities })),
                deps.retry,
              );
              textMessageId = sent?.id;
              isTextPosted = true;
            }
            if (pollRequest) {
              const sent = await retryTransient(
                () => turnContext.sendActivity(buildPollActivity(pollRequest.poll)),
                deps.retry,
              );
              cardMessageId = sent?.id;
            }
            if (questionCard) {
              const card = questionCard;
              const sent = await retryTransient(() => turnContext.sendActivity(buildQuestionCardActivity(card)), deps.retry);
              cardMessageId = sent?.id;
            }
          },
          ...(botAppId ? ([{ botAppId }] as const) : ([] as const)),
        );
        if (questionCard && senderAppId) {
          // Remembered so a click is checked against what was posted where.
          deps.cards?.register({ card: questionCard, conversationId, botAppId: senderAppId, messageId: cardMessageId });
        }
        deps.refs.recordPostResult(conversationId, senderAppId, true);
        const messageId = textMessageId ?? cardMessageId;
        return {
          outcome: 'delivered',
          ...(messageId ? { messageId } : {}),
          ...(cardMessageId ? { cardMessageId } : {}),
        };
      } catch (err) {
        const detail = err instanceof Error ? err.message : String(err);
        deps.refs.recordPostResult(conversationId, senderAppId, false, detail);
        return {
          outcome: 'unreachable',
          code: 'channel_error',
          // Never silent about half a post: the caller must not send the
          // text a second time when only the card failed.
          message: isTextPosted ? `the text was posted but the card was not: ${detail}` : detail,
          ...(isTextPosted && textMessageId ? { messageId: textMessageId } : {}),
        };
      }
    },

    /**
     * #2089 — change a posted card: into its closed form (the engine's
     * result, no inputs) or into a corrected open card with the same id.
     * `PUT …/activities/{id}` on the stored activity id, as the bot that
     * posted it, and only for a message that bot posted as a question card
     * in that conversation. A failed or refused update is reported and the
     * engine posts its result as text, which it does anyway.
     */
    async updateCard(conversationId, cardMessageId, update, opts): Promise<CardUpdateOutcome> {
      const cards = deps.cards;
      if (!cards) return { outcome: 'unreachable', code: 'channel_error', message: 'question cards are not wired in this Teams channel' };
      const botAppId = parseTeamsBotKey(opts?.asChannelKey ?? '');
      if (!botAppId) {
        return {
          outcome: 'unreachable',
          code: 'not_permitted',
          message: `'${String(opts?.asChannelKey)}' is not a Teams bot identity key (expected '${TEAMS_BOT_KEY_PREFIX}<appId>')`,
        };
      }
      if (typeof cardMessageId !== 'string' || cardMessageId.length === 0) {
        return { outcome: 'unreachable', code: 'channel_error', message: 'the id of the posted card is missing' };
      }
      if (isBotKeyPrefixedConversationId(conversationId)) {
        reportBotKeyPrefixedConversationId(conversationId, 'card update');
        return { outcome: 'unreachable', code: 'no_binding', message: botKeyPrefixedConversationMessage(conversationId) };
      }
      // ONLY a question card that THIS bot posted, in THIS conversation. The
      // kernel keeps no record of posted cards and checks identity and
      // subscription only, so this is the one place that knows: an id that
      // is another bot's card, a text message, or anything this process did
      // not post as a card is refused. (Teams would refuse another bot's
      // message itself, but would happily turn this bot's own text message
      // into a card.)
      const posted = cards.findByMessage(botAppId, cardMessageId);
      if (!posted) {
        return {
          outcome: 'unreachable',
          code: 'not_permitted',
          message: `'${cardMessageId}' is not a question card that bot '${botAppId}' posted (cards are remembered per bot until the Teams channel restarts)`,
        };
      }
      if (bareConversationId(posted.conversationId) !== bareConversationId(conversationId)) {
        return {
          outcome: 'unreachable',
          code: 'not_permitted',
          message: `the question card '${posted.card.cardId}' was posted into another conversation than '${conversationId}'`,
        };
      }
      let activity: Partial<Activity>;
      if (update.state === 'open') {
        const problem = questionCardProblem(update.card);
        if (problem) return { outcome: 'unreachable', code: 'channel_error', message: `invalid question card: ${problem}` };
        if (update.card.cardId !== posted.card.cardId) {
          return {
            outcome: 'unreachable',
            code: 'channel_error',
            message: `a corrected card keeps its id: the posted card is '${posted.card.cardId}', the update names '${update.card.cardId}'`,
          };
        }
        activity = buildQuestionCardActivity(update.card);
      } else {
        activity = {
          type: 'message',
          attachments: [
            toCardAttachment(
              buildClosedCardContent({
                title: update.title,
                resultLines: update.resultLines,
                closedAt: update.closedAt,
                ...(posted.card.locale ? { locale: posted.card.locale } : {}),
              }),
            ),
          ],
          entities: [aiLabelEntity()],
        };
      }
      const cached = await deps.refs.getOrLoad(conversationId, botAppId).catch(() => undefined);
      const ownerAppId = cached?.ref.bot?.id ? parseTeamsBotKey(cached.ref.bot.id) : undefined;
      if (!cached || (ownerAppId && ownerAppId !== botAppId)) {
        return {
          outcome: 'unreachable',
          code: 'no_binding',
          message: `bot '${botAppId}' has no Teams conversation reference for '${conversationId}'`,
        };
      }
      try {
        await deps.sendProactive(
          cached.ref,
          async (turnContext) => {
            await retryTransient(() => turnContext.updateActivity({ ...activity, id: cardMessageId }), deps.retry);
          },
          { botAppId },
        );
      } catch (err) {
        return { outcome: 'unreachable', code: 'channel_error', message: err instanceof Error ? err.message : String(err) };
      }
      if (update.state === 'open') {
        cards.register({ card: update.card, conversationId, botAppId, messageId: cardMessageId });
      } else {
        cards.markClosed(botAppId, posted.card.cardId, {
          title: update.title,
          resultLines: update.resultLines,
          closedAt: update.closedAt,
          ...(posted.card.locale ? { locale: posted.card.locale } : {}),
        });
      }
      return { outcome: 'updated' };
    },

    // Best-effort by contract, and silent by design: an activity indicator that
    // fails must never disturb — let alone fail — the work it was announcing.
    // Resolved through the same per-bot path as a message, so the dots appear
    // next to the bot that is actually thinking rather than next to whichever
    // bot happens to hold the reference.
    async sendTyping(
      ...args: [conversationId: string, opts?: { asChannelKey?: string }]
    ): Promise<void> {
      const [conversationId, opts] = args;
      const botAppId = opts?.asChannelKey ? parseTeamsBotKey(opts.asChannelKey) : undefined;
      // A named-but-unresolvable sender shows nothing rather than showing the
      // wrong bot as busy — the same refusal the message path makes.
      if (opts?.asChannelKey !== undefined && !botAppId) return;
      const cached = await deps.refs.getOrLoad(conversationId, botAppId).catch(() => undefined);
      if (!cached) return;
      const refBotAppId = cached.ref.bot?.id ? parseTeamsBotKey(cached.ref.bot.id) : undefined;
      if (botAppId && refBotAppId && refBotAppId !== botAppId) return;
      try {
        await deps.sendProactive(
          cached.ref,
          async (turnContext) => {
            await turnContext.sendActivity({ type: 'typing' });
          },
          ...(botAppId ? ([{ botAppId }] as const) : ([] as const)),
        );
      } catch {
        // Deliberately swallowed — see the contract note above.
      }
    },
  };
}

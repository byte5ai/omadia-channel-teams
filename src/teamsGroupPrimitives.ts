// #330 B2 — Teams adapters for the channel-SDK group-conversation primitives
// (roster provider, targeted send, conversation-reference capture). All three
// kernel seams are OPTIONAL CoreApi methods (feature-detected in plugin.ts),
// so this plugin keeps activating unchanged on kernels that predate them.

import { TurnContext } from 'botbuilder';
import type { ConversationReference } from 'botbuilder';
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
  retryTransient,
  type RetryOptions,
  readPollRequest,
  renderMentions,
} from './teamsOutboundMessage.js';
import { parseTeamsBotKey, TEAMS_BOT_KEY_PREFIX } from './teamsBotIdentity.js';
import { normalizeTeamsBotAppId } from './teamsConversationRefStore.js';
import type { TeamsConversationRefPersistence } from './teamsConversationRefStore.js';
import type { TeamsRosterProvider } from './teamsRoster.js';
import type {
  ChatParticipant,
  FacilitatorDeliveryOutcome,
  TargetedMention,
  TargetedSendOptions,
} from './kernel-types.js';

const MAX_CACHED_REFERENCES = 500;

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
  /** #2043 — channel conversation → AAD group id of its team
   *  (`channelData.team.aadGroupId`); Graph addresses a team by that id. */
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
      this.teamIds.set(conversationId, teamAadGroupId);
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
      const serialized = JSON.stringify({ ref, teamsType: teamsType ?? null });
      if (this.lastWritten.get(key) !== serialized) {
        this.lastWritten.set(key, serialized);
        void this.persistence.save(conversationId, ref, teamsType, botAppId).catch(() => {
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

  /** #2043 — AAD group id of the team a channel conversation belongs to,
   *  known once an activity from that channel was seen in this process. */
  teamIdFor(conversationId: string): string | undefined {
    return this.teamIds.get(conversationId);
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
    const entry = {
      ref: persisted.ref,
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

/**
 * SDK roster provider over the existing `TeamsRosterProvider` (5-min cache,
 * RSC-permission soft-fail). Opens a proactive turn on the cached
 * conversation reference; no reference = "roster unknown" (undefined), never
 * an invented roster. An EMPTY member list is reported `partial: true`: every
 * real Teams conversation has at least the requesting bot as a member, so
 * emptiness here is the roster provider's permission/transport soft-fail, not
 * an answer.
 */
export function createTeamsRosterAdapter(deps: {
  refs: TeamsConversationReferenceCache;
  roster: TeamsRosterProvider;
  sendProactive: TeamsProactiveSend;
}): ConversationRosterProvider {
  return {
    channelType: 'teams',
    async getRoster(conversationId: string): Promise<ConversationRoster | undefined> {
      const cached = await deps.refs.getOrLoad(conversationId);
      if (!cached) return undefined;

      let members: ChatParticipant[] = [];
      await deps.sendProactive(cached.ref, async (turnContext) => {
        members = await deps.roster.list(turnContext);
      });

      const bot = cached.ref.bot;
      const self: ConversationParticipant | undefined = bot
        ? {
            userRef: { kind: 'teams-aad', id: bot.id, ...(bot.name ? { displayName: bot.name } : {}) },
            isBot: true,
            externalId: bot.id,
            userPrincipalName: null,
          }
        : undefined;

      return {
        conversationType: toSdkConversationType(cached.teamsType) ?? 'direct',
        participants: members.map(toParticipant),
        ...(self ? { self } : {}),
        partial: members.length === 0,
      };
    },
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
  retry?: RetryOptions;
  // The declared return type is a SUPERSET of the SDK interface, for the same
  // reason the send method takes a rest tuple: this plugin ships independently
  // and is compiled against whatever `@omadia/channel-sdk` is installed, which
  // may predate `sendTyping`. Declaring the extra method here keeps the object
  // assignable to the narrower contract without a cast and without a version
  // floor; a kernel that does not know about it simply never calls it.
}): ConversationSendProvider & {
  sendTyping(conversationId: string, opts?: { asChannelKey?: string }): Promise<void>;
} {
  return {
    channelType: 'teams',
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
      try {
        await deps.sendProactive(
          cached.ref,
          async (turnContext) => {
            if (hasText) {
              const members =
                mentions && mentions.length > 0 && deps.roster ? await deps.roster.list(turnContext) : [];
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
          },
          ...(botAppId ? ([{ botAppId }] as const) : ([] as const)),
        );
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

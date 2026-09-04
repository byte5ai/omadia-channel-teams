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
  TargetedDeliveryOutcome,
  TargetedMessage,
  TargetedSendProvider,
} from '@omadia/channel-sdk';

import type { TeamsProactiveSend } from './messagesRouter.js';
import { parseTeamsBotKey, TEAMS_BOT_KEY_PREFIX } from './teamsBotIdentity.js';
import { normalizeTeamsBotAppId } from './teamsConversationRefStore.js';
import type { TeamsConversationRefPersistence } from './teamsConversationRefStore.js';
import type { TeamsRosterProvider } from './teamsRoster.js';
import type { ChatParticipant } from './kernel-types.js';

const MAX_CACHED_REFERENCES = 500;

/** Teams' native `conversation.conversationType` → the SDK's binary notion.
 *  'personal' is a 1:1; 'groupChat' and 'channel' are groups; unknown stays
 *  undefined (consumers treat unknown as direct — no group semantics without
 *  a positive statement). */
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

/**
 * SDK targeted-send provider (#330 B3). The kernel resolves the Principal and
 * hands over ONE user plus the cached per-user conversation reference from the
 * Conductor channel-binding store (captured on the user's inbound turns). No
 * reference = `no_binding` — creating a brand-new 1:1 conversation requires
 * the Teams app to be installed for that user and is a deliberate non-goal of
 * this slice; the outcome names the gap instead of hiding it.
 */
export function createTeamsTargetedSendAdapter(deps: {
  sendProactive: TeamsProactiveSend;
}): TargetedSendProvider {
  return {
    channelType: 'teams',
    async sendToUser(target, message) {
      if (target.conversationRef === undefined || target.conversationRef === null) {
        return {
          outcome: 'unreachable',
          code: 'no_binding',
          message: `no Teams conversation reference known for '${target.principalId}' — the user has not talked to this bot yet`,
        };
      }
      try {
        await deps.sendProactive(target.conversationRef as Partial<ConversationReference>, async (turnContext) => {
          await turnContext.sendActivity({ type: 'message', text: message.text });
        });
        return { outcome: 'delivered' };
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
 * SDK conversation-send provider (#330 C3b - the Facilitator's group nudges).
 * Delivers INTO a conversation via the proactive path on the cached
 * per-conversation reference. The kernel has already scope-checked the caller
 * (own ephemeral attachment only); this adapter only knows how to deliver.
 * No cached reference = named unreachable outcome, never a throw.
 */
export function createTeamsConversationSendAdapter(deps: {
  refs: TeamsConversationReferenceCache;
  sendProactive: TeamsProactiveSend;
}): ConversationSendProvider {
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
    ): Promise<TargetedDeliveryOutcome> {
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

      try {
        await deps.sendProactive(
          cached.ref,
          async (turnContext) => {
            await turnContext.sendActivity({ type: 'message', text: message.text });
          },
          ...(botAppId ? ([{ botAppId }] as const) : ([] as const)),
        );
        return { outcome: 'delivered' };
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

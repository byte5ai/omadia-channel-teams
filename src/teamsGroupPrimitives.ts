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
  TargetedSendProvider,
} from '@omadia/channel-sdk';

import type { TeamsProactiveSend } from './messagesRouter.js';
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
  private readonly refs = new Map<string, { ref: Partial<ConversationReference>; teamsType?: string }>();
  /** Last serialization written through per conversation — capture() fires on
   *  EVERY inbound activity, the store only needs changes. */
  private readonly lastWritten = new Map<string, string>();
  private persistence?: TeamsConversationRefPersistence;

  /** #330 field report — optional write-through store (kernel graph table
   *  `teams_conversation_refs`) so references survive restarts. Writes are
   *  fire-and-forget; a missing table degrades to cache-only behaviour. */
  attachPersistence(persistence: TeamsConversationRefPersistence): void {
    this.persistence = persistence;
  }

  capture(context: TurnContext): void {
    const conversationId = context.activity.conversation?.id;
    if (!conversationId) return;
    const ref = TurnContext.getConversationReference(context.activity);
    // Map.delete+set keeps insertion order = recency, making the FIFO an LRU.
    this.refs.delete(conversationId);
    const teamsType = context.activity.conversation?.conversationType;
    this.refs.set(conversationId, { ref, ...(teamsType ? { teamsType } : {}) });
    if (this.persistence) {
      const serialized = JSON.stringify({ ref, teamsType: teamsType ?? null });
      if (this.lastWritten.get(conversationId) !== serialized) {
        this.lastWritten.set(conversationId, serialized);
        void this.persistence.save(conversationId, ref, teamsType).catch(() => {
          // Next capture retries: forgetting the marker keeps save attempts alive.
          this.lastWritten.delete(conversationId);
        });
      }
    }
    if (this.refs.size > MAX_CACHED_REFERENCES) {
      const oldest = this.refs.keys().next().value;
      if (oldest !== undefined) {
        this.refs.delete(oldest);
        this.lastWritten.delete(oldest);
      }
    }
  }

  get(conversationId: string): { ref: Partial<ConversationReference>; teamsType?: string } | undefined {
    return this.refs.get(conversationId);
  }

  /** Cache-or-store lookup for the proactive adapters: a miss after a restart
   *  falls back to the persisted reference and re-seeds the cache. */
  async getOrLoad(conversationId: string): Promise<{ ref: Partial<ConversationReference>; teamsType?: string } | undefined> {
    const cached = this.refs.get(conversationId);
    if (cached) return cached;
    if (!this.persistence) return undefined;
    const persisted = await this.persistence.load(conversationId);
    if (!persisted) return undefined;
    this.refs.delete(conversationId);
    this.refs.set(conversationId, persisted);
    if (this.refs.size > MAX_CACHED_REFERENCES) {
      const oldest = this.refs.keys().next().value;
      if (oldest !== undefined) {
        this.refs.delete(oldest);
        this.lastWritten.delete(oldest);
      }
    }
    return persisted;
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
    async sendToConversation(conversationId, message) {
      const cached = await deps.refs.getOrLoad(conversationId);
      if (!cached) {
        return {
          outcome: 'unreachable',
          code: 'no_binding',
          message: `no Teams conversation reference cached for '${conversationId}' - no inbound activity seen yet`,
        };
      }
      try {
        await deps.sendProactive(cached.ref, async (turnContext) => {
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

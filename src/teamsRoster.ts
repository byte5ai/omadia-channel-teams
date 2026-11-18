import { TeamsInfo, type TurnContext } from 'botbuilder';
import type { ChatParticipant } from './kernel-types.js';

interface CacheEntry {
  participants: ChatParticipant[];
  fetchedAt: number;
}

const DEFAULT_TTL_MS = 5 * 60 * 1000;
const PAGE_SIZE = 100;

/** How one roster read ended. `refused` = the Bot Connector did not answer
 *  with a member list (missing RSC permission, bot not in the conversation,
 *  transport error); the participants are then empty and NOT an answer. */
export type TeamsRosterOutcome =
  | { outcome: 'fetched' | 'cached'; participants: ChatParticipant[] }
  | { outcome: 'refused'; participants: ChatParticipant[]; status?: number; detail: string };

function statusOf(err: unknown): number | undefined {
  const record = (err ?? {}) as { statusCode?: unknown; status?: unknown };
  const status = typeof record.statusCode === 'number' ? record.statusCode : record.status;
  return typeof status === 'number' ? status : undefined;
}

/**
 * Teams roster cache. One instance is shared across every bot of the plugin
 * and hands out cached rosters keyed by the ASKING BOT and `conversation.id`.
 * The bot is part of the key because a roster is one bot's view: a bot that
 * is not in a chat gets no members, and that empty answer must never be
 * served to another bot that is in it (#2043).
 *
 * Uses `TeamsInfo.getPagedMembers` which works uniformly for 1:1, group-chat
 * and channel conversations. The call needs the appropriate RSC permission
 * (`ChatMember.Read.Chat` for group chats, `ChannelMember.Read.Group` for
 * channels) — missing permission surfaces as a 403 at runtime and is
 * swallowed into an empty roster so the turn continues.
 */
export class TeamsRosterProvider {
  private readonly cache = new Map<string, CacheEntry>();
  private readonly ttlMs: number;

  constructor(options: { ttlMs?: number } = {}) {
    this.ttlMs = options.ttlMs ?? DEFAULT_TTL_MS;
  }

  /**
   * `maxAgeMs` tightens the cache for one call. The kernel's member-scoped
   * memory reads the roster as the room's audience, and a roster cached across
   * a member joining would hide the newcomer from the rule and let them read
   * what was never theirs. Membership events invalidate the entry too; this
   * bounds the window when one is missed.
   */
  async list(
    context: TurnContext,
    options: { maxAgeMs?: number } = {},
  ): Promise<ChatParticipant[]> {
    return (await this.listWithOutcome(context, options)).participants;
  }

  /** {@link list}, saying how the read ended. */
  async listWithOutcome(
    context: TurnContext,
    options: { maxAgeMs?: number } = {},
  ): Promise<TeamsRosterOutcome> {
    const convId = context.activity.conversation?.id ?? 'unknown';
    // The bot this turn runs as: the recipient of an inbound activity, the
    // bot of the conversation reference in a proactive turn.
    const botId = (context.activity.recipient?.id ?? '').toLowerCase();
    const cacheKey = TeamsRosterProvider.key(botId, convId);
    const hit = this.cache.get(cacheKey);
    const maxAge = Math.min(this.ttlMs, options.maxAgeMs ?? this.ttlMs);
    if (hit && Date.now() - hit.fetchedAt < maxAge) {
      return { outcome: 'cached', participants: hit.participants };
    }
    const members: ChatParticipant[] = [];
    let continuation: string | undefined;
    try {
      do {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const page: any = await TeamsInfo.getPagedMembers(
          context,
          PAGE_SIZE,
          continuation,
        );
        const raw: unknown[] = Array.isArray(page?.members) ? page.members : [];
        for (const entry of raw) {
          const m = entry as Record<string, unknown>;
          const channelUserId = typeof m['id'] === 'string' ? (m['id'] as string) : '';
          if (channelUserId.length === 0) continue;
          const name = typeof m['name'] === 'string' ? (m['name'] as string) : '';
          members.push({
            channelUserId,
            aadObjectId:
              typeof m['aadObjectId'] === 'string'
                ? (m['aadObjectId'] as string)
                : null,
            displayName: name.length > 0 ? name : '(unbekannt)',
            email: typeof m['email'] === 'string' ? (m['email'] as string) : null,
            userPrincipalName:
              typeof m['userPrincipalName'] === 'string'
                ? (m['userPrincipalName'] as string)
                : null,
            // Bot Framework bot ids start with `28:`; a bot is not a person in
            // the room, for mentions or for member-scoped memory.
            ...(channelUserId.startsWith('28:') ? { kind: 'agent' as const } : {}),
          });
        }
        continuation =
          typeof page?.continuationToken === 'string'
            ? (page.continuationToken as string)
            : undefined;
      } while (continuation);
    } catch (err) {
      console.error(
        `[teams-roster] getPagedMembers failed conv=${convId}:`,
        err instanceof Error ? err.message : err,
      );
      // Soft-fail: cache empty with TTL so a missing RSC permission / transient
      // error doesn't hammer the Graph on every mention attempt.
      this.cache.set(cacheKey, { participants: [], fetchedAt: Date.now() });
      const status = statusOf(err);
      return {
        outcome: 'refused',
        participants: [],
        ...(status !== undefined ? { status } : {}),
        detail: err instanceof Error ? err.message : String(err),
      };
    }
    console.error(
      `[teams-roster] fetched conv=${convId} count=${String(members.length)}`,
    );
    this.cache.set(cacheKey, { participants: members, fetchedAt: Date.now() });
    return { outcome: 'fetched', participants: members };
  }

  private static key(botId: string, conversationId: string): string {
    return `${botId}|${conversationId}`;
  }

  /**
   * Drop the cached roster for a conversation — call on `membersAdded` /
   * `membersRemoved` events so a stale roster doesn't persist across changes.
   * Optional; TTL bounds staleness either way.
   */
  invalidate(conversationId: string): void {
    // Every bot's view of the conversation: a membership change concerns all.
    const suffix = `|${conversationId}`;
    for (const key of [...this.cache.keys()]) {
      if (key.endsWith(suffix)) this.cache.delete(key);
    }
  }
}

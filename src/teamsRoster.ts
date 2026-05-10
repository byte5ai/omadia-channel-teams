import { TeamsInfo, type TurnContext } from 'botbuilder';
import type { ChatParticipant } from './kernel-types.js';

interface CacheEntry {
  participants: ChatParticipant[];
  fetchedAt: number;
}

const DEFAULT_TTL_MS = 5 * 60 * 1000;
const PAGE_SIZE = 100;

/**
 * Per-conversation Teams roster cache. One instance is shared across the bot
 * and hands out cached rosters keyed by `conversation.id`.
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

  async list(context: TurnContext): Promise<ChatParticipant[]> {
    const convId = context.activity.conversation?.id ?? 'unknown';
    const hit = this.cache.get(convId);
    if (hit && Date.now() - hit.fetchedAt < this.ttlMs) {
      return hit.participants;
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
      this.cache.set(convId, { participants: [], fetchedAt: Date.now() });
      return [];
    }
    console.error(
      `[teams-roster] fetched conv=${convId} count=${String(members.length)}`,
    );
    this.cache.set(convId, { participants: members, fetchedAt: Date.now() });
    return members;
  }

  /**
   * Drop the cached roster for a conversation — call on `membersAdded` /
   * `membersRemoved` events so a stale roster doesn't persist across changes.
   * Optional; TTL bounds staleness either way.
   */
  invalidate(conversationId: string): void {
    this.cache.delete(conversationId);
  }
}

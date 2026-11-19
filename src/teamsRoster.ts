import { TeamsInfo, type TurnContext } from 'botbuilder';
import type { ChatParticipant } from './kernel-types.js';

/** How complete a member list is (kernel spec B §5.5, byte5ai/omadia#2091). */
export type TeamsRosterStatus = 'complete' | 'partial' | 'unknown';

/**
 * One read of a conversation's members, as ONE bot sees them.
 *
 *   complete  every page was read without an error and somebody is listed
 *   partial   some pages were read, then the read failed; `participants`
 *             holds what was read and nothing may be concluded from an
 *             absence
 *   unknown   nothing could be read (the first page failed), or the read came
 *             back empty; `participants` is empty and is NOT an answer
 *
 * An unreadable roster is never an empty list presented as the roster: the
 * first live facilitation listed no participants for exactly that reason.
 */
export interface TeamsRosterRead {
  status: TeamsRosterStatus;
  participants: ChatParticipant[];
  /** When the platform answered this list (epoch ms). A cached answer keeps
   *  the time of its original read. */
  readAt: number;
  /** `true` when this answer came from the cache. */
  fromCache: boolean;
  /** HTTP status of the failed call, for `partial` / `unknown`. */
  httpStatus?: number;
  /** Operator-readable reason, for `partial` / `unknown`. */
  detail?: string;
}

export interface TeamsRosterReadOptions {
  /** Tightens the cache for this call. */
  maxAgeMs?: number;
  /** Bypass the cache: ask the platform now. */
  fresh?: boolean;
}

type CacheEntry = Omit<TeamsRosterRead, 'fromCache'>;

/** How long a `complete` member list is served from the cache. */
const DEFAULT_TTL_MS = 5 * 60 * 1000;
/** How long a `partial` or `unknown` read is kept: long enough not to hammer
 *  the Bot Connector on every mention, short enough not to pin a failure. */
const INCOMPLETE_TTL_MS = 30 * 1000;
const PAGE_SIZE = 100;

function statusOf(err: unknown): number | undefined {
  const record = (err ?? {}) as { statusCode?: unknown; status?: unknown };
  const status = typeof record.statusCode === 'number' ? record.statusCode : record.status;
  return typeof status === 'number' ? status : undefined;
}

/** One member as the Bot Connector returns it (`TeamsChannelAccount`).
 *  `email` and `userPrincipalName` are carried whenever the platform sends
 *  them; nothing is derived or guessed when it does not. */
function toChatParticipant(entry: unknown): ChatParticipant | undefined {
  const m = (entry ?? {}) as Record<string, unknown>;
  const text = (key: string): string | null => {
    const value = m[key];
    return typeof value === 'string' && value.trim().length > 0 ? value : null;
  };
  const channelUserId = text('id');
  if (!channelUserId) return undefined;
  return {
    channelUserId,
    aadObjectId: text('aadObjectId') ?? text('objectId'),
    displayName: text('name') ?? '(unbekannt)',
    email: text('email'),
    userPrincipalName: text('userPrincipalName'),
    // Bot Framework bot ids start with `28:`; a bot is not a person in the
    // room, for mentions or for member-scoped memory.
    ...(channelUserId.startsWith('28:') ? { kind: 'agent' as const } : {}),
  };
}

/**
 * Teams roster cache. One instance is shared across every bot of the plugin
 * and hands out cached rosters keyed by the ASKING BOT and `conversation.id`.
 * The bot is part of the key because a roster is one bot's view: a bot that
 * is not in a chat gets no members, and that answer must never be served to
 * another bot that is in it (#2043).
 *
 * Uses `TeamsInfo.getPagedMembers` (Bot Connector), which works for 1:1,
 * group-chat and channel conversations.
 */
export class TeamsRosterProvider {
  private readonly cache = new Map<string, CacheEntry>();
  private readonly ttlMs: number;
  private readonly incompleteTtlMs: number;
  private readonly now: () => number;

  constructor(options: { ttlMs?: number; incompleteTtlMs?: number; now?: () => number } = {}) {
    this.ttlMs = options.ttlMs ?? DEFAULT_TTL_MS;
    this.incompleteTtlMs = options.incompleteTtlMs ?? INCOMPLETE_TTL_MS;
    this.now = options.now ?? Date.now;
  }

  /**
   * The members for callers that need a list or nothing: mention resolution
   * and the room's audience for member-scoped memory. Returns the members of
   * a `complete` read and an EMPTY list otherwise, as it always did; these
   * callers read "empty" as "unknown room" and must never work on a subset.
   * Callers that have to tell the three cases apart use {@link read}.
   *
   * `maxAgeMs` tightens the cache for one call. The kernel's member-scoped
   * memory reads the roster as the room's audience, and a roster cached across
   * a member joining would hide the newcomer from the rule and let them read
   * what was never theirs. Membership events invalidate the entry too; this
   * bounds the window when one is missed.
   */
  async list(context: TurnContext, options: TeamsRosterReadOptions = {}): Promise<ChatParticipant[]> {
    const read = await this.read(context, options);
    return read.status === 'complete' ? read.participants : [];
  }

  /** Read the members, saying how complete the answer is and when it was read. */
  async read(context: TurnContext, options: TeamsRosterReadOptions = {}): Promise<TeamsRosterRead> {
    const convId = context.activity.conversation?.id ?? 'unknown';
    // The bot this turn runs as: the recipient of an inbound activity, the
    // bot of the conversation reference in a proactive turn.
    const botId = (context.activity.recipient?.id ?? '').toLowerCase();
    const cacheKey = TeamsRosterProvider.key(botId, convId);
    const hit = this.cache.get(cacheKey);
    if (hit && options.fresh !== true) {
      const ttl = hit.status === 'complete' ? this.ttlMs : this.incompleteTtlMs;
      const maxAge = Math.min(ttl, options.maxAgeMs ?? ttl);
      if (this.now() - hit.readAt < maxAge) return { ...hit, fromCache: true };
    }

    const members: ChatParticipant[] = [];
    let failure: { httpStatus?: number; detail: string } | undefined;
    let continuation: string | undefined;
    try {
      do {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const page: any = await TeamsInfo.getPagedMembers(context, PAGE_SIZE, continuation);
        const raw: unknown[] = Array.isArray(page?.members) ? page.members : [];
        for (const entry of raw) {
          const member = toChatParticipant(entry);
          if (member) members.push(member);
        }
        continuation = typeof page?.continuationToken === 'string' ? (page.continuationToken as string) : undefined;
      } while (continuation);
    } catch (err) {
      const httpStatus = statusOf(err);
      failure = {
        ...(httpStatus !== undefined ? { httpStatus } : {}),
        detail: err instanceof Error ? err.message : String(err),
      };
    }

    const readAt = this.now();
    let entry: CacheEntry;
    if (failure) {
      // What was read before the failure is kept and marked: a subset is
      // `partial`, nothing at all is `unknown`.
      entry = { status: members.length > 0 ? 'partial' : 'unknown', participants: members, readAt, ...failure };
    } else if (members.length === 0) {
      entry = {
        status: 'unknown',
        participants: [],
        readAt,
        detail: 'the platform answered with no members; every conversation has at least one',
      };
    } else {
      entry = { status: 'complete', participants: members, readAt };
    }
    console.error(
      `[teams-roster] read conv=${convId} status=${entry.status} count=${String(entry.participants.length)}` +
        (entry.httpStatus !== undefined ? ` http=${String(entry.httpStatus)}` : '') +
        (entry.detail ? ` detail=${entry.detail}` : '') +
        (options.fresh === true ? ' fresh=true' : ''),
    );
    this.cache.set(cacheKey, entry);
    return { ...entry, fromCache: false };
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

/**
 * Microsoft-Graph name + member resolution for the channels directory.
 *
 * The Bot-Framework activity stream gives us real names only for Teams
 * *channels* (`channelData.channel/team`). Group chats surface as
 * `19:<guid>@thread.skype` with no topic, 1:1 chats as bare ids, and the
 * bot-level catch-all as `28:<app-id>`. This resolver fills those gaps
 * via Graph (client-credentials on the bot's own App Registration):
 *
 *   - group chat  → GET /chats/{id}            (topic)
 *                  + GET /chats/{id}/members    (display names)
 *   - channel     → GET /teams/{aadGroupId}/members (team roster —
 *                    a deliberate v1 approximation: private/shared
 *                    channels would need /channels/{id}/members)
 *   - catch-all   → GET /organization           (tenant display name)
 *
 * Required application permissions (admin consent): `Chat.Read.All` +
 * `ChatMember.Read.All` (group chats), `TeamMember.Read.All` (channels),
 * `Organization.Read.All` (catch-all label). Every call degrades
 * gracefully: a 401/403/404 caches a negative result for the TTL and
 * logs once per status class — the directory then keeps the
 * Bot-Framework-derived labels exactly as before this feature.
 *
 * Resolution is NEVER awaited on the dashboard's render path:
 * `listKeys()` must stay synchronous-fast (see @omadia/channel-sdk
 * contract), so callers `prime()` fire-and-forget (single-flight per
 * conversation) and read the cache with `get()`.
 */

export interface ResolvedConversation {
  /** Graph-resolved label override (e.g. group-chat topic). Absent when
   *  Graph added nothing over the Bot-Framework label. */
  readonly label?: string;
  /** Capped member display names (cap: `memberCap`, default 8). */
  readonly members?: readonly string[];
  /** Total members seen (uncapped, within the first Graph page). */
  readonly memberCount?: number;
}

export interface ConversationToResolve {
  readonly conversationId: string;
  readonly conversationType?: string;
  /** AAD group id of the owning team — Teams-channel conversations only. */
  readonly teamAadGroupId?: string;
}

export interface TeamsGraphResolverOptions {
  readonly tenantId: string;
  readonly clientId: string;
  readonly clientSecret: string;
  readonly log: (msg: string) => void;
  readonly fetchImpl?: typeof fetch;
  /** Cache TTL per conversation (default 10 min). */
  readonly ttlMs?: number;
  /** Per-request timeout (default 3 s). */
  readonly timeoutMs?: number;
  /** Max member names carried per conversation (default 8). */
  readonly memberCap?: number;
  /** Clock injection for tests. */
  readonly now?: () => number;
}

const DEFAULT_TTL_MS = 10 * 60 * 1000;
const DEFAULT_TIMEOUT_MS = 3_000;
const DEFAULT_MEMBER_CAP = 8;
const TOKEN_EXPIRY_SLACK_MS = 60_000;
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const MEMBER_PAGE_SIZE = 50;

interface CacheEntry {
  readonly value: ResolvedConversation;
  readonly fetchedAt: number;
}

export class TeamsGraphResolver {
  private readonly tenantId: string;
  private readonly clientId: string;
  private readonly clientSecret: string;
  private readonly log: (msg: string) => void;
  private readonly fetchImpl: typeof fetch;
  private readonly ttlMs: number;
  private readonly timeoutMs: number;
  private readonly memberCap: number;
  private readonly now: () => number;

  private tokenState:
    | { promise: Promise<string>; expiresAt: number }
    | undefined;
  private readonly cache = new Map<string, CacheEntry>();
  private readonly inflight = new Set<string>();
  private orgEntry: CacheEntry | undefined;
  private orgInflight = false;
  private readonly loggedFailures = new Set<string>();

  constructor(opts: TeamsGraphResolverOptions) {
    this.tenantId = opts.tenantId;
    this.clientId = opts.clientId;
    this.clientSecret = opts.clientSecret;
    this.log = opts.log;
    this.fetchImpl = opts.fetchImpl ?? fetch;
    this.ttlMs = opts.ttlMs ?? DEFAULT_TTL_MS;
    this.timeoutMs = opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;
    this.memberCap = opts.memberCap ?? DEFAULT_MEMBER_CAP;
    this.now = opts.now ?? Date.now;
  }

  /** Cache read — synchronous, never triggers network. Returns undefined
   *  both for "never resolved" and for a cached negative result (the
   *  caller keeps its Bot-Framework label either way); `prime()` reads
   *  the raw cache, so negative entries still suppress refetching. */
  get(conversationId: string): ResolvedConversation | undefined {
    const value = this.fresh(this.cache.get(conversationId))?.value;
    if (!value) return undefined;
    if (value.label === undefined && value.members === undefined) {
      return undefined;
    }
    return value;
  }

  /** Cached tenant organization display name (catch-all label). */
  getOrgName(): string | undefined {
    return this.fresh(this.orgEntry)?.value.label;
  }

  /**
   * Fire-and-forget resolution. Single-flight per conversation; a fresh
   * cache entry (positive OR negative) suppresses re-fetch until the TTL
   * lapses. Errors never propagate — they only produce negative cache
   * entries.
   */
  prime(conv: ConversationToResolve): void {
    const id = conv.conversationId;
    if (this.fresh(this.cache.get(id)) || this.inflight.has(id)) return;
    if (conv.conversationType !== 'groupChat' && !conv.teamAadGroupId) {
      return; // personal / unknown — nothing Graph can add in v1
    }
    if (conv.conversationType === 'groupChat' && !isGraphChatId(id)) {
      // Bot-Framework group chats (`19:…@thread.skype`) are not Graph
      // chat threads — `/chats/{id}` answers 400 for them. Their members
      // come from the Bot-Framework roster via the observer instead.
      return;
    }
    this.inflight.add(id);
    void this.resolve(conv)
      .then((value) => {
        this.cache.set(id, { value, fetchedAt: this.now() });
      })
      .catch((err: unknown) => {
        this.cache.set(id, { value: {}, fetchedAt: this.now() });
        this.warnOnce('resolve', err);
      })
      .finally(() => {
        this.inflight.delete(id);
      });
  }

  /** Fire-and-forget tenant-name resolution for the catch-all entry. */
  primeOrg(): void {
    if (this.fresh(this.orgEntry) || this.orgInflight) return;
    this.orgInflight = true;
    void this.fetchJson(`${GRAPH_BASE}/organization?$select=displayName`)
      .then((json) => {
        const name = firstOrgDisplayName(json);
        this.orgEntry = {
          value: name ? { label: name } : {},
          fetchedAt: this.now(),
        };
      })
      .catch((err: unknown) => {
        this.orgEntry = { value: {}, fetchedAt: this.now() };
        this.warnOnce('organization', err);
      })
      .finally(() => {
        this.orgInflight = false;
      });
  }

  private fresh(entry: CacheEntry | undefined): CacheEntry | undefined {
    if (!entry) return undefined;
    return this.now() - entry.fetchedAt < this.ttlMs ? entry : undefined;
  }

  private async resolve(
    conv: ConversationToResolve,
  ): Promise<ResolvedConversation> {
    if (conv.conversationType === 'groupChat') {
      return this.resolveGroupChat(conv.conversationId);
    }
    if (conv.teamAadGroupId) {
      return this.resolveTeamMembers(conv.teamAadGroupId);
    }
    return {};
  }

  private async resolveGroupChat(
    chatId: string,
  ): Promise<ResolvedConversation> {
    const id = encodeURIComponent(chatId);
    const [chat, members] = await Promise.all([
      this.fetchJson(`${GRAPH_BASE}/chats/${id}?$select=topic`),
      this.fetchJson(
        `${GRAPH_BASE}/chats/${id}/members?$select=displayName&$top=${MEMBER_PAGE_SIZE}`,
      ),
    ]);
    const topic = typeof chat['topic'] === 'string' ? chat['topic'].trim() : '';
    const names = memberDisplayNames(members);
    return {
      ...(topic ? { label: `Teams · ${topic}` } : {}),
      ...(names.length > 0
        ? {
            members: names.slice(0, this.memberCap),
            memberCount: names.length,
          }
        : {}),
    };
  }

  private async resolveTeamMembers(
    aadGroupId: string,
  ): Promise<ResolvedConversation> {
    const json = await this.fetchJson(
      `${GRAPH_BASE}/teams/${encodeURIComponent(aadGroupId)}/members?$select=displayName&$top=${MEMBER_PAGE_SIZE}`,
    );
    const names = memberDisplayNames(json);
    if (names.length === 0) return {};
    return {
      members: names.slice(0, this.memberCap),
      memberCount: names.length,
    };
  }

  private async fetchJson(url: string): Promise<Record<string, unknown>> {
    const token = await this.accessToken();
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const res = await this.fetchImpl(url, {
        headers: { Authorization: `Bearer ${token}` },
        signal: controller.signal,
      });
      if (!res.ok) {
        throw new Error(`graph ${String(res.status)} for ${url.slice(0, 80)}`);
      }
      return (await res.json()) as Record<string, unknown>;
    } finally {
      clearTimeout(timer);
    }
  }

  private accessToken(): Promise<string> {
    if (this.tokenState && this.now() < this.tokenState.expiresAt) {
      return this.tokenState.promise;
    }
    const requestedAt = this.now();
    const promise = this.requestToken().then(({ token, expiresInSec }) => {
      // Patch the real expiry in once known (state object is replaced,
      // not mutated — the placeholder below only guards the in-flight
      // window against a token stampede).
      this.tokenState = {
        promise: Promise.resolve(token),
        expiresAt: requestedAt + expiresInSec * 1000 - TOKEN_EXPIRY_SLACK_MS,
      };
      return token;
    });
    promise.catch(() => {
      this.tokenState = undefined;
    });
    this.tokenState = {
      promise,
      expiresAt: requestedAt + TOKEN_EXPIRY_SLACK_MS,
    };
    return promise;
  }

  private async requestToken(): Promise<{
    token: string;
    expiresInSec: number;
  }> {
    const body = new URLSearchParams({
      client_id: this.clientId,
      client_secret: this.clientSecret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    });
    const res = await this.fetchImpl(
      `https://login.microsoftonline.com/${encodeURIComponent(this.tenantId)}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'content-type': 'application/x-www-form-urlencoded' },
        body: body.toString(),
      },
    );
    if (!res.ok) {
      throw new Error(`graph token ${String(res.status)}`);
    }
    const json = (await res.json()) as {
      access_token?: string;
      expires_in?: number;
    };
    if (!json.access_token) {
      throw new Error('graph token response missing access_token');
    }
    return {
      token: json.access_token,
      expiresInSec:
        typeof json.expires_in === 'number' ? json.expires_in : 300,
    };
  }

  /** One warn line per failure class — a tenant without the Graph
   *  permissions would otherwise log every TTL lapse forever. */
  private warnOnce(what: string, err: unknown): void {
    const detail = err instanceof Error ? err.message : String(err);
    const cls = `${what}:${detail.replace(/for .*$/, '')}`;
    if (this.loggedFailures.has(cls)) return;
    this.loggedFailures.add(cls);
    this.log(
      `[teams] graph ${what} resolution unavailable (${detail}) — directory keeps Bot-Framework labels; check Graph application permissions`,
    );
  }
}

/** Graph's /chats/{id} only accepts Teams-native thread ids. Classic
 *  Bot-Framework group chats end in `@thread.skype` and are rejected
 *  with 400 — they must never reach Graph. */
function isGraphChatId(conversationId: string): boolean {
  return (
    conversationId.endsWith('@thread.v2') ||
    conversationId.endsWith('@unq.gbl.spaces')
  );
}

function memberDisplayNames(json: Record<string, unknown>): string[] {
  const value = json['value'];
  if (!Array.isArray(value)) return [];
  const names: string[] = [];
  for (const m of value) {
    const name = (m as { displayName?: unknown }).displayName;
    if (typeof name === 'string' && name.trim()) names.push(name.trim());
  }
  return names;
}

function firstOrgDisplayName(
  json: Record<string, unknown>,
): string | undefined {
  const value = json['value'];
  if (!Array.isArray(value) || value.length === 0) return undefined;
  const name = (value[0] as { displayName?: unknown }).displayName;
  return typeof name === 'string' && name.trim() ? name.trim() : undefined;
}

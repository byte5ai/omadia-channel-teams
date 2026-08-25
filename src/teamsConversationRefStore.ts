// #330 field report — the per-conversation ConversationReference cache is the
// ONLY way to open a proactive turn (roster reads, group nudges via
// conversationSend), and it was in-memory: every middleware restart made
// proactive delivery answer `no_binding` until the conversation happened to
// produce a new inbound activity. This store write-throughs the cache into
// `teams_conversation_refs` (kernel graph migration 0009) so references
// survive restarts. Best-effort by design: no pool, a missing table, or a
// failing query degrade to the old cache-only behaviour — never a throw on
// the message path.
//
// #860 W0a — multi-bot: one row PER BOT per conversation. The kernel owns the
// schema; migration 0010 (ships in the monorepo, NOT here) applies:
//
//   ALTER TABLE teams_conversation_refs
//     ADD COLUMN bot_app_id TEXT NOT NULL DEFAULT '';
//   ALTER TABLE teams_conversation_refs
//     DROP CONSTRAINT teams_conversation_refs_pkey;
//   ALTER TABLE teams_conversation_refs
//     ADD PRIMARY KEY (conversation_id, bot_app_id);
//
// Contract stated here once, relied on everywhere:
//   * `bot_app_id` is the lower-cased Entra app id of the bot that captured
//     the reference. `''` (the migration's column default) marks LEGACY rows
//     written before the migration — they belong to the DEFAULT bot.
//   * The plugin-side backfill (`backfillLegacyBotAppId`, run once at
//     activation by the config-wiring layer) rewrites `''` rows to the real
//     default-bot app id; re-running is a no-op. Until it ran, reads for the
//     default bot also accept `''` rows, so nothing breaks in between.
//   * Pre-migration kernels (no `bot_app_id` column) degrade to the legacy
//     single-key behaviour: every row is treated as the default bot's.
//     Cross-bot isolation is only guaranteed once migration 0010 is applied —
//     which holds for every kernel new enough to run more than one bot.

import type { ConversationReference } from 'botbuilder';
import type { Pool } from 'pg';

export interface TeamsConversationRefPersistence {
  /** `botAppId` undefined = the default bot (legacy single-bot call sites). */
  save(conversationId: string, ref: Partial<ConversationReference>, teamsType?: string, botAppId?: string): Promise<void>;
  /** `botAppId` undefined = the default bot (legacy single-bot call sites).
   *  The returned `botAppId` is the OWNING bot of the served row (lowercase,
   *  `undefined` for legacy/sentinel rows) — cache re-seeding must key on it
   *  so a restart-loaded entry lands under the same composite key a
   *  subsequent `capture()` of the same conversation would use. */
  load(conversationId: string, botAppId?: string): Promise<{ ref: Partial<ConversationReference>; teamsType?: string; botAppId?: string } | undefined>;
}

/** Sentinel `bot_app_id` of rows written before migration 0010 backfilled
 *  them — they belong to the default bot. */
export const LEGACY_BOT_APP_ID = '';

/**
 * Canonical form of a bot app id for keying (DB rows AND in-memory cache
 * keys): trimmed + lower-cased. Azure serialises app ids lowercase but
 * operator-pasted config may differ in casing — a casing drift must never
 * split one bot's references into two key spaces. Returns `undefined` for
 * missing/blank input ("the default bot").
 */
export function normalizeTeamsBotAppId(botAppId: string | undefined): string | undefined {
  const normalized = botAppId?.trim().toLowerCase();
  return normalized ? normalized : undefined;
}

/** Bot-Framework service hosts a persisted reference may point proactive
 *  turns at. Defense in depth (review M2): a poisoned DB row must not make
 *  `continueConversationAsync` present bot credentials to an attacker host —
 *  provenance on the WRITE path is JWT-authenticated inbound activities, this
 *  guards the read path. With multiple bots this risk multiplies per bot, so
 *  the allowlist stays on every read regardless of schema mode. */
function isAllowedServiceUrl(serviceUrl: unknown): boolean {
  if (typeof serviceUrl !== 'string') return false;
  try {
    const host = new URL(serviceUrl).hostname.toLowerCase();
    return host === 'smba.trafficmanager.net' || host.endsWith('.botframework.com');
  } catch {
    return false;
  }
}

/** Postgres error codes signalling the kernel has not applied migration 0010
 *  yet: 42703 undefined_column, 42P10 invalid ON CONFLICT target (column
 *  added but composite key not installed — never the case for the atomic
 *  kernel migration, guarded anyway). */
function isPreMigrationSchemaError(err: unknown): boolean {
  const code = (err as { code?: unknown } | null)?.code;
  return code === '42703' || code === '42P10';
}

type SchemaMode = 'unknown' | 'per-bot' | 'legacy';

interface PersistedRefRow {
  ref: Partial<ConversationReference>;
  teams_type: string | null;
  bot_app_id?: string | null;
}

export class PgTeamsConversationRefStore implements TeamsConversationRefPersistence {
  /** Determined lazily from the first schema error; a restart re-probes, so a
   *  kernel upgrade (which restarts the middleware anyway) is picked up. */
  private schemaMode: SchemaMode = 'unknown';
  private readonly defaultBotAppId: string | undefined;

  constructor(
    private readonly pool: Pool,
    private readonly log?: (msg: string) => void,
    opts?: { defaultBotAppId?: string },
  ) {
    this.defaultBotAppId = normalizeTeamsBotAppId(opts?.defaultBotAppId);
  }

  /** The key a write for `botAppId` lands under: explicit bot, else the
   *  configured default bot, else the legacy sentinel. */
  private effectiveBotKey(botAppId: string | undefined): string {
    return normalizeTeamsBotAppId(botAppId) ?? this.defaultBotAppId ?? LEGACY_BOT_APP_ID;
  }

  async save(conversationId: string, ref: Partial<ConversationReference>, teamsType?: string, botAppId?: string): Promise<void> {
    const refJson = JSON.stringify(ref);
    if (this.schemaMode !== 'legacy') {
      try {
        await this.pool.query(
          `INSERT INTO teams_conversation_refs (conversation_id, bot_app_id, ref, teams_type, updated_at)
           VALUES ($1, $2, $3, $4, now())
           ON CONFLICT (conversation_id, bot_app_id)
           DO UPDATE SET ref = EXCLUDED.ref, teams_type = EXCLUDED.teams_type, updated_at = now()`,
          [conversationId, this.effectiveBotKey(botAppId), refJson, teamsType ?? null],
        );
        this.schemaMode = 'per-bot';
        return;
      } catch (err) {
        if (!isPreMigrationSchemaError(err)) throw err;
        this.schemaMode = 'legacy';
        this.log?.('[teams] teams_conversation_refs has no bot_app_id column (kernel migration 0010 pending) — falling back to single-bot keying');
      }
    }
    // Pre-migration kernel: the single row per conversation IS the default
    // bot's row; a non-default bot cannot be isolated on this schema.
    await this.pool.query(
      `INSERT INTO teams_conversation_refs (conversation_id, ref, teams_type, updated_at)
       VALUES ($1, $2, $3, now())
       ON CONFLICT (conversation_id)
       DO UPDATE SET ref = EXCLUDED.ref, teams_type = EXCLUDED.teams_type, updated_at = now()`,
      [conversationId, refJson, teamsType ?? null],
    );
  }

  async load(conversationId: string, botAppId?: string): Promise<{ ref: Partial<ConversationReference>; teamsType?: string; botAppId?: string } | undefined> {
    try {
      const row = this.schemaMode === 'legacy'
        ? await this.loadLegacy(conversationId, botAppId)
        : await this.loadPerBot(conversationId, botAppId);
      if (!row) return undefined;
      if (!isAllowedServiceUrl(row.ref.serviceUrl)) {
        this.log?.(`[teams] persisted conversation ref for '${conversationId}' dropped — serviceUrl outside Bot Framework domains`);
        return undefined;
      }
      const owningBotAppId = normalizeTeamsBotAppId(row.bot_app_id ?? undefined);
      return {
        ref: row.ref,
        ...(row.teams_type ? { teamsType: row.teams_type } : {}),
        ...(owningBotAppId !== undefined ? { botAppId: owningBotAppId } : {}),
      };
    } catch (err) {
      // Pre-migration kernel or transient pool trouble: cache-only behaviour.
      this.log?.(`[teams] conversation-ref load failed: ${err instanceof Error ? err.message : String(err)}`);
      return undefined;
    }
  }

  private async loadPerBot(conversationId: string, botAppId: string | undefined): Promise<PersistedRefRow | undefined> {
    let rows: PersistedRefRow[];
    try {
      const r = await this.pool.query<PersistedRefRow>(
        `SELECT ref, teams_type, bot_app_id FROM teams_conversation_refs WHERE conversation_id = $1`,
        [conversationId],
      );
      this.schemaMode = 'per-bot';
      rows = r.rows;
    } catch (err) {
      if (!isPreMigrationSchemaError(err)) throw err;
      this.schemaMode = 'legacy';
      return this.loadLegacy(conversationId, botAppId);
    }
    const wanted = this.effectiveBotKey(botAppId);
    const exact = rows.find((row) => (row.bot_app_id ?? LEGACY_BOT_APP_ID) === wanted);
    if (exact) return exact;
    // Pre-backfill legacy rows (`''` / NULL) belong to the DEFAULT bot only —
    // a ref written by bot A must never be loaded for bot B.
    const isDefaultBot = wanted === LEGACY_BOT_APP_ID || wanted === this.defaultBotAppId;
    if (!isDefaultBot) return undefined;
    const legacyRow = rows.find((row) => (row.bot_app_id ?? LEGACY_BOT_APP_ID) === LEGACY_BOT_APP_ID);
    if (legacyRow) return legacyRow;
    // Single-bot deployment without a configured default (pre-wiring plugin):
    // capture() keys writes on the REAL app id derived from the activity, so a
    // default-bot read may only fall through to it when it is UNAMBIGUOUS —
    // exactly one bot holds a row for this conversation.
    return wanted === LEGACY_BOT_APP_ID && rows.length === 1 ? rows[0] : undefined;
  }

  /** "Missing column = default bot": on a pre-migration schema the single
   *  row is served to the default bot; an explicitly NON-default bot gets
   *  nothing rather than another bot's credentials context. */
  private async loadLegacy(conversationId: string, botAppId: string | undefined): Promise<PersistedRefRow | undefined> {
    const wanted = normalizeTeamsBotAppId(botAppId);
    if (wanted !== undefined && this.defaultBotAppId !== undefined && wanted !== this.defaultBotAppId) {
      return undefined;
    }
    const r = await this.pool.query<PersistedRefRow>(
      `SELECT ref, teams_type FROM teams_conversation_refs WHERE conversation_id = $1`,
      [conversationId],
    );
    return r.rows[0];
  }

  /** Schema mode as far as it is known — `'unknown'` until the first query
   *  (or `backfillLegacyBotAppId`) probed the kernel schema. The config
   *  wiring uses this to warn loudly when more than one bot is configured
   *  on a pre-migration kernel (no cross-bot row isolation available). */
  get knownSchemaMode(): SchemaMode {
    return this.schemaMode;
  }

  /**
   * One-shot, idempotent, NON-destructive backfill of pre-migration rows
   * onto the legacy single-bot app id — run at activation by the
   * config-wiring layer with the app id of the bot that actually captured
   * the legacy rows (the legacy scalar `microsoft_app_id`, falling back to
   * `teams_bots[0].appId`). A single guarded UPDATE, a no-op on re-run:
   * attribute `''` rows to that bot where it has no newer row. Nothing is
   * ever deleted (coordinator decision #860 W0a: no destructive backfill in
   * the plugin — the kernel migration owns the column default); a `''` row
   * whose conversation already has a fresher target-bot row simply stays,
   * and reads prefer the exact-match row over the sentinel.
   * Best-effort: a pre-migration kernel (no column) or pool trouble is
   * logged and swallowed — activation must not fail over this.
   */
  async backfillLegacyBotAppId(defaultBotAppId: string): Promise<void> {
    const target = normalizeTeamsBotAppId(defaultBotAppId);
    if (!target) return;
    try {
      await this.pool.query(
        `UPDATE teams_conversation_refs SET bot_app_id = $1
         WHERE bot_app_id = ''
           AND NOT EXISTS (
             SELECT 1 FROM teams_conversation_refs t2
             WHERE t2.conversation_id = teams_conversation_refs.conversation_id
               AND t2.bot_app_id = $1
           )`,
        [target],
      );
      this.schemaMode = 'per-bot';
    } catch (err) {
      if (isPreMigrationSchemaError(err)) {
        this.schemaMode = 'legacy';
        this.log?.('[teams] conversation-ref backfill skipped — kernel migration 0010 not applied yet');
        return;
      }
      this.log?.(`[teams] conversation-ref backfill failed: ${err instanceof Error ? err.message : String(err)}`);
    }
  }
}

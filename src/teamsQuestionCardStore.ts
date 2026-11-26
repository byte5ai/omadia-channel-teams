// Posted question cards, kept across restarts (#2089).
//
// WHY. The instance restarts on every deploy. A vote that spans a deploy
// would otherwise leave a card with live inputs that the adapter no longer
// knows: `updateCard` could not close it, and a click could only be read
// against the card the click itself carries, which a client can forge. With
// this table the adapter knows after a restart what it posted, where, and as
// which bot.
//
// THE TABLE is `teams_question_cards`, next to `teams_conversation_refs` on
// the same pool. Unlike that one (a kernel migration), this table is the
// plugin's own: `migrations/0001_question_cards.sql`, applied by the kernel
// under the granted `permissions.sql` (omadia#838). The same DDL also runs
// idempotently on first use, a no-op after the migration, for kernels that
// run no channel migrations.
//
// POOLER-SAFE. The pool is the Neon transaction-mode pooler on Fly: nothing
// here keeps session state. Every statement is a single autocommit query
// through `pool.query`; no transaction blocks, no advisory locks, no LISTEN,
// no named prepared statements, no temp tables.
//
// BEST EFFORT, AND SAID. No pool, missing DDL rights or a failing query
// degrade to the in-memory registry, logged once: cards then work until the
// next restart, exactly as before 0.38.2.

import type { Pool } from 'pg';

import type { QuestionCard } from './kernel-types.js';
import type { ClosedCardView } from './teamsQuestionCardRender.js';

export interface StoredQuestionCard {
  /** Lower-cased Entra app id of the bot that posted the card. */
  botAppId: string;
  card: QuestionCard;
  conversationId: string;
  messageId: string | undefined;
  /** Set once the engine closed the card. */
  closed?: ClosedCardView;
}

export interface TeamsQuestionCardPersistence {
  save(record: StoredQuestionCard): Promise<void>;
  loadByCard(botAppId: string, cardId: string): Promise<StoredQuestionCard | undefined>;
  loadByMessage(botAppId: string, messageId: string): Promise<StoredQuestionCard | undefined>;
  /** Cards still open whose closing time lies at or before `before`, oldest
   *  first. Optional: a persistence without it leaves the sweep to what the
   *  registry holds in memory. */
  listOverdueOpen?(before: Date, limit: number): Promise<readonly StoredQuestionCard[]>;
}

export const QUESTION_CARDS_TABLE = 'teams_question_cards';
const DAY_MS = 24 * 60 * 60 * 1000;
/** A card is kept this long after its closing time (or after it was closed):
 *  long enough for a late update or a late click to be answered correctly. */
export const KEEP_AFTER_CLOSE_MS = 14 * DAY_MS;
/** A card without a closing time is kept this long after it was last saved. */
export const KEEP_WITHOUT_CLOSE_MS = 60 * DAY_MS;
const PURGE_INTERVAL_MS = 60 * 60 * 1000;

/**
 * The table, as `migrations/0001_question_cards.sql` creates it. The kernel applies that
 * file before activate() when the operator granted `permissions.sql`
 * (omadia#838). The statements also run at first use, as a no-op after the
 * migration, so a kernel that runs no channel migrations keeps working.
 * `tests/sqlMigrations.test.ts` keeps both equal.
 */
export const QUESTION_CARDS_DDL: readonly string[] = [
  `CREATE TABLE IF NOT EXISTS ${QUESTION_CARDS_TABLE} (
     bot_app_id       TEXT        NOT NULL,
     card_id          TEXT        NOT NULL,
     conversation_id  TEXT        NOT NULL,
     message_id       TEXT,
     proposal_id      TEXT,
     proposal_version INTEGER,
     card             JSONB       NOT NULL,
     closes_at        TIMESTAMPTZ,
     state            TEXT        NOT NULL DEFAULT 'open',
     closed           JSONB,
     expires_at       TIMESTAMPTZ NOT NULL,
     updated_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
     PRIMARY KEY (bot_app_id, card_id)
   )`,
  `CREATE INDEX IF NOT EXISTS ${QUESTION_CARDS_TABLE}_message_idx ON ${QUESTION_CARDS_TABLE} (bot_app_id, message_id)`,
  `CREATE INDEX IF NOT EXISTS ${QUESTION_CARDS_TABLE}_expires_idx ON ${QUESTION_CARDS_TABLE} (expires_at)`,
];

/** When a stored card may be deleted. */
export function expiresAt(record: Pick<StoredQuestionCard, 'card' | 'closed'>, now: number): Date {
  const closeTime = record.closed ? Date.parse(record.closed.closedAt) : record.card.closesAt ? Date.parse(record.card.closesAt) : NaN;
  return Number.isNaN(closeTime)
    ? new Date(now + KEEP_WITHOUT_CLOSE_MS)
    : new Date(Math.max(closeTime, now) + KEEP_AFTER_CLOSE_MS);
}

interface Row {
  bot_app_id: string;
  card: QuestionCard;
  conversation_id: string;
  message_id: string | null;
  state: string;
  closed: ClosedCardView | null;
}

function toRecord(row: Row): StoredQuestionCard {
  return {
    botAppId: row.bot_app_id,
    card: row.card,
    conversationId: row.conversation_id,
    messageId: row.message_id ?? undefined,
    ...(row.state === 'closed' && row.closed ? { closed: row.closed } : {}),
  };
}

const COLUMNS = 'bot_app_id, card, conversation_id, message_id, state, closed';

export class PgTeamsQuestionCardStore implements TeamsQuestionCardPersistence {
  /** `undefined` = not tried yet; the promise is shared by concurrent first
   *  uses; `false` = the table cannot be used, memory only from here on. */
  private ready: Promise<boolean> | undefined;
  private lastPurge = 0;

  constructor(
    private readonly pool: Pick<Pool, 'query'>,
    private readonly log?: (message: string) => void,
    private readonly now: () => number = Date.now,
  ) {}

  private ensureTable(): Promise<boolean> {
    this.ready ??= (async () => {
      try {
        for (const statement of QUESTION_CARDS_DDL) await this.pool.query(statement);
        return true;
      } catch (err) {
        this.log?.(
          `[teams] question cards are kept in memory only: the table '${QUESTION_CARDS_TABLE}' could not be created or reached (${err instanceof Error ? err.message : String(err)}). After a restart of the Teams channel, earlier cards cannot be closed or answered.`,
        );
        return false;
      }
    })();
    return this.ready;
  }

  async save(record: StoredQuestionCard): Promise<void> {
    if (!(await this.ensureTable())) return;
    const now = this.now();
    await this.pool.query(
      `INSERT INTO ${QUESTION_CARDS_TABLE}
         (bot_app_id, card_id, conversation_id, message_id, proposal_id, proposal_version, card, closes_at, state, closed, expires_at, updated_at)
       VALUES ($1, $2, $3, $4, $5, $6, $7::jsonb, $8, $9, $10::jsonb, $11, now())
       ON CONFLICT (bot_app_id, card_id) DO UPDATE SET
         conversation_id = EXCLUDED.conversation_id,
         message_id = EXCLUDED.message_id,
         proposal_id = EXCLUDED.proposal_id,
         proposal_version = EXCLUDED.proposal_version,
         card = EXCLUDED.card,
         closes_at = EXCLUDED.closes_at,
         state = EXCLUDED.state,
         closed = EXCLUDED.closed,
         expires_at = EXCLUDED.expires_at,
         updated_at = now()`,
      [
        record.botAppId.toLowerCase(),
        record.card.cardId,
        record.conversationId,
        record.messageId ?? null,
        record.card.proposalId ?? null,
        record.card.proposalVersion ?? null,
        JSON.stringify(record.card),
        record.card.closesAt ?? null,
        record.closed ? 'closed' : 'open',
        record.closed ? JSON.stringify(record.closed) : null,
        expiresAt(record, now).toISOString(),
      ],
    );
    void this.purgeExpired(now);
  }

  async loadByCard(botAppId: string, cardId: string): Promise<StoredQuestionCard | undefined> {
    return this.loadOne(`bot_app_id = $1 AND card_id = $2`, [botAppId.toLowerCase(), cardId]);
  }

  async loadByMessage(botAppId: string, messageId: string): Promise<StoredQuestionCard | undefined> {
    return this.loadOne(`bot_app_id = $1 AND message_id = $2`, [botAppId.toLowerCase(), messageId]);
  }

  async listOverdueOpen(before: Date, limit: number): Promise<readonly StoredQuestionCard[]> {
    if (!(await this.ensureTable())) return [];
    try {
      const result = await this.pool.query<Row>(
        `SELECT ${COLUMNS} FROM ${QUESTION_CARDS_TABLE} WHERE state = 'open' AND closes_at <= $1 AND expires_at > now() ORDER BY closes_at LIMIT $2`,
        [before.toISOString(), limit],
      );
      return result.rows.map(toRecord);
    } catch (err) {
      this.log?.(`[teams] question card lookup failed: ${err instanceof Error ? err.message : String(err)}`);
      return [];
    }
  }

  private async loadOne(where: string, params: readonly string[]): Promise<StoredQuestionCard | undefined> {
    if (!(await this.ensureTable())) return undefined;
    try {
      const result = await this.pool.query<Row>(
        `SELECT ${COLUMNS} FROM ${QUESTION_CARDS_TABLE} WHERE ${where} AND expires_at > now() ORDER BY updated_at DESC LIMIT 1`,
        [...params],
      );
      const row = result.rows[0];
      return row ? toRecord(row) : undefined;
    } catch (err) {
      // A failed read is "unknown card", never a crash of a click or an update.
      this.log?.(`[teams] question card lookup failed: ${err instanceof Error ? err.message : String(err)}`);
      return undefined;
    }
  }

  /** Delete what is past its time, at most once an hour, never awaited by a
   *  caller and never thrown. */
  private async purgeExpired(now: number): Promise<void> {
    if (now - this.lastPurge < PURGE_INTERVAL_MS) return;
    this.lastPurge = now;
    try {
      await this.pool.query(`DELETE FROM ${QUESTION_CARDS_TABLE} WHERE expires_at <= now()`);
    } catch (err) {
      this.log?.(`[teams] question card purge failed: ${err instanceof Error ? err.message : String(err)}`);
    }
  }
}

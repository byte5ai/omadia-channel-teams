// What a bot has seen of the meeting in a conversation, kept across restarts
// (#2043, rollout defect D26).
//
// WHY. The capability snapshot says `online` only while a meeting is known
// to run, and "known" is what this bot saw: a start, a join, an end. That
// lived in memory, and the instance restarts on every deploy; a deploy during
// a meeting made a running meeting look like a plain chat. Microsoft
// documents no way to ask whether a meeting is running, so the only way to
// know after a restart is to have written it down.
//
// A STORED "RUNNING" GOES STALE. If the end event is missed (it arrives while
// the process is down), the row would say "running" for ever. So a stored
// running state counts only for a limited time after the last evidence
// (`RUNNING_EVIDENCE_MAX_AGE_MS`); older than that it is unknown, and unknown
// is reported as a chat.
//
// Same rules as `teamsQuestionCardStore.ts`: the plugin creates its table on
// first use, every statement is a single autocommit query (transaction-mode
// pooler), and every failure degrades to memory with one log line.

import type { Pool } from 'pg';

import type { TeamsMeetingState } from './teamsMeetingSessions.js';

export interface StoredMeetingState {
  state: TeamsMeetingState;
  /** Epoch ms of the last meeting event this state was built from. */
  seenAt: number;
}

export interface TeamsMeetingStatePersistence {
  save(botKey: string, conversationId: string, stored: StoredMeetingState): Promise<void>;
  load(botKey: string, conversationId: string): Promise<StoredMeetingState | undefined>;
}

export const MEETING_STATES_TABLE = 'teams_meeting_states';
/** A stored "running" is believed this long after the last start or join.
 *  Longer than any ordinary meeting; short enough that a missed end event
 *  does not leave a chat marked as a meeting until the next one. */
export const RUNNING_EVIDENCE_MAX_AGE_MS = 12 * 60 * 60 * 1000;
/** Rows are deleted this long after their last event. */
const KEEP_MS = 30 * 24 * 60 * 60 * 1000;
const PURGE_INTERVAL_MS = 60 * 60 * 1000;

/**
 * The table, as `migrations/0002_meeting_states.sql` creates it. The kernel applies that
 * file before activate() when the operator granted `permissions.sql`
 * (omadia#838). The statements also run at first use, as a no-op after the
 * migration, so a kernel that runs no channel migrations keeps working.
 * `tests/sqlMigrations.test.ts` keeps both equal.
 */
export const MEETING_STATES_DDL: readonly string[] = [
  `CREATE TABLE IF NOT EXISTS ${MEETING_STATES_TABLE} (
     bot_key         TEXT        NOT NULL,
     conversation_id TEXT        NOT NULL,
     state           JSONB       NOT NULL,
     seen_at         TIMESTAMPTZ NOT NULL,
     expires_at      TIMESTAMPTZ NOT NULL,
     PRIMARY KEY (bot_key, conversation_id)
   )`,
  `CREATE INDEX IF NOT EXISTS ${MEETING_STATES_TABLE}_expires_idx ON ${MEETING_STATES_TABLE} (expires_at)`,
];

/**
 * A state as it may be USED at `now`: a running meeting whose last evidence
 * is too old is no longer claimed to run.
 */
export function believableMeetingState(stored: StoredMeetingState, now: number): TeamsMeetingState {
  return stored.state.isLive && now - stored.seenAt > RUNNING_EVIDENCE_MAX_AGE_MS
    ? { ...stored.state, isLive: false, hasSeenParticipantEvent: false }
    : stored.state;
}

export class PgTeamsMeetingStateStore implements TeamsMeetingStatePersistence {
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
        for (const statement of MEETING_STATES_DDL) await this.pool.query(statement);
        return true;
      } catch (err) {
        this.log?.(
          `[teams] meeting state is kept in memory only: the table '${MEETING_STATES_TABLE}' could not be created or reached (${err instanceof Error ? err.message : String(err)}). After a restart of the Teams channel during a meeting, its chat is reported as a plain chat until somebody joins.`,
        );
        return false;
      }
    })();
    return this.ready;
  }

  async save(botKey: string, conversationId: string, stored: StoredMeetingState): Promise<void> {
    if (!(await this.ensureTable())) return;
    await this.pool.query(
      `INSERT INTO ${MEETING_STATES_TABLE} (bot_key, conversation_id, state, seen_at, expires_at)
       VALUES ($1, $2, $3::jsonb, $4, $5)
       ON CONFLICT (bot_key, conversation_id) DO UPDATE SET
         state = EXCLUDED.state, seen_at = EXCLUDED.seen_at, expires_at = EXCLUDED.expires_at`,
      [
        botKey.toLowerCase(),
        conversationId,
        JSON.stringify(stored.state),
        new Date(stored.seenAt).toISOString(),
        new Date(stored.seenAt + KEEP_MS).toISOString(),
      ],
    );
    const now = this.now();
    if (now - this.lastPurge >= PURGE_INTERVAL_MS) {
      this.lastPurge = now;
      void this.pool
        .query(`DELETE FROM ${MEETING_STATES_TABLE} WHERE expires_at <= now()`)
        .catch((err: unknown) => this.log?.(`[teams] meeting state purge failed: ${err instanceof Error ? err.message : String(err)}`));
    }
  }

  async load(botKey: string, conversationId: string): Promise<StoredMeetingState | undefined> {
    if (!(await this.ensureTable())) return undefined;
    try {
      const result = await this.pool.query<{ state: TeamsMeetingState; seen_at: Date | string }>(
        `SELECT state, seen_at FROM ${MEETING_STATES_TABLE} WHERE bot_key = $1 AND conversation_id = $2 AND expires_at > now()`,
        [botKey.toLowerCase(), conversationId],
      );
      const row = result.rows[0];
      if (!row) return undefined;
      const seenAt = new Date(row.seen_at).getTime();
      return Number.isNaN(seenAt) ? undefined : { state: row.state, seenAt };
    } catch (err) {
      this.log?.(`[teams] meeting state lookup failed: ${err instanceof Error ? err.message : String(err)}`);
      return undefined;
    }
  }
}

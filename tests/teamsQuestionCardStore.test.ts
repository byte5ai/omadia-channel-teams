import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  believableMeetingState,
  MEETING_STATES_DDL,
  PgTeamsMeetingStateStore,
  RUNNING_EVIDENCE_MAX_AGE_MS,
  KEEP_AFTER_CLOSE_MS,
  KEEP_WITHOUT_CLOSE_MS,
  PgTeamsQuestionCardStore,
  QUESTION_CARDS_DDL,
  QUESTION_CARDS_TABLE,
  questionCardExpiresAt,
  type StoredQuestionCard,
} from '@omadia/channel-teams';

import type { QuestionCard } from '../src/kernel-types.js';

/**
 * #2089 — posted question cards in `teams_question_cards`.
 *
 * The database is a fake pool that records every statement. What these tests
 * hold is what can be held without a database: WHICH statements are sent
 * (one at a time, nothing that needs a session), with which values, and what
 * happens when they fail. That the SQL runs on Postgres behind the Neon
 * pooler is NOT verified here.
 */

const NOW = Date.parse('2026-11-19T10:00:00.000Z');
const BOT = '6E7D5E84-5216-408F-B74B-E299BBD0D947';

const CARD: QuestionCard = {
  cardId: 'c-1',
  proposalId: 'p-4',
  proposalVersion: 2,
  title: 'Welchen Termin nehmen wir?',
  closesAt: '2026-11-20T16:00:00.000Z',
  questions: [{ id: 'q', label: 'Termin', kind: 'single_choice', options: [{ id: 'a', label: 'A' }, { id: 'b', label: 'B' }] }],
};

const RECORD: StoredQuestionCard = { botAppId: BOT, card: CARD, conversationId: '19:chat@thread.v2', messageId: '1795027571374' };

function fakePool(opts: { failOn?: RegExp; rows?: Record<string, unknown>[] } = {}) {
  const statements: { sql: string; params: unknown[] }[] = [];
  return {
    statements,
    pool: {
      query: async (sql: string, params: unknown[] = []) => {
        statements.push({ sql: sql.replace(/\s+/g, ' ').trim(), params });
        if (opts.failOn?.test(sql)) throw new Error('permission denied for schema public');
        return { rows: /^\s*SELECT/i.test(sql) ? (opts.rows ?? []) : [] };
      },
    } as never,
  };
}

describe('#2089 question card store', () => {
  it('creates its table on first use, once, and stores a card with one upsert', async () => {
    const db = fakePool();
    const store = new PgTeamsQuestionCardStore(db.pool, undefined, () => NOW);

    await store.save(RECORD);
    await store.save(RECORD);

    const ddl = db.statements.filter((s) => s.sql.startsWith('CREATE'));
    assert.equal(ddl.length, QUESTION_CARDS_DDL.length, 'the DDL ran once, not per save');
    assert.ok(ddl.every((s) => s.sql.includes('IF NOT EXISTS')), 'idempotent: a no-op when the table exists');
    const inserts = db.statements.filter((s) => s.sql.startsWith('INSERT'));
    assert.equal(inserts.length, 2);
    assert.match(inserts[0]?.sql ?? '', /ON CONFLICT \(bot_app_id, card_id\) DO UPDATE/);
    assert.deepEqual(inserts[0]?.params, [
      BOT.toLowerCase(),
      'c-1',
      '19:chat@thread.v2',
      '1795027571374',
      'p-4',
      2,
      JSON.stringify(CARD),
      '2026-11-20T16:00:00.000Z',
      'open',
      null,
      // 14 days after the closing time.
      '2026-12-04T16:00:00.000Z',
    ]);
  });

  it('sends nothing that needs a database session (transaction-mode pooler)', async () => {
    const db = fakePool({ rows: [] });
    const store = new PgTeamsQuestionCardStore(db.pool, undefined, () => NOW);

    await store.save(RECORD);
    await store.loadByCard(BOT, 'c-1');
    await store.loadByMessage(BOT, '1795027571374');
    await new Promise((resolve) => setImmediate(resolve));

    assert.ok(db.statements.length >= 6);
    for (const { sql } of db.statements) {
      assert.doesNotMatch(sql, /\b(BEGIN|COMMIT|ROLLBACK|LISTEN|NOTIFY|PREPARE|DEALLOCATE|SET\s+(SESSION|search_path)|pg_advisory|TEMP|DECLARE)\b/i, sql);
      assert.equal(sql.split(';').filter((part) => part.trim().length > 0).length, 1, 'one statement per query');
    }
  });

  it('reads a card back by card id and by message id, for that bot only, and not past its expiry', async () => {
    const row = { bot_app_id: BOT.toLowerCase(), card: CARD, conversation_id: '19:chat@thread.v2', message_id: '1795027571374', state: 'open', closed: null };
    const db = fakePool({ rows: [row] });
    const store = new PgTeamsQuestionCardStore(db.pool, undefined, () => NOW);

    const byCard = await store.loadByCard(BOT, 'c-1');
    const byMessage = await store.loadByMessage(BOT, '1795027571374');

    assert.deepEqual(byCard, { botAppId: BOT.toLowerCase(), card: CARD, conversationId: '19:chat@thread.v2', messageId: '1795027571374' });
    assert.deepEqual(byMessage, byCard);
    const selects = db.statements.filter((s) => s.sql.startsWith('SELECT'));
    assert.match(selects[0]?.sql ?? '', /WHERE bot_app_id = \$1 AND card_id = \$2 AND expires_at > now\(\)/);
    assert.match(selects[1]?.sql ?? '', /WHERE bot_app_id = \$1 AND message_id = \$2 AND expires_at > now\(\)/);
    assert.deepEqual(selects.map((s) => s.params), [[BOT.toLowerCase(), 'c-1'], [BOT.toLowerCase(), '1795027571374']]);
  });

  it('stores and returns the closed state', async () => {
    const closed = { title: 't', resultLines: ['Ergebnis: A'], closedAt: '2026-11-19T10:05:00.000Z' };
    const db = fakePool({
      rows: [{ bot_app_id: BOT.toLowerCase(), card: CARD, conversation_id: 'c', message_id: 'm', state: 'closed', closed }],
    });
    const store = new PgTeamsQuestionCardStore(db.pool, undefined, () => NOW);

    await store.save({ ...RECORD, closed });
    const read = await store.loadByCard(BOT, 'c-1');

    const insert = db.statements.find((s) => s.sql.startsWith('INSERT'));
    assert.equal(insert?.params[8], 'closed');
    assert.equal(insert?.params[9], JSON.stringify(closed));
    assert.equal(insert?.params[10], '2026-12-03T10:05:00.000Z', '14 days after it was closed');
    assert.deepEqual(read?.closed, closed);
  });

  it('keeps a card 14 days after its close and 60 days when it has no closing time', () => {
    const noClose = { ...CARD, closesAt: undefined } as QuestionCard;
    assert.equal(questionCardExpiresAt({ card: CARD }, NOW).getTime(), Date.parse('2026-11-20T16:00:00.000Z') + KEEP_AFTER_CLOSE_MS);
    assert.equal(questionCardExpiresAt({ card: noClose }, NOW).getTime(), NOW + KEEP_WITHOUT_CLOSE_MS);
    // A closing time already in the past still gives the full period from now.
    const past = { ...CARD, closesAt: '2026-01-01T00:00:00.000Z' };
    assert.equal(questionCardExpiresAt({ card: past }, NOW).getTime(), NOW + KEEP_AFTER_CLOSE_MS);
  });

  it('deletes expired cards at most once an hour', async () => {
    const db = fakePool();
    const clock = { now: NOW };
    const store = new PgTeamsQuestionCardStore(db.pool, undefined, () => clock.now);

    await store.save(RECORD);
    await store.save(RECORD);
    clock.now += 61 * 60 * 1000;
    await store.save(RECORD);
    await new Promise((resolve) => setImmediate(resolve));

    const deletes = db.statements.filter((s) => s.sql.startsWith('DELETE'));
    assert.equal(deletes.length, 2);
    assert.equal(deletes[0]?.sql, `DELETE FROM ${QUESTION_CARDS_TABLE} WHERE expires_at <= now()`);
  });

  it('without the right to create the table: memory only, said once, and never a throw', async () => {
    const db = fakePool({ failOn: /CREATE TABLE/ });
    const logged: string[] = [];
    const store = new PgTeamsQuestionCardStore(db.pool, (message) => logged.push(message), () => NOW);

    await store.save(RECORD);
    await store.save(RECORD);
    assert.equal(await store.loadByCard(BOT, 'c-1'), undefined);
    assert.equal(await store.loadByMessage(BOT, 'm'), undefined);

    assert.equal(logged.length, 1);
    assert.match(logged[0] ?? '', /kept in memory only: the table 'teams_question_cards' could not be created or reached \(permission denied/);
    assert.equal(db.statements.filter((s) => !s.sql.startsWith('CREATE')).length, 0, 'no insert or select is attempted afterwards');
    assert.equal(db.statements.length, 1, 'and the DDL is not retried on every call');
  });

  it('a failing read is "unknown card", not an exception in a click', async () => {
    const db = fakePool({ failOn: /SELECT/ });
    const logged: string[] = [];
    const store = new PgTeamsQuestionCardStore(db.pool, (message) => logged.push(message), () => NOW);

    assert.equal(await store.loadByCard(BOT, 'c-1'), undefined);
    assert.match(logged[0] ?? '', /question card lookup failed/);
  });
});

describe('D26 meeting state store', () => {
  const STATE = { sessionRef: 'meeting-1', meetingType: 'Scheduled', isScheduled: true, isLive: true, hasSeenParticipantEvent: true };
  const BOT_KEY = `28:${BOT}`;

  it('creates its table once and stores a state with one upsert, nothing that needs a session', async () => {
    const db = fakePool({ rows: [{ state: STATE, seen_at: '2026-11-19T10:00:00.000Z' }] });
    const store = new PgTeamsMeetingStateStore(db.pool, undefined, () => NOW);

    await store.save(BOT_KEY, '19:meeting_x@thread.v2', { state: STATE, seenAt: NOW });
    await store.save(BOT_KEY, '19:meeting_x@thread.v2', { state: { ...STATE, isLive: false }, seenAt: NOW + 1000 });
    const read = await store.load(BOT_KEY, '19:meeting_x@thread.v2');
    await new Promise((resolve) => setImmediate(resolve));

    assert.equal(db.statements.filter((s) => s.sql.startsWith('CREATE')).length, MEETING_STATES_DDL.length);
    const insert = db.statements.find((s) => s.sql.startsWith('INSERT'));
    assert.match(insert?.sql ?? '', /ON CONFLICT \(bot_key, conversation_id\) DO UPDATE/);
    assert.deepEqual(insert?.params.slice(0, 4), [BOT_KEY.toLowerCase(), '19:meeting_x@thread.v2', JSON.stringify(STATE), '2026-11-19T10:00:00.000Z']);
    assert.deepEqual(read, { state: STATE, seenAt: NOW });
    for (const { sql } of db.statements) {
      assert.doesNotMatch(sql, /\b(BEGIN|COMMIT|ROLLBACK|LISTEN|NOTIFY|PREPARE|DEALLOCATE|SET\s+(SESSION|search_path)|pg_advisory|TEMP|DECLARE)\b/i, sql);
      assert.equal(sql.split(';').filter((part) => part.trim().length > 0).length, 1);
    }
  });

  it('believes a stored running meeting for twelve hours after the last evidence, not longer', () => {
    const stored = { state: STATE, seenAt: NOW };
    assert.equal(believableMeetingState(stored, NOW + RUNNING_EVIDENCE_MAX_AGE_MS).isLive, true);
    const stale = believableMeetingState(stored, NOW + RUNNING_EVIDENCE_MAX_AGE_MS + 1);
    assert.equal(stale.isLive, false);
    assert.equal(stale.hasSeenParticipantEvent, false, 'presence has to be proven again too');
    assert.equal(stale.sessionRef, 'meeting-1');
    // An ended meeting stays ended, whatever its age.
    assert.equal(believableMeetingState({ state: { ...STATE, isLive: false }, seenAt: NOW }, NOW + 1).isLive, false);
  });

  it('without the right to create the table: memory only, said once, and never a throw', async () => {
    const db = fakePool({ failOn: /CREATE TABLE/ });
    const logged: string[] = [];
    const store = new PgTeamsMeetingStateStore(db.pool, (message) => logged.push(message), () => NOW);

    await store.save(BOT_KEY, 'c', { state: STATE, seenAt: NOW });
    assert.equal(await store.load(BOT_KEY, 'c'), undefined);

    assert.equal(logged.length, 1);
    assert.match(logged[0] ?? '', /meeting state is kept in memory only: the table 'teams_meeting_states' could not be created/);
    assert.equal(db.statements.length, 1);
  });
});

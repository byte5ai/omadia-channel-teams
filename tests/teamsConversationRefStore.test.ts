import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';
import type { ConversationReference } from 'botbuilder';
import type { Pool } from 'pg';

// #860 W0a — per-bot conversation-reference persistence. The store module is
// imported relatively (it only has TYPE imports, so esbuild bundles it clean
// from src/, and the `src/index.ts` export surface is owned by the
// config-wiring unit); the cache comes via the package entry like in
// teamsGroupPrimitives.test.ts — bundling `teamsGroupPrimitives.ts` would
// drag the CJS botbuilder runtime into the ESM test bundle.
import {
  LEGACY_BOT_APP_ID,
  normalizeTeamsBotAppId,
  PgTeamsConversationRefStore,
} from '../src/teamsConversationRefStore.js';
import type { TeamsConversationRefPersistence } from '../src/teamsConversationRefStore.js';
import { TeamsConversationReferenceCache } from '@omadia/channel-teams';

const ALLOWED_SERVICE_URL = 'https://smba.trafficmanager.net/emea/';

function pgError(code: string, message: string): Error {
  return Object.assign(new Error(message), { code });
}

interface FakeRow {
  conversation_id: string;
  bot_app_id: string;
  ref: Record<string, unknown>;
  teams_type: string | null;
}

/**
 * In-memory emulation of the `teams_conversation_refs` table for exactly the
 * statements the store issues — stateful, so idempotency and isolation are
 * tested against real row movement, not against SQL strings. `hasBotAppId`
 * false emulates a pre-migration kernel (every reference to the column
 * raises 42703, like Postgres does).
 */
class FakeRefTable {
  readonly rows = new Map<string, FakeRow>();
  readonly queries: { text: string; params: unknown[] }[] = [];
  failEverythingWith: Error | undefined;

  constructor(private readonly hasBotAppId: boolean) {}

  seed(conversationId: string, botAppId: string, ref: Record<string, unknown>, teamsType: string | null = null): void {
    this.rows.set(`${conversationId}|${botAppId}`, { conversation_id: conversationId, bot_app_id: botAppId, ref, teams_type: teamsType });
  }

  snapshot(): string[] {
    return [...this.rows.values()]
      .map((r) => `${r.conversation_id}|${r.bot_app_id}`)
      .sort();
  }

  pool(): Pool {
    return { query: (text: string, params?: unknown[]) => this.query(text, params ?? []) } as unknown as Pool;
  }

  private convRows(conversationId: unknown): FakeRow[] {
    return [...this.rows.values()].filter((r) => r.conversation_id === conversationId);
  }

  private async query(text: string, params: unknown[]): Promise<{ rows: unknown[]; rowCount: number }> {
    this.queries.push({ text, params });
    if (this.failEverythingWith) throw this.failEverythingWith;
    const t = text.replace(/\s+/g, ' ').trim();
    const missingColumn = (): never => {
      throw pgError('42703', 'column "bot_app_id" of relation "teams_conversation_refs" does not exist');
    };

    if (t.startsWith('INSERT INTO teams_conversation_refs (conversation_id, bot_app_id')) {
      if (!this.hasBotAppId) missingColumn();
      const [conv, bot, refJson, teamsType] = params as [string, string, string, string | null];
      this.rows.set(`${conv}|${bot}`, { conversation_id: conv, bot_app_id: bot, ref: JSON.parse(refJson), teams_type: teamsType });
      return { rows: [], rowCount: 1 };
    }
    if (t.startsWith('INSERT INTO teams_conversation_refs (conversation_id, ref')) {
      // Legacy statement: single row per conversation, keyed conv-only.
      const [conv, refJson, teamsType] = params as [string, string, string | null];
      for (const existing of this.convRows(conv)) this.rows.delete(`${existing.conversation_id}|${existing.bot_app_id}`);
      this.rows.set(`${conv}|`, { conversation_id: conv, bot_app_id: '', ref: JSON.parse(refJson), teams_type: teamsType });
      return { rows: [], rowCount: 1 };
    }
    if (t.startsWith('SELECT ref, teams_type, bot_app_id')) {
      if (!this.hasBotAppId) missingColumn();
      const rows = this.convRows(params[0]).map((r) => ({ ref: r.ref, teams_type: r.teams_type, bot_app_id: r.bot_app_id }));
      return { rows, rowCount: rows.length };
    }
    if (t.startsWith('SELECT ref, teams_type FROM')) {
      const rows = this.convRows(params[0]).map((r) => ({ ref: r.ref, teams_type: r.teams_type }));
      return { rows: rows.slice(0, 1), rowCount: Math.min(rows.length, 1) };
    }
    if (t.startsWith('UPDATE teams_conversation_refs SET bot_app_id')) {
      if (!this.hasBotAppId) missingColumn();
      const [target] = params as [string];
      let n = 0;
      for (const [key, row] of [...this.rows]) {
        if (row.bot_app_id !== '') continue;
        const conflict = [...this.rows.values()].some((r2) => r2.conversation_id === row.conversation_id && r2.bot_app_id === target);
        if (conflict) continue;
        this.rows.delete(key);
        this.rows.set(`${row.conversation_id}|${target}`, { ...row, bot_app_id: target });
        n += 1;
      }
      return { rows: [], rowCount: n };
    }
    if (t.startsWith('DELETE FROM teams_conversation_refs')) {
      if (!this.hasBotAppId) missingColumn();
      let n = 0;
      for (const [key, row] of [...this.rows]) {
        if (row.bot_app_id !== '') continue;
        this.rows.delete(key);
        n += 1;
      }
      return { rows: [], rowCount: n };
    }
    throw new Error(`FakeRefTable: unexpected query: ${t}`);
  }
}

function ref(marker: string): Record<string, unknown> {
  return { serviceUrl: ALLOWED_SERVICE_URL, conversation: { id: marker } };
}

describe('normalizeTeamsBotAppId', () => {
  it('lower-cases and trims; blank/undefined mean "default bot"', () => {
    assert.equal(normalizeTeamsBotAppId('AppA-GUID'), 'appa-guid');
    assert.equal(normalizeTeamsBotAppId('  appb  '), 'appb');
    assert.equal(normalizeTeamsBotAppId(''), undefined);
    assert.equal(normalizeTeamsBotAppId('   '), undefined);
    assert.equal(normalizeTeamsBotAppId(undefined), undefined);
    assert.equal(LEGACY_BOT_APP_ID, '');
  });
});

describe('PgTeamsConversationRefStore — per-bot schema (migration 0010 applied)', () => {
  it('keys writes on (conversation_id, bot_app_id) and isolates loads per bot', async () => {
    const table = new FakeRefTable(true);
    const store = new PgTeamsConversationRefStore(table.pool());

    await store.save('conv-1', ref('from-a') as Partial<ConversationReference>, 'groupChat', 'AppA');
    await store.save('conv-1', ref('from-b') as Partial<ConversationReference>, 'groupChat', 'AppB');
    assert.deepEqual(table.snapshot(), ['conv-1|appa', 'conv-1|appb']);

    const forA = await store.load('conv-1', 'appa');
    const forB = await store.load('conv-1', 'AppB'); // casing must not matter
    assert.equal((forA?.ref.conversation as { id?: string } | undefined)?.id, 'from-a');
    assert.equal((forB?.ref.conversation as { id?: string } | undefined)?.id, 'from-b');

    // Cross-bot isolation: a ref written by bot A/B must never load for bot C.
    assert.equal(await store.load('conv-1', 'appc'), undefined);
  });

  it('upserts on the composite key: same (conversation, bot) overwrites, never duplicates', async () => {
    const table = new FakeRefTable(true);
    const store = new PgTeamsConversationRefStore(table.pool());
    await store.save('conv-1', ref('v1') as Partial<ConversationReference>, undefined, 'appa');
    await store.save('conv-1', ref('v2') as Partial<ConversationReference>, undefined, 'AppA');
    assert.deepEqual(table.snapshot(), ['conv-1|appa']);
    const hit = await store.load('conv-1', 'appa');
    assert.equal((hit?.ref.conversation as { id?: string } | undefined)?.id, 'v2');
  });

  it("serves pre-backfill legacy rows ('') to the default bot ONLY", async () => {
    const table = new FakeRefTable(true);
    table.seed('conv-legacy', '', ref('legacy'), 'personal');
    const store = new PgTeamsConversationRefStore(table.pool(), undefined, { defaultBotAppId: 'AppA' });

    const asDefaultImplicit = await store.load('conv-legacy');
    const asDefaultExplicit = await store.load('conv-legacy', 'appa');
    assert.equal((asDefaultImplicit?.ref.conversation as { id?: string } | undefined)?.id, 'legacy');
    assert.equal(asDefaultImplicit?.teamsType, 'personal');
    assert.equal((asDefaultExplicit?.ref.conversation as { id?: string } | undefined)?.id, 'legacy');
    assert.equal(await store.load('conv-legacy', 'appb'), undefined);
  });

  it('without a configured default, an unattributed read falls through only when ONE bot holds the conversation', async () => {
    const table = new FakeRefTable(true);
    table.seed('conv-solo', 'appa', ref('solo'));
    table.seed('conv-ambiguous', 'appa', ref('a'));
    table.seed('conv-ambiguous', 'appb', ref('b'));
    const store = new PgTeamsConversationRefStore(table.pool());

    const solo = await store.load('conv-solo');
    assert.equal((solo?.ref.conversation as { id?: string } | undefined)?.id, 'solo');
    assert.equal(await store.load('conv-ambiguous'), undefined);
  });

  it('drops poisoned rows per bot — serviceUrl outside Bot Framework domains never reaches the caller', async () => {
    const table = new FakeRefTable(true);
    table.seed('conv-1', 'appa', { serviceUrl: 'https://evil.example/api', conversation: { id: 'poison' } });
    table.seed('conv-1', 'appb', ref('clean'));
    const logs: string[] = [];
    const store = new PgTeamsConversationRefStore(table.pool(), (m) => logs.push(m));

    assert.equal(await store.load('conv-1', 'appa'), undefined);
    assert.ok(logs.some((l) => l.includes('serviceUrl outside Bot Framework domains')));
    const clean = await store.load('conv-1', 'appb');
    assert.equal((clean?.ref.conversation as { id?: string } | undefined)?.id, 'clean');
  });
});

describe('PgTeamsConversationRefStore — pre-migration kernel (no bot_app_id column)', () => {
  it('save degrades to the legacy single-key statement after one 42703 and stays there', async () => {
    const table = new FakeRefTable(false);
    const store = new PgTeamsConversationRefStore(table.pool());

    await store.save('conv-1', ref('v1') as Partial<ConversationReference>, 'personal', 'appa');
    await store.save('conv-1', ref('v2') as Partial<ConversationReference>, 'personal', 'appa');

    assert.deepEqual(table.snapshot(), ['conv-1|']);
    const compositeAttempts = table.queries.filter((q) => q.text.includes('bot_app_id'));
    assert.equal(compositeAttempts.length, 1, 'probes the per-bot statement exactly once');
  });

  it('load treats the single row as the default bot and refuses explicit foreign bots', async () => {
    const table = new FakeRefTable(false);
    const store = new PgTeamsConversationRefStore(table.pool(), undefined, { defaultBotAppId: 'appa' });
    await store.save('conv-1', ref('r') as Partial<ConversationReference>, 'personal', 'appa');

    const asDefault = await store.load('conv-1');
    assert.equal((asDefault?.ref.conversation as { id?: string } | undefined)?.id, 'r');
    assert.equal(asDefault?.teamsType, 'personal');
    // "missing column = default bot": bot B must not get bot A's reference.
    assert.equal(await store.load('conv-1', 'appb'), undefined);
  });

  it('load NEVER throws on the message path — pool trouble degrades to cache-only', async () => {
    const table = new FakeRefTable(true);
    table.failEverythingWith = pgError('57P01', 'terminating connection');
    const logs: string[] = [];
    const store = new PgTeamsConversationRefStore(table.pool(), (m) => logs.push(m));
    assert.equal(await store.load('conv-1', 'appa'), undefined);
    assert.ok(logs.some((l) => l.includes('conversation-ref load failed')));
  });

  it('save still surfaces non-schema errors to the fire-and-forget caller', async () => {
    const table = new FakeRefTable(true);
    table.failEverythingWith = pgError('57P01', 'terminating connection');
    const store = new PgTeamsConversationRefStore(table.pool());
    await assert.rejects(() => store.save('conv-1', ref('x') as Partial<ConversationReference>));
  });
});

describe('PgTeamsConversationRefStore.backfillLegacyBotAppId', () => {
  it('attributes legacy rows to the target bot NON-destructively and re-runs as a no-op', async () => {
    const table = new FakeRefTable(true);
    table.seed('conv-1', '', ref('legacy-1'));
    table.seed('conv-2', '', ref('legacy-2-stale'));
    table.seed('conv-2', 'appa', ref('fresh-2')); // newer default-bot row already exists
    table.seed('conv-3', 'appb', ref('other-bot'));
    const store = new PgTeamsConversationRefStore(table.pool(), undefined, { defaultBotAppId: 'appa' });

    await store.backfillLegacyBotAppId('AppA');
    const after = table.snapshot();
    // Coordinator decision (#860 W0a): NOTHING is deleted. conv-1's sentinel
    // row is attributed to the default bot; conv-2's stale sentinel STAYS
    // (its conversation already has a fresher default-bot row, and reads
    // prefer the exact match over the sentinel).
    assert.deepEqual(after, ['conv-1|appa', 'conv-2|', 'conv-2|appa', 'conv-3|appb']);
    assert.equal((table.rows.get('conv-2|appa')?.ref.conversation as { id?: string } | undefined)?.id, 'fresh-2', 'fresh row wins over stale sentinel');

    await store.backfillLegacyBotAppId('AppA');
    assert.deepEqual(table.snapshot(), after, 'second run is a no-op');

    const loaded = await store.load('conv-1', 'appa');
    assert.equal((loaded?.ref.conversation as { id?: string } | undefined)?.id, 'legacy-1');
    // The stale sentinel never shadows the fresh default-bot row.
    const loaded2 = await store.load('conv-2', 'appa');
    assert.equal((loaded2?.ref.conversation as { id?: string } | undefined)?.id, 'fresh-2');
  });

  it('is a silent no-op on a pre-migration kernel and never throws on pool trouble', async () => {
    const legacy = new FakeRefTable(false);
    const logs: string[] = [];
    await new PgTeamsConversationRefStore(legacy.pool(), (m) => logs.push(m)).backfillLegacyBotAppId('appa');
    assert.ok(logs.some((l) => l.includes('backfill skipped')));

    const broken = new FakeRefTable(true);
    broken.failEverythingWith = pgError('57P01', 'terminating connection');
    await new PgTeamsConversationRefStore(broken.pool(), (m) => logs.push(m)).backfillLegacyBotAppId('appa');
    assert.ok(logs.some((l) => l.includes('backfill failed')));
  });
});

// ————————————————————————————————————————————————————————————————————————
// TeamsConversationReferenceCache — per-(bot, conversation) keying
// ————————————————————————————————————————————————————————————————————————

function fakeContext(input: { conversationId: string; teamsType?: string; recipientId?: string }): TurnContext {
  return {
    activity: {
      type: 'message',
      channelId: 'msteams',
      serviceUrl: ALLOWED_SERVICE_URL,
      conversation: { id: input.conversationId, ...(input.teamsType ? { conversationType: input.teamsType } : {}) },
      from: { id: '29:user', name: 'User' },
      recipient: { id: input.recipientId ?? '28:bot', name: 'Omadia' },
    },
  } as unknown as TurnContext;
}

class RecordingPersistence implements TeamsConversationRefPersistence {
  readonly saves: { conversationId: string; teamsType?: string; botAppId?: string }[] = [];
  readonly loads: { conversationId: string; botAppId?: string }[] = [];
  loadResult: { ref: Partial<ConversationReference>; teamsType?: string; botAppId?: string } | undefined;

  async save(conversationId: string, _ref: Partial<ConversationReference>, teamsType?: string, botAppId?: string): Promise<void> {
    this.saves.push({ conversationId, ...(teamsType !== undefined ? { teamsType } : {}), ...(botAppId !== undefined ? { botAppId } : {}) });
  }

  async load(conversationId: string, botAppId?: string): Promise<{ ref: Partial<ConversationReference>; teamsType?: string; botAppId?: string } | undefined> {
    this.loads.push({ conversationId, ...(botAppId !== undefined ? { botAppId } : {}) });
    return this.loadResult;
  }
}

describe('TeamsConversationReferenceCache — per-bot keying (#860 W0a)', () => {
  it('captures the same conversation once PER BOT (recipient 28:<appId>) and isolates lookups', () => {
    const cache = new TeamsConversationReferenceCache();
    cache.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'groupChat', recipientId: '28:AppA' }));
    cache.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'groupChat', recipientId: '28:AppB' }));

    const forA = cache.get('conv-1', 'appa');
    const forB = cache.get('conv-1', 'AppB'); // casing must not matter
    assert.equal(forA?.ref.bot?.id, '28:AppA');
    assert.equal(forB?.ref.bot?.id, '28:AppB');
    // Cross-bot isolation: bot C sees nothing, and the ambiguous
    // no-bot-given lookup refuses to guess between A and B.
    assert.equal(cache.get('conv-1', 'appc'), undefined);
    assert.equal(cache.get('conv-1'), undefined);
  });

  it('single-bot compatibility: lookups without a botAppId keep hitting when one bot holds the conversation', () => {
    const cache = new TeamsConversationReferenceCache();
    cache.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'personal' }));
    const hit = cache.get('conv-1');
    assert.equal(hit?.teamsType, 'personal');
    assert.equal(hit?.ref.conversation?.id, 'conv-1');
  });

  it('setDefaultBotAppId resolves no-bot-given lookups deterministically in multi-bot setups', () => {
    const cache = new TeamsConversationReferenceCache();
    cache.setDefaultBotAppId('AppB');
    cache.capture(fakeContext({ conversationId: 'conv-1', recipientId: '28:AppA' }));
    cache.capture(fakeContext({ conversationId: 'conv-1', recipientId: '28:AppB' }));
    assert.equal(cache.get('conv-1')?.ref.bot?.id, '28:AppB');
  });

  it('write-throughs carry the capturing bot to the persistence layer and dedupe per (bot, conversation)', () => {
    const cache = new TeamsConversationReferenceCache();
    const persistence = new RecordingPersistence();
    cache.attachPersistence(persistence);

    cache.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'groupChat', recipientId: '28:AppA' }));
    cache.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'groupChat', recipientId: '28:AppA' })); // unchanged → deduped
    cache.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'groupChat', recipientId: '28:AppB' }));

    assert.deepEqual(persistence.saves, [
      { conversationId: 'conv-1', teamsType: 'groupChat', botAppId: 'appa' },
      { conversationId: 'conv-1', teamsType: 'groupChat', botAppId: 'appb' },
    ]);
  });

  it('getOrLoad passes the requested bot through to the store and re-seeds the cache per bot', async () => {
    const cache = new TeamsConversationReferenceCache();
    const persistence = new RecordingPersistence();
    persistence.loadResult = { ref: { serviceUrl: ALLOWED_SERVICE_URL, conversation: { id: 'conv-9' } } as Partial<ConversationReference>, teamsType: 'channel' };
    cache.attachPersistence(persistence);

    const loaded = await cache.getOrLoad('conv-9', 'AppA');
    assert.equal(loaded?.teamsType, 'channel');
    assert.deepEqual(persistence.loads, [{ conversationId: 'conv-9', botAppId: 'appa' }]);

    // Re-seeded under bot A: the next lookup is a cache hit, no second load.
    assert.equal(cache.get('conv-9', 'appa')?.teamsType, 'channel');
    await cache.getOrLoad('conv-9', 'appa');
    assert.equal(persistence.loads.length, 1);
    // …and bot B still sees nothing for that conversation.
    persistence.loadResult = undefined;
    assert.equal(cache.get('conv-9', 'appb'), undefined);
    assert.equal(await cache.getOrLoad('conv-9', 'appb'), undefined);
  });

  it('production call shape: getOrLoad without botAppId, then capture(), then get() — one bot, one cache slot', async () => {
    // Regression (#860 W0a review): with no explicit botAppId AND no default
    // configured, a restart-load used to seed under the '' sentinel while the
    // next inbound capture() seeded under the real recipient appId — two keys
    // for ONE bot, permanent ambiguity, get() missing forever.
    const cache = new TeamsConversationReferenceCache();
    const persistence = new RecordingPersistence();
    persistence.loadResult = {
      ref: { serviceUrl: ALLOWED_SERVICE_URL, conversation: { id: 'conv-r' } } as Partial<ConversationReference>,
      teamsType: 'channel',
    };
    cache.attachPersistence(persistence);

    // Restart: proactive path loads with no bot argument (teamsGroupPrimitives
    // call shape) — seeds under the sentinel, and the lookup still hits.
    assert.ok(await cache.getOrLoad('conv-r'));
    assert.ok(cache.get('conv-r'), 'sentinel-seeded entry must be readable');

    // First inbound activity re-captures under the real bot key.
    cache.capture(fakeContext({ conversationId: 'conv-r', teamsType: 'channel', recipientId: '28:AppA' }));

    // The sync lookup must now hit the real entry — no ambiguity, no
    // permanent DB round-trips, no doubled capacity for the conversation.
    const hit = cache.get('conv-r');
    assert.equal(hit?.ref.bot?.id, '28:AppA', 'lookup resolves to the recaptured real-bot entry');
    const before = persistence.loads.length;
    assert.ok(await cache.getOrLoad('conv-r'), 'getOrLoad stays a cache hit');
    assert.equal(persistence.loads.length, before, 'no extra store round-trip after recapture');
  });

  it('getOrLoad re-seeds under the OWNING bot the store reports, matching a later capture()', async () => {
    const cache = new TeamsConversationReferenceCache();
    const persistence = new RecordingPersistence();
    persistence.loadResult = {
      ref: { serviceUrl: ALLOWED_SERVICE_URL, conversation: { id: 'conv-own' }, bot: { id: '28:AppA', name: 'bot' } } as Partial<ConversationReference>,
      teamsType: 'channel',
      botAppId: 'appa',
    };
    cache.attachPersistence(persistence);

    await cache.getOrLoad('conv-own');
    // Seeded under 'appa' (the owning bot), NOT under the '' sentinel.
    assert.ok(cache.get('conv-own', 'appa'), 'entry keyed under the owning bot');
    assert.ok(cache.get('conv-own'), 'unique-holder lookup resolves it too');
    cache.capture(fakeContext({ conversationId: 'conv-own', teamsType: 'channel', recipientId: '28:AppA' }));
    assert.equal(cache.get('conv-own')?.ref.bot?.id, '28:AppA');
  });

  it('eviction under the composite key drops the oldest (bot, conversation) entry only — never another bot\'s', () => {
    const cache = new TeamsConversationReferenceCache();
    // Oldest entry: bot A in the shared conversation.
    cache.capture(fakeContext({ conversationId: 'conv-shared', recipientId: '28:AppA' }));
    // Fill to the 500-entry cap with distinct bot-A conversations.
    for (let i = 0; i < 499; i += 1) {
      cache.capture(fakeContext({ conversationId: `conv-${i}`, recipientId: '28:AppA' }));
    }
    // 501st entry: bot B in the SAME shared conversation → evicts A|conv-shared.
    cache.capture(fakeContext({ conversationId: 'conv-shared', recipientId: '28:AppB' }));

    assert.equal(cache.get('conv-shared', 'appa'), undefined, 'oldest entry evicted');
    assert.equal(cache.get('conv-shared', 'appb')?.ref.bot?.id, '28:AppB', "other bot's entry for the same conversation survives");
    assert.equal(cache.get('conv-0', 'appa')?.ref.conversation?.id, 'conv-0', 'unrelated entries survive');
  });
});

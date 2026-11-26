// omadia#838 — the channel's tables come from `migrations/` (applied by the
// kernel under `permissions.sql`, ADR-0050). The statements the stores still
// run at first use must be exactly the migration's, or the two drift.
import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

import { QUESTION_CARDS_DDL } from '../src/teamsQuestionCardStore.js';
import { MEETING_STATES_DDL } from '../src/teamsMeetingStateStore.js';

const root = process.cwd();
const normalise = (sql: string): string[] =>
  sql
    .replace(/--[^\n]*\n/g, '')
    .split(';')
    .map((s) => s.replace(/\s+/g, ' ').trim())
    .filter((s) => s.length > 0);

describe('omadia#838 — plugin migrations', () => {
  it('declares permissions.sql with the kernel-owned ledger name and ships the directory', () => {
    const manifest = readFileSync(join(root, 'manifest.yaml'), 'utf8');
    assert.match(manifest, /\n  sql:\n    migrations: "migrations"\n    ledger: "plg_omadia_channel_teams_migrations"\n/);
    assert.deepEqual(readdirSync(join(root, 'migrations')).sort(), ['0001_question_cards.sql', '0002_meeting_states.sql']);
    assert.match(readFileSync(join(root, 'scripts/build-zip.mjs'), 'utf8'), /OPTIONAL_DIRS = \[[^\]]*'migrations'/);
  });

  for (const [file, ddl] of [
    ['0001_question_cards.sql', QUESTION_CARDS_DDL],
    ['0002_meeting_states.sql', MEETING_STATES_DDL],
  ] as const) {
    it(`${file} is exactly the store's DDL, idempotent`, () => {
      const statements = normalise(readFileSync(join(root, 'migrations', file), 'utf8'));
      assert.deepEqual(statements, ddl.map((s) => s.replace(/\s+/g, ' ').trim()));
      for (const s of statements) assert.match(s, /^CREATE (TABLE|INDEX) IF NOT EXISTS /);
    });
  }
});

import { resolveGraphPool } from '../src/graphPoolAccess.js';

describe('omadia#838 — graphPool without the operator grant', () => {
  it('a refused pool is logged with the remedy and the channel runs without a database', () => {
    const warnings: string[] = [];
    const pool = resolveGraphPool(
      {
        get: () => {
          const err = new Error("plugin '@omadia/channel-teams' may not use 'graphPool': ungranted");
          err.name = 'SqlPermissionError';
          throw err;
        },
      },
      (m) => warnings.push(m),
    );
    assert.equal(pool, undefined);
    assert.equal(warnings.length, 1);
    assert.match(warnings[0]!, /ungranted.*Grant the plugin's database permission/s);
  });

  it('any other error is not swallowed', () => {
    assert.throws(
      () => resolveGraphPool({ get: () => { throw new Error('registry broken'); } }, () => undefined),
      /registry broken/,
    );
  });

  it('a granted pool is returned as is', () => {
    const fake = { query: async () => ({ rows: [] }) };
    assert.equal(resolveGraphPool({ get: <T>() => fake as T }, () => undefined), fake);
  });
});

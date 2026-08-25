import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// #860 W0a — per-bot identity surface. Imported relatively (not via the
// `@omadia/channel-teams` package entry) because the `src/index.ts` export
// line is owned by the config-wiring unit; the module itself is
// dependency-free, so esbuild bundles it straight from src/.
import {
  DEFAULT_TEAMS_BOT_APP_TYPE,
  findTeamsBotByAppId,
  findTeamsBotBySlug,
  getDefaultTeamsBot,
  parseTeamsBotKey,
  TEAMS_BOT_KEY_PREFIX,
  teamsBotKey,
  teamsBotLogLabel,
} from '../src/teamsBotIdentity.js';
import type { TeamsBotIdentity } from '../src/teamsBotIdentity.js';

const BOTS: readonly TeamsBotIdentity[] = [
  {
    botSlug: 'default',
    appId: '11111111-aaaa-bbbb-cccc-222222222222',
    tenantId: 'tenant-1',
    appPasswordSecretRef: 'microsoft_app_password',
    appType: 'MultiTenant',
    displayName: 'Omadia',
  },
  {
    botSlug: 'hr-agent',
    appId: '33333333-dddd-eeee-ffff-444444444444',
    tenantId: 'tenant-1',
    appPasswordSecretRef: 'teams_bot_hr_agent_password',
    appType: 'SingleTenant',
    displayName: 'HR Agent',
  },
];

describe('teamsBotKey', () => {
  it('builds exactly the 28:<appId> key the directory publishes and the resolver matches', () => {
    // Must stay byte-identical to the inline `28:${appId}` construction it
    // replaces (channelKeyDirectory catch-all row / activity.recipient.id).
    assert.equal(
      teamsBotKey('11111111-aaaa-bbbb-cccc-222222222222'),
      '28:11111111-aaaa-bbbb-cccc-222222222222',
    );
    assert.equal(TEAMS_BOT_KEY_PREFIX, '28:');
  });
});

describe('parseTeamsBotKey', () => {
  it('round-trips with teamsBotKey', () => {
    const appId = BOTS[1]!.appId;
    assert.equal(parseTeamsBotKey(teamsBotKey(appId)), appId);
  });

  it('rejects non-bot keys', () => {
    assert.equal(parseTeamsBotKey('29:user-id'), undefined);
    assert.equal(parseTeamsBotKey('19:meeting@thread.v2'), undefined);
    assert.equal(parseTeamsBotKey('28:'), undefined);
    assert.equal(parseTeamsBotKey(''), undefined);
  });
});

describe('findTeamsBotBySlug', () => {
  it('finds by exact slug', () => {
    assert.equal(findTeamsBotBySlug(BOTS, 'hr-agent'), BOTS[1]);
    assert.equal(findTeamsBotBySlug(BOTS, 'default'), BOTS[0]);
  });

  it('returns undefined for unknown or non-exact slugs', () => {
    assert.equal(findTeamsBotBySlug(BOTS, 'HR-Agent'), undefined);
    assert.equal(findTeamsBotBySlug(BOTS, 'missing'), undefined);
    assert.equal(findTeamsBotBySlug([], 'default'), undefined);
  });
});

describe('findTeamsBotByAppId', () => {
  it('finds by app id, case-insensitively (GUID casing may differ)', () => {
    assert.equal(
      findTeamsBotByAppId(BOTS, '33333333-dddd-eeee-ffff-444444444444'),
      BOTS[1],
    );
    assert.equal(
      findTeamsBotByAppId(BOTS, '33333333-DDDD-EEEE-FFFF-444444444444'),
      BOTS[1],
    );
  });

  it('returns undefined for unknown app ids', () => {
    assert.equal(findTeamsBotByAppId(BOTS, 'not-a-configured-app'), undefined);
    assert.equal(findTeamsBotByAppId([], BOTS[0]!.appId), undefined);
  });
});

describe('getDefaultTeamsBot', () => {
  it('returns the first configured bot (teams_bots[0] — the legacy alias target)', () => {
    assert.equal(getDefaultTeamsBot(BOTS), BOTS[0]);
  });

  it('returns undefined for an empty list', () => {
    assert.equal(getDefaultTeamsBot([]), undefined);
  });
});

describe('DEFAULT_TEAMS_BOT_APP_TYPE', () => {
  it('is SingleTenant (MultiTenant creation deprecated 07/2025)', () => {
    assert.equal(DEFAULT_TEAMS_BOT_APP_TYPE, 'SingleTenant');
  });
});

describe('teamsBotLogLabel', () => {
  it('names the bot by slug + display name only — never appId or secret ref', () => {
    const label = teamsBotLogLabel(BOTS[1]!);
    assert.equal(label, 'hr-agent ("HR Agent")');
    assert.ok(!label.includes(BOTS[1]!.appId));
    assert.ok(!label.includes(BOTS[1]!.appPasswordSecretRef));
  });
});

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// #860 W0a — teams_bots[] config surface (config-wiring unit). Runtime
// values come from the built dist/ via the package entry; types relatively.
import {
  DEFAULT_TEAMS_BOT_APP_TYPE,
  parseTeamsBotsConfig,
  TeamsBotsConfigError,
} from '@omadia/channel-teams';
import type { TeamsBotIdentity } from '../src/teamsBotIdentity.js';

const FULL_ENTRY = {
  botSlug: 'hr-agent',
  appId: '33333333-dddd-eeee-ffff-444444444444',
  tenantId: 'tenant-1',
  appPasswordSecretRef: 'teams_bot_hr_agent_password',
  appType: 'SingleTenant',
  displayName: 'HR Agent',
};

describe('parseTeamsBotsConfig — empty / absent surfaces', () => {
  it('returns [] for undefined (caller falls back to the legacy shim)', () => {
    assert.deepEqual(parseTeamsBotsConfig(undefined), []);
  });

  it('returns [] for null, empty string, and blank string', () => {
    assert.deepEqual(parseTeamsBotsConfig(null), []);
    assert.deepEqual(parseTeamsBotsConfig(''), []);
    assert.deepEqual(parseTeamsBotsConfig('   '), []);
  });

  it('returns [] for an empty array', () => {
    assert.deepEqual(parseTeamsBotsConfig([]), []);
  });
});

describe('parseTeamsBotsConfig — valid input shapes', () => {
  it('accepts a real array (install-registry value) and keeps order', () => {
    const bots = parseTeamsBotsConfig([
      { ...FULL_ENTRY, botSlug: 'default', appId: 'app-a', displayName: 'Omadia' },
      FULL_ENTRY,
    ]);
    assert.equal(bots.length, 2);
    assert.equal(bots[0]!.botSlug, 'default');
    assert.equal(bots[1]!.botSlug, 'hr-agent');
  });

  it('accepts the same array as a JSON string (setup-wizard string field)', () => {
    const bots = parseTeamsBotsConfig(JSON.stringify([FULL_ENTRY]));
    const expected: TeamsBotIdentity = {
      botSlug: 'hr-agent',
      appId: '33333333-dddd-eeee-ffff-444444444444',
      tenantId: 'tenant-1',
      appPasswordSecretRef: 'teams_bot_hr_agent_password',
      appType: 'SingleTenant',
      displayName: 'HR Agent',
    };
    assert.deepEqual(bots, [expected]);
  });

  it('defaults appType to SingleTenant — Azure deprecated MultiTenant creation 07/2025', () => {
    const { appType, ...withoutType } = FULL_ENTRY;
    void appType;
    const bots = parseTeamsBotsConfig([withoutType]);
    assert.equal(bots[0]!.appType, 'SingleTenant');
    assert.equal(bots[0]!.appType, DEFAULT_TEAMS_BOT_APP_TYPE);
  });

  it('keeps an explicitly configured MultiTenant / UserAssignedMSI appType per bot', () => {
    const bots = parseTeamsBotsConfig([
      { ...FULL_ENTRY, botSlug: 'legacy', appId: 'app-legacy', appType: 'MultiTenant' },
      { ...FULL_ENTRY, botSlug: 'msi', appId: 'app-msi', appType: 'UserAssignedMSI' },
    ]);
    assert.equal(bots[0]!.appType, 'MultiTenant');
    assert.equal(bots[1]!.appType, 'UserAssignedMSI');
  });

  it('defaults displayName to the botSlug', () => {
    const { displayName, ...withoutName } = FULL_ENTRY;
    void displayName;
    const bots = parseTeamsBotsConfig([withoutName]);
    assert.equal(bots[0]!.displayName, 'hr-agent');
  });

  it('trims whitespace off entry fields', () => {
    const bots = parseTeamsBotsConfig([
      { ...FULL_ENTRY, botSlug: ' hr-agent ', appId: ' app-x ' },
    ]);
    assert.equal(bots[0]!.botSlug, 'hr-agent');
    assert.equal(bots[0]!.appId, 'app-x');
  });
});

describe('parseTeamsBotsConfig — validation failures are loud, never silent', () => {
  const cases: readonly [string, unknown][] = [
    ['non-array JSON value', { botSlug: 'x' }],
    ['invalid JSON string', '[{not json'],
    ['non-object entry', ['bot']],
    ['missing botSlug', [{ ...FULL_ENTRY, botSlug: undefined }]],
    // Coordinator decision (#860 W0a): slugs match ^[a-z0-9][a-z0-9-]{0,62}$ —
    // no slashes, no uppercase, no whitespace, no dots/underscores.
    ['botSlug with a slash (hr/agent)', [{ ...FULL_ENTRY, botSlug: 'hr/agent' }]],
    ['botSlug with uppercase', [{ ...FULL_ENTRY, botSlug: 'HR-Agent' }]],
    ['botSlug with whitespace', [{ ...FULL_ENTRY, botSlug: 'hr agent' }]],
    ['botSlug with a dot', [{ ...FULL_ENTRY, botSlug: 'hr.agent' }]],
    ['botSlug with an underscore', [{ ...FULL_ENTRY, botSlug: 'hr_agent' }]],
    ['botSlug starting with a dash', [{ ...FULL_ENTRY, botSlug: '-hr' }]],
    ['botSlug longer than 63 chars', [{ ...FULL_ENTRY, botSlug: 'a'.repeat(64) }]],
    ['missing appId', [{ ...FULL_ENTRY, appId: '' }]],
    ['missing tenantId', [{ ...FULL_ENTRY, tenantId: '  ' }]],
    ['missing appPasswordSecretRef', [{ ...FULL_ENTRY, appPasswordSecretRef: undefined }]],
    ['unknown appType', [{ ...FULL_ENTRY, appType: 'PublicClient' }]],
  ];
  for (const [name, raw] of cases) {
    it(`rejects ${name}`, () => {
      assert.throws(() => parseTeamsBotsConfig(raw), TeamsBotsConfigError);
    });
  }

  it('rejects duplicate botSlugs', () => {
    assert.throws(
      () => parseTeamsBotsConfig([FULL_ENTRY, { ...FULL_ENTRY, appId: 'other-app' }]),
      /duplicate botSlug/,
    );
  });

  it('rejects duplicate appIds case-insensitively (Azure serialises lowercase)', () => {
    assert.throws(
      () =>
        parseTeamsBotsConfig([
          FULL_ENTRY,
          { ...FULL_ENTRY, botSlug: 'other', appId: FULL_ENTRY.appId.toUpperCase() },
        ]),
      /duplicate appId/,
    );
  });

  it('rejects an inline app password instead of quietly storing a credential in config', () => {
    assert.throws(
      () => parseTeamsBotsConfig([{ ...FULL_ENTRY, appPassword: 'hunter2' }]),
      /inline app password/,
    );
    assert.throws(
      () => parseTeamsBotsConfig([{ ...FULL_ENTRY, app_password: 'hunter2' }]),
      /inline app password/,
    );
  });

  it('never leaks secret material in the error message', () => {
    try {
      parseTeamsBotsConfig([{ ...FULL_ENTRY, appPassword: 'hunter2' }]);
      assert.fail('expected a TeamsBotsConfigError');
    } catch (err) {
      assert.ok(err instanceof TeamsBotsConfigError);
      assert.ok(!err.message.includes('hunter2'));
    }
  });
});

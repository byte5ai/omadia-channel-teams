import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// #860 W2 — teams_agent_apps[] config surface (agent-apps-config unit).
// Runtime values import relatively (precedent: teamsBotIdentity.test.ts) so
// this test executes before the wiring unit lands the src/index.ts
// re-export; the wiring unit's full-suite run re-verifies against dist.
import {
  parseTeamsAgentAppsConfig,
  TeamsAgentAppsConfigError,
  type TeamsAgentApp,
} from '../src/teamsAgentApps.js';

const FULL_ENTRY = {
  agentSlug: 'hr-agent',
  teamsAppExternalId: '11111111-aaaa-bbbb-cccc-222222222222',
  teamsAppId: '33333333-dddd-eeee-ffff-444444444444',
  displayName: 'HR Agent',
};

describe('parseTeamsAgentAppsConfig — empty / absent surfaces (feature off)', () => {
  it('returns [] for undefined — deployments without the field stay unchanged', () => {
    assert.deepEqual(parseTeamsAgentAppsConfig(undefined), []);
  });

  it('returns [] for null, empty string, and blank string', () => {
    assert.deepEqual(parseTeamsAgentAppsConfig(null), []);
    assert.deepEqual(parseTeamsAgentAppsConfig(''), []);
    assert.deepEqual(parseTeamsAgentAppsConfig('   '), []);
  });

  it('returns [] for an empty array', () => {
    assert.deepEqual(parseTeamsAgentAppsConfig([]), []);
  });
});

describe('parseTeamsAgentAppsConfig — valid input shapes', () => {
  it('accepts a real array (install-registry value) and keeps order', () => {
    const apps = parseTeamsAgentAppsConfig([
      { ...FULL_ENTRY, agentSlug: 'accounting', teamsAppExternalId: 'ext-a', teamsAppId: undefined },
      FULL_ENTRY,
    ]);
    assert.equal(apps.length, 2);
    assert.equal(apps[0]!.agentSlug, 'accounting');
    assert.equal(apps[1]!.agentSlug, 'hr-agent');
  });

  it('accepts the same array as a JSON string (setup-wizard string field)', () => {
    const apps = parseTeamsAgentAppsConfig(JSON.stringify([FULL_ENTRY]));
    const expected: TeamsAgentApp = {
      agentSlug: 'hr-agent',
      teamsAppExternalId: '11111111-aaaa-bbbb-cccc-222222222222',
      teamsAppId: '33333333-dddd-eeee-ffff-444444444444',
      displayName: 'HR Agent',
    };
    assert.deepEqual(apps, [expected]);
  });

  it('keeps teamsAppId undefined when omitted (installer falls back to catalog lookup)', () => {
    const { teamsAppId, ...withoutAppId } = FULL_ENTRY;
    void teamsAppId;
    const apps = parseTeamsAgentAppsConfig([withoutAppId]);
    assert.equal(apps[0]!.teamsAppId, undefined);
  });

  it('treats empty-string and null teamsAppId as not configured', () => {
    assert.equal(
      parseTeamsAgentAppsConfig([{ ...FULL_ENTRY, teamsAppId: '' }])[0]!.teamsAppId,
      undefined,
    );
    assert.equal(
      parseTeamsAgentAppsConfig([{ ...FULL_ENTRY, teamsAppId: null }])[0]!.teamsAppId,
      undefined,
    );
  });

  it('defaults displayName to the agentSlug', () => {
    const { displayName, ...withoutName } = FULL_ENTRY;
    void displayName;
    const apps = parseTeamsAgentAppsConfig([withoutName]);
    assert.equal(apps[0]!.displayName, 'hr-agent');
  });

  it('trims whitespace off entry fields', () => {
    const apps = parseTeamsAgentAppsConfig([
      { ...FULL_ENTRY, agentSlug: ' hr-agent ', teamsAppExternalId: ' ext-x ', teamsAppId: ' app-x ' },
    ]);
    assert.equal(apps[0]!.agentSlug, 'hr-agent');
    assert.equal(apps[0]!.teamsAppExternalId, 'ext-x');
    assert.equal(apps[0]!.teamsAppId, 'app-x');
  });
});

describe('parseTeamsAgentAppsConfig — validation failures are loud, never silent', () => {
  const cases: readonly [string, unknown][] = [
    ['non-array JSON value', { agentSlug: 'x' }],
    ['invalid JSON string', '[{not json'],
    ['non-object entry', ['hr-agent']],
    ['missing agentSlug', [{ ...FULL_ENTRY, agentSlug: undefined }]],
    // agentSlug follows the conductor-registry slug rule:
    // ^[a-z0-9](?:[a-z0-9-]{0,62}[a-z0-9])?$
    ['agentSlug with a slash (hr/agent)', [{ ...FULL_ENTRY, agentSlug: 'hr/agent' }]],
    ['agentSlug with uppercase', [{ ...FULL_ENTRY, agentSlug: 'HR-Agent' }]],
    ['agentSlug with whitespace', [{ ...FULL_ENTRY, agentSlug: 'hr agent' }]],
    ['agentSlug with an underscore', [{ ...FULL_ENTRY, agentSlug: 'hr_agent' }]],
    ['agentSlug ending with a dash', [{ ...FULL_ENTRY, agentSlug: 'hr-' }]],
    ['agentSlug longer than 64 chars', [{ ...FULL_ENTRY, agentSlug: 'a'.repeat(65) }]],
    ['missing teamsAppExternalId', [{ ...FULL_ENTRY, teamsAppExternalId: '' }]],
    ['blank teamsAppExternalId', [{ ...FULL_ENTRY, teamsAppExternalId: '   ' }]],
    ['non-string teamsAppId', [{ ...FULL_ENTRY, teamsAppId: 42 }]],
  ];
  for (const [name, raw] of cases) {
    it(`rejects ${name}`, () => {
      assert.throws(() => parseTeamsAgentAppsConfig(raw), TeamsAgentAppsConfigError);
    });
  }

  it('rejects duplicate agentSlugs', () => {
    assert.throws(
      () =>
        parseTeamsAgentAppsConfig([
          FULL_ENTRY,
          { ...FULL_ENTRY, teamsAppExternalId: 'ext-other', teamsAppId: undefined },
        ]),
      /duplicate agentSlug/,
    );
  });

  it('rejects duplicate teamsAppExternalIds case-insensitively (Graph serialises lowercase)', () => {
    assert.throws(
      () =>
        parseTeamsAgentAppsConfig([
          FULL_ENTRY,
          {
            ...FULL_ENTRY,
            agentSlug: 'other',
            teamsAppExternalId: FULL_ENTRY.teamsAppExternalId.toUpperCase(),
            teamsAppId: undefined,
          },
        ]),
      /duplicate teamsAppExternalId/,
    );
  });

  it('rejects duplicate teamsAppIds case-insensitively', () => {
    assert.throws(
      () =>
        parseTeamsAgentAppsConfig([
          FULL_ENTRY,
          {
            ...FULL_ENTRY,
            agentSlug: 'other',
            teamsAppExternalId: 'ext-other',
            teamsAppId: FULL_ENTRY.teamsAppId.toUpperCase(),
          },
        ]),
      /duplicate teamsAppId/,
    );
  });

  it('prefixes every error with the config key so operators can locate it', () => {
    try {
      parseTeamsAgentAppsConfig('not-json{');
      assert.fail('expected a TeamsAgentAppsConfigError');
    } catch (err) {
      assert.ok(err instanceof TeamsAgentAppsConfigError);
      assert.ok(err.message.startsWith('teams_agent_apps config invalid:'));
      assert.equal(err.name, 'TeamsAgentAppsConfigError');
    }
  });
});

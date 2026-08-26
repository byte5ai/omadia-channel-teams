import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// #860 W2 — installer error mapping. Imported relatively (not via the
// `@omadia/channel-teams` package entry) because the `src/index.ts` export
// line is owned by the wiring unit; the module only has type imports, so
// esbuild bundles it straight from src/ (precedent: teamsBotIdentity.test.ts).
import {
  AUTO_INSTALL_MARKER_TTL_MS,
  CONSENT_NEGATIVE_CACHE_TTL_MS,
  DEFAULT_THROTTLE_WAIT_MS,
  isConsentMissingError,
  isProvisioningThrottledError,
  MAX_THROTTLE_RETRIES,
  MAX_THROTTLE_WAIT_MS,
  TeamsAgentInstaller,
  TeamsAutoInstallMarker,
  TeamsConsentNegativeCache,
  teamsAgentAppLogLabel,
} from '../src/teamsAgentInstaller.js';
import type {
  AgentAppInstallOutcome,
  TeamsAgentAppTarget,
} from '../src/teamsAgentInstaller.js';
import type {
  GetCatalogAppResult,
  Idempotent,
  InstallToTeamRequest,
  TeamAppInstallation,
  TeamsProvisionerAccessor,
} from '../src/kernel-types.js';

const TEAM_ID = 'aaaaaaaa-1111-2222-3333-bbbbbbbbbbbb';
const TENANT_ID = 'cccccccc-4444-5555-6666-dddddddddddd';

const APP_HR: TeamsAgentAppTarget = {
  agentSlug: 'odoo-hr',
  teamsAppExternalId: '11111111-aaaa-bbbb-cccc-000000000001',
  displayName: 'Odoo HR',
};

const APP_ACCOUNTING: TeamsAgentAppTarget = {
  agentSlug: 'odoo-accounting',
  teamsAppExternalId: '11111111-aaaa-bbbb-cccc-000000000002',
  teamsAppId: 'catalog-accounting',
  displayName: 'Odoo Accounting',
};

// Test doubles for the connector's typed errors: the real classes live in
// `@omadia/integration-microsoft365` (peer-only, not installed), and the
// installer's guards branch on `name` + structure — exactly what these
// doubles replicate.
class FakeConsentMissingError extends Error {
  readonly missingScopes: readonly string[];
  readonly resource: 'graph' | 'arm';
  constructor(missingScopes: readonly string[], resource: 'graph' | 'arm' = 'graph') {
    super('consent_missing');
    this.name = 'ConsentMissingError';
    this.missingScopes = missingScopes;
    this.resource = resource;
  }
}

class FakeProvisioningThrottledError extends Error {
  readonly retryAfterSeconds?: number;
  readonly resource: 'graph' | 'arm';
  constructor(retryAfterSeconds?: number, resource: 'graph' | 'arm' = 'graph') {
    super('provisioning_throttled');
    this.name = 'ProvisioningThrottledError';
    this.resource = resource;
    if (retryAfterSeconds !== undefined) this.retryAfterSeconds = retryAfterSeconds;
  }
}

type InstallStep = (input: InstallToTeamRequest) => Promise<Idempotent<TeamAppInstallation>>;
type LookupStep = (externalId: string) => Promise<GetCatalogAppResult>;

/** Recording provisioner fake. `lookup` undefined = a 0.3.0 connector
 *  WITHOUT `getCatalogApp` (feature-detection path). */
function fakeProvisioner(opts: {
  install: InstallStep;
  lookup?: LookupStep;
}): {
  provisioner: TeamsProvisionerAccessor;
  installCalls: InstallToTeamRequest[];
  lookupCalls: string[];
} {
  const installCalls: InstallToTeamRequest[] = [];
  const lookupCalls: string[] = [];
  const provisioner: TeamsProvisionerAccessor = {
    installToTeam: async (input) => {
      installCalls.push(input);
      return opts.install(input);
    },
    ...(opts.lookup
      ? {
          getCatalogApp: async (input: { teamsAppExternalId: string }) => {
            lookupCalls.push(input.teamsAppExternalId);
            return opts.lookup!(input.teamsAppExternalId);
          },
        }
      : {}),
  };
  return { provisioner, installCalls, lookupCalls };
}

const installed = (
  teamsAppId: string,
  outcome: 'created' | 'already-existed' = 'created',
): Idempotent<TeamAppInstallation> => ({
  outcome,
  value: { teamId: TEAM_ID, teamsAppId },
});

function makeInstaller(
  provisioner: TeamsProvisionerAccessor,
  apps: readonly TeamsAgentAppTarget[],
  extra?: {
    now?: () => number;
    sleeps?: number[];
    consentCache?: TeamsConsentNegativeCache;
    marker?: TeamsAutoInstallMarker;
    maxRunSleepBudgetMs?: number;
  },
): TeamsAgentInstaller {
  return new TeamsAgentInstaller({
    provisioner,
    apps,
    ...(extra?.consentCache ? { consentCache: extra.consentCache } : {}),
    ...(extra?.marker ? { marker: extra.marker } : {}),
    ...(extra?.maxRunSleepBudgetMs !== undefined
      ? { maxRunSleepBudgetMs: extra.maxRunSleepBudgetMs }
      : {}),
    sleep: async (ms) => {
      extra?.sleeps?.push(ms);
    },
  });
}

const run = (installer: TeamsAgentInstaller) =>
  installer.installAgentApps({ teamId: TEAM_ID, tenantId: TENANT_ID });

function expectKind<K extends AgentAppInstallOutcome['kind']>(
  outcome: AgentAppInstallOutcome | undefined,
  kind: K,
): Extract<AgentAppInstallOutcome, { kind: K }> {
  assert.ok(outcome, 'expected an outcome');
  assert.equal(outcome.kind, kind);
  return outcome as Extract<AgentAppInstallOutcome, { kind: K }>;
}

describe('teamsAgentInstaller catalog resolution', () => {
  it('uses the configured teamsAppId without a catalog lookup', async () => {
    const { provisioner, installCalls, lookupCalls } = fakeProvisioner({
      install: async (input) => installed(input.teamsAppId),
      lookup: async () => ({ found: true, teamsAppId: 'never-used' }),
    });
    const result = await run(makeInstaller(provisioner, [APP_ACCOUNTING]));

    const outcome = expectKind(result.outcomes[0], 'installed');
    assert.equal(outcome.teamsAppId, 'catalog-accounting');
    assert.equal(outcome.outcome, 'created');
    assert.deepEqual(lookupCalls, []);
    assert.equal(installCalls[0]?.teamId, TEAM_ID);
  });

  it('feature-detects getCatalogApp and installs the resolved catalog id', async () => {
    const { provisioner, installCalls, lookupCalls } = fakeProvisioner({
      install: async (input) => installed(input.teamsAppId),
      lookup: async () => ({ found: true, teamsAppId: 'catalog-hr' }),
    });
    const result = await run(makeInstaller(provisioner, [APP_HR]));

    const outcome = expectKind(result.outcomes[0], 'installed');
    assert.equal(outcome.teamsAppId, 'catalog-hr');
    assert.deepEqual(lookupCalls, [APP_HR.teamsAppExternalId]);
    assert.equal(installCalls[0]?.teamsAppId, 'catalog-hr');
  });

  it('lookup miss → not-in-catalog fallback, no install call', async () => {
    const { provisioner, installCalls } = fakeProvisioner({
      install: async (input) => installed(input.teamsAppId),
      lookup: async () => ({ found: false }),
    });
    const result = await run(makeInstaller(provisioner, [APP_HR]));

    const outcome = expectKind(result.outcomes[0], 'fallback');
    assert.equal(outcome.reason, 'not-in-catalog');
    assert.equal(outcome.teamsAppExternalId, APP_HR.teamsAppExternalId);
    assert.deepEqual(installCalls, []);
    assert.equal(result.autoInstallMarked, false);
  });

  it('connector without getCatalogApp (0.3.0) → not-in-catalog fallback, never a crash', async () => {
    const { provisioner, installCalls } = fakeProvisioner({
      install: async (input) => installed(input.teamsAppId),
    });
    const result = await run(makeInstaller(provisioner, [APP_HR]));

    const outcome = expectKind(result.outcomes[0], 'fallback');
    assert.equal(outcome.reason, 'not-in-catalog');
    assert.deepEqual(installCalls, []);
  });
});

describe('teamsAgentInstaller error mapping', () => {
  it('consent 403 → typed fallback with missingScopes, never a throw', async () => {
    const { provisioner } = fakeProvisioner({
      install: async () => {
        throw new FakeConsentMissingError(['TeamsAppInstallation.ReadWriteForTeam.All']);
      },
    });
    const result = await run(makeInstaller(provisioner, [APP_ACCOUNTING]));

    const outcome = expectKind(result.outcomes[0], 'fallback');
    assert.equal(outcome.reason, 'consent-missing');
    assert.deepEqual(outcome.missingScopes, [
      'TeamsAppInstallation.ReadWriteForTeam.All',
    ]);
    assert.equal(outcome.teamsAppId, 'catalog-accounting');
    assert.equal(result.autoInstallMarked, false);
  });

  it('409 already-existed → success outcome, no marker', async () => {
    const { provisioner } = fakeProvisioner({
      install: async (input) => installed(input.teamsAppId, 'already-existed'),
    });
    const marker = new TeamsAutoInstallMarker();
    const result = await run(
      makeInstaller(provisioner, [APP_ACCOUNTING], { marker }),
    );

    const outcome = expectKind(result.outcomes[0], 'installed');
    assert.equal(outcome.outcome, 'already-existed');
    assert.equal(result.autoInstallMarked, false);
    assert.equal(marker.probe(TEAM_ID), false);
  });

  it('throttled → bounded retry honouring retryAfterSeconds, then success', async () => {
    let attempts = 0;
    const { provisioner } = fakeProvisioner({
      install: async (input) => {
        attempts += 1;
        if (attempts <= 2) throw new FakeProvisioningThrottledError(1);
        return installed(input.teamsAppId);
      },
    });
    const sleeps: number[] = [];
    const result = await run(
      makeInstaller(provisioner, [APP_ACCOUNTING], { sleeps }),
    );

    expectKind(result.outcomes[0], 'installed');
    assert.equal(attempts, 3);
    assert.deepEqual(sleeps, [1000, 1000]);
  });

  it('throttle retries exhausted → typed failed outcome carrying the hint', async () => {
    let attempts = 0;
    const { provisioner } = fakeProvisioner({
      install: async () => {
        attempts += 1;
        throw new FakeProvisioningThrottledError(7);
      },
    });
    const sleeps: number[] = [];
    const result = await run(
      makeInstaller(provisioner, [APP_ACCOUNTING], { sleeps }),
    );

    const outcome = expectKind(result.outcomes[0], 'failed');
    assert.equal(outcome.reason, 'throttled');
    assert.equal(outcome.retryAfterSeconds, 7);
    assert.equal(attempts, 1 + MAX_THROTTLE_RETRIES);
    assert.equal(sleeps.length, MAX_THROTTLE_RETRIES);
  });

  it('hintless throttle → DEFAULT_THROTTLE_WAIT_MS between retries', async () => {
    const { provisioner } = fakeProvisioner({
      install: async () => {
        throw new FakeProvisioningThrottledError();
      },
    });
    const sleeps: number[] = [];
    const result = await run(
      makeInstaller(provisioner, [APP_ACCOUNTING], { sleeps }),
    );

    const outcome = expectKind(result.outcomes[0], 'failed');
    assert.equal(outcome.reason, 'throttled');
    assert.equal(outcome.retryAfterSeconds, undefined);
    assert.deepEqual(sleeps, [
      DEFAULT_THROTTLE_WAIT_MS,
      DEFAULT_THROTTLE_WAIT_MS,
    ]);
  });

  it('a hint AT the cap is still honoured (clamp boundary)', async () => {
    const { provisioner } = fakeProvisioner({
      install: async () => {
        throw new FakeProvisioningThrottledError(MAX_THROTTLE_WAIT_MS / 1000);
      },
    });
    const sleeps: number[] = [];
    const result = await run(
      makeInstaller(provisioner, [APP_ACCOUNTING], { sleeps }),
    );

    expectKind(result.outcomes[0], 'failed');
    assert.deepEqual(sleeps, [MAX_THROTTLE_WAIT_MS, MAX_THROTTLE_WAIT_MS]);
  });

  it('a hint ABOVE the cap is non-retryable — no sleep, full hint preserved', async () => {
    let attempts = 0;
    const { provisioner } = fakeProvisioner({
      install: async () => {
        attempts += 1;
        throw new FakeProvisioningThrottledError(3600);
      },
    });
    const sleeps: number[] = [];
    const result = await run(
      makeInstaller(provisioner, [APP_ACCOUNTING], { sleeps }),
    );

    const outcome = expectKind(result.outcomes[0], 'failed');
    assert.equal(outcome.reason, 'throttled');
    // The connector's backpressure contract: a long window means
    // "reschedule", not "fire more requests into it".
    assert.equal(attempts, 1);
    assert.deepEqual(sleeps, []);
    assert.equal(outcome.retryAfterSeconds, 3600);
  });

  it('run-level sleep budget caps total installer-owned sleep', async () => {
    const { provisioner } = fakeProvisioner({
      install: async () => {
        throw new FakeProvisioningThrottledError(15);
      },
    });
    const sleeps: number[] = [];
    const result = await run(
      makeInstaller(provisioner, [APP_ACCOUNTING], {
        sleeps,
        maxRunSleepBudgetMs: 25_000,
      }),
    );

    const outcome = expectKind(result.outcomes[0], 'failed');
    assert.equal(outcome.reason, 'throttled');
    // First wait (15s) fits the 25s budget; the second (15s > 10s left)
    // does not — the retry loop stops instead of sleeping.
    assert.deepEqual(sleeps, [15_000]);
  });

  it('a throttled app short-circuits the remaining apps of the run', async () => {
    const APP_SECOND: TeamsAgentAppTarget = {
      agentSlug: 'odoo-hr',
      teamsAppExternalId: '11111111-aaaa-bbbb-cccc-000000000001',
      teamsAppId: 'catalog-hr',
      displayName: 'Odoo HR',
    };
    let attempts = 0;
    const { provisioner, installCalls } = fakeProvisioner({
      install: async () => {
        attempts += 1;
        throw new FakeProvisioningThrottledError(7);
      },
    });
    const sleeps: number[] = [];
    const result = await run(
      makeInstaller(provisioner, [APP_ACCOUNTING, APP_SECOND], { sleeps }),
    );

    const first = expectKind(result.outcomes[0], 'failed');
    assert.equal(first.reason, 'throttled');
    assert.equal(first.message, 'provisioning_throttled');
    const second = expectKind(result.outcomes[1], 'failed');
    assert.equal(second.reason, 'throttled');
    assert.equal(second.message, 'provisioning_throttled_run_short_circuit');
    assert.equal(second.retryAfterSeconds, 7);
    // Only the FIRST app reached Graph — a 429 is tenant-wide; calling on
    // for the second app would only deepen the throttle window.
    assert.equal(attempts, 1 + MAX_THROTTLE_RETRIES);
    assert.equal(installCalls.length, 1 + MAX_THROTTLE_RETRIES);
  });

  it('unexpected error → typed failed outcome, message path never throws', async () => {
    const { provisioner } = fakeProvisioner({
      install: async () => {
        throw new Error('graph exploded');
      },
    });
    const result = await run(makeInstaller(provisioner, [APP_ACCOUNTING]));

    const outcome = expectKind(result.outcomes[0], 'failed');
    assert.equal(outcome.reason, 'error');
    assert.equal(outcome.message, 'graph exploded');
  });
});

describe('teamsAgentInstaller consent negative cache', () => {
  it('403 short-circuits the remaining apps of the same run', async () => {
    const { provisioner, installCalls } = fakeProvisioner({
      install: async () => {
        throw new FakeConsentMissingError(['AppCatalog.ReadWrite.All']);
      },
      lookup: async () => ({ found: true, teamsAppId: 'catalog-hr' }),
    });
    const result = await run(
      makeInstaller(provisioner, [APP_ACCOUNTING, APP_HR]),
    );

    expectKind(result.outcomes[0], 'fallback');
    const second = expectKind(result.outcomes[1], 'fallback');
    assert.equal(second.reason, 'consent-cached');
    assert.deepEqual(second.missingScopes, ['AppCatalog.ReadWrite.All']);
    // Only the first app reached Graph.
    assert.equal(installCalls.length, 1);
  });

  it('caches per tenant across runs (case-insensitive key) — no retry storm', async () => {
    let calls = 0;
    const cache = new TeamsConsentNegativeCache();
    const { provisioner } = fakeProvisioner({
      install: async () => {
        calls += 1;
        throw new FakeConsentMissingError(['AppCatalog.ReadWrite.All']);
      },
    });
    const installer = makeInstaller(provisioner, [APP_ACCOUNTING], {
      consentCache: cache,
    });

    await run(installer);
    assert.equal(calls, 1);

    // Same tenant, different casing → still served from the cache.
    const second = await installer.installAgentApps({
      teamId: TEAM_ID,
      tenantId: TENANT_ID.toUpperCase(),
    });
    assert.equal(calls, 1);
    const outcome = expectKind(second.outcomes[0], 'fallback');
    assert.equal(outcome.reason, 'consent-cached');
  });

  it('scope sets are unioned per tenant, not replaced — the card names every grant', () => {
    const cache = new TeamsConsentNegativeCache();
    cache.markConsentMissing(TENANT_ID, ['AppCatalog.ReadWrite.All']);
    cache.markConsentMissing(TENANT_ID, [
      'TeamsAppInstallation.ReadWriteForTeam.All',
      'AppCatalog.ReadWrite.All',
    ]);

    assert.deepEqual(cache.get(TENANT_ID)?.missingScopes, [
      'AppCatalog.ReadWrite.All',
      'TeamsAppInstallation.ReadWriteForTeam.All',
    ]);
  });

  it('cache entries expire after the TTL — a granted consent is retried', () => {
    let nowMs = 1_000_000;
    const cache = new TeamsConsentNegativeCache({ now: () => nowMs });
    cache.markConsentMissing(TENANT_ID, ['AppCatalog.ReadWrite.All']);

    assert.ok(cache.get(TENANT_ID));
    nowMs += CONSENT_NEGATIVE_CACHE_TTL_MS + 1;
    assert.equal(cache.get(TENANT_ID), undefined);
  });
});

describe('teamsAgentInstaller auto-install intro marker', () => {
  it('fresh install marks the team; probe is non-consuming', async () => {
    const marker = new TeamsAutoInstallMarker();
    const { provisioner } = fakeProvisioner({
      install: async (input) => installed(input.teamsAppId),
    });
    const result = await run(
      makeInstaller(provisioner, [APP_ACCOUNTING], { marker }),
    );

    assert.equal(result.autoInstallMarked, true);
    assert.equal(marker.probe(TEAM_ID), true);
    // Non-consuming: every bot of the flock can correlate the same event.
    assert.equal(marker.probe(TEAM_ID), true);
    // Trim+lowercase canonicalisation — casing drift never splits the key.
    assert.equal(marker.probe(` ${TEAM_ID.toUpperCase()} `), true);
  });

  it('marker expires after its short TTL', () => {
    let nowMs = 5_000_000;
    const marker = new TeamsAutoInstallMarker({ now: () => nowMs });
    marker.mark(TEAM_ID);

    assert.equal(marker.probe(TEAM_ID), true);
    nowMs += AUTO_INSTALL_MARKER_TTL_MS + 1;
    assert.equal(marker.probe(TEAM_ID), false);
  });
});

describe('teamsAgentInstaller guards and labels', () => {
  it('error guards branch on name + structure, not instanceof', () => {
    assert.ok(isConsentMissingError(new FakeConsentMissingError(['S'])));
    assert.ok(isProvisioningThrottledError(new FakeProvisioningThrottledError(3)));
    assert.equal(isConsentMissingError(new Error('consent_missing')), false);
    assert.equal(isProvisioningThrottledError('throttled'), false);
    const impostor = new Error('x');
    impostor.name = 'ConsentMissingError';
    assert.equal(isConsentMissingError(impostor), false);
  });

  it('log label carries slug + display name only — never app ids', () => {
    const label = teamsAgentAppLogLabel(APP_ACCOUNTING);
    assert.equal(label, 'odoo-accounting ("Odoo Accounting")');
    assert.ok(!label.includes('catalog-accounting'));
    assert.ok(!label.includes(APP_ACCOUNTING.teamsAppExternalId));
  });
});

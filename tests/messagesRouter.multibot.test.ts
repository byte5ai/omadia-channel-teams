import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// #860 W0a — per-bot CloudAdapter + BF credential factory. The router module
// value-imports the CJS botbuilder runtime, so (like teamsGroupPrimitives.test.ts)
// runtime values come via the package entry — bundling `messagesRouter.ts`
// relatively would drag botbuilder's CJS into the ESM test bundle. Type-only
// imports are erased by esbuild and may stay relative.
import { createTeamsRouter } from '@omadia/channel-teams';
import type { TeamsRouterDeps } from '../src/messagesRouter.js';
import type { TeamsBotIdentity } from '../src/teamsBotIdentity.js';
import type { TeamsBot } from '../src/teamsBot.js';

const fakeBot = { run: async () => {} } as unknown as TeamsBot;

function identity(overrides: Partial<TeamsBotIdentity> & Pick<TeamsBotIdentity, 'botSlug' | 'appId'>): TeamsBotIdentity {
  return {
    tenantId: 'tenant-common',
    appPasswordSecretRef: `${overrides.botSlug}_password`,
    appType: 'SingleTenant',
    displayName: overrides.botSlug,
    ...overrides,
  };
}

const BOT_A = identity({ botSlug: 'hr', appId: '11111111-aaaa-4bbb-8ccc-111111111111', tenantId: 'tenant-a' });
const BOT_B = identity({ botSlug: 'accounting', appId: '22222222-aaaa-4bbb-8ccc-222222222222', tenantId: 'tenant-b' });

function multiBotDeps(): TeamsRouterDeps {
  return {
    bot: fakeBot,
    bots: [
      { identity: BOT_A, appPassword: 'secret-a' },
      { identity: BOT_B, appPassword: 'secret-b' },
    ],
  };
}

describe('createTeamsRouter — per-bot adapter/credential isolation (#860 W0a)', () => {
  it('builds one runtime per configured bot, in config order', () => {
    const artifacts = createTeamsRouter(multiBotDeps());
    assert.equal(artifacts.botRuntimes.length, 2);
    assert.equal(artifacts.botRuntimes[0]?.identity?.botSlug, 'hr');
    assert.equal(artifacts.botRuntimes[1]?.identity?.botSlug, 'accounting');
    assert.equal(artifacts.defaultBotRuntime, artifacts.botRuntimes[0]);
    // Back-compat surface: `adapter` is the DEFAULT bot's adapter.
    assert.equal(artifacts.adapter, artifacts.botRuntimes[0]?.adapter);
  });

  it('gives every bot its OWN CloudAdapter and credential factory', () => {
    const artifacts = createTeamsRouter(multiBotDeps());
    const [a, b] = artifacts.botRuntimes;
    assert.ok(a && b);
    assert.notEqual(a.adapter, b.adapter);
    assert.notEqual(a.credentialsFactory, b.credentialsFactory);
  });

  it("wires each factory with that bot's OWN appId / password / tenant — a turn for bot B never authenticates with bot A's password", () => {
    const artifacts = createTeamsRouter(multiBotDeps());
    const [a, b] = artifacts.botRuntimes;
    assert.ok(a && b);
    assert.equal(a.credentialsFactory.appId, BOT_A.appId);
    assert.equal(a.credentialsFactory.password, 'secret-a');
    assert.equal(a.credentialsFactory.tenantId, 'tenant-a');
    assert.equal(b.credentialsFactory.appId, BOT_B.appId);
    assert.equal(b.credentialsFactory.password, 'secret-b');
    assert.equal(b.credentialsFactory.tenantId, 'tenant-b');
  });

  it('resolves runtimes by appId case-insensitively (operator-pasted GUID casing)', () => {
    const artifacts = createTeamsRouter(multiBotDeps());
    assert.equal(artifacts.getBotRuntimeByAppId(BOT_B.appId.toUpperCase())?.identity?.botSlug, 'accounting');
    assert.equal(artifacts.getBotRuntimeByAppId('99999999-0000-4000-8000-000000000000'), undefined);
    assert.equal(artifacts.getBotRuntimeByAppId(''), undefined);
  });

  it('legacy scalar deps behave like a single-entry bot list (shim back-compat)', () => {
    const artifacts = createTeamsRouter({
      bot: fakeBot,
      appId: 'legacy-app-id',
      appPassword: 'legacy-secret',
      appType: 'MultiTenant',
      appTenantId: 'legacy-tenant',
    });
    assert.equal(artifacts.botRuntimes.length, 1);
    const runtime = artifacts.defaultBotRuntime;
    assert.equal(runtime.identity, undefined);
    assert.equal(runtime.appId, 'legacy-app-id');
    assert.equal(runtime.credentialsFactory.appId, 'legacy-app-id');
    assert.equal(runtime.credentialsFactory.password, 'legacy-secret');
    assert.equal(runtime.credentialsFactory.tenantId, 'legacy-tenant');
    assert.equal(artifacts.adapter, runtime.adapter);
  });

  it('rejects an empty bot list', () => {
    assert.throws(
      () => createTeamsRouter({ bot: fakeBot, bots: [] }),
      /at least one bot identity/,
    );
  });

  it('rejects duplicate appIds (also across casing) without leaking the appId in the error', () => {
    assert.throws(
      () => createTeamsRouter({
        bot: fakeBot,
        bots: [
          { identity: BOT_A, appPassword: 'secret-a' },
          { identity: identity({ botSlug: 'clone', appId: BOT_A.appId.toUpperCase() }), appPassword: 'secret-c' },
        ],
      }),
      (err: unknown) => {
        assert.ok(err instanceof Error);
        assert.match(err.message, /duplicate bot appId/);
        assert.match(err.message, /clone/);
        assert.ok(!err.message.toLowerCase().includes(BOT_A.appId.toLowerCase()));
        return true;
      },
    );
  });

  it('rejects a bot with an empty appId', () => {
    assert.throws(
      () => createTeamsRouter({
        bot: fakeBot,
        bots: [{ identity: identity({ botSlug: 'blank', appId: '  ' }), appPassword: 'x' }],
      }),
      /empty appId/,
    );
  });
});

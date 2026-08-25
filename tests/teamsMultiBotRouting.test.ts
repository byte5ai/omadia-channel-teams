import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// #860 W0a — per-bot message routes. Runtime values via the package entry
// (messagesRouter value-imports CJS botbuilder — see messagesRouter.multibot.test.ts),
// type-only imports relative (erased by esbuild).
import { createTeamsRouter } from '@omadia/channel-teams';
import type { TeamsBotIdentity } from '../src/teamsBotIdentity.js';
import type { TeamsBotRuntime, TeamsRouterArtifacts } from '../src/messagesRouter.js';
import type { TeamsBot } from '../src/teamsBot.js';

const fakeBot = { run: async () => {} } as unknown as TeamsBot;

const APP_ID_A = 'aaaaaaaa-1111-4111-8111-aaaaaaaaaaaa';
const APP_ID_B = 'bbbbbbbb-2222-4222-8222-bbbbbbbbbbbb';

function identity(botSlug: string, appId: string): TeamsBotIdentity {
  return {
    botSlug,
    appId,
    tenantId: 'tenant-common',
    appPasswordSecretRef: `${botSlug}_password`,
    appType: 'SingleTenant',
    displayName: botSlug,
  };
}

/** Replace a runtime adapter's `process` with a recorder — the real one
 *  would run Bot-Framework JWT authentication against the request. The
 *  route handlers look the method up on the adapter instance at request
 *  time, so patching the instance is enough. */
function recordProcessedTurns(runtime: TeamsBotRuntime): string[] {
  const urls: string[] = [];
  (runtime.adapter as unknown as {
    process: (req: { url?: string }, res: { end: () => void }, logic: unknown) => Promise<void>;
  }).process = async (req, res) => {
    urls.push(req.url ?? '');
    res.end();
  };
  return urls;
}

/** Outcome of pushing one fake request through the express Router. */
type DispatchResult =
  | { kind: 'handled'; status: number; body?: unknown }
  | { kind: 'fallthrough' };

/**
 * Drive the Router directly as the middleware function it is (no listening
 * server needed): express parses `req.url`/`req.method`, fills `req.params`,
 * and calls our handler; unmatched paths call `next()` — reported here as
 * `fallthrough` (the kernel app's own 404 in production).
 */
function dispatch(
  artifacts: TeamsRouterArtifacts,
  method: string,
  url: string,
): Promise<DispatchResult> {
  return new Promise((resolve, reject) => {
    const req = { method, url, headers: {}, body: undefined };
    const res = {
      statusCode: 200,
      setHeader: () => res,
      status: (code: number) => {
        res.statusCode = code;
        return res;
      },
      json: (body: unknown) => {
        resolve({ kind: 'handled', status: res.statusCode, body });
        return res;
      },
      end: () => {
        resolve({ kind: 'handled', status: res.statusCode });
        return res;
      },
    };
    const next = (err?: unknown) => {
      if (err) reject(err instanceof Error ? err : new Error(String(err)));
      else resolve({ kind: 'fallthrough' });
    };
    (artifacts.router as unknown as (rq: unknown, rs: unknown, nx: unknown) => void)(
      req,
      res,
      next,
    );
  });
}

function makeMultiBotFixture() {
  const artifacts = createTeamsRouter({
    bot: fakeBot,
    bots: [
      { identity: identity('hr', APP_ID_A), appPassword: 'secret-a' },
      { identity: identity('accounting', APP_ID_B), appPassword: 'secret-b' },
    ],
  });
  const [runtimeA, runtimeB] = artifacts.botRuntimes;
  assert.ok(runtimeA && runtimeB);
  return {
    artifacts,
    turnsA: recordProcessedTurns(runtimeA),
    turnsB: recordProcessedTurns(runtimeB),
  };
}

describe('per-bot message routes (#860 W0a)', () => {
  it("POST /teams/:botSlug/messages dispatches to THAT bot's own adapter", async () => {
    const { artifacts, turnsA, turnsB } = makeMultiBotFixture();

    const forB = await dispatch(artifacts, 'POST', '/teams/accounting/messages');
    assert.deepEqual(forB, { kind: 'handled', status: 200 });
    assert.equal(turnsA.length, 0);
    assert.equal(turnsB.length, 1);

    const forA = await dispatch(artifacts, 'POST', '/teams/hr/messages');
    assert.deepEqual(forA, { kind: 'handled', status: 200 });
    assert.equal(turnsA.length, 1);
    assert.equal(turnsB.length, 1);
  });

  it('POST /messages (live legacy path) keeps hitting the DEFAULT bot', async () => {
    const { artifacts, turnsA, turnsB } = makeMultiBotFixture();
    const result = await dispatch(artifacts, 'POST', '/messages');
    assert.deepEqual(result, { kind: 'handled', status: 200 });
    assert.equal(turnsA.length, 1); // default = bots[0]
    assert.equal(turnsB.length, 0);
  });

  it('POST /teams/messages (manifest-documented alias) hits the DEFAULT bot', async () => {
    const { artifacts, turnsA, turnsB } = makeMultiBotFixture();
    const result = await dispatch(artifacts, 'POST', '/teams/messages');
    assert.deepEqual(result, { kind: 'handled', status: 200 });
    assert.equal(turnsA.length, 1);
    assert.equal(turnsB.length, 0);
  });

  it("an unknown botSlug 404s and NEVER falls through to the default bot's credentials", async () => {
    const { artifacts, turnsA, turnsB } = makeMultiBotFixture();
    const result = await dispatch(artifacts, 'POST', '/teams/ghost/messages');
    assert.equal(result.kind, 'handled');
    assert.equal(result.kind === 'handled' && result.status, 404);
    assert.deepEqual(result.kind === 'handled' ? result.body : undefined, {
      code: 'teams.unknown_bot',
      message: 'no bot is configured under this slug',
    });
    assert.equal(turnsA.length, 0);
    assert.equal(turnsB.length, 0);
  });

  it('slug matching is exact — no case-folding, no prefix match', async () => {
    const { artifacts, turnsA, turnsB } = makeMultiBotFixture();
    for (const url of ['/teams/HR/messages', '/teams/h/messages', '/teams/hrx/messages']) {
      const result = await dispatch(artifacts, 'POST', url);
      assert.equal(result.kind === 'handled' && result.status, 404, `expected 404 for ${url}`);
    }
    assert.equal(turnsA.length, 0);
    assert.equal(turnsB.length, 0);
  });

  it('only POST is bound — GET on the messaging paths falls through', async () => {
    const { artifacts, turnsA, turnsB } = makeMultiBotFixture();
    for (const url of ['/messages', '/teams/messages', '/teams/hr/messages']) {
      const result = await dispatch(artifacts, 'GET', url);
      assert.deepEqual(result, { kind: 'fallthrough' }, `expected fallthrough for GET ${url}`);
    }
    assert.equal(turnsA.length, 0);
    assert.equal(turnsB.length, 0);
  });

  it('legacy scalar deps: both no-slug aliases serve the single bot, every slug 404s', async () => {
    const artifacts = createTeamsRouter({
      bot: fakeBot,
      appId: 'legacy-app-id',
      appPassword: 'legacy-secret',
      appType: 'MultiTenant',
      appTenantId: 'legacy-tenant',
    });
    const turns = recordProcessedTurns(artifacts.defaultBotRuntime);

    assert.deepEqual(await dispatch(artifacts, 'POST', '/messages'), { kind: 'handled', status: 200 });
    assert.deepEqual(await dispatch(artifacts, 'POST', '/teams/messages'), { kind: 'handled', status: 200 });
    assert.equal(turns.length, 2);

    const result = await dispatch(artifacts, 'POST', '/teams/anything/messages');
    assert.equal(result.kind === 'handled' && result.status, 404);
    assert.equal(turns.length, 2);
    assert.equal(artifacts.getBotRuntimeBySlug('anything'), undefined);
  });

  it('getBotRuntimeBySlug resolves configured slugs exactly', () => {
    const { artifacts } = makeMultiBotFixture();
    assert.equal(artifacts.getBotRuntimeBySlug('hr')?.appId, APP_ID_A);
    assert.equal(artifacts.getBotRuntimeBySlug('accounting')?.appId, APP_ID_B);
    assert.equal(artifacts.getBotRuntimeBySlug('HR'), undefined);
    assert.equal(artifacts.getBotRuntimeBySlug(''), undefined);
  });

  it('rejects duplicate botSlugs at construction (the later bot would be shadowed)', () => {
    assert.throws(
      () => createTeamsRouter({
        bot: fakeBot,
        bots: [
          { identity: identity('hr', APP_ID_A), appPassword: 'secret-a' },
          { identity: identity('hr', APP_ID_B), appPassword: 'secret-b' },
        ],
      }),
      /duplicate botSlug 'hr'/,
    );
  });

  it('rejects an empty botSlug (the bot would be unreachable)', () => {
    assert.throws(
      () => createTeamsRouter({
        bot: fakeBot,
        bots: [{ identity: identity('  ', APP_ID_A), appPassword: 'secret-a' }],
      }),
      /empty botSlug/,
    );
  });
});

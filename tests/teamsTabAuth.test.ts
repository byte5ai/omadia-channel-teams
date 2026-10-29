import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { generateKeyPairSync, sign } from 'node:crypto';
import type { KeyObject } from 'node:crypto';

// byte5ai/omadia#1596 follow-up — the Hub / Tab-Config router is mounted
// `auth:'custom'`, so these tests ARE the route's authentication contract.
import {
  bearerToken,
  createTeamsSsoVerifier,
  createTeamsUiRouter,
  TeamsSsoError,
} from '@omadia/channel-teams';
import type { TeamsSsoVerifierOptions } from '@omadia/channel-teams';

const TENANT = 'd6314403-d0f2-4c08-ac17-0cc36bcabe7a';
const OTHER_TENANT = '11111111-2222-3333-4444-555555555555';
const APP_ID = '737c6ddd-6d4e-4599-8dc3-260281ea906e';
const KID = 'key-1';
const NOW_MS = Date.UTC(2026, 9, 29, 12, 0, 0);
const NOW_S = Math.floor(NOW_MS / 1000);

const { privateKey, publicKey } = generateKeyPairSync('rsa', {
  modulusLength: 2048,
});
const attacker = generateKeyPairSync('rsa', { modulusLength: 2048 });

const jwks = { keys: [{ ...publicKey.export({ format: 'jwk' }), kid: KID }] };

function b64(obj: unknown): string {
  return Buffer.from(JSON.stringify(obj)).toString('base64url');
}

function mint(
  claims: Record<string, unknown>,
  opts: { key?: KeyObject; header?: Record<string, unknown> } = {},
): string {
  const header = b64({ alg: 'RS256', typ: 'JWT', kid: KID, ...opts.header });
  const payload = b64(claims);
  const sig = sign(
    'RSA-SHA256',
    Buffer.from(`${header}.${payload}`),
    opts.key ?? privateKey,
  ).toString('base64url');
  return `${header}.${payload}.${sig}`;
}

const v1Claims = {
  aud: `api://odoo-bot-harness.fly.dev/${APP_ID}`,
  iss: `https://sts.windows.net/${TENANT}/`,
  tid: TENANT,
  oid: 'user-oid',
  scp: 'access_as_user',
  appid: '1fec8e78-bce4-4aaf-ab1b-5451cc387264',
  upn: 'someone@example.com',
  iat: NOW_S - 60,
  nbf: NOW_S - 60,
  exp: NOW_S + 3600,
};

function verifier(overrides: Partial<TeamsSsoVerifierOptions> = {}) {
  const fetched: string[] = [];
  const v = createTeamsSsoVerifier({
    audiences: [{ appId: APP_ID, tenantId: TENANT }],
    fetchJwks: async (tenantId) => {
      fetched.push(tenantId);
      return jwks;
    },
    now: () => NOW_MS,
    ...overrides,
  });
  return { v, fetched };
}

async function rejects(p: Promise<unknown>, pattern: RegExp): Promise<void> {
  await assert.rejects(p, (err: unknown) => {
    assert.ok(err instanceof TeamsSsoError, `expected TeamsSsoError, got ${String(err)}`);
    assert.match(err.message, pattern);
    return true;
  });
}

describe('createTeamsSsoVerifier', () => {
  it('accepts a v1 Teams SSO token (api://<host>/<appId> audience)', async () => {
    const { v } = verifier();
    const p = await v.verify(mint(v1Claims));
    assert.equal(p.tenantId, TENANT);
    assert.equal(p.objectId, 'user-oid');
    assert.equal(p.appId, APP_ID);
    assert.equal(p.username, 'someone@example.com');
  });

  it('accepts a v2 token (bare app id audience, v2 issuer)', async () => {
    const { v } = verifier();
    const p = await v.verify(
      mint({
        ...v1Claims,
        aud: APP_ID,
        iss: `https://login.microsoftonline.com/${TENANT}/v2.0`,
        appid: undefined,
        azp: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346',
      }),
    );
    assert.equal(p.appId, APP_ID);
  });

  it('refuses a token signed with a different key', async () => {
    const { v } = verifier();
    await rejects(v.verify(mint(v1Claims, { key: attacker.privateKey })), /signature/);
  });

  it('refuses a tampered payload', async () => {
    const { v } = verifier();
    const [h, , s] = mint(v1Claims).split('.');
    const forged = `${h}.${b64({ ...v1Claims, oid: 'admin' })}.${s}`;
    await rejects(v.verify(forged), /signature/);
  });

  it('refuses an unsigned (alg:none) token', async () => {
    const { v } = verifier();
    const t = `${b64({ alg: 'none', kid: KID })}.${b64(v1Claims)}.`;
    await rejects(v.verify(t), /algorithm/);
  });

  it('refuses a foreign tenant without fetching its keys', async () => {
    const { v, fetched } = verifier();
    await rejects(
      v.verify(mint({ ...v1Claims, tid: OTHER_TENANT, iss: `https://sts.windows.net/${OTHER_TENANT}/` })),
      /tenant/,
    );
    assert.deepEqual(fetched, []);
  });

  it('refuses a token issued for another app', async () => {
    const { v } = verifier();
    await rejects(
      v.verify(mint({ ...v1Claims, aud: 'api://evil.example/00000000-0000-0000-0000-000000000000' })),
      /audience/,
    );
  });

  it('refuses an issuer that does not match the tenant', async () => {
    const { v } = verifier();
    await rejects(
      v.verify(mint({ ...v1Claims, iss: `https://sts.windows.net/${OTHER_TENANT}/` })),
      /issuer/,
    );
  });

  it('refuses an expired token', async () => {
    const { v } = verifier();
    await rejects(v.verify(mint({ ...v1Claims, exp: NOW_S - 3600 })), /expired/);
  });

  it('refuses an app-only token (no scp) for the same audience', async () => {
    const { v } = verifier();
    const { scp: _scp, ...appOnly } = v1Claims;
    await rejects(v.verify(mint({ ...appOnly, roles: ['Something.All'] })), /delegated/);
  });

  it('refuses a token minted for a non-Teams client app', async () => {
    const { v } = verifier();
    await rejects(
      v.verify(mint({ ...v1Claims, appid: '99999999-9999-9999-9999-999999999999' })),
      /Teams client/,
    );
    const { appid: _appid, ...noClient } = v1Claims;
    await rejects(v.verify(mint(noClient)), /Teams client/);
  });

  it('backs off after a failed key fetch and shares one in-flight fetch', async () => {
    let calls = 0;
    let fail = true;
    const v = createTeamsSsoVerifier({
      audiences: [{ appId: APP_ID, tenantId: TENANT }],
      fetchJwks: async () => {
        calls += 1;
        await new Promise((r) => setTimeout(r, 10));
        if (fail) throw new Error('entra down');
        return jwks;
      },
      now: () => NOW_MS,
    });
    const token = mint(v1Claims);
    await Promise.all(
      [1, 2, 3].map(() => rejects(v.verify(token), /unavailable/)),
    );
    assert.equal(calls, 1, 'concurrent callers share one fetch');
    fail = false;
    await rejects(v.verify(token), /unavailable/);
    assert.equal(calls, 1, 'no refetch inside the failure backoff');
  });

  it('does not refetch keys for an unknown kid inside the floor', async () => {
    const { v, fetched } = verifier();
    await v.verify(mint(v1Claims));
    await rejects(v.verify(mint(v1Claims, { header: { kid: 'nope' } })), /signing key/);
    await rejects(v.verify(mint(v1Claims, { header: { kid: 'nope' } })), /signing key/);
    assert.equal(fetched.length, 1);
  });

  it('bearerToken parses only a well-formed Bearer header', () => {
    assert.equal(bearerToken('Bearer abc.def.ghi'), 'abc.def.ghi');
    assert.equal(bearerToken('bearer x'), 'x');
    assert.equal(bearerToken('Basic abc'), undefined);
    assert.equal(bearerToken(undefined), undefined);
  });
});

interface Dispatched {
  readonly status: number;
  readonly headers: ReadonlyMap<string, string>;
  readonly body: string;
}

/**
 * Drive the Router directly as the middleware function it is (same approach
 * as teamsMultiBotRouting.test.ts — bundling a listening express app into the
 * ESM test build trips over express' dynamic requires).
 */
function dispatch(
  router: unknown,
  url: string,
  headers: Record<string, string> = {},
): Promise<Dispatched> {
  return new Promise((resolve, reject) => {
    const outHeaders = new Map<string, string>();
    const req = { method: 'GET', url, headers, body: undefined };
    const finish = (body: string) =>
      resolve({ status: res.statusCode, headers: outHeaders, body });
    const res = {
      statusCode: 200,
      headersSent: false,
      writableEnded: false,
      setHeader: (k: string, v: string) => {
        outHeaders.set(k.toLowerCase(), v);
        return res;
      },
      getHeader: (k: string) => outHeaders.get(k.toLowerCase()),
      status: (code: number) => {
        res.statusCode = code;
        return res;
      },
      type: () => res,
      send: (body: string) => {
        finish(body);
        return res;
      },
      json: (body: unknown) => {
        finish(JSON.stringify(body));
        return res;
      },
      end: () => {
        finish('');
        return res;
      },
    };
    const next = (err?: unknown) => {
      if (err) reject(err instanceof Error ? err : new Error(String(err)));
      else resolve({ status: 404, headers: outHeaders, body: 'fallthrough' });
    };
    (router as (rq: unknown, rs: unknown, nx: unknown) => void)(req, res, next);
  });
}

describe('Teams UI router (auth:custom)', () => {
  const SECRET_TITLE = 'Abwesenheiten-Dashboard';
  const router: unknown = createTeamsUiRouter({
    verifier: verifier().v,
    discover: () => [
      { pluginId: '@omadia/agent-odoo-hr', routeId: 'absences', path: '/absences', title: SECRET_TITLE },
    ],
  });

  it('serves /hub and /tab-config as data-free shells', async () => {
    for (const page of ['/hub', '/tab-config']) {
      const res = await dispatch(router, page);
      assert.equal(res.status, 200);
      assert.ok(!res.body.includes(SECRET_TITLE), `${page} leaked catalogue data`);
      assert.ok(!res.body.includes('agent-odoo-hr'), `${page} leaked a plugin id`);
      assert.match(res.body, /\/p\/channel-teams\/catalog/);
      assert.match(res.headers.get('content-security-policy') ?? '', /frame-ancestors/);
    }
  });

  it('answers /catalog 401 without a token', async () => {
    const res = await dispatch(router, '/catalog');
    assert.equal(res.status, 401);
    assert.equal(res.headers.get('www-authenticate'), 'Bearer');
    assert.ok(!res.body.includes(SECRET_TITLE));
  });

  it('answers /catalog 401 with a forged token', async () => {
    const res = await dispatch(router, '/catalog', {
      authorization: `Bearer ${mint(v1Claims, { key: attacker.privateKey })}`,
    });
    assert.equal(res.status, 401);
    assert.ok(!res.body.includes(SECRET_TITLE));
  });

  it('serves /catalog to a valid Teams SSO token', async () => {
    const res = await dispatch(router, '/catalog', {
      authorization: `Bearer ${mint(v1Claims)}`,
    });
    assert.equal(res.status, 200);
    assert.equal(res.headers.get('cache-control'), 'no-store');
    const body = JSON.parse(res.body) as { routes: Array<{ title: string; href: string }> };
    assert.equal(body.routes[0]?.title, SECRET_TITLE);
    assert.equal(body.routes[0]?.href, '/p/agent-odoo-hr/absences');
  });
});

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

// byte5ai/omadia#1620 — the Hub / Tab-Config shells are mounted
// `auth:'public'`, the catalogue `auth:'teams-sso'` (the KERNEL verifies the
// Teams SSO token). These tests pin the plugin's half of that contract: the
// shells carry no data, and the catalogue fails closed without the kernel's
// `req.teamsSsoUser`.
import {
  createTeamsCatalogRouter,
  createTeamsUiShellRouter,
} from '@omadia/channel-teams';

const pkgRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

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
  extra: Record<string, unknown> = {},
): Promise<Dispatched> {
  return new Promise((resolve, reject) => {
    const outHeaders = new Map<string, string>();
    const req = { method: 'GET', url, headers: {}, body: undefined, ...extra };
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

const SECRET_TITLE = 'Abwesenheiten-Dashboard';
const SSO_USER = {
  oid: 'user-oid',
  tenantId: 'd6314403-d0f2-4c08-ac17-0cc36bcabe7a',
  appId: '737c6ddd-6d4e-4599-8dc3-260281ea906e',
};

describe('Teams UI shell router (auth:public)', () => {
  const router: unknown = createTeamsUiShellRouter({
    catalogUrl: '/p/channel-teams/catalog',
  });

  it('serves /hub and /tab-config as data-free shells that fetch the catalogue', async () => {
    for (const page of ['/hub', '/tab-config']) {
      const res = await dispatch(router, page);
      assert.equal(res.status, 200);
      assert.ok(!res.body.includes(SECRET_TITLE), `${page} leaked catalogue data`);
      assert.ok(!res.body.includes('agent-odoo-hr'), `${page} leaked a plugin id`);
      assert.match(res.body, /\/p\/channel-teams\/catalog/);
      assert.match(res.body, /getAuthToken/);
      assert.match(res.body, /credentials: 'omit'/);
      assert.match(res.headers.get('content-security-policy') ?? '', /frame-ancestors/);
    }
  });

  it('does not serve the catalogue itself', async () => {
    const res = await dispatch(router, '/catalog');
    assert.equal(res.status, 404);
    assert.equal(res.body, 'fallthrough');
  });
});

describe('Teams catalogue router (auth:teams-sso)', () => {
  let discoverCalls = 0;
  const router: unknown = createTeamsCatalogRouter({
    discover: () => {
      discoverCalls += 1;
      return [
        { pluginId: '@omadia/agent-odoo-hr', routeId: 'absences', path: '/teams/absences', title: SECRET_TITLE },
      ];
    },
  });

  it('answers 401 auth.teams_sso_required without req.teamsSsoUser', async () => {
    const before = discoverCalls;
    const res = await dispatch(router, '/');
    assert.equal(res.status, 401);
    assert.deepEqual(JSON.parse(res.body), { code: 'auth.teams_sso_required' });
    assert.equal(res.headers.get('cache-control'), 'no-store');
    assert.ok(!res.body.includes(SECRET_TITLE));
    assert.equal(discoverCalls, before, 'catalogue was read for an unauthenticated request');
  });

  it('ignores a bearer header — only the kernel-set user opens it', async () => {
    const res = await dispatch(router, '/', {
      headers: { authorization: 'Bearer forged.token.value' },
    });
    assert.equal(res.status, 401);
    assert.ok(!res.body.includes(SECRET_TITLE));
  });

  it('serves the catalogue to a kernel-verified Teams SSO user', async () => {
    const res = await dispatch(router, '/', { teamsSsoUser: SSO_USER });
    assert.equal(res.status, 200);
    assert.equal(res.headers.get('cache-control'), 'no-store');
    const body = JSON.parse(res.body) as { routes: Array<{ title: string; href: string }> };
    assert.equal(body.routes[0]?.title, SECRET_TITLE);
    // Plugin descriptors point at their Teams shell (`/teams/<page>`).
    assert.equal(body.routes[0]?.href, '/p/agent-odoo-hr/teams/absences');
  });
});

describe('manifest — Teams tab routes (#1620)', () => {
  const manifestYaml = readFileSync(join(pkgRoot, 'manifest.yaml'), 'utf8');

  it('declares the public path the shell and catalogue mounts sit beneath', () => {
    assert.match(manifestYaml, /^ {2}public_paths:\n {4}- "\/p\/channel-teams"$/m);
  });

  it('requires a core whose plugin-api knows auth:teams-sso', () => {
    assert.match(manifestYaml, /^ {2}core: ">=1\.51\.0 <2\.0"$/m);
  });
});

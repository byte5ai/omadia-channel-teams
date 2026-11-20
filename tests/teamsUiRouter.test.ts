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
  TAB_CONFIG_HUB_VALUE,
  tabConfigChoices,
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

  it('requires a core whose plugin-api knows auth:teams-sso (1.51.0; the S6 channel SDK this plugin imports, 1.75.0, is above it)', () => {
    assert.match(manifestYaml, /^ {2}core: ">=1\.75\.0 <2\.0"$/m);
  });
});

/**
 * Bug seen live (2026-11-18): adding an agent app to a channel from the Teams
 * app dialog opens this page; for an agent without plugin surfaces the picker
 * was empty, `setValidityState(true)` was never called, Save stayed disabled
 * and the install looked failed although it had succeeded.
 *
 * These tests run the page's own script against a minimal DOM and Teams SDK.
 */
interface FakeElement {
  id: string;
  textContent: string;
  value: string;
  options: FakeElement[];
  selectedIndex: number;
  text: string;
  attributes: Record<string, string>;
  listeners: Record<string, () => void>;
  src?: string;
  onload?: () => void;
  appendChild(child: FakeElement): void;
  setAttribute(name: string, value: string): void;
  getAttribute(name: string): string | null;
  addEventListener(name: string, fn: () => void): void;
}

function fakeElement(id = ''): FakeElement {
  const el: FakeElement = {
    id,
    textContent: '',
    value: '',
    options: [],
    selectedIndex: 0,
    text: '',
    attributes: {},
    listeners: {},
    appendChild(child) {
      // <select>: the first option becomes the selection, as in a browser.
      el.options.push(child);
      if (el.options.length === 1) el.value = child.value;
      child.text = child.textContent;
      // <head>: a script element "loads" at once.
      if (child.onload) queueMicrotask(child.onload);
    },
    setAttribute(name, value) {
      el.attributes[name] = value;
    },
    getAttribute(name) {
      return el.attributes[name] ?? null;
    },
    addEventListener(name, fn) {
      el.listeners[name] = fn;
    },
  };
  return el;
}

async function runTabConfigPage(catalog: { status: number; routes?: unknown[] }): Promise<{
  validity: boolean[];
  status: string;
  options: { value: string; title: string }[];
  save(): Promise<{ config?: Record<string, string>; result: string }>;
}> {
  const router: unknown = createTeamsUiShellRouter({ catalogUrl: '/p/channel-teams/catalog', hubUrl: '/p/channel-teams/hub' });
  const page = await dispatch(router, '/tab-config');
  const script = /<script>([\s\S]*?)<\/script>/.exec(page.body)?.[1];
  assert.ok(script, 'the page carries its script inline');

  const elements: Record<string, FakeElement> = {
    'tab-config-select': fakeElement('tab-config-select'),
    'omadia-status': fakeElement('omadia-status'),
  };
  const validity: boolean[] = [];
  let saveHandler: ((event: { notifySuccess(): void; notifyFailure(reason?: string): void }) => void) | undefined;
  let savedConfig: Record<string, string> | undefined;
  const teams = {
    app: { initialize: async () => {} },
    authentication: { getAuthToken: async () => 'sso-token' },
    pages: {
      config: {
        registerOnSaveHandler: (handler: typeof saveHandler) => {
          saveHandler = handler;
        },
        setValidityState: (state: boolean) => {
          validity.push(state);
        },
        setConfig: async (config: Record<string, string>) => {
          savedConfig = config;
        },
      },
    },
  };
  const windowStub = { microsoftTeams: teams, location: { origin: 'https://harness.example' } };
  const documentStub = {
    readyState: 'complete',
    head: fakeElement('head'),
    getElementById: (id: string) => elements[id],
    createElement: () => fakeElement(),
    addEventListener: () => {},
  };
  const fetchStub = async () => ({
    ok: catalog.status === 200,
    status: catalog.status,
    json: async () => ({ routes: catalog.routes ?? [] }),
  });
  // eslint-disable-next-line @typescript-eslint/no-implied-eval
  new Function('window', 'document', 'fetch', 'console', 'setTimeout', 'URL', script)(
    windowStub,
    documentStub,
    fetchStub,
    { log: () => {} },
    () => 0,
    URL,
  );
  for (let i = 0; i < 20; i += 1) await new Promise((resolve) => setImmediate(resolve));

  const select = elements['tab-config-select'] as FakeElement;
  return {
    validity,
    status: (elements['omadia-status'] as FakeElement).textContent,
    options: select.options.map((o) => ({ value: o.value, title: o.textContent })),
    save: async () => {
      let result = 'no-handler';
      saveHandler?.({
        notifySuccess: () => {
          result = 'success';
        },
        notifyFailure: () => {
          result = 'failure';
        },
      });
      for (let i = 0; i < 5; i += 1) await new Promise((resolve) => setImmediate(resolve));
      return { ...(savedConfig ? { config: savedConfig } : {}), result };
    },
  };
}

describe('Teams tab configuration page enables Save', () => {
  it('for an agent WITHOUT plugin surfaces: the overview is offered and Save is enabled', async () => {
    const page = await runTabConfigPage({ status: 200, routes: [] });

    assert.deepEqual(page.options, [{ value: TAB_CONFIG_HUB_VALUE, title: 'Plugin overview' }]);
    assert.deepEqual(page.validity, [true], 'setValidityState(true) is what enables Save in Teams');
    assert.match(page.status, /no plugin surfaces yet/);

    const saved = await page.save();
    assert.equal(saved.result, 'success');
    assert.equal(saved.config?.['contentUrl'], 'https://harness.example/p/channel-teams/hub');
    assert.equal(saved.config?.['entityId'], 'channel-teams:/hub');
    assert.equal(saved.config?.['suggestedDisplayName'], 'Plugin overview');
  });

  it('for an agent with plugin surfaces: they come first, the overview last', async () => {
    const page = await runTabConfigPage({
      status: 200,
      routes: [{ pluginId: '@omadia/agent-odoo-hr', path: '/teams/absences', title: 'Absences', href: '/p/agent-odoo-hr/teams/absences' }],
    });

    assert.deepEqual(page.options.map((o) => o.value), ['@omadia/agent-odoo-hr::/teams/absences', TAB_CONFIG_HUB_VALUE]);
    assert.deepEqual(page.validity, [true]);
    assert.equal(page.status, '');
    const saved = await page.save();
    assert.equal(saved.config?.['contentUrl'], 'https://harness.example/p/agent-odoo-hr/teams/absences');
  });

  it('says on the page itself that a tab is optional and the agent is already added', async () => {
    const router: unknown = createTeamsUiShellRouter({});
    const res = await dispatch(router, '/tab-config');
    assert.match(res.body, /A tab is optional\./);
    assert.match(res.body, /already added/);
    assert.match(res.body, /close this dialog without saving/);
  });

  it('keeps Save disabled and shows the reason when the catalogue cannot be read', async () => {
    const page = await runTabConfigPage({ status: 401 });
    assert.deepEqual(page.validity, []);
    assert.match(page.status, /Not authorised \(401\)/);
  });

  it('tabConfigChoices never returns an empty list and skips unusable routes', () => {
    assert.deepEqual(tabConfigChoices([], '/hub'), [{ value: TAB_CONFIG_HUB_VALUE, title: 'Plugin overview', href: '/hub' }]);
    const choices = tabConfigChoices(
      [{ pluginId: 'p', href: '' }, { href: '/x' }, { pluginId: 'q', href: '/q' }],
      '/hub',
    );
    assert.deepEqual(choices.map((c) => [c.value, c.title]), [
      ['q::/', 'q'],
      [TAB_CONFIG_HUB_VALUE, 'Plugin overview'],
    ]);
  });
});


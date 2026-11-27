import { Router } from 'express';
import {
  html,
  htmlDoc,
  renderRoute,
  safe,
  withIframeSafeHeaders,
} from '@omadia/plugin-ui-helpers';


/**
 * Discoverable descriptor of a plugin-served UI surface. Routes for the
 * configurable Teams Tab to target. In Phase 1 a static list is supplied
 * via opts.discover(); a future iteration will query the kernel's
 * PluginRouteRegistry for live discovery.
 */
export interface DiscoveredUiRoute {
  readonly pluginId: string;
  readonly routeId: string;
  /** Path relative to the plugin mount, e.g. `/dashboard`. */
  readonly path: string;
  readonly title: string;
}

export interface TeamsUiShellRouterOptions {
  /** Absolute URL path the pages fetch the catalogue from. Default
   *  `/p/channel-teams/catalog`. */
  readonly catalogUrl?: string;
  /** Absolute URL path of the hub page, the tab that is always available.
   *  Default `/p/channel-teams/hub`. */
  readonly hubUrl?: string;
}

/** One entry of the tab picker on the configuration page. */
export interface TabConfigChoice {
  /** `<pluginId>::<path>`; also the tab's entity id source. */
  value: string;
  title: string;
  href: string;
}

/** `value` of the choice that pins the plugin overview. */
export const TAB_CONFIG_HUB_VALUE = 'channel-teams::/hub';

/**
 * The choices of the tab picker: every plugin surface of the catalogue, then
 * ALWAYS the plugin overview.
 *
 * The list must never be empty. Teams opens this page when somebody adds the
 * app to a channel, and its Save button stays disabled until the page calls
 * `setValidityState(true)`. An agent without plugin surfaces used to produce
 * an empty picker, so Save never enabled and the dialog looked like a failed
 * install although the agent was added. The overview is a tab every agent
 * has, so there is always something valid to save.
 *
 * Serialised into the page with `toString()`, so it must stay
 * self-contained: no imports, no closures.
 */
export function tabConfigChoices(
  routes: readonly { pluginId?: unknown; path?: unknown; title?: unknown; href?: unknown }[],
  hubHref: string,
): TabConfigChoice[] {
  const choices: TabConfigChoice[] = [];
  for (const route of routes) {
    if (typeof route.pluginId !== 'string' || typeof route.href !== 'string' || route.href.length === 0) continue;
    const path = typeof route.path === 'string' && route.path.length > 0 ? route.path : '/';
    choices.push({
      value: route.pluginId + '::' + path,
      title: typeof route.title === 'string' && route.title.length > 0 ? route.title : route.pluginId,
      href: route.href,
    });
  }
  choices.push({ value: 'channel-teams::/hub', title: 'Plugin overview', href: hubHref });
  return choices;
}

export interface TeamsCatalogRouterOptions {
  /** Callback that returns the current list of plugin uiRoutes. Allows
   *  the channel plugin to inject a discovery strategy without coupling
   *  to the kernel's route-registry type. */
  readonly discover: () => DiscoveredUiRoute[] | Promise<DiscoveredUiRoute[]>;
  /** Origin (scheme + host) that web-ui presents to the browser; this is
   *  the URL Teams iframes resolve against. Pass via env in production;
   *  default `''` produces relative URLs which the iframe origin resolves. */
  readonly webUiOrigin?: string;
}

/** Shape the kernel's `auth:'teams-sso'` mode leaves on the request
 *  (`TeamsSsoUser` in `@omadia/plugin-api` >= 1.51.0). Declared locally so
 *  the guard below compiles against older plugin-api typings too. */
interface TeamsSsoRequest {
  teamsSsoUser?: unknown;
}

/** Teams JS SDK, loaded by both pages — Teams talks to it via postMessage. */
const TEAMS_JS_URL =
  'https://res.cdn.office.net/teams-js/2.34.0/js/MicrosoftTeams.min.js';

/**
 * Client bootstrap shared by both pages: load teams-js, initialise, fetch an
 * SSO token, call `/catalog` with it. Resolves with the route list; rejects
 * with a human-readable reason (never with data).
 */
function catalogLoaderJs(catalogUrl: string): string {
  return `
    var omadiaLoadCatalog = function () {
      var sdk = new Promise(function (resolve, reject) {
        var s = document.createElement('script');
        s.src = ${JSON.stringify(TEAMS_JS_URL)};
        s.onload = function () { resolve(window.microsoftTeams); };
        s.onerror = function () { reject(new Error('Teams SDK failed to load')); };
        document.head.appendChild(s);
      });
      var outsideTeams = new Promise(function (_resolve, reject) {
        setTimeout(function () {
          reject(new Error('Open this page inside Microsoft Teams.'));
        }, 8000);
      });
      return Promise.race([
        sdk.then(function (t) {
          return t.app.initialize().then(function () { return t; });
        }),
        outsideTeams,
      ]).then(function (t) {
        return t.authentication.getAuthToken().then(function (token) {
          return fetch(${JSON.stringify(catalogUrl)}, {
            headers: { Authorization: 'Bearer ' + token },
            credentials: 'omit',
            cache: 'no-store',
          });
        }, function (err) {
          throw new Error('Teams sign-in failed: ' + ((err && err.message) || err));
        }).then(function (res) {
          if (!res.ok) throw new Error('Not authorised (' + res.status + ').');
          return res.json().then(function (body) {
            return { teams: t, routes: (body && body.routes) || [] };
          });
        });
      });
    };
    var omadiaShowError = function (err) {
      var box = document.getElementById('omadia-status');
      if (box) box.textContent = (err && err.message) || String(err);
    };
  `;
}

/**
 * Teams-bridge uiRoutes exposed by the channel-teams plugin.
 *
 * Shell router — `auth:'public'` at `/p/channel-teams` (#1596 / #1620):
 *
 *   GET /hub        — Personal-Tab content. Lists discovered plugin
 *                     uiRoutes as clickable cards.
 *
 *   GET /tab-config — configurationUrl for the Teams Configurable Tab,
 *                     iframed inside Teams' Tab-add modal; calls
 *                     `pages.config.setConfig({ contentUrl, entityId })`
 *                     when the user picks a target route + clicks Save.
 *
 * Teams iframes these pages without the omadia session cookie and without a
 * bearer, so they are DATA-FREE shells: this factory has no access to the
 * catalogue at all. They fetch it from {@link createTeamsCatalogRouter}.
 */
export function createTeamsUiShellRouter(
  opts: TeamsUiShellRouterOptions = {},
): Router {
  const router = Router();
  const catalogUrl = opts.catalogUrl ?? '/p/channel-teams/catalog';
  const hubUrl = opts.hubUrl ?? '/p/channel-teams/hub';

  router.get(
    '/hub',
    renderRoute(() => {
      const hubJs = `
        (function () {
          ${catalogLoaderJs(catalogUrl)}
          var render = function (result) {
            var status = document.getElementById('omadia-status');
            var list = document.getElementById('omadia-hub-list');
            if (result.routes.length === 0) {
              status.textContent = 'No plugin UIs registered yet.';
              return;
            }
            status.textContent = '';
            result.routes.forEach(function (r) {
              var li = document.createElement('li');
              var a = document.createElement('a');
              a.href = r.href;
              a.target = '_blank';
              a.rel = 'noopener';
              a.className = 'block bg-white border border-slate-200 rounded-lg p-4 hover:border-slate-400 hover:shadow-sm transition';
              var id = document.createElement('div');
              id.className = 'text-xs uppercase text-slate-500';
              id.textContent = r.pluginId;
              var title = document.createElement('div');
              title.className = 'text-base font-medium mt-1';
              title.textContent = r.title;
              var path = document.createElement('div');
              path.className = 'text-xs font-mono text-slate-400 mt-2 truncate';
              path.textContent = r.href;
              a.appendChild(id);
              a.appendChild(title);
              a.appendChild(path);
              li.appendChild(a);
              list.appendChild(li);
            });
          };
          var start = function () { omadiaLoadCatalog().then(render, omadiaShowError); };
          if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', start);
          } else {
            start();
          }
        })();
      `;
      return htmlDoc({
        title: 'Omadia — Plugin Hub',
        // Self-filling: refresh every minute so newly-installed plugin
        // uiRoutes appear in the hub without a manual reload.
        refreshSeconds: 60,
        body: html`
          <main class="max-w-3xl mx-auto p-6 space-y-6">
            <header>
              <h1 class="text-2xl font-semibold tracking-tight">Plugin Hub</h1>
              <p class="text-sm text-slate-500">
                Pin any of these surfaces as a Teams Tab, or open them
                directly here.
              </p>
              <p
                class="mt-1 text-[10px] uppercase tracking-wider text-slate-400"
                data-testid="hub-build"
              >
                channel-teams build 0.5.0
              </p>
            </header>
            <p id="omadia-status" class="text-sm text-slate-500">Loading…</p>
            <ul
              id="omadia-hub-list"
              class="grid grid-cols-2 gap-3"
              data-testid="hub-list"
            ></ul>
            <script>
              ${safe(hubJs)}
            </script>
          </main>
        `,
      });
    }),
  );

  // byte5ai/omadia#2312 (omadia ADR-0301) — the Teams dialog behind the
  // "Connect" button of an answer that needed a per-user MCP connection. It
  // holds no data and no secret: the kernel starts and completes the flow
  // only with this person's Teams SSO token, so the page is a public shell.
  router.get(
    '/connect',
    renderRoute(() => {
      const connectJs = `
        (function () {
          var de = (navigator.language || '').toLowerCase().indexOf('de') === 0;
          var T = de ? {
            working: 'Verbindung wird vorbereitet…', popup: 'Melde dich im geöffneten Fenster beim Anbieter an.',
            done: 'Verbunden. Du kannst dieses Fenster schließen und deine Frage wiederholen.',
            notPerson: 'Dein Konto ist nicht im Verzeichnis dieser Instanz. Bitte wende dich an den Betrieb.',
            failed: 'Die Verbindung ist fehlgeschlagen: ', outside: 'Öffne diese Seite in Microsoft Teams.'
          } : {
            working: 'Preparing the connection…', popup: 'Sign in to the provider in the window that opened.',
            done: 'Connected. You can close this window and ask again.',
            notPerson: 'Your account is not in this instance\'s directory. Please contact your operator.',
            failed: 'The connection failed: ', outside: 'Open this page inside Microsoft Teams.'
          };
          var status = function (m) { document.getElementById('omadia-status').textContent = m; };
          var q = new URLSearchParams(location.search);
          var server = q.get('server') || '';
          var name = q.get('name') || server;
          document.getElementById('omadia-server').textContent = name;
          var api = '/api/v1/directory/connect/mcp/';
          var post = function (t, path, body) {
            return t.authentication.getAuthToken().then(function (token) {
              return fetch(api + path, {
                method: 'POST', credentials: 'omit', cache: 'no-store',
                headers: { Authorization: 'Bearer ' + token, 'content-type': 'application/json' },
                body: JSON.stringify(body || {})
              });
            }).then(function (res) {
              return res.json().catch(function () { return {}; }).then(function (b) {
                if (!res.ok) { var e = new Error((b && b.code) || String(res.status)); e.code = b && b.code; throw e; }
                return b;
              });
            });
          };
          var sdk = new Promise(function (resolve, reject) {
            var s = document.createElement('script');
            s.src = ${JSON.stringify(TEAMS_JS_URL)};
            s.onload = function () { resolve(window.microsoftTeams); };
            s.onerror = function () { reject(new Error(T.outside)); };
            document.head.appendChild(s);
          });
          status(T.working);
          sdk.then(function (t) { return t.app.initialize().then(function () { return t; }); })
            .then(function (t) {
              if (!server) throw new Error('server');
              return post(t, encodeURIComponent(server) + '/start').then(function (b) {
                status(T.popup);
                return t.authentication.authenticate({ url: location.origin + b.startUrl, width: 600, height: 700 });
              }).then(function (result) {
                var r = JSON.parse(result);
                return post(t, 'complete', { state: r.state, code: r.code, iss: r.iss });
              });
            })
            .then(function () { status(T.done); }, function (err) {
              status(err && err.code === 'directory.not_a_person' ? T.notPerson : T.failed + ((err && err.message) || err));
            });
        })();
      `;
      return htmlDoc({
        title: 'Omadia — Connect',
        body: html`
          <main class="max-w-md mx-auto p-6 space-y-4">
            <h1 class="text-xl font-semibold tracking-tight" id="omadia-server"></h1>
            <p id="omadia-status" class="text-sm text-slate-600"></p>
            <script>
              ${safe(connectJs)}
            </script>
          </main>
        `,
      });
    }),
  );

  router.get(
    '/tab-config',
    renderRoute(() => {
      const tabConfigJs = `
        (function () {
          ${catalogLoaderJs(catalogUrl)}
          var log = function (m) { try { console.log('[tab-config]', m); } catch (e) {} };
          var tabConfigChoices = ${tabConfigChoices.toString()};
          var start = function () {
            var select = document.getElementById('tab-config-select');
            omadiaLoadCatalog().then(function (result) {
              var t = result.teams;
              var choices = tabConfigChoices(result.routes, ${JSON.stringify(hubUrl)});
              choices.forEach(function (c) {
                var opt = document.createElement('option');
                opt.value = c.value;
                opt.textContent = c.title;
                opt.setAttribute('data-href', c.href);
                select.appendChild(opt);
              });
              // Only the overview is on offer: say so instead of showing a
              // lone entry without explanation.
              document.getElementById('omadia-status').textContent =
                choices.length === 1
                  ? 'This agent has no plugin surfaces yet. You can pin the plugin overview, or close this dialog.'
                  : '';
              t.pages.config.registerOnSaveHandler(function (saveEvent) {
                var raw = select.value || '';
                var parts = raw.split('::');
                var pluginId = parts[0];
                var routePath = parts[1] || '/';
                var chosen = select.options[select.selectedIndex];
                var href = chosen.getAttribute('data-href') || '';
                var contentUrl = new URL(href, window.location.origin).toString();
                log('save → ' + contentUrl);
                t.pages.config.setConfig({
                  contentUrl: contentUrl,
                  websiteUrl: contentUrl,
                  entityId: pluginId + ':' + routePath,
                  suggestedDisplayName: chosen.text,
                }).then(function () {
                  saveEvent.notifySuccess();
                }).catch(function (err) {
                  log('setConfig error: ' + (err && err.message));
                  saveEvent.notifyFailure(err && err.message);
                });
              });
              // The picker is never empty (see tabConfigChoices), so Save is
              // enabled as soon as the page has loaded.
              t.pages.config.setValidityState(Boolean(select.value));
              select.addEventListener('change', function () {
                t.pages.config.setValidityState(Boolean(select.value));
              });
            }, omadiaShowError);
          };
          if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', start);
          } else {
            start();
          }
        })();
      `;

      return htmlDoc({
        title: 'Omadia — Configure Tab',
        body: html`
          <main class="max-w-md mx-auto p-6 space-y-4">
            <header>
              <h1 class="text-xl font-semibold">Configure Tab</h1>
              <p class="text-sm text-slate-500">
                Pick which plugin surface this Tab should show.
              </p>
              <p class="mt-2 text-sm text-slate-700" data-testid="tab-config-optional">
                A tab is optional. The agent is already added and answers in
                this team or chat; you can close this dialog without saving.
              </p>
            </header>
            <p id="omadia-status" class="text-sm text-slate-500">Loading…</p>
            <form id="tab-config-form" class="space-y-3">
              <label class="block text-sm font-medium" for="tab-config-select">
                Target route
              </label>
              <select
                id="tab-config-select"
                class="w-full border border-slate-300 rounded-md p-2 bg-white"
              ></select>
              <p class="text-xs text-slate-400">
                Click <strong>Save</strong> in Teams to pin this surface.
              </p>
            </form>
            <script>
              ${safe(tabConfigJs)}
            </script>
          </main>
        `,
      });
    }),
  );

  return router;
}

/**
 * `GET /` — the route list both shell pages render, as JSON. The ONLY
 * data-bearing endpoint of the Teams bridge. Mounted by plugin.ts with the
 * kernel's `auth:'teams-sso'` mode (#1620) at `/p/channel-teams/catalog`:
 * the kernel verifies the Teams SSO token and sets `req.teamsSsoUser`
 * before this router runs.
 *
 * The first middleware is a fail-closed guard: without `req.teamsSsoUser`
 * (an older core that does not know `'teams-sso'`, or any mis-mount) it
 * answers the kernel's own 401 and never reaches the catalogue.
 */
export function createTeamsCatalogRouter(opts: TeamsCatalogRouterOptions): Router {
  const router = Router();
  const origin = opts.webUiOrigin ?? '';

  /**
   * pluginId in the catalog is the full agentId (e.g. `@omadia/foo`),
   * but plugins mount their routers under `/p/<name-without-namespace>`
   * (e.g. `/p/foo`). Strip the leading `@<scope>/` so the rendered URL
   * matches the actual HTTP mount. Plugins on the #1620 Teams-tab pattern
   * point their descriptor `path` at `/teams/<page>`, so the href lands on
   * their data-free shell.
   */
  const mountSlug = (pluginId: string): string =>
    pluginId.replace(/^@[^/]+\//, '');

  router.use((req, res, next) => {
    withIframeSafeHeaders(res);
    res.setHeader('Cache-Control', 'no-store');
    if (!(req as TeamsSsoRequest).teamsSsoUser) {
      res.status(401).json({ code: 'auth.teams_sso_required' });
      return;
    }
    next();
  });

  router.get('/', async (_req, res, next) => {
    try {
      const routes = await opts.discover();
      res.json({
        ok: true,
        routes: routes.map((r) => ({
          pluginId: r.pluginId,
          path: r.path,
          title: r.title,
          href: `${origin}/p/${mountSlug(r.pluginId)}${r.path}`,
        })),
      });
    } catch (err) {
      next(err);
    }
  });

  return router;
}

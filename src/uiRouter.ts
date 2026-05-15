import { Router } from 'express';
import { html, htmlDoc, renderRoute, safe } from '@omadia/plugin-ui-helpers';

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

export interface TeamsUiRouterOptions {
  /** Callback that returns the current list of plugin uiRoutes. Allows
   *  the channel plugin to inject a discovery strategy without coupling
   *  to the kernel's route-registry type. */
  readonly discover: () => DiscoveredUiRoute[] | Promise<DiscoveredUiRoute[]>;
  /** Origin (scheme + host) that web-ui presents to the browser; this is
   *  the URL Teams iframes resolve against. Pass via env in production;
   *  default `''` produces relative URLs which the iframe origin resolves. */
  readonly webUiOrigin?: string;
}

/**
 * Teams-bridge uiRoutes exposed by the channel-teams plugin. Two surfaces:
 *
 *   GET /hub        — Personal-Tab content. Lists discovered plugin
 *                     uiRoutes as clickable cards. User can pin individual
 *                     routes by opening them.
 *
 *   GET /tab-config — configurationUrl for the Teams Configurable Tab. The
 *                     page is iframed inside Teams' Tab-add modal and uses
 *                     `@microsoft/teams-js` v2 to call
 *                     `pages.config.setConfig({ contentUrl, entityId })`
 *                     when the user picks a target route + clicks Save.
 *
 * Both pages are SSR HTML — no React, no client-side bundle from us. The
 * Tab-Config page does load the Teams JS SDK via CDN because the Teams
 * host needs to communicate with it via postMessage.
 */
export function createTeamsUiRouter(opts: TeamsUiRouterOptions): Router {
  const router = Router();
  const origin = opts.webUiOrigin ?? '';

  router.get(
    '/hub',
    renderRoute(async () => {
      const routes = await opts.discover();
      return htmlDoc({
        title: 'Omadia — Plugin Hub',
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
                channel-teams build 0.2.1
              </p>
            </header>

            ${routes.length === 0
              ? safe(
                  '<p class="text-sm text-slate-500">No plugin UIs registered yet.</p>',
                )
              : html`
                  <ul class="grid grid-cols-2 gap-3" data-testid="hub-list">
                    ${routes.map(
                      (r) => html`
                        <li>
                          <a
                            href="${origin}/p/${r.pluginId}${r.path}"
                            target="_blank"
                            rel="noopener"
                            class="block bg-white border border-slate-200 rounded-lg p-4 hover:border-slate-400 hover:shadow-sm transition"
                          >
                            <div class="text-xs uppercase text-slate-500">
                              ${r.pluginId}
                            </div>
                            <div class="text-base font-medium mt-1">
                              ${r.title}
                            </div>
                            <div
                              class="text-xs font-mono text-slate-400 mt-2 truncate"
                            >
                              ${r.path}
                            </div>
                          </a>
                        </li>
                      `,
                    )}
                  </ul>
                `}
          </main>
        `,
      });
    }),
  );

  router.get(
    '/tab-config',
    renderRoute(async () => {
      const routes = await opts.discover();
      const optionsHtml = routes.map(
        (r) => html`
          <option value="${r.pluginId}::${r.path}">
            ${r.pluginId} — ${r.title}
          </option>
        `,
      );

      const teamsBootstrapJs = `
        (function () {
          if (typeof window === 'undefined') return;
          var sdkPromise = new Promise(function (resolve, reject) {
            var s = document.createElement('script');
            s.src = 'https://res.cdn.office.net/teams-js/2.34.0/js/MicrosoftTeams.min.js';
            s.onload = function () { resolve(window.microsoftTeams); };
            s.onerror = function () { reject(new Error('teams-js failed to load')); };
            document.head.appendChild(s);
          });
          document.addEventListener('DOMContentLoaded', function () {
            var form = document.getElementById('tab-config-form');
            var select = document.getElementById('tab-config-select');
            sdkPromise.then(function (t) {
              t.app.initialize();
              t.pages.config.registerOnSaveHandler(function (saveEvent) {
                var raw = select.value || '';
                var parts = raw.split('::');
                var pluginId = parts[0];
                var routePath = parts[1] || '/';
                var origin = ${JSON.stringify(origin || '')};
                var contentUrl =
                  (origin || window.location.origin) +
                  '/p/' + pluginId + routePath;
                t.pages.config.setConfig({
                  contentUrl: contentUrl,
                  entityId: pluginId + ':' + routePath,
                  suggestedDisplayName:
                    select.options[select.selectedIndex].text,
                }).then(function () {
                  saveEvent.notifySuccess();
                });
              });
              t.pages.config.setValidityState(true);
            }).catch(function (err) {
              // Outside of Teams (e.g. PoC smoke), the SDK won't load; the
              // form is still usable as a preview but Save is inert.
              console.warn('teams-js unavailable:', err && err.message);
            });
          });
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
            </header>
            <form id="tab-config-form" class="space-y-3">
              <label class="block text-sm font-medium" for="tab-config-select">
                Target route
              </label>
              <select
                id="tab-config-select"
                class="w-full border border-slate-300 rounded-md p-2 bg-white"
              >
                ${optionsHtml}
              </select>
              <p class="text-xs text-slate-400">
                Click <strong>Save</strong> in Teams to pin this surface.
              </p>
            </form>
            <script>
              ${safe(teamsBootstrapJs)}
            </script>
          </main>
        `,
      });
    }),
  );

  return router;
}

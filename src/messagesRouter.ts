import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
} from 'botbuilder';
import type {
  ConversationReference,
  Activity,
  TurnContext,
} from 'botbuilder';
import { Router } from 'express';
import type { Request, Response } from 'express';
import type { TeamsBot } from './teamsBot.js';
import type { TeamsBotAppType, TeamsBotIdentity } from './teamsBotIdentity.js';
import { findTeamsBotBySlug, parseTeamsBotKey, teamsBotLogLabel } from './teamsBotIdentity.js';
import { normalizeTeamsBotAppId } from './teamsConversationRefStore.js';

/**
 * #860 W0a — multi-bot: one Bot-Framework credential set per bot. Each
 * configured {@link TeamsBotIdentity} gets its OWN
 * `ConfigurationServiceClientCredentialFactory` + `CloudAdapter`, so a turn
 * (or proactive send) for bot B can never authenticate with bot A's app
 * password. Adapters are long-lived — built once here and shared with
 * proactive-send callers so the BF authentication factory + token cache are
 * reused, never rebuilt per request.
 */

/** One bot's resolved credential set. `appPassword` is the SECRET VALUE the
 *  config-wiring layer resolved from `identity.appPasswordSecretRef` via
 *  `ctx.secrets.require(...)` — it exists only to be handed to the BF
 *  credential factory. Never log it, never persist it. */
export interface TeamsRouterBotCredentials {
  identity: TeamsBotIdentity;
  appPassword: string;
}

/**
 * Legacy single-bot dependency shape (pre-#860 plugin wiring): one scalar
 * credential set for the whole channel. Kept so single-bot deployments and
 * the config-wiring shim keep working unchanged; it behaves exactly like a
 * `bots` list with one (identity-less) entry.
 */
export interface TeamsRouterLegacyDeps {
  bot: TeamsBot;
  appId: string;
  appPassword: string;
  appType: TeamsBotAppType;
  appTenantId?: string;
  bots?: undefined;
}

/** Multi-bot dependency shape: the full `teams_bots[]` list with resolved
 *  passwords. `bots[0]` is the default bot (the one the legacy
 *  `/api/messages` alias serves). */
export interface TeamsRouterMultiBotDeps {
  bot: TeamsBot;
  bots: readonly TeamsRouterBotCredentials[];
}

export type TeamsRouterDeps = TeamsRouterLegacyDeps | TeamsRouterMultiBotDeps;

/**
 * Channel-native proactive send. Picks up a previously-captured
 * `ConversationReference` (taken by `TurnContext.getConversationReference`
 * during an inbound turn — typically stored on a routine row) and pushes
 * an outbound activity into that conversation. The routines runner uses
 * it to deliver a scheduled agent answer.
 *
 * Multi-bot dispatch resolves WHICH bot's adapter/credentials continue the
 * conversation, in priority order:
 *   1. `opts.botAppId` — explicit caller choice (e.g. from the per-bot
 *      conversation-ref store keying).
 *   2. `reference.bot.id` — captured references carry the owning bot as
 *      `28:<appId>` (this is the botAppId the ref store persists rows under).
 *   3. The default bot — best-effort fallback matching the ref store's
 *      contract: a legacy reference without bot attribution must keep
 *      delivering rather than throw.
 */
export type TeamsProactiveSend = (
  reference: Partial<ConversationReference>,
  build: (turnContext: TurnContext) => Promise<void>,
  opts?: { botAppId?: string },
) => Promise<void>;

/** One bot's long-lived runtime: its own credential factory + CloudAdapter
 *  plus a proactive sender pre-bound to that identity. */
export interface TeamsBotRuntime {
  /** Entra app id the adapter authenticates as (as configured, casing kept). */
  readonly appId: string;
  /** Full identity — `undefined` only for the legacy scalar-credential shape,
   *  which predates {@link TeamsBotIdentity}. */
  readonly identity?: TeamsBotIdentity;
  /** Long-lived per-bot CloudAdapter — share, never rebuild per request. */
  readonly adapter: CloudAdapter;
  /** The bot's own BF credential factory. SECURITY: holds the app password
   *  (`.password`) — expose to wiring/tests for identity assertions only,
   *  never log or serialise it. */
  readonly credentialsFactory: ConfigurationServiceClientCredentialFactory;
  /** `continueConversationAsync` pre-bound to THIS bot's adapter + app id. */
  readonly sendProactive: TeamsProactiveSend;
}

export interface TeamsRouterArtifacts {
  router: Router;
  /** The DEFAULT bot's long-lived CloudAdapter (back-compat surface) — share
   *  with proactive-send callers so the BF authentication factory + token
   *  cache are reused. Per-bot adapters live on {@link botRuntimes}. */
  adapter: CloudAdapter;
  /** Per-bot dispatching proactive send — see {@link TeamsProactiveSend}. */
  sendProactive: TeamsProactiveSend;
  /** One runtime per configured bot, in config order (`[0]` = default). */
  botRuntimes: readonly TeamsBotRuntime[];
  /** The default (first / shimmed) bot's runtime. */
  defaultBotRuntime: TeamsBotRuntime;
  /** Case-insensitive lookup by Entra app id — the turn/proactive resolver. */
  getBotRuntimeByAppId(appId: string): TeamsBotRuntime | undefined;
  /** Exact-match lookup by route slug (`/teams/:botSlug/messages`). Always
   *  `undefined` for the legacy scalar-credential shape (no slugs there). */
  getBotRuntimeBySlug(botSlug: string): TeamsBotRuntime | undefined;
}

/** Same read-path allowlist as the conversation-ref store (`isAllowedServiceUrl`,
 *  review M2): a poisoned reference — DB row, routine row, or cache — must not
 *  make `continueConversationAsync` present bot credentials to an attacker
 *  host. With N bots a poisoned `bot.id` could additionally pick ANOTHER bot's
 *  credentials, so the guard sits directly in front of every proactive turn,
 *  not only behind the ref store's loads. */
function isAllowedProactiveServiceUrl(serviceUrl: unknown): boolean {
  if (typeof serviceUrl !== 'string') return false;
  try {
    const host = new URL(serviceUrl).hostname.toLowerCase();
    return host === 'smba.trafficmanager.net' || host.endsWith('.botframework.com');
  } catch {
    return false;
  }
}

interface BotCredentialConfig {
  appId: string;
  appPassword: string;
  appType: TeamsBotAppType;
  tenantId?: string;
  identity?: TeamsBotIdentity;
}

/** Build one bot's long-lived credential factory + CloudAdapter. */
function buildBotRuntime(config: BotCredentialConfig): TeamsBotRuntime {
  const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: config.appId,
    MicrosoftAppPassword: config.appPassword,
    MicrosoftAppType: config.appType,
    MicrosoftAppTenantId: config.tenantId,
  });

  const botFrameworkAuthentication = createBotFrameworkAuthenticationFromConfiguration(
    null,
    credentialsFactory,
  );

  const adapter = new CloudAdapter(botFrameworkAuthentication);

  adapter.onTurnError = async (context, error) => {
    console.error('[teams] onTurnError:', error);
    try {
      await context.sendActivity('Es gab einen internen Fehler. Bitte erneut versuchen.');
    } catch (sendErr) {
      console.error('[teams] onTurnError: could not send error activity', sendErr);
    }
  };

  const sendProactive: TeamsProactiveSend = async (reference, build) => {
    if (reference.serviceUrl !== undefined && !isAllowedProactiveServiceUrl(reference.serviceUrl)) {
      throw new Error('[teams] proactive send refused — serviceUrl outside Bot Framework domains');
    }
    await adapter.continueConversationAsync(
      config.appId,
      reference as Partial<Activity>,
      build,
    );
  };

  return { appId: config.appId, identity: config.identity, adapter, credentialsFactory, sendProactive };
}

/** Normalise deps into the per-bot credential list (legacy scalars = a
 *  single-entry list). */
function toBotCredentialConfigs(deps: TeamsRouterDeps): BotCredentialConfig[] {
  if (deps.bots === undefined) {
    return [{
      appId: deps.appId,
      appPassword: deps.appPassword,
      appType: deps.appType,
      tenantId: deps.appTenantId,
    }];
  }
  return deps.bots.map(({ identity, appPassword }) => ({
    appId: identity.appId,
    appPassword,
    appType: identity.appType,
    tenantId: identity.tenantId,
    identity,
  }));
}

export function createTeamsRouter(deps: TeamsRouterDeps): TeamsRouterArtifacts {
  const configs = toBotCredentialConfigs(deps);
  if (configs.length === 0) {
    throw new Error('createTeamsRouter: at least one bot identity is required');
  }

  const botRuntimes: TeamsBotRuntime[] = [];
  const runtimesByAppId = new Map<string, TeamsBotRuntime>();
  const seenSlugs = new Set<string>();
  for (const config of configs) {
    const key = normalizeTeamsBotAppId(config.appId);
    if (!key) {
      throw new Error('createTeamsRouter: a bot identity has an empty appId');
    }
    if (runtimesByAppId.has(key)) {
      // Log-safe: slugs/display names only, never the appId itself.
      const label = config.identity ? teamsBotLogLabel(config.identity) : 'legacy bot';
      throw new Error(`createTeamsRouter: duplicate bot appId (second entry: ${label})`);
    }
    // Slug integrity gates the per-bot route: an empty slug would make a bot
    // unreachable, a duplicate would silently shadow the later entry (exact
    // first-match semantics in `findTeamsBotBySlug`). Slugs are operator-chosen
    // identifiers, so they are log-safe to include in the error.
    if (config.identity !== undefined) {
      const slug = config.identity.botSlug;
      if (slug.trim().length === 0) {
        throw new Error('createTeamsRouter: a bot identity has an empty botSlug');
      }
      if (seenSlugs.has(slug)) {
        throw new Error(`createTeamsRouter: duplicate botSlug '${slug}'`);
      }
      seenSlugs.add(slug);
    }
    const runtime = buildBotRuntime(config);
    botRuntimes.push(runtime);
    runtimesByAppId.set(key, runtime);
  }

  const defaultBotRuntime = botRuntimes[0];
  if (defaultBotRuntime === undefined) {
    // Unreachable (configs.length checked above) — narrows the indexed access.
    throw new Error('createTeamsRouter: at least one bot identity is required');
  }

  const getBotRuntimeByAppId = (appId: string): TeamsBotRuntime | undefined => {
    const key = normalizeTeamsBotAppId(appId);
    return key ? runtimesByAppId.get(key) : undefined;
  };

  /** Which bot's adapter continues this conversation — see
   *  {@link TeamsProactiveSend} for the priority order. Unknown/unattributed
   *  references fall back to the default bot (best-effort, matching the ref
   *  store's legacy-row contract) rather than throwing. */
  const resolveRuntimeForReference = (
    reference: Partial<ConversationReference>,
    explicitBotAppId: string | undefined,
  ): TeamsBotRuntime => {
    const referencedAppId =
      explicitBotAppId ?? parseTeamsBotKey(reference.bot?.id ?? '');
    if (referencedAppId === undefined) return defaultBotRuntime;
    const runtime = getBotRuntimeByAppId(referencedAppId);
    if (runtime) return runtime;
    console.warn(
      '[teams] proactive send: reference bot is not a configured bot — falling back to the default bot',
    );
    return defaultBotRuntime;
  };

  // Slug → runtime resolution goes through the canonical helper
  // (`findTeamsBotBySlug`, exact match) so route matching can never drift
  // from the identity module's slug semantics. Legacy scalar deps have no
  // identity → empty list → every `:botSlug` resolves to 404.
  const botIdentities: TeamsBotIdentity[] = botRuntimes.flatMap(
    (runtime) => (runtime.identity !== undefined ? [runtime.identity] : []),
  );
  const getBotRuntimeBySlug = (botSlug: string): TeamsBotRuntime | undefined => {
    const identity = findTeamsBotBySlug(botIdentities, botSlug);
    return identity ? getBotRuntimeByAppId(identity.appId) : undefined;
  };

  /** Hand the request to ONE bot's CloudAdapter — BF authentication happens
   *  inside `adapter.process`, against exactly that bot's credential set. */
  const processTurn = (runtime: TeamsBotRuntime, req: Request, res: Response): void => {
    runtime.adapter.process(req, res, (context) => deps.bot.run(context));
  };

  const router = Router();

  // Route surface (#860 W0a). The kernel mounts this router at EXACTLY the
  // prefix given to `core.registerRouter` (plugin.ts mounts at `/api`; the
  // route registry injects nothing), so the `teams` segment must live HERE:
  //
  //   /messages                  → default bot  (live legacy path — the
  //                                endpoint existing Azure bots point at
  //                                today, public `/api/messages`)
  //   /teams/messages            → default bot  (alias — the path the
  //                                manifest documents for Azure setup,
  //                                public `/api/teams/messages`)
  //   /teams/:botSlug/messages   → that bot's OWN adapter/credentials,
  //                                404 for unknown slugs
  //
  // Both no-slug paths are aliases of the default bot so no messaging
  // endpoint registered on an existing Azure bot breaks, whichever of the
  // two documented spellings its operator followed.
  router.post('/messages', (req: Request, res: Response) => {
    processTurn(defaultBotRuntime, req, res);
  });
  router.post('/teams/messages', (req: Request, res: Response) => {
    processTurn(defaultBotRuntime, req, res);
  });
  router.post('/teams/:botSlug/messages', (req: Request, res: Response) => {
    // Express 5 types params as string | string[] (repeatable patterns);
    // this route's `:botSlug` is a single segment — anything else is not a
    // configured slug.
    const rawSlug: string | string[] | undefined = req.params.botSlug;
    const botSlug = typeof rawSlug === 'string' ? rawSlug : '';
    const runtime = botSlug ? getBotRuntimeBySlug(botSlug) : undefined;
    if (runtime === undefined) {
      // MUST 404, never fall through to the default bot: Azure bot A's
      // traffic must not be answered with bot B's credentials. (Slugs are
      // published in the generated Teams app packages, so distinguishing
      // "unknown slug" from "known slug, bad auth" leaks nothing secret.)
      // `botSlug` is caller-controlled — JSON.stringify + cap neutralises
      // log injection before it reaches the log line.
      console.warn(
        `[teams] message for unknown botSlug ${JSON.stringify(botSlug).slice(0, 66)} — no bot is configured under this slug`,
      );
      res.status(404).json({
        code: 'teams.unknown_bot',
        message: 'no bot is configured under this slug',
      });
      return;
    }
    processTurn(runtime, req, res);
  });

  const sendProactive: TeamsProactiveSend = async (reference, build, opts) => {
    const runtime = resolveRuntimeForReference(reference, opts?.botAppId);
    await runtime.sendProactive(reference, build);
  };

  return {
    router,
    adapter: defaultBotRuntime.adapter,
    sendProactive,
    botRuntimes,
    defaultBotRuntime,
    getBotRuntimeByAppId,
    getBotRuntimeBySlug,
  };
}

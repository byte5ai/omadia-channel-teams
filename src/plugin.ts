import type Anthropic from '@anthropic-ai/sdk';
import type { Pool } from 'pg';

import {
  ROUTINES_INTEGRATION_SERVICE_NAME,
  type PluginContext,
  type RoutinesIntegration,
} from '@omadia/plugin-api';
import {
  CHANNEL_RESOLVER_SERVICE,
  InMemoryConversationHistoryStore,
  isNoReply,
  logNoReplyDrop,
  type ChannelBindingResolver,
  type ChannelHandle,
  type ChannelKeyDirectory,
  type ChatAgent,
  type CoreApi,
} from '@omadia/channel-sdk';

import { buildTeamsChannelKeyDirectory } from './channelKeyDirectory.js';
import { TeamsConversationObserver } from './teamsConversationObserver.js';

/**
 * Narrow shim for the kernel's `channelDirectoryRegistry@1` service so
 * the plugin doesn't drag in `@omadia/middleware` types. The kernel-side
 * surface has more methods (size, types, listAll) — we only consume
 * register / unregister.
 */
interface ChannelDirectoryRegistryShim {
  register(directory: ChannelKeyDirectory): void;
  unregister(channelType: string): void;
}

import type { TigrisStore } from '@omadia/diagrams';
import type { Microsoft365Accessor } from '@omadia/integration-microsoft365';
import type { ChatAgentBundle } from '@omadia/orchestrator';

import type {
  EmbeddingClient,
  TeamsConfigShim,
  TopicDetector,
  TurnContextModule,
} from './kernel-types.js';
import { createAttachmentsRouter } from './attachmentsRouter.js';
import { TeamsAttachmentStore } from './teamsAttachmentStore.js';
import { TeamsBot } from './teamsBot.js';
import { TeamsRosterProvider } from './teamsRoster.js';
import { createTeamsRouter } from './messagesRouter.js';
import { createTeamsUiRouter } from './uiRouter.js';

/**
 * Minimal shape of the kernel's UiRouteCatalog service. Declared here
 * so the plugin doesn't need a hard dep on @omadia/plugin-api's
 * UiRouteDescriptor type — keeps the plugin compilable in isolation
 * even if the host's plugin-api version skews ahead.
 */
interface UiRouteCatalogShim {
  list(): readonly {
    readonly pluginId: string;
    readonly routeId: string;
    readonly path: string;
    readonly title: string;
  }[];
}

/**
 * Microsoft Teams as a first-class ChannelPlugin. The Bot Framework App
 * credentials (app_id, tenant_id, app_password) flow in via ctx — pulled
 * from the `@omadia/integration-microsoft365` integration the plugin
 * depends on, not from `.env`. The same integration publishes a
 * `Microsoft365Accessor` service (`ctx.services.get('microsoft365.graph')`)
 * which owns the shared Graph client used for attachment downloads.
 * Teams-specific config (SSO connection name, attachment key prefix)
 * comes from this plugin's own registry entry.
 *
 * Phase 5B: standalone activate() shape — no constructor, no Deps
 * interface. Every kernel-side resource is sourced from `ctx.services`
 * (anthropicClient, tigrisStore, graphPool, graphTenantId,
 * embeddingClient, topicDetector, turnContext, microsoft365.graph,
 * chatAgent, routinesIntegration) or constructed in-package
 * (InMemoryConversationHistoryStore). This is the contract the
 * plugin-store flow needs: the dynamic-channel-resolver imports
 * `dist/plugin.js` and calls `activate(ctx, core)` directly with no
 * knowledge of plugin-specific Deps shapes.
 */
export async function activate(
  ctx: PluginContext,
  core: CoreApi,
): Promise<ChannelHandle> {
  // --- Credentials: MS App registration comes from MS365 integration ----
  const appId = ctx.config.require<string>('microsoft_app_id');
  const tenantId = ctx.config.require<string>('microsoft_tenant_id');
  const appPassword = await ctx.secrets.require('microsoft_app_password');

  // --- Channel-specific config (own registry entry) --------------------
  const config = readTeamsConfigFromEnv();
  const ssoConnectionName = ctx.config.get<string>(
    'teams_sso_connection_name',
  );
  const attachmentKeyPrefix =
    ctx.config.get<string>('teams_attachment_key_prefix') ??
    config.TEAMS_ATTACHMENT_KEY_PREFIX;

  // --- Channels-directory contribution (A+B) ---------------------------
  // Register this Teams bot with the kernel's channel-directory registry
  // so the operator dashboard at /operator/channels can show it as a
  // pickable target instead of asking the operator to memorise the
  // 28:<app-id> key. The registry is optional: a host that has not
  // landed the A+B platform change degrades to "no entry" gracefully.
  //
  // The observer is the SAME instance handed to TeamsBot below — the
  // bot records every inbound conversation into it, and the directory's
  // `listKeys()` reads them at render time. That way the operator sees
  // every Teams channel the bot has been messaged in as a bindable
  // target, alongside the bot-level catch-all.
  const conversationObserver = new TeamsConversationObserver();
  const channelDirectoryDisplayLabel = ctx.config.get<string>(
    'teams_directory_label',
  );
  const channelDirectory = buildTeamsChannelKeyDirectory({
    microsoftAppId: appId,
    microsoftTenantId: tenantId,
    ...(channelDirectoryDisplayLabel
      ? { displayLabel: channelDirectoryDisplayLabel }
      : {}),
    conversationObserver,
  });
  const channelDirectoryRegistry =
    ctx.services.get<ChannelDirectoryRegistryShim>(
      'channelDirectoryRegistry',
    );
  if (channelDirectoryRegistry) {
    channelDirectoryRegistry.register(channelDirectory);
    core.log(
      'info',
      `channel-key directory contributed: teams · 28:${appId.slice(0, 8)}…`,
    );
  } else {
    core.log(
      'info',
      'channelDirectoryRegistry@1 not published — skipping /operator/channels contribution',
    );
  }

  // --- chatAgent — late-resolved via capability registry (S+10-4b) -----
  // Manifest declares `requires: ["chatAgent@^1"]`, so the resolver guards
  // activation order. The check below is a defensive type-narrow guard:
  // `ServicesAccessor.get<T>` is typed `T | undefined` regardless of the
  // requires-declaration, so we surface a clear error if the resolver
  // ever lets us through without the capability.
  const chatAgentBundle = ctx.services.get<ChatAgentBundle>('chatAgent');
  if (!chatAgentBundle) {
    throw new Error(
      'chatAgent@^1 capability not resolved — manifest declares it as required but ServicesAccessor returned undefined. Check that the harness-orchestrator plugin published the capability before channel activation.',
    );
  }
  const chatAgent = chatAgentBundle.agent;

  // --- Phase 5B: kernel-published deps via ctx.services ----------------
  const anthropicClient = ctx.services.get<Anthropic>('anthropicClient');
  if (!anthropicClient) {
    throw new Error(
      "anthropicClient service not published — kernel must publish it before channel activation (see middleware/src/index.ts).",
    );
  }
  const turnContext = ctx.services.get<TurnContextModule>('turnContext');
  if (!turnContext) {
    throw new Error(
      'turnContext service not published — kernel must publish it before channel activation (see middleware/src/index.ts).',
    );
  }
  const embeddingClient = ctx.services.get<EmbeddingClient>('embeddingClient');
  const diagramStore = ctx.services.get<TigrisStore>('tigrisStore');
  const graphPool = ctx.services.get<Pool>('graphPool');
  const graphTenantId =
    ctx.services.get<string>('graphTenantId') ?? 'default';
  const topicDetector = ctx.services.get<TopicDetector>('topicDetector');
  const routinesIntegration = ctx.services.get<RoutinesIntegration>(
    ROUTINES_INTEGRATION_SERVICE_NAME,
  );

  // Per-channel in-memory history. Channel-prefixed scopes
  // (`teams:<conversationId>`) are isolated by construction — but a
  // dedicated instance also keeps a Teams history wipe from clobbering
  // any other channel's state on the same process.
  const conversationHistoryStore = new InMemoryConversationHistoryStore();

  if (topicDetector) {
    core.log('info', 'topic detector ready');
  } else {
    core.log(
      'info',
      'topic detector DISABLED (OLLAMA_BASE_URL not set) — history passes through',
    );
  }

  // --- Optional Teams attachment persistence ---------------------------
  let teamsAttachmentStore: TeamsAttachmentStore | undefined;
  if (
    config.TEAMS_ATTACHMENT_STORAGE_ENABLED &&
    diagramStore &&
    graphPool
  ) {
    const signingReady =
      Boolean(config.ATTACHMENT_URL_SECRET) &&
      Boolean(config.ATTACHMENT_PUBLIC_BASE_URL);
    // GraphClient comes from the microsoft365 integration's Accessor —
    // single shared instance, token cache shared across the process.
    // When the integration isn't active the attachment store degrades
    // gracefully (SharePoint-linked attachments fall through to their
    // `undefined` branch in TeamsAttachmentStore).
    const microsoft365 = ctx.services.get<Microsoft365Accessor>(
      'microsoft365.graph',
    );
    teamsAttachmentStore = new TeamsAttachmentStore({
      tigris: diagramStore,
      pool: graphPool,
      tenant: graphTenantId,
      keyPrefix: attachmentKeyPrefix,
      maxBytes: config.TEAMS_ATTACHMENT_MAX_BYTES,
      ...(signingReady
        ? {
            signing: {
              secret: config.ATTACHMENT_URL_SECRET!,
              publicBaseUrl: config.ATTACHMENT_PUBLIC_BASE_URL!,
              ttlSec: config.ATTACHMENT_SIGNED_URL_TTL_SEC,
            },
          }
        : {}),
      ...(microsoft365 ? { graphClient: microsoft365.app } : {}),
    });
    core.log(
      'info',
      `teams attachment store ready (prefix=${attachmentKeyPrefix}, maxBytes=${String(config.TEAMS_ATTACHMENT_MAX_BYTES)}, signing=${signingReady ? 'on' : 'off'})`,
    );
  } else if (config.TEAMS_ATTACHMENT_STORAGE_ENABLED) {
    core.log(
      'warn',
      'teams attachment store DISABLED — Tigris or Neon pool missing',
    );
  }

  // --- Roster provider -------------------------------------------------
  const teamsRosterProvider = new TeamsRosterProvider();
  core.log('info', 'teams roster provider ready (ttl=5min)');

  // --- Routines wiring (only when the kernel published the integration)
  const captureRoutineTurn = routinesIntegration
    ? (info: { tenant: string; userId: string; conversationRef: unknown }) =>
        routinesIntegration.captureRoutineTurn({
          tenant: info.tenant,
          userId: info.userId,
          channel: 'teams',
          conversationRef: info.conversationRef,
        })
    : undefined;
  const handleRoutineAction = routinesIntegration
    ? (input: {
        action: 'pause' | 'resume' | 'trigger_now' | 'delete';
        id: string;
      }) => routinesIntegration.handleRoutineAction(input)
    : undefined;
  const buildRoutineSmartCardAttachment = routinesIntegration
    ? (input: {
        routine: { id: string; name: string; cron: string };
        body: string;
      }) => routinesIntegration.buildRoutineSmartCardAttachment(input)
    : undefined;
  const buildRoutineListSmartCardAttachment = routinesIntegration
    ? routinesIntegration.buildRoutineListSmartCardAttachment.bind(
        routinesIntegration,
      )
    : undefined;

  // --- Per-turn Agent resolution (Phase A) -----------------------------
  // Late-resolved channelResolver@1: when the multi-orchestrator
  // registry is active AND the operator bound this Teams bot to an
  // Agent in /operator/channels, each inbound turn is routed to that
  // Agent. Without the resolver service the bot falls back to the
  // legacy chatAgent@1 (the `chatAgent` constant captured above).
  const channelResolver = ctx.services.get<ChannelBindingResolver>(
    CHANNEL_RESOLVER_SERVICE,
  );
  const resolveChatAgentForActivity = channelResolver
    ? (input: { channelType: 'teams'; channelKey: string }): {
        decision: 'bound' | 'fallback' | 'reject';
        chatAgent?: ChatAgent;
      } => {
        const decision = channelResolver.resolve(
          input.channelType,
          input.channelKey,
        );
        return {
          decision: decision.decision,
          ...(decision.chatAgent ? { chatAgent: decision.chatAgent } : {}),
        };
      }
    : undefined;

  // --- Bot + Router ----------------------------------------------------
  const bot = new TeamsBot(
    chatAgent,
    conversationHistoryStore,
    topicDetector,
    turnContext,
    teamsAttachmentStore,
    teamsRosterProvider,
    ssoConnectionName,
    graphTenantId,
    captureRoutineTurn,
    handleRoutineAction,
    buildRoutineListSmartCardAttachment,
    resolveChatAgentForActivity,
    conversationObserver,
  );

  if (resolveChatAgentForActivity) {
    core.log(
      'info',
      'Teams per-turn Agent resolution active via channelResolver@1',
    );
  } else {
    core.log(
      'info',
      'channelResolver@1 not published — Teams routes all turns to default chatAgent',
    );
  }
  if (ssoConnectionName) {
    core.log(
      'info',
      `Teams SSO enabled (connection=${ssoConnectionName}) — calendar tools will receive user tokens`,
    );
  } else {
    core.log(
      'info',
      'Teams SSO disabled (teams_sso_connection_name unset) — calendar tools surface sso_unavailable',
    );
  }
  // anthropicClient is currently unused inside this scope — TeamsBot
  // pulls Claude via `chatAgent` only. Keep the resolved variable so a
  // future direct-Anthropic feature (e.g. inline image-OCR) doesn't
  // need to re-thread DI.
  void anthropicClient;
  // embeddingClient + topicDetector are passed through as constructor
  // args above (topicDetector) or stay undefined-tracked for diagnostics.
  void embeddingClient;

  const { router, sendProactive } = createTeamsRouter({
    bot,
    appId,
    appPassword,
    appType: config.MICROSOFT_APP_TYPE,
    appTenantId: tenantId,
  });

  // Mount at /api — same prefix as before, now owned by the channel
  // runtime so deactivation cleanly returns 503 instead of crashing.
  core.registerRouter(ctx.agentId, '/api', router);
  core.log(
    'info',
    `Teams endpoint active at /api/messages (app=${appId}, type=${config.MICROSOFT_APP_TYPE}, credentials=vault)`,
  );

  // Hand the long-lived CloudAdapter back to the kernel so the routines
  // runner can deliver scheduled answers via continueConversationAsync.
  // No-op when routinesIntegration wasn't published (routines off /
  // no Postgres).
  if (routinesIntegration) {
    routinesIntegration.publishProactiveSend(
      'teams',
      async (conversationRef, message, routine) => {
        // Routines fire on a cron without an active user question — the
        // sentinel is the agent's way to signal "nothing to report this
        // run". Drop the delivery; the routine's run-history still
        // records the turn for audit, but the user's Teams chat stays
        // quiet.
        if (isNoReply(message)) {
          logNoReplyDrop('teams', {
            trigger: 'routine',
            ...(routine?.id ? { routineId: routine.id } : {}),
            ...(routine?.name ? { routineName: routine.name } : {}),
          });
          return;
        }
        await sendProactive(
          conversationRef as Parameters<typeof sendProactive>[0],
          async (turnContext) => {
            // When we have routine metadata + the integration's card
            // builder, render the Adaptive Card so the user sees this
            // is a cron-triggered delivery and can pause/delete inline.
            // Otherwise fall back to plain text (channel-agnostic
            // minimum).
            if (routine && buildRoutineSmartCardAttachment) {
              const attachment = buildRoutineSmartCardAttachment({
                routine,
                body: message.text,
              });
              await turnContext.sendActivity({
                type: 'message',
                attachments: [attachment],
              });
              return;
            }
            await turnContext.sendActivity({
              type: 'message',
              text: message.text,
            });
          },
        );
      },
    );
  }

  // --- Signed attachment proxy ----------------------------------------
  // Plugin mounts `/attachments/<signed-key>` itself (consistent with the
  // diagrams plugin pattern from S+2). Only runs when the Tigris store is
  // available AND the URL-signing secret is configured — otherwise the
  // plugin silently skips the mount; persisted-attachment flows fall back
  // to storage-native retrieval (no exception, no hard failure).
  if (diagramStore && config.ATTACHMENT_URL_SECRET) {
    const attachmentsRouter = createAttachmentsRouter({
      store: diagramStore,
      secret: config.ATTACHMENT_URL_SECRET,
      allowedPrefixes: [`${attachmentKeyPrefix}/`],
    });
    core.registerRouter(ctx.agentId, '/attachments', attachmentsRouter);
    core.log(
      'info',
      `attachments proxy mounted at /attachments/<key> (prefix=${attachmentKeyPrefix}/)`,
    );
  }

  // --- Teams bridge uiRoutes (Hub + Tab-Config) -----------------------
  // Mount at /p/channel-teams via ctx.routes.register so the URLs are
  // browser-reachable through web-ui's /p/:path* rewrite. The Teams
  // App-Manifest's staticTabs[] points at /p/channel-teams/hub and
  // configurableTabs[].configurationUrl at /p/channel-teams/tab-config.
  //
  // Discovery is live — the kernel publishes a UiRouteCatalog as the
  // `uiRouteCatalog` service. Every Hub + Tab-Config request pulls the
  // current set from it, so plugins that come and go via upload/uninstall
  // surface automatically without a channel-teams code change.
  const uiRouteCatalog = ctx.services.get<UiRouteCatalogShim>(
    'uiRouteCatalog',
  );
  if (!uiRouteCatalog) {
    throw new Error(
      "channel-teams: required service 'uiRouteCatalog' not published — kernel must publish it before plugins activate (see middleware/src/index.ts).",
    );
  }
  const teamsUiRouter = createTeamsUiRouter({
    webUiOrigin: ctx.config.get<string>('web_ui_origin') ?? '',
    discover: () =>
      uiRouteCatalog
        .list()
        // Don't list channel-teams' OWN surfaces in the Hub/Tab-Config —
        // the operator pins those via the App-Manifest's staticTabs[]
        // entry, not as targets-of-itself.
        .filter((r) => r.pluginId !== '@omadia/channel-teams')
        .map((r) => ({
          pluginId: r.pluginId,
          routeId: r.routeId,
          path: r.path,
          title: r.title,
        })),
  });
  const disposeTeamsUi = ctx.routes.register(
    '/p/channel-teams',
    teamsUiRouter,
  );
  core.log('info', 'teams bridge uiRoutes mounted at /p/channel-teams/{hub,tab-config}');

  // --- Notification handler ------------------------------------------
  // Slice 2: real Graph `sendActivityNotification` Bell-Icon push when
  // a target team-id is configured. Falls back to log-only when not
  // (lets the plugin run in dev without poking Teams). The Activity-
  // Type `pluginEvent` MUST be declared in the App-Manifest's
  // `activities.activityTypes[]`; the TeamsActivity.Send.Group
  // permission was consented when the bot was first deployed.
  const notifyTargetTeamId = ctx.config.get<string>(
    'teams_notify_team_id',
  );
  const notifyTopicWebUrl =
    ctx.config.get<string>('teams_notify_topic_url') ??
    'https://odoo-bot-harness.fly.dev/p/channel-teams/hub';
  // Re-resolve the microsoft365 accessor here — the earlier read inside
  // the attachment-store branch is scoped to that `if`-block. Reading
  // again is cheap (single ServiceRegistry lookup) and keeps the
  // notification path independent of the attachment-store config.
  const notifyMs365 = ctx.services.get<Microsoft365Accessor>(
    'microsoft365.graph',
  );
  const disposeTeamsNotify = ctx.notifications.registerChannel(
    'teams',
    async (payload) => {
      const recipientsLabel =
        payload.recipients === 'broadcast'
          ? 'broadcast'
          : `[${payload.recipients.length} recipient(s)]`;
      const previewBody =
        payload.body.length > 140
          ? `${payload.body.slice(0, 140)}…`
          : payload.body;
      core.log(
        'info',
        `[notify] teams ← plugin=${payload.pluginId} title=${JSON.stringify(payload.title)} body=${JSON.stringify(previewBody)} deepLink=${JSON.stringify(payload.deepLink ?? '')} recipients=${recipientsLabel}`,
      );

      if (!notifyTargetTeamId) {
        core.log(
          'info',
          '[notify] teams_notify_team_id not configured — skipping Graph sendActivityNotification (log-only fallback)',
        );
        return;
      }
      if (!notifyMs365) {
        core.log(
          'info',
          '[notify] microsoft365.graph service unavailable — skipping Graph sendActivityNotification',
        );
        return;
      }
      const webUrl = payload.deepLink
        ? `${notifyTopicWebUrl.replace(/\/p\/channel-teams\/hub\/?$/, '')}${payload.deepLink}`
        : notifyTopicWebUrl;
      try {
        await notifyMs365.app.sendActivityNotification({
          scope: `/teams/${encodeURIComponent(notifyTargetTeamId)}`,
          activityType: 'pluginEvent',
          previewText: previewBody,
          templateParameters: [
            { name: 'pluginId', value: payload.pluginId },
            { name: 'title', value: payload.title },
          ],
          topic: {
            source: 'text',
            value: payload.title,
            webUrl,
          },
        });
        core.log(
          'info',
          `[notify] graph sendActivityNotification ok team=${notifyTargetTeamId.slice(0, 20)} type=pluginEvent`,
        );
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        core.log(
          'info',
          `[notify] graph sendActivityNotification FAILED: ${message}`,
        );
      }
    },
  );
  core.log(
    'info',
    `notification handler registered for channel "teams" (mode=${notifyTargetTeamId ? 'graph→team:' + notifyTargetTeamId.slice(0, 12) : 'log-only'})`,
  );

  // --- Handle ----------------------------------------------------------
  return {
    close: async () => {
      // Graceful shutdown: the route registry flips the active flag so
      // incoming requests return 503. The TeamsBot holds no timers or
      // sockets we need to close explicitly; attachment store / graph
      // client rely on HTTP pools that shut down with the process.
      disposeTeamsUi();
      disposeTeamsNotify();
      // Drop our channel-directory contribution so /operator/channels
      // stops surfacing this bot when the plugin is uninstalled or
      // re-activated with new config.
      channelDirectoryRegistry?.unregister('teams');
      core.log('info', 'Teams channel closed (routes now 503)');
    },
  };
}

// ---------------------------------------------------------------------------
// TeamsConfigShim — read narrow process.env subset
// ---------------------------------------------------------------------------

function readTeamsConfigFromEnv(): TeamsConfigShim {
  const appType = process.env['MICROSOFT_APP_TYPE'];
  const microsoftAppType: TeamsConfigShim['MICROSOFT_APP_TYPE'] =
    appType === 'SingleTenant' || appType === 'UserAssignedMSI'
      ? appType
      : 'MultiTenant';

  const storageEnabled =
    String(process.env['TEAMS_ATTACHMENT_STORAGE_ENABLED'] ?? '')
      .trim()
      .toLowerCase() === 'true';

  const maxBytes = parsePositiveInt(
    process.env['TEAMS_ATTACHMENT_MAX_BYTES'],
    25 * 1024 * 1024,
  );
  const signedTtl = parsePositiveInt(
    process.env['ATTACHMENT_SIGNED_URL_TTL_SEC'],
    600,
  );

  const keyPrefix =
    process.env['TEAMS_ATTACHMENT_KEY_PREFIX']?.trim() || 'teams-attachments';

  const urlSecret = process.env['ATTACHMENT_URL_SECRET'];
  const publicBaseUrl = process.env['ATTACHMENT_PUBLIC_BASE_URL'];

  return {
    MICROSOFT_APP_TYPE: microsoftAppType,
    TEAMS_ATTACHMENT_KEY_PREFIX: keyPrefix,
    TEAMS_ATTACHMENT_STORAGE_ENABLED: storageEnabled,
    TEAMS_ATTACHMENT_MAX_BYTES: maxBytes,
    ATTACHMENT_SIGNED_URL_TTL_SEC: signedTtl,
    ...(urlSecret ? { ATTACHMENT_URL_SECRET: urlSecret } : {}),
    ...(publicBaseUrl ? { ATTACHMENT_PUBLIC_BASE_URL: publicBaseUrl } : {}),
  };
}

function parsePositiveInt(raw: string | undefined, fallback: number): number {
  if (raw === undefined || raw === '') return fallback;
  const n = Number.parseInt(raw, 10);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

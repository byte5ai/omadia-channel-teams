import type Anthropic from '@anthropic-ai/sdk';
import type { Pool } from 'pg';

import type { PluginContext } from '@omadia/plugin-api';
import type { ChannelHandle, ChannelPlugin, CoreApi } from '@omadia/channel-sdk';
import type { TigrisStore } from '@omadia/diagrams';
import type { Microsoft365Accessor } from '@omadia/integration-microsoft365';
import type { ChatAgentBundle } from '@omadia/orchestrator';

import type {
  ConversationHistoryStore,
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

/**
 * Shared runtime dependencies required to instantiate the Teams channel.
 * Constructed once by the kernel (`src/index.ts`) and passed in — the plugin
 * itself never reaches back into the kernel for value-carrying deps.
 *
 * Three of these are deliberately INSTANCES from the kernel (not classes the
 * plugin constructs):
 *   - `conversationHistoryStore` — a per-process in-memory scope map; the
 *     kernel constructs once and shares across the Teams activation.
 *   - `topicDetector` — optional, constructed by the kernel iff an
 *     `embeddingClient` is available. Undefined disables topic detection.
 *   - `turnContext` — the kernel's `AsyncLocalStorage` namespace. The plugin
 *     and the orchestrator MUST share the same instance or `runWithChatParticipants`
 *     silently fails to hand the roster provider to the orchestrator.
 */
export interface TeamsChannelPluginDeps {
  anthropicClient: Anthropic;
  embeddingClient: EmbeddingClient | undefined;
  diagramStore: TigrisStore | undefined;
  graphPool: Pool | undefined;
  graphTenantId: string;
  /** Pre-constructed by the kernel. Shared across the Teams activation. */
  conversationHistoryStore: ConversationHistoryStore;
  /** Pre-constructed by the kernel iff `embeddingClient` is available. */
  topicDetector: TopicDetector | undefined;
  /** Kernel-owned AsyncLocalStorage module — injected so the adapter and the
   *  orchestrator share the same turn-scope storage. */
  turnContext: TurnContextModule;
  /** Narrow subset of the kernel `.env` Config the Teams plugin reads. */
  config: TeamsConfigShim;
  /**
   * Optional kernel-owned hook invoked at the very start of every inbound
   * Teams turn, BEFORE the chatAgent is called. The kernel uses it to
   * install the routines per-turn AsyncLocalStorage so the
   * `manage_routine` tool can attribute `create` calls to the right user
   * and capture the channel-native delivery handle. The plugin treats it
   * as opaque — pass everything we know about the inbound turn and let
   * the kernel decide what to do. Undefined → no-op (routines feature
   * disabled or running without Postgres).
   */
  captureRoutineTurn?: (info: {
    tenant: string;
    userId: string;
    conversationRef: unknown;
  }) => void;
  /**
   * Optional callback invoked once at activation. The plugin hands back a
   * `(conversationRef, message) => Promise<void>` bound to the long-lived
   * BotFramework CloudAdapter; the kernel registers it as the Teams
   * `ProactiveSender` so the routines runner can deliver scheduled
   * answers. Undefined → routines feature off / no proactive delivery.
   */
  publishProactiveSend?: (
    send: (
      conversationRef: unknown,
      message: { text: string },
      routine?: { id: string; name: string; cron: string },
    ) => Promise<void>,
  ) => void;
  /**
   * Optional kernel-owned handler invoked when the user clicks a Pause /
   * Delete button on the routine smart-card the bot delivered. The bot
   * detects `activity.value.kind === 'routine.action'` on the inbound
   * activity and forwards the payload here. Returns a short confirmation
   * string the bot renders back into the chat.
   */
  handleRoutineAction?: (input: {
    action: 'pause' | 'resume' | 'trigger_now' | 'delete';
    id: string;
  }) => Promise<string>;
  /**
   * Optional builder that turns routine metadata + the agent's prose
   * answer into a Bot-Framework Attachment (Adaptive Card). The kernel
   * supplies it from the routines plugin; the bot wraps the result
   * inside `turnContext.sendActivity({ attachments: [...] })` instead of
   * sending plain text. Undefined → fall back to plain-text delivery.
   */
  buildRoutineSmartCardAttachment?: (input: {
    routine: { id: string; name: string; cron: string };
    body: string;
  }) => { contentType: string; content: unknown };
  /**
   * Builder for the routine LIST smart card emitted by
   * `manage_routine.list`. Same shape as the single-routine builder but
   * takes the list payload (filter + totals + rows). Kernel supplies this
   * from the routines plugin; the bot wraps the result in
   * `MessageFactory.attachment(...)` and posts it as a sidecar after the
   * agent's prose answer.
   */
  buildRoutineListSmartCardAttachment?: (input: {
    filter: 'all' | 'active' | 'paused';
    totals: { all: number; active: number; paused: number };
    routines: Array<{
      id: string;
      name: string;
      cron: string;
      prompt: string;
      status: 'active' | 'paused';
      lastRunAt: string | null;
      lastRunStatus: 'ok' | 'error' | 'timeout' | null;
    }>;
  }) => { contentType: string; content: unknown };
}

/**
 * Microsoft Teams as a first-class ChannelPlugin. The Bot Framework App
 * credentials (app_id, tenant_id, app_password) flow in via ctx — pulled
 * from the `de.byte5.integration.microsoft365` integration the plugin
 * depends on, not from `.env`. The same integration publishes a
 * `Microsoft365Accessor` service (`ctx.services.get('microsoft365.graph')`)
 * which owns the shared Graph client used for attachment downloads.
 * Teams-specific config (SSO connection name, attachment key prefix)
 * comes from this plugin's own registry entry.
 */
export class TeamsChannelPlugin implements ChannelPlugin {
  constructor(private readonly deps: TeamsChannelPluginDeps) {}

  async activate(ctx: PluginContext, core: CoreApi): Promise<ChannelHandle> {
    // --- Credentials: MS App registration comes from MS365 integration ----
    const appId = ctx.config.require<string>('microsoft_app_id');
    const tenantId = ctx.config.require<string>('microsoft_tenant_id');
    const appPassword = await ctx.secrets.require('microsoft_app_password');

    // --- Channel-specific config (own registry entry) --------------------
    const ssoConnectionName = ctx.config.get<string>(
      'teams_sso_connection_name',
    );
    const attachmentKeyPrefix =
      ctx.config.get<string>('teams_attachment_key_prefix') ??
      this.deps.config.TEAMS_ATTACHMENT_KEY_PREFIX;

    // --- chatAgent — late-resolved via capability registry (S+10-4b) -----
    // Manifest declares `requires: ["chatAgent@^1"]`, so the resolver guards
    // activation order. The check below is a defensive type-narrow guard:
    // `ServicesAccessor.get<T>` is typed `T | undefined` regardless of the
    // requires-declaration, so we surface a clear error if the resolver
    // ever lets us through without the capability.
    const chatAgentBundle =
      ctx.services.get<ChatAgentBundle>('chatAgent');
    if (!chatAgentBundle) {
      throw new Error(
        'chatAgent@^1 capability not resolved — manifest declares it as required but ServicesAccessor returned undefined. Check that the harness-orchestrator plugin published the capability before channel activation.',
      );
    }
    const chatAgent = chatAgentBundle.agent;

    // --- Shared deps from bootstrap --------------------------------------
    const {
      config,
      diagramStore,
      graphPool,
      graphTenantId,
      conversationHistoryStore,
      topicDetector,
      turnContext,
      captureRoutineTurn,
    } = this.deps;

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
      this.deps.handleRoutineAction,
      this.deps.buildRoutineListSmartCardAttachment,
    );
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
    // No-op when the kernel didn't wire the callback (routines off).
    if (this.deps.publishProactiveSend) {
      const buildSmartCardAttachment =
        this.deps.buildRoutineSmartCardAttachment;
      this.deps.publishProactiveSend(async (conversationRef, message, routine) => {
        await sendProactive(
          conversationRef as Parameters<typeof sendProactive>[0],
          async (turnContext) => {
            // When we have routine metadata + a card-builder, render the
            // Adaptive Card so the user sees this is a cron-triggered
            // delivery and can pause/delete inline. Otherwise fall back
            // to plain text (channel-agnostic minimum).
            if (routine && buildSmartCardAttachment) {
              const attachment = buildSmartCardAttachment({
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
      });
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

    // --- Handle ----------------------------------------------------------
    return {
      close: async () => {
        // Graceful shutdown: the route registry flips the active flag so
        // incoming requests return 503. The TeamsBot holds no timers or
        // sockets we need to close explicitly; attachment store / graph
        // client rely on HTTP pools that shut down with the process.
        core.log('info', 'Teams channel closed (routes now 503)');
      },
    };
  }
}

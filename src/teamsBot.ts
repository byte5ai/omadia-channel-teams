import {
  ActivityTypes,
  CardFactory,
  MessageFactory,
  TeamsActivityHandler,
  TurnContext,
  type Activity,
  type SigninStateVerificationQuery,
} from 'botbuilder';
import {
  formatSessionScope,
  isNoReply,
  logNoReplyDrop,
  unsharedConversationScope,
} from '@omadia/channel-sdk';
import type { ChannelUserRef, ConversationMembershipEvent } from '@omadia/channel-sdk';
import { attributeGroupMessage, toSdkConversationType } from './teamsGroupPrimitives.js';
import { pickChatAgentForTurn } from './agentForTurn.js';

/**
 * What resolving an activity yields: the agent that must answer it, and the
 * provisioned-bot key it resolved through (when it resolved through one).
 *
 * Returned rather than stored, because several bots in one group chat receive
 * the SAME message and their turns overlap on this single instance — see the
 * note on `agentResolutionIsPerTurn`.
 */
interface TurnAgentResolution {
  readonly chatAgent?: ChatAgent;
  readonly identityBotKey?: string;
}
import { parseTeamsBotKey, teamsBotKey } from './teamsBotIdentity.js';
import type { TeamsProactiveSend } from './messagesRouter.js';
import type { PrivacyReceipt } from '@omadia/plugin-api';
import type {
  CaptureDisclosure,
  ChatAgent,
  AgentConsultation,
  ConversationHistoryStore,
  DelegatedAnswer,
  FollowUpOption,
  OutgoingAttachment,
  OutgoingChoiceCard,
  OutgoingSlotPicker,
  RunTracePayload,
  TopicDetector,
  TurnContextModule,
  VerifierBadge,
} from './kernel-types.js';
import type {
  PersistedAttachment,
  TeamsAttachmentStore,
} from './teamsAttachmentStore.js';
import {
  aiLabelEntity,
  buildAgentAppsResultCard,
  buildAnswerCard,
  buildChoiceAskCard,
  buildDirectLineOnlyCard,
  buildFollowUpsOnlyCard,
  buildSlotPickerCard,
  buildTopicAskCard,
  parseAgentAppsRecheckValue,
  parseApprovalValue,
  parseBookSlotValue,
  parseChoiceAskValue,
  parseDirectLineValue,
  parseFollowUpValue,
  parseFreshCheckValue,
  parseRoutineCardActionValue,
  parseRoutineListFilterValue,
  parseTopicDecisionValue,
  stripFoldedAiDisclosure,
  type AgentAppsRecheckValue,
  type ApprovalValue,
} from './teamsCard.js';
import type {
  InstallAgentAppsRequest,
  TeamsAgentInstallResult,
} from './teamsAgentInstaller.js';
import { buildRecalledContextCard } from './teamsRecall.js';
import type { TeamsRosterProvider } from './teamsRoster.js';
import type { TeamsConversationObserver } from './teamsConversationObserver.js';
import {
  resolveMentions,
  stripMentionTokens,
  type MentionEntity,
} from './teamsMentions.js';

// Teams accepts long messages but the UI truncates around 28k chars. We leave a safety
// margin and split on paragraph boundaries so tables aren't cut in half.
const TEAMS_MAX_MESSAGE_CHARS = 25_000;

// Frequency of typing indicators during long-running orchestrator calls. The Teams
// client dims the "typing" animation after ~10 s of silence, so we refresh faster.
const TYPING_INTERVAL_MS = 5_000;

/**
 * #575 D7 — the Teams turn's session scope, derived through the channel SDK's
 * typed resolver instead of being hand-built.
 *
 * The bug this closes is measured, not theoretical. The scope used to be
 * `` `teams-${conversation?.id ?? 'unknown'}` ``, so **every** Teams activity
 * arriving without a conversation id landed in the single literal bucket
 * `teams-unknown` — and that string keys conversation history *and* the
 * knowledge-graph partition. Two unrelated callers who both hit that gap shared
 * one memory. It is the `'http-default'` hole again, in production Teams, and
 * it is why `SHARED_SCOPE_TOKENS` lists `teams-unknown` at all.
 *
 * `unsharedConversationScope` gives each such turn its own scope, keyed on the
 * Bot Framework activity id so a **retry of the same message** stays continuous
 * rather than forking a new conversation.
 *
 * ## Why the id is checked here rather than handed over verbatim
 *
 * Passing `` `teams-${id ?? 'unknown'}` `` straight to the resolver looks
 * equivalent and very nearly is: `teams-unknown` is a known shared token, so
 * the SDK neutralises it either way. It differs in exactly one case, and that
 * case is a live hole — an **empty** conversation id builds `teams-`, which is
 * *not* a known shared token, so the resolver classifies it as a perfectly
 * ordinary conversation and every caller with a blank id shares it. Refusing to
 * build a scope from a falsy id is what actually closes that. Verified by
 * mutation: without the check, every other assertion in the scope test still
 * passes.
 *
 * Two properties this deliberately preserves:
 *
 *  - **A present conversation id round-trips byte-identically.** The SDK's
 *    adapter keeps `teams-<id>` an opaque conversation scope precisely so that
 *    introducing the type moves no scope string, and therefore orphans no
 *    existing graph partition.
 *  - **The spelling stays `teams-<id>`, not the canonical `teams::<id>`.**
 *    Re-spelling carries exactly the cost #575 D3 carries — every live Teams
 *    conversation's partition would move — so it is its own migration, not a
 *    rider on an isolation fix.
 *
 * Exported so the behaviour is testable without standing up a Bot Framework
 * adapter; the parameter is structurally typed for the same reason.
 */
export function teamsSessionScope(
  activity: {
    readonly conversation?: { readonly id?: string } | undefined;
    readonly id?: string | undefined;
  },
  /**
   * The provisioned-bot key this turn resolved through (`28:<appId>`), or
   * `undefined` when it did not resolve through one.
   *
   * ## Why the scope has to carry it
   *
   * The scope keys the conversation HISTORY, and the history was shared by
   * every bot in the chat. Each bot therefore received the OTHER bots' replies
   * as its own prior assistant turns — and continued them. Measured, not
   * theorised: a bot that had just joined a chat reported `pool=0` (no memory
   * of its own) and `history=9` in the same turn, and answered "I am Karen",
   * the name of the bot that had been speaking there. The mirror image
   * followed minutes later, with Karen quoting the other bot's scripture.
   *
   * An identity line in the system prompt cannot win that argument: one
   * sentence against nine turns of transcript is not a fair fight, and making
   * it fairer would still leave it a matter of persuasion. Splitting the
   * history is the structural answer — each bot sees what IT said, and nothing
   * it can mistake for its own voice.
   *
   * ## Why only when a provisioned identity resolved
   *
   * Re-spelling a scope moves every live partition behind it (history AND the
   * knowledge-graph tree), so it is not something to do to deployments that do
   * not need it. `undefined` covers exactly those: a single-bot deployment, a
   * binding-routed turn, an unattributable one. All three keep
   * `teams-<conversationId>` byte-identical and lose nothing.
   *
   * A deployment that DOES run several bots per chat pays a one-off: each bot
   * starts fresh in chats it has already been in. That is the intended
   * outcome — the histories it would otherwise inherit are the contaminated
   * ones this fixes.
   */
  identityBotKey?: string,
): string {
  const rawConversationId = activity.conversation?.id;
  // The bot key first, so the conversation id stays the trailing, most
  // variable segment — the shape existing readers already scan for.
  const scope = rawConversationId
    ? identityBotKey
      ? `teams-${identityBotKey}-${rawConversationId}`
      : `teams-${rawConversationId}`
    : undefined;
  return formatSessionScope(
    unsharedConversationScope({
      scope,
      uniqueSuffix: typeof activity.id === 'string' ? activity.id : undefined,
    }),
  );
}

/** Honest confirmation for an approve/reject click, worded by the kernel's resolution outcome so we
 *  never claim "approved → done" when a quorum='all' await still waits, or on a stale/double click. */
function approvalAckText(
  outcome: 'resumed' | 'recorded' | 'already_resolved' | 'not_a_holder',
  approved: boolean,
): string {
  switch (outcome) {
    case 'resumed':
      return approved
        ? '✅ Genehmigt — danke! Der Workflow läuft weiter.'
        : '❌ Abgelehnt — der Workflow wurde entsprechend fortgesetzt.';
    case 'recorded':
      return approved
        ? '🗳️ Deine Zustimmung wurde erfasst — wir warten noch auf weitere Freigeber.'
        : '🗳️ Deine Ablehnung wurde erfasst — wir warten noch auf weitere Freigeber.';
    case 'already_resolved':
      return 'Diese Freigabe wurde bereits bearbeitet.';
    case 'not_a_holder':
      return 'Du bist für diese Freigabe nicht berechtigt.';
  }
}

/**
 * Teams bot that delegates every turn to the orchestrator. No session stickiness —
 * each incoming message becomes a fresh orchestrator chat. Conversation continuity
 * comes from the persistent memory store, not from a stored session.
 */
export class TeamsBot extends TeamsActivityHandler {
  /**
   * THE RESOLVED AGENT IS NO LONGER INSTANCE STATE.
   *
   * It used to be, with the reasoning that "concurrent turns won't trample
   * each other under normal load (one event-loop tick per Bot Framework
   * invocation)". A group chat with several provisioned bots is precisely the
   * load that breaks that: Teams delivers the SAME message to EVERY bot in the
   * chat, so several `handleMessage` calls overlap on this one instance —
   * there is exactly one `TeamsBot` for all of them — and they interleave
   * across every `await`. The second resolution overwrote the first before it
   * was read, and whichever bot resolved last answered for both.
   *
   * That is not a subtle drift. It is impersonation: the reply carried one
   * bot's name and another agent's permissions, and the routing log said the
   * right thing while the turn did the wrong one.
   *
   * So the decision now travels WITH the turn — returned by
   * {@link resolveChatAgentForTurn}, carried in the turn input, read where it
   * is used. Threading one value through is the noise the old comment wanted
   * to avoid; it is much cheaper than the bug.
   */
  private readonly agentResolutionIsPerTurn = true;

  constructor(
    /** Default / legacy chatAgent. Used when no per-Agent resolver returns
     *  anything for this turn (e.g. no binding configured for this Teams
     *  bot yet, or pre-Phase-A boot). */
    private readonly defaultOrchestrator: ChatAgent,
    private readonly history: ConversationHistoryStore,
    private readonly topicDetector: TopicDetector | undefined,
    /**
     * Kernel-owned `turnContext` AsyncLocalStorage module. Injected by the
     * plugin's `activate()` so the Teams adapter and the orchestrator share
     * the same storage instance — critical for `runWithChatParticipants`
     * handing the roster provider down into the orchestrator's own `run()`.
     */
    private readonly turnContext: TurnContextModule,
    private readonly attachmentStore?: TeamsAttachmentStore,
    private readonly rosterProvider?: TeamsRosterProvider,
    /**
     * Bot Framework OAuth Connection Name. When set, every inbound turn
     * attempts a silent `UserTokenClient.getUserToken` lookup; the resulting
     * JWT is threaded as `ssoAssertion` so the calendar tools can OBO-exchange
     * for delegated Graph scopes. Missing connection OR missing token → the
     * assertion stays undefined and calendar tools return `sso_unavailable`.
     */
    private readonly ssoConnectionName?: string,
    /**
     * Tenant id passed into the routines per-turn context. Comes from the
     * kernel's `GRAPH_TENANT_ID` (same key the verifier-store + graph use)
     * so routines stay aligned with the rest of the user-scoped data.
     */
    private readonly tenantId?: string,
    /**
     * Optional kernel hook called once per inbound turn, BEFORE the
     * chatAgent is invoked. The kernel installs the routines per-turn
     * AsyncLocalStorage from this hook so the `manage_routine` tool can
     * attribute `create` to the right user and capture the conversation
     * reference for proactive delivery. Undefined when routines is off.
     */
    private readonly captureRoutineTurn?: (info: {
      tenant: string;
      userId: string;
      /** Operator-addressable id (the user's email) for the Conductor channel-binding key (P2). */
      principalRef?: string;
      conversationRef: unknown;
    }) => void,
    /**
     * Optional kernel-owned dispatcher for routine smart-card actions.
     * Invoked when the bot detects an inbound Activity whose `value`
     * matches `{kind: 'routine.action', action: 'pause'|'delete', id}`.
     * Returns a short user-facing confirmation string the bot renders
     * back into the chat.
     */
    private readonly handleRoutineAction?: (input: {
      action: 'pause' | 'resume' | 'trigger_now' | 'delete';
      id: string;
      /**
       * byte5ai/omadia#1029 — the principal this click belongs to, so the
       * kernel can scope the mutation instead of trusting the card's id.
       *
       * The card path is dispatched out-of-band and never reaches
       * `runOrchestratorTurn`, which is the only place that installs the
       * routines per-turn ALS. So the kernel has no context to read here and
       * falls back to acting UNSCOPED — the card carries the routine id, so a
       * replayed payload would otherwise reach pause/resume/trigger/delete for
       * any row, and `trigger_now` delivers into the routine's own
       * `conversationRef` (a message pushed into someone else's conversation).
       *
       * Same id space as `captureRoutineTurn` on the orchestrator path: the
       * tenant from the kernel's `GRAPH_TENANT_ID`, the user from the shared
       * `userId` derivation in `handleMessage` (`from.aadObjectId`, falling
       * back to `from.id`). Omitted entirely when either half is missing —
       * a half-filled principal is worse than none, because the kernel would
       * scope to it.
       */
      actor?: { tenant: string; userId: string };
    }) => Promise<string>,
    /**
     * Builder for the routine LIST smart card sidecar (rendered after the
     * agent's prose answer when `result.interactive.kind === 'routine_list'`).
     * Kernel-supplied; undefined → list-card path skipped silently.
     */
    private readonly buildRoutineListSmartCardAttachment?: (input: {
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
    }) => { contentType: string; content: unknown },
    /**
     * Phase A+B follow-up — per-turn ChatAgent resolver from the
     * multi-orchestrator runtime. Given the activity's recipient
     * identity (`28:<bot-app-id>` in Teams), the kernel's
     * `channelResolver@1` is asked which Agent owns this channel. When
     * unset (no DATABASE_URL / pre-Phase-A boot / no binding configured)
     * the bot falls back to `defaultOrchestrator` and behaves
     * identically to before. Errors inside the resolver are caught by
     * the caller in plugin.ts; this hook receives a swallow-failures
     * contract.
     */
    /**
     * Resolves a key to a ChatAgent. Returns `decision: 'bound'` when an
     * explicit binding OR the agent's own provisioned identity matches the
     * key — `'fallback'` means "nothing matched, the platform fallback
     * Agent would answer", `'reject'` means "nothing matched and no
     * fallback configured".
     *
     * `exclusive` marks the one case that outranks everything: the key IS
     * this agent's provisioned bot. See {@link pickChatAgentForTurn} for the
     * full precedence order — this callback answers about a single key and
     * has no opinion about which key wins.
     */
    private readonly resolveChatAgentForActivity?: (input: {
      readonly channelType: 'teams';
      readonly channelKey: string;
      readonly conversationId: string;
    }) => {
      readonly decision: 'bound' | 'fallback' | 'reject';
      readonly chatAgent?: ChatAgent;
      readonly exclusive?: boolean;
    },
    /**
     * Observer that records every inbound conversation so the
     * `/operator/channels` dashboard can list known Teams channels /
     * DMs / group chats as bindable targets. Optional — without it,
     * routing falls back to bot-level (`recipient.id`) and the
     * dashboard only sees the catch-all entry.
     */
    private readonly conversationObserver?: TeamsConversationObserver,
    /**
     * US4 (Conductor Surface) — fire-and-forget emitter that surfaces an inbound Teams activity as a
     * Conductor domain event so workflows can trigger on real activity. Supplied by activate() from
     * `ctx.events` (present iff the manifest declares `permissions.events.emit`). Undefined → no emit.
     */
    private readonly emitConductorEvent?: (eventId: string, payload: Record<string, unknown>) => void,
    /**
     * P2c — resolve a user's SMTP email by AAD object id via the M365 Graph (app perm User.Read.All).
     * The reliable email source for the Conductor binding key, especially in 1:1 chats where the
     * conversation roster exposes no member email. Undefined when the M365 integration isn't wired.
     */
    private readonly resolveEmailByAad?: (aadObjectId: string) => Promise<string | null>,
    /**
     * Resolves a Conductor human-await when the user clicks an approve/reject card button —
     * in-process via the kernel's `conductorAwaitResolver` service (no HTTP). `responderId` is the
     * user's email (matches the email-keyed await holder). Undefined when Conductor isn't wired.
     */
    private readonly resolveConductorAwait?: (
      awaitId: string,
      responderId: string,
      approved: boolean,
    ) => Promise<'resumed' | 'recorded' | 'already_resolved' | 'not_a_holder'>,
    /**
     * #330 B2 — group-conversation hooks, present only when the kernel exposes
     * the group primitives (feature-detected in plugin.ts). `capture` feeds the
     * per-conversation reference cache the roster provider resolves through;
     * `emitMembershipEvent` forwards bot-invited / members-changed events
     * (`bot_added` is the Facilitator's explicit, announced entry).
     */
    private readonly groupPrimitives?: {
      captureConversationReference: (context: TurnContext) => void;
      emitMembershipEvent: (event: ConversationMembershipEvent) => void;
    },
    /**
     * #860 W2 (issue #21) — auto-invite seam. Injected by plugin.ts ONLY
     * when `teams_agent_apps[]` is configured AND the `teamsProvisioner@1`
     * service resolved; absent → the whole auto-invite feature is off and
     * `membersAdded` behaves exactly as before. All Graph traffic stays
     * behind `installAgentApps` (the installer → the connector) — the bot
     * itself never talks to Graph for installs.
     */
    private readonly autoInvite?: TeamsAutoInviteDeps,
  ) {
    super();

    this.onMessage(async (context, next) => {
      this.groupPrimitives?.captureConversationReference(context);
      await this.handleMessage(context);
      await next();
    });

    // #330 B2 — membership lifecycle. Runs regardless of groupPrimitives so
    // the roster cache invalidation below stays correct on old kernels too.
    // #860 W2 — the same membersAdded event anchors the auto-invite hook
    // (AFTER the ref capture inside handleMembershipChange).
    this.onMembersAdded(async (context, next) => {
      this.handleMembershipChange(context, 'added');
      await this.maybeRunAutoInvite(context);
      await next();
    });
    this.onMembersRemoved(async (context, next) => {
      this.handleMembershipChange(context, 'removed');
      await next();
    });
  }

  /** #330 B2 — translate a Bot-Framework conversationUpdate into the SDK's
   *  membership event, invalidate the roster cache, and (when the added
   *  members include this bot) emit `bot_added` with the inviter. */
  private handleMembershipChange(context: TurnContext, change: 'added' | 'removed'): void {
    try {
      const activity = context.activity;
      const conversationId = activity.conversation?.id;
      if (!conversationId) return;
      this.rosterProvider?.invalidate(conversationId);
      this.groupPrimitives?.captureConversationReference(context);
      if (!this.groupPrimitives) return;

      const accounts = (change === 'added' ? activity.membersAdded : activity.membersRemoved) ?? [];
      const members: ChannelUserRef[] = accounts.map((m) => ({
        kind: 'teams-aad',
        id: m.aadObjectId ?? m.id,
        ...(m.name ? { displayName: m.name } : {}),
      }));
      const botAdded = change === 'added' && accounts.some((m) => m.id === activity.recipient?.id);
      const inviter: ChannelUserRef | undefined =
        botAdded && activity.from?.id && activity.from.id !== activity.recipient?.id
          ? {
              kind: 'teams-aad',
              id: activity.from.aadObjectId ?? activity.from.id,
              ...(activity.from.name ? { displayName: activity.from.name } : {}),
            }
          : undefined;

      const conversationType = toSdkConversationType(activity.conversation?.conversationType);
      // #860 W2 — intro suppression: when THIS bot's own membersAdded
      // correlates with a just-run auto-install (the team's marker is
      // fresh), it is one of the freshly installed agent apps. Downgrade
      // `bot_added` to `members_added` so the announced entry (the default
      // intro) fires only for the bot that ran the installer — a flock
      // install must not post N identical intros.
      const suppressIntro =
        botAdded &&
        shouldSuppressAutoInstallIntro(
          this.autoInvite &&
            ((teamId: string) => this.autoInvite!.probeAutoInstallMarker(teamId)),
          activity,
        );
      this.groupPrimitives.emitMembershipEvent({
        kind:
          botAdded && !suppressIntro
            ? 'bot_added'
            : change === 'added'
              ? 'members_added'
              : 'members_removed',
        channelId: 'de.byte5.channel.teams',
        channelType: 'teams',
        conversationId,
        ...(conversationType ? { conversationType } : {}),
        members,
        ...(inviter ? { addedBy: inviter } : {}),
        occurredAt:
          activity.timestamp instanceof Date ? activity.timestamp.toISOString() : new Date().toISOString(),
        rawEvent: activity,
      });
    } catch (err) {
      // A membership event must never break the conversationUpdate turn.
      console.error('[teams] membership event handling failed:', err instanceof Error ? err.message : err);
    }
  }

  /**
   * #860 W2 (issue #21) — onboarding hook: when THIS bot was just added to
   * a team, run the agent-app installer for that team and post the result /
   * fallback card. Anchored on the same `membersAdded` event that captures
   * the conversation reference. No-op unless plugin.ts injected the
   * auto-invite seam. When the team's auto-install marker is fresh, this
   * bot IS one of the freshly installed agent apps — it stays silent
   * (no second installer run, no second summary card).
   */
  private async maybeRunAutoInvite(context: TurnContext): Promise<void> {
    if (!this.autoInvite) return;
    try {
      const outcome = await runTeamsAutoInviteHook(this.autoInvite, context);
      if (outcome !== 'skipped') {
        const conversationIdShort =
          (context.activity.conversation?.id ?? '-').slice(0, 24);
        console.error(
          `[teams] auto-invite hook ${outcome} conv=${conversationIdShort}…`,
        );
      }
    } catch (err) {
      // Auto-invite must never break the conversationUpdate turn.
      console.error(
        '[teams] auto-invite hook failed:',
        err instanceof Error ? err.message : err,
      );
    }
  }

  /**
   * userId (AAD object id, or the 29:-channel id) → { resolved SMTP email or null, when fetched }.
   * The Conductor channel-binding key must be the operator-addressable id — the user's real email,
   * NOT the UPN (which can differ from the primary SMTP, so binding by it would silently miss). Email
   * isn't on the inbound activity, so we resolve it (Graph first, roster fallback) and cache BOTH hits
   * AND misses (a `null` value) with a TTL, so an unresolvable user doesn't re-hit Graph every turn.
   * P2 identity-bridge.
   */
  private readonly emailByUser = new Map<string, { value: string | null; fetchedAt: number }>();

  /** conversationId → last `bot_present` emit (epoch ms). Re-emitted after
   *  the interval so the kernel's 24h eligibility TTL cannot outlive a
   *  long-running process's only emit. */
  private readonly botPresentEmittedAt = new Map<string, number>();
  private static readonly BOT_PRESENT_REEMIT_MS = 12 * 60 * 60 * 1000;

  private maybeEmitBotPresent(context: TurnContext): void {
    try {
      if (!this.groupPrimitives) return;
      const activity = context.activity;
      const conversationId = activity.conversation?.id;
      if (!conversationId) return;
      if (toSdkConversationType(activity.conversation?.conversationType) !== 'group') return;
      const now = Date.now();
      const last = this.botPresentEmittedAt.get(conversationId);
      if (last !== undefined && now - last < TeamsBot.BOT_PRESENT_REEMIT_MS) return;
      this.botPresentEmittedAt.set(conversationId, now);
      const sender: ChannelUserRef | undefined = activity.from?.id
        ? {
            kind: 'teams-aad',
            id: activity.from.aadObjectId ?? activity.from.id,
            ...(activity.from.name ? { displayName: activity.from.name } : {}),
          }
        : undefined;
      this.groupPrimitives.emitMembershipEvent({
        kind: 'bot_present',
        channelId: 'de.byte5.channel.teams',
        channelType: 'teams',
        conversationId,
        conversationType: 'group',
        members: activity.recipient?.id
          ? [{ kind: 'teams-aad', id: activity.recipient.id, ...(activity.recipient.name ? { displayName: activity.recipient.name } : {}) }]
          : [],
        ...(sender ? { addedBy: sender } : {}),
        occurredAt: activity.timestamp instanceof Date ? activity.timestamp.toISOString() : new Date().toISOString(),
      });
    } catch (err) {
      // Eligibility signalling must never break the message turn.
      console.error('[teams] bot_present emit failed:', err instanceof Error ? err.message : err);
    }
  }

  /**
   * Resolve this user's SMTP email for the Conductor binding key — M365 Graph first (reliable in 1:1
   * AND group), conversation roster as fallback. Awaited BEFORE the turn's `captureRoutineTurn`, so
   * the FIRST turn already binds by email (no stranded AAD-keyed binding). Returns undefined — leaving
   * the binding channel-native-keyed, with a log line — when no SMTP email is resolvable (we never
   * bind a UPN guess). Both hits and misses are cached (hits 30 min, misses 5 min), so the awaited
   * Graph call runs at most once per user per window, never per turn.
   */
  private async resolveBindingEmail(context: TurnContext, userId: string): Promise<string | undefined> {
    const POSITIVE_TTL_MS = 30 * 60 * 1000;
    const NEGATIVE_TTL_MS = 5 * 60 * 1000;
    const cached = this.emailByUser.get(userId);
    if (cached && Date.now() - cached.fetchedAt < (cached.value ? POSITIVE_TTL_MS : NEGATIVE_TTL_MS)) {
      return cached.value ?? undefined;
    }
    // 1) M365 Graph by AAD object id (app perm User.Read.All) — reliable in BOTH 1:1 and group, unlike
    // the conversation roster (no member email in personal scope). Only when userId is a canonical AAD
    // object id (UUID); the channel-native `29:`/`28:` ids would just 404.
    if (this.resolveEmailByAad && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(userId)) {
      try {
        const mail = (await this.resolveEmailByAad(userId))?.trim().toLowerCase();
        if (mail) {
          this.emailByUser.set(userId, { value: mail, fetchedAt: Date.now() });
          return mail;
        }
      } catch (err) {
        console.error('[teams] conductor binding graph email resolve failed:', err);
      }
    }
    // 2) Fallback: the channel roster (exposes member email in group chats).
    if (this.rosterProvider) {
      try {
        const participants = await this.rosterProvider.list(context);
        const match = participants.find((p) => p.aadObjectId === userId || p.channelUserId === userId);
        const email = match?.email?.trim().toLowerCase();
        if (email) {
          this.emailByUser.set(userId, { value: email, fetchedAt: Date.now() });
          return email;
        }
      } catch (err) {
        console.error('[teams] conductor binding email resolve failed:', err);
      }
    }
    // Unresolved after Graph + roster. Keep serving a prior email if we had one; otherwise negative-cache
    // (so we don't re-resolve every turn) and log ONCE so this silent-miss class is diagnosable.
    if (cached?.value) return cached.value;
    console.error(
      `[teams] conductor binding: email unresolved after Graph+roster for user=${userId} — binding stays channel-native-keyed (email-addressed reminders may miss; check the User.Read.All grant)`,
    );
    this.emailByUser.set(userId, { value: null, fetchedAt: Date.now() });
    return undefined;
  }

  /**
   * Resolve the ChatAgent for THIS turn based on the bot framework
   * activity. Sets `this.currentChatAgent` so the helpers downstream
   * (which all read it via `this.currentChatAgent ?? this.defaultOrchestrator`)
   * see the same agent for the whole turn even if the resolver later
   * changes (binding edit mid-turn). Called once at the top of
   * handleMessage.
   *
   * TWO keys are probed, and probing order is deliberately NOT precedence
   * order. The bot identity (`activity.recipient.id`, `28:<bot-app-id>`) goes
   * first because it is the only key that can come back exclusive — a bot
   * the platform provisioned to BE one agent. Failing that, the conversation
   * binding wins, then the bot-level catch-all, then the platform fallback.
   * {@link pickChatAgentForTurn} owns that rule and its tests.
   *
   * Both keys are matched by exact string equality, so they are normalized
   * through `teamsBotKey()` — the same helper the `ChannelKeyDirectory`
   * contribution publishes with, so what the operator picked in
   * `/operator/channels` IS what the resolver matches at runtime.
   */
  private resolveChatAgentForTurn(
    activity: TurnContext['activity'],
  ): TurnAgentResolution {
    if (!this.resolveChatAgentForActivity) return {};
    const conversationId = activity.conversation?.id;
    // Normalize the bot-identity key through the canonical helper: the
    // directory publishes `teamsBotKey(appId)` (lowercase-normalized) and the
    // platform projects provisioned identities the same way, so the runtime
    // lookup must use the exact same normalization or a mixed-casing
    // `recipient.id` could miss both the binding and the identity.
    const rawRecipientId =
      typeof activity.recipient?.id === 'string'
        ? activity.recipient.id
        : undefined;
    const recipientAppId =
      rawRecipientId !== undefined ? parseTeamsBotKey(rawRecipientId) : undefined;
    const recipientId =
      recipientAppId !== undefined ? teamsBotKey(recipientAppId) : rawRecipientId;
    if (!conversationId && !recipientId) return {};
    const resolveForActivity = this.resolveChatAgentForActivity;
    // Local to this call — nothing about this turn touches the instance.
    let identityBotKey: string | undefined;
    try {
      const chatAgent = pickChatAgentForTurn({
        ...(conversationId ? { conversationId } : {}),
        ...(recipientId ? { botKey: recipientId } : {}),
        resolve: (channelKey) => {
          const decision = resolveForActivity({
            channelType: 'teams',
            channelKey,
            conversationId: conversationId ?? 'unknown',
          });
          // `exclusive` means the key IS this agent's own provisioned bot —
          // the platform's own word for "this deployment runs per-bot
          // agents". That is exactly the condition under which the shared
          // conversation history has to be split, and the only one: a legacy
          // single-bot deployment never sees it and keeps its scopes byte-
          // identical. Recorded here rather than derived later because this
          // is the one place the platform's verdict is visible.
          if (decision.exclusive === true && channelKey === recipientId) {
            identityBotKey = channelKey;
          }
          return decision;
        },
      });
      return {
        ...(chatAgent ? { chatAgent } : {}),
        ...(identityBotKey ? { identityBotKey } : {}),
      };
    } catch (err) {
      console.error(
        `[teams] channelResolver threw (conv=${conversationId?.slice(0, 16) ?? '-'}, recipient=${recipientId?.slice(0, 16) ?? '-'}…) — falling back to default agent:`,
        err,
      );
      return {};
    }
  }

  /**
   * Whether this inbound activity explicitly @mentions our bot.
   *
   * Teams enforces "channel posts require @mention" at platform level
   * (channel-scoped activities only flow to bots that are tagged). Group
   * chats are different: depending on tenant policy the bot can receive
   * every message in the chat once it has been added. Personal 1:1 chats
   * never require a mention.
   *
   * We treat "not mentioned + not personal" as silence — the bot drops
   * the turn without responding so it doesn't accidentally talk in group
   * chats where it was added but the current user wasn't addressing it.
   * The observer still captures the conversation (the directory should
   * surface it as a bindable target regardless of whether THIS message
   * was addressed at us).
   */
  private isMentioned(activity: TurnContext['activity']): boolean {
    const recipientId = activity.recipient?.id;
    if (!recipientId) return false;
    const entities = activity.entities ?? [];
    for (const entity of entities) {
      if (entity.type !== 'mention') continue;
      const mentioned = (entity as { mentioned?: { id?: string } }).mentioned;
      if (mentioned?.id === recipientId) return true;
    }
    return false;
  }

  private async handleMessage(context: TurnContext): Promise<void> {
    // Record this conversation so /operator/channels can list it as a
    // bindable target — idempotent + cheap. Done BEFORE the mention
    // filter so operators can pre-bind a chat even if nobody has
    // @-mentioned the bot yet (a "Mention me later" pre-config flow).
    this.conversationObserver?.observe(context);
    // Phase A+B follow-up — pick the ChatAgent for THIS turn before any
    // chat work runs. Falls back to defaultOrchestrator on miss / error.
    const turnAgent = this.resolveChatAgentForTurn(context.activity);

    // Mention-only policy: in any non-personal context (group chat,
    // channel thread, meeting chat) drop the turn unless the bot was
    // explicitly @mentioned. Otherwise the bot would chime in on every
    // message of a group chat the moment it joins — noisy + unwanted.
    const conversationType =
      context.activity.conversation?.conversationType ?? 'unknown';
    const isPersonal = conversationType === 'personal';
    // Action.Submit card interactions (follow-up buttons, choice/slot/routine
    // cards) always carry a `value` payload and are unambiguously directed at
    // THIS bot — Teams routes them to us regardless of an @-mention, and a card
    // click cannot inject a mention entity (messageBack has no entities). Treat
    // them as bot-directed so channel follow-ups actually trigger a turn;
    // otherwise the mention-gate silently drops every button click in a channel.
    const isCardAction =
      typeof context.activity.value === 'object' &&
      context.activity.value !== null;
    if (!isPersonal && !isCardAction && !this.isMentioned(context.activity)) {
      const conversationIdShort =
        (context.activity.conversation?.id ?? '-').slice(0, 24);
      console.error(
        `[teams] drop non-mention conv=${conversationIdShort}… type=${conversationType}`,
      );
      return;
    }

    // `conversationId` deliberately keeps its old `'unknown'` fallback — it is
    // read by the attachment store (`conversationId !== 'unknown'`) and a dozen
    // log lines. Only the SCOPE changes; the reasoning is on `teamsSessionScope`.
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    // Bot-qualified in a multi-bot deployment — see `teamsSessionScope`. The
    // agent was resolved a few lines up, so the key is already known here.
    const sessionScope = teamsSessionScope(
      context.activity,
      turnAgent.identityBotKey,
    );
    const from = context.activity.from;
    const userId =
      from?.aadObjectId ??
      (typeof from?.id === 'string' ? from.id : undefined);
    // Hard entry-point log — stderr so Fly can't silently drop it. Tells
    // us a Teams message actually reached the bot process, independent
    // of every downstream log that might get filtered under load.
    console.error(
      `[teams] inbound conv=${conversationId} aad=${from?.aadObjectId ?? '-'} type=${context.activity.type ?? '-'} hasText=${String(Boolean(context.activity.text))} hasAttachments=${String((context.activity.attachments?.length ?? 0) > 0)}`,
    );
    // Phase-B+ diagnostic — print the full Teams routing context once per
    // inbound so we can tell genuine "same conversation" from look-alike
    // contexts (group-chat vs. channel-thread vs. meeting-chat, same vs.
    // different recipient bot id). Helps debug the operator-channels
    // dashboard when a conv-id appears stable across what the operator
    // believes are different chats.
    {
      const conv = context.activity.conversation;
      const recipient = context.activity.recipient;
      const channelData = context.activity.channelData as
        | {
            team?: { id?: string; name?: string };
            channel?: { id?: string; name?: string };
            tenant?: { id?: string };
          }
        | undefined;
      console.error(
        `[teams] inbound-meta conv=${conversationId} ` +
          `conversationType=${conv?.conversationType ?? '-'} ` +
          `convName=${conv?.name ?? '-'} ` +
          `team=${channelData?.team?.id?.slice(0, 12) ?? '-'}(${channelData?.team?.name ?? '-'}) ` +
          `channel=${channelData?.channel?.id?.slice(0, 12) ?? '-'}(${channelData?.channel?.name ?? '-'}) ` +
          `tenant=${channelData?.tenant?.id?.slice(0, 12) ?? '-'} ` +
          `recipient=${recipient?.id?.slice(0, 16) ?? '-'}(${recipient?.name ?? '-'})`,
      );
    }

    // Persist any file / image attachments BEFORE delegating to the
    // orchestrator — we want to thread the stored metadata (storage key +
    // signed URL + filename) into the user message so Claude can decide
    // whether to treat it as a brand asset (Logo → /memories/_brand/logo.md)
    // or reference it in a later diagram. Failures are swallowed by the
    // store itself and yield an empty list, so the bot still replies.
    const persistedAttachments = await this.persistAttachmentsIfAny(
      context,
      conversationId,
      userId,
    );

    // Adaptive-card button click: the activity carries a structured `value`.
    // Resolve any pending topic decision *before* we treat it like a normal
    // text message, otherwise the user would see "Ich habe keinen Text…".
    const decision = parseTopicDecisionValue(context.activity.value);
    if (decision) {
      await this.handleTopicDecision(context, sessionScope, userId, decision);
      return;
    }

    // "🔄 Fresh Check" — user clicked the button on a previous answer card
    // to re-run the same question while bypassing memory + context block.
    const freshCheck = parseFreshCheckValue(context.activity.value);
    if (freshCheck) {
      await this.handleFreshCheck(context, sessionScope, userId, freshCheck.originalMessage);
      return;
    }

    // `ask_user_choice` option click — blocking clarification resolved.
    // The chosen `value` is injected as the next user turn so the LLM can
    // re-run with the disambiguated input.
    const choiceAsk = parseChoiceAskValue(context.activity.value);
    if (choiceAsk) {
      await this.handleChoiceAsk(context, sessionScope, userId, choiceAsk);
      return;
    }

    // `suggest_follow_ups` refinement click — non-blocking. The full
    // `prompt` (not just the label) is sent as the next user turn; the
    // LLM treats it as a fresh self-contained question.
    const followUp = parseFollowUpValue(context.activity.value);
    if (followUp) {
      await this.handleFollowUp(context, sessionScope, userId, followUp);
      return;
    }

    // Slot-picker click from `find_free_slots` sidecar card. The synthetic
    // user message carries the opaque slotId so the LLM can call
    // `book_meeting({slotId, subject})` on the next turn with the subject
    // inferred from prior conversation context.
    const bookSlot = parseBookSlotValue(context.activity.value);
    if (bookSlot) {
      await this.handleBookSlotClick(context, sessionScope, userId, bookSlot);
      return;
    }

    // Routine smart-card click (Pausieren / Aktivieren / Löschen — either
    // on a cron-triggered delivery or on a routine row inside the list
    // smart-card). Dispatched out-of-band — never reaches the orchestrator
    // because the action is purely lifecycle, not a conversation turn.
    const routineAction = parseRoutineCardActionValue(context.activity.value);
    if (routineAction && this.handleRoutineAction) {
      try {
        // #1029 — hand the principal along. `userId` is the value this same
        // function already computed for every other branch, and the one the
        // orchestrator path passes to `captureRoutineTurn`, so both doors onto
        // `manage_routine`'s mutations agree on who the caller is. Both halves
        // or neither: an `actor` missing its tenant would scope to a partial
        // principal, which is worse than the documented unscoped fallback.
        const actor =
          this.tenantId && userId
            ? { actor: { tenant: this.tenantId, userId } }
            : {};
        const ack = await this.handleRoutineAction({
          action: routineAction.action,
          id: routineAction.id,
          ...actor,
        });
        await context.sendActivity(ack);
      } catch (err) {
        const detail = err instanceof Error ? err.message : String(err);
        const verb =
          routineAction.action === 'pause'
            ? 'pausieren'
            : routineAction.action === 'resume'
              ? 'aktivieren'
              : routineAction.action === 'trigger_now'
                ? 'jetzt auslösen'
                : 'löschen';
        await context.sendActivity(
          `Konnte die Routine nicht ${verb}: ${detail}`,
        );
      }
      return;
    }

    // Conductor approve/reject button click (from a proactive human-await reminder
    // card). Out-of-band like the routine action — resolves the await in-process via
    // the kernel callback; never reaches the orchestrator (it's a decision, not a turn).
    const approval = parseApprovalValue(context.activity.value);
    if (approval) {
      await this.handleApprovalDecision(context, approval);
      return;
    }

    // #860 W2 (issue #21) — "🔄 Prüfen" click on the auto-invite result
    // card: re-run the installer and update the card in place (fresh post
    // when the update is rejected). Out-of-band like the routine action —
    // a lifecycle decision, never an orchestrator turn.
    const agentAppsRecheck = parseAgentAppsRecheckValue(context.activity.value);
    if (agentAppsRecheck) {
      if (!this.autoInvite) {
        await context.sendActivity(
          'Auto-Invite ist auf dieser Installation nicht (mehr) konfiguriert — bitte `teams_agent_apps` und den M365-Connector prüfen.',
        );
        return;
      }
      try {
        await handleTeamsAgentAppsRecheck(
          this.autoInvite,
          context,
          agentAppsRecheck,
        );
      } catch (err) {
        console.error(
          '[teams] agent-apps re-check failed:',
          err instanceof Error ? err.message : err,
        );
        await context.sendActivity(
          'Die Prüfung ist fehlgeschlagen — bitte später erneut versuchen.',
        );
      }
      return;
    }

    // Filter-pill click on the routine list smart-card. Synthesize a
    // German user message that the agent recognises and dispatches via
    // `manage_routine.list` with the chosen filter. Same pattern as the
    // follow-up / choice-ask buttons: the synthetic user-turn is
    // chronologically visible in the transcript so the user sees the
    // intent was registered.
    const routineFilter = parseRoutineListFilterValue(context.activity.value);
    if (routineFilter) {
      const phrase =
        routineFilter.filter === 'all'
          ? 'Zeig mir alle meine Routinen.'
          : routineFilter.filter === 'active'
            ? 'Zeig mir nur die aktiven Routinen.'
            : 'Zeig mir nur die pausierten Routinen.';
      this.history.clearPending(sessionScope);
      const priorTurns = this.history.get(sessionScope);
      await this.runOrchestratorTurnDetached(context, {
        conversationId,
        sessionScope,
        userId,
        userMessage: phrase,
        priorTurns,
      });
      return;
    }

    // #332 — a "💬 Direkt mit <Agent>" button click carries no text, only a
    // submit value. Synthesize the Direct-Line directive (`#<token> <message>`)
    // as the user message so it flows through the normal turn (history,
    // privacy, persistence) and the core dispatcher routes it straight to the
    // named specialist. Bypasses extractUserMessage's mention-strip — the
    // directive is already clean.
    const directLineSubmit = parseDirectLineValue(context.activity.value);
    const rawUserMessage = directLineSubmit
      ? `#${directLineSubmit.token} ${directLineSubmit.originalMessage}`
      : extractUserMessage(context);

    // Magic-code intercept for the bypass sign-in flow. When the user pastes
    // the 6-digit code from the BF token-service page into the chat, Bot
    // Framework does NOT intercept it automatically (that would require the
    // stock OAuthCard we deliberately sidestepped). We call `getUserToken`
    // with the code as the magic-code param — success stores the token
    // against user+connection so subsequent `extractSsoAssertion` calls
    // succeed silently.
    const channelUserId = context.activity.from?.id;
    if (this.ssoConnectionName && /^\d{6}$/.test(rawUserMessage) && channelUserId) {
      const exchanged = await this.tryExchangeMagicCode(
        context,
        channelUserId,
        rawUserMessage,
      );
      if (exchanged) {
        console.error(
          `[teams] magic-code exchange ok conv=${conversationId} user=${userId}`,
        );
        await context.sendActivity(
          'Kalender-Zugriff freigegeben ✓ — stell deine Terminfrage einfach nochmal, dann buche ich direkt.',
        );
        return;
      }
      console.error(
        `[teams] magic-code exchange failed conv=${conversationId} — continuing as normal message`,
      );
    }
    // Empty text + no attachment → ask for a real question. Empty text
    // WITH attachments is treated as "hier ein File, merk dir das" — we
    // let the orchestrator see the attachment-hint and decide how to
    // react (usually: store as brand asset + short acknowledgement).
    if (rawUserMessage.length === 0 && persistedAttachments.length === 0) {
      await context.sendActivity(
        'Ich habe keinen Text in deiner Nachricht gefunden. Schreib mir ruhig eine Frage — z. B. "Wer sind unsere umsatzstärksten Kunden?".',
      );
      return;
    }
    const attachmentHint = this.buildAttachmentHint(persistedAttachments);
    const userMessage = attachmentHint
      ? `${rawUserMessage || '(keine Texteingabe — nur eine Datei)'}${attachmentHint}`
      : rawUserMessage;

    // A fresh typed message while a topic-ask was pending resolves that
    // ambiguity implicitly — the user didn't click the buttons, they moved
    // on. We clear the pending flag so the bot doesn't keep asking.
    this.history.clearPending(sessionScope);

    // Topic detection runs before the orchestrator so we don't pay a full
    // agent turn just to hand the user a clarification card. The detector
    // is optional — if unavailable (sidecar down, config off), we fall
    // through to the previous behaviour.
    const priorTurns = this.history.get(sessionScope);
    let effectiveHistory = priorTurns;
    if (this.topicDetector && priorTurns.length > 0) {
      try {
        const verdict = await this.topicDetector.classify({
          userMessage,
          history: priorTurns,
        });
        // stderr so Fly's log aggregator doesn't silently drop it — this is
        // the only signal that tells us which branch fired per turn, and we
        // saw other stdout-INFO lines vanish under load.
        console.error(
          `[topic] decision=${verdict.decision} reason=${verdict.reason}` +
            (verdict.similarity !== undefined
              ? ` sim=${verdict.similarity.toFixed(3)}`
              : '') +
            (verdict.classifier ? ` cls=${verdict.classifier}` : '') +
            ` conv=${conversationId}`,
        );
        if (verdict.decision === 'reset') {
          this.history.resetTurns(sessionScope);
          effectiveHistory = [];
        } else if (verdict.decision === 'ask') {
          const lastTurn = priorTurns.at(-1);
          const summary = lastTurn
            ? `Deine letzte Frage: "${truncate(lastTurn.userMessage, 180)}"`
            : 'Der vorherige Chat-Kontext liegt vor.';
          await sendTopicAsk(context, summary, userMessage);
          this.history.markPending(sessionScope, {
            userMessage,
            askedAt: Date.now(),
          });
          console.error(
            `[topic] asked user conv=${conversationId} msg="${truncate(userMessage, 60)}"`,
          );
          return;
        }
      } catch (err) {
        console.error(
          '[topic] classify failed — continuing with full history:',
          err instanceof Error ? err.message : err,
        );
      }
    }

    // #330 field report round 3 — a bot that is ALREADY a group member never
    // gets another conversationUpdate, so the kernel's invite index (the
    // facilitation auto-bind scope guard) stayed closed for exactly the chats
    // people talk to the bot in; two live facilitation attempts died on this.
    // An inbound GROUP message is transport-verified proof of membership, so
    // emit ONE synthetic `bot_present` per conversation per interval —
    // ELIGIBILITY only, consumers must not eagerly bind on it. The verified
    // sender rides along as addedBy (the person engaging the bot).
    this.maybeEmitBotPresent(context);

    // US4 (Conductor Surface) — surface a GENUINE user message as a Conductor domain event so a
    // workflow can trigger on real Teams activity. Placed on the main user-text path: it runs ONLY
    // after the card-action handlers + the SSO magic-code intercept have early-returned, so it never
    // leaks a one-time sign-in code (H3) and never fires for button clicks (M1), and it emits the
    // CLEANED message text (M3). Fire-and-forget — the wrapper swallows sync throws + async
    // rejections, so a workflow trigger can never delay or break the turn.
    if (this.emitConductorEvent) {
      const mentioned = this.isMentioned(context.activity);
      const channelData = context.activity.channelData as { tenant?: { id?: string } } | undefined;
      const eventPayload: Record<string, unknown> = {
        conversationId,
        activityId: context.activity.id ?? null, // idempotency key — Bot Framework redelivers on 5xx/timeout
        userId: userId ?? null,
        userName: from?.name ?? null,
        text: userMessage,
        mentioned,
        // WHICH BOT THE PERSON ADDRESSED — the same `28:<appId>` key routing
        // resolves an Agent from (`activity.recipient.id`).
        //
        // Without it a workflow triggered by this event cannot be bound to the
        // permissions of the bot that was actually spoken to. That is not
        // hypothetical: a run configured to execute as the platform fallback
        // agent — which is granted every installed plugin — answered a message
        // addressed to an agent holding no grants at all, under that agent's
        // name and avatar. The kernel now refuses such a run when it cannot
        // attribute it, so this field is what lets a legitimate workflow keep
        // working rather than being refused.
        //
        // Null rather than omitted when the activity carries no recipient, so
        // a consumer reads "no bot identified" as data instead of inferring it
        // from an absent key.
        botId: context.activity.recipient?.id ?? null,
        // The inbound AAD tenant (not the kernel graph tenant) so workflows can route/isolate by tenant.
        tenantId: channelData?.tenant?.id ?? context.activity.conversation?.tenantId ?? null,
      };
      this.emitConductorEvent('teams.message.posted', eventPayload);
      if (mentioned) this.emitConductorEvent('teams.mention', eventPayload);
    }
    await this.runOrchestratorTurn(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage,
      priorTurns: effectiveHistory,
      ...(turnAgent.chatAgent ? { chatAgent: turnAgent.chatAgent } : {}),
    });
  }

  /**
   * Downloads + persists attached files to Tigris + Neon, then returns the
   * list of successfully stored metadata records. Swallows all errors
   * internally (they land in the store's logger) and returns `[]` — the
   * Teams reply continues in every case.
   *
   * We await this (rather than fire-and-forget) so the caller can thread
   * the metadata into the orchestrator's user message. Typical logo upload
   * is sub-second; 25 MB cap + 30 s download timeout bounds the worst case.
   */
  private async persistAttachmentsIfAny(
    context: TurnContext,
    conversationId: string,
    userId: string | undefined,
  ): Promise<PersistedAttachment[]> {
    const activity = context.activity;
    const originalAttachments = activity.attachments ?? [];
    const turnTime = activity.timestamp
      ? new Date(activity.timestamp)
      : undefined;

    // Channel messages often put the file reference OUTSIDE `attachments[]`
    // — in `channelData`, `entities[]`, or as a bare URL in the message text.
    // Probe all of them, synthesise pseudo-attachments so the store's
    // existing resolver + Graph-download path can handle them uniformly.
    const synthesised = this.extractOutOfBandAttachments(activity);

    // Group-chat fallback: Teams does NOT forward file attachments to
    // the bot Activity even with RSC `ChatMessage.Read.Chat` granted
    // (empirical observation, 2026-04-20). We use the same RSC permission
    // to pull the message via Graph, which *does* surface attachments.
    // Runs only when we have message + chat ids AND the store has a
    // graph client configured.
    const messageId =
      typeof activity.id === 'string' && activity.id.length > 0
        ? activity.id
        : undefined;
    let graphSynthesised: Array<{
      contentType: string;
      contentUrl: string;
      name: string;
    }> = [];
    if (this.attachmentStore && messageId && conversationId !== 'unknown') {
      graphSynthesised = await this.attachmentStore.discoverFromGraphMessage(
        conversationId,
        messageId,
      );
    }

    // Always log, even on empty — the 0-case tells us "a message arrived
    // that we could NOT find a file reference in", which is exactly what
    // we need to triage channel upload failures.
    const channelDataKeys = activity.channelData
      ? Object.keys(activity.channelData as Record<string, unknown>)
          .slice(0, 8)
          .join(',')
      : '-';
    const entityTypes = Array.isArray(activity.entities)
      ? activity.entities
          .map((e) => String(e?.type ?? '?'))
          .slice(0, 8)
          .join(',')
      : '-';
    // stderr: Fly's log aggregator drops stdout INFO under load; this
    // diag line is the only signal that tells us what shape the inbound
    // Teams activity has, so it MUST make it through. Logged
    // unconditionally (even when attachmentStore is disabled) so we can
    // triage inbound shape independently of the store wiring.
    console.error(
      `[teams] activity conv=${conversationId} attachments=${String(originalAttachments.length)} entities=${String(activity.entities?.length ?? 0)} entityTypes=${entityTypes} channelDataKeys=${channelDataKeys} synthesised=${String(synthesised.length)} graphSynthesised=${String(graphSynthesised.length)} textLen=${String(activity.text?.length ?? 0)} storeReady=${String(Boolean(this.attachmentStore))}`,
    );

    if (!this.attachmentStore) return [];

    const allAttachments = [
      ...originalAttachments,
      ...synthesised,
      ...graphSynthesised,
    ];
    if (allAttachments.length === 0) return [];

    try {
      return await this.attachmentStore.persistTurn({
        conversationId,
        ...(userId ? { userId } : {}),
        ...(turnTime ? { turnTime } : {}),
        attachments: allAttachments,
      });
    } catch (err) {
      console.error(
        '[teams] attachment persistence threw (swallowed):',
        err instanceof Error ? err.message : err,
      );
      return [];
    }
  }

  /**
   * Gather file references that aren't in `activity.attachments[]`:
   *   1. SharePoint / OneDrive URLs in the message text (plain text or
   *      extracted from HTML `<a href>` / `<img src>`).
   *   2. Likely the same URLs surfaced in `entities[]`.
   *   3. Teams-specific `channelData.attachments` when present.
   *
   * Each hit becomes a minimal `Attachment`-shaped object with
   * `contentUrl = <sharing-url>`, so the store's existing SharePoint
   * resolver can route it through the Graph client.
   */
  private extractOutOfBandAttachments(activity: {
    text?: string;
    entities?: unknown[];
    channelData?: unknown;
  }): Array<{ contentType: string; contentUrl: string; name: string }> {
    const out: Array<{ contentType: string; contentUrl: string; name: string }> = [];
    const seen = new Set<string>();

    const push = (url: string, nameHint?: string): void => {
      if (seen.has(url)) return;
      if (!isSharePointLike(url)) return;
      seen.add(url);
      const name = nameHint && nameHint.trim().length > 0
        ? nameHint.trim()
        : fileNameFromUrl(url) ?? 'sharepoint-file';
      out.push({ contentType: '', contentUrl: url, name });
    };

    // 1) scan message.text (Teams sometimes delivers HTML, sometimes plain)
    const text = activity.text ?? '';
    if (text) {
      for (const m of text.matchAll(/https?:\/\/[^\s"'<>]+/g)) {
        const url = m[0];
        push(url);
      }
    }

    // 2) scan channelData (shape varies; dump generically)
    if (activity.channelData && typeof activity.channelData === 'object') {
      for (const url of collectStringsDeep(activity.channelData)) {
        if (typeof url === 'string') push(url);
      }
    }

    // 3) entities[]
    if (Array.isArray(activity.entities)) {
      for (const ent of activity.entities) {
        if (ent && typeof ent === 'object') {
          for (const url of collectStringsDeep(ent)) {
            if (typeof url === 'string') push(url);
          }
        }
      }
    }

    return out;
  }

  /**
   * Produce a short system-style block that the orchestrator gets appended
   * to the user's message. Tells Claude exactly which files arrived in
   * this turn, their signed URLs (if any), and the memory convention for
   * promoting them to persistent brand assets. Returns '' when nothing to
   * announce.
   */
  private buildAttachmentHint(
    persisted: readonly PersistedAttachment[],
  ): string {
    if (persisted.length === 0) return '';
    const lines: string[] = [
      '',
      '---',
      `[attachments-info] ${String(persisted.length)} Datei(en) in diesem Turn hochgeladen + persistiert:`,
    ];
    for (const a of persisted) {
      const sizeKb = Math.round(a.sizeBytes / 1024);
      lines.push(
        `- ${a.fileName} (${a.contentType}, ${String(sizeKb)} KB) · storage_key=${a.storageKey}${a.signedUrl ? ` · signed_url=${a.signedUrl}` : ''}`,
      );
    }
    lines.push(
      '',
      'Wenn der User diese Datei als **Brand-Asset** markiert ("das ist unser Logo", "unser Firmenlogo", "das Team-Banner"), schreibe/aktualisiere die Memory-Datei `/memories/_brand/<asset-name>.md` (z.B. `/memories/_brand/logo.md`) mit folgenden Feldern:',
      '```yaml',
      'storage_key: <storage_key von oben>',
      'signed_url: <signed_url von oben>',
      'file_name: <Dateiname>',
      'content_type: <MIME-Type>',
      'uploaded_at: <heutiges Datum>',
      'asset_role: logo | banner | favicon | …',
      '```',
      'Die signed_url ist 7 Tage gültig — wenn du sie später in einem Diagramm brauchst und sie abgelaufen ist, frag nach oder lass sie vom Middleware via storage_key neu signieren.',
      'Wenn der User die Datei NICHT als Asset markiert (nur "hier ein Screenshot", "schau mal drauf"), NICHT ins Brand-Memory schreiben — nur inhaltlich antworten.',
    );
    return lines.join('\n');
  }

  /**
   * Resolve a Conductor human-await from an approve/reject card click. Out-of-band (no orchestrator
   * turn). Responder is keyed by the user's email (matching the email-keyed await holder); falls back
   * to the AAD object id when the email can't be resolved. Errors surface a friendly message — a
   * resolve failure must never throw out of the turn.
   */
  private async handleApprovalDecision(
    context: TurnContext,
    approval: ApprovalValue,
  ): Promise<void> {
    if (!this.resolveConductorAwait) {
      await context.sendActivity(
        'Conductor ist gerade nicht verfügbar — die Freigabe konnte nicht verarbeitet werden.',
      );
      return;
    }
    const from = context.activity.from;
    const userId =
      from?.aadObjectId ?? (typeof from?.id === 'string' ? from.id : undefined);
    // Fail closed: resolve the responder into the SAME id space as the await holder (the user's
    // email). Do NOT fall back to the AAD object id — it never matches an email-keyed holder, so the
    // kernel rejects it and a quorum='all' vote would be silently stranded (review H2). Ask to retry.
    const responderId = userId
      ? await this.resolveBindingEmail(context, userId)
      : undefined;
    if (!responderId) {
      await context.sendActivity(
        'Ich konnte deine Identität für diese Freigabe nicht bestätigen — bitte versuche es gleich noch einmal.',
      );
      return;
    }
    const approved = approval.decision === 'approve';
    try {
      const outcome = await this.resolveConductorAwait(approval.awaitId, responderId, approved);
      await context.sendActivity(approvalAckText(outcome, approved));
    } catch (err) {
      const detail = err instanceof Error ? err.message : String(err);
      await context.sendActivity(`Konnte die Freigabe nicht verarbeiten: ${detail}`);
    }
  }

  private async handleTopicDecision(
    context: TurnContext,
    sessionScope: string,
    userId: string | undefined,
    decision: { choice: 'continue' | 'reset'; originalMessage: string },
  ): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    const pending = this.history.getPending(sessionScope);
    // The card carries the original user message so we can proceed even if
    // the bot restarted in the meantime. `pending` just confirms we were in
    // the "asking" state; we don't hard-require it.
    if (!pending) {
      console.warn(
        `[topic] decision received but no pending entry for scope=${sessionScope}`,
      );
    }
    this.history.clearPending(sessionScope);
    if (decision.choice === 'reset') {
      this.history.resetTurns(sessionScope);
      console.error(
        `[topic] user chose RESET conv=${conversationId} msg="${truncate(decision.originalMessage, 60)}"`,
      );
    } else {
      console.error(
        `[topic] user chose CONTINUE conv=${conversationId} msg="${truncate(decision.originalMessage, 60)}"`,
      );
    }
    const priorTurns =
      decision.choice === 'reset' ? [] : this.history.get(sessionScope);
    await this.runOrchestratorTurnDetached(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage: decision.originalMessage,
      priorTurns,
    });
  }

  /**
   * User clicked the "🔄 Fresh Check" button on a previous answer card.
   * Re-runs the original question while bypassing the memory-read
   * convention AND the FTS context block — the orchestrator is told
   * explicitly to treat the turn as isolated. We intentionally do NOT
   * pass `priorTurns` so the in-memory conversation history can't
   * re-contaminate the turn either.
   */
  private async handleFreshCheck(
    context: TurnContext,
    sessionScope: string,
    userId: string | undefined,
    originalMessage: string,
  ): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    console.error(
      `[teams] fresh-check conv=${conversationId} msg="${truncate(originalMessage, 80)}"`,
    );
    // Empty priorTurns array — the point of fresh-check is strict isolation.
    const priorTurns: ReturnType<ConversationHistoryStore['get']> = [];
    await this.runOrchestratorTurn(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage: originalMessage,
      priorTurns,
      freshCheck: true,
    });
  }

  /**
   * User clicked one of the option buttons on an `ask_user_choice` card.
   * The chosen `value` is sent as the next user turn with the full chat
   * history preserved — the LLM then has the disambiguated input and can
   * answer normally.
   */
  private async handleChoiceAsk(
    context: TurnContext,
    sessionScope: string,
    userId: string | undefined,
    choice: { label: string; value: string },
  ): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    console.error(
      `[teams] choice-ask click conv=${conversationId} label="${truncate(choice.label, 40)}" value="${truncate(choice.value, 80)}"`,
    );
    this.history.clearPending(sessionScope);
    const priorTurns = this.history.get(sessionScope);
    await this.runOrchestratorTurnDetached(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage: choice.value,
      priorTurns,
    });
  }

  /**
   * User clicked a follow-up refinement button on a previous answer. The
   * option's full `prompt` becomes the next user turn — no free-text
   * extraction from the label needed because the LLM pre-composed the
   * self-contained question when invoking `suggest_follow_ups`.
   */
  private async handleFollowUp(
    context: TurnContext,
    sessionScope: string,
    userId: string | undefined,
    followUp: { label: string; prompt: string },
  ): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    console.error(
      `[teams] follow-up click conv=${conversationId} label="${truncate(followUp.label, 40)}" prompt="${truncate(followUp.prompt, 80)}"`,
    );
    this.history.clearPending(sessionScope);
    const priorTurns = this.history.get(sessionScope);
    await this.runOrchestratorTurnDetached(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage: followUp.prompt,
      priorTurns,
    });
  }

  /**
   * Silent SSO path: Teams exchanges the user's token for our resource
   * without a popup. Fires when the user had previously consented and the
   * cached token was still valid. We acknowledge so the user knows the
   * original turn can be retried; the next message will find the token
   * silently via `extractSsoAssertion`.
   */
  protected override async handleTeamsSigninTokenExchange(
    context: TurnContext,
    _query: SigninStateVerificationQuery,
  ): Promise<void> {
    await this.ackConsentAndPromptRetry(context);
  }

  /**
   * Magic-code / popup path: user completed the interactive sign-in. Same
   * ack as the silent path — we don't auto-retry because we never stored
   * the original message, and asking the user to retype is clearer than a
   * phantom replay.
   */
  protected override async handleTeamsSigninVerifyState(
    context: TurnContext,
    _query: SigninStateVerificationQuery,
  ): Promise<void> {
    await this.ackConsentAndPromptRetry(context);
  }

  private async ackConsentAndPromptRetry(context: TurnContext): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    console.error(`[teams] consent granted conv=${conversationId}`);
    await context.sendActivity(
      'Kalender-Zugriff freigegeben ✓ — stell deine Terminfrage einfach nochmal, dann buche ich direkt.',
    );
  }

  /**
   * Exchange a user-supplied 6-digit magic code from the Bot Framework
   * sign-in redirect page. Returns `true` when a token is now bound to
   * (user, connection), `false` otherwise (wrong code / expired flow /
   * adapter shape mismatch). Swallows errors — the caller falls back to
   * treating the text as a normal message.
   */
  private async tryExchangeMagicCode(
    context: TurnContext,
    userId: string,
    magicCode: string,
  ): Promise<boolean> {
    try {
      const adapter = context.adapter as unknown as { UserTokenClientKey?: symbol };
      if (!adapter.UserTokenClientKey || !this.ssoConnectionName) return false;
      const client = context.turnState.get(adapter.UserTokenClientKey) as
        | {
            getUserToken?: (
              userId: string,
              connectionName: string,
              channelId: string,
              magicCode?: string,
            ) => Promise<{ token?: string } | undefined>;
          }
        | undefined;
      if (!client?.getUserToken) return false;
      const channelId = context.activity.channelId ?? 'msteams';
      const result = await client.getUserToken(
        userId,
        this.ssoConnectionName,
        channelId,
        magicCode,
      );
      return Boolean(result?.token);
    } catch (err) {
      console.error(
        '[teams] magic-code exchange threw:',
        err instanceof Error ? err.message : err,
      );
      return false;
    }
  }

  /**
   * User clicked a slot button on a `find_free_slots` sidecar card. We
   * synthesize a user message that carries the opaque `slotId` so the
   * LLM recognises the booking intent and invokes `book_meeting(slotId, …)`.
   * The synthetic message also surfaces the human-readable label so the
   * transcript stays meaningful.
   */
  private async handleBookSlotClick(
    context: TurnContext,
    sessionScope: string,
    userId: string | undefined,
    bookSlot: { slotId: string; label: string },
  ): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    console.error(
      `[teams] book-slot click conv=${conversationId} slotId=${bookSlot.slotId} label="${truncate(bookSlot.label, 40)}"`,
    );
    this.history.clearPending(sessionScope);
    const priorTurns = this.history.get(sessionScope);
    const synthMessage =
      `Bitte diesen Termin buchen: ${bookSlot.label}. ` +
      `Nutze book_meeting mit slotId="${bookSlot.slotId}" und leite einen sinnvollen Betreff aus dem vorherigen Kontext ab.`;
    await this.runOrchestratorTurn(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage: synthMessage,
      priorTurns,
    });
  }

  /** Late-bound proactive sender (the adapter exists only after the router is
   *  created, which needs this bot) — see attachProactiveSend. */
  private proactiveSend?: TeamsProactiveSend;

  attachProactiveSend(send: TeamsProactiveSend): void {
    this.proactiveSend = send;
  }

  /**
   * Card Action.Submit clicks (choice / follow-up / topic decision / routine
   * filter): the Teams client renders "Something went wrong. Please try
   * again." whenever the submit's HTTP turn stays open past its ~15s budget —
   * which EVERY orchestrator turn does — even though the bot keeps processing
   * and answers fine. So: capture what only the live inbound activity can
   * give us (SSO assertion, verified sender name), end the inbound turn
   * immediately, and run the heavy work on a proactive continuation of the
   * same conversation. Without an attached proactive sender this falls back
   * to the old inline behaviour.
   */
  private async runOrchestratorTurnDetached(
    context: TurnContext,
    input: {
      conversationId: string;
      sessionScope: string;
      userId: string | undefined;
      userMessage: string;
      priorTurns: ReturnType<ConversationHistoryStore['get']>;
      freshCheck?: boolean;
    },
  ): Promise<void> {
    if (!this.proactiveSend) {
      await this.runOrchestratorTurn(context, input);
      return;
    }
    const presetSsoAssertion = this.ssoConnectionName
      ? await extractSsoAssertion(context, this.ssoConnectionName)
      : undefined;
    const presetSenderName = context.activity.from?.name;
    const reference = TurnContext.getConversationReference(context.activity);
    void this.proactiveSend(reference, async (proactive) => {
      await this.runOrchestratorTurn(proactive, {
        ...input,
        ...(presetSsoAssertion ? { presetSsoAssertion } : {}),
        ...(presetSenderName ? { presetSenderName } : {}),
      });
    }).catch((err: unknown) => {
      console.error(
        `[teams] detached card turn failed conv=${input.conversationId}:`,
        err instanceof Error ? err.message : err,
      );
    });
  }

  private async runOrchestratorTurn(
    context: TurnContext,
    input: {
      conversationId: string;
      sessionScope: string;
      userId: string | undefined;
      userMessage: string;
      priorTurns: ReturnType<ConversationHistoryStore['get']>;
      /** When true, the orchestrator skips context retrieval + memory read. */
      freshCheck?: boolean;
      /** Pre-extracted on the ORIGINAL inbound turn for detached card clicks —
       *  a proactive continuation cannot read the user token itself. */
      presetSsoAssertion?: string;
      /** Verified sender name captured on the original inbound turn — the
       *  proactive continuation's `from` is the bot, not the user. */
      presetSenderName?: string;
      /**
       * The agent this turn resolved to, carried rather than looked up again.
       *
       * Absent means "no per-agent resolver answered for this activity", which
       * is the same thing the instance field used to mean when it was
       * `undefined` — the default orchestrator answers. What is NOT the same
       * is another turn's agent arriving here, which is what a shared field
       * allowed.
       */
      chatAgent?: ChatAgent;
    },
  ): Promise<void> {
    // Scope-local binding so the ALS wrapper can forward the roster accessor
    // to the orchestrator (which surfaces `get_chat_participants` to Claude)
    // and sendAnswer (which resolves `<at>…</at>` tokens back into Mention
    // entities). 1:1 conversations skip ALS entirely — no roster, no mentions.
    const isGroupConversation =
      context.activity.conversation?.conversationType !== 'personal';
    const rosterProvider =
      this.rosterProvider && isGroupConversation ? this.rosterProvider : undefined;

    const run = async (): Promise<void> => {
      const stopTyping = startTypingLoop(context);
      const ssoAssertion =
        input.presetSsoAssertion ??
        (this.ssoConnectionName ? await extractSsoAssertion(context, this.ssoConnectionName) : undefined);
      console.log(
        `[teams] turn start conv=${input.conversationId} user=${input.userId ?? 'anon'} history=${String(input.priorTurns.length)} freshCheck=${String(Boolean(input.freshCheck))} roster=${rosterProvider ? 'on' : 'off'} sso=${ssoAssertion ? 'yes' : 'no'}`,
      );
      try {
        const userTimeZone =
          (context.activity as { localTimezone?: string }).localTimezone ??
          'Europe/Berlin';
        // Routines-feature hook: install the kernel's per-turn ALS so the
        // `manage_routine` tool can attribute `create` to this (tenant,
        // user) and capture the channel-native delivery handle. Cheap +
        // synchronous; no-op when the kernel didn't wire the callback.
        if (this.captureRoutineTurn && input.userId && this.tenantId) {
          try {
            const conversationRef =
              TurnContext.getConversationReference(context.activity);
            // P2 identity-bridge: key the Conductor channel binding by the user's email (the id an
            // operator addresses in a human step / role holder) when we've resolved it. The roster
            // lookup is async but this path must stay sync, so we read a cache and populate it in the
            // background — the binding becomes email-keyed from the first resolved turn onward,
            // falling back to the AAD object id until then.
            // Resolve the email BEFORE capture so the very first turn binds by it (await is fine — we
            // are in the async run loop; the roster is cached and the orchestrator turn dwarfs this).
            const principalRef = await this.resolveBindingEmail(context, input.userId);
            this.captureRoutineTurn({
              tenant: this.tenantId,
              userId: input.userId,
              ...(principalRef ? { principalRef } : {}),
              conversationRef,
            });
          } catch (refErr) {
            // Non-fatal: the rest of the turn proceeds normally; manage_routine
            // create will surface a friendly error if invoked this turn.
            console.error('[teams] routines turn-context capture failed:', refErr);
          }
        }
        const chatAgent =
          input.chatAgent ??
          this.resolveChatAgentForTurn(context.activity).chatAgent ??
          this.defaultOrchestrator;
        // #330 field report — the attributed form is what the agent sees AND
        // what lands in history; sendAnswer keeps the raw text for mention
        // resolution. See attributeGroupMessage for the why.
        const attributedMessage = attributeGroupMessage(input.userMessage, {
          isGroup: isGroupConversation,
          senderName: input.presetSenderName ?? context.activity.from?.name,
        });
        const result = await chatAgent.chat({
          userMessage: attributedMessage,
          sessionScope: input.sessionScope,
          ...(input.userId ? { userId: input.userId } : {}),
          ...(input.priorTurns.length > 0 ? { priorTurns: input.priorTurns } : {}),
          ...(input.freshCheck ? { freshCheck: true } : {}),
          ...(ssoAssertion ? { ssoAssertion } : {}),
          userTimeZone,
        });
        // Post-S+7.5: ChatAgent.chat returns the SDK's SemanticAnswer. We
        // narrow the discriminated `interactive` union into the legacy
        // pendingUserChoice / pendingSlotCard slots the Teams renderer
        // consumes. `runTrace` is no longer exposed — the trace panel reads
        // `undefined` and degrades gracefully (no section rendered).
        const pendingChoice =
          result.interactive?.kind === 'choice' ? result.interactive : undefined;
        const pendingSlots =
          result.interactive?.kind === 'slots' ? result.interactive : undefined;
        const pendingRoutineList =
          result.interactive?.kind === 'routine_list'
            ? result.interactive
            : undefined;
        if (isNoReply(result)) {
          logNoReplyDrop('teams', {
            userId: input.userId,
            sessionScope: input.sessionScope,
            conversationId: input.conversationId,
          });
          return;
        }
        // Cross-session KG-recall — surface what the per-turn probe pulled
        // from PRIOR sessions as a read-only card ABOVE the answer (context
        // first), mirroring the web-ui RecalledContextCard. No-op when empty.
        if (result.recalled) {
          const recallCard = buildRecalledContextCard(result.recalled);
          if (recallCard) {
            await context.sendActivity(MessageFactory.attachment(recallCard));
          }
        }
        // AI-Act marking (#643/#644): the kernel folds the disclosure
        // sentence into `text` for wire-only channels. Teams marks the
        // answer itself (✨ AI-generated card chip + activity-level AI
        // label entity), so the folded sentence would double-mark — strip
        // it, keeping any operator-authored addendum.
        const answerText = stripFoldedAiDisclosure(
          result.text,
          result.aiDisclosure?.text,
        );
        await sendAnswer(
          context,
          answerText,
          undefined,
          result.attachments,
          result.verifier,
          input.userMessage,
          rosterProvider,
          pendingChoice,
          result.followUps,
          pendingSlots,
          result.oauthConsentPending ? this.ssoConnectionName : undefined,
          result.captureDisclosure,
          result.privacyReceipt,
          result.maskedValues,
          {
            ...(result.agentsConsulted
              ? { agentsConsulted: result.agentsConsulted }
              : {}),
            ...(result.delegatedAnswer
              ? { delegatedAnswer: result.delegatedAnswer }
              : {}),
            ...(result.memoryUsed ? { memoryUsed: true } : {}),
          },
        );

        // Routine-list smart card (sidecar). Rendered AFTER the agent's
        // prose answer so the user sees the narration first ("Hier sind
        // deine 3 Routinen") and the actionable card right below.
        if (pendingRoutineList && this.buildRoutineListSmartCardAttachment) {
          const att = this.buildRoutineListSmartCardAttachment({
            filter: pendingRoutineList.filter,
            totals: pendingRoutineList.totals,
            routines: pendingRoutineList.routines,
          });
          await context.sendActivity(MessageFactory.attachment(att));
        }

        // History append skipped for blocking `ask_user_choice` turns —
        // the "answer" is just the clarification prompt, and appending it
        // would pollute the verbatim tail with a non-substantive turn.
        // Follow-up turns (suggest_follow_ups) DO get appended, since the
        // answer is a real response.
        if (!pendingChoice) {
          this.history.append(input.sessionScope, {
            userMessage: input.userMessage,
            // Stripped variant — the disclosure sentence is channel chrome,
            // not conversation content, and would pollute the verbatim tail.
            assistantAnswer: answerText,
            at: Date.now(),
          });
        }

        const attachmentCount = result.attachments?.length ?? 0;
        console.log(
          `[teams] turn done (attach=${attachmentCount}, history=${String(input.priorTurns.length + 1)}, conv=${input.conversationId}, user=${input.userId ?? 'anon'})`,
        );
      } catch (err) {
        console.error('[teams] orchestrator failure:', err);
        const detail = err instanceof Error ? err.message : String(err);
        await context.sendActivity(
          `Entschuldigung, beim Verarbeiten deiner Anfrage ist ein Fehler aufgetreten: ${detail}. Versuch es gleich nochmal oder stell die Frage anders.`,
        );
      } finally {
        stopTyping();
      }
    };

    if (rosterProvider) {
      await this.turnContext.runWithChatParticipants(
        () => rosterProvider.list(context),
        run,
      );
    } else {
      await run();
    }
  }
}

async function sendTopicAsk(
  context: TurnContext,
  previousSummary: string,
  originalMessage: string,
): Promise<void> {
  const card = buildTopicAskCard({ previousSummary, originalMessage });
  const activity = MessageFactory.attachment(card);
  await context.sendActivity(activity);
}

function truncate(value: string, max: number): string {
  const single = value.replace(/\s+/g, ' ').trim();
  if (single.length <= max) return single;
  return `${single.slice(0, max - 1)}…`;
}

function extractUserMessage(context: TurnContext): string {
  const cleaned = TurnContext.removeRecipientMention(context.activity) ?? '';
  // Teams sends message body twice (plain text + html attachment) for @mentions;
  // removeRecipientMention cleans the plain text, the html duplicate is on the
  // `attachments` array and is ignored here.
  return cleaned.replace(/\s+/g, ' ').trim();
}

function startTypingLoop(context: TurnContext): () => void {
  const sendTyping = (): void => {
    context.sendActivity({ type: ActivityTypes.Typing }).catch((err) => {
      console.warn('[teams] typing send failed:', err instanceof Error ? err.message : err);
    });
  };
  sendTyping();
  const handle = setInterval(sendTyping, TYPING_INTERVAL_MS);
  return () => clearInterval(handle);
}

/**
 * Sends the final turn reply:
 * - Short-answer path (≤ TEAMS_MAX_MESSAGE_CHARS): one Adaptive Card with the
 *   AI label, the answer and — if present — a collapsed tool-trace section.
 *   ToggleVisibility keeps the trace user-local so a group chat stays quiet.
 * - Long-answer fallback: paragraph-split plain text (Adaptive Cards can't
 *   be chunked). AI-label entity is attached to the first message.
 */
async function sendAnswer(
  context: TurnContext,
  text: string,
  runTrace: RunTracePayload | undefined,
  attachments: OutgoingAttachment[] | undefined,
  verifier: VerifierBadge | undefined,
  originalUserMessage: string | undefined,
  rosterProvider: TeamsRosterProvider | undefined,
  pendingUserChoice: OutgoingChoiceCard | undefined,
  followUpOptions: FollowUpOption[] | undefined,
  pendingSlotCard: OutgoingSlotPicker | undefined,
  consentConnectionName: string | undefined,
  captureDisclosure: CaptureDisclosure | undefined,
  privacyReceipt: PrivacyReceipt | undefined,
  maskedValues: readonly string[] | undefined,
  // #332 — harness-built agent transparency (footer + dynamic Direct-Line
  // buttons) and the Direct Line attributed-answer marker. `memoryUsed`
  // rides along as the Fresh-Check gate: the button renders only when the
  // kernel reported that memory actually influenced this answer.
  directLine?: {
    agentsConsulted?: readonly AgentConsultation[];
    delegatedAnswer?: DelegatedAnswer;
    memoryUsed?: boolean;
  },
): Promise<void> {
  // Blocking clarification — render a standalone Choice-Card instead of
  // an answer. The follow-up / fresh-check actions aren't meaningful here
  // since there's no concrete answer yet; attachments are ignored because
  // the orchestrator dropped them in this branch (see drainFollowUps).
  if (pendingUserChoice) {
    const card = buildChoiceAskCard({
      question: pendingUserChoice.question,
      ...(pendingUserChoice.rationale
        ? { rationale: pendingUserChoice.rationale }
        : {}),
      options: pendingUserChoice.options,
    });
    const activity = MessageFactory.attachment(card);
    activity.entities = [...(activity.entities ?? []), aiLabelEntity()];
    await context.sendActivity(activity);
    return;
  }

  const trimmed = text.trim();
  const hasAttachments = (attachments?.length ?? 0) > 0;
  // Images are the whole point when diagrams come back — always use the
  // Adaptive-Card path even if the accompanying text is empty.
  if (trimmed.length === 0 && !hasAttachments) {
    await context.sendActivity('Es ist keine Antwort zurückgekommen. Bitte versuch es erneut.');
    return;
  }

  // Resolve `<at>Display Name</at>` tokens against the roster BEFORE the
  // card/plain-text branching. Resolved mentions get threaded into the
  // Adaptive Card's `msteams.entities` so Teams renders the blue @-pill
  // directly inside the card — no separate plain-text activity on top.
  // Unresolved names get stripped (typo never pings a random person).
  const mentionPrep = await prepareMentions(
    context,
    trimmed.length === 0 && hasAttachments ? '' : trimmed,
    rosterProvider,
  );
  const resolvedText = mentionPrep.text;
  const mentionEntities =
    mentionPrep.kind === 'mentioned' ? mentionPrep.entities : [];

  const answerForCard = resolvedText.length === 0 && hasAttachments
    ? 'Hier das angeforderte Diagramm:'
    : resolvedText;

  if (answerForCard.length <= TEAMS_MAX_MESSAGE_CHARS) {
    const card = buildAnswerCard({
      answer: answerForCard,
      ...(runTrace ? { runTrace } : {}),
      ...(attachments ? { attachments } : {}),
      ...(verifier ? { verifier } : {}),
      ...(originalUserMessage ? { originalUserMessage } : {}),
      ...(directLine?.memoryUsed ? { showFreshCheck: true } : {}),
      ...(mentionEntities.length > 0 ? { mentions: mentionEntities } : {}),
      ...(followUpOptions && followUpOptions.length > 0
        ? { followUpOptions }
        : {}),
      ...(captureDisclosure ? { captureDisclosure } : {}),
      ...(privacyReceipt ? { privacyReceipt } : {}),
      ...(maskedValues && maskedValues.length > 0 ? { maskedValues } : {}),
      ...(directLine?.agentsConsulted && directLine.agentsConsulted.length > 0
        ? { agentsConsulted: directLine.agentsConsulted }
        : {}),
      ...(directLine?.delegatedAnswer
        ? { delegatedAnswer: directLine.delegatedAnswer }
        : {}),
    });
    const activity = MessageFactory.attachment(card);
    // Activity-level entities: AI label + mention entities for Teams'
    // notification pipeline. The card also carries the mentions via
    // `msteams.entities` — that's what drives the in-card pill rendering.
    activity.entities = [
      ...(activity.entities ?? []),
      ...mentionEntities,
      aiLabelEntity(),
    ];
    await context.sendActivity(activity);
    if (pendingSlotCard && pendingSlotCard.slots.length > 0) {
      const slotCard = buildSlotPickerCard({
        question: pendingSlotCard.question,
        slots: pendingSlotCard.slots.map((s) => ({
          slotId: s.slotId,
          label: s.label,
          confidence: s.confidence,
        })),
      });
      await context.sendActivity(MessageFactory.attachment(slotCard));
    }
    if (consentConnectionName) {
      const consentCard = await buildConsentSignInCard(context, consentConnectionName);
      await context.sendActivity(MessageFactory.attachment(consentCard));
    }
    return;
  }

  const chunks = splitAtParagraphs(resolvedText, TEAMS_MAX_MESSAGE_CHARS);
  for (let i = 0; i < chunks.length; i++) {
    const chunk = chunks[i];
    if (chunk === undefined) continue;
    const activity = MessageFactory.text(chunk);
    // AI label + mentions on the first chunk only — Teams uses the first
    // message in a split as the canonical one for notifications and labels.
    if (i === 0) {
      activity.entities = [
        ...(activity.entities ?? []),
        ...mentionEntities,
        aiLabelEntity(),
      ];
    }
    await context.sendActivity(activity);
  }

  // Long-answer fallback ate the card, so follow-up buttons need their own
  // mini-card appended after the chunks. Rare path (>25 KB answers); worth
  // the extra activity to keep refinements working.
  if (followUpOptions && followUpOptions.length > 0) {
    const card = buildFollowUpsOnlyCard(followUpOptions);
    await context.sendActivity(MessageFactory.attachment(card));
  }
  // #332 — same fallback for the agent-transparency footer + Direct-Line
  // buttons; otherwise the trust affordance silently vanishes on long answers.
  if (directLine?.agentsConsulted?.length || directLine?.delegatedAnswer) {
    const dlCard = buildDirectLineOnlyCard({
      ...(directLine.agentsConsulted
        ? { agentsConsulted: directLine.agentsConsulted }
        : {}),
      ...(directLine.delegatedAnswer
        ? { delegatedAnswer: directLine.delegatedAnswer }
        : {}),
      ...(originalUserMessage ? { originalUserMessage } : {}),
    });
    if (dlCard) await context.sendActivity(MessageFactory.attachment(dlCard));
  }
  if (pendingSlotCard && pendingSlotCard.slots.length > 0) {
    const slotCard = buildSlotPickerCard({
      question: pendingSlotCard.question,
      slots: pendingSlotCard.slots.map((s) => ({
        slotId: s.slotId,
        label: s.label,
        confidence: s.confidence,
      })),
    });
    await context.sendActivity(MessageFactory.attachment(slotCard));
  }
  if (consentConnectionName) {
    const consentCard = await buildConsentSignInCard(context, consentConnectionName);
    await context.sendActivity(MessageFactory.attachment(consentCard));
  }
}

/**
 * Sidecar sign-in card rendered when a calendar tool reported missing consent.
 *
 * Why not `CardFactory.oauthCard`: Teams' OAuthCard flow asks the client to
 * look up the sign-in URL via Bot Framework Token Service internally, which
 * sporadically fails under SingleTenant configurations with a generic
 * "Something went wrong" that gives zero server-side signal. Generating the
 * URL here via `UserTokenClient.getSignInResource` sidesteps that entire
 * handshake — we hand Teams a ready-to-open URL. The redirect from AAD still
 * lands at `token.botframework.com/.auth/web/redirect`, so the subsequent
 * `signin/verifyState` invoke on the bot side works exactly as before.
 *
 * On any failure to resolve the URL (no UserTokenClient in turnState, token
 * service down) we fall back to the stock OAuthCard — worst case we're back
 * to the old failure mode rather than a crashed turn.
 */
async function buildConsentSignInCard(
  context: TurnContext,
  connectionName: string,
): Promise<ReturnType<typeof CardFactory.adaptiveCard> | ReturnType<typeof CardFactory.oauthCard>> {
  const signInUrl = await tryGetSignInUrl(context, connectionName);
  const text =
    'Ich brauche einmalig deine Zustimmung, um auf deinen M365-Kalender zugreifen zu dürfen (freie Slots finden + Termine buchen). Klick auf den Button, bestätige im Tab — dann läuft es im Hintergrund. Zurück in Teams kannst du die Frage direkt nochmal stellen.';
  if (!signInUrl) {
    console.error('[teams] sign-in URL lookup failed — falling back to OAuthCard');
    return CardFactory.oauthCard(connectionName, 'Kalender freigeben', text);
  }
  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body: [
      {
        type: 'TextBlock',
        text: '🔐 Kalender-Zugriff erforderlich',
        weight: 'Bolder',
        size: 'Medium',
      },
      { type: 'TextBlock', text, wrap: true },
    ],
    actions: [
      {
        type: 'Action.OpenUrl',
        title: 'Kalender freigeben',
        url: signInUrl,
        style: 'positive',
      },
    ],
  });
}

async function tryGetSignInUrl(
  context: TurnContext,
  connectionName: string,
): Promise<string | undefined> {
  try {
    const adapter = context.adapter as unknown as { UserTokenClientKey?: symbol };
    if (!adapter.UserTokenClientKey) return undefined;
    const client = context.turnState.get(adapter.UserTokenClientKey) as
      | {
          getSignInResource?: (
            connectionName: string,
            activity: unknown,
            finalRedirect?: string,
          ) => Promise<{ signInLink?: string } | undefined>;
        }
      | undefined;
    if (!client?.getSignInResource) return undefined;
    const resource = await client.getSignInResource(connectionName, context.activity);
    return resource?.signInLink;
  } catch (err) {
    console.error(
      '[teams] getSignInResource failed:',
      err instanceof Error ? err.message : err,
    );
    return undefined;
  }
}

/**
 * Pull a silent Teams SSO assertion for the current user via the Bot
 * Framework `UserTokenClient` that CloudAdapter installs into turnState.
 * Returns `undefined` when:
 *   - the connection is missing/misconfigured,
 *   - the user hasn't consented yet (first-ever call returns no token),
 *   - the adapter shape differs (older BotFrameworkAdapter without the key).
 *
 * The caller treats every `undefined` identically: calendar tools stay
 * dormant and surface `sso_unavailable`. Errors are logged (stderr so Fly
 * keeps them under load) but never thrown — a Graph hiccup mustn't break
 * a regular non-calendar turn.
 */
async function extractSsoAssertion(
  context: TurnContext,
  connectionName: string,
): Promise<string | undefined> {
  try {
    const adapter = context.adapter as unknown as {
      UserTokenClientKey?: symbol;
    };
    if (!adapter.UserTokenClientKey) return undefined;
    const client = context.turnState.get(adapter.UserTokenClientKey) as
      | {
          getUserToken?: (
            userId: string,
            connectionName: string,
            channelId: string,
            magicCode?: string,
          ) => Promise<{ token?: string } | undefined>;
        }
      | undefined;
    if (!client?.getUserToken) return undefined;
    const fromId = context.activity.from?.id;
    const channelId = context.activity.channelId ?? 'msteams';
    if (!fromId) return undefined;
    const res = await client.getUserToken(fromId, connectionName, channelId);
    return res?.token;
  } catch (err) {
    console.error(
      '[teams] sso token extraction failed (continuing without):',
      err instanceof Error ? err.message : err,
    );
    return undefined;
  }
}

/**
 * Inspect the finished answer for `<at>Display Name</at>` tokens, attempt
 * to resolve them via the roster, and decide on the Activity shape:
 *
 * - `kind: 'plain'` — no tokens, or roster unavailable and tokens were
 *   stripped to plain text. Callers stay on the Adaptive-Card path.
 * - `kind: 'mentioned'` — at least one token resolved; send plain-text
 *   activity with matching Mention entities.
 *
 * Unresolved names are stripped and logged — we never forward a raw
 * `<at>…</at>` literal to Teams.
 */
async function prepareMentions(
  context: TurnContext,
  text: string,
  rosterProvider: TeamsRosterProvider | undefined,
): Promise<
  | { kind: 'plain'; text: string }
  | { kind: 'mentioned'; text: string; entities: MentionEntity[] }
> {
  if (text.length === 0 || !rosterProvider) {
    return { kind: 'plain', text: text.length === 0 ? text : stripMentionTokens(text) };
  }
  if (!text.includes('<at>')) {
    return { kind: 'plain', text };
  }
  let roster;
  try {
    roster = await rosterProvider.list(context);
  } catch (err) {
    console.error(
      '[teams] roster fetch failed — stripping mention tokens:',
      err instanceof Error ? err.message : err,
    );
    return { kind: 'plain', text: stripMentionTokens(text) };
  }
  const resolution = resolveMentions(text, roster);
  if (resolution.unresolved.length > 0) {
    console.error(
      `[teams] unresolved mentions stripped: ${resolution.unresolved.join(', ')}`,
    );
  }
  if (resolution.entities.length === 0) {
    return { kind: 'plain', text: stripMentionTokens(text) };
  }
  // Keep the resolved `<at>Name</at>` tokens intact (Teams renders the pill
  // around them), strip only the unresolved ones.
  let outText = text;
  for (const name of resolution.unresolved) {
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    outText = outText.replace(
      new RegExp(`<at>${escaped}</at>`, 'g'),
      name,
    );
  }
  return {
    kind: 'mentioned',
    text: outText,
    entities: resolution.entities,
  };
}

function splitAtParagraphs(text: string, maxLen: number): string[] {
  const chunks: string[] = [];
  let remaining = text;
  while (remaining.length > maxLen) {
    const window = remaining.slice(0, maxLen);
    const lastParagraph = window.lastIndexOf('\n\n');
    // Prefer paragraph break; fall back to line break; fall back to hard cut.
    const lastLine = window.lastIndexOf('\n');
    const cutAt =
      lastParagraph > maxLen / 2 ? lastParagraph : lastLine > maxLen / 2 ? lastLine : maxLen;
    chunks.push(remaining.slice(0, cutAt).trimEnd());
    remaining = remaining.slice(cutAt).trimStart();
  }
  if (remaining.length > 0) chunks.push(remaining);
  return chunks;
}

function isSharePointLike(url: string): boolean {
  try {
    const host = new URL(url).host.toLowerCase();
    return (
      host.endsWith('.sharepoint.com') ||
      host.endsWith('.sharepoint-df.com') ||
      host === '1drv.ms' ||
      host.endsWith('.1drv.ms')
    );
  } catch {
    return false;
  }
}

function fileNameFromUrl(url: string): string | undefined {
  try {
    const u = new URL(url);
    const last = u.pathname.split('/').filter(Boolean).pop();
    if (!last) return undefined;
    const decoded = decodeURIComponent(last);
    return decoded.length > 200 ? decoded.slice(0, 200) : decoded;
  } catch {
    return undefined;
  }
}

/**
 * Flat-walk any object / array, yielding every string value encountered.
 * Used to scavenge Teams-specific channelData + entities for SharePoint
 * URLs without hard-coding each schema variant (which Microsoft changes
 * regularly).
 */
function* collectStringsDeep(value: unknown): Iterable<string> {
  if (typeof value === 'string') {
    yield value;
    return;
  }
  if (Array.isArray(value)) {
    for (const v of value) yield* collectStringsDeep(v);
    return;
  }
  if (value && typeof value === 'object') {
    for (const v of Object.values(value)) yield* collectStringsDeep(v);
  }
}

// ---------------------------------------------------------------------------
// #860 W2 (issue #21) — auto-invite onboarding hook
// ---------------------------------------------------------------------------
// Module-level (exported) so tests drive the hook logic directly with fake
// contexts — instantiating a full TeamsBot needs the whole dependency fan-in.
// The class methods above are thin delegates onto these functions.

/**
 * The seam the wiring layer (plugin.ts) injects: the installer run plus the
 * non-consuming auto-install marker probe — BOTH backed by the same
 * `TeamsAgentInstaller` instance so hook and installer correlate on the
 * same in-memory marker. No Graph call happens outside this seam.
 */
export interface TeamsAutoInviteDeps {
  installAgentApps(
    request: InstallAgentAppsRequest,
  ): Promise<TeamsAgentInstallResult>;
  probeAutoInstallMarker(teamId: string): boolean;
}

/** Minimal structural slice of `TurnContext` the hook needs — lets tests
 *  drive it without a Bot-Framework adapter (the real `TurnContext`
 *  satisfies it structurally). */
export interface TeamsAutoInviteTurnContext {
  readonly activity: Activity;
  sendActivity(activity: Partial<Activity>): Promise<unknown>;
  updateActivity(activity: Partial<Activity>): Promise<unknown>;
}

/**
 * Graph team (group) id + tenant id of a TEAM-scope activity, from Teams
 * `channelData`. `undefined` for personal / group-chat scopes — and also
 * when the event carries no `aadGroupId`: `installToTeam` targets Graph
 * `POST /teams/{group-id}/installedApps`, and only the AAD group id will
 * do (the `19:…` thread id is not a Graph team id).
 */
export function teamsTeamScopeFromActivity(
  activity: Activity,
): { teamId: string; tenantId: string } | undefined {
  const channelData = activity.channelData as
    | {
        team?: { id?: string; aadGroupId?: string };
        tenant?: { id?: string };
      }
    | undefined;
  const teamId = channelData?.team?.aadGroupId?.trim();
  const tenantId = (
    channelData?.tenant?.id ?? activity.conversation?.tenantId
  )?.trim();
  if (!teamId || !tenantId) return undefined;
  return { teamId, tenantId };
}

/**
 * `true` when a `membersAdded` activity is the ECHO of a just-run
 * auto-install: THIS bot is among the added members, the conversation is a
 * team scope, and the team's auto-install marker is still fresh. Consumed
 * by `handleMembershipChange` to downgrade `bot_added` → `members_added`
 * (suppressing the announced-entry intro) and by the hook to skip a second
 * installer run. The probe is non-consuming — every bot of the flock can
 * correlate the same event.
 */
export function shouldSuppressAutoInstallIntro(
  probeAutoInstallMarker: ((teamId: string) => boolean) | undefined,
  activity: Activity,
): boolean {
  if (!probeAutoInstallMarker) return false;
  const recipientId = activity.recipient?.id;
  if (!recipientId) return false;
  if (!(activity.membersAdded ?? []).some((m) => m.id === recipientId)) {
    return false;
  }
  const team = teamsTeamScopeFromActivity(activity);
  if (!team) return false;
  return probeAutoInstallMarker(team.teamId);
}

/** What one hook invocation did — surfaced for logging + tests. */
export type TeamsAutoInviteHookResult =
  | 'skipped'
  | 'suppressed'
  | 'posted-result-card';

/**
 * The onboarding hook body: run the installer for the team THIS bot was
 * just added to and post the result / fallback card into the conversation.
 * Trigger conditions (all must hold, otherwise `'skipped'`):
 *   * the added member IS this bot (`recipient` id match) — a human
 *     joining the team must never trigger installs,
 *   * the conversation is a team scope with a Graph group id,
 *   * the team's auto-install marker is NOT fresh (else `'suppressed'` —
 *     this bot is itself a freshly auto-installed agent app, and only the
 *     bot that ran the installer posts the summary).
 */
export async function runTeamsAutoInviteHook(
  deps: TeamsAutoInviteDeps,
  context: TeamsAutoInviteTurnContext,
): Promise<TeamsAutoInviteHookResult> {
  const activity = context.activity;
  const recipientId = activity.recipient?.id;
  const accounts = activity.membersAdded ?? [];
  if (!recipientId || !accounts.some((m) => m.id === recipientId)) {
    return 'skipped';
  }
  const team = teamsTeamScopeFromActivity(activity);
  if (!team) return 'skipped';
  if (deps.probeAutoInstallMarker(team.teamId)) return 'suppressed';

  const result = await deps.installAgentApps(team);
  if (result.outcomes.length === 0) return 'skipped';
  const card = buildAgentAppsResultCard({
    outcomes: result.outcomes,
    teamId: team.teamId,
    tenantId: team.tenantId,
  });
  await context.sendActivity({
    type: ActivityTypes.Message,
    attachments: [card],
  });
  return 'posted-result-card';
}

/**
 * "🔄 Prüfen" re-check: re-run the installer and UPDATE the existing card
 * (the submit's `replyToId` names the card activity); falls back to a
 * fresh post when the channel rejects the update. Card `data` is
 * client-editable, so the transport-derived team/tenant of the submit
 * activity wins over the round-tripped values whenever present.
 */
export async function handleTeamsAgentAppsRecheck(
  deps: TeamsAutoInviteDeps,
  context: TeamsAutoInviteTurnContext,
  submitted: AgentAppsRecheckValue,
): Promise<void> {
  const transport = teamsTeamScopeFromActivity(context.activity);
  const request: InstallAgentAppsRequest = transport ?? {
    teamId: submitted.teamId,
    tenantId: submitted.tenantId,
  };
  const result = await deps.installAgentApps(request);
  const card = buildAgentAppsResultCard({
    outcomes: result.outcomes,
    teamId: request.teamId,
    tenantId: request.tenantId,
  });
  const cardActivityId = context.activity.replyToId;
  if (cardActivityId) {
    try {
      await context.updateActivity({
        type: ActivityTypes.Message,
        id: cardActivityId,
        conversation: context.activity.conversation,
        attachments: [card],
      });
      return;
    } catch {
      // Some contexts reject activity updates — fall through to a fresh post.
    }
  }
  await context.sendActivity({
    type: ActivityTypes.Message,
    attachments: [card],
  });
}

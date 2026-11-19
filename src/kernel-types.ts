/**
 * Narrow structural shims for kernel-internal types that the Teams channel
 * package consumes but does not own. Everything here is *structural* — no
 * value imports from kernel paths. Value-carrying dependencies (concrete
 * classes like `ConversationHistoryStore`, `TopicDetector`, and the
 * `turnContext` AsyncLocalStorage namespace) are supplied via dependency
 * injection through `TeamsChannelPluginDeps`; the kernel constructs them
 * once and passes the instances into the plugin, which keeps this file
 * path-independent.
 *
 * Stability: post-S+7.5 cut, `ChatAgent.chat(…)` returns `SemanticAnswer`
 * from `@omadia/channel-sdk`; the previous `ChatTurnResult` +
 * DiagramAttachment/FollowUpOption/PendingUserChoice/PendingSlotCard/
 * VerifierResultSummary shims are gone. RunTracePayload stays because the
 * Adaptive-Card trace panel still reads it via an internal-observability
 * bridge (see teamsCard.ts); it is not channel-contract territory.
 */
import type {
  CaptureDisclosure,
  FollowUpOption as SdkFollowUpOption,
  OutgoingAttachment,
  OutgoingChoiceCard,
  OutgoingSlotPicker,
  SemanticAnswer,
  VerifierBadge,
  AgentConsultation,
  DelegatedAnswer,
  TurnOrigin,
} from '@omadia/channel-sdk';

// Re-export SDK shapes under the names the Teams package uses internally.
// Keeps local imports one-line (`from './kernel-types.js'`) instead of a
// mix of './kernel-types.js' + '@omadia/channel-sdk' everywhere.
export type {
  CaptureDisclosure,
  SdkFollowUpOption as FollowUpOption,
  OutgoingAttachment,
  OutgoingChoiceCard,
  OutgoingSlotPicker,
  SemanticAnswer,
  VerifierBadge,
  // #332 — agent-transparency footer + Direct Line attributed answer.
  AgentConsultation,
  DelegatedAnswer,
};

// ---------------------------------------------------------------------------
// Participants — mirror of src/services/chatParticipants.ts
// ---------------------------------------------------------------------------

export interface ChatParticipant {
  channelUserId: string;
  aadObjectId: string | null;
  displayName: string;
  email: string | null;
  userPrincipalName: string | null;
  /** Kernel #1018 — `'agent'` marks a bot; absent = a person. */
  kind?: 'human' | 'agent';
}

/**
 * `completeRoster` mirrors the kernel flag (member-scoped memory): the provider
 * promises its non-empty roster lists EVERY member of the chat. The kernel
 * gives a group member-scoped knowledge only when this is set.
 */
export type ChatParticipantsProvider = (() => Promise<ChatParticipant[]>) & {
  readonly completeRoster?: boolean;
};

// ---------------------------------------------------------------------------
// Turn context — mirror of src/services/turnContext.ts (structural)
// ---------------------------------------------------------------------------

export interface TurnContextValue {
  turnId: string;
  turnDate: string;
  chatParticipants?: ChatParticipantsProvider;
}

/**
 * The turn-context namespace object the kernel exposes. Injected into the
 * Teams plugin via `deps.turnContext` so the package and the kernel share
 * the same AsyncLocalStorage instance — critical for the orchestrator to
 * observe the `chatParticipants` provider the Teams adapter installed.
 */
export interface TurnContextModule {
  run<T>(value: TurnContextValue, fn: () => Promise<T>): Promise<T>;
  enter(value: TurnContextValue): void;
  runWithChatParticipants<T>(
    chatParticipants: ChatParticipantsProvider,
    fn: () => Promise<T>,
  ): Promise<T>;
  current(): TurnContextValue | undefined;
  currentTurnId(): string | undefined;
  currentTurnDate(): string;
}

// ---------------------------------------------------------------------------
// Conversation history — mirror of src/services/conversationHistory.ts
// ---------------------------------------------------------------------------

export interface ConversationTurn {
  userMessage: string;
  assistantAnswer: string;
  /** Unix millis of the user message. */
  at: number;
}

export interface PendingTopicDecision {
  userMessage: string;
  askedAt: number;
}

/**
 * Structural shape of the kernel's `ConversationHistoryStore` class. The
 * concrete instance is injected via `deps.conversationHistoryStore` so the
 * package never has to reach into the kernel to `new` one up.
 */
export interface ConversationHistoryStore {
  get(scope: string): ConversationTurn[];
  append(scope: string, turn: ConversationTurn): void;
  resetTurns(scope: string): void;
  markPending(scope: string, pending: PendingTopicDecision): void;
  getPending(scope: string): PendingTopicDecision | undefined;
  clearPending(scope: string): void;
  size(): number;
  clear(): void;
}

// ---------------------------------------------------------------------------
// Embedding client — mirror of src/services/embeddingClient.ts
// ---------------------------------------------------------------------------

export interface EmbeddingClient {
  embed(text: string): Promise<number[]>;
}

// ---------------------------------------------------------------------------
// Topic detector — mirror of src/services/topicDetector.ts (structural)
// ---------------------------------------------------------------------------

export type TopicDecision = 'continue' | 'reset' | 'ask';

export interface TopicClassifyInput {
  userMessage: string;
  history: readonly ConversationTurn[];
}

export interface TopicClassifyResult {
  decision: TopicDecision;
  reason: string;
  similarity?: number;
  classifier?: 'continue' | 'reset' | 'unsure';
}

export interface TopicDetectorOptions {
  upperThreshold?: number;
  lowerThreshold?: number;
  centroidDepth?: number;
  classifierModel?: string;
  classifierMaxTokens?: number;
  fallbackDecision?: TopicDecision;
}

/**
 * Structural shape of the kernel's `TopicDetector` class. The concrete
 * instance is injected via `deps.topicDetector` (conditionally, only when
 * the embedding client is available at kernel boot).
 */
export interface TopicDetector {
  classify(input: TopicClassifyInput): Promise<TopicClassifyResult>;
}

// ---------------------------------------------------------------------------
// Run trace — mirror of src/services/runTraceCollector.ts (structural)
// ---------------------------------------------------------------------------

/**
 * Structural shim for `RunTracePayload = Omit<RunTrace, 'turnId'>`. The
 * Teams adapter reads the fields below to render a trace panel in the
 * Adaptive Card — keep these in sync with
 * `src/services/knowledgeGraph.ts:RunTrace`.
 */
export type RunStatus = 'success' | 'failed' | string;

export interface RunToolCallShim {
  toolName: string;
  durationMs: number;
  isError: boolean;
}

export interface RunAgentInvocationShim {
  agentName: string;
  status: RunStatus;
  durationMs: number;
  subIterations: number;
  toolCalls: RunToolCallShim[];
}

export interface RunTracePayload {
  status: RunStatus;
  durationMs: number;
  iterations: number;
  orchestratorToolCalls: RunToolCallShim[];
  agentInvocations: RunAgentInvocationShim[];
}

// ---------------------------------------------------------------------------
// Chat agent — mirror of ChatAgent + ChatTurnInput
// ---------------------------------------------------------------------------

/**
 * Answer-size scale — mirror of the kernel's `ANSWER_VERBOSITY_LEVELS`
 * (`@omadia/channel-sdk` ≥ the phase-3 cut; mirrored here because the SDK
 * this package compiles against predates it). The kernel parses the wire
 * value and ignores anything outside the scale, so a drift here can only
 * make a button inert, never pick a wrong level.
 */
export const ANSWER_VERBOSITY_LEVELS = [
  'tldr',
  'brief',
  'standard',
  'detailed',
  'max',
] as const;
export type AnswerVerbosityLevel = (typeof ANSWER_VERBOSITY_LEVELS)[number];

/** The level an answer was generated under (`SemanticAnswer.answerVerbosity`). */
export interface AnswerVerbosityInfo {
  readonly effective: AnswerVerbosityLevel;
  readonly source: 'turn' | 'configured';
}

export function isAnswerVerbosityLevel(raw: unknown): raw is AnswerVerbosityLevel {
  return (
    typeof raw === 'string' &&
    (ANSWER_VERBOSITY_LEVELS as readonly string[]).includes(raw)
  );
}

/**
 * Read the level off a kernel answer structurally: the SDK type this package
 * compiles against may predate the field, and an older kernel simply omits
 * it — in which case the card offers no size buttons.
 */
export function readAnswerVerbosity(answer: unknown): AnswerVerbosityInfo | undefined {
  if (!answer || typeof answer !== 'object') return undefined;
  const raw = (answer as { answerVerbosity?: unknown }).answerVerbosity;
  if (!raw || typeof raw !== 'object') return undefined;
  const { effective, source } = raw as { effective?: unknown; source?: unknown };
  if (!isAnswerVerbosityLevel(effective)) return undefined;
  if (source !== 'turn' && source !== 'configured') return undefined;
  return { effective, source };
}

export interface ChatTurnInput {
  userMessage: string;
  sessionScope?: string;
  userId?: string;
  priorTurns?: Array<{ userMessage: string; assistantAnswer: string }>;
  extraSystemHint?: string;
  freshCheck?: boolean;
  /**
   * Answer-size pick for THIS turn only (the "Kürzer" / "Mehr Details" card
   * buttons re-asking the same question). Replaces the configured level for
   * the one turn; an older kernel ignores the field.
   */
  answerVerbosity?: AnswerVerbosityLevel;
  ssoAssertion?: string;
  userTimeZone?: string;
  /**
   * W5 memory-ACL (#870 §5) — the turn's context origin, so the kernel can
   * scope `/memories/` to this chat context instead of to the agent alone.
   *
   * `TurnOrigin` is imported as a real SDK type rather than mirrored, unlike
   * everything else in this file: the shims here exist for KERNEL-INTERNAL
   * types the plugin cannot import, whereas this one is published by
   * `@omadia/channel-sdk` precisely so channel plugins can name it. Mirroring
   * it structurally would be worse than useless — a drift between the two
   * spellings is a memory-isolation bug, and the compiler would not see it.
   *
   * Optional, because the kernel treats a missing origin as context-free: an
   * older kernel ignores the field, a newer one behaves exactly as today when
   * it is absent.
   */
  origin?: TurnOrigin;
}

/**
 * Options of `ConversationRosterProvider.getRoster` (`@omadia/channel-sdk`
 * 0.3.2, #2043). `asChannelKey` is the asking agent's own identity
 * (`28:<appId>`): the roster is then read as THAT bot, or not at all.
 */
export interface RosterRequestOptions {
  asChannelKey?: string;
  /** Bypass any cache and ask the platform now (kernel spec B §5.5, #2091). */
  fresh?: boolean;
}

/** How complete a roster answer is (kernel spec B §5.5, #2091). */
export type RosterStatus = 'complete' | 'partial' | 'unknown';

/**
 * `ConversationRoster` with the roster status of #2091, in the shape agreed
 * with the kernel for `@omadia/channel-sdk` (additive there):
 *   status        how complete `participants` is
 *   readAt        ISO 8601 UTC, when the platform answered this list; a
 *                 cached answer keeps the time of its read
 *   statusDetail  operator-readable reason when the status is not `complete`
 * `partial` stays and is true unless the status is `complete`, so a kernel
 * that does not know `status` keeps working.
 */
export type TeamsConversationRoster = import('@omadia/channel-sdk').ConversationRoster & {
  status?: RosterStatus;
  readAt?: string;
  statusDetail?: string;
};

/**
 * `TurnOrigin` with the field `@omadia/channel-sdk` 0.4.0 adds (#2043):
 * the conversation as the CHANNEL names it, here `activity.conversation.id`
 * as received (no `teams-` prefix, no bot key, `;messageid=<root>` kept).
 * The kernel uses it for the tool call context, the subscription check and
 * the own-post counter instead of deriving an id from the session scope.
 *
 * Mirrored as an intersection, so it is the SDK's own type as soon as the
 * SDK declares the field, and an older kernel just ignores it.
 */
export type TeamsTurnOrigin = TurnOrigin & {
  readonly channelConversationId?: string;
  /**
   * channel-sdk 0.3.2 — who spoke, as the platform stated it: the Entra
   * object id and display name of `activity.from`. The kernel builds the
   * tool call context's `callerUserRef` from it, so a facilitator can name
   * the person who started a facilitation. (`ChatTurnInput.channelIdentity`
   * is NOT set by this channel: it has knowledge-graph meaning.)
   */
  readonly speaker?: { readonly kind: string; readonly id: string; readonly displayName?: string };
};

/**
 * Channel-facing chat agent. `chat()` returns the SDK's SemanticAnswer —
 * Teams doesn't consume `chatStream` (that's HTTP-route territory), so the
 * shim omits it. The concrete kernel class (Orchestrator or VerifierService)
 * still implements both — the structural shim is intentionally narrower
 * than the real interface.
 */
export interface ChatAgent {
  chat(input: ChatTurnInput): Promise<SemanticAnswer>;
}

// ---------------------------------------------------------------------------
// teamsProvisioner@1 — mirror of @omadia/integration-microsoft365
// src/teamsProvisioner (structural)
// ---------------------------------------------------------------------------
//
// #860 W2 — the auto-invite installer (`teamsAgentInstaller.ts`) consumes the
// M365 connector's `teamsProvisioner@1` service. The connector is a
// peerDependency ('*', not in node_modules), so — like every other
// kernel/plugin type in this file — its contract is mirrored STRUCTURALLY,
// never imported. Source of truth: connector `src/teamsProvisioner/`
// (types.ts / catalog.ts / install.ts / errors.ts / index.ts).
//
// The accessor shim is intentionally NARROWER than the real
// `TeamsProvisionerAccessor` (precedent: the `ChatAgent` shim above): only
// the two chain steps the installer calls are mirrored. Resolving the
// service object still typechecks because TypeScript matches structurally.

/**
 * Idempotency signal of provisioner write steps — Graph 409 "already exists"
 * paths are NOT errors; callers branch on `outcome` instead of
 * string-matching error bodies.
 */
export type IdempotentOutcome = 'created' | 'already-existed';

/** Result wrapper carrying the {@link IdempotentOutcome} alongside the value. */
export interface Idempotent<T> {
  readonly outcome: IdempotentOutcome;
  readonly value: T;
}

/** An app installation into one team (`POST /teams/{id}/installedApps`). */
export interface TeamAppInstallation {
  readonly teamId: string;
  readonly teamsAppId: string;
  /** Graph installation id when the API returned/located one. */
  readonly installationId?: string;
}

/**
 * Install request for the team-install step. Mirror of the connector's
 * `InstallToTeamRequest` minus the optional `consentedPermissionSet` (RSC
 * consent) — the installer never sends one, and an absent optional field
 * keeps this narrow shape assignable to the real request type.
 */
export interface InstallToTeamRequest {
  readonly teamId: string;
  /** Catalog id (`CatalogTeamsApp.teamsAppId`). */
  readonly teamsAppId: string;
}

/** Input for the catalog-lookup probe (`getCatalogApp`, connector >= 0.3.1). */
export interface GetCatalogAppInput {
  /** Manifest id (`externalId`) of the catalog app to resolve. */
  readonly teamsAppExternalId: string;
}

/** Lookup miss — no catalog app carries the requested `externalId`. */
export interface CatalogAppNotFound {
  readonly found: false;
}

/** Lookup hit. `displayName`/`publishedVersion` are optional on purpose —
 *  Graph can omit either on a thin-but-installable catalog entry. */
export interface CatalogAppFound {
  readonly found: true;
  /** Catalog id (`teamsApp.id`) — what installs reference. */
  readonly teamsAppId: string;
  readonly displayName?: string;
  readonly publishedVersion?: string;
}

/** Result of the catalog-lookup probe — `{ found: false }` is a plain
 *  outcome, never an exception. */
export type GetCatalogAppResult = CatalogAppNotFound | CatalogAppFound;

/**
 * Structural shim of the `teamsProvisioner@1` service object (published
 * under the ServiceRegistry key `teamsProvisioner`). Narrowed to the chain
 * steps the Teams plugin calls; the real accessor also exposes
 * createAppRegistration/createBot/uploadToCatalog/… for the middleware
 * agent factory.
 *
 * `getCatalogApp` ships with connector 0.3.1 — it is OPTIONAL here and the
 * installer feature-detects it. Never require the method: a 0.3.0 connector
 * without it must degrade to the fallback-card outcome, not crash.
 */
export interface TeamsProvisionerAccessor {
  /**
   * Lookup probe — resolve an EXISTING catalog app by manifest id
   * (`externalId`) without uploading a package (connector >= 0.3.1).
   */
  getCatalogApp?(input: GetCatalogAppInput): Promise<GetCatalogAppResult>;
  /** Chain step 5 — install the catalog app into one team. */
  installToTeam(
    input: InstallToTeamRequest,
  ): Promise<Idempotent<TeamAppInstallation>>;
}

/**
 * Structural shape of the connector's `ConsentMissingError` (Graph/ARM 403:
 * application permission or admin consent missing). The class itself is not
 * importable here, and cross-package `instanceof` would be unreliable
 * anyway — the connector sets `this.name` explicitly, so consumers branch
 * on `name` (see `isConsentMissingError` in `teamsAgentInstaller.ts`).
 */
export interface ConsentMissingErrorShape extends Error {
  readonly name: 'ConsentMissingError';
  /** The scopes/app roles the caller must have granted. */
  readonly missingScopes: readonly string[];
  /** Which API rejected the call. */
  readonly resource: 'graph' | 'arm';
}

/**
 * Structural shape of the connector's `ProvisioningThrottledError` (thrown
 * after ITS 429 retry/backoff budget is exhausted). Carries the last
 * `Retry-After` hint, when the API sent one.
 */
export interface ProvisioningThrottledErrorShape extends Error {
  readonly name: 'ProvisioningThrottledError';
  /** Seconds from the final `Retry-After` header, if the API provided it. */
  readonly retryAfterSeconds?: number;
  /** Which API throttled the call. */
  readonly resource: 'graph' | 'arm';
}

// ---------------------------------------------------------------------------
// Config — narrow subset the Teams plugin consumes from kernel `Config`
// ---------------------------------------------------------------------------

/**
 * The set of shared-`.env` knobs the Teams plugin reads from the kernel-
 * owned `Config` object. Kept narrow so adding a field to the kernel config
 * doesn't implicitly widen the Teams package's surface; adding a new field
 * here is an intentional act.
 */
export interface TeamsConfigShim {
  // #860 W0a — MICROSOFT_APP_TYPE left this shim: the app registration
  // type is per-bot now (`TeamsBotIdentity.appType`), never process-global.
  // The legacy env knob is consumed only by the scalar-credential shim
  // (`legacyTeamsBotFromScalars`) when no `teams_bots[]` list is set.
  TEAMS_ATTACHMENT_KEY_PREFIX: string;
  TEAMS_ATTACHMENT_STORAGE_ENABLED: boolean;
  TEAMS_ATTACHMENT_MAX_BYTES: number;
  ATTACHMENT_URL_SECRET?: string;
  ATTACHMENT_PUBLIC_BASE_URL?: string;
  ATTACHMENT_SIGNED_URL_TTL_SEC: number;
}

// ---------------------------------------------------------------------------
// Native facilitators (byte5ai/omadia#2043) — mirror of the channel-sdk 0.3.0
// / plugin-api 1.55.0 additions
// ---------------------------------------------------------------------------
//
// Mirrored structurally for the reason given at the top of this file: the
// plugin ships independently of the kernel and is compiled against whatever
// SDK the host tree carries. Sources of truth:
//   channel-sdk  src/conversationMessage.ts, src/sessionCapabilities.ts,
//                src/targetedSend.ts, src/coreApi.ts
//   plugin-api   src/conversationMessages.ts, src/sessionCapabilities.ts
// Every CoreApi member below is OPTIONAL in the kernel contract and is
// feature-detected in plugin.ts; a kernel without them leaves this plugin's
// behaviour exactly as it was.

/** A channel user as the kernel addresses one (`ChannelUserRef`, narrowed to
 *  the Teams namespace this plugin emits). */
export interface TeamsUserRef {
  kind: 'teams-aad';
  /** AAD object id when the activity carries one, else the channel id. */
  id: string;
  displayName?: string;
  email?: string;
}

export type ConversationMessageKind =
  | 'created'
  | 'edited'
  | 'deleted'
  | 'reaction'
  | 'card_answer';

/** One inbound message event. This is not a turn: emitting it starts no
 *  model call. */
export interface ConversationMessageEvent {
  /** The emitting channel plugin's catalog id. */
  channelId: string;
  channelType: string;
  conversationId: string;
  conversationType?: 'direct' | 'group';
  /** The bot identity that received the event (`28:<appId>`). */
  receivedByChannelKey: string;
  /** Platform message id; for 'reaction' / 'card_answer' the message reacted
   *  to / the card answered. */
  messageId: string;
  /** 0 for the original; incremented per platform edit. */
  revision: number;
  kind: ConversationMessageKind;
  /** The platform-verified sender. Never derived from message text. */
  author: TeamsUserRef;
  tenantId?: string;
  /** Plain text, mentions resolved, markup stripped. Absent for 'deleted'. */
  text?: string;
  replyToMessageId?: string;
  quoted?: boolean;
  forwarded?: boolean;
  reaction?: { name: string; removed: boolean };
  /** `answererVerified: false` when the platform did not state who clicked
   *  (no verified user id on the activity). */
  cardAnswer?: { cardId: string; value: unknown; answererVerified?: boolean };
  /** ISO 8601 timestamp of the platform event. */
  occurredAt: string;
  /** The original platform event, opaque to the kernel and not stored. */
  rawEvent?: unknown;
}

/** The readiness codes the kernel contract documents (plugin-api
 *  `ChannelReadinessCode`). */
export type ChannelReadinessCode =
  | 'no_identity'
  | 'app_not_installed'
  | 'consent_missing'
  | 'cannot_post'
  | 'graph_forbidden'
  | 'graph_error'
  | 'graph_unavailable'
  | 'team_unknown'
  | 'unsupported_conversation'
  | 'package_outdated';

export interface ChannelReadinessIssue {
  /** One of {@link ChannelReadinessCode}, or a code a newer channel added. */
  code: ChannelReadinessCode | (string & {});
  /** Precise, operator-readable reason. */
  message: string;
}

export interface ChannelReadinessReport {
  ready: boolean;
  installed: boolean;
  /** Teams: RSC `ChatMessage.Read.Chat` consented in this chat. */
  canReadAllMessages: boolean;
  canPost: boolean;
  /** Empty when `ready`. One entry per thing that is missing. */
  missing: readonly ChannelReadinessIssue[];
  /** Set by the kernel: `false` when no channel checked. A provider's own
   *  report never sets it. */
  verified?: boolean;
}

/** Outcome of a consent repair (plugin-api `ChannelConsentRepairOutcome`). */
export type ChannelConsentRepairOutcome =
  | { outcome: 'upgraded'; installationId?: string; permissions?: readonly string[] }
  | { outcome: 'not_needed' }
  | { outcome: 'failed'; code: ChannelReadinessCode | 'unsupported' | (string & {}); message: string };

/** RETURNS what is missing instead of throwing. */
export interface ChannelReadinessProvider {
  channelType: string;
  check(conversationId: string, opts: { asChannelKey: string }): Promise<ChannelReadinessReport>;
  /** Optional: repair missing consent for that identity in the conversation
   *  (Teams: upgrade the app installation). RETURNS a failure, never throws. */
  repairConsent?(conversationId: string, opts: { asChannelKey: string }): Promise<ChannelConsentRepairOutcome>;
}

export type SessionEventKind = 'session_started' | 'session_ended' | 'participant_joined' | 'participant_left';

/** A meeting lifecycle event (channel-sdk `SessionEvent`). The adapter emits
 *  every one it receives, once per receiving bot; the kernel filters by
 *  subscription and removes the duplicates. */
export interface SessionEvent {
  channelId: string;
  channelType: string;
  conversationId: string;
  /** The live session (meeting) the event is about. */
  sessionRef: string;
  kind: SessionEventKind;
  /** Present for 'participant_joined' / 'participant_left'. */
  participants?: readonly TeamsUserRef[];
  /** ISO 8601 timestamp of the platform event. */
  occurredAt: string;
}

export type SessionInputCapability =
  | 'text'
  | 'message_edits'
  | 'reactions'
  | 'card_answers'
  | 'presence'
  | 'hand_raise'
  | 'transcribed_speech'
  | 'shared_documents';

export type SessionOutputCapability =
  | 'text'
  | 'mentions'
  | 'cards'
  | 'polls'
  | 'private_messages'
  | 'pinned_surface'
  | 'voice'
  | 'screen';

export interface SessionCapabilitySnapshot {
  channelType: string;
  conversationId: string;
  /** The live session (meeting) the snapshot is about; absent for the
   *  conversation's asynchronous baseline. */
  sessionRef?: string;
  modality: 'chat' | 'online' | 'hybrid';
  inputs: readonly SessionInputCapability[];
  outputs: readonly SessionOutputCapability[];
  /** ISO 8601. */
  capturedAt: string;
}

/** `undefined` when the provider knows nothing about the conversation. */
export interface MeetingSourceProvider {
  channelType: string;
  snapshot(
    conversationId: string,
    opts: { asChannelKey?: string; sessionRef?: string },
  ): Promise<SessionCapabilitySnapshot | undefined>;
}

/** A person to @mention; `text` occurs in the message text. */
export interface TargetedMention {
  user: TeamsUserRef | { kind: string; id: string; displayName?: string; email?: string };
  /** The display text standing in the message (e.g. '@Anna Köhler'). */
  text: string;
}

/** A live poll riding on a message (`TargetedMessage.poll`, channel-sdk
 *  0.3.0 / kernel K3). The kernel validates it before the provider sees it. */
export interface TargetedPoll {
  /** 1..128 characters; round-tripped as `cardAnswer.cardId`. */
  pollId: string;
  question: string;
  /** 2..6 options, ids unique. */
  options: readonly { id: string; label: string }[];
  hint?: string;
}

/**
 * What a send provider declares it implements (channel-sdk
 * `SendProviderCapabilities`). The kernel does not trust a provider to honour
 * a send option it did not declare: without `asChannelKey: true` it never
 * calls the provider for an own-identity send and answers `no_identity`.
 * Absent = the provider implements none.
 */
export interface SendProviderCapabilities {
  /** Honours `opts.asChannelKey`: posts as exactly that identity or returns
   *  `unreachable`, never as another one. */
  asChannelKey?: boolean;
  /** Turns `message.mentions` into platform mentions. */
  mentions?: boolean;
  /** Posts `message.poll` as a poll and refuses an invalid one. */
  poll?: boolean;
  /**
   * channel-sdk 0.3.2 — the channel marks EVERY message of these sends as
   * AI-generated with its own label (Teams: the `AIGeneratedContent`
   * entity). The kernel then adds no disclosure sentence to the text;
   * without the declaration it folds the instance's disclosure line into the
   * first message an agent posts as itself into a conversation.
   */
  aiLabel?: boolean;
}

/** Third argument of `TargetedSendProvider.sendToUser`. */
export interface TargetedSendOptions {
  /** Identity the DM must come from (`28:<appId>`). Never substituted. */
  asChannelKey?: string;
}

/** `TargetedDeliveryOutcome` including the 0.3.0 additions (`messageId`,
 *  `cardMessageId`, `app_not_installed`, `blocked`). A poll may go out as two
 *  messages: `cardMessageId` is the card, and the `messageId` of every
 *  `card_answer` event for it. */
export type FacilitatorDeliveryOutcome =
  | { outcome: 'delivered'; messageId?: string; cardMessageId?: string }
  | {
      outcome: 'unreachable';
      code:
        | 'no_binding'
        | 'user_not_found'
        | 'channel_error'
        | 'not_permitted'
        | 'app_not_installed'
        | 'blocked';
      message: string;
    };

/** The CoreApi members this plugin feature-detects for #2043. */
export interface NativeFacilitatorCoreApi {
  emitConversationMessage?(event: ConversationMessageEvent): void;
  /** Synchronous and kernel-cached; `false` for anything not positively
   *  subscribed. Absent means the mention gate always applies. */
  isAllMessagesSubscribed?(channelType: string, conversationId: string, channelKey: string): boolean;
  registerReadinessProvider?(channelId: string, provider: ChannelReadinessProvider): void;
  registerMeetingSourceProvider?(channelId: string, provider: MeetingSourceProvider): void;
  emitSessionCapabilities?(snapshot: SessionCapabilitySnapshot): void;
  /** Meeting start / end and participants joining / leaving (kernel K5). */
  emitSessionEvent?(event: SessionEvent): void;
}

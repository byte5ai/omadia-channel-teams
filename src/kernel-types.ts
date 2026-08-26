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
}

export type ChatParticipantsProvider = () => Promise<ChatParticipant[]>;

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

export interface ChatTurnInput {
  userMessage: string;
  sessionScope?: string;
  userId?: string;
  priorTurns?: Array<{ userMessage: string; assistantAnswer: string }>;
  extraSystemHint?: string;
  freshCheck?: boolean;
  ssoAssertion?: string;
  userTimeZone?: string;
}

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

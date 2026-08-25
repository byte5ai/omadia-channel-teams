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

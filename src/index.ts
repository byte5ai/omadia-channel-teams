// ---------------------------------------------------------------------------
// Channel plugin entry point
// ---------------------------------------------------------------------------
// Phase 5B: standalone activate() shape replaces the legacy class. The
// dynamic-channel-resolver picks `activate` off the module export
// (`mod.activate ?? mod.default?.activate ?? mod.default`).
export { activate } from './plugin.js';

// ---------------------------------------------------------------------------
// Bot runtime + adapters (re-exported for the kernel bootstrap call-site
// + the answer-card tests under test/)
// ---------------------------------------------------------------------------
export { RETIRED_CARD_ACTION_TEXT, TeamsBot, teamsSessionScope, teamsTurnOrigin } from './teamsBot.js';

// ---------------------------------------------------------------------------
// #860 W2 — auto-invite agent apps (issues #20/#21/#22). Config surface,
// installer + state primitives, onboarding hook and card builders —
// re-exported for the kernel/provisioner call-sites and the tests under
// tests/ (scripts/test.mjs runs them against the built dist/).
// ---------------------------------------------------------------------------
export {
  parseTeamsAgentAppsConfig,
  TeamsAgentAppsConfigError,
  type TeamsAgentApp,
} from './teamsAgentApps.js';
export {
  AUTO_INSTALL_MARKER_TTL_MS,
  CONSENT_NEGATIVE_CACHE_TTL_MS,
  DEFAULT_THROTTLE_WAIT_MS,
  isConsentMissingError,
  isProvisioningThrottledError,
  MAX_RUN_SLEEP_BUDGET_MS,
  MAX_THROTTLE_RETRIES,
  MAX_THROTTLE_WAIT_MS,
  TeamsAgentInstaller,
  TeamsAutoInstallMarker,
  TeamsConsentNegativeCache,
  teamsAgentAppLogLabel,
  type AgentAppFallbackReason,
  type AgentAppFailureReason,
  type AgentAppInstallOutcome,
  type InstallAgentAppsRequest,
  type TeamsAgentAppTarget,
  type TeamsAgentInstallerOptions,
  type TeamsAgentInstallResult,
} from './teamsAgentInstaller.js';
export {
  handleTeamsAgentAppsRecheck,
  runTeamsAutoInviteHook,
  shouldSuppressAutoInstallIntro,
  teamsTeamScopeFromActivity,
  AGENT_APPS_RECHECK_UNSCOPED_MESSAGE,
  type TeamsAgentAppsRecheckResult,
  type TeamsAutoInviteDeps,
  type TeamsAutoInviteHookResult,
  type TeamsAutoInviteTurnContext,
} from './teamsBot.js';
export {
  TeamsAttachmentStore,
  graphMessageLookup,
  type TeamsAttachmentStoreOptions,
  type PersistedAttachment,
  type PersistTurnInput,
} from './teamsAttachmentStore.js';
export { TeamsRosterProvider } from './teamsRoster.js';
// #330 B2 — group-conversation adapters (kernel roster / targeted-send /
// membership-event seams). Re-exported for the tests under tests/.
export {
  attributeGroupMessage,
  createTeamsConversationSendAdapter,
  createTeamsRosterAdapter,
  createTeamsTargetedSendAdapter,
  TeamsConversationReferenceCache,
  toSdkConversationType,
  type ConversationPostOutcome,
} from './teamsGroupPrimitives.js';
// #2043 — native facilitators: authenticated inbound message events and the
// live poll card. Re-exported for the tests under tests/.
export {
  plainMessageText,
  platformAuthor,
  readQuoteInfo,
  receivingBotKey,
  TeamsConversationMessages,
  type TeamsConversationMessageSink,
} from './teamsConversationMessages.js';
export {
  buildPollCard,
  buildPollText,
  cardActionData,
  parsePollAnswerValue,
  POLL_ANSWER_VALUE_TYPE,
  POLL_ANSWER_VERB,
  validatePollInput,
  type BuildPollCardInput,
  type PollAnswerValue,
  type PollOption,
} from './teamsPollCard.js';
export {
  registerTeamsSendProviders,
  stampChannel,
  TEAMS_CHANNEL_TYPE,
  type TeamsSendProviderDeps,
} from './teamsKernelRegistration.js';
export {
  readTeamsAppTemplateInfo,
  templateRevisionOfAppVersion,
  type TeamsAppTemplateInfo,
} from './teamsAppTemplate.js';
export {
  APP_CATALOG_GRAPH_PERMISSION,
  CHANNEL_READ_ALL_RSC_PERMISSION,
  CHAT_CONSENT_GRAPH_PERMISSION,
  CHAT_INSTALL_READ_GRAPH_PERMISSION,
  CHAT_READ_ALL_RSC_PERMISSION,
  createTeamsReadinessProvider,
  TEAM_INSTALL_READ_GRAPH_PERMISSION,
  type ConsentRepairOutcome,
  type TeamsReadinessDeps,
  type TeamsReadinessGraph,
  type TeamsReadinessProvider,
} from './teamsReadiness.js';
export {
  createTeamsMeetingSource,
  isTeamsMeetingChat,
  MEETING_LIFECYCLE_RSC_PERMISSION,
  MEETING_PARTICIPANT_RSC_PERMISSION,
  type TeamsMeetingSourceDeps,
} from './teamsMeetingSource.js';
export {
  MEETING_END_EVENT,
  MEETING_PARTICIPANT_JOIN_EVENT,
  MEETING_PARTICIPANT_LEAVE_EVENT,
  MEETING_START_EVENT,
  participantsOf,
  sessionEventKind,
  TeamsMeetingSessions,
  type TeamsMeetingSessionSink,
  type TeamsMeetingState,
  type TeamsSessionEvent,
  type TeamsSessionEventKind,
} from './teamsMeetingSessions.js';
export { PgTeamsConversationRefStore } from './teamsConversationRefStore.js';
export type { TeamsConversationRefPersistence } from './teamsConversationRefStore.js';

// ---------------------------------------------------------------------------
// #860 W0a — per-bot identity + teams_bots[] config surface (config-wiring).
// Re-exported for the kernel/provisioner call-sites and the tests under
// tests/ (scripts/test.mjs runs them against the built dist/).
// ---------------------------------------------------------------------------
export {
  DEFAULT_TEAMS_BOT_APP_TYPE,
  findTeamsBotByAppId,
  findTeamsBotBySlug,
  getDefaultTeamsBot,
  parseTeamsBotKey,
  TEAMS_BOT_KEY_PREFIX,
  teamsBotKey,
  teamsBotLogLabel,
  type TeamsBotAppType,
  type TeamsBotIdentity,
} from './teamsBotIdentity.js';
export {
  LEGACY_TEAMS_BOT_DISPLAY_NAME,
  LEGACY_TEAMS_BOT_SECRET_REF,
  LEGACY_TEAMS_BOT_SLUG,
  legacyTeamsBotFromScalars,
  mergeLegacyTeamsBot,
  parseTeamsBotsConfig,
  TeamsBotsConfigError,
  type LegacyTeamsBotScalars,
} from './teamsBotsConfig.js';

// ---------------------------------------------------------------------------
// Channels-directory contribution + Graph name/member resolution
// (re-exported for the tests under tests/)
// ---------------------------------------------------------------------------
export { buildTeamsChannelKeyDirectory } from './channelKeyDirectory.js';
// People directory (Conductor human-step principal resolution via Graph /users)
export { createTeamsPeopleDirectoryProvider } from './teamsPeopleDirectory.js';
export { GraphRequestError, type GraphUser } from './teamsGraphResolver.js';
export {
  TeamsConversationObserver,
  type TeamsConversation,
} from './teamsConversationObserver.js';
export {
  TeamsGraphResolver,
  type ConversationToResolve,
  type ResolvedConversation,
  type TeamsGraphResolverOptions,
} from './teamsGraphResolver.js';

// ---------------------------------------------------------------------------
// Mentions
// ---------------------------------------------------------------------------
export {
  resolveMentions,
  stripMentionTokens,
  type MentionEntity,
  type MentionResolution,
} from './teamsMentions.js';

// ---------------------------------------------------------------------------
// Adaptive-Card builders + parsers
// ---------------------------------------------------------------------------
export {
  aiChipDue,
  aiLabelEntity,
  AI_MARKER_CHIP,
  buildAgentAppsResultCard,
  buildAnswerCard,
  buildChoiceAskCard,
  buildFollowUpsOnlyCard,
  buildSlotPickerCard,
  buildTopicAskCard,
  parseAgentAppsRecheckValue,
  parseBookSlotValue,
  parseChoiceAskValue,
  parseFollowUpValue,
  parseTopicDecisionValue,
  stripFoldedAiDisclosure,
  AGENT_APPS_RECHECK_VALUE_TYPE,
  type AgentAppsRecheckValue,
  type BuildAgentAppsCardInput,
  type BuildAnswerCardInput,
  type BuildChoiceAskCardInput,
  type ChoiceAskValue,
  type FollowUpValue,
  type BookSlotValue,
  type TopicAskCardInput,
  type TopicDecisionChoice,
  type TopicDecisionValue,
  CHOICE_ASK_VALUE_TYPE,
  FOLLOW_UP_VALUE_TYPE,
  BOOK_SLOT_VALUE_TYPE,
  TOPIC_DECISION_VALUE_TYPE,
  TEAMS_CARD_CONTENT_TYPE,
} from './teamsCard.js';

// ---------------------------------------------------------------------------
// Routers (will move to plugin-side `ctx.routes.register` in phase-3.1-4 —
// exported here for the kernel bootstrap call-site while the intermediate
// state persists)
// ---------------------------------------------------------------------------
export {
  createTeamsRouter,
  type TeamsProactiveSend,
  type TeamsRouterBotCredentials,
  type TeamsRouterDeps,
} from './messagesRouter.js';
export { createAttachmentsRouter } from './attachmentsRouter.js';

// ---------------------------------------------------------------------------
// Attachment signing — exported for the kernel-mounted /attachments proxy
// route + the attachmentSigning test suite
// ---------------------------------------------------------------------------
export {
  signAttachmentUrl,
  verifyAttachmentSig,
  type SignAttachmentUrlParams,
  type VerifyAttachmentSigParams,
} from './attachmentSigning.js';

// ---------------------------------------------------------------------------
// byte5ai/omadia#1596 / #1620 — Hub / Tab-Config shell router
// (`auth:'public'`) and the catalogue router (`auth:'teams-sso'`; the kernel
// verifies the Teams SSO token). Exported for tests/teamsUiRouter.test.ts,
// which runs against the built dist/.
// ---------------------------------------------------------------------------
export {
  createTeamsCatalogRouter,
  createTeamsUiShellRouter,
  TAB_CONFIG_HUB_VALUE,
  tabConfigChoices,
  type DiscoveredUiRoute,
} from './uiRouter.js';

// #2043 — which conversation id goes where (kernel key vs. Graph).
export {
  bareConversationId,
  isBotKeyPrefixedConversationId,
  isChannelConversationId,
  threadRootMessageId,
} from './teamsConversationIds.js';

// spec 2043 §5.14 — the bot's own reaction on a message (Bot Connector).
export { TEAMS_REACTION_TYPES, teamsReactionUrl, type TeamsReactionFetch } from './teamsReactions.js';

// AI marking of deliveries outside a chat turn (routines, approvals).
export { buildRoutineDeliveryActivity, createTeamsRoutineDelivery } from './teamsRoutineDelivery.js';
export { PROACTIVE_POSTS_CARRY_AI_LABEL } from './teamsOutboundMessage.js';

// #2089 — question cards (votes, consent, assignment, free text, questionnaires).
export { OVERDUE_GRACE_MS, TeamsQuestionCards, type QuestionCardAnswer } from './teamsQuestionCards.js';
export {
  buildClosedCardActivity,
  createOverdueCardSweep,
  OVERDUE_SWEEP_INTERVAL_MS,
  type OverdueCardSweepDeps,
} from './teamsQuestionCardSweep.js';
export {
  buildClosedCardContent,
  buildOwnAnswerCardContent,
  buildQuestionCardContent,
  QUESTION_CARD_ANSWER_VERB,
  QUESTION_CARD_CHANGE_VERB,
  QUESTION_CARD_DATA_TYPE,
  QUESTION_CARD_VIEW_VERB,
  questionCardActionData,
} from './teamsQuestionCardRender.js';
export { inputId, readCardSubmission, VOLUNTEER_PRESET } from './teamsQuestionCardAnswers.js';
export {
  MAX_CARD_QUESTIONS,
  questionCardProblem,
  RENDERED_CARD_KINDS,
  TEAMS_CARD_CAPABILITY,
} from './teamsQuestionCardSpec.js';

// D24 — how agent messages are marked as AI-generated, and the debug capture.
export {
  isOutgoingLogEnabled,
  LOG_OUTGOING_ENV,
  logOutgoingActivity,
  NATIVE_AI_LABEL_CONFIG_KEY,
  parseBooleanSetting,
  resolveNativeAiLabel,
} from './teamsAiMarkingMode.js';

// #2086 — the Teams side of the kernel's decoration registry, the answer's
// attributed parts (colleagues' contributions, system notices) and the details
// element of an answer card.
export {
  admitTeamsDecorations,
  TEAMS_DECORATION_IDS,
} from './teamsDecorations.js';
export {
  contributionLabel,
  contributionsAsText,
  SYSTEM_NOTICE_LABEL,
  systemNoticeAsText,
} from './teamsAnswerParts.js';
export { buildAnswerDetailsElements, ANSWER_DETAILS_CONTAINER_ID } from './teamsCardDetails.js';
// Spec 2042 §5.8 — pace: typing, one acknowledgement, no progress chatter.
export { TEAMS_PACE_PROFILE, teamsPaceHooks } from './teamsPace.js';
// #2089 — posted question cards kept across restarts.
export {
  expiresAt as questionCardExpiresAt,
  KEEP_AFTER_CLOSE_MS,
  KEEP_WITHOUT_CLOSE_MS,
  PgTeamsQuestionCardStore,
  QUESTION_CARDS_DDL,
  QUESTION_CARDS_TABLE,
  type StoredQuestionCard,
  type TeamsQuestionCardPersistence,
} from './teamsQuestionCardStore.js';

// D26 — what a bot saw of a meeting, kept across restarts.
export {
  believableMeetingState,
  MEETING_STATES_DDL,
  MEETING_STATES_TABLE,
  PgTeamsMeetingStateStore,
  RUNNING_EVIDENCE_MAX_AGE_MS,
  type StoredMeetingState,
  type TeamsMeetingStatePersistence,
} from './teamsMeetingStateStore.js';

export { incomingInvokeLogLine, invokeValueForLog, INVOKE_LOG_MAX_TEXT } from './teamsInvokeLog.js';

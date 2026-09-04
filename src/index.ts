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
export { TeamsBot, teamsSessionScope } from './teamsBot.js';

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
} from './teamsGroupPrimitives.js';
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
  aiLabelEntity,
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
  parseFreshCheckValue,
  parseTopicDecisionValue,
  stripFoldedAiDisclosure,
  AGENT_APPS_RECHECK_VALUE_TYPE,
  type AgentAppsRecheckValue,
  type BuildAgentAppsCardInput,
  type BuildAnswerCardInput,
  type BuildChoiceAskCardInput,
  type FreshCheckValue,
  type ChoiceAskValue,
  type FollowUpValue,
  type BookSlotValue,
  type TopicAskCardInput,
  type TopicDecisionChoice,
  type TopicDecisionValue,
  FRESH_CHECK_VALUE_TYPE,
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

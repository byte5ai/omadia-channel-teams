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
  buildAnswerCard,
  buildChoiceAskCard,
  buildFollowUpsOnlyCard,
  buildSlotPickerCard,
  buildTopicAskCard,
  parseBookSlotValue,
  parseChoiceAskValue,
  parseFollowUpValue,
  parseFreshCheckValue,
  parseTopicDecisionValue,
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
export { createTeamsRouter, type TeamsRouterDeps } from './messagesRouter.js';
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

// ---------------------------------------------------------------------------
// Channel plugin entry point
// ---------------------------------------------------------------------------
export { TeamsChannelPlugin } from './plugin.js';
export type { TeamsChannelPluginDeps } from './plugin.js';

// ---------------------------------------------------------------------------
// Bot runtime + adapters (re-exported for the kernel bootstrap call-site
// + the answer-card tests under test/)
// ---------------------------------------------------------------------------
export { TeamsBot } from './teamsBot.js';
export {
  TeamsAttachmentStore,
  type TeamsAttachmentStoreOptions,
  type PersistedAttachment,
  type PersistTurnInput,
} from './teamsAttachmentStore.js';
export { TeamsRosterProvider } from './teamsRoster.js';

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

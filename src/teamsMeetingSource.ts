// Meeting source for native facilitators (byte5ai/omadia#2043 §6.1 / §6.3,
// issue #44): what a Teams session offers the facilitator as inputs and
// outputs. The baseline is the chat of the meeting and cards; see
// docs/teams-meeting-apis.md for what each entry rests on and for everything
// that is deliberately absent (presence, hand raise, audio, surfaces).
//
// A capability is listed only when this plugin implements it. The snapshot
// says what the bot can do in the conversation, not what Teams could offer
// with more permissions.

import type {
  MeetingSourceProvider,
  SessionCapabilitySnapshot,
  SessionInputCapability,
  SessionOutputCapability,
} from './kernel-types.js';
import { parseTeamsBotKey } from './teamsBotIdentity.js';
import type { TeamsConversationReferenceCache } from './teamsGroupPrimitives.js';

/** Teams names the chat of a scheduled or instant meeting `19:meeting_…`. */
const MEETING_CHAT_PREFIX = '19:meeting_';

const CHAT_INPUTS: readonly SessionInputCapability[] = ['text', 'message_edits', 'reactions', 'card_answers'];
const CHAT_OUTPUTS: readonly SessionOutputCapability[] = ['text', 'mentions', 'cards', 'polls', 'private_messages'];
/** A 1:1 has nobody to mention and no group to poll. */
const DIRECT_OUTPUTS: readonly SessionOutputCapability[] = ['text', 'cards'];

export function isTeamsMeetingChat(conversationId: string): boolean {
  return conversationId.startsWith(MEETING_CHAT_PREFIX);
}

export interface TeamsMeetingSourceDeps {
  refs: Pick<TeamsConversationReferenceCache, 'getOrLoad'>;
  now?: () => Date;
}

export function createTeamsMeetingSource(deps: TeamsMeetingSourceDeps): MeetingSourceProvider {
  const now = deps.now ?? ((): Date => new Date());
  return {
    channelType: 'teams',
    async snapshot(conversationId, opts): Promise<SessionCapabilitySnapshot | undefined> {
      const botAppId = opts.asChannelKey ? parseTeamsBotKey(opts.asChannelKey) : undefined;
      // A named identity that is not a bot key knows nothing about any chat.
      if (opts.asChannelKey !== undefined && !botAppId) return undefined;
      const known = await deps.refs.getOrLoad(conversationId, botAppId).catch(() => undefined);
      if (!known) return undefined;
      const isDirect = known.teamsType === 'personal';
      return {
        channelType: 'teams',
        conversationId,
        ...(opts.sessionRef ? { sessionRef: opts.sessionRef } : {}),
        // 'hybrid' is never claimed: Teams does not tell a bot whether people
        // share a room.
        modality: isTeamsMeetingChat(conversationId) ? 'online' : 'chat',
        inputs: CHAT_INPUTS,
        outputs: isDirect ? DIRECT_OUTPUTS : CHAT_OUTPUTS,
        capturedAt: now().toISOString(),
      };
    },
  };
}

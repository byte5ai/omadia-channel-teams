// Meeting source for native facilitators (byte5ai/omadia#2043 §6.1 / §6.3,
// issue #44): what a Teams session offers the facilitator as inputs and
// outputs. The baseline is the chat of the meeting and cards; see
// docs/teams-meeting-apis.md for what each entry rests on and for everything
// that is deliberately absent (presence, hand raise, audio, surfaces).
//
// A capability is listed only when this plugin implements it. The snapshot
// says what the bot can do in the conversation, not what Teams could offer
// with more permissions.
//
//   always            chat (text, edits), cards and polls, mentions, private
//                     messages. The roster is available too, through the
//                     kernel's roster provider, not through this snapshot.
//   modality          `chat`, also for the chat of a meeting. `online` only
//                     while a meeting is known to be running there: between
//                     its start (or a join) and its end, as this bot saw it.
//   scheduled only    `presence`, and only while it is proven: a scheduled
//                     meeting is running, a join or leave event has arrived
//                     for it, and the install holds the two meeting RSC
//                     permissions (read from Graph). Ad-hoc calls and
//                     Meet-now keep the "always" set. The meeting lifecycle
//                     itself has no capability name in the contract; it
//                     reaches the kernel as session events and as a pushed
//                     snapshot at start and end.
//   never             `hand_raise` (no bot event exists) and
//                     `transcribed_speech` / `voice` (real-time media needs a
//                     .NET service on Windows; not on this stack).

import type {
  MeetingSourceProvider,
  SessionCapabilitySnapshot,
  SessionInputCapability,
  SessionOutputCapability,
} from './kernel-types.js';
import { parseTeamsBotKey, teamsBotKey } from './teamsBotIdentity.js';
import type { TeamsMeetingSessions } from './teamsMeetingSessions.js';
import type { TeamsConversationReferenceCache } from './teamsGroupPrimitives.js';

/** Teams names the chat of a scheduled or instant meeting `19:meeting_…`. */
const MEETING_CHAT_PREFIX = '19:meeting_';

// `reactions` is deliberately NOT declared. Teams documents the reaction
// event only for messages the bot itself sent; whether a reaction to another
// person's message reaches the bot is unproven, and the snapshot has no way
// to say "on my own messages only". Reaction events that do arrive are still
// emitted; an engine may use them for its own posts (confirming an echo) but
// must not build on reactions to participants' messages.
const CHAT_INPUTS: readonly SessionInputCapability[] = ['text', 'message_edits', 'card_answers'];
const CHAT_OUTPUTS: readonly SessionOutputCapability[] = ['text', 'mentions', 'cards', 'polls', 'private_messages'];
/** A 1:1 has nobody to mention and no group to poll. */
const DIRECT_OUTPUTS: readonly SessionOutputCapability[] = ['text', 'cards'];

export function isTeamsMeetingChat(conversationId: string): boolean {
  return conversationId.startsWith(MEETING_CHAT_PREFIX);
}

/** RSC a chat must have granted for meeting start/end and participant events. */
export const MEETING_LIFECYCLE_RSC_PERMISSION = 'OnlineMeeting.ReadBasic.Chat';
export const MEETING_PARTICIPANT_RSC_PERMISSION = 'OnlineMeetingParticipant.Read.Chat';

export interface TeamsMeetingSourceDeps {
  refs: Pick<TeamsConversationReferenceCache, 'getOrLoad'>;
  /** What has been seen of the meeting in a conversation, per bot. */
  sessions?: Pick<TeamsMeetingSessions, 'stateFor'> & Partial<Pick<TeamsMeetingSessions, 'stateOrLoad'>>;
  /** The RSC permissions granted to the bot's app in the chat, read from
   *  Graph; `undefined` when they cannot be read. */
  grantedPermissions?(conversationId: string, asChannelKey: string): Promise<readonly string[] | undefined>;
  now?: () => Date;
}

export function createTeamsMeetingSource(deps: TeamsMeetingSourceDeps): MeetingSourceProvider {
  const now = deps.now ?? ((): Date => new Date());

  /** What the bot saw of the meeting: from memory, else from the stored
   *  state (after a restart), where the sessions offer it. */
  const meetingState = async (conversationId: string, botKey: string) =>
    deps.sessions?.stateOrLoad
      ? deps.sessions.stateOrLoad(conversationId, botKey)
      : deps.sessions?.stateFor(conversationId, botKey);

  /**
   * `presence` is declared only when ALL of this holds, none of it assumed:
   *   - a scheduled meeting is running in this conversation (seen: its start
   *     event, or a participant event, which exists for scheduled meetings
   *     only);
   *   - a join or leave event has actually arrived for this meeting. The
   *     per-app "Meeting event subscriptions" setting in Developer Portal has
   *     no documented manifest field or API, so the only proof that it is set
   *     is an event;
   *   - the install in this chat really holds both meeting RSC permissions
   *     (the grant is read from Graph).
   * Meet-now and calls never qualify: Teams sends them no participant events.
   */
  const hasPresence = async (conversationId: string, asChannelKey: string | undefined): Promise<boolean> => {
    if (!asChannelKey || !deps.sessions || !deps.grantedPermissions) return false;
    const state = await meetingState(conversationId, asChannelKey);
    if (!state?.isLive || !state.isScheduled || !state.hasSeenParticipantEvent) return false;
    const granted = await deps.grantedPermissions(conversationId, asChannelKey).catch(() => undefined);
    return (
      granted !== undefined &&
      granted.includes(MEETING_LIFECYCLE_RSC_PERMISSION) &&
      granted.includes(MEETING_PARTICIPANT_RSC_PERMISSION)
    );
  };

  return {
    channelType: 'teams',
    async snapshot(conversationId, opts): Promise<SessionCapabilitySnapshot | undefined> {
      const botAppId = opts.asChannelKey ? parseTeamsBotKey(opts.asChannelKey) : undefined;
      // A named identity that is not a bot key knows nothing about any chat.
      if (opts.asChannelKey !== undefined && !botAppId) return undefined;
      const known = await deps.refs.getOrLoad(conversationId, botAppId).catch(() => undefined);
      if (!known) return undefined;
      const isDirect = known.teamsType === 'personal';
      const botKey = botAppId ? teamsBotKey(botAppId) : undefined;
      const meeting = botKey ? await meetingState(conversationId, botKey) : undefined;
      const sessionRef = opts.sessionRef ?? (meeting?.isLive ? meeting.sessionRef : undefined);
      const inputs = (await hasPresence(conversationId, botKey))
        ? [...CHAT_INPUTS, 'presence' as const]
        : CHAT_INPUTS;
      return {
        channelType: 'teams',
        conversationId,
        ...(sessionRef ? { sessionRef } : {}),
        // `online` is a statement that a meeting is RUNNING, and an engine
        // acts on it. It is made only on evidence: this bot saw the meeting
        // start (or a participant join it) and has not seen it end. The chat
        // of a meeting is an ordinary chat before, after and between
        // sessions, and also when the state is unknown, as after a restart.
        // Microsoft documents no way to read "is this meeting running" on
        // demand (see docs/teams-meeting-apis.md), so there is nothing to
        // ask; unknown is reported as `chat`, never as a meeting.
        // 'hybrid' is never claimed: Teams does not tell a bot whether people
        // share a room.
        modality: meeting?.isLive ? 'online' : 'chat',
        inputs,
        outputs: isDirect ? DIRECT_OUTPUTS : CHAT_OUTPUTS,
        capturedAt: now().toISOString(),
      };
    },
  };
}

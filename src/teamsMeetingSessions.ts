// Meeting lifecycle and participant events (byte5ai/omadia#2043 §6.3, issue
// #44). Teams sends four `event` activities into the chat of a meeting:
//
//   application/vnd.microsoft.meetingStart            RSC OnlineMeeting.ReadBasic.Chat
//   application/vnd.microsoft.meetingEnd              (channel: ChannelMeeting.ReadBasic.Group)
//   application/vnd.microsoft.meetingParticipantJoin  RSC OnlineMeetingParticipant.Read.Chat
//   application/vnd.microsoft.meetingParticipantLeave AND the per-app "Meeting event
//                                                     subscriptions" in Developer Portal
//
// Names and payload fields follow Microsoft Learn "meeting-apps-apis", read
// 2026-11-18. Start and end are documented for scheduled and channel meetings,
// participant events for scheduled meetings only. None of it was observed on
// a live tenant.
//
// This module turns the activities into session events for the kernel and
// remembers, per bot and conversation, what it has actually SEEN: whether a
// meeting is running, of which type, and whether a join or leave event ever
// arrived for it. The capability snapshot is built from that record, never
// from what the app package requests.

import type { Activity } from 'botbuilder';

import type { SessionEvent, SessionEventKind, TeamsUserRef } from './kernel-types.js';
import { receivingBotKey } from './teamsConversationMessages.js';

export const MEETING_START_EVENT = 'application/vnd.microsoft.meetingStart';
export const MEETING_END_EVENT = 'application/vnd.microsoft.meetingEnd';
export const MEETING_PARTICIPANT_JOIN_EVENT = 'application/vnd.microsoft.meetingParticipantJoin';
export const MEETING_PARTICIPANT_LEAVE_EVENT = 'application/vnd.microsoft.meetingParticipantLeave';

const CHANNEL_TYPE = 'teams';
/** Replaced with the catalog id by plugin.ts; the bot cannot know it. */
const CHANNEL_ID_PLACEHOLDER = 'de.byte5.channel.teams';
const MAX_TRACKED_MEETINGS = 500;
/** `value.MeetingType` values of a meeting created in the calendar. */
const SCHEDULED_MEETING_TYPES: readonly string[] = ['scheduled', 'recurring'];

/** The contract's session event; these names are kept for this module's API. */
export type TeamsSessionEventKind = SessionEventKind;
export type TeamsSessionEvent = SessionEvent & { participants?: TeamsUserRef[] };

/** What has been seen of the meeting in one conversation, for one bot. */
export interface TeamsMeetingState {
  sessionRef: string;
  /** `value.MeetingType` of the start event, when one arrived. */
  meetingType?: string;
  isScheduled: boolean;
  isLive: boolean;
  /** A join or leave event arrived for THIS meeting: the per-app event
   *  subscription is evidently set. */
  hasSeenParticipantEvent: boolean;
}

type InboundActivity = Partial<Activity>;

interface MeetingValue {
  MeetingType?: unknown;
  StartTime?: unknown;
  EndTime?: unknown;
  members?: unknown;
}

function isoOrUndefined(raw: unknown): string | undefined {
  if (raw instanceof Date) return raw.toISOString();
  if (typeof raw === 'string' && !Number.isNaN(Date.parse(raw))) return new Date(raw).toISOString();
  return undefined;
}

/** The meeting id Teams asks apps to use: `channelData.meeting.id`. The id in
 *  `value.Id` of a start or end event is explicitly not to be used. */
function sessionRefOf(activity: InboundActivity): string | undefined {
  const meetingId = (activity.channelData as { meeting?: { id?: unknown } } | undefined)?.meeting?.id;
  if (typeof meetingId === 'string' && meetingId.trim().length > 0) return meetingId.trim();
  return activity.conversation?.id;
}

/** `value.members[].user` of a participant event as kernel user references. */
export function participantsOf(activity: InboundActivity): TeamsUserRef[] {
  const members = (activity.value as MeetingValue | undefined)?.members;
  if (!Array.isArray(members)) return [];
  return members.flatMap((member) => {
    const user = (member as { user?: { id?: unknown; aadObjectId?: unknown; name?: unknown } } | null)?.user;
    const aadObjectId = typeof user?.aadObjectId === 'string' ? user.aadObjectId.trim() : '';
    const teamsId = typeof user?.id === 'string' ? user.id.trim() : '';
    const id = aadObjectId || teamsId;
    if (!id) return [];
    const name = typeof user?.name === 'string' ? user.name.trim() : '';
    return [{ kind: 'teams-aad' as const, id, ...(name ? { displayName: name } : {}) }];
  });
}

export function sessionEventKind(eventName: string | undefined): TeamsSessionEventKind | undefined {
  switch (eventName?.trim()) {
    case MEETING_START_EVENT:
      return 'session_started';
    case MEETING_END_EVENT:
      return 'session_ended';
    case MEETING_PARTICIPANT_JOIN_EVENT:
      return 'participant_joined';
    case MEETING_PARTICIPANT_LEAVE_EVENT:
      return 'participant_left';
    default:
      return undefined;
  }
}

export interface TeamsMeetingSessionSink {
  /** `core.emitSessionEvent`, already feature-detected; absent on a kernel
   *  without it. The record below is kept either way. */
  emit?(event: TeamsSessionEvent): void;
  /** The meeting started or ended: what the conversation offers changed. */
  onLifecycleChange?(conversationId: string, botKey: string): void;
}

/** One instance serves every bot of the plugin; state is keyed by the bot. */
export class TeamsMeetingSessions {
  private readonly meetings = new Map<string, TeamsMeetingState>();

  constructor(private readonly sink: TeamsMeetingSessionSink = {}) {}

  /** What this bot has seen of the meeting in this conversation. */
  stateFor(conversationId: string, botKey: string): TeamsMeetingState | undefined {
    return this.meetings.get(`${botKey}|${conversationId}`);
  }

  /** Record one meeting event activity and pass it to the kernel. Returns the
   *  event, or `undefined` for an activity that is not a usable meeting event.
   *  Never throws: a meeting event must not break the turn that carried it. */
  handle(activity: InboundActivity): TeamsSessionEvent | undefined {
    try {
      const kind = sessionEventKind(activity.name);
      const conversationId = activity.conversation?.id;
      const botKey = receivingBotKey(activity);
      const sessionRef = sessionRefOf(activity);
      if (!kind || !conversationId || !botKey || !sessionRef) return undefined;

      const key = `${botKey}|${conversationId}`;
      const previous = this.meetings.get(key);
      const value = (activity.value ?? {}) as MeetingValue;
      const isParticipantEvent = kind === 'participant_joined' || kind === 'participant_left';
      const participants = isParticipantEvent ? participantsOf(activity) : [];
      if (isParticipantEvent && participants.length === 0) return undefined;

      // A new meeting in the same chat starts a new record; anything else
      // continues the one we have.
      const isSameMeeting = previous?.sessionRef === sessionRef;
      const base: TeamsMeetingState = isSameMeeting && previous
        ? previous
        : { sessionRef, isScheduled: false, isLive: false, hasSeenParticipantEvent: false };
      const meetingType = typeof value.MeetingType === 'string' ? value.MeetingType : base.meetingType;
      const next: TeamsMeetingState = {
        ...base,
        ...(meetingType ? { meetingType } : {}),
        // Participant events exist for scheduled meetings only, so one
        // arriving is itself the statement that this meeting is scheduled.
        isScheduled:
          isParticipantEvent ||
          (meetingType ? SCHEDULED_MEETING_TYPES.includes(meetingType.toLowerCase()) : base.isScheduled),
        isLive: kind === 'session_ended' ? false : kind === 'session_started' ? true : base.isLive || isParticipantEvent,
        hasSeenParticipantEvent: kind === 'session_started' ? false : base.hasSeenParticipantEvent || isParticipantEvent,
      };
      this.meetings.delete(key);
      this.meetings.set(key, next);
      if (this.meetings.size > MAX_TRACKED_MEETINGS) {
        const oldest = this.meetings.keys().next().value;
        if (oldest !== undefined) this.meetings.delete(oldest);
      }

      const occurredAt =
        (kind === 'session_started' ? isoOrUndefined(value.StartTime) : undefined) ??
        (kind === 'session_ended' ? isoOrUndefined(value.EndTime) : undefined) ??
        isoOrUndefined(activity.timestamp) ??
        new Date().toISOString();
      const event: TeamsSessionEvent = {
        channelId: CHANNEL_ID_PLACEHOLDER,
        channelType: CHANNEL_TYPE,
        conversationId,
        sessionRef,
        kind,
        ...(isParticipantEvent ? { participants } : {}),
        occurredAt,
      };
      this.sink.emit?.(event);
      const isLifecycle = kind === 'session_started' || kind === 'session_ended';
      // The first participant event of a meeting also changes the snapshot.
      if (isLifecycle || (isParticipantEvent && !base.hasSeenParticipantEvent)) {
        this.sink.onLifecycleChange?.(conversationId, botKey);
      }
      return event;
    } catch (err) {
      console.error('[teams] meeting event handling failed:', err instanceof Error ? err.message : err);
      return undefined;
    }
  }
}

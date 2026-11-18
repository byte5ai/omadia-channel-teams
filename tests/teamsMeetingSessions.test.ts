import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  createTeamsMeetingSource,
  MEETING_LIFECYCLE_RSC_PERMISSION,
  MEETING_PARTICIPANT_RSC_PERMISSION,
  TeamsConversationReferenceCache,
  TeamsMeetingSessions,
} from '@omadia/channel-teams';
import type { TeamsSessionEvent } from '@omadia/channel-teams';

import {
  ANNA,
  BERT,
  CARLA,
  GROUP_CHAT_ID,
  JONAS,
  MEETING_CHAT_ID,
  MEETING_ID,
  meetingLifecycleActivity,
  meetingParticipantActivity,
  MIRA,
  SimulatedKernel,
  simulatedBot,
} from './support/teamsSimulation.js';

// byte5ai/omadia#2043 §6.3, issue #44 — meeting lifecycle and participant
// events. Payload shapes follow Microsoft Learn "meeting-apps-apis"; none of
// it was observed on a live tenant.

interface Harness {
  events: TeamsSessionEvent[];
  pushes: string[];
  sessions: TeamsMeetingSessions;
  refs: TeamsConversationReferenceCache;
  granted: string[] | undefined;
  grantReads: number;
  source: ReturnType<typeof createTeamsMeetingSource>;
  deliver: ReturnType<typeof simulatedBot>['deliver'];
  sent: unknown[];
  turns: string[];
}

function harness(): Harness {
  const events: TeamsSessionEvent[] = [];
  const pushes: string[] = [];
  const refs = new TeamsConversationReferenceCache();
  const sessions = new TeamsMeetingSessions({
    emit: (event) => {
      events.push(event);
    },
    onLifecycleChange: (conversationId, botKey) => {
      pushes.push(`${botKey}|${conversationId}`);
    },
  });
  const h = {
    events,
    pushes,
    sessions,
    refs,
    granted: [MEETING_LIFECYCLE_RSC_PERMISSION, MEETING_PARTICIPANT_RSC_PERMISSION, 'ChatMessage.Read.Chat'] as string[] | undefined,
    grantReads: 0,
  } as Harness;
  h.source = createTeamsMeetingSource({
    refs,
    sessions,
    grantedPermissions: async () => {
      h.grantReads += 1;
      return h.granted;
    },
    now: () => new Date('2026-11-18T10:30:00.000Z'),
  });
  const sim = simulatedBot(new SimulatedKernel(), {
    captureConversationReference: (context: TurnContext) => refs.capture(context),
    emitMembershipEvent: () => {},
    meetingSessions: sessions,
  });
  h.deliver = sim.deliver;
  h.sent = sim.sent;
  h.turns = sim.turns;
  return h;
}

const inputsOf = async (h: Harness, conversationId = MEETING_CHAT_ID, bot = MIRA): Promise<readonly string[] | undefined> =>
  (await h.source.snapshot(conversationId, { asChannelKey: bot.key }))?.inputs;

describe('#44 meeting start and end', () => {
  it('emits session_started and session_ended with the meeting id from channelData', async () => {
    const h = harness();

    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start', at: '2026-11-18T10:00:00.000Z' }));
    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'end', at: '2026-11-18T10:45:00.000Z' }));

    assert.deepEqual(h.events, [
      {
        channelId: 'de.byte5.channel.teams',
        channelType: 'teams',
        conversationId: MEETING_CHAT_ID,
        sessionRef: MEETING_ID,
        kind: 'session_started',
        occurredAt: '2026-11-18T10:00:00.000Z',
      },
      {
        channelId: 'de.byte5.channel.teams',
        channelType: 'teams',
        conversationId: MEETING_CHAT_ID,
        sessionRef: MEETING_ID,
        kind: 'session_ended',
        occurredAt: '2026-11-18T10:45:00.000Z',
      },
    ]);
    assert.deepEqual(h.turns, [], 'a meeting event is not a turn');
    assert.deepEqual(h.sent, [], 'and the bot says nothing');
  });

  it('handles the event name with the leading space the documentation prints', async () => {
    const h = harness();

    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start', paddedName: true }));

    assert.deepEqual(h.events.map((e) => e.kind), ['session_started']);
  });

  it('pushes a fresh snapshot at start and at end', async () => {
    const h = harness();

    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));
    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'end' }));

    assert.deepEqual(h.pushes, [`${MIRA.key}|${MEETING_CHAT_ID}`, `${MIRA.key}|${MEETING_CHAT_ID}`]);
  });

  it('falls back to the chat id as session reference when the activity names no meeting', async () => {
    const h = harness();
    const start = meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' });

    await h.deliver({ ...start, channelData: { tenant: { id: 'x' } } });

    assert.equal(h.events[0]?.sessionRef, MEETING_CHAT_ID, 'never value.Id, which Teams says not to use');
  });

  it('keeps each bot\'s record apart in one meeting chat', async () => {
    const h = harness();

    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));

    assert.equal(h.sessions.stateFor(MEETING_CHAT_ID, MIRA.key)?.isLive, true);
    assert.equal(h.sessions.stateFor(MEETING_CHAT_ID, JONAS.key), undefined);
  });
});

describe('#44 participants joining and leaving', () => {
  it('emits the people with their Entra object id, trimmed', async () => {
    const h = harness();
    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));

    await h.deliver(meetingParticipantActivity({ bot: MIRA, from: ANNA, phase: 'join', people: [ANNA, BERT], at: '2026-11-18T10:01:00.000Z' }));
    await h.deliver(meetingParticipantActivity({ bot: MIRA, from: BERT, phase: 'leave', people: [BERT], at: '2026-11-18T10:20:00.000Z' }));

    assert.deepEqual(
      h.events.slice(1).map((e) => [e.kind, e.sessionRef, e.participants, e.occurredAt]),
      [
        [
          'participant_joined',
          MEETING_ID,
          [
            { kind: 'teams-aad', id: ANNA.aadObjectId, displayName: ANNA.name },
            { kind: 'teams-aad', id: BERT.aadObjectId, displayName: BERT.name },
          ],
          '2026-11-18T10:01:00.000Z',
        ],
        ['participant_left', MEETING_ID, [{ kind: 'teams-aad', id: BERT.aadObjectId, displayName: BERT.name }], '2026-11-18T10:20:00.000Z'],
      ],
    );
  });

  it('a late joiner arrives as participant_joined with the time of joining', async () => {
    const h = harness();
    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start', at: '2026-11-18T10:00:00.000Z' }));
    await h.deliver(meetingParticipantActivity({ bot: MIRA, from: ANNA, phase: 'join', people: [ANNA], at: '2026-11-18T10:00:30.000Z' }));

    await h.deliver(meetingParticipantActivity({ bot: MIRA, from: CARLA, phase: 'join', people: [CARLA], at: '2026-11-18T10:25:00.000Z' }));

    const late = h.events.at(-1);
    assert.equal(late?.kind, 'participant_joined');
    assert.deepEqual(late?.participants, [{ kind: 'teams-aad', id: CARLA.aadObjectId, displayName: CARLA.name }]);
    assert.equal(late?.occurredAt, '2026-11-18T10:25:00.000Z');
  });

  it('ignores a participant event that names nobody', async () => {
    const h = harness();
    await h.deliver(meetingParticipantActivity({ bot: MIRA, from: ANNA, phase: 'join', people: [] }));
    assert.deepEqual(h.events, []);
  });

  it('works without a kernel that takes session events: the record is kept', async () => {
    const sessions = new TeamsMeetingSessions();
    sessions.handle(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));
    assert.equal(sessions.stateFor(MEETING_CHAT_ID, MIRA.key)?.isLive, true);
  });
});

describe('#44 capability snapshot: presence only when it is proven', () => {
  it('declares presence for a running scheduled meeting once a participant event arrived and the grants are there', async () => {
    const h = harness();
    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));
    assert.deepEqual(await inputsOf(h), ['text', 'message_edits', 'card_answers'], 'start alone proves no participant events');

    await h.deliver(meetingParticipantActivity({ bot: MIRA, from: ANNA, phase: 'join', people: [ANNA] }));

    const snapshot = await h.source.snapshot(MEETING_CHAT_ID, { asChannelKey: MIRA.key });
    assert.deepEqual(snapshot?.inputs, ['text', 'message_edits', 'card_answers', 'presence']);
    assert.equal(snapshot?.modality, 'online');
    assert.equal(snapshot?.sessionRef, MEETING_ID);
    assert.equal(h.pushes.length, 2, 'the first participant event pushes a fresh snapshot too');
  });

  it('does not declare presence when the install does not hold both meeting permissions', async () => {
    const h = harness();
    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));
    await h.deliver(meetingParticipantActivity({ bot: MIRA, from: ANNA, phase: 'join', people: [ANNA] }));

    h.granted = [MEETING_LIFECYCLE_RSC_PERMISSION, 'ChatMessage.Read.Chat'];
    assert.equal((await inputsOf(h))?.includes('presence'), false);
    h.granted = undefined;
    assert.equal((await inputsOf(h))?.includes('presence'), false, 'grants that cannot be read count as not granted');
    assert.ok(h.grantReads >= 2, 'the grant is read, not assumed');
  });

  it('does not declare presence for Meet-now or a call', async () => {
    for (const meetingType of ['MeetNow', 'GroupCall', 'OneToOneCall']) {
      const h = harness();
      await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start', meetingType, conversationId: GROUP_CHAT_ID }));

      const snapshot = await h.source.snapshot(GROUP_CHAT_ID, { asChannelKey: MIRA.key });
      assert.deepEqual(snapshot?.inputs, ['text', 'message_edits', 'card_answers'], meetingType);
      assert.equal(snapshot?.modality, 'online', 'a running call is an online session');
      assert.equal(h.grantReads, 0);
    }
  });

  it('withdraws presence when the meeting ends, and starts unproven in the next meeting', async () => {
    const h = harness();
    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));
    await h.deliver(meetingParticipantActivity({ bot: MIRA, from: ANNA, phase: 'join', people: [ANNA] }));
    assert.equal((await inputsOf(h))?.includes('presence'), true);

    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'end' }));
    assert.equal((await inputsOf(h))?.includes('presence'), false);

    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start', at: '2026-11-25T10:00:00.000Z' }));
    assert.equal((await inputsOf(h))?.includes('presence'), false, 'each occurrence has to prove it again');
  });

  it('never declares hand raise, audio or reactions', async () => {
    const h = harness();
    await h.deliver(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));
    await h.deliver(meetingParticipantActivity({ bot: MIRA, from: ANNA, phase: 'join', people: [ANNA] }));

    const snapshot = await h.source.snapshot(MEETING_CHAT_ID, { asChannelKey: MIRA.key });
    for (const never of ['hand_raise', 'transcribed_speech', 'reactions']) {
      assert.equal((snapshot?.inputs as readonly string[]).includes(never), false, never);
    }
    for (const never of ['voice', 'screen', 'pinned_surface']) {
      assert.equal((snapshot?.outputs as readonly string[]).includes(never), false, never);
    }
  });
});

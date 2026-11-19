import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  createTeamsConversationSendAdapter,
  createTeamsMeetingSource,
  createTeamsReadinessProvider,
  createTeamsTargetedSendAdapter,
  registerTeamsSendProviders,
  stampChannel,
  TEAMS_CHANNEL_TYPE,
  TeamsConversationMessages,
  TeamsConversationReferenceCache,
  TeamsMeetingSessions,
  TeamsQuestionCards,
} from '@omadia/channel-teams';
import type { ConversationMessageEvent, SessionEvent } from '../src/kernel-types.js';

import {
  ANNA,
  GROUP_CHAT_ID,
  meetingLifecycleActivity,
  messageActivity,
  MIRA,
  SimulatedConnector,
} from './support/teamsSimulation.js';

// byte5ai/omadia#2043 — the kernel only uses a send option a provider
// DECLARES, and only accepts events that carry the emitting plugin's own
// catalog id and channel type. These tests look at the objects that are
// registered, not at copies.

const CATALOG_ID = '@omadia/channel-teams';

function register(withCards = false): {
  registered: { kind: string; channelId: string; provider: Record<string, unknown> }[];
  connector: SimulatedConnector;
  refs: TeamsConversationReferenceCache;
} {
  const registered: { kind: string; channelId: string; provider: Record<string, unknown> }[] = [];
  const connector = new SimulatedConnector();
  const refs = new TeamsConversationReferenceCache();
  registerTeamsSendProviders(
    {
      registerTargetedSendProvider: (channelId, provider) => {
        registered.push({ kind: 'targeted', channelId, provider: provider as unknown as Record<string, unknown> });
      },
      registerConversationSendProvider: (channelId, provider) => {
        registered.push({ kind: 'conversation', channelId, provider: provider as unknown as Record<string, unknown> });
      },
    },
    CATALOG_ID,
    {
      refs,
      sendProactive: connector.sendProactive,
      roster: { list: async () => [] },
      openDirectConversation: connector.openDirectConversation,
      botLabel: (botAppId) => (botAppId === MIRA.appId ? MIRA.name : undefined),
      ...(withCards ? { cards: new TeamsQuestionCards() } : {}),
    },
  );
  return { registered, connector, refs };
}

describe('declared send capabilities on the registered providers', () => {
  it('the conversation send provider declares own identity, mentions and polls', () => {
    const { registered } = register();
    const conversation = registered.find((entry) => entry.kind === 'conversation');

    assert.ok(conversation);
    assert.equal(conversation.channelId, CATALOG_ID);
    assert.equal(conversation.provider['channelType'], 'teams');
    assert.deepEqual(conversation.provider['supports'], { asChannelKey: true, mentions: true, poll: true, aiLabel: true });
  });

  it('wired as the plugin wires it, the conversation provider declares question cards and can update them', () => {
    // plugin.ts passes the card registry; then the registered object says
    // which question kinds it renders (#2089) and offers `updateCard`.
    const { registered } = register(true);
    const conversation = registered.find((entry) => entry.kind === 'conversation');

    assert.ok(conversation);
    assert.deepEqual(conversation.provider['supports'], {
      asChannelKey: true,
      mentions: true,
      poll: true,
      aiLabel: true,
      card: {
        kinds: ['single_choice', 'multiple_choice', 'ranking', 'consent', 'assignment', 'text'],
        maxQuestions: 5,
        update: true,
      },
    });
    assert.equal(typeof conversation.provider['updateCard'], 'function');
  });

  it('the targeted send provider declares own identity', () => {
    const { registered } = register();
    const targeted = registered.find((entry) => entry.kind === 'targeted');

    assert.ok(targeted);
    assert.equal(targeted.channelId, CATALOG_ID);
    assert.equal(targeted.provider['channelType'], 'teams');
    assert.deepEqual(targeted.provider['supports'], { asChannelKey: true, aiLabel: true });
  });

  it('what is declared is what the registered objects do: a post and a DM as the named bot', async () => {
    const { registered, connector, refs } = register();
    refs.capture({ activity: messageActivity({ bot: MIRA, from: ANNA, text: 'hi' }) } as unknown as TurnContext);
    const conversation = registered.find((entry) => entry.kind === 'conversation')?.provider as unknown as ReturnType<typeof createTeamsConversationSendAdapter>;
    const targeted = registered.find((entry) => entry.kind === 'targeted')?.provider as unknown as ReturnType<typeof createTeamsTargetedSendAdapter>;

    const post = await conversation.sendToConversation(
      GROUP_CHAT_ID,
      { text: 'Abstimmung', poll: { pollId: 'p', question: 'q', options: [{ id: 'a', label: 'A' }, { id: 'b', label: 'B' }] } },
      { asChannelKey: MIRA.key },
    );
    const dm = await targeted.sendToUser(
      { principalId: 'user:anna', userRef: { kind: 'teams-aad', id: ANNA.aadObjectId } },
      { text: 'x' },
      { asChannelKey: MIRA.key },
    );

    assert.equal(post.outcome, 'delivered');
    assert.equal(dm.outcome, 'delivered');
    assert.ok(connector.sent.every((sent) => sent.asBot === MIRA.key));
  });

  it('does not declare what it is not wired for', () => {
    const connector = new SimulatedConnector();
    const refs = new TeamsConversationReferenceCache();
    assert.deepEqual(createTeamsTargetedSendAdapter({ sendProactive: connector.sendProactive }).supports, {
      asChannelKey: false,
      aiLabel: true,
    });
    assert.deepEqual(createTeamsConversationSendAdapter({ refs, sendProactive: connector.sendProactive }).supports, {
      asChannelKey: true,
      mentions: false,
      poll: true,
      aiLabel: true,
    });
  });
});

describe('everything handed to the kernel carries this plugin\'s catalog id and channel type', () => {
  it('message events', () => {
    const emitted: ConversationMessageEvent[] = [];
    const messages = new TeamsConversationMessages({
      isAllMessagesSubscribed: () => true,
      emit: (event) => {
        emitted.push(stampChannel(event, CATALOG_ID));
      },
    });

    messages.emitCreated(messageActivity({ bot: MIRA, from: ANNA, text: 'A' }));

    assert.equal(emitted.length, 1);
    assert.equal(emitted[0]?.channelId, CATALOG_ID);
    assert.equal(emitted[0]?.channelType, TEAMS_CHANNEL_TYPE);
  });

  it('session events', () => {
    const emitted: SessionEvent[] = [];
    const sessions = new TeamsMeetingSessions({
      emit: (event) => {
        emitted.push(stampChannel(event, CATALOG_ID));
      },
    });

    sessions.handle(meetingLifecycleActivity({ bot: MIRA, from: ANNA, phase: 'start' }));

    assert.equal(emitted[0]?.channelId, CATALOG_ID);
    assert.equal(emitted[0]?.channelType, TEAMS_CHANNEL_TYPE);
  });

  it('stampChannel overrides whatever the emitting module filled in', () => {
    const stamped = stampChannel({ channelId: 'placeholder', channelType: 'other', kind: 'x' }, CATALOG_ID);
    assert.deepEqual(stamped, { channelId: CATALOG_ID, channelType: 'teams', kind: 'x' });
  });

  it('readiness provider, meeting source and their snapshots are of channel type teams', async () => {
    const refs = new TeamsConversationReferenceCache();
    refs.capture({ activity: messageActivity({ bot: MIRA, from: ANNA, text: 'hi' }) } as unknown as TurnContext);
    const readiness = createTeamsReadinessProvider({ botLabel: () => undefined, graph: undefined, refs });
    const meetings = createTeamsMeetingSource({ refs });

    assert.equal(readiness.channelType, TEAMS_CHANNEL_TYPE);
    assert.equal(typeof readiness.repairConsent, 'function');
    assert.equal(meetings.channelType, TEAMS_CHANNEL_TYPE);
    assert.equal((await meetings.snapshot(GROUP_CHAT_ID, { asChannelKey: MIRA.key }))?.channelType, TEAMS_CHANNEL_TYPE);
  });
});

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import { createTeamsMeetingSource, isTeamsMeetingChat, TeamsConversationReferenceCache } from '@omadia/channel-teams';

import { ANNA, GROUP_CHAT_ID, JONAS, MEETING_CHAT_ID, messageActivity, MIRA } from './support/teamsSimulation.js';

// byte5ai/omadia#2043 §6.1, issue #44 — the capability snapshot.

function sourceWith(seen: { conversationId: string; personal?: boolean }[]): ReturnType<typeof createTeamsMeetingSource> {
  const refs = new TeamsConversationReferenceCache();
  for (const entry of seen) {
    refs.capture({
      activity: messageActivity({
        bot: MIRA,
        from: ANNA,
        text: 'hi',
        conversationId: entry.conversationId,
        ...(entry.personal ? { conversationType: 'personal' as const } : {}),
      }),
    } as unknown as TurnContext);
  }
  return createTeamsMeetingSource({ refs, now: () => new Date('2026-11-18T10:00:00.000Z') });
}

describe('#44 Teams meeting source', () => {
  it('declares the chat-and-cards baseline for a meeting chat as an online session', async () => {
    const source = sourceWith([{ conversationId: MEETING_CHAT_ID }]);

    const snapshot = await source.snapshot(MEETING_CHAT_ID, { asChannelKey: MIRA.key, sessionRef: 'meeting-42' });

    assert.deepEqual(snapshot, {
      channelType: 'teams',
      conversationId: MEETING_CHAT_ID,
      sessionRef: 'meeting-42',
      modality: 'online',
      inputs: ['text', 'message_edits', 'card_answers'],
      outputs: ['text', 'mentions', 'cards', 'polls', 'private_messages'],
      capturedAt: '2026-11-18T10:00:00.000Z',
    });
  });

  it('claims nothing the plugin does not implement', async () => {
    const snapshot = await sourceWith([{ conversationId: MEETING_CHAT_ID }]).snapshot(MEETING_CHAT_ID, {
      asChannelKey: MIRA.key,
    });
    assert.ok(snapshot);
    for (const absent of ['reactions', 'presence', 'hand_raise', 'transcribed_speech', 'shared_documents']) {
      assert.equal((snapshot.inputs as readonly string[]).includes(absent), false, absent);
    }
    for (const absent of ['voice', 'screen', 'pinned_surface']) {
      assert.equal((snapshot.outputs as readonly string[]).includes(absent), false, absent);
    }
    assert.notEqual(snapshot.modality, 'hybrid');
  });

  it('reports a group chat as chat and a 1:1 without group outputs', async () => {
    const source = sourceWith([{ conversationId: GROUP_CHAT_ID }, { conversationId: 'a:1personal', personal: true }]);

    const group = await source.snapshot(GROUP_CHAT_ID, { asChannelKey: MIRA.key });
    const direct = await source.snapshot('a:1personal', { asChannelKey: MIRA.key });

    assert.equal(group?.modality, 'chat');
    assert.equal('sessionRef' in (group ?? {}), false);
    assert.deepEqual(direct?.outputs, ['text', 'cards']);
  });

  it('knows nothing about a chat the asked bot is not in', async () => {
    const source = sourceWith([{ conversationId: MEETING_CHAT_ID }]);

    assert.equal(await source.snapshot(GROUP_CHAT_ID, { asChannelKey: MIRA.key }), undefined);
    assert.equal(await source.snapshot(MEETING_CHAT_ID, { asChannelKey: JONAS.key }), undefined);
    assert.equal(await source.snapshot(MEETING_CHAT_ID, { asChannelKey: '29:user' }), undefined);
  });

  it('recognises meeting chats by their id', () => {
    assert.equal(isTeamsMeetingChat(MEETING_CHAT_ID), true);
    assert.equal(isTeamsMeetingChat(GROUP_CHAT_ID), false);
  });
});

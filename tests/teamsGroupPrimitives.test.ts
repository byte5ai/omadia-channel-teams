import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  createTeamsRosterAdapter,
  createTeamsTargetedSendAdapter,
  TeamsConversationReferenceCache,
  toSdkConversationType,
} from '@omadia/channel-teams';
import type { TeamsRosterProvider } from '@omadia/channel-teams';
import type { TeamsProactiveSend } from '../src/messagesRouter.js';
import type { ChatParticipant } from '../src/kernel-types.js';

// #330 B2 — Teams adapters for the kernel group-conversation seams.

function fakeContext(input: { conversationId: string; teamsType?: string }): TurnContext {
  return {
    activity: {
      type: 'message',
      channelId: 'msteams',
      serviceUrl: 'https://smba.example',
      conversation: { id: input.conversationId, ...(input.teamsType ? { conversationType: input.teamsType } : {}) },
      from: { id: '29:user', name: 'User' },
      recipient: { id: '28:bot', name: 'Omadia' },
    },
  } as unknown as TurnContext;
}

describe('toSdkConversationType', () => {
  it("maps personal→direct, groupChat/channel→group, unknown→undefined", () => {
    assert.equal(toSdkConversationType('personal'), 'direct');
    assert.equal(toSdkConversationType('groupChat'), 'group');
    assert.equal(toSdkConversationType('channel'), 'group');
    assert.equal(toSdkConversationType('weird'), undefined);
    assert.equal(toSdkConversationType(undefined), undefined);
  });
});

describe('TeamsConversationReferenceCache', () => {
  it('captures per conversation and serves the reference + teams type back', () => {
    const cache = new TeamsConversationReferenceCache();
    cache.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'groupChat' }));

    const hit = cache.get('conv-1');
    assert.ok(hit);
    assert.equal(hit.teamsType, 'groupChat');
    assert.equal(hit.ref.conversation?.id, 'conv-1');
    assert.equal(cache.get('conv-unknown'), undefined);
  });
});

const MEMBERS: ChatParticipant[] = [
  { channelUserId: '29:alice', aadObjectId: 'aad-alice', displayName: 'Alice', email: 'alice@co.com', userPrincipalName: 'alice@co.com' },
  { channelUserId: '28:botapp', aadObjectId: null, displayName: 'Omadia', email: null, userPrincipalName: null },
];

function rosterAdapterHarness(members: ChatParticipant[]): ReturnType<typeof createTeamsRosterAdapter> {
  const cache = new TeamsConversationReferenceCache();
  cache.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'groupChat' }));
  const sendProactive: TeamsProactiveSend = async (_ref, build) => {
    await build(fakeContext({ conversationId: 'conv-1' }));
  };
  const roster = { list: async () => members } as unknown as TeamsRosterProvider;
  return createTeamsRosterAdapter({ refs: cache, roster, sendProactive });
}

describe('createTeamsRosterAdapter', () => {
  it('resolves the roster through a proactive turn: AAD-keyed refs, bot detection, group type', async () => {
    const adapter = rosterAdapterHarness(MEMBERS);
    const roster = await adapter.getRoster('conv-1');

    assert.ok(roster);
    assert.equal(roster.conversationType, 'group');
    assert.equal(roster.partial, false);
    assert.deepEqual(
      roster.participants.map((p) => ({ id: p.userRef.id, isBot: p.isBot })),
      [{ id: 'aad-alice', isBot: false }, { id: '28:botapp', isBot: true }],
    );
    assert.equal(roster.self?.isBot, true);
    assert.equal(roster.self?.userRef.id, '28:bot');
  });

  it('reports an EMPTY member list as partial — the soft-fail is not an answer', async () => {
    const adapter = rosterAdapterHarness([]);
    const roster = await adapter.getRoster('conv-1');
    assert.ok(roster);
    assert.equal(roster.partial, true);
  });

  it('returns undefined for a conversation without a captured reference', async () => {
    const adapter = rosterAdapterHarness(MEMBERS);
    assert.equal(await adapter.getRoster('conv-never-seen'), undefined);
  });
});

describe('createTeamsTargetedSendAdapter', () => {
  it("returns 'no_binding' when the kernel has no conversation reference for the user", async () => {
    const adapter = createTeamsTargetedSendAdapter({ sendProactive: async () => undefined as never });
    const outcome = await adapter.sendToUser({ principalId: 'jane@co.com' }, { text: 'hi' });
    assert.deepEqual(outcome.outcome, 'unreachable');
    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'no_binding');
  });

  it('delivers through the proactive path when a reference exists', async () => {
    const sent: string[] = [];
    const sendProactive: TeamsProactiveSend = async (_ref, build) => {
      await build({
        sendActivity: async (activity: { text?: string }) => {
          sent.push(activity.text ?? '');
          return undefined;
        },
      } as unknown as TurnContext);
    };
    const adapter = createTeamsTargetedSendAdapter({ sendProactive });
    const outcome = await adapter.sendToUser({ principalId: 'jane@co.com', conversationRef: { c: 1 } }, { text: 'report' });

    assert.deepEqual(outcome, { outcome: 'delivered' });
    assert.deepEqual(sent, ['report']);
  });

  it("maps a proactive-send throw to 'channel_error' instead of throwing", async () => {
    const adapter = createTeamsTargetedSendAdapter({
      sendProactive: async () => {
        throw new Error('BF token expired');
      },
    });
    const outcome = await adapter.sendToUser({ principalId: 'jane@co.com', conversationRef: {} }, { text: 'x' });
    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'channel_error');
  });
});

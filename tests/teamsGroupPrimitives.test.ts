import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import { PgTeamsConversationRefStore } from '@omadia/channel-teams';
import {
  attributeGroupMessage,
  createTeamsConversationSendAdapter,
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

describe('createTeamsConversationSendAdapter (#330 C3b)', () => {
  it('delivers into a conversation with a cached reference via the proactive path', async () => {
    const cache = new TeamsConversationReferenceCache();
    cache.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'groupChat' }));
    const sent: string[] = [];
    const sendProactive: TeamsProactiveSend = async (_ref, build) => {
      await build({
        sendActivity: async (activity: { text?: string }) => {
          sent.push(activity.text ?? '');
          return undefined;
        },
      } as unknown as TurnContext);
    };
    const adapter = createTeamsConversationSendAdapter({ refs: cache, sendProactive });
    const out = await adapter.sendToConversation('conv-1', { text: 'nudge: wer fehlt noch?' });
    assert.deepEqual(out, { outcome: 'delivered' });
    assert.deepEqual(sent, ['nudge: wer fehlt noch?']);
  });

  it("returns 'no_binding' without a cached reference and 'channel_error' on a proactive throw", async () => {
    const cache = new TeamsConversationReferenceCache();
    const adapter = createTeamsConversationSendAdapter({
      refs: cache,
      sendProactive: async () => {
        throw new Error('BF token expired');
      },
    });
    const missing = await adapter.sendToConversation('conv-never-seen', { text: 'x' });
    assert.equal(missing.outcome === 'unreachable' ? missing.code : '', 'no_binding');

    cache.capture(fakeContext({ conversationId: 'conv-1' }));
    const threw = await adapter.sendToConversation('conv-1', { text: 'x' });
    assert.equal(threw.outcome === 'unreachable' ? threw.code : '', 'channel_error');
  });
});

describe('attributeGroupMessage (#330 field report — who is speaking?)', () => {
  it('prefixes group turns with the verified sender name', () => {
    assert.equal(
      attributeGroupMessage('Business Entscheider als Zielgruppe.', { isGroup: true, senderName: 'Marcel Wege' }),
      '[Marcel Wege]: Business Entscheider als Zielgruppe.',
    );
  });

  it('leaves 1:1 turns untouched — no prefix noise where the speaker is unambiguous', () => {
    assert.equal(attributeGroupMessage('hallo', { isGroup: false, senderName: 'Marcel Wege' }), 'hallo');
  });

  it('degrades to the raw text when the sender name is missing or blank', () => {
    assert.equal(attributeGroupMessage('hallo', { isGroup: true }), 'hallo');
    assert.equal(attributeGroupMessage('hallo', { isGroup: true, senderName: '   ' }), 'hallo');
  });

  it('trims the sender name (Bot Framework pads some display names)', () => {
    assert.equal(attributeGroupMessage('x', { isGroup: true, senderName: '  Christian Wendler ' }), '[Christian Wendler]: x');
  });
});

// #330 field report — references survive restarts: write-through on capture,
// store fallback on cache miss. Losing the store degrades to the old
// cache-only behaviour, never a throw on the message path.
describe('TeamsConversationReferenceCache persistence', () => {
  function fakeStore() {
    const rows = new Map<string, { ref: unknown; teamsType?: string }>();
    return {
      saves: [] as string[],
      rows,
      async save(conversationId: string, ref: unknown, teamsType?: string) {
        this.saves.push(conversationId);
        rows.set(conversationId, { ref, ...(teamsType ? { teamsType } : {}) });
      },
      async load(conversationId: string) {
        return rows.get(conversationId) as { ref: Partial<import('botbuilder').ConversationReference>; teamsType?: string } | undefined;
      },
    };
  }

  it('writes captures through once per distinct reference, and survives a "restart"', async () => {
    const store = fakeStore();
    const before = new TeamsConversationReferenceCache();
    before.attachPersistence(store);
    before.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'groupChat' }));
    before.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'groupChat' }));
    await Promise.resolve();
    assert.deepEqual(store.saves, ['conv-1'], 'identical re-captures must not re-write');

    // "Restart": a fresh cache over the same backing store.
    const after = new TeamsConversationReferenceCache();
    after.attachPersistence(store);
    assert.equal(after.get('conv-1'), undefined, 'sync path stays cache-only');
    const loaded = await after.getOrLoad('conv-1');
    assert.equal(loaded?.teamsType, 'groupChat');
    // Re-seeded: the next lookup is a plain cache hit.
    assert.ok(after.get('conv-1'));
  });

  it('conversationSend delivers after a restart via the persisted reference', async () => {
    const store = fakeStore();
    const before = new TeamsConversationReferenceCache();
    before.attachPersistence(store);
    before.capture(fakeContext({ conversationId: 'conv-1', teamsType: 'groupChat' }));
    await Promise.resolve();

    const after = new TeamsConversationReferenceCache();
    after.attachPersistence(store);
    const sent: string[] = [];
    const adapter = createTeamsConversationSendAdapter({
      refs: after,
      sendProactive: async (_ref, build) => {
        await build({
          sendActivity: async (activity: { text?: string }) => {
            sent.push(activity.text ?? '');
            return undefined;
          },
        } as unknown as TurnContext);
      },
    });
    const out = await adapter.sendToConversation('conv-1', { text: 'nudge nach neustart' });
    assert.deepEqual(out, { outcome: 'delivered' });
    assert.deepEqual(sent, ['nudge nach neustart']);
  });

  it('without persistence the old cache-only behaviour is unchanged', async () => {
    const cache = new TeamsConversationReferenceCache();
    assert.equal(await cache.getOrLoad('conv-never'), undefined);
  });
});

// Review M2 — a poisoned DB row must not point proactive turns at an
// attacker host; only Bot-Framework service hosts pass the load path.
describe('PgTeamsConversationRefStore — serviceUrl allowlist', () => {
  function storeWithRow(ref: Record<string, unknown>) {
    const pool = {
      query: async () => ({ rows: [{ ref, teams_type: 'groupChat' }] }),
    } as unknown as import('pg').Pool;
    return new PgTeamsConversationRefStore(pool);
  }

  it('returns a reference pointing at Bot Framework hosts', async () => {
    const ok = await storeWithRow({ serviceUrl: 'https://smba.trafficmanager.net/emea/', conversation: { id: 'c1' } }).load('c1');
    assert.equal(ok?.teamsType, 'groupChat');
    const okBf = await storeWithRow({ serviceUrl: 'https://europe.botframework.com/', conversation: { id: 'c1' } }).load('c1');
    assert.ok(okBf);
  });

  it('drops rows with an attacker or missing serviceUrl', async () => {
    assert.equal(await storeWithRow({ serviceUrl: 'https://evil.example.com/', conversation: { id: 'c1' } }).load('c1'), undefined);
    assert.equal(await storeWithRow({ conversation: { id: 'c1' } }).load('c1'), undefined);
    assert.equal(await storeWithRow({ serviceUrl: 'not a url' }).load('c1'), undefined);
  });
});

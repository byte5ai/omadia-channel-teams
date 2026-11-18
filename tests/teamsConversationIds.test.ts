import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  bareConversationId,
  isBotKeyPrefixedConversationId,
  isChannelConversationId,
  teamsTurnOrigin,
  threadRootMessageId,
} from '@omadia/channel-teams';

// byte5ai/omadia#2043 — ids as captured live on 2026-11-18.
const CHANNEL = '19:fdsWodnEDi7J3FI5tgPf1_4KxrURUShMFpNJrmcN8ns1@thread.tacv2';
const THREAD = `${CHANNEL};messageid=1795023950656`;
const MEETING_CHAT = '19:meeting_Mzc1OTg2YjctYmNkMC00Nzk4LTk3MjQtN2M1Y2M5NmU1MDlh@thread.v2';
const BOT_KEY = '28:6e7d5e84-5216-408f-b74b-e299bbd0d947';

describe('#2043 Teams conversation ids', () => {
  it('cuts the thread part off for Graph and leaves every other id alone', () => {
    assert.equal(bareConversationId(THREAD), CHANNEL);
    assert.equal(bareConversationId(CHANNEL), CHANNEL);
    assert.equal(bareConversationId(MEETING_CHAT), MEETING_CHAT);
    assert.equal(threadRootMessageId(THREAD), '1795023950656');
    assert.equal(threadRootMessageId(CHANNEL), undefined);
    assert.equal(threadRootMessageId(`${CHANNEL};messageid=`), undefined);
  });

  it('knows a channel by its id, with or without a thread', () => {
    assert.equal(isChannelConversationId(THREAD), true);
    assert.equal(isChannelConversationId(CHANNEL), true);
    assert.equal(isChannelConversationId(MEETING_CHAT), false);
    assert.equal(isChannelConversationId('19:5b1e2c3d_6c2f3d4e@unq.gbl.spaces'), false);
  });

  it('recognises an id that was cut out of a bot-qualified session scope', () => {
    // The live failure: `teams-28:<appId>-19:meeting_…` minus `teams-`.
    assert.equal(isBotKeyPrefixedConversationId(`${BOT_KEY}-${MEETING_CHAT}`), true);
    assert.equal(isBotKeyPrefixedConversationId(`${BOT_KEY.toUpperCase()}-${THREAD}`), true);
    assert.equal(isBotKeyPrefixedConversationId(MEETING_CHAT), false);
    assert.equal(isBotKeyPrefixedConversationId(THREAD), false);
    assert.equal(isBotKeyPrefixedConversationId('a:1dRsHCobZ1AxURzY05Dc_personal'), false);
    // A bot key alone is not a conversation id either, but not this bug.
    assert.equal(isBotKeyPrefixedConversationId(BOT_KEY), false);
  });
});

describe('#2043 teamsTurnOrigin — the channel-native conversation id', () => {
  const channelIdOf = (origin: unknown): unknown => (origin as { channelConversationId?: unknown }).channelConversationId;

  it('states `activity.conversation.id` unchanged, whatever the session scope looks like', () => {
    const cases: [string, string, string][] = [
      ['a channel thread', THREAD, 'channel'],
      ['a meeting chat', MEETING_CHAT, 'groupChat'],
      ['a 1:1', 'a:1dRsHCobZ1AxURzY05Dc_personal', 'personal'],
    ];
    for (const [name, id, conversationType] of cases) {
      const origin = teamsTurnOrigin(
        { conversation: { id, conversationType } },
        // Bot-qualified, as for every facilitator.
        `teams-${BOT_KEY}-${id}`,
        'aad-oid-1',
      );
      assert.equal(channelIdOf(origin), id, name);
    }
  });

  it('states nothing for an activity without a conversation id', () => {
    for (const conversation of [{}, { id: '' }, { id: '   ' }]) {
      const origin = teamsTurnOrigin({ conversation }, 'teams-x', 'aad-oid-1');
      assert.equal(Object.hasOwn(origin, 'channelConversationId'), false);
    }
  });
});

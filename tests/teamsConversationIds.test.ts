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

describe('#2043 teamsTurnOrigin — who spoke', () => {
  const ANNA_OID = '8f4c1d2e-3a5b-4c6d-9e7f-0a1b2c3d4e5f';
  const speakerOf = (origin: unknown): unknown => (origin as { speaker?: unknown }).speaker;
  const conversations: [string, { id: string; conversationType: string }][] = [
    ['a channel thread', { id: THREAD, conversationType: 'channel' }],
    ['a group chat', { id: '19:b8a1c0f2d3e44f5a9b6c7d8e9f0a1b2c@thread.v2', conversationType: 'groupChat' }],
    ['a meeting chat', { id: MEETING_CHAT, conversationType: 'groupChat' }],
    ['a 1:1', { id: 'a:1dRsHCobZ1AxURzY05Dc_personal', conversationType: 'personal' }],
  ];

  for (const [name, conversation] of conversations) {
    it(`names the sender with Entra object id and display name in ${name}`, () => {
      const origin = teamsTurnOrigin(
        { conversation, from: { aadObjectId: ANNA_OID, name: 'Anna Köhler' } },
        `teams-${BOT_KEY}-${conversation.id}`,
        ANNA_OID,
      );

      assert.deepEqual(speakerOf(origin), { kind: 'teams-aad', id: ANNA_OID, displayName: 'Anna Köhler' });
      // The same id the principal carries.
      assert.equal((origin.principal as { userId?: string } | undefined)?.userId, ANNA_OID);
    });
  }

  it('omits the display name when the activity has none', () => {
    for (const from of [{ aadObjectId: ANNA_OID }, { aadObjectId: ANNA_OID, name: '' }, { aadObjectId: ANNA_OID, name: '  ' }]) {
      const origin = teamsTurnOrigin({ conversation: { id: THREAD, conversationType: 'channel' }, from }, 'teams-x', ANNA_OID);
      assert.deepEqual(speakerOf(origin), { kind: 'teams-aad', id: ANNA_OID });
    }
  });

  it('omits the speaker without an Entra object id, also when a Teams id and a name are there', () => {
    // An anonymous meeting guest: `from.id` only. `userId` falls back to it,
    // the speaker does not.
    const guest: { aadObjectId?: string; name?: string } = { name: 'Gast' };
    const origin = teamsTurnOrigin({ conversation: { id: MEETING_CHAT, conversationType: 'groupChat' }, from: guest }, 'teams-x', '8:teamsvisitor:abc');
    assert.equal(Object.hasOwn(origin, 'speaker'), false);

    const none = teamsTurnOrigin({ conversation: { id: MEETING_CHAT, conversationType: 'groupChat' } }, 'teams-x', undefined);
    assert.equal(Object.hasOwn(none, 'speaker'), false);
  });
});

describe('channel-sdk 0.3.10 teamsTurnOrigin — replyToMessageId on private turns', () => {
  const PERSONAL = { id: 'a:1dRsHCobZ1AxURzY05Dc_personal', conversationType: 'personal' };
  const replyTo = (origin: unknown): unknown => (origin as { replyToMessageId?: unknown }).replyToMessageId;
  const quote = (itemId: string) => ({
    contentType: 'text/html',
    content: `<blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="${itemId}"><p>Erinnerung</p></blockquote><p>erledigt</p>`,
  });

  it('a quoted reply in a 1:1 names the quoted message id (the id a send reports)', () => {
    const origin = teamsTurnOrigin(
      { conversation: PERSONAL, id: '1700000000999', attachments: [quote('1700000000123')] },
      'teams-x',
      'aad-oid-1',
    );
    assert.equal(replyTo(origin), '1700000000123');
  });

  it('a 1:1 activity with replyToId names it', () => {
    const origin = teamsTurnOrigin({ conversation: PERSONAL, id: '2', replyToId: '1700000000123' }, 'teams-x', 'aad-oid-1');
    assert.equal(replyTo(origin), '1700000000123');
  });

  it('a quotedReply entity in a 1:1 names its message id', () => {
    const origin = teamsTurnOrigin(
      { conversation: PERSONAL, id: '2', entities: [{ type: 'quotedReply', quotedReply: { messageId: '1700000000123' } } as never] },
      'teams-x',
      'aad-oid-1',
    );
    assert.equal(replyTo(origin), '1700000000123');
  });

  it('a plain 1:1 message has none', () => {
    assert.equal(replyTo(teamsTurnOrigin({ conversation: PERSONAL, id: '2' }, 'teams-x', 'aad-oid-1')), undefined);
  });

  it('is not set on group turns', () => {
    const origin = teamsTurnOrigin(
      { conversation: { id: THREAD, conversationType: 'channel' }, id: '2', attachments: [quote('1700000000123')] },
      'teams-x',
      'aad-oid-1',
    );
    assert.equal(replyTo(origin), undefined);
  });
});

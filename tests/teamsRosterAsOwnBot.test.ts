import { afterEach, beforeEach, describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { TeamsInfo, type TurnContext } from 'botbuilder';

import {
  createTeamsRosterAdapter,
  TeamsConversationReferenceCache,
  TeamsRosterProvider,
} from '@omadia/channel-teams';

import {
  ANNA,
  BERT,
  CHANNEL_ID,
  channelThreadId,
  GROUP_CHAT_ID,
  JONAS,
  messageActivity,
  MIRA,
  OTHER_GROUP_CHAT_ID,
  SimulatedConnector,
  type SimBot,
  type SimUser,
} from './support/teamsSimulation.js';

/**
 * byte5ai/omadia#2043 — THE ROSTER IS ONE BOT'S VIEW.
 *
 * First live facilitation (2026-11-18): the handshake listed no participants.
 * The roster adapter looked the conversation up without a bot, i.e. as the
 * instance's DEFAULT bot, and that bot is not in the facilitator's chat. With
 * `asChannelKey` (channel-sdk 0.3.2) the roster is read as the asking agent's
 * own bot, or not at all.
 *
 * The Bot Connector's member list is simulated: it answers a bot that is in
 * the conversation and refuses (403) one that is not, as Teams does. The
 * real `TeamsRosterProvider` and the real adapter run on top of it.
 */

const original = TeamsInfo.getPagedMembers;

interface Room {
  bots: SimBot[];
  people: SimUser[];
}

/** Who is where, as Teams knows it. */
function teams(rooms: Record<string, Room>): { asked: { as: string; conversationId: string }[] } {
  const asked: { as: string; conversationId: string }[] = [];
  (TeamsInfo as unknown as { getPagedMembers: unknown }).getPagedMembers = async (context: TurnContext) => {
    const conversationId = context.activity.conversation?.id ?? '';
    const as = context.activity.recipient?.id ?? '';
    asked.push({ as, conversationId });
    const room = rooms[conversationId];
    if (!room || !room.bots.some((bot) => bot.key === as)) {
      throw Object.assign(new Error('The bot is not part of the conversation roster.'), { statusCode: 403 });
    }
    return {
      members: [
        ...room.people.map((person) => ({ id: person.id, aadObjectId: person.aadObjectId, name: person.name })),
        ...room.bots.map((bot) => ({ id: bot.key, name: bot.name })),
      ],
    };
  };
  return { asked };
}

function harness(defaultBot: SimBot = JONAS): {
  refs: TeamsConversationReferenceCache;
  adapter: ReturnType<typeof createTeamsRosterAdapter>;
  seen(bot: SimBot, conversationId: string, conversationType?: 'groupChat' | 'channel'): void;
} {
  const refs = new TeamsConversationReferenceCache();
  refs.setDefaultBotAppId(defaultBot.appId);
  const connector = new SimulatedConnector();
  const adapter = createTeamsRosterAdapter({ refs, roster: new TeamsRosterProvider(), sendProactive: connector.sendProactive });
  return {
    refs,
    adapter,
    seen: (bot, conversationId, conversationType = 'groupChat') =>
      refs.capture({ activity: messageActivity({ bot, from: ANNA, text: 'hi', conversationId, conversationType }) } as unknown as TurnContext),
  };
}

const names = (roster: { participants: readonly { userRef: { displayName?: string } }[] } | undefined): (string | undefined)[] =>
  (roster?.participants ?? []).map((participant) => participant.userRef.displayName);

describe('#2043 roster asked as the facilitator\'s own bot', () => {
  const lines: string[] = [];
  const originalError = console.error;

  beforeEach(() => {
    lines.length = 0;
    console.error = (...args: unknown[]) => {
      lines.push(args.map(String).join(' '));
    };
  });

  afterEach(() => {
    console.error = originalError;
    (TeamsInfo as unknown as { getPagedMembers: unknown }).getPagedMembers = original;
  });

  it('reads the members as that bot in a chat the default bot is not in', async () => {
    // Mira is the facilitator's bot; Jonas is the instance's default bot and
    // is not in this chat.
    const net = teams({ [GROUP_CHAT_ID]: { bots: [MIRA], people: [ANNA, BERT] } });
    const h = harness();
    h.seen(MIRA, GROUP_CHAT_ID);

    const roster = await h.adapter.getRoster(GROUP_CHAT_ID, { asChannelKey: MIRA.key });

    assert.deepEqual(names(roster), [ANNA.name, BERT.name, MIRA.name]);
    assert.equal(roster?.partial, false);
    assert.equal(roster?.conversationType, 'group');
    assert.equal(roster?.self?.userRef.id, MIRA.key, 'self is the asking bot');
    assert.equal(roster?.self?.isBot, true);
    assert.deepEqual(net.asked, [{ as: MIRA.key, conversationId: GROUP_CHAT_ID }]);
    assert.equal(roster?.status, 'complete');
    assert.ok(lines.some((line) => /\[teams-roster\] roster conv=.* as=28:[0-9a-f]{8}… status=complete count=3 readAt=/.test(line)), lines.join('\n'));
  });

  it('answers undefined when the asked bot holds no reference there, never another bot\'s view', async () => {
    const net = teams({ [GROUP_CHAT_ID]: { bots: [MIRA], people: [ANNA, BERT] } });
    const h = harness();
    h.seen(MIRA, GROUP_CHAT_ID);

    const roster = await h.adapter.getRoster(GROUP_CHAT_ID, { asChannelKey: JONAS.key });

    assert.equal(roster, undefined);
    assert.deepEqual(net.asked, [], 'nobody was asked on the other bot\'s behalf');
    assert.ok(lines.some((line) => line.includes('no reference: this bot has not received an activity from the conversation')));
  });

  it('refuses a key that is not a bot identity', async () => {
    teams({});
    const h = harness();
    h.seen(MIRA, GROUP_CHAT_ID);

    assert.equal(await h.adapter.getRoster(GROUP_CHAT_ID, { asChannelKey: 'mira' }), undefined);
    assert.ok(lines.some((line) => line.includes("refused: 'mira' is not a Teams bot identity key")));
  });

  it('one bot\'s refused read is not served to another bot of the same chat', async () => {
    // The default bot holds an old reference for the chat but was removed.
    const net = teams({ [GROUP_CHAT_ID]: { bots: [MIRA], people: [ANNA, BERT] } });
    const h = harness();
    h.seen(JONAS, GROUP_CHAT_ID);
    h.seen(MIRA, GROUP_CHAT_ID);

    const asDefault = await h.adapter.getRoster(GROUP_CHAT_ID);
    const asMira = await h.adapter.getRoster(GROUP_CHAT_ID, { asChannelKey: MIRA.key });

    assert.deepEqual(names(asDefault), [], 'the default bot sees nobody');
    assert.equal(asDefault?.status, 'unknown', 'and says that this is not an answer');
    assert.equal(asDefault?.partial, true);
    assert.match(asDefault?.statusDetail ?? '', /^HTTP 403: The bot is not part of the conversation roster\.$/);
    assert.deepEqual(names(asMira), [ANNA.name, BERT.name, MIRA.name], 'the cached empty read of the other bot is not reused');
    assert.deepEqual(net.asked.map((call) => call.as), [JONAS.key, MIRA.key]);
    assert.ok(lines.some((line) => /as=default-bot status=unknown count=0 http=403 detail=The bot is not part of the conversation roster\./.test(line)), lines.join('\n'));
  });

  it('without asChannelKey behaves as before: the default bot\'s reference and view', async () => {
    const net = teams({ [OTHER_GROUP_CHAT_ID]: { bots: [JONAS], people: [ANNA] } });
    const h = harness();
    h.seen(JONAS, OTHER_GROUP_CHAT_ID);

    const roster = await h.adapter.getRoster(OTHER_GROUP_CHAT_ID);

    assert.deepEqual(names(roster), [ANNA.name, JONAS.name]);
    assert.equal(roster?.self?.userRef.id, JONAS.key);
    assert.deepEqual(net.asked, [{ as: JONAS.key, conversationId: OTHER_GROUP_CHAT_ID }]);
    assert.equal(await h.adapter.getRoster(GROUP_CHAT_ID), undefined, 'no reference: unknown, as before');
  });

  it('asks for the members of the CHANNEL when the conversation is one of its threads', async () => {
    const thread = channelThreadId('1795023950656');
    const net = teams({ [CHANNEL_ID]: { bots: [MIRA], people: [ANNA, BERT] } });
    const h = harness();
    h.seen(MIRA, thread, 'channel');

    const roster = await h.adapter.getRoster(thread, { asChannelKey: MIRA.key });

    assert.deepEqual(names(roster), [ANNA.name, BERT.name, MIRA.name]);
    assert.deepEqual(net.asked, [{ as: MIRA.key, conversationId: CHANNEL_ID }], 'the bare channel id, not the thread id');
  });

  it('refuses an id cut out of a session scope and says so in the log', async () => {
    const net = teams({});
    const h = harness();
    h.seen(MIRA, GROUP_CHAT_ID);

    assert.equal(await h.adapter.getRoster(`${MIRA.key}-${GROUP_CHAT_ID}`, { asChannelKey: MIRA.key }), undefined);
    assert.deepEqual(net.asked, []);
    assert.ok(lines.some((line) => line.includes('refused: not a Teams conversation id (bot-key-prefixed)')));
  });
});

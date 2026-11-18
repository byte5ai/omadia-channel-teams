import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  createTeamsConversationSendAdapter,
  createTeamsTargetedSendAdapter,
  parsePollAnswerValue,
  TeamsConversationReferenceCache,
} from '@omadia/channel-teams';
import type { TeamsRosterProvider } from '@omadia/channel-teams';
import type { ChatParticipant } from '../src/kernel-types.js';

import {
  ANNA,
  BERT,
  botNotInstalledError,
  connectorError,
  directConversationId,
  GROUP_CHAT_ID,
  JONAS,
  messageActivity,
  messageWritesBlockedError,
  MIRA,
  referenceFor,
  SERVICE_URL,
  SimulatedConnector,
  type SimBot,
  type SimUser,
} from './support/teamsSimulation.js';

// byte5ai/omadia#2043 §6.2, issue #43 — proactive posts and DMs as the
// facilitator's own bot. The Bot Connector is simulated.

const ROSTER: ChatParticipant[] = [ANNA, BERT].map((user) => ({
  channelUserId: user.id,
  aadObjectId: user.aadObjectId,
  displayName: user.name,
  email: null,
  userPrincipalName: null,
}));

function capture(cache: TeamsConversationReferenceCache, bot: SimBot, from: SimUser, conversationId: string, personal = false): void {
  cache.capture({
    activity: messageActivity({ bot, from, text: 'hi', conversationId, ...(personal ? { conversationType: 'personal' as const } : {}) }),
  } as unknown as TurnContext);
}

function groupHarness(): {
  connector: SimulatedConnector;
  send: ReturnType<typeof createTeamsConversationSendAdapter>['sendToConversation'];
} {
  const connector = new SimulatedConnector();
  const refs = new TeamsConversationReferenceCache();
  capture(refs, MIRA, ANNA, GROUP_CHAT_ID);
  capture(refs, JONAS, ANNA, GROUP_CHAT_ID);
  const roster = { list: async () => ROSTER } as unknown as TeamsRosterProvider;
  const adapter = createTeamsConversationSendAdapter({ refs, sendProactive: connector.sendProactive, roster });
  return { connector, send: adapter.sendToConversation.bind(adapter) };
}

describe('#43 sendToConversation as the own bot', () => {
  it('turns a mention into a real Teams mention and returns the activity id', async () => {
    const { connector, send } = groupHarness();

    const outcome = await send(
      GROUP_CHAT_ID,
      {
        text: '@Anna Köhler deine Stimme fehlt noch.',
        mentions: [{ user: { kind: 'teams-aad', id: ANNA.aadObjectId }, text: '@Anna Köhler' }],
      },
      { asChannelKey: MIRA.key },
    );

    assert.equal(connector.sent.length, 1);
    const sent = connector.sent[0];
    assert.ok(sent);
    assert.deepEqual(outcome, { outcome: 'delivered', messageId: sent.id });
    assert.equal(sent.asBot, MIRA.key);
    assert.equal(sent.activity['text'], '<at>Anna Köhler</at> deine Stimme fehlt noch.');
    assert.deepEqual(sent.activity['entities'], [
      { type: 'mention', text: '<at>Anna Köhler</at>', mentioned: { id: ANNA.id, name: 'Anna Köhler' } },
    ]);
  });

  it('leaves a mention of somebody outside the chat as plain text', async () => {
    const { connector, send } = groupHarness();

    await send(
      GROUP_CHAT_ID,
      {
        text: '@Carla Voss und @Bert Maier: bitte abstimmen.',
        mentions: [
          { user: { kind: 'teams-aad', id: 'not-in-this-chat' }, text: '@Carla Voss' },
          { user: { kind: 'teams-aad', id: BERT.aadObjectId }, text: '@Bert Maier' },
        ],
      },
      { asChannelKey: MIRA.key },
    );

    const sent = connector.sent[0];
    assert.equal(sent?.activity['text'], '@Carla Voss und <at>Bert Maier</at>: bitte abstimmen.');
    assert.equal((sent?.activity['entities'] as unknown[]).length, 1);
  });

  it('posts a poll as a card whose message id the answers will name', async () => {
    const { connector, send } = groupHarness();
    const message = {
      text: 'Abstimmung',
      poll: { pollId: 'poll-7', question: 'Welcher Termin?', options: [{ id: 'a', label: 'Mo' }, { id: 'b', label: 'Di' }] },
    };

    const outcome = await send(GROUP_CHAT_ID, message, { asChannelKey: MIRA.key });

    const sent = connector.sent[0];
    assert.ok(sent);
    assert.deepEqual(outcome, { outcome: 'delivered', messageId: sent.id });
    const attachments = sent.activity['attachments'] as { content: { actions: { data: unknown }[] } }[];
    assert.equal(attachments.length, 1);
    assert.deepEqual(parsePollAnswerValue(attachments[0]?.content.actions[0]?.data)?.pollId, 'poll-7');
  });

  it('refuses an unusable poll before posting anything', async () => {
    const { connector, send } = groupHarness();

    const outcome = await send(
      GROUP_CHAT_ID,
      { text: 'x', poll: { pollId: 'p', question: 'q', options: [{ id: 'a', label: 'A' }] } } as never,
      { asChannelKey: MIRA.key },
    );

    assert.equal(outcome.outcome, 'unreachable');
    assert.match(outcome.outcome === 'unreachable' ? outcome.message : '', /at least two options/);
    assert.deepEqual(connector.sent, []);
  });

  it('two bots in one chat: each post goes out as the bot that was asked', async () => {
    const { connector, send } = groupHarness();

    await send(GROUP_CHAT_ID, { text: 'von Mira' }, { asChannelKey: MIRA.key });
    await send(GROUP_CHAT_ID, { text: 'von Jonas' }, { asChannelKey: JONAS.key });

    assert.deepEqual(
      connector.sent.map((s) => [s.asBot, s.activity['text']]),
      [
        [MIRA.key, 'von Mira'],
        [JONAS.key, 'von Jonas'],
      ],
    );
  });

  it('does not fall back to another bot when the asked bot is not in the chat', async () => {
    const connector = new SimulatedConnector();
    const refs = new TeamsConversationReferenceCache();
    capture(refs, JONAS, ANNA, GROUP_CHAT_ID);
    const adapter = createTeamsConversationSendAdapter({ refs, sendProactive: connector.sendProactive });

    const outcome = await adapter.sendToConversation(GROUP_CHAT_ID, { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'unreachable');
    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'no_binding');
    assert.deepEqual(connector.sent, []);
  });
});

function directHarness(configured: SimBot[] = [MIRA, JONAS]): {
  connector: SimulatedConnector;
  refs: TeamsConversationReferenceCache;
  sendToUser: ReturnType<typeof createTeamsTargetedSendAdapter>['sendToUser'];
} {
  const connector = new SimulatedConnector();
  const refs = new TeamsConversationReferenceCache();
  const adapter = createTeamsTargetedSendAdapter({
    sendProactive: connector.sendProactive,
    direct: {
      refs,
      openDirectConversation: connector.openDirectConversation,
      botLabel: (botAppId) => configured.find((bot) => bot.appId === botAppId)?.name,
    },
  });
  return { connector, refs, sendToUser: adapter.sendToUser.bind(adapter) };
}

function target(user: SimUser, conversationRef?: unknown): Parameters<ReturnType<typeof createTeamsTargetedSendAdapter>['sendToUser']>[0] {
  return {
    principalId: `user:${user.aadObjectId}`,
    userRef: { kind: 'teams-aad', id: user.aadObjectId, displayName: user.name },
    ...(conversationRef !== undefined ? { conversationRef } : {}),
  };
}

describe('#43 sendToUser as the own bot', () => {
  it('first contact: opens the 1:1 as the asked bot and returns the message id', async () => {
    const { connector, sendToUser } = directHarness();

    const outcome = await sendToUser(target(ANNA), { text: 'Deine Stimme fehlt noch.' }, { asChannelKey: MIRA.key });

    assert.deepEqual(connector.opened, [{ asBot: MIRA.key, userId: ANNA.aadObjectId }]);
    assert.equal(connector.sent.length, 1);
    assert.deepEqual(outcome, { outcome: 'delivered', messageId: connector.sent[0]?.id });
    assert.equal(connector.sent[0]?.asBot, MIRA.key);
  });

  it('reuses the 1:1 the bot already has with the user', async () => {
    const { connector, refs, sendToUser } = directHarness();
    const dm = directConversationId(`${MIRA.appId}|${ANNA.aadObjectId}`);
    capture(refs, MIRA, ANNA, dm, true);

    const outcome = await sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'delivered');
    assert.deepEqual(connector.opened, [], 'no new conversation is opened');
    assert.equal(connector.sent[0]?.conversationId, dm);
  });

  it('uses the regional endpoint the bot was last reached on for a first contact', async () => {
    const { connector, refs, sendToUser } = directHarness();
    capture(refs, MIRA, BERT, GROUP_CHAT_ID);

    await sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(connector.opened[0]?.serviceUrl, SERVICE_URL);
  });

  it('never uses another bot\'s conversation with the same user', async () => {
    const { connector, refs, sendToUser } = directHarness();
    const jonasDm = directConversationId(`${JONAS.appId}|${ANNA.aadObjectId}`);
    capture(refs, JONAS, ANNA, jonasDm, true);
    // The kernel's binding reference also belongs to Jonas.
    const kernelRef = { ...referenceFor(JONAS, jonasDm), conversation: { id: jonasDm, conversationType: 'personal' } };

    const outcome = await sendToUser(target(ANNA, kernelRef), { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'delivered');
    assert.deepEqual(connector.sent.map((s) => s.asBot), [MIRA.key]);
    assert.notEqual(connector.sent[0]?.conversationId, jonasDm);
  });

  it('returns app_not_installed when the user never installed the bot', async () => {
    const { connector, sendToUser } = directHarness();
    connector.failDirect(MIRA, ANNA, botNotInstalledError());

    const outcome = await sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'unreachable');
    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'app_not_installed');
    assert.match(outcome.outcome === 'unreachable' ? outcome.message : '', /'Mira' is not installed for this user/);
    assert.deepEqual(connector.sent, [], 'no other bot steps in');
  });

  it('returns blocked when the user blocked the bot, on first contact and on a known 1:1', async () => {
    const first = directHarness();
    first.connector.failDirect(MIRA, ANNA, messageWritesBlockedError());
    const firstOutcome = await first.sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });
    assert.equal(firstOutcome.outcome === 'unreachable' ? firstOutcome.code : '', 'blocked');

    const known = directHarness();
    capture(known.refs, MIRA, ANNA, directConversationId(`${MIRA.appId}|${ANNA.aadObjectId}`), true);
    known.connector.failDirect(MIRA, ANNA, messageWritesBlockedError());
    const knownOutcome = await known.sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });
    assert.equal(knownOutcome.outcome === 'unreachable' ? knownOutcome.code : '', 'blocked');
    assert.deepEqual(known.connector.sent, []);
  });

  it('maps an unknown user to user_not_found and anything else to channel_error', async () => {
    const h = directHarness();
    h.connector.failDirect(MIRA, ANNA, connectorError(404, 'NotFound', 'User not found'));
    h.connector.failDirect(MIRA, BERT, connectorError(502, 'BadGateway', 'upstream'));

    const notFound = await h.sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });
    const other = await h.sendToUser(target(BERT), { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(notFound.outcome === 'unreachable' ? notFound.code : '', 'user_not_found');
    assert.equal(other.outcome === 'unreachable' ? other.code : '', 'channel_error');
  });

  it('refuses a sender that is not a configured bot', async () => {
    const { connector, sendToUser } = directHarness([JONAS]);

    const unknown = await sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });
    const notABot = await sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: '29:someone' });

    assert.equal(unknown.outcome === 'unreachable' ? unknown.code : '', 'not_permitted');
    assert.equal(notABot.outcome === 'unreachable' ? notABot.code : '', 'not_permitted');
    assert.deepEqual(connector.opened, []);
    assert.deepEqual(connector.sent, []);
  });

  it('finds the 1:1 again after a restart through a new first contact', async () => {
    const before = directHarness();
    await before.sendToUser(target(ANNA), { text: 'vor dem Neustart' }, { asChannelKey: MIRA.key });
    const conversationBefore = before.connector.sent[0]?.conversationId;

    // New process: empty caches, same Teams.
    const after = directHarness();
    const outcome = await after.sendToUser(target(ANNA), { text: 'nach dem Neustart' }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'delivered');
    assert.equal(after.connector.sent[0]?.conversationId, conversationBefore, 'Teams returns the existing 1:1');
  });

  it('without asChannelKey keeps the #330 behaviour', async () => {
    const { connector, sendToUser } = directHarness();

    const noRef = await sendToUser(target(ANNA), { text: 'x' });
    const dm = directConversationId(`${JONAS.appId}|${ANNA.aadObjectId}`);
    const withRef = await sendToUser(target(ANNA, { ...referenceFor(JONAS, dm) }), { text: 'x' });

    assert.equal(noRef.outcome === 'unreachable' ? noRef.code : '', 'no_binding');
    assert.equal(withRef.outcome, 'delivered');
    assert.deepEqual(connector.opened, []);
    assert.equal(connector.sent[0]?.asBot, JONAS.key);
  });
});

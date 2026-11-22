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
  conversationBlockedError,
  documentedNotInstalledError,
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

function entitiesOf(activity: Record<string, unknown> | undefined): { type?: string; additionalType?: string[] }[] {
  return (activity?.['entities'] as { type?: string; additionalType?: string[] }[] | undefined) ?? [];
}

function mentionEntities(activity: Record<string, unknown> | undefined): unknown[] {
  return entitiesOf(activity).filter((entity) => entity.type === 'mention');
}

/** Teams' "AI generated" label: the activity-level entity. */
function aiLabels(activity: Record<string, unknown> | undefined): unknown[] {
  return entitiesOf(activity).filter((entity) => entity.additionalType?.includes('AIGeneratedContent'));
}

describe('#2043 every message posted as an agent identity carries the AI label, once', () => {
  it('a text post, a poll card and a post with a mention each carry exactly one label', async () => {
    const { connector, send } = groupHarness();

    await send(GROUP_CHAT_ID, { text: 'Die Abstimmung läuft.' }, { asChannelKey: MIRA.key });
    await send(
      GROUP_CHAT_ID,
      { text: 'Bitte wählt.', poll: { pollId: 'p1', question: 'Welche Option?', options: [{ id: 'a', label: 'A' }, { id: 'b', label: 'B' }] } },
      { asChannelKey: MIRA.key },
    );
    await send(
      GROUP_CHAT_ID,
      { text: '@Anna Köhler bitte.', mentions: [{ user: { kind: 'teams-aad', id: ANNA.aadObjectId }, text: '@Anna Köhler' }] },
      { asChannelKey: MIRA.key },
    );

    assert.equal(connector.sent.length, 4, 'text, text + card, text');
    for (const sent of connector.sent) {
      assert.equal(aiLabels(sent.activity).length, 1, JSON.stringify(sent.activity).slice(0, 120));
    }
    // The label is the marking: no disclosure sentence is added to the text.
    assert.equal(connector.sent[0]?.activity['text'], 'Die Abstimmung läuft.');
  });

  it('a private reminder carries the label too', async () => {
    const connector = new SimulatedConnector();
    const refs = new TeamsConversationReferenceCache();
    const adapter = createTeamsTargetedSendAdapter({
      sendProactive: connector.sendProactive,
      direct: { refs, openDirectConversation: connector.openDirectConversation, botLabel: () => MIRA.name },
    });

    const outcome = await adapter.sendToUser(
      { principalId: `user:${BERT.aadObjectId}`, userRef: { kind: 'teams-aad', id: BERT.aadObjectId } },
      { text: 'Deine Stimme fehlt noch.' },
      { asChannelKey: MIRA.key },
    );

    assert.equal(outcome.outcome, 'delivered');
    assert.equal(aiLabels(connector.sent[0]?.activity).length, 1);
  });
});

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
    assert.deepEqual(mentionEntities(sent.activity), [
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
    assert.equal(mentionEntities(sent?.activity).length, 1);
  });

  it('sends text and poll card as separate activities and returns both ids', async () => {
    const { connector, send } = groupHarness();
    const message = {
      text: 'Abstimmung',
      poll: { pollId: 'poll-7', question: 'Welcher Termin?', options: [{ id: 'a', label: 'Mo' }, { id: 'b', label: 'Di' }] },
    };

    const outcome = await send(GROUP_CHAT_ID, message, { asChannelKey: MIRA.key });

    assert.equal(connector.sent.length, 2);
    const [text, card] = connector.sent;
    assert.ok(text && card);
    assert.deepEqual(outcome, { outcome: 'delivered', messageId: text.id, cardMessageId: card.id });
    assert.equal(text.activity['text'], 'Abstimmung');
    assert.equal('attachments' in text.activity, false, 'the text activity carries no card');
    assert.equal('text' in card.activity, false, 'the card activity carries no text');
    const attachments = card.activity['attachments'] as { content: { actions: { data: unknown }[] } }[];
    assert.equal(attachments.length, 1);
    assert.deepEqual(parsePollAnswerValue(attachments[0]?.content.actions[0]?.data)?.pollId, 'poll-7');
  });

  it('a poll without text is one activity and its id is the message id', async () => {
    const { connector, send } = groupHarness();

    const outcome = await send(
      GROUP_CHAT_ID,
      { text: '', poll: { pollId: 'p', question: 'q', options: [{ id: 'a', label: 'A' }, { id: 'b', label: 'B' }] } },
      { asChannelKey: MIRA.key },
    );

    assert.equal(connector.sent.length, 1);
    assert.deepEqual(outcome, { outcome: 'delivered', messageId: connector.sent[0]?.id, cardMessageId: connector.sent[0]?.id });
  });

  it('a multiple-answer poll (omadia#2155) is posted as text, never as the single-answer card', async () => {
    const { connector, send } = groupHarness();
    const poll = {
      pollId: 'approval-1',
      question: 'Welche Punkte gebt ihr frei?',
      hint: 'Mehrere Antworten möglich.',
      options: [{ id: 'A', label: 'Budget' }, { id: 'B', label: 'Zeitplan' }, { id: 'C', label: 'Team' }],
      multipleAnswers: true,
    };

    const outcome = await send(GROUP_CHAT_ID, { text: '', poll }, { asChannelKey: MIRA.key });

    assert.equal(connector.sent.length, 1);
    const activity = connector.sent[0]?.activity ?? {};
    assert.equal('attachments' in activity, false, 'no card, so no click can replace a written ballot');
    assert.equal(
      activity['text'],
      '**Welche Punkte gebt ihr frei?**\n\nMehrere Antworten möglich.\n\nA. Budget\nB. Zeitplan\nC. Team',
    );
    assert.deepEqual(outcome, { outcome: 'delivered', messageId: connector.sent[0]?.id, cardMessageId: connector.sent[0]?.id });
  });

  it('multipleAnswers false keeps the poll card', async () => {
    const { connector, send } = groupHarness();
    await send(
      GROUP_CHAT_ID,
      { text: '', poll: { pollId: 'p', question: 'q', options: [{ id: 'a', label: 'A' }, { id: 'b', label: 'B' }], multipleAnswers: false } },
      { asChannelKey: MIRA.key },
    );
    assert.equal((connector.sent[0]?.activity['attachments'] as unknown[] | undefined)?.length, 1);
  });

  it('resolves a mention by Entra object id only', async () => {
    const { connector, send } = groupHarness();

    await send(
      GROUP_CHAT_ID,
      { text: '@Anna Köhler bitte.', mentions: [{ user: { kind: 'teams-aad', id: ANNA.id }, text: '@Anna Köhler' }] },
      { asChannelKey: MIRA.key },
    );

    assert.equal(connector.sent[0]?.activity['text'], '@Anna Köhler bitte.', 'a 29: id is not an Entra object id');
    assert.deepEqual(mentionEntities(connector.sent[0]?.activity), []);
  });

  it('refuses an unusable poll before posting anything', async () => {
    const { connector, send } = groupHarness();

    const outcome = await send(
      GROUP_CHAT_ID,
      { text: 'x', poll: { pollId: 'p', question: 'q', options: [{ id: 'a', label: 'A' }] } },
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
  waits: number[];
  connector: SimulatedConnector;
  refs: TeamsConversationReferenceCache;
  sendToUser: ReturnType<typeof createTeamsTargetedSendAdapter>['sendToUser'];
} {
  const connector = new SimulatedConnector();
  const refs = new TeamsConversationReferenceCache();
  const waits: number[] = [];
  const adapter = createTeamsTargetedSendAdapter({
    sendProactive: connector.sendProactive,
    retry: {
      sleep: async (ms) => {
        waits.push(ms);
      },
    },
    direct: {
      refs,
      openDirectConversation: connector.openDirectConversation,
      botLabel: (botAppId) => configured.find((bot) => bot.appId === botAppId)?.name,
    },
  });
  return { waits, connector, refs, sendToUser: adapter.sendToUser.bind(adapter) };
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
    // The 1:1 conversation the message went into is stated (0.39.0): with the
    // message id it is what a later private reply (a quote) refers to.
    assert.deepEqual(outcome, {
      outcome: 'delivered',
      messageId: connector.sent[0]?.id,
      conversationId: directConversationId(`${MIRA.appId}|${ANNA.aadObjectId}`),
    });
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

  it('returns app_not_installed for the live answer (403 BadArgument) and for the documented one', async () => {
    for (const failure of [botNotInstalledError(), documentedNotInstalledError()]) {
      const { connector, sendToUser } = directHarness();
      connector.failDirect(MIRA, ANNA, failure);

      const outcome = await sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });

      assert.equal(outcome.outcome, 'unreachable');
      assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'app_not_installed');
      assert.match(outcome.outcome === 'unreachable' ? outcome.message : '', /'Mira' is not installed for this user/);
      assert.match(outcome.outcome === 'unreachable' ? outcome.message : '', /sharing a group chat with it is not enough/);
      assert.deepEqual(connector.sent, [], 'no other bot steps in');
    }
  });

  it('does not read every 403 BadArgument as not installed', async () => {
    const { connector, sendToUser } = directHarness();
    connector.failDirect(MIRA, ANNA, connectorError(403, 'BadArgument', 'Invalid tenant id'));

    const outcome = await sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'channel_error');
  });

  it('blocked shows on the send, not on creating the conversation: first contact', async () => {
    const { connector, sendToUser } = directHarness();
    connector.block(MIRA, ANNA);

    const outcome = await sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(connector.opened.length, 1, 'the conversation was created (201 while blocked)');
    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'blocked');
    assert.deepEqual(connector.sent, []);
  });

  it('blocked on a known 1:1, and delivered again after the user unblocks', async () => {
    const { connector, refs, sendToUser } = directHarness();
    capture(refs, MIRA, ANNA, directConversationId(`${MIRA.appId}|${ANNA.aadObjectId}`), true);
    connector.block(MIRA, ANNA);

    const blocked = await sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });
    connector.unblock(MIRA, ANNA);
    const after = await sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(blocked.outcome === 'unreachable' ? blocked.code : '', 'blocked');
    assert.equal(after.outcome, 'delivered');
  });

  it('still maps the documented MessageWritesBlocked shape to blocked', async () => {
    const { connector, sendToUser } = directHarness();
    connector.block(MIRA, ANNA, messageWritesBlockedError());

    const outcome = await sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'blocked');
  });

  it('never retries a 403: one attempt, no waiting', async () => {
    const notInstalled = directHarness();
    notInstalled.connector.failDirect(MIRA, ANNA, botNotInstalledError());
    await notInstalled.sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });
    assert.equal(notInstalled.connector.opened.length, 1);
    assert.deepEqual(notInstalled.waits, []);

    const blocked = directHarness();
    blocked.connector.block(MIRA, ANNA, conversationBlockedError());
    await blocked.sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });
    assert.equal(blocked.connector.opened.length, 1);
    assert.deepEqual(blocked.waits, []);
  });

  it('retries 429, 412, 502 and 504 with exponential backoff and then delivers once', async () => {
    for (const status of [429, 412, 502, 504]) {
      const h = directHarness();
      h.connector.failDirect(MIRA, ANNA, connectorError(status, 'Transient', 'try later'), 2);

      const outcome = await h.sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });

      assert.equal(outcome.outcome, 'delivered', String(status));
      assert.deepEqual(h.waits, [500, 1000]);
      assert.equal(h.connector.sent.length, 1, 'the message is posted exactly once');
    }
  });

  it('gives up after four attempts and reports the error', async () => {
    const h = directHarness();
    h.connector.failDirect(MIRA, ANNA, connectorError(429, 'TooManyRequests', 'throttled'));

    const outcome = await h.sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'channel_error');
    assert.equal(h.connector.opened.length, 4);
    assert.deepEqual(h.waits, [500, 1000, 2000]);
  });

  it('does not match error codes neither seen live nor documented', async () => {
    const h = directHarness();
    h.connector.failDirect(MIRA, ANNA, connectorError(403, 'BotNotInConversationRoster', 'The bot is not part of the conversation roster.'));

    const outcome = await h.sendToUser(target(ANNA), { text: 'x' }, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'channel_error');
  });

  it('opens the conversation by Entra object id only: a 29: id is per bot and is not used', async () => {
    const h = directHarness();

    const outcome = await h.sendToUser(
      { principalId: 'user:x', userRef: { kind: 'teams-aad', id: ANNA.id } },
      { text: 'x' },
      { asChannelKey: MIRA.key },
    );

    assert.equal(outcome.outcome === 'unreachable' ? outcome.code : '', 'no_binding');
    assert.deepEqual(h.connector.opened, []);
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

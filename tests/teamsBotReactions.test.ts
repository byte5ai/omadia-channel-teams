import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  createTeamsConversationSendAdapter,
  TEAMS_REACTION_TYPES,
  TeamsConversationReferenceCache,
  type TeamsReactionFetch,
} from '@omadia/channel-teams';

import {
  ANNA,
  channelThreadId,
  GROUP_CHAT_ID,
  JONAS,
  messageActivity,
  MIRA,
  SERVICE_URL,
  SimulatedConnector,
  type SimBot,
} from './support/teamsSimulation.js';

/**
 * spec 2043 §5.14, omadia #2061 — the bot's own reaction on a message.
 *
 * Teams: `PUT {serviceUrl}/v3/conversations/{conversationId}/activities/
 * {activityId}/reactions/{reactionType}` with the bot's own token
 * (learn.microsoft.com/microsoftteams/platform/agents-in-teams/agent-reactions;
 * microsoft/teams.ts ReactionClient.add). The thumbs-up is `like`.
 *
 * NOT VERIFIED HERE: what a Teams client shows. The request is checked
 * against the documented route; the live check is on the owner checklist.
 */

interface RecordedRequest {
  url: string;
  method: string;
  headers: Record<string, string>;
}

type Answer = { status: number; body?: unknown } | Error;

/** The Bot Connector's reaction endpoint: records requests, answers in order
 *  (the last answer repeats). */
function fakeConnectorHttp(answers: Answer[] = [{ status: 200 }]) {
  const requests: RecordedRequest[] = [];
  const fetch: TeamsReactionFetch = async (url, init) => {
    requests.push({ url, method: init.method, headers: { ...init.headers } });
    const answer = answers[Math.min(requests.length - 1, answers.length - 1)] as Answer;
    if (answer instanceof Error) throw answer;
    const body = answer.body === undefined ? '' : JSON.stringify(answer.body);
    return { ok: answer.status >= 200 && answer.status < 300, status: answer.status, text: async () => body };
  };
  return { requests, fetch };
}

function harness(answers?: Answer[]) {
  const refs = new TeamsConversationReferenceCache();
  const connector = new SimulatedConnector();
  const http = fakeConnectorHttp(answers);
  const sleeps: number[] = [];
  const provider = createTeamsConversationSendAdapter({
    refs,
    sendProactive: connector.sendProactive,
    reactionFetch: http.fetch,
    retry: { sleep: async (ms) => void sleeps.push(ms) },
  });
  /** A person's message as the bot received it. Its activity id is the
   *  `messageId` of the plugin's message event (`inboundMessageId`). */
  function inbound(bot: SimBot, conversationId = GROUP_CHAT_ID, conversationType = 'groupChat'): string {
    const activity = messageActivity({ bot, from: ANNA, text: 'Mein Vorschlag', conversationId, conversationType });
    refs.capture({ activity } as unknown as TurnContext);
    return String(activity.id);
  }
  return { refs, connector, http, sleeps, provider, inbound };
}

function reactionUrl(conversationId: string, messageId: string): string {
  return `${SERVICE_URL}v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(messageId)}/reactions/like`;
}

describe('bot reactions: declaration', () => {
  it('declares reactions and maps acknowledge to the Teams thumbs-up', () => {
    const { provider } = harness();
    assert.equal(provider.supports.reactions, true);
    assert.equal(provider.supports.asChannelKey, true);
    assert.equal(typeof provider.react, 'function');
    assert.deepEqual(TEAMS_REACTION_TYPES, { acknowledge: 'like' });
  });
});

describe('bot reactions: the request', () => {
  it('PUTs a like on the message, signed with the token of the named bot, and posts nothing', async () => {
    const h = harness();
    const messageId = h.inbound(MIRA);

    const outcome = await h.provider.react(GROUP_CHAT_ID, messageId, 'acknowledge', { asChannelKey: MIRA.key });

    assert.deepEqual(outcome, { outcome: 'reacted' });
    assert.deepEqual(h.http.requests, [
      { url: reactionUrl(GROUP_CHAT_ID, messageId), method: 'PUT', headers: { Authorization: `Bearer token-of-${MIRA.appId}` } },
    ]);
    assert.deepEqual(h.connector.sent, []);
  });

  it('addresses a message in a channel thread under the thread conversation id', async () => {
    const h = harness();
    const thread = channelThreadId('1763456789012');
    const messageId = h.inbound(MIRA, thread, 'channel');

    const outcome = await h.provider.react(thread, messageId, 'acknowledge', { asChannelKey: MIRA.key });

    assert.deepEqual(outcome, { outcome: 'reacted' });
    assert.equal(h.http.requests[0]?.url, reactionUrl(thread, messageId));
    assert.match(h.http.requests[0]?.url ?? '', /%3Bmessageid%3D1763456789012\/activities\//);
  });
});

describe('bot reactions: identity', () => {
  it('reacts as exactly the bot asChannelKey names', async () => {
    const h = harness();
    const messageId = h.inbound(MIRA);
    h.inbound(JONAS);

    await h.provider.react(GROUP_CHAT_ID, messageId, 'acknowledge', { asChannelKey: JONAS.key });

    assert.equal(h.http.requests.length, 1);
    assert.equal(h.http.requests[0]?.headers['Authorization'], `Bearer token-of-${JONAS.appId}`);
  });

  it('refuses a key that is not a Teams bot identity, without a request', async () => {
    const h = harness();
    const messageId = h.inbound(MIRA);

    const outcome = await h.provider.react(GROUP_CHAT_ID, messageId, 'acknowledge', { asChannelKey: 'mira' });

    assert.equal(outcome.outcome, 'unreachable');
    assert.equal(outcome.outcome === 'unreachable' && outcome.code, 'not_permitted');
    assert.equal(h.http.requests.length, 0);
  });

  it('does not react through another bot when the named bot has no reference in the chat', async () => {
    const h = harness();
    const messageId = h.inbound(MIRA);

    const outcome = await h.provider.react(GROUP_CHAT_ID, messageId, 'acknowledge', { asChannelKey: JONAS.key });

    assert.equal(outcome.outcome, 'unreachable');
    assert.equal(outcome.outcome === 'unreachable' && outcome.code, 'no_binding');
    assert.equal(h.http.requests.length, 0);
  });

  it('refuses when the proactive turn runs under other credentials than the named bot', async () => {
    const h = harness();
    const messageId = h.inbound(MIRA);
    // A transport that would continue the conversation as Jonas.
    const provider = createTeamsConversationSendAdapter({
      refs: h.refs,
      sendProactive: (reference, build) => h.connector.sendProactive(reference, build, { botAppId: JONAS.appId }),
      reactionFetch: h.http.fetch,
    });

    const outcome = await provider.react(GROUP_CHAT_ID, messageId, 'acknowledge', { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome === 'unreachable' && outcome.code, 'not_permitted');
    assert.equal(h.http.requests.length, 0);
  });

  it('refuses a session-scope fragment instead of a conversation id', async () => {
    const h = harness();
    const messageId = h.inbound(MIRA);

    const outcome = await h.provider.react(`${MIRA.key}-${GROUP_CHAT_ID}`, messageId, 'acknowledge', { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome === 'unreachable' && outcome.code, 'no_binding');
    assert.equal(h.http.requests.length, 0);
  });

  it('refuses an empty message id and an unknown reaction, without a request', async () => {
    const h = harness();
    h.inbound(MIRA);

    const empty = await h.provider.react(GROUP_CHAT_ID, '', 'acknowledge', { asChannelKey: MIRA.key });
    const unknown = await h.provider.react(GROUP_CHAT_ID, '1', 'heart' as 'acknowledge', { asChannelKey: MIRA.key });

    assert.equal(empty.outcome === 'unreachable' && empty.code, 'channel_error');
    assert.equal(unknown.outcome === 'unreachable' && unknown.code, 'channel_error');
    assert.equal(h.http.requests.length, 0);
  });
});

describe('bot reactions: platform errors are returned, never thrown', () => {
  it('a deleted message: channel_error with the platform message, not retried', async () => {
    const h = harness([{ status: 404, body: { error: { code: 'NotFound', message: 'Activity not found' } } }]);
    const messageId = h.inbound(MIRA);

    const outcome = await h.provider.react(GROUP_CHAT_ID, messageId, 'acknowledge', { asChannelKey: MIRA.key });

    assert.deepEqual(outcome, {
      outcome: 'unreachable',
      code: 'channel_error',
      message: 'Bot Connector answered 404: NotFound: Activity not found',
    });
    assert.equal(h.http.requests.length, 1);
  });

  it('a bot removed from the chat (403) is not retried', async () => {
    const h = harness([{ status: 403, body: { error: { code: 'Forbidden', message: 'Bot is not part of the conversation roster.' } } }]);
    const messageId = h.inbound(MIRA);

    const outcome = await h.provider.react(GROUP_CHAT_ID, messageId, 'acknowledge', { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome === 'unreachable' && outcome.code, 'channel_error');
    assert.match(outcome.outcome === 'unreachable' ? outcome.message : '', /403: Forbidden: Bot is not part/);
    assert.equal(h.http.requests.length, 1);
  });

  it('the rate limit (429) is retried with backoff', async () => {
    const h = harness([{ status: 429 }, { status: 200 }]);
    const messageId = h.inbound(MIRA);

    const outcome = await h.provider.react(GROUP_CHAT_ID, messageId, 'acknowledge', { asChannelKey: MIRA.key });

    assert.deepEqual(outcome, { outcome: 'reacted' });
    assert.equal(h.http.requests.length, 2);
    assert.deepEqual(h.sleeps, [500]);
  });

  it('a lasting rate limit ends in channel_error', async () => {
    const h = harness([{ status: 429, body: 'Too Many Requests' }]);
    const messageId = h.inbound(MIRA);

    const outcome = await h.provider.react(GROUP_CHAT_ID, messageId, 'acknowledge', { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome === 'unreachable' && outcome.code, 'channel_error');
    assert.match(outcome.outcome === 'unreachable' ? outcome.message : '', /429/);
    assert.equal(h.http.requests.length, 4);
  });

  it('a network failure and a refused proactive turn are channel_error', async () => {
    const h = harness([new Error('socket hang up')]);
    const messageId = h.inbound(MIRA);
    const network = await h.provider.react(GROUP_CHAT_ID, messageId, 'acknowledge', { asChannelKey: MIRA.key });

    const refused = createTeamsConversationSendAdapter({
      refs: h.refs,
      sendProactive: async () => {
        throw new Error('[teams] proactive send refused — serviceUrl outside Bot Framework domains');
      },
      reactionFetch: h.http.fetch,
    });
    const turn = await refused.react(GROUP_CHAT_ID, messageId, 'acknowledge', { asChannelKey: MIRA.key });

    assert.deepEqual(network, { outcome: 'unreachable', code: 'channel_error', message: 'socket hang up' });
    assert.equal(turn.outcome === 'unreachable' && turn.code, 'channel_error');
    assert.match(turn.outcome === 'unreachable' ? turn.message : '', /serviceUrl outside/);
  });
});

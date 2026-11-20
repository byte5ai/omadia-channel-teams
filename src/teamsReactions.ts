// A bot's own reaction on a message (spec 2043 §5.14, omadia #2061).
//
// Teams lets an agent attach reactions to any message of a conversation it is
// in: `PUT {serviceUrl}/v3/conversations/{conversationId}/activities/
// {activityId}/reactions/{reactionType}` on the Bot Connector, with the bot's
// own token, no body.
//   https://learn.microsoft.com/en-us/microsoftteams/platform/agents-in-teams/agent-reactions
//   github.com/microsoft/teams.ts packages/api/src/clients/reaction/reaction.ts
// Adding is idempotent (a second add is no error). Adds and removes are
// rate-limited to two per second across all conversations of the bot; a 429 is
// retried with backoff by the caller (`retryTransient`). A deleted message or
// a bot removed from the chat fails and is not retried.
//
// botbuilder 4.23 has no method for this route. The request is therefore made
// here, signed by the credentials of the Bot Connector client that CloudAdapter
// put into the proactive turn of exactly the requested bot: no other bot's
// token can end up on it, and it goes only to the turn's own service URL.
// Graph `chatMessage: setReaction` is delegated-only and not used.

import type { TurnContext } from 'botbuilder';

import type { ConversationReaction } from './kernel-types.js';
import { isAllowedProactiveServiceUrl } from './teamsServiceUrl.js';

/** The Teams reaction id of each kernel reaction: `like` is the thumbs-up. */
export const TEAMS_REACTION_TYPES: Readonly<Record<ConversationReaction, string>> = Object.freeze({
  acknowledge: 'like',
});

/** The part of `fetch` this module uses; injected in tests. */
export type TeamsReactionFetch = (
  url: string,
  init: { method: 'PUT'; headers: Record<string, string> },
) => Promise<{ ok: boolean; status: number; text(): Promise<string> }>;

/** A failed reaction request. `statusCode` lets `retryTransient` retry a 429. */
export class TeamsReactionError extends Error {
  constructor(
    message: string,
    readonly code: 'channel_error' | 'not_permitted',
    readonly statusCode?: number,
  ) {
    super(message);
    this.name = 'TeamsReactionError';
  }
}

/** The Bot Connector URL of one reaction. Ids are encoded the way botbuilder's
 *  own connector client encodes them (a thread id carries `;messageid=`). */
export function teamsReactionUrl(serviceUrl: string, conversationId: string, activityId: string, reactionType: string): string {
  const base = serviceUrl.endsWith('/') ? serviceUrl : `${serviceUrl}/`;
  return `${base}v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(activityId)}/reactions/${encodeURIComponent(reactionType)}`;
}

/** What a request signer needs: botbuilder's `ServiceClientCredentials`. */
interface TurnCredentials {
  appId?: unknown;
  signRequest(webResource: { headers: { set(name: string, value: string): void } }): Promise<unknown>;
}

/** The credentials of the Bot Connector client of this turn, if any. */
function turnCredentials(turnContext: TurnContext): TurnCredentials | undefined {
  const key = (turnContext.adapter as { ConnectorClientKey?: symbol } | undefined)?.ConnectorClientKey;
  if (key === undefined) return undefined;
  const client = turnContext.turnState.get(key) as { credentials?: Partial<TurnCredentials> } | undefined;
  const credentials = client?.credentials;
  return typeof credentials?.signRequest === 'function' ? (credentials as TurnCredentials) : undefined;
}

/** The headers of a request signed by these credentials. */
async function signedHeaders(credentials: TurnCredentials): Promise<Record<string, string>> {
  const headers: Record<string, string> = {};
  await credentials.signRequest({ headers: { set: (name, value) => (headers[name] = value) } });
  return headers;
}

/** "403 Forbidden: <code>: <message>" from the service's error body. */
async function platformErrorMessage(response: { status: number; text(): Promise<string> }): Promise<string> {
  const body = await response.text().catch(() => '');
  try {
    const error = (JSON.parse(body) as { error?: { code?: unknown; message?: unknown } }).error;
    const parts = [error?.code, error?.message].filter((part): part is string => typeof part === 'string' && part.length > 0);
    if (parts.length > 0) return `Bot Connector answered ${response.status}: ${parts.join(': ')}`;
  } catch {
    // Not JSON: the raw body below.
  }
  return `Bot Connector answered ${response.status}${body ? `: ${body.slice(0, 300)}` : ''}`;
}

/**
 * Set one reaction as the bot of this proactive turn. Throws a
 * {@link TeamsReactionError}: `not_permitted` when the turn's credentials are
 * not those of `botAppId`, `channel_error` for everything the platform says.
 */
export async function putTeamsReaction(
  turnContext: TurnContext,
  request: { botAppId: string; conversationId: string; messageId: string; reactionType: string },
  fetchImpl: TeamsReactionFetch,
): Promise<void> {
  const serviceUrl = turnContext.activity.serviceUrl;
  if (!isAllowedProactiveServiceUrl(serviceUrl)) {
    throw new TeamsReactionError('the conversation reference names no Bot Framework service URL', 'channel_error');
  }
  const credentials = turnCredentials(turnContext);
  if (!credentials) {
    throw new TeamsReactionError('the proactive turn carries no Bot Connector client to sign the request with', 'channel_error');
  }
  const credentialAppId = typeof credentials.appId === 'string' ? credentials.appId.toLowerCase() : undefined;
  if (credentialAppId !== request.botAppId.toLowerCase()) {
    throw new TeamsReactionError(
      `the turn's credentials are not those of bot '${request.botAppId}' - refusing to react under another identity`,
      'not_permitted',
    );
  }
  const url = teamsReactionUrl(serviceUrl, request.conversationId, request.messageId, request.reactionType);
  const response = await fetchImpl(url, { method: 'PUT', headers: await signedHeaders(credentials) });
  if (!response.ok) {
    throw new TeamsReactionError(await platformErrorMessage(response), 'channel_error', response.status);
  }
}

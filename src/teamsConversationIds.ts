// Teams conversation ids, and which form goes where (#2043).
//
// THE CHANNEL-NATIVE ID is `activity.conversation.id`, unchanged. This plugin
// uses exactly that id for conversation references, the subscription lookup,
// message events and posts, and states it in the turn origin. The kernel and
// the engine treat it as opaque.
//
//   chat / meeting chat   19:…@thread.v2, 19:meeting_…@thread.v2
//   team channel          19:…@thread.tacv2;messageid=<root message id>
//
// A channel message carries the THREAD in its id: every root post is its own
// conversation, and a reply in that thread arrives under the same id. A
// facilitation in a channel is therefore scoped to the thread it was started
// in; a post for that conversation goes into the thread.
//
// GRAPH knows neither threads nor this plugin's session scopes. It is asked
// for a chat by the plain chat id and for a channel by its team. Nothing else
// is ever put into a Graph path.

import { TEAMS_BOT_KEY_PREFIX } from './teamsBotIdentity.js';

const THREAD_MARKER = ';messageid=';
const CHANNEL_SUFFIX = '@thread.tacv2';
/** `28:<app id>-`: how this plugin's SESSION SCOPE qualifies a conversation
 *  for a bot with its own identity (`teams-28:<appId>-<conversation id>`). A
 *  conversation id that starts like this was cut out of a session scope. */
const BOT_KEY_PREFIXED = new RegExp(
  `^${TEAMS_BOT_KEY_PREFIX}[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}-`,
  'i',
);

/** The id without the thread part: the channel itself, or the chat. */
export function bareConversationId(conversationId: string): string {
  const at = conversationId.indexOf(THREAD_MARKER);
  return at === -1 ? conversationId : conversationId.slice(0, at);
}

/** The root message id of a channel thread id, else `undefined`. */
export function threadRootMessageId(conversationId: string): string | undefined {
  const at = conversationId.indexOf(THREAD_MARKER);
  const root = at === -1 ? '' : conversationId.slice(at + THREAD_MARKER.length);
  return root.length > 0 ? root : undefined;
}

/** A team channel (or one of its threads), by the id alone. */
export function isChannelConversationId(conversationId: string): boolean {
  return bareConversationId(conversationId).endsWith(CHANNEL_SUFFIX);
}

/** `true` when the id is a session-scope fragment (`28:<appId>-<id>`), not a
 *  Teams conversation id. That is a bug of the caller. */
export function isBotKeyPrefixedConversationId(conversationId: string): boolean {
  return BOT_KEY_PREFIXED.test(conversationId);
}

/** The operator-readable refusal for such an id. */
export function botKeyPrefixedConversationMessage(conversationId: string): string {
  return `'${conversationId}' is not a Teams conversation id: it starts with a bot key ('${TEAMS_BOT_KEY_PREFIX}<appId>-'), which is how the Teams channel's session scope names a conversation for one bot. The caller has to pass the channel-native conversation id (the part after the bot key, as stated in the turn origin).`;
}

const MAX_REPORTED_IDS = 200;
const reported = new Set<string>();

/** Log a bot-key-prefixed id once per id (bounded), then stay quiet. */
export function reportBotKeyPrefixedConversationId(conversationId: string, where: string): void {
  if (reported.has(conversationId) || reported.size >= MAX_REPORTED_IDS) return;
  reported.add(conversationId);
  console.error(`[teams] ${where}: caller bug, ${botKeyPrefixedConversationMessage(conversationId)}`);
}

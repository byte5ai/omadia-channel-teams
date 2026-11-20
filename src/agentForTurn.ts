import type { AgentSpeaker, ChatAgent, SystemNotice } from './kernel-types.js';

/**
 * Which agent answers this turn.
 *
 * A Teams turn carries TWO routing keys and they answer different questions:
 *
 *   - `activity.recipient.id` (`28:<appId>`) — WHICH BOT was addressed.
 *   - `activity.conversation.id` — WHERE it was said.
 *
 * For a deployment with one bot the conversation is the interesting key: the
 * operator binds each channel to a different agent and the bot key is a
 * catch-all. Provisioning inverts that. Once several agents each have their
 * OWN Entra app, Azure bot and app package, they all sit in the same group
 * chat — so the conversation identifies a room full of bots and cannot say
 * which one was spoken to. Only the bot key can.
 *
 * Hence: a bot key that IS an agent's provisioned identity (the platform
 * reports `exclusive: true`) wins outright, and everything else keeps the
 * order it always had. Kept as a pure function so the rule is testable
 * without a Bot-Framework adapter, and so it has one place to be read.
 *
 * Contract of the management layer (byte5ai/omadia#2055, spec 2042 INV-7):
 * there is no fallback agent. A key yields `bound` (an agent answers) or
 * `reject`, and BOTH carry a `chatAgent`: the management layer, which handles
 * `/agent`, `/agents` and, on `reject`, answers with the `notice` as `system`.
 * The channel hands it the text and posts what comes back.
 */

/** The platform's answer for one key. Structural — see the channel SDK's
 *  `ChannelResolveResult`, of which this is the part we act on. */
export interface TurnAgentDecision {
  readonly decision: 'bound' | 'reject';
  /** Always set by a platform with a registry, on `reject` too. */
  readonly chatAgent?: ChatAgent;
  /**
   * The key IS this agent's provisioned identity, not a binding someone
   * chose. Also set on a `reject` of such a key: a provisioned bot never
   * answers for another agent, so its own refusal outranks every other key.
   */
  readonly exclusive?: boolean;
  /** With `reject`: the `system` notice the sender gets. */
  readonly notice?: SystemNotice;
  /** With `bound`: the agent that will answer, known before the answer. */
  readonly speaker?: AgentSpeaker;
}

/** The decision of a turn plus the ChatAgent that carries out `/agent` for it. */
export interface PickedTurnDecision extends TurnAgentDecision {
  /**
   * The management layer under the NARROWEST key of the turn. A selection
   * hangs on the `channelKey` of the resolve call plus the principal (S1), so
   * `/agent` must reach the ChatAgent of the conversation, not the one of a
   * bot-wide catch-all binding that happened to win the routing: otherwise the
   * choice would be stored for every chat of that bot. For an exclusive hit
   * (the bot IS the agent) it is the bot's own.
   */
  readonly selectionAgent?: ChatAgent;
}

export interface PickTurnDecisionInput {
  /** `activity.conversation.id` — the chat the message arrived in. */
  readonly conversationId?: string;
  /** `activity.recipient.id`, normalized: `28:<appId>` lowercased. */
  readonly botKey?: string;
  /** Ask the platform about one key. Called at most once per distinct key. */
  readonly resolve: (channelKey: string) => TurnAgentDecision;
}

/**
 * Pick the platform's decision for a turn, in precedence order:
 *
 *   1. the bot's own provisioned identity (`exclusive`), `bound` or `reject`:
 *      nothing overrides an agent's own bot, and a refused one stays refused;
 *   2. a binding on the conversation — legacy per-channel routing;
 *   3. a binding on the bot key — the catch-all for everything else;
 *   4. otherwise the refusal: the conversation's (it is where the sender is
 *      typing, so a `/agent` selection is stored for it), else the bot's.
 *
 * Returns `undefined` only when there is no key to ask about.
 */
export function pickTurnDecision(input: PickTurnDecisionInput): PickedTurnDecision | undefined {
  const { conversationId, botKey, resolve } = input;

  // The bot key is probed FIRST because only it can be exclusive — but a
  // non-exclusive hit here is still ranked below the conversation, so
  // probing order and precedence order deliberately differ.
  const bot = botKey ? resolve(botKey) : undefined;
  if (bot?.exclusive === true) return { ...bot, ...(bot.chatAgent ? { selectionAgent: bot.chatAgent } : {}) };

  const conversation =
    conversationId && conversationId !== botKey ? resolve(conversationId) : undefined;
  const selectionAgent = conversation?.chatAgent ?? bot?.chatAgent;
  const withSelection = (decision: TurnAgentDecision): PickedTurnDecision => ({
    ...decision,
    ...(selectionAgent ? { selectionAgent } : {}),
  });

  if (conversation?.decision === 'bound') return withSelection(conversation);
  if (bot?.decision === 'bound') return withSelection(bot);
  const refusal = conversation ?? bot;
  return refusal ? withSelection(refusal) : undefined;
}

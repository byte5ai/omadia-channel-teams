import type { ChatAgent } from './kernel-types.js';

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
 */

/** The platform's answer for one key. Structural — see the channel SDK's
 *  `ChannelResolveResult`, of which this is the part we act on. */
export interface TurnAgentDecision {
  readonly decision: 'bound' | 'fallback' | 'reject';
  readonly chatAgent?: ChatAgent;
  /**
   * The key IS this agent's provisioned identity, not a binding someone
   * chose. Optional: an older platform never sets it, and then this module
   * behaves exactly as it did before the field existed.
   */
  readonly exclusive?: boolean;
}

export interface PickChatAgentInput {
  /** `activity.conversation.id` — the chat the message arrived in. */
  readonly conversationId?: string;
  /** `activity.recipient.id`, normalized: `28:<appId>` lowercased. */
  readonly botKey?: string;
  /** Ask the platform about one key. Called at most once per distinct key. */
  readonly resolve: (channelKey: string) => TurnAgentDecision;
}

/** A decision only counts when it actually carries an agent to talk to. */
function agentOf(
  decision: TurnAgentDecision | undefined,
  kind: TurnAgentDecision['decision'],
): ChatAgent | undefined {
  if (!decision || decision.decision !== kind) return undefined;
  return decision.chatAgent;
}

/**
 * Pick the {@link ChatAgent} for a turn, in precedence order:
 *
 *   1. the bot's own provisioned identity (`exclusive`) — nothing overrides
 *      an agent's own bot;
 *   2. a binding on the conversation — legacy per-channel routing;
 *   3. a binding on the bot key — the catch-all for everything else;
 *   4. the platform fallback agent, whichever key offered it first.
 *
 * Returns `undefined` when no key yields an agent; the caller then keeps its
 * legacy default orchestrator.
 */
export function pickChatAgentForTurn(
  input: PickChatAgentInput,
): ChatAgent | undefined {
  const { conversationId, botKey, resolve } = input;

  // The bot key is probed FIRST because only it can be exclusive — but a
  // non-exclusive hit here is still ranked below the conversation, so
  // probing order and precedence order deliberately differ.
  const bot = botKey ? resolve(botKey) : undefined;
  const exclusive = bot?.exclusive === true ? agentOf(bot, 'bound') : undefined;
  if (exclusive) return exclusive;

  const conversation =
    conversationId && conversationId !== botKey
      ? resolve(conversationId)
      : undefined;

  return (
    agentOf(conversation, 'bound') ??
    agentOf(bot, 'bound') ??
    agentOf(conversation, 'fallback') ??
    agentOf(bot, 'fallback')
  );
}

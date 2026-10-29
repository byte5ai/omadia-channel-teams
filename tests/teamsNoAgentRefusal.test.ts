import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { NO_AGENT_FOR_CONVERSATION_TEXT, TeamsBot } from '../src/teamsBot.js';

/**
 * byte5ai/omadia#1598 — when the platform resolver is wired but serves no
 * agent for a chat (it rejected, or it threw), the turn is REFUSED. It used to
 * fall back to `defaultOrchestrator`, the kernel's legacy bundle, which has no
 * per-agent plugin grants and answers with every plugin's tools.
 */

type Resolve = (input: { channelKey: string }) => {
  decision: 'bound' | 'fallback' | 'reject';
  chatAgent?: unknown;
  exclusive?: boolean;
};

function botWith(resolve: Resolve | undefined): {
  resolveChatAgentForTurn(activity: unknown): { chatAgent?: unknown; refused?: true };
} {
  const bot = Object.create(TeamsBot.prototype) as Record<string, unknown>;
  Object.defineProperty(bot, 'resolveChatAgentForActivity', { value: resolve });
  return bot as never;
}

const ACTIVITY = {
  conversation: { id: '19:chat@thread.v2' },
  recipient: { id: '28:3d78d742-eefb-4fb2-bae5-3687f24c46fc' },
};

describe('#1598 — Teams refuses a turn no agent serves', () => {
  it('resolver rejects every key → refused, no agent', () => {
    const r = botWith(() => ({ decision: 'reject' })).resolveChatAgentForTurn(ACTIVITY);
    assert.equal(r.refused, true);
    assert.equal(r.chatAgent, undefined);
  });

  it('resolver throws → refused, never the default bundle', () => {
    const r = botWith(() => {
      throw new Error('registry down');
    }).resolveChatAgentForTurn(ACTIVITY);
    assert.equal(r.refused, true);
  });

  it('a bound agent is used and nothing is refused', () => {
    const agent = { slug: 'hr' };
    const r = botWith(() => ({ decision: 'bound', chatAgent: agent })).resolveChatAgentForTurn(
      ACTIVITY,
    );
    assert.equal(r.chatAgent, agent);
    assert.equal(r.refused, undefined);
  });

  it('no resolver wired (legacy single-agent install) → not refused', () => {
    const r = botWith(undefined).resolveChatAgentForTurn(ACTIVITY);
    assert.equal(r.refused, undefined);
  });

  it('the refusal text names no agent or plugin', () => {
    assert.match(NO_AGENT_FOR_CONVERSATION_TEXT, /kein Agent zugeordnet/);
  });
});

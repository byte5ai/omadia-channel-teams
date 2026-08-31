import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// Pure precedence logic — no botbuilder, no adapter, so the relative import
// is bundled by the test transpiler and needs no built `dist/`.
import { pickChatAgentForTurn, type TurnAgentDecision } from '../src/agentForTurn.js';
import type { ChatAgent } from '../src/kernel-types.js';

const HR = { id: 'hr' } as unknown as ChatAgent;
const MESSIAS = { id: 'messias' } as unknown as ChatAgent;
const DEFAULT = { id: 'fallback' } as unknown as ChatAgent;

const HR_BOT = '28:3d78d742-eefb-4fb2-bae5-3687f24c46fc';
const MESSIAS_BOT = '28:19ad2729-f7d3-4099-9d2a-7da1230c9533';
/** The one group chat every provisioned bot sits in. */
const GROUP = '19:5c8a1f60deadbeef@thread.skype';

/** Build a resolver over a fixed key→decision table, recording probe order. */
function resolverOver(
  table: Record<string, TurnAgentDecision>,
): { resolve: (key: string) => TurnAgentDecision; probes: string[] } {
  const probes: string[] = [];
  return {
    probes,
    resolve: (key) => {
      probes.push(key);
      return table[key] ?? { decision: 'reject' };
    },
  };
}

describe('pickChatAgentForTurn', () => {
  it('routes each provisioned bot in a shared chat to its own agent', () => {
    const table: Record<string, TurnAgentDecision> = {
      [HR_BOT]: { decision: 'bound', chatAgent: HR, exclusive: true },
      [MESSIAS_BOT]: { decision: 'bound', chatAgent: MESSIAS, exclusive: true },
    };

    const hr = pickChatAgentForTurn({
      conversationId: GROUP,
      botKey: HR_BOT,
      resolve: resolverOver(table).resolve,
    });
    const messias = pickChatAgentForTurn({
      conversationId: GROUP,
      botKey: MESSIAS_BOT,
      resolve: resolverOver(table).resolve,
    });

    assert.equal(hr, HR);
    assert.equal(messias, MESSIAS);
  });

  it('an exclusive bot hit beats a binding on the shared conversation', () => {
    // This is the regression that made the whole feature a no-op the other
    // way round: one binding on the group chat, three bots, one voice.
    const { resolve } = resolverOver({
      [GROUP]: { decision: 'bound', chatAgent: MESSIAS },
      [HR_BOT]: { decision: 'bound', chatAgent: HR, exclusive: true },
    });

    assert.equal(
      pickChatAgentForTurn({ conversationId: GROUP, botKey: HR_BOT, resolve }),
      HR,
    );
  });

  it('does not probe the conversation once the bot key answered exclusively', () => {
    const { resolve, probes } = resolverOver({
      [HR_BOT]: { decision: 'bound', chatAgent: HR, exclusive: true },
    });

    pickChatAgentForTurn({ conversationId: GROUP, botKey: HR_BOT, resolve });

    assert.deepEqual(probes, [HR_BOT]);
  });

  it('keeps legacy precedence: a conversation binding still beats a bot catch-all', () => {
    // One legacy bot, many channels: the operator binds each channel to a
    // different agent and the bot key is only the catch-all. No provisioned
    // identity is involved, so nothing here may change.
    const { resolve } = resolverOver({
      [GROUP]: { decision: 'bound', chatAgent: MESSIAS },
      [HR_BOT]: { decision: 'bound', chatAgent: HR },
    });

    assert.equal(
      pickChatAgentForTurn({ conversationId: GROUP, botKey: HR_BOT, resolve }),
      MESSIAS,
    );
  });

  it('falls through to the bot catch-all when the conversation is unbound', () => {
    const { resolve } = resolverOver({
      [GROUP]: { decision: 'fallback', chatAgent: DEFAULT },
      [HR_BOT]: { decision: 'bound', chatAgent: HR },
    });

    assert.equal(
      pickChatAgentForTurn({ conversationId: GROUP, botKey: HR_BOT, resolve }),
      HR,
    );
  });

  it('accepts the platform fallback when nothing is bound', () => {
    const { resolve } = resolverOver({
      [GROUP]: { decision: 'fallback', chatAgent: DEFAULT },
      [HR_BOT]: { decision: 'fallback', chatAgent: DEFAULT },
    });

    assert.equal(
      pickChatAgentForTurn({ conversationId: GROUP, botKey: HR_BOT, resolve }),
      DEFAULT,
    );
  });

  it('an unknown bot key still reaches the platform fallback', () => {
    const { resolve } = resolverOver({
      '28:unprovisioned': { decision: 'fallback', chatAgent: DEFAULT },
    });

    assert.equal(
      pickChatAgentForTurn({ botKey: '28:unprovisioned', resolve }),
      DEFAULT,
    );
  });

  it('returns undefined when every key is rejected', () => {
    const { resolve } = resolverOver({});

    assert.equal(
      pickChatAgentForTurn({ conversationId: GROUP, botKey: HR_BOT, resolve }),
      undefined,
    );
  });

  it('probes a key that is both conversation and bot exactly once', () => {
    const { resolve, probes } = resolverOver({
      [HR_BOT]: { decision: 'bound', chatAgent: HR },
    });

    const picked = pickChatAgentForTurn({
      conversationId: HR_BOT,
      botKey: HR_BOT,
      resolve,
    });

    assert.equal(picked, HR);
    assert.deepEqual(probes, [HR_BOT]);
  });

  it('ignores a decision that carries no agent', () => {
    // `bound` without a chatAgent is a resolver that matched an agent the
    // registry could not build. Treat it as no answer, not as a null agent.
    const { resolve } = resolverOver({
      [HR_BOT]: { decision: 'bound', exclusive: true },
      [GROUP]: { decision: 'bound', chatAgent: MESSIAS },
    });

    assert.equal(
      pickChatAgentForTurn({ conversationId: GROUP, botKey: HR_BOT, resolve }),
      MESSIAS,
    );
  });

  it('is a no-op without keys', () => {
    const { resolve, probes } = resolverOver({});
    assert.equal(pickChatAgentForTurn({ resolve }), undefined);
    assert.deepEqual(probes, []);
  });
});

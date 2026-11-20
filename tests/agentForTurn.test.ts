import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// Pure precedence logic — no botbuilder, no adapter, so the relative import
// is bundled by the test transpiler and needs no built `dist/`.
import { pickTurnDecision, type TurnAgentDecision } from '../src/agentForTurn.js';
import type { ChatAgent, SystemNotice } from '../src/kernel-types.js';

const HR = { id: 'hr' } as unknown as ChatAgent;
const MESSIAS = { id: 'messias' } as unknown as ChatAgent;
/** The management layer in front of no agent: what a `reject` carries (#2055). */
const MANAGEMENT = { id: 'management' } as unknown as ChatAgent;
const CONVERSATION_MANAGEMENT = { id: 'management-of-the-conversation' } as unknown as ChatAgent;
const NOTICE = { code: 'no_agent', text: 'Kein Agent.', locale: 'de' } as unknown as SystemNotice;
const DISABLED_NOTICE = { code: 'agent_disabled', text: 'Mira ist deaktiviert.', locale: 'de' } as unknown as SystemNotice;

const HR_BOT = '28:3d78d742-eefb-4fb2-bae5-3687f24c46fc';
const MESSIAS_BOT = '28:19ad2729-f7d3-4099-9d2a-7da1230c9533';
/** The one group chat every provisioned bot sits in. */
const GROUP = '19:5c8a1f60deadbeef@thread.skype';

const REJECT: TurnAgentDecision = { decision: 'reject', chatAgent: MANAGEMENT, notice: NOTICE };

/** Build a resolver over a fixed key→decision table, recording probe order. */
function resolverOver(table: Record<string, TurnAgentDecision>): {
  resolve: (key: string) => TurnAgentDecision;
  probes: string[];
} {
  const probes: string[] = [];
  return {
    probes,
    resolve: (key) => {
      probes.push(key);
      return table[key] ?? REJECT;
    },
  };
}

describe('pickTurnDecision', () => {
  it('routes each provisioned bot in a shared chat to its own agent', () => {
    const table: Record<string, TurnAgentDecision> = {
      [HR_BOT]: { decision: 'bound', chatAgent: HR, exclusive: true },
      [MESSIAS_BOT]: { decision: 'bound', chatAgent: MESSIAS, exclusive: true },
    };

    const hr = pickTurnDecision({ conversationId: GROUP, botKey: HR_BOT, resolve: resolverOver(table).resolve });
    const messias = pickTurnDecision({ conversationId: GROUP, botKey: MESSIAS_BOT, resolve: resolverOver(table).resolve });

    assert.equal(hr?.chatAgent, HR);
    assert.equal(messias?.chatAgent, MESSIAS);
  });

  it('an exclusive bot hit beats a binding on the shared conversation', () => {
    // This is the regression that made the whole feature a no-op the other
    // way round: one binding on the group chat, three bots, one voice.
    const { resolve } = resolverOver({
      [GROUP]: { decision: 'bound', chatAgent: MESSIAS },
      [HR_BOT]: { decision: 'bound', chatAgent: HR, exclusive: true },
    });

    assert.equal(pickTurnDecision({ conversationId: GROUP, botKey: HR_BOT, resolve })?.chatAgent, HR);
  });

  it('an exclusive REJECT of the bot also beats a binding on the conversation', () => {
    // A provisioned bot whose agent is disabled never answers for another
    // agent bound to the shared chat: the refusal is its own (#2055).
    const refusal: TurnAgentDecision = {
      decision: 'reject',
      chatAgent: MANAGEMENT,
      notice: DISABLED_NOTICE,
      exclusive: true,
    };
    const { resolve, probes } = resolverOver({
      [GROUP]: { decision: 'bound', chatAgent: MESSIAS },
      [HR_BOT]: refusal,
    });

    const picked = pickTurnDecision({ conversationId: GROUP, botKey: HR_BOT, resolve });

    assert.equal(picked, refusal);
    assert.deepEqual(probes, [HR_BOT]);
  });

  it('does not probe the conversation once the bot key answered exclusively', () => {
    const { resolve, probes } = resolverOver({
      [HR_BOT]: { decision: 'bound', chatAgent: HR, exclusive: true },
    });

    pickTurnDecision({ conversationId: GROUP, botKey: HR_BOT, resolve });

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

    assert.equal(pickTurnDecision({ conversationId: GROUP, botKey: HR_BOT, resolve })?.chatAgent, MESSIAS);
  });

  it('falls through to the bot catch-all when the conversation is rejected', () => {
    const { resolve } = resolverOver({
      [GROUP]: REJECT,
      [HR_BOT]: { decision: 'bound', chatAgent: HR },
    });

    assert.equal(pickTurnDecision({ conversationId: GROUP, botKey: HR_BOT, resolve })?.chatAgent, HR);
  });

  it('there is no fallback agent: when nothing is bound the answer is a reject with the notice', () => {
    const { resolve } = resolverOver({});

    const picked = pickTurnDecision({ conversationId: GROUP, botKey: HR_BOT, resolve });

    assert.equal(picked?.decision, 'reject');
    assert.equal(picked?.notice, NOTICE);
    assert.equal(picked?.chatAgent, MANAGEMENT, 'the management layer still carries /agent');
  });

  it('a reject names the conversation\'s management layer: a /agent selection is stored for the chat where it was typed', () => {
    const { resolve } = resolverOver({
      [GROUP]: { decision: 'reject', chatAgent: CONVERSATION_MANAGEMENT, notice: NOTICE },
      [HR_BOT]: { decision: 'reject', chatAgent: MANAGEMENT, notice: NOTICE },
    });

    assert.equal(pickTurnDecision({ conversationId: GROUP, botKey: HR_BOT, resolve })?.chatAgent, CONVERSATION_MANAGEMENT);
  });

  it('an unknown bot key without a conversation is its own reject', () => {
    const { resolve } = resolverOver({});

    const picked = pickTurnDecision({ botKey: '28:unprovisioned', resolve });

    assert.equal(picked?.decision, 'reject');
    assert.equal(picked?.chatAgent, MANAGEMENT);
  });

  it('hands the speaker of the winning decision along (known before the answer)', () => {
    const speaker = { kind: 'agent', agentId: 'a-hr', slug: 'hr', displayName: 'Hanna' } as const;
    const { resolve } = resolverOver({ [HR_BOT]: { decision: 'bound', chatAgent: HR, exclusive: true, speaker } });

    assert.deepEqual(pickTurnDecision({ conversationId: GROUP, botKey: HR_BOT, resolve })?.speaker, speaker);
  });

  it('probes a key that is both conversation and bot exactly once', () => {
    const { resolve, probes } = resolverOver({
      [HR_BOT]: { decision: 'bound', chatAgent: HR },
    });

    const picked = pickTurnDecision({ conversationId: HR_BOT, botKey: HR_BOT, resolve });

    assert.equal(picked?.chatAgent, HR);
    assert.deepEqual(probes, [HR_BOT]);
  });

  it('is a no-op without keys', () => {
    const { resolve, probes } = resolverOver({});
    assert.equal(pickTurnDecision({ resolve }), undefined);
    assert.deepEqual(probes, []);
  });
});

import { afterEach, describe, it, mock } from 'node:test';
import { strict as assert } from 'node:assert';

import { paceAckText } from '@omadia/channel-sdk';

import { TEAMS_REPLY_WINDOW_MS } from '../src/teamsPace.js';

import { ANNA, messageActivity, MIRA, SimulatedKernel, simulatedBot } from './support/teamsSimulation.js';

/**
 * Spec 2042 §6 criterion 10, byte5ai/omadia-channel-teams#42 — the reply
 * window. A message turn that outlives the Bot Framework's window is answered
 * as a follow-up on a proactive continuation of the conversation; a turn that
 * fails ends in a notice and never in silence.
 */

const MIRA_SPEAKER = { kind: 'agent', agentId: 'a-mira', slug: 'mira', displayName: 'Mira' };

type Sent = Record<string, unknown>;
const cards = (sent: unknown[]): Sent[] =>
  (sent as unknown[]).filter((a): a is Sent => typeof a === 'object' && a !== null && (a as Sent)['type'] !== 'typing');

function botTakingMs(ms: number, answerText = 'Fertig.') {
  mock.timers.enable({ apis: ['setTimeout', 'setInterval', 'Date'] });
  const sim = simulatedBot(new SimulatedKernel(), {
    captureConversationReference: () => {},
    emitMembershipEvent: () => {},
    resolve: () => ({ decision: 'bound', chatAgent: sim.agent, speaker: MIRA_SPEAKER }),
    chat: async () => {
      mock.timers.tick(ms);
      await Promise.resolve();
      return { text: answerText, speaker: MIRA_SPEAKER };
    },
  });
  const followUps: { sent: unknown[] }[] = [];
  sim.bot.attachProactiveSend(async (_reference, build) => {
    const sent: unknown[] = [];
    followUps.push({ sent });
    await build({
      activity: {},
      sendActivity: async (outgoing: unknown) => {
        sent.push(outgoing);
        return { id: 'follow-up' };
      },
      sendActivities: async (outgoing: unknown[]) => {
        sent.push(...outgoing);
        return outgoing.map(() => ({ id: 'follow-up' }));
      },
      turnState: new Map<unknown, unknown>(),
    } as never);
  });
  return { sim, followUps };
}

describe('reply window of a Teams message turn (#42)', () => {
  afterEach(() => mock.timers.reset());

  it('is the 15 s of the Bot Framework, above the acknowledgement threshold', () => {
    assert.equal(TEAMS_REPLY_WINDOW_MS, 15_000);
  });

  it('a turn that outlives the window: one acknowledgement through the live turn, the answer as a follow-up', async () => {
    const { sim, followUps } = botTakingMs(TEAMS_REPLY_WINDOW_MS + 1_000);

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'lange arbeit', mention: true }));

    assert.deepEqual((sim.sent as unknown[]).filter((a) => typeof a === 'string'), [paceAckText('de')], 'the live turn only acknowledged');
    assert.equal(cards(sim.sent).length, 0, 'no answer through the live turn context');
    assert.equal(followUps.length, 1);
    const answer = cards(followUps[0]?.sent ?? []);
    assert.equal(answer.length, 1);
    assert.ok(JSON.stringify(answer[0]).includes('Fertig.'));
  });

  it('a turn within the window answers through the live turn, no follow-up', async () => {
    const { sim, followUps } = botTakingMs(9_000);

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'mittel', mention: true }));

    assert.equal(followUps.length, 0);
    assert.ok(JSON.stringify(cards(sim.sent)).includes('Fertig.'));
  });

  it('when the follow-up cannot be posted the live turn is the last way to the person', async () => {
    const { sim } = botTakingMs(TEAMS_REPLY_WINDOW_MS + 1_000);
    sim.bot.attachProactiveSend(async () => {
      throw new Error('conversation reference expired');
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'lange arbeit', mention: true }));

    assert.ok(JSON.stringify(cards(sim.sent)).includes('Fertig.'), 'the answer was not lost');
  });
});

describe('a failing Teams turn ends in a notice (criterion 10)', () => {
  afterEach(() => mock.timers.reset());

  it('the agent call throwing is answered with a visible error line, after the acknowledgement, never silence', async () => {
    mock.timers.enable({ apis: ['setTimeout', 'setInterval', 'Date'] });
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: () => ({ decision: 'bound', chatAgent: sim.agent, speaker: MIRA_SPEAKER }),
      chat: async () => {
        mock.timers.tick(9_000);
        await Promise.resolve();
        throw new Error('modell nicht erreichbar');
      },
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'lange arbeit', mention: true }));

    const texts = (sim.sent as unknown[]).filter((a): a is string => typeof a === 'string');
    assert.equal(texts.length, 2, 'the acknowledgement, then the notice');
    assert.equal(texts[0], paceAckText('de'));
    assert.match(texts[1] ?? '', /Fehler aufgetreten/);
    assert.equal(sim.appended.length, 0, 'a failed turn is not conversation content');
  });
});

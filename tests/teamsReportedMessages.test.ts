import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { ANNA, GROUP_CHAT_ID, messageActivity, MIRA, SimulatedKernel, simulatedBot } from './support/teamsSimulation.js';

/**
 * Kernel #2086 (28da9b54a): a reply that answers an open question ("2") must
 * not also start the turn the channel runs for an addressed message. In the
 * `addressed` mode Teams reports the message with `reportConversationMessage`
 * before the turn and starts none when the kernel consumed it. A slow or
 * failing kernel never delays or blocks the turn.
 */

function setup(reportMode: SimulatedKernel['reportMode'], mode: 'all' | 'addressed' = 'addressed') {
  const kernel = new SimulatedKernel();
  kernel.reportMode = reportMode;
  kernel.subscribe(GROUP_CHAT_ID, MIRA, mode);
  const sim = simulatedBot(kernel, { captureConversationReference: () => {}, emitMembershipEvent: () => {}, answerText: 'Antwort.' });
  return { kernel, sim };
}
const mention = (text: string) => messageActivity({ bot: MIRA, from: ANNA, text, mention: true });

describe('a typed answer to an open question starts no turn (#2086)', () => {
  it('consumed by the kernel: no turn, no typing, no acknowledgement, no answer', async () => {
    const { kernel, sim } = setup('consume');

    await sim.deliver(mention('2'));

    assert.equal(kernel.reported.length, 1);
    assert.equal(kernel.reported[0]?.addressedToBot, true);
    assert.equal(kernel.reported[0]?.text, 'Mira 2');
    assert.deepEqual(sim.turns, []);
    assert.deepEqual(sim.sent, []);
  });

  it('not consumed: the event is recorded and the turn runs as before', async () => {
    const { kernel, sim } = setup('pass');

    await sim.deliver(mention('wie ist der Stand?'));

    assert.equal(kernel.reported.length, 1);
    assert.equal(sim.turns.length, 1);
    assert.ok(sim.sent.length > 0, 'the turn answered');
  });

  it('a verdict that does not come in time does not delay the turn: it runs, answeredCard unknown is false', async () => {
    const { sim } = setup('hang');
    const started = Date.now();

    await sim.deliver(mention('2'));

    assert.equal(sim.turns.length, 1);
    assert.ok(Date.now() - started < 1_000, 'bounded by the report timeout, not by the kernel');
  });

  it('a failing report is logged and the turn runs', async () => {
    const { sim } = setup('throw');

    await sim.deliver(mention('2'));

    assert.equal(sim.turns.length, 1);
  });

  it('a kernel without reportConversationMessage: the message is emitted and the turn runs, as before', async () => {
    const { kernel, sim } = setup(undefined);

    await sim.deliver(mention('2'));

    assert.equal(kernel.reported.length, 0);
    assert.equal(kernel.events.length, 1, 'emitted');
    assert.equal(sim.turns.length, 1);
  });

  it('the bot\'s own 1:1 chat is addressed too: a consumed reply starts no turn there either', async () => {
    const kernel = new SimulatedKernel();
    kernel.reportMode = 'consume';
    kernel.subscribe('a:1personalchat', MIRA, 'addressed');
    const sim = simulatedBot(kernel, { captureConversationReference: () => {}, emitMembershipEvent: () => {} });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: '2', conversationType: 'personal', conversationId: 'a:1personalchat' }));

    assert.equal(kernel.reported[0]?.addressedToBot, true);
    assert.deepEqual(sim.turns, []);
  });

  it('a message nobody addressed is not reported and starts no turn (the mention gate)', async () => {
    const { kernel, sim } = setup('consume');

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'unter uns' }));

    assert.deepEqual(kernel.reported, []);
    assert.deepEqual(sim.turns, []);
  });

  it('where the engine drives every turn (all) nothing changes: the message is emitted, no report, no turn', async () => {
    const { kernel, sim } = setup('consume', 'all');

    await sim.deliver(mention('2'));

    assert.deepEqual(kernel.reported, []);
    assert.equal(kernel.events.length, 1);
    assert.deepEqual(sim.turns, []);
  });
});

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  createOverdueCardSweep,
  createTeamsConversationSendAdapter,
  OVERDUE_GRACE_MS,
  QUESTION_CARD_ANSWER_VERB,
  TeamsConversationReferenceCache,
  TeamsQuestionCards,
} from '@omadia/channel-teams';
import type { QuestionCard } from '../src/kernel-types.js';
import type { StoredQuestionCard, TeamsQuestionCardPersistence } from '../src/teamsQuestionCardStore.js';

import {
  ANNA,
  cardExecuteActivity,
  GROUP_CHAT_ID,
  messageActivity,
  MIRA,
  simulatedBot,
  SimulatedConnector,
  SimulatedKernel,
} from './support/teamsSimulation.js';

/**
 * 0.40.2 — A CARD WHOSE CLOSE NEVER ARRIVED CLOSES ITSELF.
 *
 * Rollout pass 6, observation O2: an operator's kill switch made the kernel
 * refuse the engine's card update. The refusal happens in the kernel, so the
 * Teams channel was never called; the engine does not retry a card update.
 * The row stayed `open` and the card kept its inputs. These tests reproduce
 * that from the channel's side (a card is posted, no `updateCard` follows)
 * over the simulated transport. What a Teams client shows after the
 * replacement is not verified here.
 */

type Json = Record<string, unknown>;

const NOW = Date.parse('2026-11-19T12:36:00.000Z');
const CLOSES_AT = '2026-11-19T13:36:00.000Z';
const MINUTE = 60_000;

const CARD: QuestionCard = {
  cardId: 'f9-v1',
  proposalId: 'p-1',
  proposalVersion: 1,
  title: 'Abstimmung: Raum für den Testlauf',
  locale: 'de',
  closesAt: CLOSES_AT,
  questions: [
    {
      id: 'raum',
      label: 'Welcher Raum?',
      kind: 'single_choice',
      required: true,
      options: [
        { id: 'a', label: 'A' },
        { id: 'b', label: 'B' },
      ],
    },
  ],
};

/** The table, in memory; records go through JSON as the column stores them. */
class CardTable implements TeamsQuestionCardPersistence {
  readonly rows = new Map<string, StoredQuestionCard>();

  async save(record: StoredQuestionCard): Promise<void> {
    this.rows.set(`${record.botAppId}|${record.card.cardId}`, JSON.parse(JSON.stringify(record)) as StoredQuestionCard);
  }

  async loadByCard(botAppId: string, cardId: string): Promise<StoredQuestionCard | undefined> {
    return this.rows.get(`${botAppId.toLowerCase()}|${cardId}`);
  }

  async loadByMessage(botAppId: string, messageId: string): Promise<StoredQuestionCard | undefined> {
    return [...this.rows.values()].find((row) => row.botAppId === botAppId.toLowerCase() && row.messageId === messageId);
  }

  async listOverdueOpen(before: Date, limit: number): Promise<readonly StoredQuestionCard[]> {
    return [...this.rows.values()]
      .filter((row) => !row.closed && row.card.closesAt !== undefined && Date.parse(row.card.closesAt) <= before.getTime())
      .slice(0, limit);
  }

  stateOf(cardId: string): string {
    return this.rows.get(`${MIRA.appId.toLowerCase()}|${cardId}`)?.closed ? 'closed' : 'open';
  }
}

function harness() {
  const clock = { now: NOW };
  const table = new CardTable();
  const kernel = new SimulatedKernel();
  kernel.subscribe(GROUP_CHAT_ID, MIRA);
  const refs = new TeamsConversationReferenceCache();
  const sim = simulatedBot(kernel, { captureConversationReference: (context) => refs.capture(context), emitMembershipEvent: () => {} });
  refs.capture({ activity: messageActivity({ bot: MIRA, from: ANNA, text: 'hi' }) } as unknown as TurnContext);
  const connector = new SimulatedConnector();
  const logged: string[] = [];

  /** Start (or restart) the channel: new in-memory registry, same table. */
  function start() {
    const cards = new TeamsQuestionCards({ now: () => clock.now, persistence: table });
    sim.bot.attachQuestionCards(cards);
    const provider = createTeamsConversationSendAdapter({ refs, sendProactive: connector.sendProactive, cards });
    const sweep = createOverdueCardSweep({ cards, refs, sendProactive: connector.sendProactive, log: (m) => logged.push(m) });
    return { provider, sweep };
  }

  async function post(provider: ReturnType<typeof start>['provider'], card: QuestionCard): Promise<{ content: Json; messageId: string }> {
    const outcome = await provider.sendToConversation(GROUP_CHAT_ID, { text: '', card } as never, { asChannelKey: MIRA.key });
    assert.equal(outcome.outcome, 'delivered', JSON.stringify(outcome));
    const content = (connector.sent.at(-1)?.activity['attachments'] as { content: Json }[])[0]?.content as Json;
    return { content, messageId: (outcome as { cardMessageId?: string }).cardMessageId ?? '' };
  }

  async function click(messageId: string, data: Json): Promise<{ type: string; value: unknown }> {
    const before = sim.sent.length;
    await sim.deliver(cardExecuteActivity({ bot: MIRA, from: ANNA, conversationId: GROUP_CHAT_ID, cardMessageId: messageId, verb: QUESTION_CARD_ANSWER_VERB, data }));
    const response = (sim.sent.slice(before) as Json[]).find((activity) => activity['type'] === 'invokeResponse');
    assert.ok(response, 'the click was answered');
    return (response['value'] as { body: { type: string; value: unknown } }).body;
  }

  return { clock, table, kernel, connector, logged, start, post, click };
}

const actionData = (content: Json): Json => ((content['actions'] as { data: Json }[]).at(-1) as { data: Json }).data;

describe('teams · a question card whose close never arrived (O2)', () => {
  it('stays as it is until its closing time has passed by the grace period', async () => {
    const h = harness();
    const { provider, sweep } = h.start();
    await h.post(provider, CARD);

    assert.equal(await sweep(), 0, 'open and not due');
    h.clock.now = Date.parse(CLOSES_AT) + OVERDUE_GRACE_MS - 1;
    assert.equal(await sweep(), 0, 'due, but the engine may still bring its result');

    assert.equal(h.table.stateOf(CARD.cardId), 'open');
    assert.deepEqual(h.connector.updated, []);
  });

  it('is replaced by its closed form and its row is closed, also after a restart of the channel', async () => {
    const h = harness();
    const { messageId, content } = await h.post(h.start().provider, CARD);
    // The engine closed the vote at 12:38; the kernel refused its card
    // update, so nothing reached the channel. A deploy follows.
    h.clock.now = Date.parse(CLOSES_AT) + OVERDUE_GRACE_MS;
    const { sweep } = h.start();

    assert.equal(await sweep(), 1);

    assert.equal(h.table.stateOf(CARD.cardId), 'closed');
    assert.deepEqual(h.connector.updated.map((update) => [update.asBot, update.id]), [[MIRA.key, messageId]]);
    const replaced = JSON.stringify(h.connector.updated[0]?.activity);
    assert.ok(replaced.includes(CARD.title) && replaced.includes('Diese Abstimmung ist nicht mehr aktiv.'));
    assert.ok(!replaced.includes('Action.Execute'), 'no inputs left on the card');
    assert.equal(h.logged.length, 1);
    assert.match(h.logged[0] ?? '', /'f9-v1' was still open past its closing time .* replaced by its closed form/);

    // A late click gets the closed card; nothing reaches the kernel.
    const late = await h.click(messageId, { ...actionData(content), q0: 'a' });
    assert.equal(late.type, 'application/vnd.microsoft.card.adaptive');
    assert.deepEqual(h.kernel.events, []);
    // And the next sweep finds nothing to do.
    assert.equal(await sweep(), 0);
    assert.equal(h.connector.updated.length, 1);
  });

  it('closes the row even when Teams refuses the replacement, and says so once', async () => {
    const h = harness();
    const { provider, sweep } = h.start();
    await h.post(provider, CARD);
    h.clock.now = Date.parse(CLOSES_AT) + 10 * MINUTE;
    h.connector.updateFailure = Object.assign(new Error('message not found'), { statusCode: 404 });

    assert.equal(await sweep(), 1);

    assert.equal(h.table.stateOf(CARD.cardId), 'closed');
    assert.match(h.logged[0] ?? '', /could not be replaced \(message not found\)/);
    assert.equal(await sweep(), 0, 'not tried again');
  });

  it("leaves alone a card the engine closed, a card without a closing time, and still takes the engine's later result", async () => {
    const h = harness();
    const { provider, sweep } = h.start();
    const closedByEngine = await h.post(provider, { ...CARD, cardId: 'c-engine' });
    await provider.updateCard(
      GROUP_CHAT_ID,
      closedByEngine.messageId,
      { state: 'closed', title: CARD.title, resultLines: ['Ergebnis: A'], closedAt: '2026-11-19T12:38:16.000Z' },
      { asChannelKey: MIRA.key },
    );
    await h.post(provider, { ...CARD, cardId: 'c-open-end', closesAt: undefined } as QuestionCard);
    const overdue = await h.post(provider, CARD);
    h.clock.now = Date.parse(CLOSES_AT) + 10 * MINUTE;

    assert.equal(await sweep(), 1);

    assert.deepEqual(h.connector.updated.map((update) => update.id), [closedByEngine.messageId, overdue.messageId]);
    assert.ok(JSON.stringify(h.table.rows.get(`${MIRA.appId.toLowerCase()}|c-engine`)?.closed).includes('Ergebnis: A'), "the engine's result stays");
    assert.equal(h.table.stateOf('c-open-end'), 'open');

    // The engine's update arrives after all (the kill switch was released
    // and something retried): the result replaces the channel's own close.
    const late = await provider.updateCard(
      GROUP_CHAT_ID,
      overdue.messageId,
      { state: 'closed', title: CARD.title, resultLines: ['Ergebnis: B'], closedAt: '2026-11-19T13:50:00.000Z' },
      { asChannelKey: MIRA.key },
    );
    assert.deepEqual(late, { outcome: 'updated' });
    assert.ok(JSON.stringify(h.table.rows.get(`${MIRA.appId.toLowerCase()}|${CARD.cardId}`)?.closed).includes('Ergebnis: B'));
  });
});

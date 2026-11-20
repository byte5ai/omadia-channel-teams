import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { buildAnswerCard } from '@omadia/channel-teams';

import { readSideChannelData, SIDE_CHANNEL_MAX_ROWS } from '../src/teamsSideChannelData.js';
import { ANNA, messageActivity, MIRA, SimulatedKernel, simulatedBot } from './support/teamsSimulation.js';

/**
 * Spec 2042 §5.13, criteria 25/35 (kernel #2085): rows the agent shows people
 * without its model reading them. `SemanticAnswer.data` is drawn as native
 * tables next to `textWithoutData`; an answer without the fields renders from
 * `text` as before. The fields are read structurally (S4 is not on every SDK).
 */

const table = (rowCount: number, extra: Record<string, unknown> = {}) => ({
  kind: 'table' as const,
  caption: 'Offene Rechnungen',
  columns: ['Nr', 'Kunde', 'Betrag'],
  rows: Array.from({ length: rowCount }, (_, n) => [`R-${String(n)}`, `Kunde ${String(n)}`, n * 10]),
  ...extra,
});

const cardJson = (card: unknown): { body: Record<string, unknown>[] } =>
  (card as { content: { body: Record<string, unknown>[] } }).content;
const tablesOf = (card: unknown) => cardJson(card).body.filter((e) => e['type'] === 'Table');
const textsOf = (card: unknown): string[] =>
  cardJson(card).body.filter((e) => e['type'] === 'TextBlock').map((e) => String(e['text']));

describe('readSideChannelData', () => {
  it('reads data together with textWithoutData', () => {
    const read = readSideChannelData({ text: 'a', data: [table(2)], textWithoutData: 'Hier die Liste.' });
    assert.equal(read?.tables.length, 1);
    assert.equal(read?.textWithoutData, 'Hier die Liste.');
  });

  for (const [name, answer] of [
    ['no data', { text: 'a' }],
    ['empty data', { text: 'a', data: [], textWithoutData: 'a' }],
    ['data without textWithoutData: the tables are inside text already, drawing them again would show them twice', { text: 'a', data: [table(1)] }],
    ['a payload this plugin cannot draw: text stays the source', { text: 'a', data: [table(1), { kind: 'chart' }], textWithoutData: 'a' }],
    ['a malformed table', { text: 'a', data: [{ kind: 'table', columns: 'x', rows: [] }], textWithoutData: 'a' }],
  ] as const) {
    it(`is undefined for ${name}`, () => {
      assert.equal(readSideChannelData(answer), undefined);
    });
  }
});

describe('side-channel data on the answer card', () => {
  it('draws a native Table with caption, header and rows, under the text without the data rendering', () => {
    const card = buildAnswerCard({ answer: 'Hier die Liste.', dataTables: [table(3)] });

    const tables = tablesOf(card);
    assert.equal(tables.length, 1);
    assert.equal((tables[0]?.['rows'] as unknown[]).length, 1 + 3, 'header + 3 rows');
    const texts = textsOf(card);
    assert.deepEqual(texts.slice(0, 1), ['Hier die Liste.']);
    assert.ok(texts.includes('Offene Rechnungen'));
    assert.ok(JSON.stringify(card).includes('Kunde 2'));
  });

  it('shows at most 50 rows and says how many more there are', () => {
    const card = buildAnswerCard({ answer: 'Liste', dataTables: [table(SIDE_CHANNEL_MAX_ROWS + 12)] });

    assert.equal((tablesOf(card)[0]?.['rows'] as unknown[]).length, 1 + SIDE_CHANNEL_MAX_ROWS);
    assert.ok(textsOf(card).includes('… 12 weitere Zeilen'));
  });

  it('draws two tables in order, null as an empty cell, long cells cut', () => {
    const card = buildAnswerCard({
      answer: 'Zwei Listen',
      dataTables: [
        { kind: 'table', columns: ['A'], rows: [[null], ['x'.repeat(500)]] },
        table(1, { caption: 'Zweite' }),
      ],
    });

    assert.equal(tablesOf(card).length, 2);
    const json = JSON.stringify(card);
    assert.ok(!json.includes('x'.repeat(201)), 'a cell is cut at 200 characters');
    assert.ok(json.includes('Zweite'));
  });

  it('an empty table says so instead of drawing nothing', () => {
    const card = buildAnswerCard({ answer: 'Nichts gefunden', dataTables: [{ kind: 'table', caption: 'Leer', columns: ['A'], rows: [] }] });

    assert.equal(tablesOf(card).length, 0);
    assert.ok(textsOf(card).includes('—'));
  });

  it('shrinks the rows, not the answer text, when the card would exceed the size cap', () => {
    const wide = {
      kind: 'table' as const,
      columns: ['A', 'B', 'C', 'D'],
      rows: Array.from({ length: SIDE_CHANNEL_MAX_ROWS }, (_, n) => [1, 2, 3, 4].map((c) => `${'wert '.repeat(30)}${String(n)}-${String(c)}`)),
    };
    const card = buildAnswerCard({ answer: 'Ganzer Text bleibt.', dataTables: [wide] });

    assert.ok(Buffer.byteLength(JSON.stringify(card), 'utf8') <= 28_000, 'under the Teams card cap');
    assert.ok(textsOf(card).includes('Ganzer Text bleibt.'));
    assert.ok(textsOf(card).some((t) => /weitere Zeilen$/.test(t)));
  });

  it('colours a masked value in a data table like in a markdown table', () => {
    const card = buildAnswerCard({ answer: 'Liste', dataTables: [table(1)], maskedValues: ['Kunde 0'] });

    assert.ok(JSON.stringify(card).includes('"color":"Accent"'));
  });
});

describe('side-channel data through the bot', () => {
  const MARKDOWN = '| Nr | Kunde | Betrag |\n| --- | --- | --- |\n| R-0 | Kunde 0 | 0 |';
  async function deliver(answer: Record<string, unknown>) {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      answer,
    });
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'zeig die Rechnungen', mention: true }));
    return sim;
  }
  const posted = (sim: { sent: unknown[] }) => (sim.sent as Record<string, unknown>[]).filter((a) => a['type'] !== 'typing');

  it('draws the data natively and shows the text without its rendering: the table is on the card once', async () => {
    const sim = await deliver({ text: `Hier die Liste.\n\n${MARKDOWN}`, textWithoutData: 'Hier die Liste.', data: [table(1)] });

    const [card] = posted(sim);
    const attachment = (card?.['attachments'] as unknown[])[0];
    assert.equal(tablesOf(attachment).length, 1, 'exactly one table');
    assert.ok(textsOf(attachment).includes('Hier die Liste.'));
    assert.ok(!JSON.stringify(attachment).includes('| --- |'), 'no markdown table next to the native one');
  });

  it('an answer without the fields renders from text, the markdown table becomes the native table as before', async () => {
    const sim = await deliver({ text: `Hier die Liste.\n\n${MARKDOWN}` });

    const attachment = (posted(sim)[0]?.['attachments'] as unknown[])[0];
    assert.equal(tablesOf(attachment).length, 1);
  });

  it('the history never holds the rows: it reaches a model', async () => {
    const sim = await deliver({ text: `Hier die Liste.\n\n${MARKDOWN}`, textWithoutData: 'Hier die Liste.', data: [table(1)] });

    assert.deepEqual(sim.appended.map((turn) => turn.assistantAnswer), ['Hier die Liste.']);
  });

  it('a long answer goes out as plain text with the tables in it, once', async () => {
    const prose = Array.from({ length: 700 }, (_, n) => `Absatz ${String(n)}: ${'lorem ipsum '.repeat(4)}`).join('\n\n');
    const sim = await deliver({ text: `${prose}\n\n${MARKDOWN}`, textWithoutData: prose, data: [table(1)] });

    const all = posted(sim).map((a) => String(a['text'] ?? '')).join('\n');
    assert.equal(all.split('| R-0 |').length - 1, 1, 'the table is in the plain text exactly once');
    assert.ok(posted(sim).every((a) => a['attachments'] === undefined));
  });
});

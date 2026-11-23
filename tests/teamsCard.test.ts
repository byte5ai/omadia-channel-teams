import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

import {
  aiChipDue,
  aiLabelEntity,
  AI_MARKER_CHIP,
  ANSWER_DETAILS_CONTAINER_ID,
  admitTeamsDecorations,
  buildAnswerCard,
  buildAnswerDetailsElements,
  buildChoiceAskCard,
  buildFollowUpsOnlyCard,
  contributionLabel,
  stripFoldedAiDisclosure,
  SYSTEM_NOTICE_LABEL,
  TEAMS_DECORATION_IDS,
} from '@omadia/channel-teams';
import {
  buildAnswerDetails,
  decideDisclosureMarkers,
  DECORATION_REGISTRY,
  InMemoryDisclosureMarkerStore,
  isDecorationAllowed,
  REMOVED_DECORATIONS,
  type AiDisclosure,
  type AiDisclosureFrequency,
  type Contribution,
  type SemanticAnswer,
  type Speaker,
} from '@omadia/channel-sdk';

/**
 * byte5ai/omadia-channel-teams#42 — agents as colleagues (spec 2042 §3 INV-3/4/5,
 * §5.4, §5.14, criteria 5, 29, 31): a minimal answer card, contributions
 * attributed to the consulted agent, `system` notices that never speak as an
 * agent, ONE collapsed details element, every decoration through the registry.
 */

type El = { type?: string; id?: string; text?: string; title?: string; items?: El[]; actions?: El[]; body?: El[]; targetElements?: string[]; isVisible?: boolean; color?: string; style?: string };
type Content = { body: El[]; actions?: El[]; msteams?: Record<string, unknown> };

function contentOf(card: ReturnType<typeof buildAnswerCard>): Content {
  return card.content as Content;
}

/** Every element of a card, depth first, actions of ActionSets and the card included. */
function walk(content: Content): El[] {
  const out: El[] = [];
  const visit = (el: El): void => {
    out.push(el);
    for (const child of [...(el.items ?? []), ...(el.actions ?? [])]) visit(child);
  };
  for (const el of content.body) visit(el);
  for (const action of content.actions ?? []) visit(action);
  return out;
}

const MIRA: Speaker = { kind: 'agent', agentId: 'a-mira', slug: 'mira', displayName: 'Mira' };
const JONAS: Speaker = { kind: 'agent', agentId: 'a-jonas', slug: 'jonas', displayName: 'Jonas' };
const BRIEFING = 'Gern: **Kernpunkte** der Woche.\n\n- Umsatz stabil\n- Zwei Risiken';

const delivered: Contribution = {
  id: 'c1',
  speaker: { kind: 'agent', agentId: 'a-jonas', slug: 'jonas', displayName: 'Jonas' },
  text: BRIEFING,
  status: 'delivered',
};
const failed: Contribution = {
  id: 'c2',
  speaker: { kind: 'agent', agentId: 'a-theo', slug: 'theo', displayName: 'Theo' },
  text: 'Theo konnte nicht antworten: Zeitüberschreitung.',
  status: 'failed',
};

function disclosure(frequency: AiDisclosureFrequency, markers: { textLine: boolean; chip: boolean }): AiDisclosure {
  return { text: 'KI-Hinweis.', level: 'standard', locale: 'de', source: 'default', frequency, markers };
}

describe('aiLabelEntity', () => {
  it('returns the AIGeneratedContent schema.org entity Teams expects', () => {
    const e = aiLabelEntity() as Record<string, unknown>;
    assert.equal(e['type'], 'https://schema.org/Message');
    assert.deepEqual(e['additionalType'], ['AIGeneratedContent']);
  });
});

describe('a normal answer: the text and the AI disclosure, nothing else (criterion 29)', () => {
  it('snapshot: chip plus text, no buttons, no footer, no details', () => {
    const card = buildAnswerCard({ answer: 'Die Antwort ist 42.', aiMarker: true });
    assert.deepEqual(contentOf(card).body, [
      {
        type: 'Container',
        id: 'ai_disclosure_chip',
        spacing: 'None',
        items: [{ type: 'TextBlock', text: AI_MARKER_CHIP, size: 'Small', weight: 'Bolder', color: 'Accent', wrap: true }],
      },
      { type: 'TextBlock', text: 'Die Antwort ist 42.', wrap: true },
    ]);
    assert.equal(contentOf(card).actions, undefined);
  });

  it('snapshot: with the chip not due the card is the text alone', () => {
    const card = buildAnswerCard({ answer: 'Die Antwort ist 42.', aiMarker: false });
    assert.deepEqual(contentOf(card).body, [{ type: 'TextBlock', text: 'Die Antwort ist 42.', wrap: true }]);
    assert.equal(contentOf(card).actions, undefined);
  });

  it('a caller without kernel data keeps the card of old: the chip is drawn', () => {
    assert.ok(JSON.stringify(buildAnswerCard({ answer: 'x' }).content).includes(AI_MARKER_CHIP));
  });
});

describe('AI disclosure frequency at the card renderer (#2087): all four modes', () => {
  const store = new InMemoryDisclosureMarkerStore();
  const chipOf = (aiDisclosure: AiDisclosure | undefined): boolean =>
    JSON.stringify(buildAnswerCard({ answer: 'Antwort', aiMarker: aiChipDue(aiDisclosure) }).content).includes(AI_MARKER_CHIP);
  const decide = (frequency: AiDisclosureFrequency, scope: string, agentId: string, firstTurnOfScope: boolean, now: number) =>
    decideDisclosureMarkers({ frequency, firstTurnOfScope, scope, agentId, now, store });

  it('channel_default: a chip on every card, also after the first answer of the conversation', () => {
    assert.equal(chipOf(disclosure('channel_default', decide('channel_default', 'c-1', 'a-mira', true, 0))), true);
    assert.equal(chipOf(disclosure('channel_default', decide('channel_default', 'c-1', 'a-mira', false, 1))), true);
  });

  it('per_answer: a chip on every card', () => {
    assert.equal(chipOf(disclosure('per_answer', decide('per_answer', 'c-2', 'a-mira', true, 0))), true);
    assert.equal(chipOf(disclosure('per_answer', decide('per_answer', 'c-2', 'a-mira', false, 1))), true);
  });

  it('per_conversation: the chip when due, then not, then again for another agent and after a pause', () => {
    const minute = 60_000;
    assert.equal(chipOf(disclosure('per_conversation', decide('per_conversation', 'c-3', 'a-mira', true, 0))), true);
    assert.equal(chipOf(disclosure('per_conversation', decide('per_conversation', 'c-3', 'a-mira', false, minute))), false);
    assert.equal(chipOf(disclosure('per_conversation', decide('per_conversation', 'c-3', 'a-jonas', false, 2 * minute))), true);
    assert.equal(chipOf(disclosure('per_conversation', decide('per_conversation', 'c-3', 'a-jonas', false, 3 * minute))), false);
    assert.equal(chipOf(disclosure('per_conversation', decide('per_conversation', 'c-3', 'a-jonas', false, 200 * minute))), true);
  });

  it('off: the operator turned the disclosure off, the kernel supplies none, no chip', () => {
    assert.equal(chipOf(undefined), false);
  });

  it('a disclosure from a kernel without markers behaves like channel_default', () => {
    const { markers: _omitted, ...withoutMarkers } = disclosure('channel_default', { textLine: true, chip: true });
    assert.equal(chipOf(withoutMarkers), true);
  });
});

describe('consultation block (spec 2042 §5.4, INV-5)', () => {
  it('snapshot: A\'s text stays A\'s, B\'s block is attributed to B and carries B\'s text unchanged', () => {
    const card = buildAnswerCard({
      answer: 'Ich habe Jonas gefragt, hier seine Einschätzung.',
      aiMarker: false,
      speaker: MIRA,
      contributions: [delivered],
    });
    assert.deepEqual(contentOf(card).body, [
      { type: 'TextBlock', text: 'Ich habe Jonas gefragt, hier seine Einschätzung.', wrap: true },
      {
        type: 'Container',
        id: 'contribution-c1',
        style: 'emphasis',
        spacing: 'Medium',
        separator: true,
        items: [
          { type: 'TextBlock', text: 'Jonas (konsultiert von Mira)', weight: 'Bolder', size: 'Small', color: 'Accent', wrap: true },
          { type: 'TextBlock', text: BRIEFING, wrap: true, spacing: 'Small' },
        ],
      },
    ]);
  });

  it('a failed contribution is shown as failed, with its failure line, never dropped', () => {
    const card = buildAnswerCard({ answer: 'Mein Teil.', aiMarker: false, speaker: MIRA, contributions: [delivered, failed] });
    const block = walk(contentOf(card)).find((el) => el.id === 'contribution-c2');
    assert.ok(block, 'the failed contribution is on the card');
    assert.equal(block.items?.[0]?.text, 'Theo (konsultiert von Mira) – fehlgeschlagen');
    assert.equal(block.items?.[0]?.color, 'Attention');
    assert.equal(block.items?.[1]?.text, failed.text);
  });

  it('keeps the kernel\'s order and names the consulted agent even without a speaker on the answer', () => {
    assert.equal(contributionLabel(delivered, undefined), 'Jonas (konsultiert)');
    assert.equal(contributionLabel(delivered, { kind: 'system' }), 'Jonas (konsultiert)');
    const card = buildAnswerCard({ answer: 'x', aiMarker: false, contributions: [failed, delivered] });
    const ids = walk(contentOf(card)).filter((el) => el.id?.startsWith('contribution-')).map((el) => el.id);
    assert.deepEqual(ids, ['contribution-c2', 'contribution-c1']);
  });

  it('never shortens a contribution: an oversized answer is cut, the colleague\'s block is not', () => {
    const card = buildAnswerCard({
      answer: 'A'.repeat(40_000),
      aiMarker: false,
      speaker: MIRA,
      contributions: [delivered],
    });
    assert.ok(Buffer.byteLength(JSON.stringify(card.content), 'utf8') <= 28_000, 'under the Teams cap');
    assert.ok(JSON.stringify(card.content).includes(JSON.stringify(BRIEFING).slice(1, -1)));
    assert.ok(contentOf(card).body.some((el) => el.text?.endsWith('…')), 'the speaking agent\'s own text was cut');
  });
});

describe('system notice (INV-3, INV-7): never an agent speaking', () => {
  const NOTICE = 'Mira ist deaktiviert und kann gerade nicht antworten.';

  it('snapshot: its own heading, the notice text, no chip, no details, no buttons', () => {
    const card = buildAnswerCard({
      answer: 'Mira ist deaktiviert',
      aiMarker: true,
      speaker: { kind: 'system' },
      systemNotice: { code: 'agent_disabled', text: NOTICE, locale: 'de' },
      followUpOptions: [{ label: 'A', prompt: 'B' }],
    });
    assert.deepEqual(contentOf(card).body, [
      {
        type: 'Container',
        id: 'system-notice',
        style: 'warning',
        items: [
          { type: 'TextBlock', text: SYSTEM_NOTICE_LABEL, weight: 'Bolder', size: 'Small', isSubtle: true, wrap: true },
          { type: 'TextBlock', text: NOTICE, wrap: true, spacing: 'Small' },
        ],
      },
    ]);
    assert.equal(contentOf(card).actions, undefined);
    assert.ok(!JSON.stringify(card.content).includes(AI_MARKER_CHIP), 'a template of the management layer is not AI output');
  });

  it('a withheld answer keeps its colleagues\' contributions, attributed (INV-5, E2E 2026-11-23)', () => {
    const card = buildAnswerCard({
      answer: 'Minervas Antwort wurde zurückgehalten.',
      aiMarker: true,
      speaker: { kind: 'system' },
      systemNotice: {
        code: 'identity_guard_blocked',
        agent: { slug: 'strategy', displayName: 'Minerva', agentId: 'a-minerva' },
        text: 'Minervas Antwort wurde zurückgehalten.',
        locale: 'de',
      },
      contributions: [delivered],
    });
    const body = contentOf(card).body;
    assert.equal((body[0] as { id?: string }).id, 'system-notice', 'the notice first');
    const json = JSON.stringify(card.content);
    assert.ok(json.includes(JSON.stringify(BRIEFING).slice(1, -1)), 'the colleague\'s text, unchanged');
    assert.ok(json.includes('(konsultiert von Minerva)'), 'attributed, consulted by the withheld agent');
    assert.ok(!json.includes(AI_MARKER_CHIP));
  });
});

describe('decision card stays where the agent lets the user choose', () => {
  it('a choice card keeps its option buttons', () => {
    const card = buildChoiceAskCard({
      question: 'Welche Variante?',
      options: [
        { label: 'Kurz', value: 'kurz' },
        { label: 'Lang', value: 'lang' },
      ],
    });
    const content = card.content as { actions?: El[] };
    assert.deepEqual(content.actions?.map((a) => a.title), ['Kurz', 'Lang']);
  });

  it('proposed options become buttons under the answer; without them there is no button row', () => {
    const withOptions = buildAnswerCard({ answer: 'x', aiMarker: false, followUpOptions: [{ label: 'Weiter', prompt: 'Mach weiter' }] });
    assert.deepEqual(
      walk(contentOf(withOptions)).filter((el) => el.type === 'Action.Submit').map((el) => el.title),
      ['Weiter'],
    );
    assert.equal(buildFollowUpsOnlyCard([])?.content, undefined);
    assert.equal(walk(contentOf(buildAnswerCard({ answer: 'x', aiMarker: false }))).filter((el) => el.type?.startsWith('Action')).length, 0);
  });
});

describe('ONE collapsed details element (spec 2042 §5.14, decoration answer_details)', () => {
  const answer: SemanticAnswer = {
    text: 'Antwort',
    speaker: MIRA,
    verifier: { status: 'verified' },
    contributions: [delivered, failed],
    answerSource: 'privacy-render',
    maskedValues: ['4.200 EUR'],
    privacyReceipt: { datasetsInterned: 2, fieldsMasked: 3, fieldsCleartext: 0, verbsExecuted: [], pseudonymProjectionUsed: false } as unknown as SemanticAnswer['privacyReceipt'],
    memoryUsed: true,
    agentsConsulted: [{ label: 'Strategist', status: 'success', toolCalls: 2 }],
  };

  it('is built from the SDK projection: one hidden container, one toggle button', () => {
    const details = buildAnswerDetails(answer);
    const card = buildAnswerCard({ answer: 'Antwort', aiMarker: false, details });
    const container = walk(contentOf(card)).find((el) => el.id === ANSWER_DETAILS_CONTAINER_ID);
    assert.ok(container);
    assert.equal(container.isVisible, false, 'collapsed until the viewer expands it');
    assert.deepEqual(contentOf(card).actions, [
      { type: 'Action.ToggleVisibility', title: 'Details', targetElements: [ANSWER_DETAILS_CONTAINER_ID] },
    ]);
    const text = JSON.stringify(container);
    for (const expected of ['Alle Angaben geprüft', 'Quellsystem', 'Jonas hat beigetragen', 'Theo hat nicht antworten können', '3 Felder wurden maskiert', 'Frühere Gespräche', 'Eingesetzt: Strategist']) {
      assert.ok(text.includes(expected), expected);
    }
  });

  it('a plain answer has no details projection and therefore no element at all', () => {
    assert.equal(buildAnswerDetails({ text: 'Hallo' }), undefined);
    assert.equal(buildAnswerDetailsElements(undefined), undefined);
    const card = buildAnswerCard({ answer: 'Hallo', aiMarker: false, details: buildAnswerDetails({ text: 'Hallo' }) });
    assert.equal(walk(contentOf(card)).some((el) => el.id === ANSWER_DETAILS_CONTAINER_ID), false);
    assert.equal(contentOf(card).actions, undefined);
  });
});

describe('every decoration goes through the registry (criteria 29, 31)', () => {
  const fullAnswer: SemanticAnswer = {
    text: 'Antwort',
    speaker: MIRA,
    verifier: { status: 'partial' },
    contributions: [delivered],
    answerSource: 'privacy-render',
    memoryUsed: true,
    agentsConsulted: [{ label: 'Strategist', status: 'success' }],
  };

  function loadedCard(): Content {
    return contentOf(
      buildAnswerCard({
        answer: 'Antwort',
        aiMarker: true,
        speaker: MIRA,
        details: buildAnswerDetails(fullAnswer),
        contributions: [delivered],
        followUpOptions: [{ label: 'Weiter', prompt: 'Mach weiter' }],
      }),
    );
  }

  it('every decoration id the Teams card code uses is declared for teams and none is a removed one', () => {
    for (const id of Object.values(TEAMS_DECORATION_IDS)) {
      assert.ok(isDecorationAllowed(id, 'teams'), `${id} is declared for teams`);
      assert.ok(!(REMOVED_DECORATIONS as readonly string[]).includes(id), `${id} is not a removed decoration`);
    }
    assert.ok(DECORATION_REGISTRY.length >= Object.values(TEAMS_DECORATION_IDS).length);
  });

  it('the gate drops what the registry does not declare for teams, removed ones included', () => {
    const refused: string[] = [];
    const admitted = admitTeamsDecorations(
      ['ai_disclosure_chip', 'answer_details', 'verifier_badge', 'direct_line_buttons', 'verbosity_buttons', 'consulted_footer', 'made_up_chip'],
      (refusal) => refused.push(refusal.id),
    );
    assert.deepEqual([...admitted].sort(), ['ai_disclosure_chip', 'answer_details']);
    assert.deepEqual(refused.sort(), ['consulted_footer', 'direct_line_buttons', 'made_up_chip', 'verbosity_buttons', 'verifier_badge']);
  });

  it('a fully loaded card holds no undeclared element: only the chip, the colleague, details and proposed buttons', () => {
    const card = loadedCard();
    const ids = walk(card).map((el) => el.id).filter((id): id is string => id !== undefined);
    for (const id of ids) {
      const known = isDecorationAllowed(id, 'teams') || id.startsWith('contribution-');
      assert.ok(known, `element id '${id}' is declared`);
    }
    // The only buttons: the Details toggle and what the agent proposed.
    const buttons = walk(card).filter((el) => el.type?.startsWith('Action.')).map((el) => `${el.type}:${el.title}`);
    assert.deepEqual(buttons.sort(), ['Action.Submit:Weiter', 'Action.ToggleVisibility:Details']);
  });

  it('what the removed decorations showed appears nowhere outside the details element', () => {
    const card = loadedCard();
    const outside = JSON.stringify({
      ...card,
      body: card.body.filter((el) => el.id !== ANSWER_DETAILS_CONTAINER_ID),
    });
    for (const gone of ['Konsultiert:', 'Direkte Antwort', 'Kürzer', 'Mehr Details', 'Direkt mit', 'Fresh Check', 'Nächste Varianten', 'Originaldaten', 'Antwort geprüft', 'Teilweise bestätigt', '🔎', '💬', '🗄']) {
      assert.ok(!outside.includes(gone), `'${gone}' must not be on the card body`);
    }
  });

  it('no source file mentions a Direct Line any more (UI and handler are gone)', () => {
    const srcDir = join(process.cwd(), 'src');
    const offenders = readdirSync(srcDir)
      .filter((name) => name.endsWith('.ts'))
      .filter((name) => /direct.?line/i.test(readFileSync(join(srcDir, name), 'utf8')));
    assert.deepEqual(offenders, []);
  });
});

describe('stripFoldedAiDisclosure', () => {
  const LINE = 'Diese Antwort wurde von einem KI-System erzeugt.';

  it('removes the folded disclosure paragraph from the end of the answer', () => {
    assert.equal(stripFoldedAiDisclosure(`pong 🏓\n\n${LINE}`, LINE), 'pong 🏓');
  });

  it('keeps an operator note that follows the marking line', () => {
    assert.equal(
      stripFoldedAiDisclosure(`Antwort.\n\n${LINE}\n\nOperator-Hinweis.`, LINE),
      'Antwort.\n\nOperator-Hinweis.',
    );
  });

  it('returns the text unchanged when the line was not folded this turn', () => {
    assert.equal(stripFoldedAiDisclosure('Antwort ohne Marking.', LINE), 'Antwort ohne Marking.');
  });

  it('keeps a disclosure-only answer (never empties the card body)', () => {
    assert.equal(stripFoldedAiDisclosure(LINE, LINE), LINE);
  });

  it('is a no-op without a structured disclosure', () => {
    assert.equal(stripFoldedAiDisclosure(`Antwort.\n\n${LINE}`, undefined), `Antwort.\n\n${LINE}`);
  });
});

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import {
  buildAnswerCard,
  buildDirectLineOnlyCard,
  buildFollowUpsOnlyCard,
  DEFAULT_TEAMS_CARD_FEATURES,
  disabledTeamsCardFeatures,
  FOLLOW_UP_VALUE_TYPE,
  FRESH_CHECK_VALUE_TYPE,
  parseBooleanSetting,
  resolveTeamsCardFeatures,
  TEAMS_CARD_FEATURE_CONFIG_KEYS,
  TEAMS_CARD_FEATURE_KEYS,
  VERBOSITY_STEP_VALUE_TYPE,
  type BuildAnswerCardInput,
  type TeamsCardFeatures,
} from '@omadia/channel-teams';

/** Not exported as a constant — the wire string is the contract. */
const DIRECT_LINE_TYPE = 'direct_line';
const CONSULTED_CHIP = '🔎 Konsultiert:';
const AI_CHIP = '✨ AI-generated';

type FeatureKey = keyof TeamsCardFeatures;
type Card = ReturnType<typeof buildAnswerCard>;
interface CardContent {
  readonly body: readonly unknown[];
  readonly actions?: readonly unknown[];
}

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const sampleTrace: NonNullable<BuildAnswerCardInput['runTrace']> = {
  scope: 'teams-test',
  userId: 'u1',
  startedAt: '2026-04-19T10:00:00.000Z',
  finishedAt: '2026-04-19T10:00:02.000Z',
  durationMs: 2000,
  status: 'success',
  iterations: 1,
  orchestratorToolCalls: [
    {
      callId: 'orch-1',
      toolName: 'query_knowledge_graph',
      durationMs: 30,
      isError: false,
      agentContext: 'orchestrator',
    },
  ],
  agentInvocations: [
    {
      index: 0,
      agentName: 'query_strategist',
      durationMs: 1500,
      subIterations: 1,
      status: 'success',
      toolCalls: [],
    },
  ],
};

const sampleDisclosure: NonNullable<BuildAnswerCardInput['captureDisclosure']> = {
  persisted: true,
  reasons: ['scorer-skipped:level=minimal'],
  entryType: 'memory',
  visibility: 'team',
  significance: null,
  embedded: true,
  privacyBlocksStripped: 0,
  hintTagsProcessed: 0,
};

const sampleReceipt: NonNullable<BuildAnswerCardInput['privacyReceipt']> = {
  datasetsInterned: 1,
  fieldsMasked: 2,
  fieldsCleartext: 3,
  verbsExecuted: ['sum'],
  pseudonymProjectionUsed: false,
};

const fullInput: BuildAnswerCardInput = {
  answer: 'Das ist die Antwort.',
  runTrace: sampleTrace,
  captureDisclosure: sampleDisclosure,
  privacyReceipt: sampleReceipt,
  originalUserMessage: 'Wie ist der Stand im Projekt?',
  showFreshCheck: true,
  followUpOptions: [
    { label: 'Kürzer fassen', prompt: 'Fasse das kürzer zusammen.' },
    { label: 'Als Tabelle', prompt: 'Zeig mir das als Tabelle.' },
  ],
  agentsConsulted: [{ label: 'Strategist', agentId: 'strategist', status: 'success' }],
  answerVerbosity: { effective: 'standard', source: 'configured' },
};

// ---------------------------------------------------------------------------
// Typed helpers
// ---------------------------------------------------------------------------

function contentOf(card: Card): CardContent {
  return card.content as CardContent;
}

function jsonOf(card: Card): string {
  return JSON.stringify(card.content);
}

/** Every `data.type` anywhere in the card (card actions AND body ActionSets). */
function payloadTypes(node: unknown): string[] {
  if (Array.isArray(node)) return node.flatMap((child) => payloadTypes(child));
  if (typeof node !== 'object' || node === null) return [];
  const record = node as Record<string, unknown>;
  const own: string[] = [];
  const data = record['data'];
  if (typeof data === 'object' && data !== null) {
    const type = (data as Record<string, unknown>)['type'];
    if (typeof type === 'string') own.push(type);
  }
  return [...own, ...Object.values(record).flatMap((child) => payloadTypes(child))];
}

/** One probe = one marker that must appear when the control renders. */
type Probe = (json: string, types: readonly string[]) => boolean;

const text =
  (needle: string): Probe =>
  (json) =>
    json.includes(needle);
const payload =
  (type: string): Probe =>
  (_json, types) =>
    types.includes(type);

/** Markers per switch: all present when on, all absent when off. */
const CONTROL_PROBES: Readonly<Record<FeatureKey, readonly Probe[]>> = {
  toolTrace: [text('🔧 Tool-Trace'), text('"trace-container"')],
  memoryDisclosure: [text('🧠 Memory-Auswirkung'), text('"capture-disclosure-container"')],
  privacyReceipt: [text('🛡 Privacy Guard'), text('"privacy-receipt-container"')],
  freshCheck: [payload(FRESH_CHECK_VALUE_TYPE), text('🔄 Fresh Check')],
  followUps: [text('💡 **Nächste Varianten**'), payload(FOLLOW_UP_VALUE_TYPE)],
  directLine: [payload(DIRECT_LINE_TYPE), text('💬 Direkt mit Strategist')],
  verbositySteps: [payload(VERBOSITY_STEP_VALUE_TYPE), text('🔽 Kürzer'), text('🔼 Mehr Details')],
};

/** Per switch: which probes hit (`true`) for this card. */
function probeResults(card: Card, feature: FeatureKey): boolean[] {
  const json = jsonOf(card);
  const types = payloadTypes(card.content);
  return CONTROL_PROBES[feature].map((probe) => probe(json, types));
}

function assertControlPresent(card: Card, feature: FeatureKey): void {
  const results = probeResults(card, feature);
  assert.ok(results.every(Boolean), `${feature}: expected every marker present, got ${String(results)}`);
}

function assertControlAbsent(card: Card, feature: FeatureKey): void {
  const results = probeResults(card, feature);
  assert.ok(results.every((hit) => !hit), `${feature}: expected every marker absent, got ${String(results)}`);
}

function withFeatures(overrides: Partial<TeamsCardFeatures>): TeamsCardFeatures {
  return { ...DEFAULT_TEAMS_CARD_FEATURES, ...overrides };
}

/** Defaults with exactly one switch off (local copy — defaults stay frozen). */
function switchedOff(feature: FeatureKey): TeamsCardFeatures {
  const copy: Record<FeatureKey, boolean> = { ...DEFAULT_TEAMS_CARD_FEATURES };
  copy[feature] = false;
  return copy;
}

const ALL_OFF: TeamsCardFeatures = {
  toolTrace: false,
  memoryDisclosure: false,
  privacyReceipt: false,
  freshCheck: false,
  followUps: false,
  directLine: false,
  verbositySteps: false,
};

// ---------------------------------------------------------------------------
// Pure switch resolution
// ---------------------------------------------------------------------------

describe('teamsCardFeatures — setting resolution', () => {
  it('parseBooleanSetting passes booleans, parses true/false strings, falls back on everything else', () => {
    assert.equal(parseBooleanSetting(true, false), true);
    assert.equal(parseBooleanSetting(false, true), false);
    assert.equal(parseBooleanSetting('true', false), true);
    assert.equal(parseBooleanSetting('TRUE ', false), true);
    assert.equal(parseBooleanSetting(' false', true), false);
    assert.equal(parseBooleanSetting('False', true), false);
    const notBooleans: readonly unknown[] = [undefined, null, '', '1', '0', 'yes', 1];
    for (const value of notBooleans) {
      assert.equal(parseBooleanSetting(value, true), true, `fallback true for ${String(value)}`);
      assert.equal(parseBooleanSetting(value, false), false, `fallback false for ${String(value)}`);
    }
  });

  it('resolveTeamsCardFeatures reads every manifest key, defaults missing to on, returns a fresh frozen object', () => {
    assert.deepEqual(TEAMS_CARD_FEATURE_CONFIG_KEYS, {
      toolTrace: 'teams_card_tool_trace',
      memoryDisclosure: 'teams_card_memory_disclosure',
      privacyReceipt: 'teams_card_privacy_receipt',
      freshCheck: 'teams_card_fresh_check',
      followUps: 'teams_card_follow_ups',
      directLine: 'teams_card_direct_line',
      verbositySteps: 'teams_card_verbosity_steps',
    });
    const defaultsBefore = { ...DEFAULT_TEAMS_CARD_FEATURES };
    const config: Readonly<Record<string, unknown>> = {
      teams_card_tool_trace: 'false',
      teams_card_follow_ups: false,
      teams_card_direct_line: 'yes', // not a boolean → default
    };
    const reads: string[] = [];
    const resolved = resolveTeamsCardFeatures((key) => {
      reads.push(key);
      return config[key];
    });

    assert.deepEqual(
      reads,
      TEAMS_CARD_FEATURE_KEYS.map((feature) => TEAMS_CARD_FEATURE_CONFIG_KEYS[feature]),
    );
    assert.deepEqual(resolved, withFeatures({ toolTrace: false, followUps: false }));
    assert.ok(Object.isFrozen(resolved));
    assert.notEqual(resolved, DEFAULT_TEAMS_CARD_FEATURES);
    assert.ok(Object.isFrozen(DEFAULT_TEAMS_CARD_FEATURES));
    assert.deepEqual({ ...DEFAULT_TEAMS_CARD_FEATURES }, defaultsBefore);
    assert.ok(Object.values(DEFAULT_TEAMS_CARD_FEATURES).every((on) => on === true));

    // Custom defaults: used for missing keys, never written to.
    const customDefaults: TeamsCardFeatures = withFeatures({ freshCheck: false });
    const customBefore = { ...customDefaults };
    const fromCustom = resolveTeamsCardFeatures(
      (key) => (key === 'teams_card_verbosity_steps' ? 'FALSE' : undefined),
      customDefaults,
    );
    assert.deepEqual(fromCustom, withFeatures({ freshCheck: false, verbositySteps: false }));
    assert.deepEqual(customDefaults, customBefore);
    assert.notEqual(fromCustom, customDefaults);
  });

  it('disabledTeamsCardFeatures lists the off switches in manifest order', () => {
    assert.deepEqual(TEAMS_CARD_FEATURE_KEYS, [
      'toolTrace',
      'memoryDisclosure',
      'privacyReceipt',
      'freshCheck',
      'followUps',
      'directLine',
      'verbositySteps',
    ]);
    assert.deepEqual(disabledTeamsCardFeatures(DEFAULT_TEAMS_CARD_FEATURES), []);
    assert.deepEqual(
      disabledTeamsCardFeatures(withFeatures({ verbositySteps: false, toolTrace: false, followUps: false })),
      ['toolTrace', 'followUps', 'verbositySteps'],
    );
    assert.deepEqual(disabledTeamsCardFeatures(ALL_OFF), [...TEAMS_CARD_FEATURE_KEYS]);
  });
});

// ---------------------------------------------------------------------------
// Answer card gates
// ---------------------------------------------------------------------------

describe('buildAnswerCard — operator switches', () => {
  it('renders all seven controls for a fully loaded input when features are omitted', () => {
    const card = buildAnswerCard(fullInput);
    for (const feature of TEAMS_CARD_FEATURE_KEYS) assertControlPresent(card, feature);
    assert.ok(jsonOf(card).includes(CONSULTED_CHIP));
  });

  it('is byte-identical with features omitted vs. DEFAULT_TEAMS_CARD_FEATURES', () => {
    assert.deepStrictEqual(
      buildAnswerCard(fullInput).content,
      buildAnswerCard({ ...fullInput, features: DEFAULT_TEAMS_CARD_FEATURES }).content,
    );
  });

  for (const feature of TEAMS_CARD_FEATURE_KEYS) {
    it(`${feature}: false removes only that control, the other six stay`, () => {
      const card = buildAnswerCard({ ...fullInput, features: switchedOff(feature) });
      assertControlAbsent(card, feature);
      for (const other of TEAMS_CARD_FEATURE_KEYS) {
        if (other !== feature) assertControlPresent(card, other);
      }
      // Chips are transparency markers, never switched off.
      const json = jsonOf(card);
      assert.ok(json.includes(CONSULTED_CHIP), 'Konsultiert chip must stay');
      assert.ok(json.includes(AI_CHIP), 'AI chip must stay');
    });
  }

  it('all seven off → no actions key, no hidden containers, chips still present', () => {
    const card = buildAnswerCard({ ...fullInput, features: ALL_OFF });
    const content = contentOf(card);
    const json = jsonOf(card);
    assert.equal(content.actions, undefined);
    assert.ok(!json.includes('"actions":[]'), 'no empty actions array anywhere');
    for (const id of ['trace-container', 'capture-disclosure-container', 'privacy-receipt-container']) {
      assert.ok(!json.includes(`"${id}"`), `container ${id} must be gone`);
    }
    for (const feature of TEAMS_CARD_FEATURE_KEYS) assertControlAbsent(card, feature);
    assert.deepEqual(payloadTypes(card.content), []);
    assert.ok(json.includes(AI_CHIP));
    assert.ok(json.includes(CONSULTED_CHIP));
    assert.ok(json.includes('Das ist die Antwort.'));
  });
});

// ---------------------------------------------------------------------------
// Fallback cards (long-answer path)
// ---------------------------------------------------------------------------

describe('buildFollowUpsOnlyCard — followUps switch', () => {
  const options = fullInput.followUpOptions ?? [];

  it('renders follow_up actions by default, nothing when switched off or empty', () => {
    const card = buildFollowUpsOnlyCard(options);
    assert.ok(card, 'default must render a card');
    const types = payloadTypes(card.content);
    assert.equal(types.filter((type) => type === FOLLOW_UP_VALUE_TYPE).length, options.length);

    assert.equal(buildFollowUpsOnlyCard(options, withFeatures({ followUps: false })), undefined);
    assert.equal(buildFollowUpsOnlyCard([]), undefined);
    assert.equal(buildFollowUpsOnlyCard([], DEFAULT_TEAMS_CARD_FEATURES), undefined);
  });
});

describe('buildDirectLineOnlyCard — directLine / verbositySteps switches', () => {
  const slimInput = {
    agentsConsulted: fullInput.agentsConsulted,
    originalUserMessage: fullInput.originalUserMessage,
    answerVerbosity: fullInput.answerVerbosity,
  } satisfies Parameters<typeof buildDirectLineOnlyCard>[0];

  it('respects both switches and returns undefined when nothing remains', () => {
    const byDefault = buildDirectLineOnlyCard(slimInput);
    assert.ok(byDefault, 'default must render');
    assert.ok(payloadTypes(byDefault.content).includes(DIRECT_LINE_TYPE));
    assert.ok(payloadTypes(byDefault.content).includes(VERBOSITY_STEP_VALUE_TYPE));

    const noDirectLine = buildDirectLineOnlyCard({
      ...slimInput,
      features: withFeatures({ directLine: false }),
    });
    assert.ok(noDirectLine, 'chip + size steps keep the card alive');
    assert.ok(jsonOf(noDirectLine).includes(CONSULTED_CHIP), 'Konsultiert chip must stay');
    assert.ok(!payloadTypes(noDirectLine.content).includes(DIRECT_LINE_TYPE));
    assert.ok(payloadTypes(noDirectLine.content).includes(VERBOSITY_STEP_VALUE_TYPE));

    const noSteps = buildDirectLineOnlyCard({
      ...slimInput,
      features: withFeatures({ verbositySteps: false }),
    });
    assert.ok(noSteps, 'chip + direct-line keep the card alive');
    assert.ok(!payloadTypes(noSteps.content).includes(VERBOSITY_STEP_VALUE_TYPE));
    assert.ok(payloadTypes(noSteps.content).includes(DIRECT_LINE_TYPE));

    assert.equal(
      buildDirectLineOnlyCard({
        originalUserMessage: slimInput.originalUserMessage,
        answerVerbosity: slimInput.answerVerbosity,
        features: withFeatures({ verbositySteps: false }),
      }),
      undefined,
    );
  });

  it('keeps the chip-only card for a delegated answer when direct line is off', () => {
    const card = buildDirectLineOnlyCard({
      delegatedAnswer: { label: 'Strategist', status: 'success' },
      originalUserMessage: fullInput.originalUserMessage,
      features: withFeatures({ directLine: false }),
    });
    assert.ok(card, 'the attribution chip is transparency, not a control');
    assert.ok(jsonOf(card).includes('💬 Direkte Antwort von Strategist'));
    assert.ok(!payloadTypes(card.content).includes(DIRECT_LINE_TYPE));
  });
});

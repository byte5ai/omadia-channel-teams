/**
 * Operator switches for the optional buttons on the Teams answer card.
 *
 * Every answer card may carry up to seven affordances that are NOT part of
 * the answer itself: three ToggleVisibility buttons (Tool-Trace,
 * Memory-Auswirkung, Privacy Guard), the Fresh-Check re-ask, the follow-up
 * suggestions, the Direct-Line "💬 Direkt mit <Agent>" buttons and the
 * answer-size steps ("🔽 Kürzer" / "🔼 Mehr Details"). Until now each one was
 * gated only by its data being present; this module adds the operator's
 * say via seven boolean setup fields (`teams_card_*` in manifest.yaml).
 *
 * Deliberately out of reach: the blocking cards (choice ask, topic
 * decision, Conductor approval, slot picker, calendar consent, agent-apps
 * re-check, routine smart cards) — their buttons are the only way to finish
 * a flow — and the status chips (✨ AI-generated, verifier, Originaldaten,
 * Konsultiert), which are transparency markers, not controls. The Direct-Line
 * switch therefore removes the buttons but keeps the "🔎 Konsultiert" chip.
 *
 * Side-effect-free: no ctx access, no logging — plugin.ts feeds the raw
 * config values in, tests feed fixtures. Resolved once at activate(); the
 * kernel re-activates the plugin when its config is saved.
 */

export interface TeamsCardFeatures {
  /** "🔧 Tool-Trace" ToggleVisibility button + its hidden container. */
  readonly toolTrace: boolean;
  /** "🧠 Memory-Auswirkung" ToggleVisibility button + container (capture disclosure). */
  readonly memoryDisclosure: boolean;
  /** "🛡 Privacy Guard" ToggleVisibility button + container (privacy receipt). */
  readonly privacyReceipt: boolean;
  /** "🔄 Fresh Check (ohne Memory)" re-ask button. */
  readonly freshCheck: boolean;
  /** "💡 Nächste Varianten" follow-up buttons (card + long-answer mini-card). */
  readonly followUps: boolean;
  /** "💬 Direkt mit <Agent>" buttons — the "🔎 Konsultiert" chip always stays. */
  readonly directLine: boolean;
  /** "🔽 Kürzer" / "🔼 Mehr Details" answer-size steps. */
  readonly verbositySteps: boolean;
}

export type TeamsCardFeatureKey = keyof TeamsCardFeatures;

/** Every switch on — an installation without configured values renders exactly as before. */
export const DEFAULT_TEAMS_CARD_FEATURES: TeamsCardFeatures = Object.freeze({
  toolTrace: true,
  memoryDisclosure: true,
  privacyReceipt: true,
  freshCheck: true,
  followUps: true,
  directLine: true,
  verbositySteps: true,
});

/** Feature → `setup.fields[].key` in manifest.yaml. */
export const TEAMS_CARD_FEATURE_CONFIG_KEYS: Readonly<Record<TeamsCardFeatureKey, string>> =
  Object.freeze({
    toolTrace: 'teams_card_tool_trace',
    memoryDisclosure: 'teams_card_memory_disclosure',
    privacyReceipt: 'teams_card_privacy_receipt',
    freshCheck: 'teams_card_fresh_check',
    followUps: 'teams_card_follow_ups',
    directLine: 'teams_card_direct_line',
    verbositySteps: 'teams_card_verbosity_steps',
  });

/** Stable iteration order (matches the manifest field order). */
export const TEAMS_CARD_FEATURE_KEYS: ReadonlyArray<TeamsCardFeatureKey> = Object.freeze([
  'toolTrace',
  'memoryDisclosure',
  'privacyReceipt',
  'freshCheck',
  'followUps',
  'directLine',
  'verbositySteps',
]);

/**
 * Read a boolean setup value the way the kernel hands it over: a real boolean
 * when the install registry stored it, or the strings `"true"` / `"false"`
 * (the setup form and `booleanSetupDefault` both speak that dialect).
 * Anything else — `undefined`, `null`, `""`, `"1"`, `"yes"` — is not a
 * boolean and yields the fallback, so a typo can never silently switch a
 * button off.
 */
export function parseBooleanSetting(value: unknown, fallback: boolean): boolean {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'true') return true;
    if (normalized === 'false') return false;
  }
  return fallback;
}

/**
 * Resolve the seven switches from a config reader (`ctx.config.get`).
 * Returns a fresh frozen object; never mutates the defaults or the reader.
 */
export function resolveTeamsCardFeatures(
  get: (key: string) => unknown,
  defaults: TeamsCardFeatures = DEFAULT_TEAMS_CARD_FEATURES,
): TeamsCardFeatures {
  const resolved = Object.fromEntries(
    TEAMS_CARD_FEATURE_KEYS.map((feature) => [
      feature,
      parseBooleanSetting(get(TEAMS_CARD_FEATURE_CONFIG_KEYS[feature]), defaults[feature]),
    ]),
  ) as Record<TeamsCardFeatureKey, boolean>;
  return Object.freeze({ ...resolved });
}

/** The switches that are off — for the activate() log line (names only, never values). */
export function disabledTeamsCardFeatures(
  features: TeamsCardFeatures,
): ReadonlyArray<TeamsCardFeatureKey> {
  return TEAMS_CARD_FEATURE_KEYS.filter((feature) => !features[feature]);
}

// How a message of an agent is marked as AI-generated in Teams (#643, #2043,
// rollout defect D24).
//
// Two markings exist:
//   - the PLATFORM LABEL: Teams shows "AI Generated" next to the agent's name
//     when the message carries the `AIGeneratedContent` entity;
//   - the DISCLOSURE SENTENCE: the instance's line, which the kernel folds
//     into an answer's text for channels without a label.
// A third thing looks like a label and is neither: the "✨ AI-generated" chip
// that this plugin's ANSWER CARD draws itself.
//
// WHAT WENT WRONG (D24). From 0.36.6 on, this plugin told the kernel that it
// labels natively (`supports.aiLabel`) and stripped the sentence everywhere.
// In the client, the facilitator's proactive posts then showed no marking at
// all. The "label" that had been seen on turn answers was the card's own
// chip; the platform label was never observed on any message of this plugin.
// So plain-text messages (engine posts, reminders, long answers, routine
// deliveries) went out unmarked.
//
// THE RULE NOW. The platform label counts as a marking only once somebody
// has SEEN it in a Teams client and switched `teams_native_ai_label` on.
// Until then (the default):
//   - the send providers do not declare `supports.aiLabel`, so the kernel
//     adds the disclosure sentence to an agent's own posts;
//   - the adapter keeps the sentence wherever no card chip of its own is
//     shown (long plain-text answers, routine deliveries);
//   - the answer card keeps its chip and drops the sentence, as before.
// The entity is attached in both modes, so the label can be verified live
// without another release. A message is never left without a visible mark;
// while the switch is off, a client that does render the label shows the
// label and the sentence. That is accepted: marked twice is a blemish,
// not marked at all is a disclosure failure.


/** Setup key: "the platform label was seen in our Teams client". */
export const NATIVE_AI_LABEL_CONFIG_KEY = 'teams_native_ai_label';

/** Environment flag that logs every outgoing activity as JSON. */
export const LOG_OUTGOING_ENV = 'TEAMS_LOG_OUTGOING_ACTIVITIES';

/** `true` only when the operator confirmed the platform label. */
export function resolveNativeAiLabel(get: (key: string) => unknown): boolean {
  return parseBooleanSetting(get(NATIVE_AI_LABEL_CONFIG_KEY), false);
}

export function isOutgoingLogEnabled(env: Record<string, string | undefined> = process.env): boolean {
  const value = env[LOG_OUTGOING_ENV]?.trim().toLowerCase();
  return value === '1' || value === 'true';
}

/**
 * One line per outgoing activity, as it leaves for the Bot Connector (after
 * the SDK applied the conversation reference): type, text, entities,
 * attachments, channelData, replyToId. `path` says how it was sent, so a
 * reply inside a turn and a proactive post can be compared field by field.
 * No credentials are part of an activity; it does contain message text and
 * names, which is why this is off unless the flag is set.
 */
export function logOutgoingActivity(path: 'turn' | 'proactive' | 'direct' | 'update', activity: unknown): void {
  let json: string;
  try {
    json = JSON.stringify(activity);
  } catch {
    json = '"(not serialisable)"';
  }
  console.error(`[teams] outgoing path=${path} activity=${json}`);
}

/**
 * Read a boolean setup value the way the kernel hands it over: a real boolean
 * when the install registry stored it, or the strings `"true"` / `"false"`
 * (the setup form and `booleanSetupDefault` both speak that dialect).
 * Anything else (`undefined`, `null`, `""`, `"1"`, `"yes"`) is not a boolean
 * and yields the fallback, so a typo can never silently flip a switch.
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

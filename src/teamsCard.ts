import type { Attachment } from 'botbuilder';
import { CardFactory } from 'botbuilder';
import type {
  PrivacyDetection,
  PrivacyDetectorRun,
  PrivacyReceipt,
} from '@omadia/plugin-api';
import type {
  CaptureDisclosure,
  OutgoingAttachment,
  RunTracePayload,
  VerifierBadge,
} from './kernel-types.js';
import type { MentionEntity } from './teamsMentions.js';

/**
 * Post-S+7.5 alias — the kernel/SemanticAnswer naming is "OutgoingAttachment",
 * but the Teams card code still calls the local var "DiagramAttachment"
 * throughout. Alias keeps the churn contained to this file's imports rather
 * than rewriting every call site. Same structure, 1:1 field mapping
 * (kind === 'image', url, altText, producer/cacheHit optional). Note: the
 * kernel's `diagramKind` is now encoded as `producer: 'diagram.<kind>'`.
 */
type DiagramAttachment = OutgoingAttachment;

const ADAPTIVE_CARD_CONTENT_TYPE = 'application/vnd.microsoft.card.adaptive';

/** Keep a safety margin below Teams' 28 KB hard cap for a single AC attachment. */
const CARD_BUDGET_BYTES = 26_000;
/** If the answer alone exceeds this, we drop the trace section entirely. */
const ANSWER_MAX_BYTES_WHEN_TRACED = 20_000;

const AI_LABEL_ENTITY = {
  type: 'https://schema.org/Message',
  '@type': 'Message',
  '@context': 'https://schema.org',
  additionalType: ['AIGeneratedContent'],
} as const;

/**
 * AI-Generated-Content marker attached to a Teams message activity. Teams
 * renders this as a "✨ AI generated" label under the bot's reply.
 * Returns the exact shape Teams expects in `activity.entities`.
 */
export function aiLabelEntity(): Readonly<typeof AI_LABEL_ENTITY> {
  return AI_LABEL_ENTITY;
}

export interface BuildAnswerCardInput {
  answer: string;
  runTrace?: RunTracePayload;
  /**
   * Image attachments to embed inline under the answer. Today only diagram
   * renders (render_diagram tool) populate this; the shape is generic so
   * future attachment producers can reuse the same path.
   */
  attachments?: DiagramAttachment[];
  /**
   * Answer-verifier summary. When present the card shows a small badge
   * ("✓ geprüft" / "⚠ teilweise bestätigt" / "↻ korrigiert") so the user
   * knows the reply has been re-checked against the source of truth.
   */
  verifier?: VerifierBadge;
  /**
   * When present, renders a `🔄 Fresh Check` button that re-asks the same
   * question while bypassing memory read + FTS context block. Used when
   * the user suspects an old memory entry is poisoning the answer.
   */
  originalUserMessage?: string;
  /**
   * Teams @-mention entities resolved against the active chat's roster.
   * When non-empty, embedded in the card's `msteams.entities` so Teams
   * renders the blue @-pill around matching `<at>Display Name</at>` tokens
   * inside the answer text. Empty / undefined → card renders without
   * mentions. Tokens without a matching entity are NOT stripped here —
   * the caller is expected to strip unresolved ones before calling in.
   */
  mentions?: readonly MentionEntity[];
  /**
   * Non-blocking 1-click refinement buttons emitted by `suggest_follow_ups`.
   * Rendered as `Action.Submit` entries in the card's actions array (next
   * to Tool-Trace + Fresh-Check). Up to 4 options; each carries a full
   * `prompt` that is submitted as a fresh user turn on click.
   */
  followUpOptions?: ReadonlyArray<{ label: string; prompt: string }>;
  /**
   * Palaia capture-disclosure (OB-81): summary of what the orchestrator
   * persisted into the knowledge graph for this turn. Rendered as a
   * collapsed `Container` + `Action.ToggleVisibility` button parallel to
   * the Tool-Trace section. Omitted (no button shown) when undefined.
   * Connector-side cosmetic only — the same data is also delivered via the
   * channel-agnostic `SemanticAnswer.captureDisclosure` field.
   */
  captureDisclosure?: CaptureDisclosure;
  /**
   * Privacy-Proxy receipt (Slice 6): summary of what the `privacy.redact@1`
   * provider did to the outbound LLM payload — detections, actions,
   * routing decision. Rendered as a collapsed `Container` +
   * `Action.ToggleVisibility` button parallel to Tool-Trace and
   * Memory-Auswirkung. Receipt is PII-free by contract (see
   * `@omadia/plugin-api/privacyReceipt`) so we render fields directly
   * without masking. Omitted (no button shown) when undefined.
   */
  privacyReceipt?: PrivacyReceipt;
}

/** Adaptive-Card Submit payload for the fresh-check button. */
export const FRESH_CHECK_VALUE_TYPE = 'fresh_check';

export interface FreshCheckValue {
  type: typeof FRESH_CHECK_VALUE_TYPE;
  originalMessage: string;
}

export function parseFreshCheckValue(value: unknown): FreshCheckValue | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const v = value as Record<string, unknown>;
  if (v['type'] !== FRESH_CHECK_VALUE_TYPE) return undefined;
  const originalMessage = v['originalMessage'];
  if (typeof originalMessage !== 'string' || originalMessage.length === 0) {
    return undefined;
  }
  return { type: FRESH_CHECK_VALUE_TYPE, originalMessage };
}

/** Adaptive-Card Submit payload for `ask_user_choice` option clicks. */
export const CHOICE_ASK_VALUE_TYPE = 'choice_ask';

export interface ChoiceAskValue {
  type: typeof CHOICE_ASK_VALUE_TYPE;
  /** Label text of the clicked button (for logging + user-facing echo). */
  label: string;
  /** The `value` field from the original option; this is what gets
   *  injected into the next user turn. Usually equals `label`. */
  value: string;
}

export function parseChoiceAskValue(value: unknown): ChoiceAskValue | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const v = value as Record<string, unknown>;
  if (v['type'] !== CHOICE_ASK_VALUE_TYPE) return undefined;
  const label = v['label'];
  const choiceValue = v['value'];
  if (typeof label !== 'string' || label.length === 0) return undefined;
  if (typeof choiceValue !== 'string' || choiceValue.length === 0) {
    return undefined;
  }
  return { type: CHOICE_ASK_VALUE_TYPE, label, value: choiceValue };
}

/** Adaptive-Card Submit payload for `suggest_follow_ups` button clicks. */
export const FOLLOW_UP_VALUE_TYPE = 'follow_up';

export interface FollowUpValue {
  type: typeof FOLLOW_UP_VALUE_TYPE;
  /** Full user-message sent as the next turn. */
  prompt: string;
  /** Short button label (for logging). */
  label: string;
}

export function parseFollowUpValue(value: unknown): FollowUpValue | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const v = value as Record<string, unknown>;
  if (v['type'] !== FOLLOW_UP_VALUE_TYPE) return undefined;
  const prompt = v['prompt'];
  const label = v['label'];
  if (typeof prompt !== 'string' || prompt.length === 0) return undefined;
  if (typeof label !== 'string' || label.length === 0) return undefined;
  return { type: FOLLOW_UP_VALUE_TYPE, prompt, label };
}

/** Adaptive-Card Submit payload for `find_free_slots` slot-picker clicks. */
export const BOOK_SLOT_VALUE_TYPE = 'book_slot';

export interface BookSlotValue {
  type: typeof BOOK_SLOT_VALUE_TYPE;
  /** Opaque id from `find_free_slots`. Round-tripped via `SlotCache`. */
  slotId: string;
  /** Human-readable label (for logging + user-echo). */
  label: string;
}

export function parseBookSlotValue(value: unknown): BookSlotValue | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const v = value as Record<string, unknown>;
  if (v['type'] !== BOOK_SLOT_VALUE_TYPE) return undefined;
  const slotId = v['slotId'];
  const label = v['label'];
  if (typeof slotId !== 'string' || slotId.length === 0) return undefined;
  if (typeof label !== 'string' || label.length === 0) return undefined;
  return { type: BOOK_SLOT_VALUE_TYPE, slotId, label };
}

/** Cap round-tripped option values so an abnormally long LLM response
 *  can't bloat the card payload. */
const CHOICE_VALUE_MAX = 500;
const FOLLOW_UP_PROMPT_MAX = 500;
const FOLLOW_UP_MAX_OPTIONS = 4;
const SLOT_MAX_OPTIONS = 5;

export interface BuildChoiceAskCardInput {
  question: string;
  rationale?: string;
  options: ReadonlyArray<{ label: string; value: string }>;
}

/**
 * Standalone Adaptive Card for `ask_user_choice` — blocking clarification
 * question with 2–4 option buttons. The turn has already ended by the time
 * this card is rendered; the orchestrator's `pendingUserChoice` field on
 * the `ChatTurnResult` drives the call site.
 *
 * Design mirrors `buildTopicAskCard` (also a button-choice card). The
 * first option gets `style: 'positive'` so Teams highlights the most
 * likely answer.
 */
export function buildChoiceAskCard(input: BuildChoiceAskCardInput): Attachment {
  const question =
    input.question.length > 500
      ? `${input.question.slice(0, 499)}…`
      : input.question;
  const rationale =
    input.rationale && input.rationale.length > 280
      ? `${input.rationale.slice(0, 279)}…`
      : input.rationale;

  const body: unknown[] = [
    {
      type: 'TextBlock',
      text: '🤔 Kurze Rückfrage',
      weight: 'Bolder',
      size: 'Medium',
    },
    {
      type: 'TextBlock',
      text: question,
      wrap: true,
    },
  ];
  if (rationale && rationale.trim().length > 0) {
    body.push({
      type: 'TextBlock',
      text: rationale,
      wrap: true,
      size: 'Small',
      isSubtle: true,
      spacing: 'Small',
    });
  }

  const actions = input.options.map((opt, idx) => {
    const trimmedValue =
      opt.value.length > CHOICE_VALUE_MAX
        ? opt.value.slice(0, CHOICE_VALUE_MAX)
        : opt.value;
    const submit: Record<string, unknown> = {
      type: 'Action.Submit',
      title: opt.label,
      data: {
        type: CHOICE_ASK_VALUE_TYPE,
        label: opt.label,
        value: trimmedValue,
        // See followUpAction for rationale. Same Teams quirk applies —
        // Action.Submit on a slow-responding bot shows "Something went
        // wrong" unless we flag it as messageBack.
        msteams: {
          type: 'messageBack',
          text: trimmedValue,
          displayText: opt.label,
        },
      } satisfies ChoiceAskValue & { msteams: Record<string, unknown> },
    };
    if (idx === 0) submit['style'] = 'positive';
    return submit;
  });

  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body,
    actions,
  });
}

/**
 * Minimal standalone card carrying ONLY follow-up buttons, used for the
 * long-answer fallback where the answer is sent as plain-text chunks
 * (which can't carry Adaptive-Card actions). Attached as a final activity
 * after the chunks so the buttons still appear.
 */
export function buildFollowUpsOnlyCard(
  options: ReadonlyArray<{ label: string; prompt: string }>,
): Attachment {
  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body: [
      {
        type: 'TextBlock',
        text: '💡 Nächste Varianten',
        size: 'Small',
        weight: 'Bolder',
        isSubtle: true,
      },
    ],
    actions: options.slice(0, FOLLOW_UP_MAX_OPTIONS).map((opt) => followUpAction(opt)),
  });
}

function followUpAction(opt: { label: string; prompt: string }): Record<string, unknown> {
  const prompt =
    opt.prompt.length > FOLLOW_UP_PROMPT_MAX
      ? opt.prompt.slice(0, FOLLOW_UP_PROMPT_MAX)
      : opt.prompt;
  return {
    type: 'Action.Submit',
    title: opt.label,
    tooltip: prompt,
    data: {
      type: FOLLOW_UP_VALUE_TYPE,
      label: opt.label,
      prompt,
      // Teams `messageBack` hint — tells the client: "render the label as
      // if the user typed it, then route to bot as a normal message". Two
      // effects: (1) the label appears in the chat transcript as the
      // user's own message (better context than a silent click), and (2)
      // avoids the "Something went wrong" cosmetic error that Teams shows
      // on Action.Submit inside ActionSet when the bot reply takes >5s.
      // `activity.value` still receives the full data object, so the
      // parser continues to work unchanged.
      msteams: {
        type: 'messageBack',
        text: prompt,
        displayText: opt.label,
      },
    } satisfies FollowUpValue & { msteams: Record<string, unknown> },
  };
}

/** Teams rejects card images over 1 MB; our service caps at ~900 KB. */
const MAX_ATTACHMENTS_PER_CARD = 3;

/**
 * Builds the final Teams Adaptive Card for a completed turn:
 * - AI-label container at the top
 * - Answer as markdown (truncated if too large)
 * - Collapsed `Action.ToggleVisibility` → expandable tool-trace section
 *
 * Size-aware: if the serialised card would exceed Teams' 28 KB cap we
 * progressively degrade — first drop sub-tool-call detail, then drop the
 * trace entirely, then truncate the answer. `trace` is always user-local
 * (ToggleVisibility doesn't ping the channel), so keeping it off by default
 * keeps group-chat surface area quiet.
 */
export function buildAnswerCard(input: BuildAnswerCardInput): Attachment {
  const cardBody = buildCardBody(input);
  const mentions = input.mentions ?? [];
  // Teams requires mention entities on `msteams.entities` of the card itself
  // to render the blue @-pill around `<at>Display Name</at>` tokens inside
  // the Adaptive-Card TextBlocks. Activity-level entities trigger the
  // notification / at-sign icon; the card-level copy drives the in-card
  // rendering. Both are needed for the full UX.
  const msteams: Record<string, unknown> = { width: 'Full' };
  if (mentions.length > 0) {
    msteams['entities'] = mentions;
  }
  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    body: cardBody.body,
    ...(cardBody.actions ? { actions: cardBody.actions } : {}),
    msteams,
  });
}

interface CardBody {
  body: unknown[];
  actions?: unknown[];
}

function buildCardBody(input: BuildAnswerCardInput): CardBody {
  const attachments = (input.attachments ?? []).slice(0, MAX_ATTACHMENTS_PER_CARD);
  const originalMessage = input.originalUserMessage;
  const followUps = (input.followUpOptions ?? []).slice(0, FOLLOW_UP_MAX_OPTIONS);
  const disclosure = input.captureDisclosure;
  const privacy = input.privacyReceipt;

  // Drop-order in size pressure: trace-detail → trace → disclosure → privacy
  // → truncate answer. Privacy survives longer than disclosure because it's
  // a per-turn trust signal the operator may need to refer back to ("which
  // detector ran, what did it do") — disclosure can be re-derived from the
  // memory store, privacy receipt cannot.

  // Tier 1: full answer + full trace + disclosure + privacy.
  const full = assemble(
    input.answer,
    input.runTrace,
    'full',
    attachments,
    input.verifier,
    originalMessage,
    followUps,
    disclosure,
    privacy,
  );
  if (bytes(full) <= CARD_BUDGET_BYTES) return full;

  // Tier 2: full answer + compact trace + disclosure + privacy.
  if (input.runTrace) {
    const compact = assemble(
      input.answer,
      input.runTrace,
      'compact',
      attachments,
      input.verifier,
      originalMessage,
      followUps,
      disclosure,
      privacy,
    );
    if (bytes(compact) <= CARD_BUDGET_BYTES) return compact;
  }

  // Tier 3: full answer + disclosure + privacy, no trace.
  const noTrace = assemble(
    input.answer,
    undefined,
    'full',
    attachments,
    input.verifier,
    originalMessage,
    followUps,
    disclosure,
    privacy,
  );
  if (bytes(noTrace) <= CARD_BUDGET_BYTES) return noTrace;

  // Tier 4: drop disclosure but keep privacy.
  const noDisclosure = assemble(
    input.answer,
    undefined,
    'full',
    attachments,
    input.verifier,
    originalMessage,
    followUps,
    undefined,
    privacy,
  );
  if (bytes(noDisclosure) <= CARD_BUDGET_BYTES) return noDisclosure;

  // Tier 5: drop privacy too — last collapsed section to go.
  const noPrivacy = assemble(
    input.answer,
    undefined,
    'full',
    attachments,
    input.verifier,
    originalMessage,
    followUps,
    undefined,
    undefined,
  );
  if (bytes(noPrivacy) <= CARD_BUDGET_BYTES) return noPrivacy;

  // Tier 6: truncate the answer but still keep the images — they carry the
  // actual payload the user asked for.
  const truncatedAnswer = `${input.answer.slice(0, ANSWER_MAX_BYTES_WHEN_TRACED - 1)}…`;
  return assemble(
    truncatedAnswer,
    undefined,
    'full',
    attachments,
    input.verifier,
    originalMessage,
    followUps,
    undefined,
    undefined,
  );
}

function bytes(body: CardBody): number {
  return Buffer.byteLength(JSON.stringify(body), 'utf8');
}

function assemble(
  answer: string,
  runTrace: RunTracePayload | undefined,
  traceLevel: 'full' | 'compact',
  attachments: readonly DiagramAttachment[] = [],
  verifier?: VerifierBadge,
  originalMessage?: string,
  followUps: ReadonlyArray<{ label: string; prompt: string }> = [],
  disclosure?: CaptureDisclosure,
  privacy?: PrivacyReceipt,
): CardBody {
  const chipItems: unknown[] = [
    {
      type: 'TextBlock',
      text: '✨ AI-generated',
      size: 'Small',
      weight: 'Bolder',
      color: 'Accent',
      wrap: true,
    },
  ];
  if (verifier) {
    const { label, color } = verifierChip(verifier);
    chipItems.push({
      type: 'TextBlock',
      text: label,
      size: 'Small',
      weight: 'Bolder',
      color,
      wrap: true,
      spacing: 'Small',
    });
  }
  const body: unknown[] = [
    // AI marker chip (visual, in addition to the activity-level entity).
    {
      type: 'Container',
      spacing: 'None',
      items: chipItems,
    },
    // Answer — `wrap: true` handles long lines, `rtl: false` + markdown support.
    {
      type: 'TextBlock',
      text: answer,
      wrap: true,
    },
  ];

  // Image attachments sit directly under the answer. Teams renders PNG/JPEG/GIF
  // inline up to 1024x1024 / 1 MB; SVG is not supported (service never emits).
  for (const att of attachments) {
    if (att.kind !== 'image') continue;
    body.push({
      type: 'Image',
      url: att.url,
      altText: att.altText,
      size: 'stretch',
      spacing: 'Medium',
      separator: true,
      // One-click "open full size" — Teams still does not follow redirects on
      // card images, but Action.OpenUrl opens the browser, which does.
      selectAction: { type: 'Action.OpenUrl', url: att.url },
    });
  }

  const actions: unknown[] = [];

  if (runTrace) {
    const summary = traceSummary(runTrace);
    // The trace-container starts hidden (`isVisible: false`). ToggleVisibility
    // actions are evaluated on the *viewer's* client only — expanding the
    // trace never broadcasts to the rest of the group chat.
    body.push({
      type: 'Container',
      id: 'trace-container',
      isVisible: false,
      spacing: 'Medium',
      separator: true,
      items: buildTraceItems(runTrace, traceLevel),
    });
    actions.push({
      type: 'Action.ToggleVisibility',
      title: `🔧 Tool-Trace (${summary})`,
      targetElements: ['trace-container'],
    });
  }

  if (disclosure) {
    // Same ToggleVisibility pattern as Tool-Trace — the disclosure stays
    // local to the viewer's client and is invisible to the rest of the
    // group chat until they expand it themselves.
    body.push({
      type: 'Container',
      id: 'capture-disclosure-container',
      isVisible: false,
      spacing: 'Medium',
      separator: true,
      items: buildCaptureDisclosureItems(disclosure),
    });
    actions.push({
      type: 'Action.ToggleVisibility',
      title: `🧠 Memory-Auswirkung (${captureDisclosureSummary(disclosure)})`,
      targetElements: ['capture-disclosure-container'],
    });
  }

  if (privacy) {
    // Per-turn audit row mirroring web-ui's <PrivacyReceiptCard>. Same
    // ToggleVisibility pattern as Tool-Trace + Memory-Auswirkung — local
    // to the viewer, doesn't broadcast on expand.
    body.push({
      type: 'Container',
      id: 'privacy-receipt-container',
      isVisible: false,
      spacing: 'Medium',
      separator: true,
      items: buildPrivacyReceiptItems(privacy),
    });
    actions.push({
      type: 'Action.ToggleVisibility',
      title: `🛡 Privacy Guard (${privacyReceiptSummary(privacy)})`,
      targetElements: ['privacy-receipt-container'],
    });
  }

  if (originalMessage && originalMessage.trim().length > 0) {
    // Trim to a sane length — Adaptive-Card Action.Submit `data` is
    // serialised into the card payload, and we don't want a 25 kB
    // attachment-info block inflating every reply.
    const trimmed =
      originalMessage.length > 2000
        ? `${originalMessage.slice(0, 1999)}…`
        : originalMessage;
    actions.push({
      type: 'Action.Submit',
      title: '🔄 Fresh Check (ohne Memory)',
      tooltip:
        'Stelle die gleiche Frage erneut — ohne Memory-Lese-Konvention und ohne FTS-Kontext-Block. Nützlich, wenn du vermutest, dass ein alter Memory-Eintrag die Antwort verfälscht hat.',
      data: {
        type: FRESH_CHECK_VALUE_TYPE,
        originalMessage: trimmed,
      } satisfies FreshCheckValue,
    });
  }

  // Follow-up refinements live INSIDE the body as an ActionSet, visually
  // separated from the meta-actions (Tool-Trace, Fresh-Check) in the card's
  // `actions[]`. Rationale: they are *content* — they generate the next
  // answer — while Tool-Trace + Fresh-Check are diagnostic affordances.
  // Mixing both in the same action row made it hard to spot which button
  // triggers a new turn vs. expands debug info. The separator line + header
  // makes the hierarchy explicit.
  if (followUps.length > 0) {
    body.push({
      type: 'TextBlock',
      text: '💡 **Nächste Varianten**',
      size: 'Small',
      weight: 'Bolder',
      color: 'Accent',
      spacing: 'Large',
      separator: true,
      wrap: true,
    });
    body.push({
      type: 'ActionSet',
      spacing: 'Small',
      actions: followUps.map((opt) => ({
        ...followUpAction(opt),
        style: 'positive',
      })),
    });
  }

  return actions.length > 0 ? { body, actions } : { body };
}

function verifierChip(verifier: VerifierBadge): {
  label: string;
  color: 'Good' | 'Warning' | 'Attention';
} {
  // Post-S+7.5: SDK VerifierBadge only carries `status` + optional `hint`.
  // Shadow-mode suffix + unverifiedCount details dropped — they were debug
  // tooltip text, not user-facing correctness signals. Hint is rendered if
  // set (forward-compat; the kernel doesn't populate it today).
  const hintSuffix = verifier.hint ? ` — ${verifier.hint}` : '';
  switch (verifier.status) {
    case 'verified':
      return { label: `✓ Antwort geprüft${hintSuffix}`, color: 'Good' };
    case 'corrected':
      return {
        label: `↻ Antwort korrigiert (nach Verifier-Prüfung)${hintSuffix}`,
        color: 'Warning',
      };
    case 'partial':
      return {
        label: `⚠ Teilweise bestätigt${hintSuffix}`,
        color: 'Warning',
      };
    case 'failed':
      return {
        label: `⚠ Verifier-Widerspruch${hintSuffix}`,
        color: 'Attention',
      };
  }
}

function traceSummary(trace: RunTracePayload): string {
  const totalTools =
    trace.orchestratorToolCalls.length +
    trace.agentInvocations.reduce((n, inv) => n + inv.toolCalls.length, 0);
  return `${String(trace.agentInvocations.length)} Agent${trace.agentInvocations.length === 1 ? '' : 'en'}, ${String(totalTools)} Call${totalTools === 1 ? '' : 's'}, ${formatMs(trace.durationMs)}`;
}

function buildTraceItems(
  trace: RunTracePayload,
  level: 'full' | 'compact',
): unknown[] {
  const items: unknown[] = [];

  items.push({
    type: 'FactSet',
    facts: [
      { title: 'Status', value: trace.status === 'success' ? '✓ OK' : '✗ FAIL' },
      { title: 'Dauer', value: formatMs(trace.durationMs) },
      { title: 'Iterationen', value: String(trace.iterations) },
    ],
  });

  if (trace.orchestratorToolCalls.length > 0) {
    items.push({
      type: 'TextBlock',
      text: '**Orchestrator-Tools**',
      weight: 'Bolder',
      size: 'Small',
      spacing: 'Medium',
      wrap: true,
    });
    for (const tc of trace.orchestratorToolCalls) {
      items.push(toolRow(tc.toolName, tc.durationMs, tc.isError));
    }
  }

  for (const inv of trace.agentInvocations) {
    const statusIcon = inv.status === 'success' ? '✓' : '✗';
    items.push({
      type: 'TextBlock',
      text: `**${statusIcon} 🤖 ${inv.agentName}** · ${formatMs(inv.durationMs)} · ${String(inv.subIterations)} iter · ${String(inv.toolCalls.length)} call${inv.toolCalls.length === 1 ? '' : 's'}`,
      spacing: 'Medium',
      wrap: true,
    });
    if (level === 'full') {
      for (const tc of inv.toolCalls) {
        items.push(toolRow(tc.toolName, tc.durationMs, tc.isError, true));
      }
    }
  }

  return items;
}

/**
 * Header summary for the Memory-Auswirkung toggle button.
 * Examples: "persisted · memory · 0.74", "stripped 1× private",
 * "verworfen · score 0.08".
 */
function captureDisclosureSummary(d: CaptureDisclosure): string {
  if (!d.persisted) {
    if (d.significance !== null) {
      return `verworfen · score ${d.significance.toFixed(2)}`;
    }
    return 'verworfen';
  }
  const parts: string[] = ['persisted'];
  if (d.entryType) parts.push(d.entryType);
  if (d.significance !== null) parts.push(d.significance.toFixed(2));
  if (d.privacyBlocksStripped > 0) {
    parts.push(`–${String(d.privacyBlocksStripped)}×privat`);
  }
  return parts.join(' · ');
}

function buildCaptureDisclosureItems(d: CaptureDisclosure): unknown[] {
  const items: unknown[] = [];

  items.push({
    type: 'TextBlock',
    text: '**🧠 Was landet in Palaia?**',
    weight: 'Bolder',
    size: 'Small',
    wrap: true,
  });

  const facts: Array<{ title: string; value: string }> = [
    {
      title: 'Persistiert',
      value: d.persisted ? '✓ ja' : '✗ verworfen',
    },
  ];
  if (d.entryType) {
    facts.push({ title: 'Eintrag-Typ', value: d.entryType });
  }
  if (d.visibility) {
    facts.push({ title: 'Sichtbarkeit', value: d.visibility });
  }
  if (d.significance !== null) {
    facts.push({ title: 'Significance', value: d.significance.toFixed(2) });
  }
  facts.push({
    title: 'Embedding',
    value: d.embedded ? '✓ vektorisiert' : '— kein Vektor',
  });
  if (d.privacyBlocksStripped > 0) {
    facts.push({
      title: '<private>-Blöcke entfernt',
      value: String(d.privacyBlocksStripped),
    });
  }
  if (d.hintTagsProcessed > 0) {
    facts.push({
      title: '<palaia-hint>-Tags',
      value: String(d.hintTagsProcessed),
    });
  }
  items.push({ type: 'FactSet', facts });

  if (d.reasons.length > 0) {
    items.push({
      type: 'TextBlock',
      text: '**Begründung**',
      weight: 'Bolder',
      size: 'Small',
      spacing: 'Medium',
      wrap: true,
    });
    for (const reason of d.reasons) {
      items.push({
        type: 'TextBlock',
        text: `• \`${reason}\``,
        size: 'Small',
        spacing: 'Small',
        wrap: true,
      });
    }
  }

  if (d.graphRefs && d.graphRefs.entityNodeIds.length > 0) {
    items.push({
      type: 'TextBlock',
      text: `**Verknüpfte Entitäten** (${String(d.graphRefs.entityNodeIds.length)})`,
      weight: 'Bolder',
      size: 'Small',
      spacing: 'Medium',
      wrap: true,
    });
    for (const id of d.graphRefs.entityNodeIds.slice(0, 8)) {
      items.push({
        type: 'TextBlock',
        text: `• \`${id}\``,
        size: 'Small',
        spacing: 'Small',
        wrap: true,
      });
    }
    if (d.graphRefs.entityNodeIds.length > 8) {
      items.push({
        type: 'TextBlock',
        text: `… +${String(d.graphRefs.entityNodeIds.length - 8)} weitere`,
        size: 'Small',
        isSubtle: true,
        spacing: 'Small',
        wrap: true,
      });
    }
  }

  return items;
}

/**
 * Header summary for the Privacy-Guard toggle button.
 * Examples: "0 Erkennungen · public-llm", "3 Erkennungen · public-llm",
 * "blocked: customer_data", "DEBUG · 2 Erkennungen".
 */
function privacyReceiptSummary(r: PrivacyReceipt): string {
  const totalHits = r.detections.reduce((n, d) => n + d.count, 0);
  const debugPrefix = r.debug ? 'DEBUG · ' : '';
  const routing =
    r.routing === 'blocked'
      ? `blockiert${r.routingReason ? ': ' + r.routingReason : ''}`
      : r.routing;
  const hits = `${String(totalHits)} Erkennung${totalHits === 1 ? '' : 'en'}`;
  return `${debugPrefix}${hits} · ${routing}`;
}

function buildPrivacyReceiptItems(r: PrivacyReceipt): unknown[] {
  const items: unknown[] = [];

  items.push({
    type: 'TextBlock',
    text: '**🛡 Privacy-Guard · was wurde vor dem LLM-Call gefiltert?**',
    weight: 'Bolder',
    size: 'Small',
    wrap: true,
  });

  if (r.debug) {
    items.push({
      type: 'TextBlock',
      text: '⚠ DEBUG-MODUS aktiv — Receipt enthält PII-Werte (operator-toggle, niemals in Prod).',
      color: 'Warning',
      weight: 'Bolder',
      size: 'Small',
      wrap: true,
      spacing: 'Small',
    });
  }

  const facts: Array<{ title: string; value: string }> = [
    { title: 'Modus', value: renderPolicyMode(r.policyMode) },
    { title: 'Routing', value: renderRouting(r.routing) },
  ];
  if (r.routingReason) {
    facts.push({ title: 'Grund', value: r.routingReason });
  }
  facts.push({ title: 'Latenz', value: `${String(r.latencyMs)} ms` });
  facts.push({
    title: 'Audit-ID',
    value: `${r.receiptId} (${r.auditHash.slice(0, 8)}…)`,
  });
  items.push({ type: 'FactSet', facts });

  if (r.detections.length === 0) {
    items.push({
      type: 'TextBlock',
      text: '_Keine Erkennungen._',
      size: 'Small',
      isSubtle: true,
      wrap: true,
      spacing: 'Medium',
    });
  } else {
    items.push({
      type: 'TextBlock',
      text: '**Erkennungen**',
      weight: 'Bolder',
      size: 'Small',
      spacing: 'Medium',
      wrap: true,
    });
    for (const d of r.detections) {
      items.push(privacyDetectionRow(d));
    }
  }

  // Detector run statuses surface fail-open cases (skipped/timeout/error)
  // that would otherwise look identical to "0 Erkennungen" — mirrors web-ui.
  const nonOk = r.detectorRuns.filter((run) => run.status !== 'ok');
  if (nonOk.length > 0) {
    items.push({
      type: 'TextBlock',
      text: '**Detector-Status**',
      weight: 'Bolder',
      size: 'Small',
      spacing: 'Medium',
      wrap: true,
    });
    for (const run of nonOk) {
      items.push(privacyDetectorRunRow(run));
    }
  }

  return items;
}

function renderPolicyMode(m: PrivacyReceipt['policyMode']): string {
  return m === 'data-residency' ? 'Data-Residency' : 'PII-Shield';
}

function renderRouting(r: PrivacyReceipt['routing']): string {
  switch (r) {
    case 'public-llm':
      return 'Public LLM';
    case 'local-llm':
      return 'Local LLM (Ollama)';
    case 'blocked':
      return 'Blockiert';
  }
}

function renderDetectionAction(a: PrivacyDetection['action']): string {
  switch (a) {
    case 'redacted':
      return 'redigiert';
    case 'tokenized':
      return 'tokenisiert';
    case 'blocked':
      return 'blockiert';
    case 'passed':
      return 'durchgelassen';
  }
}

function privacyDetectionRow(d: PrivacyDetection): unknown {
  const confidence = `${(d.confidenceMin * 100).toFixed(0)}%`;
  const valuesPart =
    d.values && d.values.length > 0
      ? ` · Werte: ${d.values.slice(0, 3).join(', ')}${d.values.length > 3 ? '…' : ''}`
      : '';
  return {
    type: 'TextBlock',
    text: `• **${d.type}** × ${String(d.count)} → ${renderDetectionAction(d.action)} · ${d.detector} · min. Confidence ${confidence}${valuesPart}`,
    size: 'Small',
    wrap: true,
    spacing: 'Small',
  };
}

function privacyDetectorRunRow(run: PrivacyDetectorRun): unknown {
  const icon =
    run.status === 'timeout' ? '⏱' : run.status === 'error' ? '✗' : '⊘';
  const reason = run.reason ? ` · ${run.reason}` : '';
  return {
    type: 'TextBlock',
    text: `${icon} **${run.detector}** · ${run.status}${reason} · ${String(run.latencyMs)} ms`,
    size: 'Small',
    wrap: true,
    spacing: 'Small',
    color: run.status === 'error' ? 'Attention' : 'Warning',
  };
}

function toolRow(
  name: string,
  durationMs: number,
  isError: boolean,
  indented = false,
): unknown {
  const icon = isError ? '⚠' : '🔧';
  const prefix = indented ? '    ' : '';
  return {
    type: 'TextBlock',
    text: `${prefix}${icon} \`${name}\` · ${formatMs(durationMs)}`,
    size: 'Small',
    color: isError ? 'Attention' : 'Default',
    spacing: 'Small',
    wrap: true,
  };
}

function formatMs(ms: number): string {
  if (ms < 1000) return `${String(ms)}ms`;
  return `${(ms / 1000).toFixed(1)}s`;
}

export const TEAMS_CARD_CONTENT_TYPE = ADAPTIVE_CARD_CONTENT_TYPE;

/**
 * Adaptive card with two explicit action buttons, used when the topic-detector
 * can't decide whether the latest user message is a follow-up or a new topic.
 * The button clicks arrive back on the `/api/messages` endpoint as regular
 * message activities with `value` set to one of the payloads below, which the
 * TeamsBot handles by resolving the pending state and continuing the turn.
 */
export interface TopicAskCardInput {
  /** Short summary of the previous exchange, shown in the card body. */
  previousSummary: string;
  /** The original user message we're clarifying on. Round-tripped via the
   *  button value so we don't need any additional server state. */
  originalMessage: string;
}

export const TOPIC_DECISION_VALUE_TYPE = 'topic_decision';
export type TopicDecisionChoice = 'continue' | 'reset';

export interface TopicDecisionValue {
  type: typeof TOPIC_DECISION_VALUE_TYPE;
  choice: TopicDecisionChoice;
  originalMessage: string;
}

export function buildTopicAskCard(input: TopicAskCardInput): Attachment {
  const summary = input.previousSummary.length > 240
    ? `${input.previousSummary.slice(0, 239)}…`
    : input.previousSummary;
  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body: [
      {
        type: 'TextBlock',
        text: '🤔 Kurze Rückfrage',
        weight: 'Bolder',
        size: 'Medium',
      },
      {
        type: 'TextBlock',
        text: `Deine letzte Nachricht könnte eine Follow-up auf unseren vorherigen Chat sein oder ein neues Thema. Wie soll ich sie verstehen?`,
        wrap: true,
      },
      {
        type: 'Container',
        style: 'emphasis',
        items: [
          {
            type: 'TextBlock',
            text: '**Letzter Chat-Stand:**',
            size: 'Small',
            weight: 'Bolder',
            wrap: true,
          },
          {
            type: 'TextBlock',
            text: summary,
            size: 'Small',
            wrap: true,
            spacing: 'Small',
          },
        ],
      },
    ],
    actions: [
      {
        type: 'Action.Submit',
        title: '↪︎ Folge-Frage',
        style: 'positive',
        data: {
          type: TOPIC_DECISION_VALUE_TYPE,
          choice: 'continue',
          originalMessage: input.originalMessage,
        } satisfies TopicDecisionValue,
      },
      {
        type: 'Action.Submit',
        title: '✨ Neues Thema',
        data: {
          type: TOPIC_DECISION_VALUE_TYPE,
          choice: 'reset',
          originalMessage: input.originalMessage,
        } satisfies TopicDecisionValue,
      },
    ],
  });
}

/**
 * Runtime guard for activity.value payloads. Teams stringifies the submit
 * `data` for us, but we can't assume the shape — a user could also hand-craft
 * a JSON message.
 */
export function parseTopicDecisionValue(
  value: unknown,
): TopicDecisionValue | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const v = value as Record<string, unknown>;
  if (v['type'] !== TOPIC_DECISION_VALUE_TYPE) return undefined;
  const choice = v['choice'];
  if (choice !== 'continue' && choice !== 'reset') return undefined;
  const originalMessage = v['originalMessage'];
  if (typeof originalMessage !== 'string' || originalMessage.length === 0) {
    return undefined;
  }
  return { type: TOPIC_DECISION_VALUE_TYPE, choice, originalMessage };
}

export interface BuildSlotPickerCardInput {
  question: string;
  slots: ReadonlyArray<{
    slotId: string;
    label: string;
    confidence: number;
  }>;
}

/**
 * Sidecar card rendered after the answer when `find_free_slots` scheduled
 * a slot-picker. Up to 5 clickable slots; clicks arrive back as normal
 * message activities with `value = BookSlotValue`. The first slot (highest
 * Graph confidence) gets `style: 'positive'` so Teams visually flags it.
 */
export function buildSlotPickerCard(input: BuildSlotPickerCardInput): Attachment {
  const slots = input.slots.slice(0, SLOT_MAX_OPTIONS);
  const body: unknown[] = [
    {
      type: 'TextBlock',
      text: '📅 Freie Termine',
      weight: 'Bolder',
      size: 'Small',
      isSubtle: true,
    },
    {
      type: 'TextBlock',
      text: input.question,
      wrap: true,
      size: 'Small',
    },
  ];

  const actions = slots.map((slot, idx) => {
    const action: Record<string, unknown> = {
      type: 'Action.Submit',
      title: slot.label,
      tooltip: `Confidence: ${String(slot.confidence)}%`,
      data: {
        type: BOOK_SLOT_VALUE_TYPE,
        slotId: slot.slotId,
        label: slot.label,
        msteams: {
          type: 'messageBack',
          text: `Bitte Termin buchen: ${slot.label}`,
          displayText: `📅 ${slot.label}`,
        },
      } satisfies BookSlotValue & { msteams: Record<string, unknown> },
    };
    if (idx === 0) action['style'] = 'positive';
    return action;
  });

  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body,
    actions,
  });
}

// -----------------------------------------------------------------------------
// Routine smart-card Submit values
// -----------------------------------------------------------------------------

/**
 * Submit-value shape for the Pause / Löschen buttons on the routine
 * smart-card the bot delivers when a cron-triggered routine fires. Mirror
 * of `RoutineCardActionPayload` in the routines plugin — kept duplicated
 * here so the Teams package stays free of cross-package source imports
 * into `middleware/src`. The shape is stable and validated structurally.
 */
export const ROUTINE_ACTION_VALUE_KIND = 'routine.action';
export const ROUTINE_LIST_FILTER_VALUE_KIND = 'routine.list.filter';

export interface RoutineActionValue {
  kind: typeof ROUTINE_ACTION_VALUE_KIND;
  action: 'pause' | 'resume' | 'trigger_now' | 'delete';
  id: string;
}

export interface RoutineListFilterValue {
  kind: typeof ROUTINE_LIST_FILTER_VALUE_KIND;
  filter: 'all' | 'active' | 'paused';
}

export function parseRoutineCardActionValue(
  value: unknown,
): RoutineActionValue | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const v = value as Record<string, unknown>;
  if (v['kind'] !== ROUTINE_ACTION_VALUE_KIND) return undefined;
  const action = v['action'];
  if (
    action !== 'pause' &&
    action !== 'resume' &&
    action !== 'trigger_now' &&
    action !== 'delete'
  ) {
    return undefined;
  }
  const id = v['id'];
  if (typeof id !== 'string' || id.length === 0) return undefined;
  return { kind: ROUTINE_ACTION_VALUE_KIND, action, id };
}

export function parseRoutineListFilterValue(
  value: unknown,
): RoutineListFilterValue | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const v = value as Record<string, unknown>;
  if (v['kind'] !== ROUTINE_LIST_FILTER_VALUE_KIND) return undefined;
  const filter = v['filter'];
  if (filter !== 'all' && filter !== 'active' && filter !== 'paused') {
    return undefined;
  }
  return { kind: ROUTINE_LIST_FILTER_VALUE_KIND, filter };
}

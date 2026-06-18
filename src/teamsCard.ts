import type { Attachment } from 'botbuilder';
import { CardFactory } from 'botbuilder';
import type { PrivacyReceipt } from '@omadia/plugin-api';
import type {
  AgentConsultation,
  CaptureDisclosure,
  DelegatedAnswer,
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
  /**
   * Privacy Shield v4 — real values rendered into `answer` that the LLM
   * never saw (resolved server-side behind the data-plane boundary). Their
   * occurrences in the answer are highlighted in the card's accent colour
   * (Teams renders accent as its brand violet) so the asker sees at a
   * glance which data was protected. When present, the answer is rendered
   * as a `RichTextBlock` of `TextRun`s instead of a markdown `TextBlock`.
   */
  maskedValues?: readonly string[];
  /**
   * #332 Layer 1 — tamper-evident agent transparency. The harness-built list
   * of sub-agents actually invoked this turn (from the deterministic run
   * trace, NOT the LLM's prose). Rendered as a "🔎 Konsultiert: …" footer chip
   * and as one dynamic "💬 Direkt mit <Agent>" button per consulted agent.
   * Omitted → no chip, no buttons.
   */
  agentsConsulted?: readonly AgentConsultation[];
  /**
   * #332 Layer 2 — Direct Line. Present when this turn's answer is a named
   * specialist's verbatim reply, delivered by the harness. Renders an
   * attribution chip ("💬 Direkte Antwort von <Agent>") above the answer so the
   * user knows the body is the specialist's own words, not the orchestrator's.
   */
  delegatedAnswer?: DelegatedAnswer;
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

/**
 * #332 Layer 2 — Adaptive-Card Submit payload for a dynamic "💬 Direkt mit
 * <Agent>" button. Carries the directive `token` (resolved server-side against
 * the orchestrator's whitelisted sub-agents) plus the original user message, so
 * the click re-issues the question as `#<token> <message>` — a harness-routed
 * Direct Line, no typing required. The token is derived per agent at render
 * time (NEVER hardcoded).
 */
export const DIRECT_LINE_VALUE_TYPE = 'direct_line';

export interface DirectLineValue {
  type: typeof DIRECT_LINE_VALUE_TYPE;
  /** Space-free directive token derived from the agent label/id. */
  token: string;
  /** The question to re-route verbatim to the specialist. */
  originalMessage: string;
}

export function parseDirectLineValue(value: unknown): DirectLineValue | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const v = value as Record<string, unknown>;
  if (v['type'] !== DIRECT_LINE_VALUE_TYPE) return undefined;
  const token = v['token'];
  const originalMessage = v['originalMessage'];
  if (typeof token !== 'string' || token.length === 0) return undefined;
  if (typeof originalMessage !== 'string' || originalMessage.length === 0) {
    return undefined;
  }
  return { type: DIRECT_LINE_VALUE_TYPE, token, originalMessage };
}

/**
 * Derive a space-free directive token for an agent. Prefers the stable agent
 * id's last segment, falls back to the label; strips everything but
 * alphanumerics so it survives the channel mention/markup + the core parser's
 * `#<token>` grammar (and the core resolver normalizes the same way).
 */
export function directLineToken(agent: {
  agentId?: string;
  label: string;
}): string {
  const base = agent.agentId
    ? (agent.agentId.split(/[./]/).pop() ?? agent.agentId)
    : agent.label;
  return base.replace(/[^A-Za-z0-9]/g, '');
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

/**
 * #332 — slim card carrying ONLY the agent-transparency chips + Direct-Line
 * buttons, for the long-answer fallback (answer > Teams' single-message cap)
 * where the full answer card is replaced by plain text chunks. Without this,
 * the transparency footer and Direct-Line affordance would silently vanish on
 * long answers. Returns undefined when there is nothing to surface.
 */
export function buildDirectLineOnlyCard(
  input: Pick<
    BuildAnswerCardInput,
    'agentsConsulted' | 'delegatedAnswer' | 'originalUserMessage'
  >,
): Attachment | undefined {
  const hasConsulted = (input.agentsConsulted?.length ?? 0) > 0;
  if (!hasConsulted && !input.delegatedAnswer) return undefined;
  const base: CardBody = {
    body: [{ type: 'Container', spacing: 'None', items: [] }],
    actions: [],
  };
  const decorated = decorateDirectLine(base, input as BuildAnswerCardInput);
  const head = decorated.body[0] as { items?: unknown[] } | undefined;
  const hasChips = (head?.items?.length ?? 0) > 0;
  const hasActions = (decorated.actions?.length ?? 0) > 0;
  if (!hasChips && !hasActions) return undefined;
  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    body: decorated.body,
    ...(hasActions ? { actions: decorated.actions } : {}),
    msteams: { width: 'Full' },
  });
}

function buildCardBody(input: BuildAnswerCardInput): CardBody {
  // Size-tier selection stays in `buildCardBodyBase`; the #332 transparency
  // chips + Direct-Line buttons are decorated onto the chosen tier here.
  const card = decorateDirectLine(buildCardBodyBase(input), input);
  // The decoration runs AFTER the base fitter, so on an answer fitted near the
  // budget the added buttons could push the card over Teams' hard cap → silent
  // render failure. The chips are tiny (keep them); the buttons carry the bulk,
  // so shed them if we're over budget. The Direct-Line buttons also surface in
  // the long-answer fallback's slim card, so dropping them here is not a total
  // loss of the affordance.
  if (bytes(card) > CARD_BUDGET_BYTES && card.actions) {
    card.actions = card.actions.filter((a) => !isDirectLineAction(a));
    if (card.actions.length === 0) delete card.actions;
  }
  return card;
}

/** True for an `Action.Submit` carrying a Direct-Line button payload. */
function isDirectLineAction(action: unknown): boolean {
  const data = (action as { data?: { type?: unknown } } | undefined)?.data;
  return data?.type === DIRECT_LINE_VALUE_TYPE;
}

/**
 * #332 — add the tamper-evident "🔎 Konsultiert" footer chip, the Direct-Line
 * attribution chip, and one dynamic "💬 Direkt mit <Agent>" button per consulted
 * agent. All sourced from harness-built fields (`agentsConsulted` /
 * `delegatedAnswer`), never hardcoded. Mutates a copy of the assembled card.
 */
function decorateDirectLine(
  card: CardBody,
  input: BuildAnswerCardInput,
): CardBody {
  const consulted = input.agentsConsulted ?? [];
  const delegated = input.delegatedAnswer;
  if (consulted.length === 0 && !delegated) return card;

  // The chip Container is body[0] in every `assemble` tier; push there so the
  // chips sit next to the "✨ AI-generated" / verifier chips. Guard defensively.
  const head = card.body[0] as { type?: string; items?: unknown[] } | undefined;
  const chipItems =
    head?.type === 'Container' && Array.isArray(head.items)
      ? head.items
      : undefined;

  if (delegated && chipItems) {
    chipItems.push({
      type: 'TextBlock',
      text: `💬 Direkte Antwort von ${delegated.label}${
        delegated.status === 'error' ? ' (Fehler)' : ''
      }`,
      size: 'Small',
      weight: 'Bolder',
      color: delegated.status === 'error' ? 'Attention' : 'Good',
      wrap: true,
      spacing: 'Small',
    });
  }

  if (consulted.length > 0 && chipItems) {
    chipItems.push({
      type: 'TextBlock',
      text: `🔎 Konsultiert: ${consulted
        .map(
          (c) =>
            `${c.label} ${c.status === 'success' ? '✓' : '✗'}${
              c.toolCalls ? ` · ${String(c.toolCalls)}` : ''
            }`,
        )
        .join(' · ')}`,
      size: 'Small',
      isSubtle: true,
      wrap: true,
      spacing: 'Small',
    });
  }

  // One dynamic Direct-Line button per consulted agent — re-issues the user's
  // question routed straight to that specialist. Deduped by token; capped so a
  // many-agent turn doesn't flood the action row.
  // Strip any leading `#<token> ` directive so a re-click doesn't compound the
  // prefix (`#strategist #strategist …`) and leak stale routing syntax into the
  // specialist's payload. Cap short (600) — four buttons each embed this, so a
  // 2 KB message would be the dominant card-size cost.
  const original = input.originalUserMessage?.replace(/^#[A-Za-z0-9._-]+\s+/, '');
  if (consulted.length > 0 && original && original.trim().length > 0) {
    const trimmed =
      original.length > 600 ? `${original.slice(0, 599)}…` : original;
    const actions = card.actions ?? [];
    const seen = new Set<string>();
    for (const agent of consulted) {
      const token = directLineToken(agent);
      if (token.length === 0 || seen.has(token)) continue;
      seen.add(token);
      actions.push({
        type: 'Action.Submit',
        title: `💬 Direkt mit ${agent.label}`,
        tooltip: `Stelle dieselbe Frage erneut — direkt an ${agent.label}. Die Antwort kommt wortwörtlich vom Spezialisten, ohne dass der Orchestrator sie umformulieren oder unterdrücken kann.`,
        data: {
          type: DIRECT_LINE_VALUE_TYPE,
          token,
          originalMessage: trimmed,
        } satisfies DirectLineValue,
      });
      if (seen.size >= 4) break;
    }
    card.actions = actions;
  }

  return card;
}

function buildCardBodyBase(input: BuildAnswerCardInput): CardBody {
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
    input.maskedValues,
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
      input.maskedValues,
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
    input.maskedValues,
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
    input.maskedValues,
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
    input.maskedValues,
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
    input.maskedValues,
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
  maskedValues: readonly string[] = [],
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
    // Answer. A v4 table renders as a native Adaptive-Card `Table` with
    // masked cells in the accent (violet) colour; prose / non-table answers
    // stay a markdown `TextBlock` (plus a compact violet line when they
    // resolved masked values).
    ...buildAnswerElements(answer, maskedValues),
  ];

  // Attachments sit directly under the answer.
  //  - image: rendered inline (Teams renders PNG/JPEG/GIF up to 1024x1024 / 1 MB;
  //    SVG is not supported, the service never emits it).
  //  - file (e.g. create_xlsx .xlsx, create_docx .docx): Teams has no inline
  //    element for an arbitrary blob, so we surface a one-click download button.
  //    Action.OpenUrl opens the signed /documents URL in the browser, which
  //    streams the file. (Native upload into the channel's Files tab would need
  //    a Graph file-consent flow — deferred.)
  for (const att of attachments) {
    if (att.kind === 'image') {
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
      continue;
    }
    if (att.kind === 'file') {
      const sizeHint =
        typeof att.sizeBytes === 'number' && att.sizeBytes > 0
          ? ` (${String(Math.max(1, Math.round(att.sizeBytes / 1024)))} KB)`
          : '';
      body.push({
        type: 'ActionSet',
        spacing: 'Medium',
        separator: true,
        actions: [
          {
            type: 'Action.OpenUrl',
            title: `📥 ${att.altText}${sizeHint}`,
            url: att.url,
          },
        ],
      });
      continue;
    }
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

/**
 * Build the answer body element(s). A v4 table answer is rendered as a
 * native Adaptive-Card `Table` so it displays as a real table on Teams (a
 * markdown `TextBlock` cannot render tables at all) — masked cells get the
 * accent colour, which Teams renders in its brand violet. Prose and
 * non-table answers stay a markdown `TextBlock`; when such an answer
 * resolved masked values, a compact violet line is appended beneath it.
 */
function buildAnswerElements(
  answer: string,
  maskedValues: readonly string[],
): unknown[] {
  const masked = new Set(maskedValues.filter((v) => v.trim().length > 0));
  const table = extractMarkdownTable(answer);
  if (table !== undefined) {
    const out: unknown[] = [];
    if (table.prose.length > 0) {
      out.push({ type: 'TextBlock', text: table.prose, wrap: true });
    }
    out.push(buildTableElement(table, masked));
    if (table.trailing.length > 0) {
      out.push({
        type: 'TextBlock',
        text: table.trailing,
        wrap: true,
        spacing: 'Small',
      });
    }
    return out;
  }
  const out: unknown[] = [{ type: 'TextBlock', text: answer, wrap: true }];
  if (masked.size > 0) {
    out.push(maskedValuesLine(masked));
  }
  return out;
}

/** A parsed GFM table — the shape the v4 Materializer emits. */
interface ParsedAnswerTable {
  readonly prose: string;
  readonly trailing: string;
  readonly headers: readonly string[];
  readonly rows: ReadonlyArray<readonly string[]>;
}

/**
 * Extract the first GFM table from a v4 materialized answer. Returns
 * `undefined` when there is no table — the caller then keeps a plain
 * markdown `TextBlock`.
 */
function extractMarkdownTable(answer: string): ParsedAnswerTable | undefined {
  const lines = answer.split('\n');
  for (let i = 0; i + 1 < lines.length; i += 1) {
    const headerLine = lines[i];
    const sepLine = lines[i + 1];
    if (headerLine === undefined || sepLine === undefined) continue;
    if (!headerLine.includes('|') || !sepLine.includes('|')) continue;
    const headers = parseTableRow(headerLine);
    const sepCells = parseTableRow(sepLine);
    if (headers.length === 0 || sepCells.length !== headers.length) continue;
    if (!sepCells.every((c) => /^:?-+:?$/.test(c))) continue;
    const rows: string[][] = [];
    let j = i + 2;
    for (; j < lines.length; j += 1) {
      const line = lines[j];
      if (line === undefined || line.trim() === '' || !line.includes('|')) {
        break;
      }
      rows.push(parseTableRow(line));
    }
    return {
      prose: lines.slice(0, i).join('\n').trim(),
      trailing: lines.slice(j).join('\n').trim(),
      headers,
      rows,
    };
  }
  return undefined;
}

/**
 * Split one pipe-delimited row into trimmed cell values — dropping the
 * empty fields the bounding pipes produce and un-escaping `\|` (the
 * Materializer escapes literal pipes inside a cell).
 */
function parseTableRow(line: string): string[] {
  const parts = line.split(/(?<!\\)\|/);
  if (parts.length > 0 && (parts[0] ?? '').trim() === '') parts.shift();
  if (parts.length > 0 && (parts[parts.length - 1] ?? '').trim() === '') {
    parts.pop();
  }
  return parts.map((c) => c.trim().replace(/\\\|/g, '|'));
}

/** Render a parsed table as a native Adaptive-Card `Table` element. */
function buildTableElement(
  table: ParsedAnswerTable,
  masked: ReadonlySet<string>,
): unknown {
  const headerRow = {
    type: 'TableRow',
    cells: table.headers.map((h) => tableCell(h, 'header')),
  };
  const dataRows = table.rows.map((row) => ({
    type: 'TableRow',
    cells: table.headers.map((_, ci) => {
      const value = row[ci] ?? '';
      return tableCell(value, masked.has(value) ? 'masked' : 'plain');
    }),
  }));
  return {
    type: 'Table',
    columns: table.headers.map(() => ({ width: 1 })),
    firstRowAsHeaders: true,
    gridStyle: 'default',
    rows: [headerRow, ...dataRows],
  };
}

/** One Adaptive-Card `TableCell`. `masked` cells render in accent/violet. */
function tableCell(
  text: string,
  kind: 'header' | 'masked' | 'plain',
): unknown {
  const block: Record<string, unknown> = {
    type: 'TextBlock',
    text,
    wrap: true,
    size: 'Small',
  };
  if (kind === 'header') {
    block['weight'] = 'Bolder';
    block['isSubtle'] = true;
  } else if (kind === 'masked') {
    // Server-resolved value the LLM never saw — Teams renders Accent violet.
    block['weight'] = 'Bolder';
    block['color'] = 'Accent';
  }
  return { type: 'TableCell', items: [block] };
}

/**
 * A compact violet line listing the masked values — appended under a
 * non-table answer so the asker still sees what the boundary resolved.
 */
function maskedValuesLine(masked: ReadonlySet<string>): unknown {
  const inlines: unknown[] = [
    { type: 'TextRun', text: '🛡 Server-seitig eingesetzt: ' },
  ];
  const values = [...masked];
  values.forEach((value, idx) => {
    inlines.push({
      type: 'TextRun',
      text: value,
      color: 'Accent',
      weight: 'Bolder',
    });
    if (idx < values.length - 1) {
      inlines.push({ type: 'TextRun', text: ' · ' });
    }
  });
  return { type: 'RichTextBlock', spacing: 'Small', inlines };
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
 * Header summary for the Privacy-Shield toggle button — e.g.
 * "1 Tool-Ergebnis · 1 Feld maskiert".
 */
function privacyReceiptSummary(r: PrivacyReceipt): string {
  const datasets = `${String(r.datasetsInterned)} Tool-Ergebnis${
    r.datasetsInterned === 1 ? '' : 'se'
  }`;
  const masked =
    r.fieldsMasked > 0
      ? ` · ${String(r.fieldsMasked)} Feld${
          r.fieldsMasked === 1 ? '' : 'er'
        } maskiert`
      : '';
  const onWire = r.identityValuesOnWire ?? 0;
  const named =
    onWire > 0
      ? ` · ⚠ ${String(onWire)} Name${onWire === 1 ? '' : 'n'} ans Modell`
      : '';
  return `${datasets}${masked}${named}`;
}

function buildPrivacyReceiptItems(r: PrivacyReceipt): unknown[] {
  const facts: Array<{ title: string; value: string }> = [
    { title: 'Tool-Ergebnisse interned', value: String(r.datasetsInterned) },
    { title: 'Felder maskiert', value: String(r.fieldsMasked) },
    { title: 'Felder Klartext', value: String(r.fieldsCleartext) },
    {
      title: 'Verben',
      value: r.verbsExecuted.length > 0 ? r.verbsExecuted.join(', ') : '—',
    },
    {
      title: 'Pseudonym-Projektion',
      value: r.pseudonymProjectionUsed ? 'verwendet' : 'nicht verwendet',
    },
  ];
  // Transparency notice — identity values the requester named themselves.
  // Shown only when it actually happened.
  if ((r.identityValuesOnWire ?? 0) > 0) {
    facts.push({
      title: 'Namen ans Modell (selbst genannt)',
      value: String(r.identityValuesOnWire),
    });
  }
  return [
    {
      type: 'TextBlock',
      text: '**🛡 Privacy Shield v4 · Data-Plane-Boundary**',
      weight: 'Bolder',
      size: 'Small',
      wrap: true,
    },
    { type: 'FactSet', facts },
    {
      type: 'TextBlock',
      text: 'Rohe Tool-Ergebnisse blieben server-seitig — das Modell sah nur einen identitätsfreien Digest. Server-seitig eingesetzte Werte sind in der Antwort violett hervorgehoben.',
      size: 'Small',
      isSubtle: true,
      wrap: true,
      spacing: 'Small',
    },
  ];
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

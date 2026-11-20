import type { Attachment } from 'botbuilder';
import { CardFactory } from 'botbuilder';
import type { AnswerDetails } from '@omadia/channel-sdk';
import type {
  AiDisclosure,
  Contribution,
  OutgoingAttachment,
  Speaker,
  SystemNotice,
} from './kernel-types.js';
import type { MentionEntity } from './teamsMentions.js';
import type { AgentAppInstallOutcome } from './teamsAgentInstaller.js';
import { contributionBlocks, systemNoticeBody } from './teamsAnswerParts.js';
import { buildAnswerDetailsElements } from './teamsCardDetails.js';
import {
  admitTeamsDecorations,
  assertTeamsDecoration,
  TEAMS_DECORATION_IDS,
} from './teamsDecorations.js';

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
/** Answers that do not fit the card budget are cut to this many characters. */
const ANSWER_MAX_BYTES_WHEN_TRUNCATED = 20_000;

const AI_LABEL_ENTITY = {
  type: 'https://schema.org/Message',
  '@type': 'Message',
  '@context': 'https://schema.org',
  // Empty on purpose. Microsoft's Teams SDK types this key as required with
  // the comment "Must be left blank. This is for Bot Framework schema", and
  // its `addAiGenerated()` always sends it. Until 0.38.2 this plugin sent the
  // entity WITHOUT it, and the platform label was never seen on a message of
  // this plugin (D24).
  '@id': '',
  additionalType: ['AIGeneratedContent'],
} as const;

/**
 * AI-Generated-Content marker attached to a Teams message activity: the
 * entity Microsoft's Teams SDK produces for `addAiGenerated()` (microsoft/
 * teams.ts, `packages/api/src/activities/activity.ts` and
 * `models/entity/message-entity.ts`, read 2026-11-19). Teams shows it as
 * "AI Generated" next to the agent's name in the message header.
 *
 * NOT TO BE CONFUSED with the "✨ AI-generated" chip that the answer card
 * draws itself (`statusChips` below): that chip is this plugin's own text
 * inside the card and is visible whether or not Teams renders its label.
 */
export function aiLabelEntity(): Readonly<typeof AI_LABEL_ENTITY> {
  return AI_LABEL_ENTITY;
}

/**
 * AI-Act Art. 50 (#643/#644) — remove the disclosure line the kernel folded
 * into `SemanticAnswer.text` for wire-only channels. Teams has its own
 * provenance surface (the "✨ AI-generated" card chip + the activity-level
 * {@link aiLabelEntity}), so the folded sentence would mark the answer twice.
 * Only the marking line is removed; an operator-authored addendum
 * (`aiDisclosure.operatorNote`, folded as its own paragraph after the line)
 * stays — the operator asked for it explicitly. When the answer consists of
 * nothing but the marking line, it is kept: an empty card body would
 * otherwise degrade to the "keine Antwort" error path.
 */
export function stripFoldedAiDisclosure(
  text: string,
  disclosureLine: string | undefined,
): string {
  if (!disclosureLine || disclosureLine.trim().length === 0) return text;
  if (text === disclosureLine) return text;
  const folded = `\n\n${disclosureLine}`;
  if (text.endsWith(folded)) return text.slice(0, -folded.length);
  // Line folded mid-text (operator note follows as its own paragraph).
  const withNote = `${folded}\n\n`;
  const idx = text.lastIndexOf(withNote);
  if (idx !== -1) return text.slice(0, idx) + text.slice(idx + folded.length);
  return text;
}


export interface BuildAnswerCardInput {
  answer: string;
  /**
   * Image attachments to embed inline under the answer. Today only diagram
   * renders (render_diagram tool) populate this; the shape is generic so
   * future attachment producers can reuse the same path.
   */
  attachments?: DiagramAttachment[];
  /**
   * The SDK's one projection of the answer's details
   * (`buildAnswerDetails(answer)`): verification, data provenance, colleagues
   * who contributed, privacy counters, memory use. Drawn as ONE collapsed
   * "Details" element (decoration `answer_details`); never as chips or
   * footers on the card. `undefined`: no details element at all.
   */
  details?: AnswerDetails;
  /**
   * Draw the card's discreet "✨ AI-generated" chip (decoration
   * `ai_disclosure_chip`). `false` means the kernel did not ask for a visible
   * marker on this answer: the operator turned the disclosure off, or the
   * frequency setting says it is not due (byte5ai/omadia#2087). Absent or
   * `true`: the chip is drawn (a caller without kernel disclosure data keeps
   * the card of old). {@link aiChipDue} derives the value from a kernel
   * answer.
   */
  aiMarker?: boolean;
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
   * Concrete next steps the agent itself proposed (`suggest_follow_ups`) and
   * expects the user to pick from: decoration `follow_up_options`. Rendered as
   * `Action.Submit` buttons under the answer; up to 4 options, each carrying a
   * full `prompt` that is submitted as a fresh user turn on click. Without
   * proposed options there is no button row.
   */
  followUpOptions?: ReadonlyArray<{ label: string; prompt: string }>;
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
   * #2045 — who the message is from. Teams shows an agent natively (its own
   * bot), so an agent speaker adds nothing to the card; a `system` speaker
   * turns the card into a management-layer notice.
   */
  speaker?: Speaker;
  /**
   * "Mira:" — shown above the answer where one bot serves several agents and
   * the speaker would be unclear (INV-4; the SDK's `speakerLine` decides).
   * Never on a card of an agent's own bot, never on a `system` notice.
   */
  speakerLine?: string;
  /** #2045 — the notice text of a `system` speaker; defaults to `answer`. */
  systemNotice?: SystemNotice;
  /**
   * #2047 — colleagues' delivered contributions (spec 2042 §5.4), rendered
   * attributed to the consulted agent, with its text unchanged, below the
   * speaking agent's own text. A failed contribution is shown as failed.
   */
  contributions?: readonly Contribution[];
}

/**
 * Whether the card draws its AI chip for this kernel answer (decoration
 * `ai_disclosure_chip`, byte5ai/omadia#2087). Fail-safe in the direction of
 * marking:
 *
 *  - no `aiDisclosure` at all: the operator turned the disclosure `off`, no
 *    chip;
 *  - `aiDisclosure` without `markers` (a kernel that predates the frequency
 *    setting): the default `channel_default`, i.e. a chip on every card;
 *  - otherwise the chip is drawn exactly when `markers.chip` is true.
 */
export function aiChipDue(aiDisclosure: AiDisclosure | undefined): boolean {
  if (aiDisclosure === undefined) return false;
  return aiDisclosure.markers?.chip !== false;
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
  assertTeamsDecoration(TEAMS_DECORATION_IDS.decisionCard);
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
 * after the chunks so the buttons still appear. Returns undefined when there
 * is nothing to offer or the decoration registry does not admit the buttons.
 */
export function buildFollowUpsOnlyCard(
  options: ReadonlyArray<{ label: string; prompt: string }>,
): Attachment | undefined {
  if (options.length === 0) return undefined;
  if (!admitTeamsDecorations([TEAMS_DECORATION_IDS.followUps]).has(TEAMS_DECORATION_IDS.followUps)) {
    return undefined;
  }
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
 * Builds the final Teams Adaptive Card for a completed turn. A normal answer
 * is the answer text and nothing else, plus (spec 2042 §5.14, every element
 * declared in the channel SDK's decoration registry and admitted through
 * {@link admitTeamsDecorations}):
 *
 *  - the discreet "✨ AI-generated" chip, when the kernel's disclosure
 *    settings ask for a visible marker on this answer (`ai_disclosure_chip`);
 *  - the colleagues' contributions, attributed to the consulted agent;
 *  - ONE collapsed "Details" element from the SDK's details projection
 *    (`answer_details`; `Action.ToggleVisibility` is evaluated on the viewer's
 *    client, so expanding never shows anything to the rest of a group chat);
 *  - buttons only for next steps the agent itself proposed
 *    (`follow_up_options`).
 *
 * A `system` speaker renders as a management-layer notice instead.
 *
 * Size-aware: if the serialised card would exceed Teams' 28 KB cap the answer
 * is truncated; images and contributions are never shortened.
 */
export function buildAnswerCard(input: BuildAnswerCardInput): Attachment {
  const cardBody =
    input.speaker?.kind === 'system'
      ? { body: systemNoticeBody(input.systemNotice?.text ?? input.answer) }
      : buildCardBody(input);
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
    ...('actions' in cardBody && cardBody.actions ? { actions: cardBody.actions } : {}),
    msteams,
  });
}

interface CardBody {
  body: unknown[];
  actions?: unknown[];
}

/** The decoration parts of one card, after the registry gate. */
interface AdmittedParts {
  readonly chip: boolean;
  readonly followUps: ReadonlyArray<{ label: string; prompt: string }>;
  readonly details: AnswerDetails | undefined;
}

function truncateAnswer(answer: string): string {
  return `${answer.slice(0, ANSWER_MAX_BYTES_WHEN_TRUNCATED - 1)}…`;
}

function buildCardBody(input: BuildAnswerCardInput): CardBody {
  const attachments = (input.attachments ?? []).slice(0, MAX_ATTACHMENTS_PER_CARD);
  const proposed = (input.followUpOptions ?? []).slice(0, FOLLOW_UP_MAX_OPTIONS);
  const hasDetails = buildAnswerDetailsElements(input.details) !== undefined;
  // The registry gate: nothing below is attached unless it was admitted.
  const admitted = admitTeamsDecorations([
    ...(input.aiMarker !== false ? [TEAMS_DECORATION_IDS.aiChip] : []),
    ...(hasDetails ? [TEAMS_DECORATION_IDS.details] : []),
    ...(proposed.length > 0 ? [TEAMS_DECORATION_IDS.followUps] : []),
  ]);
  const parts: AdmittedParts = {
    chip: admitted.has(TEAMS_DECORATION_IDS.aiChip),
    followUps: admitted.has(TEAMS_DECORATION_IDS.followUps) ? proposed : [],
    details: admitted.has(TEAMS_DECORATION_IDS.details) ? input.details : undefined,
  };

  const full = assemble(input, input.answer, attachments, parts);
  if (bytes(full) <= CARD_BUDGET_BYTES) return full;
  // Truncate the answer but still keep the images and the colleagues' words:
  // they carry the actual payload the user asked for.
  return assemble(input, truncateAnswer(input.answer), attachments, parts);
}

function bytes(body: CardBody): number {
  return Buffer.byteLength(JSON.stringify(body), 'utf8');
}

/** The answer card's own AI marking. NOT Teams' platform label: this is a
 *  text line this plugin draws, in place of the kernel's disclosure sentence. */
export const AI_MARKER_CHIP = '✨ AI-generated';

function aiChipContainer(): unknown {
  return {
    type: 'Container',
    id: TEAMS_DECORATION_IDS.aiChip,
    spacing: 'None',
    items: [
      {
        type: 'TextBlock',
        text: AI_MARKER_CHIP,
        size: 'Small',
        weight: 'Bolder',
        color: 'Accent',
        wrap: true,
      },
    ],
  };
}

/**
 * Attachments sit directly under the answer.
 *  - image: rendered inline (Teams renders PNG/JPEG/GIF up to 1024x1024 / 1 MB;
 *    SVG is not supported, the service never emits it).
 *  - file (e.g. create_xlsx .xlsx, create_docx .docx): Teams has no inline
 *    element for an arbitrary blob, so we surface a one-click download button.
 *    Action.OpenUrl opens the signed /documents URL in the browser, which
 *    streams the file. (Native upload into the channel's Files tab would need
 *    a Graph file-consent flow — deferred.)
 */
function attachmentElements(attachments: readonly DiagramAttachment[]): unknown[] {
  const elements: unknown[] = [];
  for (const att of attachments) {
    if (att.kind === 'image') {
      elements.push({
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
      elements.push({
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
    }
  }
  return elements;
}

function assemble(
  input: BuildAnswerCardInput,
  answer: string,
  attachments: readonly DiagramAttachment[],
  parts: AdmittedParts,
): CardBody {
  const body: unknown[] = [
    ...(input.speakerLine
      ? [{ type: 'TextBlock', text: input.speakerLine, weight: 'Bolder', size: 'Small', wrap: true }]
      : []),
    ...(parts.chip ? [aiChipContainer()] : []),
    // Answer. A v4 table renders as a native Adaptive-Card `Table` with
    // masked cells in the accent (violet) colour; prose / non-table answers
    // stay a markdown `TextBlock` (plus a compact violet line when they
    // resolved masked values).
    ...buildAnswerElements(answer, input.maskedValues ?? []),
    ...attachmentElements(attachments),
    // Colleagues' words: attributed to them, unchanged, never dropped.
    ...contributionBlocks(input.contributions, input.speaker),
  ];

  const actions: unknown[] = [];
  const detailElements = buildAnswerDetailsElements(parts.details);
  if (detailElements) {
    body.push(detailElements.container);
    actions.push(detailElements.toggle);
  }

  // Next steps the agent proposed: the buttons replace typing the option.
  if (parts.followUps.length > 0) {
    body.push({
      type: 'ActionSet',
      id: TEAMS_DECORATION_IDS.followUps,
      spacing: 'Medium',
      actions: parts.followUps.map((opt) => ({
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
  assertTeamsDecoration(TEAMS_DECORATION_IDS.decisionCard);
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

// --- Conductor human-await approval card (Ja/Nein) -----------------------
// A proactive reminder for a Conductor human step, rendered as an Adaptive Card
// that shows WHAT is being approved + WHERE we are in the workflow, with two
// Action.Submit buttons. A click arrives back as a normal message activity whose
// `value` is an ApprovalValue; the bot resolves the await in-process.
export const APPROVAL_VALUE_TYPE = 'conductor.approval';
export type ApprovalDecision = 'approve' | 'reject';

export interface ApprovalValue {
  type: typeof APPROVAL_VALUE_TYPE;
  decision: ApprovalDecision;
  /** the pending await this click resolves — round-tripped from the card. */
  awaitId: string;
}

export interface BuildApprovalCardInput {
  awaitId: string;
  /** WHAT is being approved — the human step's message. */
  question: string;
  workflowName: string;
  stepLabel: string;
  stepIndex?: number;
  totalSteps?: number;
  quorum: 'any' | 'all';
}

export function buildApprovalCard(input: BuildApprovalCardInput): Attachment {
  assertTeamsDecoration(TEAMS_DECORATION_IDS.decisionCard);
  const progress =
    input.stepIndex !== undefined && input.totalSteps !== undefined
      ? `Schritt ${String(input.stepIndex)}/${String(input.totalSteps)} · ${input.stepLabel}`
      : input.stepLabel;
  const body: Record<string, unknown>[] = [
    {
      type: 'TextBlock',
      text: '🔔 Freigabe erforderlich',
      weight: 'Bolder',
      size: 'Medium',
      wrap: true,
    },
    {
      type: 'TextBlock',
      text: `**${input.workflowName}** · ${progress}`,
      size: 'Small',
      isSubtle: true,
      spacing: 'None',
      wrap: true,
    },
    {
      type: 'Container',
      style: 'emphasis',
      items: [
        {
          type: 'TextBlock',
          text: input.question,
          wrap: true,
        },
      ],
    },
  ];
  if (input.quorum === 'all') {
    body.push({
      type: 'TextBlock',
      text: 'ℹ️ Alle Freigeber müssen zustimmen.',
      size: 'Small',
      isSubtle: true,
      wrap: true,
    });
  }
  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body,
    actions: [
      {
        type: 'Action.Submit',
        title: '✓ Ja, genehmigen',
        style: 'positive',
        data: {
          type: APPROVAL_VALUE_TYPE,
          decision: 'approve',
          awaitId: input.awaitId,
        } satisfies ApprovalValue,
      },
      {
        type: 'Action.Submit',
        title: '✗ Nein, ablehnen',
        data: {
          type: APPROVAL_VALUE_TYPE,
          decision: 'reject',
          awaitId: input.awaitId,
        } satisfies ApprovalValue,
      },
    ],
  });
}

/** Runtime guard for an approve/reject button submit (mirrors parseTopicDecisionValue).
 *  Some Teams clients deliver the Action.Submit `data` as a JSON STRING in `activity.value`
 *  rather than a parsed object — accept both. */
export function parseApprovalValue(value: unknown): ApprovalValue | undefined {
  let candidate: unknown = value;
  if (typeof candidate === 'string') {
    try {
      candidate = JSON.parse(candidate);
    } catch {
      return undefined;
    }
  }
  if (!candidate || typeof candidate !== 'object') return undefined;
  const v = candidate as Record<string, unknown>;
  if (v['type'] !== APPROVAL_VALUE_TYPE) return undefined;
  const decision = v['decision'];
  if (decision !== 'approve' && decision !== 'reject') return undefined;
  const awaitId = v['awaitId'];
  if (typeof awaitId !== 'string' || awaitId.length === 0) return undefined;
  return { type: APPROVAL_VALUE_TYPE, decision, awaitId };
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
  assertTeamsDecoration(TEAMS_DECORATION_IDS.decisionCard);
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

// -----------------------------------------------------------------------------
// #860 W2 (issue #21) — auto-invite result / fallback card
// -----------------------------------------------------------------------------
// Rendered by the onboarding hook after a `teamsAgentInstaller` run: one line
// per configured agent app, a per-reason hint block (consent missing vs. app
// not in catalog), install deep links for apps that could not be installed,
// and a "Prüfen" re-check Action.Submit that re-runs the installer and
// UPDATES this card in place. Deep links carry ONLY public Teams app ids.

/** Deep-link base for a published Teams app — public app id only. */
const TEAMS_APP_DEEP_LINK_BASE = 'https://teams.microsoft.com/l/app/';

export const AGENT_APPS_RECHECK_VALUE_TYPE = 'agent_apps_recheck';

/**
 * Submit payload of the "Prüfen" button — the DISCRIMINATOR AND NOTHING
 * ELSE (#1030). Card `data` is client-editable, so a round-tripped
 * `teamId`/`tenantId` was an install target the clicker could choose:
 * `handleTeamsAgentAppsRecheck` derives the scope from the submit
 * activity's own `channelData` and refuses the click when it cannot,
 * exactly as the Conductor approve/reject branch derives its responder
 * from the session. Cards posted before this change still carry the two
 * fields; the parser ignores every extra key, so old cards keep working
 * and the values are simply never read.
 */
export interface AgentAppsRecheckValue {
  type: typeof AGENT_APPS_RECHECK_VALUE_TYPE;
}

/** Runtime guard for the re-check submit. Some Teams clients deliver the
 *  Action.Submit `data` as a JSON STRING (parseApprovalValue precedent) —
 *  accept both. */
export function parseAgentAppsRecheckValue(
  value: unknown,
): AgentAppsRecheckValue | undefined {
  let candidate: unknown = value;
  if (typeof candidate === 'string') {
    try {
      candidate = JSON.parse(candidate);
    } catch {
      return undefined;
    }
  }
  if (!candidate || typeof candidate !== 'object') return undefined;
  const v = candidate as Record<string, unknown>;
  if (v['type'] !== AGENT_APPS_RECHECK_VALUE_TYPE) return undefined;
  // Deliberately narrowing: whatever else the payload carries is dropped
  // here, so no client-supplied value can reach an install path (#1030).
  return { type: AGENT_APPS_RECHECK_VALUE_TYPE };
}

export interface BuildAgentAppsCardInput {
  /** Per-app outcomes of one installer run, in configured order. */
  readonly outcomes: readonly AgentAppInstallOutcome[];
}

/** One status line per app — the card's scannable core. */
function agentAppOutcomeLine(outcome: AgentAppInstallOutcome): Record<string, unknown> {
  let text: string;
  switch (outcome.kind) {
    case 'installed':
      text =
        outcome.outcome === 'created'
          ? `✅ **${outcome.displayName}** — installiert`
          : `✅ **${outcome.displayName}** — war bereits installiert`;
      break;
    case 'fallback':
      text =
        outcome.reason === 'not-in-catalog'
          ? `⚠️ **${outcome.displayName}** — nicht im Teams-App-Katalog gefunden`
          : `⚠️ **${outcome.displayName}** — Admin-Zustimmung fehlt`;
      break;
    case 'failed':
      text =
        outcome.reason === 'throttled'
          ? `⏳ **${outcome.displayName}** — Microsoft drosselt gerade, bitte später erneut prüfen`
          : `❌ **${outcome.displayName}** — Installation fehlgeschlagen`;
      break;
  }
  return { type: 'TextBlock', text, wrap: true, spacing: 'Small' };
}

/**
 * Result / fallback card for one auto-invite installer run.
 *
 * Distinguishes the two fallback classes explicitly: CONSENT MISSING lists
 * the Graph application permissions an admin still has to grant (scope
 * names are card material, never log material); NOT IN CATALOG points at
 * the org-catalog upload. Fallback apps with a known catalog id get an
 * `Action.OpenUrl` install deep link (`https://teams.microsoft.com/l/app/
 * <teamsAppId>` — public id only). Unless every app is already installed,
 * a "Prüfen" Action.Submit re-runs the installer.
 */
export function buildAgentAppsResultCard(
  input: BuildAgentAppsCardInput,
): Attachment {
  const body: Record<string, unknown>[] = [
    {
      type: 'TextBlock',
      text: '🤝 Agenten-Apps für dieses Team',
      weight: 'Bolder',
      size: 'Medium',
      wrap: true,
    },
    ...input.outcomes.map(agentAppOutcomeLine),
  ];

  const consentFallbacks = input.outcomes.filter(
    (o): o is Extract<AgentAppInstallOutcome, { kind: 'fallback' }> =>
      o.kind === 'fallback' && o.reason !== 'not-in-catalog',
  );
  const missingScopes = [
    ...new Set(consentFallbacks.flatMap((o) => o.missingScopes ?? [])),
  ];
  if (consentFallbacks.length > 0) {
    body.push({
      type: 'Container',
      style: 'emphasis',
      spacing: 'Medium',
      items: [
        {
          type: 'TextBlock',
          text:
            missingScopes.length > 0
              ? `Ein Admin muss der App-Registrierung noch folgende Graph-Berechtigungen zustimmen: ${missingScopes.map((s) => `\`${s}\``).join(', ')} (siehe Setup-Guide, Abschnitt Auto-Invite).`
              : 'Ein Admin muss der App-Registrierung noch die Graph-Berechtigungen für Auto-Invite zustimmen (siehe Setup-Guide).',
          size: 'Small',
          wrap: true,
        },
        {
          type: 'TextBlock',
          text: 'Bis dahin kannst du die Apps über die Links unten manuell hinzufügen.',
          size: 'Small',
          isSubtle: true,
          wrap: true,
          spacing: 'Small',
        },
      ],
    });
  }
  if (input.outcomes.some((o) => o.kind === 'fallback' && o.reason === 'not-in-catalog')) {
    body.push({
      type: 'TextBlock',
      text: 'ℹ️ Nicht gefundene Apps zuerst in den Teams-App-Katalog der Organisation hochladen — danach hier „Prüfen" klicken.',
      size: 'Small',
      isSubtle: true,
      wrap: true,
      spacing: 'Medium',
    });
  }

  const actions: Record<string, unknown>[] = [];
  for (const outcome of input.outcomes) {
    if (outcome.kind === 'fallback' && outcome.teamsAppId !== undefined) {
      actions.push({
        type: 'Action.OpenUrl',
        title: `${outcome.displayName} öffnen`,
        url: `${TEAMS_APP_DEEP_LINK_BASE}${encodeURIComponent(outcome.teamsAppId)}`,
      });
    }
  }
  const allInstalled = input.outcomes.every((o) => o.kind === 'installed');
  if (!allInstalled) {
    actions.push({
      type: 'Action.Submit',
      title: '🔄 Prüfen',
      data: {
        type: AGENT_APPS_RECHECK_VALUE_TYPE,
        // Same Teams quirk as followUpAction/choiceAsk — Action.Submit on a
        // slow-responding bot shows "Something went wrong" unless flagged
        // as messageBack.
        msteams: {
          type: 'messageBack',
          text: 'Prüfen',
          displayText: '🔄 Prüfen',
        },
      } satisfies AgentAppsRecheckValue & { msteams: Record<string, unknown> },
    });
  }

  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body,
    ...(actions.length > 0 ? { actions } : {}),
  });
}

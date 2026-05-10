import { CardFactory, type Attachment } from 'botbuilder';

import type { ParsedNudge } from '@omadia/plugin-api';

/**
 * Palaia Phase 8 (OB-77 Slice 4) — Teams Adaptive-Card builder for
 * `palaia.*`-style nudges.
 *
 * Renders a parsed nudge as a separate Adaptive-Card the bot attaches
 * alongside the answer. CTA gets `Action.Submit` so the click fires a
 * tool-call-like next-turn payload via the bot's existing message
 * handler. The "Nicht mehr anzeigen"-link is a second `Action.Submit`
 * so the bot's value-router can dispatch to `nudgeStateStore.suppress`.
 *
 * Notes:
 *  - The card stays minimal: a TextBlock with the hint, an optional
 *    CTA-button, and a suppress-link. No accent-bar / no avatar / no
 *    nested containers — the nudge is a coaching surface, not a
 *    primary answer-card. Operators can grow it later via this same
 *    builder once the pattern proves out.
 *  - `version: '1.5'` mirrors the buildTopicAskCard / buildChoiceAskCard
 *    conventions in this package — Teams desktop + mobile both render
 *    1.5 features cleanly.
 *  - Action.Submit's `data` is round-tripped by Teams as the next user
 *    message's `value`. The bot's existing turn-handler dispatches on
 *    `data.type` — the two value-types below mirror the existing
 *    TopicDecisionValue / ChoiceAskValue pattern.
 */

export const NUDGE_CTA_VALUE_TYPE = 'palaia.nudge.cta';
export const NUDGE_SUPPRESS_VALUE_TYPE = 'palaia.nudge.suppress';

export interface NudgeCtaValue {
  type: typeof NUDGE_CTA_VALUE_TYPE;
  /** Provider id (e.g. `palaia.process-promote`) — used by the state
   *  store to mark the CTA as followed. */
  nudgeId: string;
  /** Tool name the CTA invokes (e.g. `write_process`). */
  toolName: string;
  /** Pre-filled tool arguments. The bot forwards these to the
   *  orchestrator's next-turn dispatch. */
  toolArgs: Record<string, unknown>;
}

export interface NudgeSuppressValue {
  type: typeof NUDGE_SUPPRESS_VALUE_TYPE;
  /** Provider id whose state row should get its `suppressed_until` set. */
  nudgeId: string;
}

/**
 * Build an Adaptive-Card attachment for a parsed nudge. Returns
 * `undefined` when there's nothing to render (`text` is empty), so the
 * caller can splice the optional value into its attachments array
 * without a guard.
 */
export function buildNudgeCard(nudge: ParsedNudge): Attachment | undefined {
  const text = nudge.text.trim();
  if (text.length === 0) return undefined;

  const actions: Array<Record<string, unknown>> = [];
  if (nudge.cta) {
    actions.push({
      type: 'Action.Submit',
      title: nudge.cta.label,
      style: 'positive',
      data: {
        type: NUDGE_CTA_VALUE_TYPE,
        nudgeId: nudge.id,
        toolName: nudge.cta.toolCall.name,
        toolArgs: nudge.cta.toolCall.arguments,
      } satisfies NudgeCtaValue,
    });
  }
  actions.push({
    type: 'Action.Submit',
    title: 'Nicht mehr anzeigen',
    data: {
      type: NUDGE_SUPPRESS_VALUE_TYPE,
      nudgeId: nudge.id,
    } satisfies NudgeSuppressValue,
  });

  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body: [
      {
        type: 'TextBlock',
        text: '💡 Vorschlag',
        weight: 'Bolder',
        size: 'Small',
        color: 'Accent',
        spacing: 'None',
      },
      {
        type: 'TextBlock',
        text,
        wrap: true,
        spacing: 'Small',
      },
    ],
    actions,
  });
}

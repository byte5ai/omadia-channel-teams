import { answerDetailsText, type AnswerDetails } from '@omadia/channel-sdk';

/**
 * The one collapsed "Details" element of an answer card (spec 2042 §5.14,
 * decoration `answer_details`).
 *
 * Everything that used to sit on the card as its own chip, footer or button
 * (verifier badge, system-data chip, "Konsultiert" footer, privacy and memory
 * sections) is part of the channel SDK's single projection of the answer's
 * details, `buildAnswerDetails(answer)`. This module only draws that
 * projection: the card carries one "Details" button and a container that stays
 * hidden until the viewer expands it (`Action.ToggleVisibility` is evaluated on
 * the viewer's client, so expanding never shows anything to the rest of a group
 * chat). No projection (nothing to show) means no element at all.
 */

/** Id of the hidden container the "Details" button toggles. */
export const ANSWER_DETAILS_CONTAINER_ID = 'answer_details';

/** Title of the one button that opens the details. */
export const ANSWER_DETAILS_TITLE = 'Details';

/** Cap the specialist list so a many-agent turn cannot grow the element without bound. */
const SPECIALISTS_MAX = 6;

export interface AnswerDetailsElements {
  /** The hidden container; goes into the card body. */
  readonly container: Record<string, unknown>;
  /** The "Details" button; goes into the card's `actions`. */
  readonly toggle: Record<string, unknown>;
}

function specialistsLine(specialists: AnswerDetails['specialistsInvoked']): string | undefined {
  if (specialists.length === 0) return undefined;
  const shown = specialists.slice(0, SPECIALISTS_MAX);
  const extra = specialists.length - shown.length;
  const list = shown
    .map((s) => `${s.label} ${s.status === 'success' ? '✓' : '✗'}${s.toolCalls ? ` · ${String(s.toolCalls)}` : ''}`)
    .join(' · ');
  return `Eingesetzt: ${list}${extra > 0 ? ` · +${String(extra)}` : ''}`;
}

/**
 * The details container and its button, or `undefined` when there is nothing
 * to show. The caller passes the result through the decoration gate.
 */
export function buildAnswerDetailsElements(
  details: AnswerDetails | undefined,
  locale: 'de' | 'en' = 'de',
): AnswerDetailsElements | undefined {
  if (!details) return undefined;
  const lines = [answerDetailsText(details, locale), specialistsLine(details.specialistsInvoked)]
    .filter((line): line is string => line !== undefined)
    .flatMap((block) => block.split('\n'));
  if (lines.length === 0) return undefined;
  return {
    container: {
      type: 'Container',
      id: ANSWER_DETAILS_CONTAINER_ID,
      isVisible: false,
      spacing: 'Medium',
      separator: true,
      items: lines.map((text, index) => ({
        type: 'TextBlock',
        text,
        size: 'Small',
        isSubtle: true,
        wrap: true,
        ...(index > 0 ? { spacing: 'None' } : {}),
      })),
    },
    toggle: {
      type: 'Action.ToggleVisibility',
      title: ANSWER_DETAILS_TITLE,
      targetElements: [ANSWER_DETAILS_CONTAINER_ID],
    },
  };
}

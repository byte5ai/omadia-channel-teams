/**
 * The Teams side of the kernel's decoration registry (spec 2042 §5.14,
 * byte5ai/omadia#2086, ADR-0021).
 *
 * A decoration is everything on or around an answer that is not the answer
 * itself: a chip, a footer, a button, a details element. The channel SDK
 * declares which ones may exist and which channels render them
 * (`DECORATION_REGISTRY`); this module is the one place where the Teams card
 * code asks for them. Nothing reaches a card without passing
 * {@link admitTeamsDecorations}, so a decoration the registry does not declare
 * for `teams` (or has removed) is dropped instead of rendered.
 */

import { admitDecorations, type DecorationRefusal } from '@omadia/channel-sdk';

import { TEAMS_CHANNEL_TYPE } from './teamsKernelRegistration.js';


/**
 * Teams shows one identity per agent (the agent's own bot: name and icon), so
 * an answer carries no name line (spec 2042 INV-4, §5.3). Documented here as a
 * constant because the SDK's plain-text renderers take it as an option.
 */
export const TEAMS_NATIVE_IDENTITY = true;

/** Every decoration id this channel attaches, declared once. */
export const TEAMS_DECORATION_IDS = Object.freeze({
  aiChip: 'ai_disclosure_chip',
  details: 'answer_details',
  followUps: 'follow_up_options',
  decisionCard: 'decision_card',
  consentCard: 'consent_card',
  routineList: 'routine_list',
  facilitationPoll: 'facilitation_poll',
} as const);

export type TeamsDecorationId = (typeof TEAMS_DECORATION_IDS)[keyof typeof TEAMS_DECORATION_IDS];

/** Where refusals go: the operator log, once per refused decoration. */
function reportRefusal(refusal: DecorationRefusal): void {
  console.warn(
    `[teams] decoration '${refusal.id}' refused by the channel SDK (${refusal.reason}); not rendered`,
  );
}

/**
 * The subset of `ids` the registry admits for Teams. Fail-closed: an id the
 * registry does not know, has removed, or does not list for `teams` is absent
 * from the result.
 */
export function admitTeamsDecorations(
  ids: readonly string[],
  onRefused: (refusal: DecorationRefusal) => void = reportRefusal,
): ReadonlySet<string> {
  const { admitted } = admitDecorations(
    TEAMS_CHANNEL_TYPE,
    ids.map((id) => ({ id })),
    onRefused,
  );
  return new Set(admitted.map((decoration) => decoration.id));
}

/** True when the registry admits this one decoration for Teams. */
export function isTeamsDecorationAdmitted(id: string): boolean {
  return admitTeamsDecorations([id]).has(id);
}

/**
 * For a standalone card whose whole purpose is one registered decoration (a
 * decision card, the consent card, a poll): fail closed when the registry does
 * not admit it for Teams, instead of posting an undeclared card.
 */
export function assertTeamsDecoration(id: TeamsDecorationId): void {
  if (!isTeamsDecorationAdmitted(id)) {
    throw new Error(`decoration '${id}' is not admitted for ${TEAMS_CHANNEL_TYPE} by the channel SDK registry`);
  }
}

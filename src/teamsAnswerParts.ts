/**
 * The parts of an answer card that carry other speakers' words (spec 2042
 * INV-3, INV-4, INV-5, §5.4): a colleague's contribution and a `system`
 * notice.
 *
 * Teams shows the speaker of the whole message natively (the agent's own bot:
 * name and icon), so the card has no name line; only where one bot serves
 * several agents does the card carry the SDK's name line (INV-4). Two things
 * are different and must never look like the bot speaking for itself:
 *
 *  - A consulted agent B answers inside A's message. B's block is attributed
 *    to B ("B (konsultiert von A)") and carries B's text exactly as delivered; A's
 *    own text stays A's. A failed contribution is shown as failed.
 *  - A `system` notice is a message of the management layer, not of any agent.
 *    It is rendered under its own heading ("System") so it is never read as the
 *    words of the bot it is posted through.
 */

import type { Contribution, Speaker, SystemNotice } from './kernel-types.js';

/** Heading of a management-layer notice. */
export const SYSTEM_NOTICE_LABEL = 'System';

/** Id prefix of a contribution container; the contribution's own id follows. */
export const CONTRIBUTION_CONTAINER_ID_PREFIX = 'contribution-';

/**
 * "Mira (konsultiert von Jonas)". Without a speaker on the answer (a kernel that
 * predates it) the attribution still names the consulted agent.
 */
export function contributionLabel(
  contribution: Contribution,
  consultedBy: Speaker | undefined,
): string {
  const name = contribution.speaker.displayName;
  const by = consultedBy?.kind === 'agent' ? ` von ${consultedBy.displayName}` : '';
  const failed = contribution.status === 'failed' ? ' – fehlgeschlagen' : '';
  return `${name} (konsultiert${by})${failed}`;
}

function contributionBlock(contribution: Contribution, consultedBy: Speaker | undefined): unknown {
  const isFailed = contribution.status === 'failed';
  return {
    type: 'Container',
    id: `${CONTRIBUTION_CONTAINER_ID_PREFIX}${contribution.id}`,
    style: 'emphasis',
    spacing: 'Medium',
    separator: true,
    items: [
      {
        type: 'TextBlock',
        text: contributionLabel(contribution, consultedBy),
        weight: 'Bolder',
        size: 'Small',
        color: isFailed ? 'Attention' : 'Accent',
        wrap: true,
      },
      // B's text exactly as delivered: no trimming, no rewording.
      { type: 'TextBlock', text: contribution.text, wrap: true, spacing: 'Small' },
    ],
  };
}

/** The card elements of the contributions, in the order the kernel delivered them. */
export function contributionBlocks(
  contributions: readonly Contribution[] | undefined,
  consultedBy: Speaker | undefined,
): unknown[] {
  return (contributions ?? []).map((contribution) => contributionBlock(contribution, consultedBy));
}

/** Total length of the contribution texts; the card/plain-text decision counts it. */
export function contributionsLength(contributions: readonly Contribution[] | undefined): number {
  return (contributions ?? []).reduce((sum, contribution) => sum + contribution.text.length, 0);
}

/** The plain-text form (long answers are posted as text chunks): same attribution, same text. */
export function contributionsAsText(
  contributions: readonly Contribution[] | undefined,
  consultedBy: Speaker | undefined,
): string {
  return (contributions ?? [])
    .map((contribution) => `${contributionLabel(contribution, consultedBy)}:\n${contribution.text}`)
    .join('\n\n');
}

/**
 * Who consulted the colleagues of an answer the management layer replaced:
 * the agent the notice is about. A withheld answer keeps its colleagues'
 * delivered contributions (INV-5): they are theirs, not the withheld prose.
 */
export function consultedByOf(speaker: Speaker | undefined, notice: SystemNotice | undefined): Speaker | undefined {
  if (speaker?.kind !== 'system') return speaker;
  const agent = notice?.agent;
  return agent ? { kind: 'agent', agentId: agent.agentId ?? agent.slug, slug: agent.slug, displayName: agent.displayName } : undefined;
}

/** The card body of a `system` notice: its own heading, then the notice text. */
export function systemNoticeBody(text: string): unknown[] {
  return [
    {
      type: 'Container',
      id: 'system-notice',
      style: 'warning',
      items: [
        {
          type: 'TextBlock',
          text: SYSTEM_NOTICE_LABEL,
          weight: 'Bolder',
          size: 'Small',
          isSubtle: true,
          wrap: true,
        },
        { type: 'TextBlock', text, wrap: true, spacing: 'Small' },
      ],
    },
  ];
}

/** The plain-text form of a `system` notice, for the long-answer path. */
export function systemNoticeAsText(text: string): string {
  return `${SYSTEM_NOTICE_LABEL}: ${text}`;
}

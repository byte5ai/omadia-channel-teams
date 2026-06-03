import { CardFactory, type Attachment } from 'botbuilder';

import type { RecalledContext } from '@omadia/plugin-api';

/**
 * Cross-session KG-recall — Teams Adaptive-Card builder.
 *
 * The orchestrator's per-turn probe surfaces relevant Plans / Processes /
 * curated Insights from PRIOR sessions on `SemanticAnswer.recalled`. The
 * web-ui renders this as the "Aus früheren Sessions" card; Teams gets the
 * same content as a read-only Adaptive Card the bot sends just before the
 * answer (context first, then the prose reply).
 *
 * Read-only: no actions. Returns `undefined` when every leg was empty so
 * the caller can splice it in without a guard. Mirrors `buildNudgeCard`
 * (version 1.5, minimal body, hardcoded German labels).
 */
export function buildRecalledContextCard(
  recalled: RecalledContext,
): Attachment | undefined {
  const { plans, processes, insights } = recalled;
  if (plans.length === 0 && processes.length === 0 && insights.length === 0) {
    return undefined;
  }

  const body: Array<Record<string, unknown>> = [
    {
      type: 'TextBlock',
      text: '🧠 Aus früheren Sessions',
      weight: 'Bolder',
      size: 'Small',
      color: 'Accent',
      spacing: 'None',
    },
  ];

  const sectionHeader = (text: string): Record<string, unknown> => ({
    type: 'TextBlock',
    text,
    weight: 'Bolder',
    size: 'Small',
    isSubtle: true,
    spacing: 'Small',
  });
  const line = (text: string): Record<string, unknown> => ({
    type: 'TextBlock',
    text,
    wrap: true,
    size: 'Small',
    spacing: 'None',
  });

  if (plans.length > 0) {
    body.push(sectionHeader('Offene Pläne'));
    for (const p of plans) {
      const label = p.strategy && p.strategy.length > 0 ? p.strategy : 'Plan';
      const open =
        p.openStepGoals.length > 0
          ? ` · offen: ${p.openStepGoals.join('; ')}`
          : '';
      body.push(
        line(`• ${label} (${p.doneCount}/${p.totalCount} Schritte erledigt)${open}`),
      );
    }
  }

  if (processes.length > 0) {
    body.push(sectionHeader('Gespeicherte Prozesse'));
    for (const pr of processes) {
      body.push(line(`• ${pr.title} (${pr.stepCount} Schritte)`));
    }
  }

  if (insights.length > 0) {
    body.push(sectionHeader('Verwandte Erkenntnisse'));
    for (const ins of insights) {
      body.push(line(`• ${ins.kind}: ${ins.summary}`));
    }
  }

  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body,
  });
}

/**
 * Tiny in-memory observer for the Teams conversations the bot has been
 * messaged in. Populated lazily — every inbound activity that lands in
 * `TeamsBot.handleMessage` records the conversation here so the
 * `/operator/channels` dashboard can show it as a bindable target.
 *
 * Why in-memory: a Teams bot in a healthy production deployment sees its
 * conversations again every few minutes. After a restart the operator
 * waits for one message per channel they want to bind — acceptable for
 * v1. A second iteration can persist this to vault / Postgres if the
 * "first message after restart routes to fallback" gap is a problem.
 */

import type { TurnContext } from 'botbuilder';

export interface TeamsConversation {
  /** `activity.conversation.id` — opaque routing key. */
  readonly conversationId: string;
  /** Teams-side conversation type if Bot Framework reported one: usually
   *  `'channel'` (Teams channel), `'personal'` (1:1), `'groupChat'`. */
  readonly conversationType?: string;
  /** Display name for the operator dashboard. Built from `channelData`
   *  when the bot is in a Teams channel:
   *    "#<channelName> in <teamName>"
   *  Falls back to the conversation type + first-12 chars of the id. */
  readonly label: string;
  /** Bot-Framework tenant id when present — same value the kernel uses
   *  for graph-tenant-id (multi-tenant deployments). */
  readonly tenantId?: string;
  /** AAD group id of the owning team (`channelData.team.aadGroupId`) —
   *  Teams-channel conversations only. Lets the Graph resolver fetch the
   *  team roster without parsing the conversation id. */
  readonly teamAadGroupId?: string;
  /** `activity.from.name` for 1:1 chats — the human on the other side.
   *  Makes the DM label self-describing without any Graph call. */
  readonly peerName?: string;
  /** ms since epoch of the most recent inbound. Useful for sorting + for
   *  a future "drop stale entries after N days" sweep. */
  readonly lastSeenAt: number;
}

export class TeamsConversationObserver {
  private readonly seen = new Map<string, TeamsConversation>();

  /** Optional hook fired after every observe() — the plugin wires the
   *  Graph resolver's fire-and-forget prime() here. Must never throw. */
  constructor(
    private readonly onObserved?: (conv: TeamsConversation) => void,
  ) {}

  /**
   * Capture an inbound activity. Idempotent — re-recording an existing
   * conversation just bumps `lastSeenAt` and refreshes the label if
   * the channel/team name changed (rename in Teams).
   */
  observe(context: TurnContext): void {
    const conversationId = context.activity.conversation?.id;
    if (!conversationId) return;

    // `channelData` is Teams-specific and not typed by botbuilder. Cast
    // to a narrow shape locally; missing fields degrade to undefined.
    const channelData = context.activity.channelData as
      | {
          channel?: { id?: string; name?: string };
          team?: { id?: string; name?: string; aadGroupId?: string };
          tenant?: { id?: string };
        }
      | undefined;

    const channelName = channelData?.channel?.name;
    const teamName = channelData?.team?.name;
    const teamAadGroupId = channelData?.team?.aadGroupId;
    const conversationType = context.activity.conversation?.conversationType;
    const peerName =
      conversationType === 'personal'
        ? context.activity.from?.name?.trim() || undefined
        : undefined;

    const label = buildLabel({
      channelName,
      teamName,
      conversationType,
      peerName,
      fallbackId: conversationId,
    });

    const tenantId =
      channelData?.tenant?.id ??
      context.activity.conversation?.tenantId ??
      undefined;

    const entry: TeamsConversation = {
      conversationId,
      ...(conversationType ? { conversationType } : {}),
      label,
      ...(tenantId ? { tenantId } : {}),
      ...(teamAadGroupId ? { teamAadGroupId } : {}),
      ...(peerName ? { peerName } : {}),
      lastSeenAt: Date.now(),
    };
    this.seen.set(conversationId, entry);
    try {
      this.onObserved?.(entry);
    } catch {
      // The hook is best-effort enrichment — it must never break a turn.
    }
  }

  list(): readonly TeamsConversation[] {
    return Array.from(this.seen.values()).sort(
      (a, b) => b.lastSeenAt - a.lastSeenAt,
    );
  }

  /** Diagnostic count — used by activate() to log "N conversations seen". */
  size(): number {
    return this.seen.size;
  }
}

function buildLabel(input: {
  channelName?: string;
  teamName?: string;
  conversationType?: string;
  peerName?: string;
  fallbackId: string;
}): string {
  const { channelName, teamName, conversationType, peerName, fallbackId } =
    input;
  if (channelName && teamName) {
    return `Teams · #${channelName} in ${teamName}`;
  }
  if (channelName) return `Teams · #${channelName}`;
  if (teamName) return `Teams · ${teamName}`;
  if (conversationType === 'personal') {
    if (peerName) return `Teams · DM · ${peerName}`;
    return `Teams · DM (${fallbackId.slice(0, 12)}…)`;
  }
  if (conversationType === 'groupChat') {
    return `Teams · Group chat (${fallbackId.slice(0, 12)}…)`;
  }
  return `Teams · ${fallbackId.slice(0, 16)}…`;
}

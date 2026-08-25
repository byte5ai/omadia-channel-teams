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
  /** Member display names from the Bot-Framework roster (uncapped here;
   *  the directory caps for the payload). Filled asynchronously after
   *  observe() for group conversations — `19:…@thread.skype` group chats
   *  are NOT addressable via Graph `/chats/{id}` (Graph only knows
   *  `@thread.v2` threads and answers 400), so the roster the bot
   *  already fetches per turn is the only member source for them. */
  readonly members?: readonly string[];
  readonly memberCount?: number;
  /** ms since epoch of the most recent inbound. Useful for sorting + for
   *  a future "drop stale entries after N days" sweep. */
  readonly lastSeenAt: number;
}

export interface ObservedMembers {
  readonly names: readonly string[];
  readonly count: number;
}

export class TeamsConversationObserver {
  private readonly seen = new Map<string, TeamsConversation>();

  /** Optional hook fired after every observe() — the plugin wires the
   *  Graph resolver's fire-and-forget prime() here. Must never throw.
   *  `fetchMembers` is called fire-and-forget for group conversations
   *  (roster lookup via the turn's context) — the result is merged into
   *  the stored entry via noteMembers(). */
  constructor(
    private readonly onObserved?: (conv: TeamsConversation) => void,
    private readonly fetchMembers?: (
      context: TurnContext,
    ) => Promise<ObservedMembers | undefined>,
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
    // Named group chats sometimes carry their topic as
    // `conversation.name` on the activity — the only Bot-Framework
    // source for it, since `@thread.skype` threads are invisible to
    // Graph. Unnamed groups leave it empty.
    const groupName =
      conversationType === 'groupChat'
        ? context.activity.conversation?.name?.trim() || undefined
        : undefined;
    const peerName =
      conversationType === 'personal'
        ? context.activity.from?.name?.trim() || undefined
        : undefined;

    const previousEntry = this.seen.get(conversationId);
    let label = buildLabel({
      channelName,
      teamName,
      conversationType,
      groupName,
      peerName,
      fallbackId: conversationId,
    });
    // A re-observe of an unnamed group chat must not regress a
    // roster-derived label back to the opaque-id fallback while the
    // async roster refresh is still in flight.
    if (
      previousEntry &&
      isGroupIdFallbackLabel(label) &&
      !isGroupIdFallbackLabel(previousEntry.label)
    ) {
      label = previousEntry.label;
    }

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
      // Keep the last known roster across re-observes — it refreshes
      // asynchronously below and must not flicker away in between.
      ...(previousEntry?.members !== undefined
        ? { members: previousEntry.members }
        : {}),
      ...(previousEntry?.memberCount !== undefined
        ? { memberCount: previousEntry.memberCount }
        : {}),
      lastSeenAt: Date.now(),
    };
    this.seen.set(conversationId, entry);
    try {
      this.onObserved?.(entry);
    } catch {
      // The hook is best-effort enrichment — it must never break a turn.
    }
    if (
      this.fetchMembers &&
      (conversationType === 'groupChat' || conversationType === 'channel')
    ) {
      void this.fetchMembers(context)
        .then((members) => {
          if (members && members.names.length > 0) {
            this.noteMembers(conversationId, members);
          }
        })
        .catch(() => {
          // Roster enrichment is best-effort; keep the entry as-is.
        });
    }
  }

  /** Merge roster names into an already-observed conversation. When the
   *  label is still the opaque-id fallback (unnamed group chat), derive
   *  a human-readable one from the members — mirroring how Teams itself
   *  titles unnamed groups by their participant list. */
  noteMembers(conversationId: string, members: ObservedMembers): void {
    const current = this.seen.get(conversationId);
    if (!current) return;
    const label =
      current.conversationType === 'groupChat' &&
      isGroupIdFallbackLabel(current.label)
        ? buildGroupMembersLabel(members)
        : current.label;
    this.seen.set(conversationId, {
      ...current,
      label,
      members: members.names,
      memberCount: members.count,
    });
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
  groupName?: string;
  peerName?: string;
  fallbackId: string;
}): string {
  const {
    channelName,
    teamName,
    conversationType,
    groupName,
    peerName,
    fallbackId,
  } = input;
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
    if (groupName) return `Teams · ${groupName}`;
    return `Teams · Group chat (${fallbackId.slice(0, 12)}…)`;
  }
  return `Teams · ${fallbackId.slice(0, 16)}…`;
}

/** Label for an unnamed group chat: first names of up to three members,
 *  "+N" for the rest — e.g. "Teams · Group chat: Teresita, Christian +2".
 *  Mirrors how Teams itself titles unnamed groups by participant list. */
function buildGroupMembersLabel(members: ObservedMembers): string {
  const firstNames = members.names
    .map((n) => n.split(/[\s,-]+/)[0] ?? n)
    .filter((n) => n.length > 0);
  const shown = firstNames.slice(0, 3);
  const rest = members.count - shown.length;
  const suffix = rest > 0 ? ` +${String(rest)}` : '';
  return `Teams · Group chat: ${shown.join(', ')}${suffix}`;
}

/** True for the opaque-id fallback label of an unnamed group chat. */
function isGroupIdFallbackLabel(label: string): boolean {
  return label.startsWith('Teams · Group chat (');
}

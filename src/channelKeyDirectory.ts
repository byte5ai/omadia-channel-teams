import type {
  ChannelKeyDirectory,
  ChannelKeyEntry,
} from '@omadia/channel-sdk';

import type { TeamsConversationObserver } from './teamsConversationObserver.js';

const PLUGIN_ID = '@omadia/channel-teams';

/**
 * Build the directory contribution for the operator `/operator/channels`
 * dashboard. Composes:
 *
 *   1. A bot-level catch-all entry  — `key='28:<app_id>'`. Routes every
 *      Teams turn whose conversation has no specific binding. This is
 *      the default "all of Teams → Agent X" lever.
 *   2. One entry per observed conversation — `key='<conversation.id>'`.
 *      Lets the operator bind individual Teams channels / DMs / group
 *      chats to different Agents. Discovery is lazy: a conversation
 *      appears here as soon as the bot has been messaged in it once
 *      (after restart this list is empty until the first message lands).
 *
 * The runtime resolver in `TeamsBot.resolveChatAgentForTurn` tries the
 * conversation-id binding first, then the bot-level binding, then the
 * platform fallback. The keys we emit here are exactly the keys the
 * resolver looks up — what the operator picks IS what routes at
 * runtime.
 */
export function buildTeamsChannelKeyDirectory(opts: {
  readonly microsoftAppId: string;
  readonly microsoftTenantId: string;
  /** Optional operator-set label override for the bot-level entry,
   *  e.g. "Production · Marketing". Falls back to a generated label. */
  readonly displayLabel?: string;
  /** Observer that records every inbound conversation. Asked at
   *  `listKeys()` time for the live set — the kernel calls listKeys()
   *  once per `/operator/channels` page render. */
  readonly conversationObserver: TeamsConversationObserver;
}): ChannelKeyDirectory {
  const botKey = `28:${opts.microsoftAppId}`;
  const shortApp =
    opts.microsoftAppId.length > 12
      ? `${opts.microsoftAppId.slice(0, 8)}…`
      : opts.microsoftAppId;
  const botLabel = opts.displayLabel?.trim() || `Teams · ${shortApp} (all)`;
  const tenantHint =
    opts.microsoftTenantId.length > 12
      ? `tenant ${opts.microsoftTenantId.slice(0, 8)}…`
      : `tenant ${opts.microsoftTenantId}`;

  return {
    channelType: 'teams',
    originPluginId: PLUGIN_ID,
    async listKeys(): Promise<readonly ChannelKeyEntry[]> {
      const entries: ChannelKeyEntry[] = [
        {
          key: botKey,
          label: botLabel,
          hint: `${tenantHint} · catch-all`,
        },
      ];
      for (const conv of opts.conversationObserver.list()) {
        entries.push({
          key: conv.conversationId,
          label: conv.label,
          hint:
            conv.conversationType === 'channel'
              ? 'Teams channel'
              : conv.conversationType === 'personal'
                ? '1:1 chat'
                : conv.conversationType === 'groupChat'
                  ? 'Group chat'
                  : 'conversation',
        });
      }
      return entries;
    },
  };
}

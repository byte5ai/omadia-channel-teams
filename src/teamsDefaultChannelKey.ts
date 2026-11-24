/**
 * The deployment's own Teams bot, as the kernel's Conductor knows it.
 *
 * The kernel lets exactly one unknown bot key per channel route to a binding
 * or the fallback agent: the channel's default bot. Other channels answer it
 * through `agentIdentities@1.defaultChannelKey`; Teams may not register an
 * identity provider (its agent bots are provisioned by the kernel into
 * `agent_teams_identities`), so it answers through the narrow
 * `registerDefaultChannelKeyProvider` seam instead.
 *
 * Which bot: the Microsoft 365 app registration — the plugin's
 * `microsoft_app_id`, else env `MICROSOFT_APP_ID` (the order the kernel used)
 * — and only when this plugin serves it. Never an operator-listed agent bot:
 * the default bot is the one bot an unknown key may route to the fallback
 * agent, so a guess would widen routing.
 */
import type { DefaultChannelKeyProvider } from './kernel-types.js';
import { teamsBotKey } from './teamsBotIdentity.js';
import type { TeamsBotIdentity } from './teamsBotIdentity.js';

export function resolveTeamsDefaultChannelKey(
  appIdCandidates: readonly (string | undefined)[],
  servedBots: readonly TeamsBotIdentity[],
): string | undefined {
  for (const candidate of appIdCandidates) {
    const wanted = candidate?.trim().toLowerCase();
    if (wanted === undefined || wanted === '') continue;
    const bot = servedBots.find((served) => served.appId.trim().toLowerCase() === wanted);
    return bot ? teamsBotKey(bot.appId.trim()) : undefined;
  }
  return undefined;
}

/**
 * The provider registered with the kernel. `readAppIdCandidates` is read per
 * call, so the answer follows the plugin's config the same way the kernel's
 * old per-call read did.
 */
export function createTeamsDefaultChannelKeyProvider(
  readAppIdCandidates: () => readonly (string | undefined)[],
  servedBots: readonly TeamsBotIdentity[],
): DefaultChannelKeyProvider {
  return {
    channelType: 'teams',
    defaultChannelKey: () => resolveTeamsDefaultChannelKey(readAppIdCandidates(), servedBots),
  };
}

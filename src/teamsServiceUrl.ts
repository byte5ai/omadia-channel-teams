// Where a bot may present its Bot Framework credentials. Kept free of
// botbuilder imports so every module can use it.

/** Same read-path allowlist as the conversation-ref store (`isAllowedServiceUrl`,
 *  review M2): a poisoned reference — DB row, routine row, or cache — must not
 *  make `continueConversationAsync` present bot credentials to an attacker
 *  host. With N bots a poisoned `bot.id` could additionally pick ANOTHER bot's
 *  credentials, so the guard sits directly in front of every proactive turn,
 *  not only behind the ref store's loads. */
export function isAllowedProactiveServiceUrl(serviceUrl: unknown): boolean {
  if (typeof serviceUrl !== 'string') return false;
  try {
    const host = new URL(serviceUrl).hostname.toLowerCase();
    return host === 'smba.trafficmanager.net' || host.endsWith('.botframework.com');
  } catch {
    return false;
  }
}

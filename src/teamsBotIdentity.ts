/**
 * Per-bot identity for the multi-bot Teams channel (epic byte5ai/omadia#860,
 * wave W0a).
 *
 * One omadia agent = one Entra app registration = one Azure bot = one entry
 * in the `teams_bots[]` plugin config. This module owns exactly three
 * things:
 *
 *   1. The {@link TeamsBotIdentity} shape the config-wiring layer maps the
 *      `teams_bots[]` list (and the legacy scalar-credential env shim) onto.
 *   2. Pure lookup helpers over that list — by `botSlug` for the
 *      `/api/teams/:botSlug/messages` route, by `appId` for the turn
 *      resolver, plus the default-bot accessor the legacy
 *      `/api/teams/messages` alias needs.
 *   3. The `28:<appId>` bot-key helper. That key is what
 *      `buildTeamsChannelKeyDirectory` publishes as the bot-level catch-all
 *      row AND what `TeamsBot.resolveChatAgentForTurn` matches against
 *      `activity.recipient.id` — it must be built in exactly one place so
 *      directory rows and runtime resolution can never drift apart.
 *
 * Deliberately dependency-free: no config reading, no env access, no
 * logging. Populating the bot list (env shim, manifest config schema) is
 * the config-wiring unit's job.
 */

/**
 * Entra app registration type — PER BOT, not process-global.
 *
 * The legacy plugin read one `MICROSOFT_APP_TYPE` from `process.env` for
 * the whole process; with multiple bots each identity carries its own
 * (e.g. an old MultiTenant registration next to newly provisioned
 * SingleTenant ones).
 */
export type TeamsBotAppType = 'MultiTenant' | 'SingleTenant' | 'UserAssignedMSI';

/**
 * App type for NEWLY provisioned bots when none is configured.
 *
 * SingleTenant — Azure deprecated creating new MultiTenant app
 * registrations in 07/2025. Only the legacy env shim may still map an
 * existing deployment's `MICROSOFT_APP_TYPE` (historically defaulted to
 * MultiTenant) onto `teams_bots[0]`; every other bot defaults to this.
 */
export const DEFAULT_TEAMS_BOT_APP_TYPE: TeamsBotAppType = 'SingleTenant';

/**
 * One named bot identity in a Teams tenant.
 *
 * SECURITY: `appPasswordSecretRef` is the NAME of the secret holding the
 * Bot Framework app password (resolved via `ctx.secrets.require(...)` at
 * activate() time, cf. manifest `auth_from: microsoft_app_password`) —
 * never the password value itself. Do not widen this type to carry the
 * resolved secret, and keep `appId` / secret refs out of log lines
 * (use {@link teamsBotLogLabel}).
 */
export interface TeamsBotIdentity {
  /** Stable, URL-safe identifier — path segment of the per-bot messaging
   *  route `/api/teams/:botSlug/messages`. Unique within the list. */
  readonly botSlug: string;
  /** Entra application (client) id of the bot's app registration. */
  readonly appId: string;
  /** Entra tenant the app registration lives in. */
  readonly tenantId: string;
  /** Secret NAME (reference) of the bot's app password — see above. */
  readonly appPasswordSecretRef: string;
  /** Per-bot app registration type — see {@link TeamsBotAppType}. */
  readonly appType: TeamsBotAppType;
  /** Human-readable name — the channel-directory catch-all row label and
   *  the name operators see in `/operator/channels`. */
  readonly displayName: string;
}

/** Bot-Framework id prefix for bot identities (`28:` + Entra app id). */
export const TEAMS_BOT_KEY_PREFIX = '28:';

/**
 * The Bot-Framework identity key of a bot: `28:<appId>`.
 *
 * Single source of truth for the key that (a) the channel-key directory
 * publishes as the bot-level catch-all row and (b) the turn resolver sees
 * as `activity.recipient.id`. Exact string equality is what routes — never
 * rebuild this inline.
 */
export function teamsBotKey(appId: string): string {
  return `${TEAMS_BOT_KEY_PREFIX}${appId}`;
}

/**
 * Inverse of {@link teamsBotKey}: extract the app id from a `28:<appId>`
 * bot key. Returns `undefined` for anything that is not a bot identity key
 * (user ids `29:…`, conversation ids, empty app id).
 */
export function parseTeamsBotKey(key: string): string | undefined {
  if (!key.startsWith(TEAMS_BOT_KEY_PREFIX)) return undefined;
  const appId = key.slice(TEAMS_BOT_KEY_PREFIX.length);
  return appId.length > 0 ? appId : undefined;
}

/**
 * Look up a bot by its route slug (`/api/teams/:botSlug/messages`).
 * Exact match — slugs are operator-chosen identifiers, not GUIDs.
 */
export function findTeamsBotBySlug(
  bots: readonly TeamsBotIdentity[],
  botSlug: string,
): TeamsBotIdentity | undefined {
  return bots.find((bot) => bot.botSlug === botSlug);
}

/**
 * Look up a bot by Entra app id — the turn resolver's lookup after
 * stripping the `28:` prefix off `activity.recipient.id`.
 *
 * GUID comparison is case-insensitive: Azure serialises app ids
 * lowercase, but operator-pasted config values may differ in casing.
 */
export function findTeamsBotByAppId(
  bots: readonly TeamsBotIdentity[],
  appId: string,
): TeamsBotIdentity | undefined {
  const wanted = appId.toLowerCase();
  return bots.find((bot) => bot.appId.toLowerCase() === wanted);
}

/**
 * The default bot — first entry of the configured list. This is the bot
 * the legacy `/api/teams/messages` alias (no `:botSlug` segment) serves,
 * and the one the env shim maps the scalar `MICROSOFT_APP_*` credentials
 * onto (`teams_bots[0]`). `undefined` when the list is empty — the caller
 * decides whether that is a hard activation error.
 */
export function getDefaultTeamsBot(
  bots: readonly TeamsBotIdentity[],
): TeamsBotIdentity | undefined {
  return bots[0];
}

/**
 * Log-safe label for a bot: slug + display name only. Use this in log
 * lines instead of interpolating the identity — it can never leak the
 * `appId` or `appPasswordSecretRef`.
 */
export function teamsBotLogLabel(bot: TeamsBotIdentity): string {
  return `${bot.botSlug} ("${bot.displayName}")`;
}

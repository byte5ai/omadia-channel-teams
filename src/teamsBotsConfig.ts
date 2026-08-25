/**
 * `teams_bots[]` config surface + legacy scalar-credential shim (#860 W0a,
 * config-wiring unit).
 *
 * The plugin historically read exactly one credential set:
 *
 *   ctx.config.require('microsoft_app_id')
 *   ctx.config.require('microsoft_tenant_id')
 *   await ctx.secrets.require('microsoft_app_password')
 *   process.env.MICROSOFT_APP_TYPE          (default 'MultiTenant')
 *
 * Multi-bot replaces that with a `teams_bots` list of
 * {@link TeamsBotIdentity} entries. This module owns the two pure mapping
 * steps:
 *
 *   1. {@link parseTeamsBotsConfig} — validate the raw `teams_bots` config
 *      value (a real array when set programmatically in the install
 *      registry, or a JSON string when typed into the setup wizard's
 *      string field) into a `TeamsBotIdentity[]`.
 *   2. {@link legacyTeamsBotFromScalars} — map the legacy scalar
 *      credentials onto `teams_bots[0]` so every existing single-bot
 *      deployment keeps working with zero config changes.
 *
 * SECURITY: entries carry `appPasswordSecretRef` — the NAME of a secret in
 * this plugin's vault namespace, resolved via `ctx.secrets.require(...)` by
 * the wiring layer. Passwords are never accepted inline in config; a raw
 * `appPassword` key in an entry is rejected loudly so an operator cannot
 * accidentally persist a credential into the (non-secret) config store.
 *
 * Deliberately side-effect-free: no ctx access, no env reads, no logging —
 * plugin.ts feeds the raw values in, tests feed fixtures.
 */

import {
  DEFAULT_TEAMS_BOT_APP_TYPE,
  type TeamsBotAppType,
  type TeamsBotIdentity,
} from './teamsBotIdentity.js';

/**
 * Deterministic `botSlug` of the shimmed legacy bot. The per-bot route
 * `/api/teams/default/messages`, the `/api/messages` + `/api/teams/messages`
 * default-bot aliases, and the one-shot `bot_app_id` ref-store backfill all
 * derive from `teams_bots[0]` — pinning the slug keeps every one of those
 * surfaces agreeing on the same identity across restarts.
 */
export const LEGACY_TEAMS_BOT_SLUG = 'default';

/** Secret name the legacy scalar deployment stores its app password under
 *  (the `microsoft_app_password` vault key the plugin has always read). */
export const LEGACY_TEAMS_BOT_SECRET_REF = 'microsoft_app_password';

/** Display name of the shimmed legacy bot when the operator set no
 *  `teams_directory_label`. */
export const LEGACY_TEAMS_BOT_DISPLAY_NAME = 'Microsoft Teams Bot';

/** `botSlug` is a URL path segment of `/api/teams/:botSlug/messages` —
 *  constrain it to unambiguous, log-safe characters. */
const BOT_SLUG_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$/;

const APP_TYPES: readonly TeamsBotAppType[] = [
  'MultiTenant',
  'SingleTenant',
  'UserAssignedMSI',
];

function isTeamsBotAppType(value: unknown): value is TeamsBotAppType {
  return typeof value === 'string' && (APP_TYPES as readonly string[]).includes(value);
}

function asTrimmedString(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim().length > 0 ? value.trim() : undefined;
}

/** Config-shape error. Thrown at activate() time — a malformed bot list is
 *  an operator mistake that must fail loudly, never a silently-empty list. */
export class TeamsBotsConfigError extends Error {
  constructor(message: string) {
    super(`teams_bots config invalid: ${message}`);
    this.name = 'TeamsBotsConfigError';
  }
}

function parseEntry(raw: unknown, index: number): TeamsBotIdentity {
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
    throw new TeamsBotsConfigError(`entry ${String(index)} must be an object`);
  }
  const entry = raw as Record<string, unknown>;

  // Loud inline-credential rejection — see module doc.
  if ('appPassword' in entry || 'app_password' in entry) {
    throw new TeamsBotsConfigError(
      `entry ${String(index)} carries an inline app password — store the secret in the vault and reference it via appPasswordSecretRef`,
    );
  }

  const botSlug = asTrimmedString(entry['botSlug']);
  if (!botSlug) {
    throw new TeamsBotsConfigError(`entry ${String(index)} is missing botSlug`);
  }
  if (!BOT_SLUG_PATTERN.test(botSlug)) {
    throw new TeamsBotsConfigError(
      `entry ${String(index)} has an invalid botSlug '${botSlug}' — use 1-64 letters, digits, '.', '_' or '-', starting alphanumeric`,
    );
  }
  const appId = asTrimmedString(entry['appId']);
  if (!appId) {
    throw new TeamsBotsConfigError(`entry '${botSlug}' is missing appId`);
  }
  const tenantId = asTrimmedString(entry['tenantId']);
  if (!tenantId) {
    throw new TeamsBotsConfigError(`entry '${botSlug}' is missing tenantId`);
  }
  const appPasswordSecretRef = asTrimmedString(entry['appPasswordSecretRef']);
  if (!appPasswordSecretRef) {
    throw new TeamsBotsConfigError(`entry '${botSlug}' is missing appPasswordSecretRef`);
  }

  const rawAppType = entry['appType'];
  let appType: TeamsBotAppType;
  if (rawAppType === undefined || rawAppType === null || rawAppType === '') {
    // New bots default SingleTenant — Azure deprecated creating MultiTenant
    // app registrations 07/2025. Only the legacy shim below may still
    // default to MultiTenant, and only for teams_bots[0].
    appType = DEFAULT_TEAMS_BOT_APP_TYPE;
  } else if (isTeamsBotAppType(rawAppType)) {
    appType = rawAppType;
  } else {
    throw new TeamsBotsConfigError(
      `entry '${botSlug}' has an invalid appType — expected MultiTenant | SingleTenant | UserAssignedMSI`,
    );
  }

  return {
    botSlug,
    appId,
    tenantId,
    appPasswordSecretRef,
    appType,
    displayName: asTrimmedString(entry['displayName']) ?? botSlug,
  };
}

/**
 * Validate the raw `teams_bots` config value into a `TeamsBotIdentity[]`.
 *
 * Accepts:
 *   - `undefined` / `null` / `''` / `[]` → `[]` (caller falls back to the
 *     legacy scalar shim),
 *   - an array of entry objects (install-registry / agent-config value),
 *   - a JSON string encoding that array (the setup wizard's string field).
 *
 * Entry shape: `{ botSlug, appId, tenantId, appPasswordSecretRef,
 * appType?, displayName? }`. `appType` defaults to
 * {@link DEFAULT_TEAMS_BOT_APP_TYPE} (SingleTenant); `displayName` defaults
 * to the slug. `teams_bots[0]` is the DEFAULT bot (legacy route aliases,
 * ref-store backfill target).
 *
 * @throws TeamsBotsConfigError on any malformed, duplicate, or
 *   inline-credential entry — never silently drops entries.
 */
export function parseTeamsBotsConfig(raw: unknown): TeamsBotIdentity[] {
  let value: unknown = raw;
  if (typeof value === 'string') {
    if (value.trim().length === 0) return [];
    try {
      value = JSON.parse(value);
    } catch {
      throw new TeamsBotsConfigError('value is a string but not valid JSON');
    }
  }
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value)) {
    throw new TeamsBotsConfigError('value must be a JSON array of bot entries');
  }

  const bots = value.map((entry, index) => parseEntry(entry, index));

  const seenSlugs = new Set<string>();
  const seenAppIds = new Set<string>();
  for (const bot of bots) {
    if (seenSlugs.has(bot.botSlug)) {
      throw new TeamsBotsConfigError(`duplicate botSlug '${bot.botSlug}'`);
    }
    seenSlugs.add(bot.botSlug);
    const appIdKey = bot.appId.toLowerCase();
    if (seenAppIds.has(appIdKey)) {
      throw new TeamsBotsConfigError(
        `duplicate appId (second entry: '${bot.botSlug}')`,
      );
    }
    seenAppIds.add(appIdKey);
  }
  return bots;
}

/** Inputs of the legacy scalar-credential shim — the values plugin.ts has
 *  always read (config keys from the microsoft365 integration chain, the
 *  process-global `MICROSOFT_APP_TYPE` env knob). */
export interface LegacyTeamsBotScalars {
  /** `ctx.config.require('microsoft_app_id')`. */
  readonly appId: string;
  /** `ctx.config.require('microsoft_tenant_id')`. */
  readonly tenantId: string;
  /** `process.env.MICROSOFT_APP_TYPE` — raw, unvalidated. */
  readonly appTypeEnv?: string | undefined;
  /** Operator label (`teams_directory_label`) reused as the shimmed bot's
   *  display name when present. */
  readonly displayName?: string | undefined;
}

/**
 * Map the legacy scalar credentials onto `teams_bots[0]`.
 *
 * appType keeps the historical `readTeamsConfigFromEnv` semantics for THIS
 * bot only: `MICROSOFT_APP_TYPE` of SingleTenant / UserAssignedMSI is
 * honoured, anything else (unset, unknown) falls back to MultiTenant —
 * existing single-bot deployments were provisioned under that default and
 * must keep authenticating unchanged. New bots (list entries above) default
 * SingleTenant instead.
 */
export function legacyTeamsBotFromScalars(
  scalars: LegacyTeamsBotScalars,
): TeamsBotIdentity {
  const appType: TeamsBotAppType =
    scalars.appTypeEnv === 'SingleTenant' || scalars.appTypeEnv === 'UserAssignedMSI'
      ? scalars.appTypeEnv
      : 'MultiTenant';
  return {
    botSlug: LEGACY_TEAMS_BOT_SLUG,
    appId: scalars.appId,
    tenantId: scalars.tenantId,
    appPasswordSecretRef: LEGACY_TEAMS_BOT_SECRET_REF,
    appType,
    displayName:
      asTrimmedString(scalars.displayName) ?? LEGACY_TEAMS_BOT_DISPLAY_NAME,
  };
}

/**
 * `teams_agent_apps[]` config surface (#860 W2, auto-invite wave).
 *
 * Each entry links one omadia agent to the Teams app package that gives it
 * a named, separately @mention-able identity inside a team. The W2
 * installer (`teamsAgentInstaller`) consumes this list to look the app up
 * in the org catalog and auto-install it via Graph
 * `POST /teams/{id}/installedApps`; the onboarding hook uses `displayName`
 * for the result / fallback cards.
 *
 * This module owns exactly one pure mapping step:
 * {@link parseTeamsAgentAppsConfig} — validate the raw `teams_agent_apps`
 * config value (a real array when set programmatically in the install
 * registry, or a JSON string when typed into the setup wizard's string
 * field, mirroring `teams_bots`) into a `TeamsAgentApp[]`.
 *
 * NOT identity, NOT routing: `agentSlug` names the omadia agent (conductor
 * slug) for labels and status only — it is unrelated to
 * `teams_bots[].botSlug` and drives no message routing in this wave.
 *
 * SECURITY: entries carry only PUBLIC identifiers (Teams app ids appear in
 * `https://teams.microsoft.com/l/app/<appId>` deep links) — never secrets.
 * Bot credentials stay in `teams_bots[]` / the vault.
 *
 * Deliberately side-effect-free: no ctx access, no env reads, no logging —
 * plugin.ts feeds the raw value in, tests feed fixtures (precedent:
 * `teamsBotsConfig.ts`).
 */

/**
 * One agent ↔ Teams-app link.
 *
 * All ids here are public app identifiers — safe for logs, cards, and
 * deep-link URLs.
 */
export interface TeamsAgentApp {
  /** Omadia agent (conductor) slug — label/status surface only, NOT a
   *  `teams_bots[].botSlug` and never used for routing. */
  readonly agentSlug: string;
  /** `id` from the generated Teams app manifest (`externalId` in the Graph
   *  `teamsApp` resource) — the org-catalog lookup key when
   *  {@link TeamsAgentApp.teamsAppId} is not configured. */
  readonly teamsAppExternalId: string;
  /** Graph catalog id of the published `teamsApp`, when already known.
   *  Configuring it skips the catalog lookup entirely. */
  readonly teamsAppId?: string | undefined;
  /** Human-readable agent name for result / fallback cards. Defaults to
   *  {@link TeamsAgentApp.agentSlug}. */
  readonly displayName: string;
}

/** `agentSlug` must be a valid conductor slug — same rule the middleware
 *  install registry enforces (`harness-orchestrator` configStore `SLUG_RE`):
 *  lowercase letters, digits and inner dashes, max 64 chars. */
const AGENT_SLUG_PATTERN = /^[a-z0-9](?:[a-z0-9-]{0,62}[a-z0-9])?$/;

/** Config-shape error. Thrown at activate() time — a malformed app list is
 *  an operator mistake that must fail loudly, never a silently-empty list
 *  (precedent: `TeamsBotsConfigError`). */
export class TeamsAgentAppsConfigError extends Error {
  constructor(message: string) {
    super(`teams_agent_apps config invalid: ${message}`);
    this.name = 'TeamsAgentAppsConfigError';
  }
}

function asTrimmedString(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim().length > 0 ? value.trim() : undefined;
}

function parseEntry(raw: unknown, index: number): TeamsAgentApp {
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
    throw new TeamsAgentAppsConfigError(`entry ${String(index)} must be an object`);
  }
  const entry = raw as Record<string, unknown>;

  const agentSlug = asTrimmedString(entry['agentSlug']);
  if (!agentSlug) {
    throw new TeamsAgentAppsConfigError(`entry ${String(index)} is missing agentSlug`);
  }
  if (!AGENT_SLUG_PATTERN.test(agentSlug)) {
    throw new TeamsAgentAppsConfigError(
      `entry ${String(index)} has an invalid agentSlug '${agentSlug}' — use 1-64 lowercase letters, digits or inner '-', starting and ending with a letter or digit`,
    );
  }

  const teamsAppExternalId = asTrimmedString(entry['teamsAppExternalId']);
  if (!teamsAppExternalId) {
    throw new TeamsAgentAppsConfigError(
      `entry '${agentSlug}' is missing teamsAppExternalId`,
    );
  }

  const rawTeamsAppId = entry['teamsAppId'];
  let teamsAppId: string | undefined;
  if (rawTeamsAppId === undefined || rawTeamsAppId === null || rawTeamsAppId === '') {
    teamsAppId = undefined;
  } else {
    teamsAppId = asTrimmedString(rawTeamsAppId);
    if (!teamsAppId) {
      throw new TeamsAgentAppsConfigError(
        `entry '${agentSlug}' has a non-string teamsAppId — omit it or configure the Graph catalog id`,
      );
    }
  }

  return {
    agentSlug,
    teamsAppExternalId,
    teamsAppId,
    displayName: asTrimmedString(entry['displayName']) ?? agentSlug,
  };
}

/**
 * Validate the raw `teams_agent_apps` config value into a
 * `TeamsAgentApp[]`.
 *
 * Accepts:
 *   - `undefined` / `null` / `''` / `[]` → `[]` (auto-invite feature OFF —
 *     deployments without the field keep behaving exactly as before),
 *   - an array of entry objects (install-registry / agent-config value),
 *   - a JSON string encoding that array (the setup wizard's string field).
 *
 * Entry shape: `{ agentSlug, teamsAppExternalId, teamsAppId?,
 * displayName? }`. `displayName` defaults to the slug.
 *
 * @throws TeamsAgentAppsConfigError on any malformed or duplicate entry —
 *   never silently drops entries.
 */
export function parseTeamsAgentAppsConfig(raw: unknown): TeamsAgentApp[] {
  let value: unknown = raw;
  if (typeof value === 'string') {
    if (value.trim().length === 0) return [];
    try {
      value = JSON.parse(value);
    } catch {
      throw new TeamsAgentAppsConfigError('value is a string but not valid JSON');
    }
  }
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value)) {
    throw new TeamsAgentAppsConfigError('value must be a JSON array of agent-app entries');
  }

  const apps = value.map((entry, index) => parseEntry(entry, index));

  const seenSlugs = new Set<string>();
  const seenExternalIds = new Set<string>();
  const seenAppIds = new Set<string>();
  for (const app of apps) {
    if (seenSlugs.has(app.agentSlug)) {
      throw new TeamsAgentAppsConfigError(`duplicate agentSlug '${app.agentSlug}'`);
    }
    seenSlugs.add(app.agentSlug);

    // App ids are GUIDs Graph serialises lowercase — compare case-insensitively
    // (same policy as the teams_bots appId dedupe).
    const externalIdKey = app.teamsAppExternalId.toLowerCase();
    if (seenExternalIds.has(externalIdKey)) {
      throw new TeamsAgentAppsConfigError(
        `duplicate teamsAppExternalId (second entry: '${app.agentSlug}')`,
      );
    }
    seenExternalIds.add(externalIdKey);

    if (app.teamsAppId !== undefined) {
      const appIdKey = app.teamsAppId.toLowerCase();
      if (seenAppIds.has(appIdKey)) {
        throw new TeamsAgentAppsConfigError(
          `duplicate teamsAppId (second entry: '${app.agentSlug}')`,
        );
      }
      seenAppIds.add(appIdKey);
    }
  }
  return apps;
}

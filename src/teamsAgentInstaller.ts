/**
 * `teamsAgentInstaller` — auto-install agent Teams apps into a team via the
 * M365 connector's `teamsProvisioner@1` capability (#860 W2, issue #20).
 *
 * For every configured agent app (`teams_agent_apps[]`, parsed by
 * `teamsAgentApps.ts`) the installer resolves the org-catalog app and runs
 * the provisioner's `installToTeam` (Graph `POST /teams/{id}/installedApps`).
 * It NEVER reimplements Graph calls — all network traffic stays inside the
 * connector.
 *
 * Catalog resolution order (coordinator override on issue #20):
 *   1. configured `teamsAppId` — skips the lookup entirely,
 *   2. `getCatalogApp({ teamsAppExternalId })` — FEATURE-DETECTED (connector
 *      >= 0.3.1); never required,
 *   3. neither → the "app not in catalog" fallback-card outcome.
 *
 * Error mapping — the message path never sees a throw from here:
 *   * `ConsentMissingError` (403) → typed `'fallback'` outcome carrying
 *     `missingScopes` for the deep-link consent card (issue #21), plus a
 *     per-tenant NEGATIVE CACHE entry so a consent-missing tenant is not
 *     hammered on every `membersAdded` (no retry storm),
 *   * `Idempotent.outcome === 'already-existed'` (409) → success,
 *   * `ProvisioningThrottledError` (429 budget exhausted connector-side) →
 *     bounded retry honouring `retryAfterSeconds`, then a typed `'failed'`
 *     outcome,
 *   * anything else → typed `'failed'` outcome, logged without ids.
 *
 * INTRO THROTTLE — freshly installed agent bots each receive `membersAdded`,
 * so a flock install would post N identical intros. This module owns the
 * marker primitive only: {@link TeamsAutoInstallMarker} records "team X was
 * just auto-installed" for a short TTL; the onboarding hook (issue #21,
 * another slice) probes it to suppress the default intro of every bot except
 * the one that ran the installer.
 *
 * PROCESS-LOCAL STATE, EXPLICITLY: both the consent negative cache and the
 * auto-install marker are in-memory. This plugin cannot ship a DB migration
 * (`teamsConversationRefStore.ts`: "The kernel owns the schema"), so with
 * more than one middleware instance a second instance can still retry a
 * consent-missing tenant or post a duplicate intro. The TTLs bound that
 * blast radius — they are deliberately NOT pretending to be a global lock.
 *
 * SECURITY: log lines carry `agentSlug`/`displayName` labels only — app ids
 * and scope names stay out of logs (W0a `teamsBotLogLabel` precedent). The
 * typed outcomes DO carry ids/scopes: they feed cards, not logs.
 */

import type {
  ConsentMissingErrorShape,
  GetCatalogAppResult,
  IdempotentOutcome,
  ProvisioningThrottledErrorShape,
  TeamsProvisionerAccessor,
} from './kernel-types.js';

/**
 * One agent ↔ Teams-app link to install. Structural twin of the
 * `TeamsAgentApp` config entry (`teamsAgentApps.ts`, sibling W2 unit) —
 * declared here as well so this module compiles standalone; TypeScript's
 * structural typing lets the wiring layer pass parsed config entries
 * directly.
 */
export interface TeamsAgentAppTarget {
  /** Omadia agent (conductor) slug — label/status surface only, NOT a
   *  `teams_bots[].botSlug`; drives no routing in this wave. */
  readonly agentSlug: string;
  /** Manifest id (`externalId`) — catalog lookup key when
   *  {@link TeamsAgentAppTarget.teamsAppId} is not configured. */
  readonly teamsAppExternalId: string;
  /** Graph catalog id, when already known. Skips the lookup entirely. */
  readonly teamsAppId?: string | undefined;
  /** Human-readable agent name for result/fallback cards. */
  readonly displayName: string;
}

/** Why an app ended on the deep-link fallback card instead of installed. */
export type AgentAppFallbackReason =
  /** Graph/ARM 403 — admin consent / application permission missing. */
  | 'consent-missing'
  /** Skipped without a Graph call — the tenant hit 403 within the TTL. */
  | 'consent-cached'
  /** No configured `teamsAppId`, and the lookup (or its absence) found
   *  no catalog app for `teamsAppExternalId`. */
  | 'not-in-catalog';

/** Terminal failure reasons (retry exhausted / unexpected error). */
export type AgentAppFailureReason = 'throttled' | 'error';

/** Per-app outcome of one installer run. Discriminate on `kind`. */
export type AgentAppInstallOutcome =
  | {
      readonly kind: 'installed';
      readonly agentSlug: string;
      readonly displayName: string;
      /** Catalog id the install referenced — deep-link material. */
      readonly teamsAppId: string;
      /** `'created'` = fresh Graph install; `'already-existed'` = 409
       *  idempotent success (the app was in the team all along). */
      readonly outcome: IdempotentOutcome;
    }
  | {
      readonly kind: 'fallback';
      readonly agentSlug: string;
      readonly displayName: string;
      readonly reason: AgentAppFallbackReason;
      readonly teamsAppExternalId: string;
      /** Catalog id when known — lets the card render the
       *  `https://teams.microsoft.com/l/app/<appId>` deep link. */
      readonly teamsAppId?: string;
      /** Scopes to grant, on the consent reasons (card material). */
      readonly missingScopes?: readonly string[];
    }
  | {
      readonly kind: 'failed';
      readonly agentSlug: string;
      readonly displayName: string;
      readonly reason: AgentAppFailureReason;
      /** Last `Retry-After` hint when `reason === 'throttled'`. */
      readonly retryAfterSeconds?: number;
      readonly message: string;
    };

/** Result of one {@link TeamsAgentInstaller.installAgentApps} run. */
export interface TeamsAgentInstallResult {
  readonly teamId: string;
  readonly outcomes: readonly AgentAppInstallOutcome[];
  /** `true` when at least one FRESH install happened and the auto-install
   *  marker was set for the team (intro-throttle window opened). */
  readonly autoInstallMarked: boolean;
}

/** `name`-based guard for the connector's `ConsentMissingError` — the class
 *  is not importable (peer-only dependency) and cross-package `instanceof`
 *  would be unreliable anyway; the connector sets `this.name` explicitly. */
export function isConsentMissingError(
  err: unknown,
): err is ConsentMissingErrorShape {
  return (
    err instanceof Error &&
    err.name === 'ConsentMissingError' &&
    Array.isArray((err as Partial<ConsentMissingErrorShape>).missingScopes)
  );
}

/** `name`-based guard for the connector's `ProvisioningThrottledError`. */
export function isProvisioningThrottledError(
  err: unknown,
): err is ProvisioningThrottledErrorShape {
  return err instanceof Error && err.name === 'ProvisioningThrottledError';
}

/**
 * Canonical key form for tenant/team ids — trim + lowercase, the
 * `normalizeTeamsBotAppId` policy (`teamsConversationRefStore.ts`): both are
 * GUIDs whose casing may drift between config and Graph responses, and a
 * casing drift must never split one tenant/team into two key spaces.
 * Returns `undefined` for missing/blank input.
 */
function normalizeIdKey(id: string | undefined): string | undefined {
  const normalized = id?.trim().toLowerCase();
  return normalized ? normalized : undefined;
}

/** Default consent negative-cache TTL — long enough to absorb a burst of
 *  `membersAdded` events, short enough that granted consent is picked up
 *  without a restart. */
export const CONSENT_NEGATIVE_CACHE_TTL_MS = 15 * 60 * 1000;

/**
 * Per-tenant negative cache for consent-missing (403) results. PROCESS-LOCAL
 * (see module doc) — a TTL-bounded courtesy throttle, not a global lock.
 */
export class TeamsConsentNegativeCache {
  private readonly entries = new Map<
    string,
    { readonly missingScopes: readonly string[]; readonly expiresAt: number }
  >();
  private readonly ttlMs: number;
  private readonly now: () => number;

  constructor(options?: { ttlMs?: number; now?: () => number }) {
    this.ttlMs = options?.ttlMs ?? CONSENT_NEGATIVE_CACHE_TTL_MS;
    this.now = options?.now ?? Date.now;
  }

  /** Record a 403 for the tenant. Blank tenant ids are ignored. */
  markConsentMissing(
    tenantId: string,
    missingScopes: readonly string[],
  ): void {
    const key = normalizeIdKey(tenantId);
    if (!key) return;
    this.entries.set(key, {
      missingScopes,
      expiresAt: this.now() + this.ttlMs,
    });
  }

  /** The cached scope set while the entry is fresh, else `undefined`. */
  get(tenantId: string): { missingScopes: readonly string[] } | undefined {
    const key = normalizeIdKey(tenantId);
    if (!key) return undefined;
    const entry = this.entries.get(key);
    if (!entry) return undefined;
    if (entry.expiresAt <= this.now()) {
      this.entries.delete(key);
      return undefined;
    }
    return { missingScopes: entry.missingScopes };
  }
}

/** Default auto-install marker TTL — `membersAdded` for a fresh install
 *  arrives within seconds; two minutes covers Graph fan-out lag without
 *  suppressing intros of genuinely later installs. */
export const AUTO_INSTALL_MARKER_TTL_MS = 2 * 60 * 1000;

/**
 * Short-lived "team X was just auto-installed" marker. Set by the installer
 * on a fresh install; PROBED (non-consuming — several bots correlate the
 * same event) by the onboarding hook (issue #21) so only the bot that ran
 * the installer posts the summary intro. PROCESS-LOCAL (see module doc).
 */
export class TeamsAutoInstallMarker {
  private readonly expiries = new Map<string, number>();
  private readonly ttlMs: number;
  private readonly now: () => number;

  constructor(options?: { ttlMs?: number; now?: () => number }) {
    this.ttlMs = options?.ttlMs ?? AUTO_INSTALL_MARKER_TTL_MS;
    this.now = options?.now ?? Date.now;
  }

  /** Open the intro-throttle window for a team. Blank ids are ignored. */
  mark(teamId: string): void {
    const key = normalizeIdKey(teamId);
    if (!key) return;
    this.expiries.set(key, this.now() + this.ttlMs);
  }

  /** `true` while the team's marker is fresh. Non-consuming. */
  probe(teamId: string): boolean {
    const key = normalizeIdKey(teamId);
    if (!key) return false;
    const expiresAt = this.expiries.get(key);
    if (expiresAt === undefined) return false;
    if (expiresAt <= this.now()) {
      this.expiries.delete(key);
      return false;
    }
    return true;
  }
}

/** Retries AFTER the first throttled attempt (the connector already burned
 *  its own 429 backoff budget per call — this is a second, outer bound). */
export const MAX_THROTTLE_RETRIES = 2;
/** Wait when the throttle carried no `Retry-After` hint. */
const DEFAULT_THROTTLE_WAIT_MS = 2000;
/** Upper bound on honoured `Retry-After` hints — the message path must not
 *  hang for minutes on an aggressive hint. */
const MAX_THROTTLE_WAIT_MS = 30_000;

export interface TeamsAgentInstallerOptions {
  /** The resolved `teamsProvisioner@1` service object. */
  readonly provisioner: TeamsProvisionerAccessor;
  /** Parsed `teams_agent_apps[]` entries (empty list = feature off). */
  readonly apps: readonly TeamsAgentAppTarget[];
  /** Shared negative cache — inject to share across installers; a fresh
   *  process-local one is created otherwise. */
  readonly consentCache?: TeamsConsentNegativeCache;
  /** Shared auto-install marker — same injection contract. */
  readonly marker?: TeamsAutoInstallMarker;
  readonly log?: (msg: string) => void;
  /** Test seam for throttle waits. */
  readonly sleep?: (ms: number) => Promise<void>;
  /** Bounded throttle retries per app (default {@link MAX_THROTTLE_RETRIES}). */
  readonly maxThrottleRetries?: number;
}

/** Install request — one team in one tenant. */
export interface InstallAgentAppsRequest {
  /** Graph team (group) id — target of `POST /teams/{id}/installedApps`. */
  readonly teamId: string;
  /** Entra tenant id the team lives in — negative-cache key. */
  readonly tenantId: string;
}

/** Log-safe label: slug + display name only (never app ids or scopes) —
 *  the `teamsBotLogLabel` policy applied to agent apps. */
export function teamsAgentAppLogLabel(app: TeamsAgentAppTarget): string {
  return `${app.agentSlug} ("${app.displayName}")`;
}

export class TeamsAgentInstaller {
  private readonly provisioner: TeamsProvisionerAccessor;
  private readonly apps: readonly TeamsAgentAppTarget[];
  private readonly consentCache: TeamsConsentNegativeCache;
  private readonly marker: TeamsAutoInstallMarker;
  private readonly log: (msg: string) => void;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly maxThrottleRetries: number;

  constructor(options: TeamsAgentInstallerOptions) {
    this.provisioner = options.provisioner;
    this.apps = options.apps;
    this.consentCache = options.consentCache ?? new TeamsConsentNegativeCache();
    this.marker = options.marker ?? new TeamsAutoInstallMarker();
    this.log = options.log ?? (() => {});
    this.sleep =
      options.sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
    this.maxThrottleRetries =
      options.maxThrottleRetries ?? MAX_THROTTLE_RETRIES;
  }

  /** The shared marker — the wiring layer hands this to the onboarding
   *  hook (issue #21) so it probes the same instance the installer sets. */
  get autoInstallMarker(): TeamsAutoInstallMarker {
    return this.marker;
  }

  /**
   * Install every configured agent app into the team. Never throws — every
   * app answers a typed {@link AgentAppInstallOutcome}. A 403 on one app
   * short-circuits the REST of the run onto the fallback path too (same
   * tenant, same missing grant — calling on would only repeat the 403).
   */
  async installAgentApps(
    request: InstallAgentAppsRequest,
  ): Promise<TeamsAgentInstallResult> {
    const outcomes: AgentAppInstallOutcome[] = [];
    let consentBlocked = this.consentCache.get(request.tenantId);
    let freshInstall = false;

    for (const app of this.apps) {
      if (consentBlocked) {
        outcomes.push({
          kind: 'fallback',
          agentSlug: app.agentSlug,
          displayName: app.displayName,
          reason: 'consent-cached',
          teamsAppExternalId: app.teamsAppExternalId,
          ...(app.teamsAppId !== undefined
            ? { teamsAppId: app.teamsAppId }
            : {}),
          missingScopes: consentBlocked.missingScopes,
        });
        continue;
      }

      const outcome = await this.installOne(request, app);
      outcomes.push(outcome);
      if (outcome.kind === 'installed' && outcome.outcome === 'created') {
        freshInstall = true;
      }
      if (outcome.kind === 'fallback' && outcome.reason === 'consent-missing') {
        // Freshly observed 403 — cached inside installOne; short-circuit
        // the remaining apps of this run through the cache branch above.
        consentBlocked = this.consentCache.get(request.tenantId) ?? {
          missingScopes: outcome.missingScopes ?? [],
        };
      }
    }

    if (freshInstall) this.marker.mark(request.teamId);
    return { teamId: request.teamId, outcomes, autoInstallMarked: freshInstall };
  }

  /** Resolve + install one app, mapping every contract signal to an outcome. */
  private async installOne(
    request: InstallAgentAppsRequest,
    app: TeamsAgentAppTarget,
  ): Promise<AgentAppInstallOutcome> {
    const label = teamsAgentAppLogLabel(app);
    try {
      const teamsAppId = await this.resolveCatalogAppId(app);
      if (teamsAppId === undefined) {
        this.log(`teamsAgentInstaller: ${label} not in catalog → fallback card`);
        return {
          kind: 'fallback',
          agentSlug: app.agentSlug,
          displayName: app.displayName,
          reason: 'not-in-catalog',
          teamsAppExternalId: app.teamsAppExternalId,
        };
      }
      const installed = await this.withThrottleRetry(label, () =>
        this.provisioner.installToTeam({ teamId: request.teamId, teamsAppId }),
      );
      this.log(
        `teamsAgentInstaller: ${label} install outcome=${installed.outcome}`,
      );
      return {
        kind: 'installed',
        agentSlug: app.agentSlug,
        displayName: app.displayName,
        teamsAppId,
        outcome: installed.outcome,
      };
    } catch (err) {
      return this.mapError(request, app, label, err);
    }
  }

  /**
   * Catalog id for the app: configured `teamsAppId` first; else the
   * feature-detected `getCatalogApp` lookup; else `undefined` (= the
   * "app not in catalog" fallback outcome — connector 0.3.0 has no lookup
   * and a pre-existing catalog app is not resolvable from `externalId`).
   */
  private async resolveCatalogAppId(
    app: TeamsAgentAppTarget,
  ): Promise<string | undefined> {
    if (app.teamsAppId !== undefined) return app.teamsAppId;
    if (typeof this.provisioner.getCatalogApp !== 'function') return undefined;
    const label = teamsAgentAppLogLabel(app);
    const result: GetCatalogAppResult = await this.withThrottleRetry(
      label,
      () =>
        this.provisioner.getCatalogApp!({
          teamsAppExternalId: app.teamsAppExternalId,
        }),
    );
    return result.found ? result.teamsAppId : undefined;
  }

  /** Bounded outer retry over the connector's throttle signal. */
  private async withThrottleRetry<T>(
    label: string,
    step: () => Promise<T>,
  ): Promise<T> {
    for (let attempt = 0; ; attempt += 1) {
      try {
        return await step();
      } catch (err) {
        if (!isProvisioningThrottledError(err) || attempt >= this.maxThrottleRetries) {
          throw err;
        }
        const hintedMs =
          err.retryAfterSeconds !== undefined
            ? err.retryAfterSeconds * 1000
            : DEFAULT_THROTTLE_WAIT_MS;
        const waitMs = Math.min(Math.max(hintedMs, 0), MAX_THROTTLE_WAIT_MS);
        this.log(
          `teamsAgentInstaller: ${label} throttled — retry ${String(attempt + 1)}/${String(this.maxThrottleRetries)} in ${String(waitMs)}ms`,
        );
        await this.sleep(waitMs);
      }
    }
  }

  /** Map a thrown contract error to the typed outcome (never rethrows). */
  private mapError(
    request: InstallAgentAppsRequest,
    app: TeamsAgentAppTarget,
    label: string,
    err: unknown,
  ): AgentAppInstallOutcome {
    if (isConsentMissingError(err)) {
      this.consentCache.markConsentMissing(request.tenantId, err.missingScopes);
      this.log(
        `teamsAgentInstaller: ${label} consent missing (${err.resource}) → fallback card, tenant negative-cached`,
      );
      return {
        kind: 'fallback',
        agentSlug: app.agentSlug,
        displayName: app.displayName,
        reason: 'consent-missing',
        teamsAppExternalId: app.teamsAppExternalId,
        ...(app.teamsAppId !== undefined ? { teamsAppId: app.teamsAppId } : {}),
        missingScopes: err.missingScopes,
      };
    }
    if (isProvisioningThrottledError(err)) {
      this.log(`teamsAgentInstaller: ${label} throttle retries exhausted`);
      return {
        kind: 'failed',
        agentSlug: app.agentSlug,
        displayName: app.displayName,
        reason: 'throttled',
        ...(err.retryAfterSeconds !== undefined
          ? { retryAfterSeconds: err.retryAfterSeconds }
          : {}),
        message: 'provisioning_throttled',
      };
    }
    const message = err instanceof Error ? err.message : String(err);
    this.log(`teamsAgentInstaller: ${label} install failed (${message})`);
    return {
      kind: 'failed',
      agentSlug: app.agentSlug,
      displayName: app.displayName,
      reason: 'error',
      message,
    };
  }
}

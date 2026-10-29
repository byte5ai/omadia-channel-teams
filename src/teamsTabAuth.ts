import { createPublicKey, verify as verifySignature } from 'node:crypto';
import type { webcrypto } from 'node:crypto';

type JsonWebKey = webcrypto.JsonWebKey;

/**
 * Teams-tab SSO verification for the Hub / Tab-Config surfaces
 * (byte5ai/omadia#1596 follow-up).
 *
 * Teams iframes `/p/channel-teams/hub` and `/tab-config` WITHOUT the omadia
 * session cookie, so the kernel's `auth:'session'` gate can never pass there.
 * The routes are therefore registered `auth:'custom'` behind a granted
 * `permissions.public_paths` entry, and this module is the "custom" part:
 *
 *   - The HTML pages themselves are data-free shells.
 *   - The page asks Teams for an SSO token (`authentication.getAuthToken()`,
 *     resource = the app package's `webApplicationInfo`) and sends it as a
 *     Bearer token to the JSON endpoint.
 *   - The JSON endpoint serves the plugin catalogue only after this module
 *     verified the token: RS256 signature against Entra's published keys,
 *     issuer + `tid` bound to a configured bot tenant, audience bound to that
 *     bot's app id, issued to a Microsoft Teams / M365 client (`azp` /
 *     `appid`), a delegated user token (`scp` + `oid`), and lifetime.
 *
 *   Anyone signed in to a configured tenant passes — including B2B guests,
 *   whose tokens carry the resource tenant's `tid`. That is acceptable for
 *   what `/catalog` serves (plugin ids, titles, paths); the pages it links to
 *   enforce their own authentication.
 *
 * The tab context (`app.getContext()`) is NOT used for authentication: it is
 * client-side data any page can fabricate.
 *
 * Dependency-free on purpose (node:crypto JWK import + fetch): the plugin
 * resolves its runtime dependencies from the host's node_modules, and a JWT
 * library the host does not ship would fail at activate().
 */

/** One bot registration whose SSO tokens are accepted. */
export interface TeamsSsoAudience {
  readonly appId: string;
  readonly tenantId: string;
}

export interface TeamsSsoPrincipal {
  readonly tenantId: string;
  readonly objectId: string;
  readonly appId: string;
  /** `preferred_username` / `upn` when present — for logs only. */
  readonly username?: string;
}

export interface TeamsSsoVerifierOptions {
  readonly audiences: readonly TeamsSsoAudience[];
  /** HTTP client for key discovery. Pass `ctx.http.fetch` so the call runs
   *  through the kernel's `network.outbound` allow-list; defaults to the
   *  global `fetch` (older cores without `ctx.http`). */
  readonly fetchImpl?: (url: string, init?: RequestInit) => Promise<Response>;
  /** Fetch the JWKS document for a tenant. Injected for tests; overrides
   *  `fetchImpl`. */
  readonly fetchJwks?: (tenantId: string) => Promise<unknown>;
  /** Clock in ms. Injected for tests. */
  readonly now?: () => number;
}

export interface TeamsSsoVerifier {
  /** Resolves with the principal, rejects with {@link TeamsSsoError}. */
  verify(token: string): Promise<TeamsSsoPrincipal>;
}

export class TeamsSsoError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'TeamsSsoError';
  }
}

const CLOCK_SKEW_SECONDS = 300;
const JWKS_TTL_MS = 60 * 60 * 1000;
/** Floor between two key fetches for one tenant, so a token with an unknown
 *  `kid` cannot turn every request into an outbound call. */
const JWKS_REFETCH_FLOOR_MS = 5 * 60 * 1000;
const MAX_TOKEN_LENGTH = 16 * 1024;

/** After a failed key fetch, refuse (instead of refetching) for this long,
 *  so an anonymous caller cannot turn a slow or failing Entra endpoint into
 *  one outbound request per inbound request. */
const JWKS_FAILURE_BACKOFF_MS = 60 * 1000;

const ENTRA_HOST = 'https://login.microsoftonline.com';

/**
 * Microsoft's own client apps that request a tab SSO token on the user's
 * behalf (Teams desktop/mobile + web, and the Microsoft 365 / Outlook hosts
 * that run Teams apps). A token's `azp` (v2) / `appid` (v1) must be one of
 * them: that pins the token to one a Microsoft host minted for this tab, not
 * one some other app in the tenant obtained for a look-alike identifier URI.
 */
export const TEAMS_SSO_CLIENT_IDS: readonly string[] = [
  '1fec8e78-bce4-4aaf-ab1b-5451cc387264', // Teams desktop / mobile
  '5e3ce6c0-2b1f-4285-8d4b-75ee78787346', // Teams web
  '4765445b-32c6-49b0-83e6-1d93765276ca', // Microsoft 365 web
  '0ec893e0-5785-4de6-99da-4ed124e5296c', // Microsoft 365 desktop
  'd3590ed6-52b3-4102-aeff-aad2292ab01c', // Microsoft 365 mobile / Outlook desktop
  'bc59ab01-8403-45c6-8796-ac3ef710b3e3', // Outlook web
  '27922004-5251-4030-b22d-91ecd9a37ea4', // Outlook mobile
];

function jwksFetcher(
  fetchImpl: (url: string, init?: RequestInit) => Promise<Response>,
): (tenantId: string) => Promise<unknown> {
  return async (tenantId) => {
    const res = await fetchImpl(
      `${ENTRA_HOST}/${encodeURIComponent(tenantId)}/discovery/v2.0/keys`,
      { signal: AbortSignal.timeout(5_000) },
    );
    if (!res.ok) {
      throw new TeamsSsoError(`key discovery answered ${String(res.status)}`);
    }
    return res.json();
  };
}

function decodeSegment(segment: string): Record<string, unknown> {
  let parsed: unknown;
  try {
    parsed = JSON.parse(Buffer.from(segment, 'base64url').toString('utf8'));
  } catch {
    throw new TeamsSsoError('token segment is not JSON');
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new TeamsSsoError('token segment is not an object');
  }
  return parsed as Record<string, unknown>;
}

function stringClaim(
  claims: Record<string, unknown>,
  name: string,
): string | undefined {
  const value = claims[name];
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

function numberClaim(
  claims: Record<string, unknown>,
  name: string,
): number | undefined {
  const value = claims[name];
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined;
}

/** Audience forms Entra issues for a Teams SSO resource: the bare app id
 *  (v2 tokens) or the `api://<host>/<appId>` identifier URI (v1 tokens). */
function audienceMatches(aud: string, appId: string): boolean {
  const want = appId.toLowerCase();
  const got = aud.toLowerCase();
  if (got === want) return true;
  const match = /^api:\/\/[^/]+\/([^/]+)$/.exec(got);
  return match?.[1] === want;
}

interface CachedKeys {
  readonly keys: ReadonlyMap<string, JsonWebKey>;
  readonly fetchedAt: number;
  /** Set when the fetch failed; `keys` is empty then. */
  readonly failed?: boolean;
}

function parseJwks(doc: unknown): ReadonlyMap<string, JsonWebKey> {
  const keys = new Map<string, JsonWebKey>();
  const list =
    typeof doc === 'object' && doc !== null
      ? (doc as { keys?: unknown }).keys
      : undefined;
  if (!Array.isArray(list)) return keys;
  for (const entry of list) {
    if (typeof entry !== 'object' || entry === null) continue;
    const jwk = entry as JsonWebKey & { kid?: unknown };
    if (typeof jwk.kid === 'string' && jwk.kty === 'RSA') {
      keys.set(jwk.kid, jwk);
    }
  }
  return keys;
}

export function createTeamsSsoVerifier(
  opts: TeamsSsoVerifierOptions,
): TeamsSsoVerifier {
  const fetchJwks =
    opts.fetchJwks ?? jwksFetcher(opts.fetchImpl ?? ((url, init) => fetch(url, init)));
  const now = opts.now ?? Date.now;
  const audiencesByTenant = new Map<string, string[]>();
  for (const { appId, tenantId } of opts.audiences) {
    const key = tenantId.toLowerCase();
    audiencesByTenant.set(key, [...(audiencesByTenant.get(key) ?? []), appId]);
  }
  const keyCache = new Map<string, CachedKeys>();
  /** One key fetch per tenant at a time; concurrent callers share it. */
  const inflight = new Map<string, Promise<CachedKeys>>();

  function refresh(tenantId: string): Promise<CachedKeys> {
    const running = inflight.get(tenantId);
    if (running) return running;
    const task = (async (): Promise<CachedKeys> => {
      let entry: CachedKeys;
      try {
        entry = { keys: parseJwks(await fetchJwks(tenantId)), fetchedAt: now() };
      } catch {
        entry = { keys: new Map(), fetchedAt: now(), failed: true };
      }
      keyCache.set(tenantId, entry);
      return entry;
    })().finally(() => inflight.delete(tenantId));
    inflight.set(tenantId, task);
    return task;
  }

  async function keyFor(tenantId: string, kid: string): Promise<JsonWebKey> {
    const cached = keyCache.get(tenantId);
    const age = cached ? now() - cached.fetchedAt : Infinity;
    if (cached?.failed && age < JWKS_FAILURE_BACKOFF_MS) {
      throw new TeamsSsoError('key discovery unavailable');
    }
    const hit = cached?.keys.get(kid);
    if (hit && age < JWKS_TTL_MS) return hit;
    if (cached && !cached.failed && !hit && age < JWKS_REFETCH_FLOOR_MS) {
      throw new TeamsSsoError('unknown signing key');
    }
    const fresh = await refresh(tenantId);
    if (fresh.failed) throw new TeamsSsoError('key discovery unavailable');
    const key = fresh.keys.get(kid);
    if (!key) throw new TeamsSsoError('unknown signing key');
    return key;
  }

  return {
    async verify(token) {
      if (token.length === 0 || token.length > MAX_TOKEN_LENGTH) {
        throw new TeamsSsoError('token missing or oversized');
      }
      const parts = token.split('.');
      if (parts.length !== 3) throw new TeamsSsoError('token is not a JWS');
      const [headerSeg, payloadSeg, signatureSeg] = parts as [
        string,
        string,
        string,
      ];
      const header = decodeSegment(headerSeg);
      if (header['alg'] !== 'RS256') {
        throw new TeamsSsoError('unsupported token algorithm');
      }
      const kid = stringClaim(header, 'kid');
      if (!kid) throw new TeamsSsoError('token has no key id');

      // Claims are read BEFORE the signature check only to pick the tenant,
      // and only a configured tenant ever reaches key discovery — an
      // attacker-chosen `tid` cannot make us fetch anything.
      const claims = decodeSegment(payloadSeg);
      const tenantId = stringClaim(claims, 'tid')?.toLowerCase();
      const allowedAppIds = tenantId
        ? audiencesByTenant.get(tenantId)
        : undefined;
      if (!tenantId || !allowedAppIds) {
        throw new TeamsSsoError('token tenant is not a configured bot tenant');
      }

      const jwk = await keyFor(tenantId, kid);
      let valid = false;
      try {
        valid = verifySignature(
          'RSA-SHA256',
          Buffer.from(`${headerSeg}.${payloadSeg}`),
          createPublicKey({ key: jwk, format: 'jwk' }),
          Buffer.from(signatureSeg, 'base64url'),
        );
      } catch {
        valid = false;
      }
      if (!valid) throw new TeamsSsoError('bad token signature');

      const iss = stringClaim(claims, 'iss');
      const issuers = [
        `${ENTRA_HOST}/${tenantId}/v2.0`,
        `https://sts.windows.net/${tenantId}/`,
      ];
      if (!iss || !issuers.includes(iss.toLowerCase())) {
        throw new TeamsSsoError('unexpected token issuer');
      }

      const aud = stringClaim(claims, 'aud');
      const appId = aud
        ? allowedAppIds.find((id) => audienceMatches(aud, id))
        : undefined;
      if (!appId) throw new TeamsSsoError('token audience is not a bot app');

      const client = (
        stringClaim(claims, 'azp') ?? stringClaim(claims, 'appid')
      )?.toLowerCase();
      if (!client || !TEAMS_SSO_CLIENT_IDS.includes(client)) {
        throw new TeamsSsoError('token was not issued to a Teams client');
      }

      const nowSeconds = Math.floor(now() / 1000);
      const exp = numberClaim(claims, 'exp');
      if (exp === undefined || exp + CLOCK_SKEW_SECONDS < nowSeconds) {
        throw new TeamsSsoError('token expired');
      }
      const nbf = numberClaim(claims, 'nbf');
      if (nbf !== undefined && nbf - CLOCK_SKEW_SECONDS > nowSeconds) {
        throw new TeamsSsoError('token not yet valid');
      }

      // A Teams SSO token is a DELEGATED token for a signed-in user. An
      // app-only token for the same audience (client credentials) has
      // `roles` but no `scp` — it must not open a user-facing surface.
      const objectId = stringClaim(claims, 'oid');
      if (!stringClaim(claims, 'scp') || !objectId) {
        throw new TeamsSsoError('token is not a delegated user token');
      }

      const username =
        stringClaim(claims, 'preferred_username') ?? stringClaim(claims, 'upn');
      return {
        tenantId,
        objectId,
        appId,
        ...(username ? { username } : {}),
      };
    },
  };
}

/** `Authorization: Bearer <token>` → token, else undefined. */
export function bearerToken(header: string | undefined): string | undefined {
  if (!header) return undefined;
  const match = /^Bearer\s+(\S+)\s*$/i.exec(header);
  return match?.[1];
}

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
 *     bot's app id, a delegated user token (`scp` + `oid`), and lifetime.
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
  /** Fetch the JWKS document for a tenant. Injected for tests. */
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

const ENTRA_HOST = 'https://login.microsoftonline.com';

async function defaultFetchJwks(tenantId: string): Promise<unknown> {
  const res = await fetch(
    `${ENTRA_HOST}/${encodeURIComponent(tenantId)}/discovery/v2.0/keys`,
    { signal: AbortSignal.timeout(5_000) },
  );
  if (!res.ok) {
    throw new TeamsSsoError(`key discovery answered ${String(res.status)}`);
  }
  return res.json();
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
  const fetchJwks = opts.fetchJwks ?? defaultFetchJwks;
  const now = opts.now ?? Date.now;
  const audiencesByTenant = new Map<string, string[]>();
  for (const { appId, tenantId } of opts.audiences) {
    const key = tenantId.toLowerCase();
    audiencesByTenant.set(key, [...(audiencesByTenant.get(key) ?? []), appId]);
  }
  const keyCache = new Map<string, CachedKeys>();

  async function keyFor(tenantId: string, kid: string): Promise<JsonWebKey> {
    const cached = keyCache.get(tenantId);
    const age = cached ? now() - cached.fetchedAt : Infinity;
    const hit = cached?.keys.get(kid);
    if (hit && age < JWKS_TTL_MS) return hit;
    if (cached && !hit && age < JWKS_REFETCH_FLOOR_MS) {
      throw new TeamsSsoError('unknown signing key');
    }
    const keys = parseJwks(await fetchJwks(tenantId));
    keyCache.set(tenantId, { keys, fetchedAt: now() });
    const fresh = keys.get(kid);
    if (!fresh) throw new TeamsSsoError('unknown signing key');
    return fresh;
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

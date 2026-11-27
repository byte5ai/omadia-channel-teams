// Spike A-S1 (byte5ai/omadia#584, omadia ADR-0287): an ACS token for the
// agent's own tenant user, without an interactive login at run time.
//
// SPIKE CODE, not wired, and gated: nothing here runs unless
// TEAMS_MEETING_AUDIO_SPIKE=1 (see `assertMeetingAudioSpikeEnabled`).
//
// Two steps:
//   1. An Entra token of the tenant user for the ACS scopes
//      (`https://auth.msft.communication.azure.com/Teams.ManageCalls` and
//      `…/Teams.ManageChats`). Prototyped here: the refresh-token grant of a
//      public client, the refresh token coming from ONE device-code login by
//      an admin when the user is provisioned (90-day rolling lifetime; a
//      password reset, "revoke sessions" or a sign-in-frequency policy forces
//      a new login). The password grant (ROPC) is deliberately not
//      implemented: Microsoft deprecates it and it fails for any account
//      under MFA. The Entra Agent ID route (agent user, no password, token via
//      the blueprint's credential) is the preferred target but has no Node
//      flow Microsoft supports ("don't implement the token exchange
//      directly"); it needs a live test (docs/meeting-audio-spikes.md).
//   2. The ACS exchange `POST {endpoint}/teamsUser/:exchangeAccessToken`
//      (api-version 2023-10-01), HMAC-signed with the ACS resource key. The
//      result is the token the sidecar's `TeamsCallAgent` uses.
//
// Secrets (refresh token, ACS key) are parameters; this module never reads
// or logs them, and errors carry the HTTP status and error code only.

import { createHash, createHmac } from 'node:crypto';

export const MEETING_AUDIO_SPIKE_FLAG = 'TEAMS_MEETING_AUDIO_SPIKE';

export const ACS_TEAMS_SCOPES = [
  'https://auth.msft.communication.azure.com/Teams.ManageCalls',
  'https://auth.msft.communication.azure.com/Teams.ManageChats',
] as const;

export const ACS_EXCHANGE_API_VERSION = '2023-10-01';

export type FetchLike = (url: string, init: { method: string; headers: Record<string, string>; body: string }) => Promise<{
  readonly status: number;
  text(): Promise<string>;
}>;

export class MeetingAudioSpikeError extends Error {
  override readonly name = 'MeetingAudioSpikeError';
  constructor(
    message: string,
    readonly status?: number,
    readonly code?: string,
  ) {
    super(message);
  }
}

export function assertMeetingAudioSpikeEnabled(env: Readonly<Record<string, string | undefined>>): void {
  if (env[MEETING_AUDIO_SPIKE_FLAG] !== '1') {
    throw new MeetingAudioSpikeError(`${MEETING_AUDIO_SPIKE_FLAG}=1 is required for the meeting-audio spike`);
  }
}

function errorCodeOf(body: string): string | undefined {
  try {
    const parsed = JSON.parse(body) as { error?: unknown };
    if (typeof parsed.error === 'string') return parsed.error;
    if (parsed.error && typeof parsed.error === 'object') {
      const code = (parsed.error as { code?: unknown }).code;
      return typeof code === 'string' ? code : undefined;
    }
  } catch {
    // not JSON: no code
  }
  return undefined;
}

async function readJson(response: { status: number; text(): Promise<string> }, step: string): Promise<Record<string, unknown>> {
  const body = await response.text();
  if (response.status < 200 || response.status >= 300) {
    const code = errorCodeOf(body);
    throw new MeetingAudioSpikeError(`${step} failed: HTTP ${response.status}${code ? ` ${code}` : ''}`, response.status, code);
  }
  try {
    return JSON.parse(body) as Record<string, unknown>;
  } catch {
    throw new MeetingAudioSpikeError(`${step} answered with a body that is not JSON`, response.status);
  }
}

function requireString(value: unknown, field: string, step: string): string {
  if (typeof value !== 'string' || value.length === 0) throw new MeetingAudioSpikeError(`${step}: response has no ${field}`);
  return value;
}

export interface EntraRefreshInput {
  readonly tenantId: string;
  /** Public-client app registration with the two ACS delegated permissions. */
  readonly clientId: string;
  readonly refreshToken: string;
}

export interface EntraTokens {
  readonly accessToken: string;
  /** Rolling: store it back, the old one may stop working. */
  readonly refreshToken?: string;
  readonly expiresInSec?: number;
}

/** Step 1: refresh-token grant against the tenant's v2 token endpoint. */
export async function refreshEntraToken(input: EntraRefreshInput, fetchImpl: FetchLike): Promise<EntraTokens> {
  const body = new URLSearchParams({
    client_id: input.clientId,
    grant_type: 'refresh_token',
    refresh_token: input.refreshToken,
    scope: [...ACS_TEAMS_SCOPES, 'offline_access'].join(' '),
  }).toString();
  const response = await fetchImpl(
    `https://login.microsoftonline.com/${encodeURIComponent(input.tenantId)}/oauth2/v2.0/token`,
    { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body },
  );
  const json = await readJson(response, 'entra refresh');
  return {
    accessToken: requireString(json.access_token, 'access_token', 'entra refresh'),
    ...(typeof json.refresh_token === 'string' ? { refreshToken: json.refresh_token } : {}),
    ...(typeof json.expires_in === 'number' ? { expiresInSec: json.expires_in } : {}),
  };
}

export interface AcsExchangeInput {
  /** ACS resource endpoint, e.g. `https://<name>.communication.azure.com`. */
  readonly endpoint: string;
  /** Base64 access key of the ACS resource. */
  readonly accessKey: string;
  /** The tenant user's Entra token for the ACS scopes. */
  readonly entraAccessToken: string;
  /** Must equal the token's `appid`. */
  readonly clientId: string;
  /** Must equal the token's `oid`: the agent's tenant user. */
  readonly userObjectId: string;
  /** Injectable for tests (RFC 1123 date). */
  readonly now?: () => Date;
}

export interface AcsToken {
  readonly token: string;
  readonly expiresOn: string;
}

/** ACS HMAC-SHA256 request signature (Azure Communication Services REST auth). */
export function signAcsRequest(args: {
  readonly method: string;
  readonly url: URL;
  readonly body: string;
  readonly accessKey: string;
  readonly date: Date;
}): Record<string, string> {
  const contentHash = createHash('sha256').update(args.body, 'utf8').digest('base64');
  const date = args.date.toUTCString();
  const host = args.url.host;
  const stringToSign = `${args.method.toUpperCase()}\n${args.url.pathname}${args.url.search}\n${date};${host};${contentHash}`;
  const signature = createHmac('sha256', Buffer.from(args.accessKey, 'base64')).update(stringToSign, 'utf8').digest('base64');
  return {
    'x-ms-date': date,
    'x-ms-content-sha256': contentHash,
    authorization: `HMAC-SHA256 SignedHeaders=x-ms-date;host;x-ms-content-sha256&Signature=${signature}`,
  };
}

/** Step 2: exchange the Entra token for an ACS token of the Teams user. */
export async function exchangeForAcsToken(input: AcsExchangeInput, fetchImpl: FetchLike): Promise<AcsToken> {
  const url = new URL(`/teamsUser/:exchangeAccessToken?api-version=${ACS_EXCHANGE_API_VERSION}`, input.endpoint);
  if (url.protocol !== 'https:') throw new MeetingAudioSpikeError('ACS endpoint must be https');
  const body = JSON.stringify({ token: input.entraAccessToken, appId: input.clientId, userId: input.userObjectId });
  const headers = {
    'content-type': 'application/json',
    ...signAcsRequest({ method: 'POST', url, body, accessKey: input.accessKey, date: (input.now ?? (() => new Date()))() }),
  };
  const json = await readJson(await fetchImpl(url.toString(), { method: 'POST', headers, body }), 'acs exchange');
  return {
    token: requireString(json.token, 'token', 'acs exchange'),
    expiresOn: requireString(json.expiresOn, 'expiresOn', 'acs exchange'),
  };
}

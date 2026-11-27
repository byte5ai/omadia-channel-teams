// Spike A-S1, live run: get an ACS token for the test tenant user.
//
//   TEAMS_MEETING_AUDIO_SPIKE=1 TENANT_ID=… CLIENT_ID=… USER_OID=… \
//   ACS_ENDPOINT=https://….communication.azure.com node out/live-token.mjs
//
// First run: device-code login (a person signs in once as the test user,
// MFA included). Every later run: the refresh token from the macOS keychain,
// no person needed — which is exactly what A-S1 has to prove.
// Secrets live only in the macOS keychain (`security`): the ACS key
// (service omadia-w44-acs-key), the rolling refresh token
// (omadia-w44-refresh-token) and the resulting ACS token
// (omadia-w44-acs-token, read by run.ts via ACS_TOKEN=$(security …)).
// Printed are only non-secret claims and expiry times.

import { execFileSync } from 'node:child_process';

import {
  ACS_TEAMS_SCOPES,
  assertMeetingAudioSpikeEnabled,
  exchangeForAcsToken,
  refreshEntraToken,
  type FetchLike,
} from '../src/teamsUserToken.js';

const KEYCHAIN = {
  acsKey: 'omadia-w44-acs-key',
  refresh: 'omadia-w44-refresh-token',
  acsToken: 'omadia-w44-acs-token',
} as const;
const DEVICE_GRANT = 'urn:ietf:params:oauth:grant-type:device_code';

function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value) throw new Error(`${name} is required`);
  return value;
}

function keychainRead(service: string): string | undefined {
  try {
    return execFileSync('security', ['find-generic-password', '-s', service, '-w'], { encoding: 'utf8' }).trim() || undefined;
  } catch {
    return undefined;
  }
}

function keychainWrite(service: string, account: string, secret: string): void {
  execFileSync('security', ['add-generic-password', '-a', account, '-s', service, '-w', secret, '-U']);
}

const fetchImpl: FetchLike = async (url, init) => fetch(url, init);

async function postForm(url: string, form: Record<string, string>): Promise<{ status: number; json: Record<string, unknown> }> {
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(form).toString(),
  });
  return { status: response.status, json: (await response.json()) as Record<string, unknown> };
}

async function deviceCodeLogin(tenantId: string, clientId: string): Promise<string> {
  const base = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0`;
  const scope = [...ACS_TEAMS_SCOPES, 'offline_access'].join(' ');
  const start = await postForm(`${base}/devicecode`, { client_id: clientId, scope });
  if (start.status !== 200) throw new Error(`devicecode failed: HTTP ${start.status} ${String(start.json.error ?? '')}`);
  console.log(String(start.json.message));
  const intervalMs = Number(start.json.interval ?? 5) * 1000;
  const deadline = Date.now() + Number(start.json.expires_in ?? 900) * 1000;
  while (Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
    const poll = await postForm(`${base}/token`, { client_id: clientId, grant_type: DEVICE_GRANT, device_code: String(start.json.device_code) });
    if (poll.status === 200 && typeof poll.json.refresh_token === 'string') return poll.json.refresh_token;
    const error = String(poll.json.error ?? '');
    if (error !== 'authorization_pending' && error !== 'slow_down') throw new Error(`device login failed: ${error}`);
  }
  throw new Error('device login expired');
}

/** Non-secret claims of a JWT; tokens that are not plain JWTs (opaque or
 *  encrypted) report their shape instead. */
function claimsOf(token: string): Record<string, unknown> {
  const parts = token.split('.');
  try {
    return JSON.parse(Buffer.from(parts[1] ?? '', 'base64url').toString('utf8')) as Record<string, unknown>;
  } catch {
    return { notAJwt: true, parts: parts.length, length: token.length };
  }
}

async function main(): Promise<void> {
  assertMeetingAudioSpikeEnabled(process.env);
  const tenantId = requireEnv('TENANT_ID');
  const clientId = requireEnv('CLIENT_ID');
  const userObjectId = requireEnv('USER_OID');
  const endpoint = requireEnv('ACS_ENDPOINT');
  const accessKey = keychainRead(KEYCHAIN.acsKey);
  if (!accessKey) throw new Error(`keychain item ${KEYCHAIN.acsKey} is missing`);

  const stored = keychainRead(KEYCHAIN.refresh);
  const refreshToken = stored ?? (await deviceCodeLogin(tenantId, clientId));
  console.log(stored ? 'using the stored refresh token (no person involved)' : 'device login done');

  const entra = await refreshEntraToken({ tenantId, clientId, refreshToken }, fetchImpl);
  keychainWrite(KEYCHAIN.refresh, userObjectId, entra.refreshToken ?? refreshToken);
  const claims = claimsOf(entra.accessToken);
  console.log(JSON.stringify({ entra: { shape: claims.notAJwt ? claims : undefined, aud: claims.aud, scp: claims.scp, oid: claims.oid, appid: claims.appid, idtyp: claims.idtyp, expiresInSec: entra.expiresInSec } }));

  const acs = await exchangeForAcsToken({ endpoint, accessKey, entraAccessToken: entra.accessToken, clientId, userObjectId }, fetchImpl);
  keychainWrite(KEYCHAIN.acsToken, userObjectId, acs.token);
  const acsClaims = claimsOf(acs.token);
  console.log(JSON.stringify({ acs: { expiresOn: acs.expiresOn, skypeid: acsClaims.skypeid, scp: acsClaims.scp } }));
}

main().catch((err: unknown) => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exitCode = 1;
});

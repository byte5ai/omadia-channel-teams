import { createHmac, timingSafeEqual } from 'node:crypto';

/**
 * HMAC-SHA256 signing for the `/attachments/<key>` proxy — structurally
 * identical to the diagram signing, but kept in its own file so the two
 * subsystems can rotate secrets independently and pick different TTLs.
 *
 * Scheme:  `/attachments/<encoded-key>?exp=<unix-seconds>&sig=<hex>`
 * Payload: `<key>.<exp>` (prevents replay for a different key)
 *
 * The secret lives in ATTACHMENT_URL_SECRET and never leaves the process.
 * Rotate by setting a new value — in-flight URLs become invalid immediately.
 */

export interface SignAttachmentUrlParams {
  key: string;
  secret: string;
  ttlSec: number;
  publicBaseUrl: string;
  nowSec?: number;
}

export function signAttachmentUrl(params: SignAttachmentUrlParams): string {
  const now = params.nowSec ?? Math.floor(Date.now() / 1000);
  const exp = now + params.ttlSec;
  const sig = createHmac('sha256', params.secret)
    .update(`${params.key}.${String(exp)}`)
    .digest('hex');
  const encKey = params.key
    .split('/')
    .map((segment) => encodeURIComponent(segment))
    .join('/');
  const base = params.publicBaseUrl.replace(/\/+$/, '');
  return `${base}/attachments/${encKey}?exp=${String(exp)}&sig=${sig}`;
}

export interface VerifyAttachmentSigParams {
  key: string;
  exp: number;
  sig: string;
  secret: string;
  nowSec?: number;
}

export function verifyAttachmentSig(params: VerifyAttachmentSigParams): boolean {
  const now = params.nowSec ?? Math.floor(Date.now() / 1000);
  if (!Number.isFinite(params.exp) || params.exp < now) return false;
  if (!/^[0-9a-f]+$/i.test(params.sig)) return false;
  const expected = createHmac('sha256', params.secret)
    .update(`${params.key}.${String(params.exp)}`)
    .digest('hex');
  const a = Buffer.from(expected, 'hex');
  const b = Buffer.from(params.sig, 'hex');
  if (a.length === 0 || a.length !== b.length) return false;
  return timingSafeEqual(a, b);
}

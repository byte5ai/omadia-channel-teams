import { Router } from 'express';
import type { Request, Response } from 'express';
import { isNotFound, type TigrisStore } from '@omadia/diagrams';
import { verifyAttachmentSig } from './attachmentSigning.js';

interface AttachmentsRouterDeps {
  store: TigrisStore;
  secret: string;
  /**
   * Optional prefix whitelist. Only keys starting with one of these prefixes
   * are served via this proxy. Defaults to `['teams-attachments/']` so a
   * compromised signer cannot hand out diagram PNGs through this route.
   */
  allowedPrefixes?: readonly string[];
}

/**
 * HMAC-signed proxy for persisted Teams attachments (logos, PDFs, images
 * referenced by the bot). Mirror of the diagrams router, but keyed to a
 * separate secret + prefix namespace so the two subsystems can rotate
 * independently.
 *
 * URL format: `/attachments/<url-encoded-key>?exp=<unix-seconds>&sig=<hex>`
 *
 *   - 403 on invalid signature / expired / disallowed prefix.
 *   - 404 when Tigris has no such object.
 *   - 200 streams the raw bytes with the stored content-type.
 *
 * Like the diagrams proxy, we stream ourselves rather than redirecting to a
 * presigned S3 URL — keeps the tenant audit trail inside the middleware and
 * avoids surprises with Teams / Kroki not following redirects.
 */
export function createAttachmentsRouter(deps: AttachmentsRouterDeps): Router {
  const router = Router();
  const allowed = deps.allowedPrefixes ?? ['teams-attachments/'];

  router.get('/*key', async (req: Request, res: Response) => {
    const rawSegments = req.params['key'];
    const segments = Array.isArray(rawSegments)
      ? rawSegments
      : typeof rawSegments === 'string'
        ? [rawSegments]
        : [];
    if (segments.length === 0) {
      res.status(400).type('text/plain').send('missing key');
      return;
    }
    const key = segments.join('/');

    if (!allowed.some((p) => key.startsWith(p))) {
      res.status(403).type('text/plain').send('disallowed prefix');
      return;
    }

    const exp = Number(req.query['exp']);
    const sig = typeof req.query['sig'] === 'string' ? req.query['sig'] : '';
    const ok = verifyAttachmentSig({ key, exp, sig, secret: deps.secret });
    if (!ok) {
      res.status(403).type('text/plain').send('invalid or expired signature');
      return;
    }

    try {
      const { stream, contentType, contentLength } = await deps.store.getStream(key);
      res.setHeader('content-type', contentType ?? 'application/octet-stream');
      if (contentLength !== undefined) {
        res.setHeader('content-length', String(contentLength));
      }
      res.setHeader('cache-control', 'private, max-age=900, immutable');
      stream.on('error', (err) => {
        console.error('[attachments] stream error:', err);
        if (!res.headersSent) {
          res.status(500).type('text/plain').send('internal error');
        } else {
          res.destroy(err);
        }
      });
      stream.pipe(res);
    } catch (err) {
      if (isNotFound(err)) {
        res.status(404).type('text/plain').send('not found');
        return;
      }
      console.error('[attachments] proxy failed:', err);
      res.status(500).type('text/plain').send('internal error');
    }
  });

  return router;
}

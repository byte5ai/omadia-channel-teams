import { createHash } from 'node:crypto';
import type { Pool } from 'pg';
import type { Attachment } from 'botbuilder';
import type { TigrisStore } from '@omadia/diagrams';
import { signAttachmentUrl } from './attachmentSigning.js';
import type { GraphClient } from '@omadia/integration-microsoft365';

/**
 * Persists Teams message attachments to Tigris and indexes them in Neon
 * (`teams_attachments` — migration 0008).
 *
 * Scope today: pure storage. No OCR, no text extraction, no ingest into
 * the knowledge graph. Future consumers can query the table by
 * conversation / sha256 and stream the blob via Tigris on demand.
 *
 * Attachment sources we recognise:
 *   1. Teams file-download-info (SharePoint / OneDrive) — `contentType ===
 *      'application/vnd.microsoft.teams.file.download.info'`. `content.downloadUrl`
 *      is a pre-signed URL (~5 min TTL). No bot auth needed.
 *   2. Inline images / direct-URL attachments — `contentType` starts with
 *      `image/` or `application/`. `contentUrl` points at a Teams CDN that
 *      requires bot JWT auth. Today we attempt anonymous fetch and log
 *      failures; wiring the adapter's JWT is a later step.
 *
 * All failures are soft: a failing download / upload logs on stderr and
 * skips the row. The Teams reply is NEVER blocked by attachment
 * persistence.
 */

export interface TeamsAttachmentStoreOptions {
  tigris: TigrisStore;
  pool: Pool;
  tenant: string;
  /** Tigris key prefix; keeps uploads out of the diagram namespace. */
  keyPrefix?: string;
  /** Cap per file. Anything bigger is skipped + logged. */
  maxBytes?: number;
  /** HTTP timeout for the source download. */
  downloadTimeoutMs?: number;
  log?: (msg: string) => void;
  fetchImpl?: typeof fetch;
  /**
   * Optional Graph client. Enables downloading SharePoint / OneDrive
   * channel attachments that arrive as sharing URLs rather than
   * pre-signed download URLs. Requires the bot app to have the
   * `Files.Read.All` application permission. Without it we still handle
   * 1:1 file.download.info attachments fine — just skip channel uploads
   * and log the skip.
   */
  graphClient?: GraphClient;
  /**
   * Optional signing config. When present, `persistTurn` returns a signed
   * URL that can be handed to Kroki / embedded into diagram specs. Missing
   * config means the URL stays undefined and the caller must rely on the
   * storage key alone.
   */
  signing?: {
    secret: string;
    publicBaseUrl: string;
    ttlSec: number;
  };
}

export interface PersistTurnInput {
  conversationId: string;
  userId?: string;
  turnTime?: Date;
  attachments: readonly Attachment[];
}

export interface PersistedAttachment {
  storageKey: string;
  fileName: string;
  contentType: string;
  sizeBytes: number;
  sha256: string;
  source: AttachmentSource;
  /** Signed proxy URL, present iff the store was configured with a signing key. */
  signedUrl?: string;
}

type AttachmentSource =
  | 'teams.file'
  | 'teams.inline_image'
  | 'teams.sharepoint_link'
  | 'other';

const TEAMS_FILE_CONTENT_TYPE =
  'application/vnd.microsoft.teams.file.download.info';

const DEFAULTS = {
  keyPrefix: 'teams-attachments',
  maxBytes: 25 * 1024 * 1024, // 25 MB — generous but still caps PDFs
  downloadTimeoutMs: 30_000,
};

export class TeamsAttachmentStore {
  private readonly tigris: TigrisStore;
  private readonly pool: Pool;
  private readonly tenant: string;
  private readonly keyPrefix: string;
  private readonly maxBytes: number;
  private readonly downloadTimeoutMs: number;
  private readonly log: (msg: string) => void;
  private readonly fetchImpl: typeof fetch;
  private readonly signing?: {
    secret: string;
    publicBaseUrl: string;
    ttlSec: number;
  };
  private readonly graphClient?: GraphClient;

  constructor(opts: TeamsAttachmentStoreOptions) {
    this.tigris = opts.tigris;
    this.pool = opts.pool;
    this.tenant = opts.tenant;
    this.keyPrefix = opts.keyPrefix ?? DEFAULTS.keyPrefix;
    this.maxBytes = opts.maxBytes ?? DEFAULTS.maxBytes;
    this.downloadTimeoutMs = opts.downloadTimeoutMs ?? DEFAULTS.downloadTimeoutMs;
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
    this.fetchImpl = opts.fetchImpl ?? fetch;
    if (opts.signing) this.signing = opts.signing;
    if (opts.graphClient) this.graphClient = opts.graphClient;
  }

  /**
   * Fallback path for Teams **group-chat** uploads: when the inbound
   * Activity's `attachments[]` only carries the mention-HTML (no file-
   * download-info), we re-fetch the full message via Graph using the RSC
   * `ChatMessage.Read.Chat` permission. The Graph representation
   * includes proper attachments with SharePoint `contentUrl`s — we
   * return them shaped as botbuilder-`Attachment` objects so the
   * existing `resolveAttachment` + Graph-download pipeline takes over
   * untouched.
   */
  async discoverFromGraphMessage(
    chatId: string,
    messageId: string,
  ): Promise<Array<{ contentType: string; contentUrl: string; name: string }>> {
    if (!this.graphClient) return [];
    // Adaptive-card submits and some invoke activities carry synthetic IDs
    // like `f:<uuid>` that Graph's /chats/{id}/messages/{id} endpoint rejects
    // with HTTP 400. Skip them silently — these activities never have
    // uploadable attachments anyway (they are button clicks).
    if (/^[a-z]:/i.test(messageId)) return [];
    try {
      const { attachments } = await this.graphClient.fetchChatMessage(
        chatId,
        messageId,
      );
      const out: Array<{ contentType: string; contentUrl: string; name: string }> =
        [];
      for (const att of attachments) {
        // Graph's `contentType === 'reference'` marks file links. Anything
        // else (`messageReference`, `codeSnippet`, …) is not a file.
        if (att.contentType !== 'reference') continue;
        if (!att.contentUrl) continue;
        out.push({
          contentType: '',
          contentUrl: att.contentUrl,
          name: att.name,
        });
      }
      this.log(
        `[teams-attachments] graph-discover chat=${chatId.slice(0, 40)} msg=${messageId.slice(0, 40)} raw=${String(attachments.length)} files=${String(out.length)}`,
      );
      return out;
    } catch (err) {
      this.log(
        `[teams-attachments] graph-discover FAIL chat=${chatId.slice(0, 40)} err=${errMsg(err)}`,
      );
      return [];
    }
  }

  /**
   * Build a fresh signed URL for a previously stored key. Returns undefined
   * if signing isn't configured. Used when the bot reads the Logo from
   * memory and needs a current, short-lived URL to embed in a diagram.
   */
  signUrl(key: string): string | undefined {
    if (!this.signing) return undefined;
    return signAttachmentUrl({
      key,
      secret: this.signing.secret,
      publicBaseUrl: this.signing.publicBaseUrl,
      ttlSec: this.signing.ttlSec,
    });
  }

  /**
   * Persist every recognisable attachment on the given turn. Returns the
   * successfully stored ones — callers can use the list to build reply
   * confirmations. Never throws; errors land in the log only.
   */
  async persistTurn(
    input: PersistTurnInput,
  ): Promise<PersistedAttachment[]> {
    this.log(
      `[teams-attachments] turn conv=${input.conversationId} attachments=${String(input.attachments.length)}`,
    );
    const out: PersistedAttachment[] = [];
    for (const att of input.attachments) {
      const resolved = resolveAttachment(att);
      if (!resolved) {
        this.log(
          `[teams-attachments] skip unsupported contentType=${String(att.contentType ?? '-')} name=${String(att.name ?? '-')} contentUrlPrefix=${truncateUrl(att.contentUrl)} contentKeys=${contentKeys(att.content)} contentPreview=${contentPreview(att.content)}`,
        );
        continue;
      }
      try {
        const stored = await this.persistOne(input, resolved);
        if (stored) out.push(stored);
      } catch (err) {
        this.log(
          `[teams-attachments] persist FAIL name=${resolved.name} src=${resolved.source} err=${errMsg(err)}`,
        );
      }
    }
    return out;
  }

  private async persistOne(
    input: PersistTurnInput,
    att: ResolvedAttachment,
  ): Promise<PersistedAttachment | undefined> {
    // Channel uploads arrive as SharePoint sharing URLs that require Graph
    // auth. Everything else (1:1 file-download-info, inline images with
    // pre-signed URLs) works with a plain fetch.
    let bytes: Buffer | undefined;
    let resolvedContentType = att.contentType;
    let resolvedName = att.name;
    if (att.source === 'teams.sharepoint_link') {
      if (!this.graphClient) {
        this.log(
          `[teams-attachments] sharepoint link but no Graph client configured — skip url=${att.downloadUrl.slice(0, 120)}`,
        );
        return undefined;
      }
      try {
        const graph = await this.graphClient.downloadBySharingUrl(
          att.downloadUrl,
        );
        bytes = graph.bytes;
        resolvedContentType = graph.contentType || att.contentType;
        if (graph.fileName && (resolvedName === 'unnamed' || resolvedName === '')) {
          resolvedName = graph.fileName;
        }
      } catch (err) {
        this.log(
          `[teams-attachments] graph download FAIL name=${att.name} err=${errMsg(err)}`,
        );
        return undefined;
      }
    } else {
      bytes = await this.download(att.downloadUrl);
    }
    if (!bytes) return undefined;
    if (bytes.byteLength > this.maxBytes) {
      this.log(
        `[teams-attachments] skip oversize name=${resolvedName} bytes=${String(bytes.byteLength)} cap=${String(this.maxBytes)}`,
      );
      return undefined;
    }

    const sha256 = hash(bytes);
    const ext = inferExtension(resolvedName, resolvedContentType);
    const turnIso = (input.turnTime ?? new Date()).toISOString();
    const key = buildKey(
      this.keyPrefix,
      this.tenant,
      input.conversationId,
      turnIso,
      sha256,
      ext,
    );

    await this.tigris.put(key, bytes, resolvedContentType);

    await this.pool.query(
      `INSERT INTO teams_attachments
         (tenant, conversation_id, user_id, turn_time, storage_key,
          file_name, content_type, size_bytes, sha256, source,
          source_url, teams_unique_id)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)
       ON CONFLICT (storage_key) DO NOTHING`,
      [
        this.tenant,
        input.conversationId,
        input.userId ?? null,
        input.turnTime ?? null,
        key,
        resolvedName,
        resolvedContentType,
        bytes.byteLength,
        sha256,
        att.source,
        att.downloadUrl,
        att.teamsUniqueId ?? null,
      ],
    );

    const signedUrl = this.signUrl(key);
    this.log(
      `[teams-attachments] stored conv=${input.conversationId} name=${resolvedName} bytes=${String(bytes.byteLength)} sha256=${sha256.slice(0, 12)} src=${att.source} key=${key}${signedUrl ? ' (signed)' : ''}`,
    );
    return {
      storageKey: key,
      fileName: resolvedName,
      contentType: resolvedContentType,
      sizeBytes: bytes.byteLength,
      sha256,
      source: att.source,
      ...(signedUrl ? { signedUrl } : {}),
    };
  }

  private async download(url: string): Promise<Buffer | undefined> {
    const ctrl = new AbortController();
    const timeout = setTimeout(
      () => ctrl.abort(new Error('download timeout')),
      this.downloadTimeoutMs,
    );
    try {
      const response = await this.fetchImpl(url, { signal: ctrl.signal });
      if (!response.ok) {
        this.log(
          `[teams-attachments] download HTTP ${String(response.status)} url=${url.slice(0, 120)}`,
        );
        return undefined;
      }
      const arr = await response.arrayBuffer();
      return Buffer.from(arr);
    } catch (err) {
      this.log(
        `[teams-attachments] download FAIL url=${url.slice(0, 120)} err=${errMsg(err)}`,
      );
      return undefined;
    } finally {
      clearTimeout(timeout);
    }
  }
}

// --- attachment normalisation -------------------------------------------

interface ResolvedAttachment {
  name: string;
  contentType: string;
  downloadUrl: string;
  source: AttachmentSource;
  teamsUniqueId?: string;
}

function resolveAttachment(att: Attachment): ResolvedAttachment | undefined {
  const contentType = att.contentType ?? '';
  const name = att.name && att.name.trim().length > 0 ? att.name : 'unnamed';

  if (contentType === TEAMS_FILE_CONTENT_TYPE) {
    const info = att.content as
      | { downloadUrl?: unknown; uniqueId?: unknown; fileType?: unknown }
      | undefined;
    const downloadUrl =
      typeof info?.downloadUrl === 'string' ? info.downloadUrl : '';
    if (!downloadUrl) return undefined;
    const uniqueId = typeof info?.uniqueId === 'string' ? info.uniqueId : undefined;
    const fileType =
      typeof info?.fileType === 'string' ? info.fileType : undefined;
    const resolvedContentType = fileType
      ? mimeFromExtension(fileType) ?? 'application/octet-stream'
      : 'application/octet-stream';
    return {
      name,
      contentType: resolvedContentType,
      downloadUrl,
      source: 'teams.file',
      ...(uniqueId ? { teamsUniqueId: uniqueId } : {}),
    };
  }

  // Inline images / direct-content attachments
  if (typeof att.contentUrl === 'string' && att.contentUrl.length > 0) {
    const url = att.contentUrl;
    const isSharePoint = isSharePointOrOneDriveUrl(url);
    const source: AttachmentSource = isSharePoint
      ? 'teams.sharepoint_link'
      : contentType.startsWith('image/')
        ? 'teams.inline_image'
        : 'other';
    return {
      name,
      contentType: contentType || 'application/octet-stream',
      downloadUrl: url,
      source,
    };
  }

  // HTML fragment with embedded SharePoint/OneDrive link. Teams uses this
  // shape when a file is dropped into a channel post or pasted inline —
  // `content` is a string containing `<img src="…">` or `<a href="…">`
  // pointing at the tenant's SharePoint. We scan for the first
  // SharePoint-like URL and route it through the Graph client downstream.
  if (contentType === 'text/html' && typeof att.content === 'string') {
    const spUrl = firstSharePointUrl(att.content);
    if (spUrl) {
      return {
        name,
        contentType: 'application/octet-stream', // Graph will report the real one
        downloadUrl: spUrl,
        source: 'teams.sharepoint_link',
      };
    }
  }

  return undefined;
}

/**
 * Extract the first SharePoint / OneDrive URL from an HTML fragment.
 * Tolerant — matches `href="…"`, `src="…"`, or bare URLs, all at the
 * same host set we accept elsewhere.
 */
function firstSharePointUrl(html: string): string | undefined {
  const matches = html.match(/https?:\/\/[^\s"'<>]+/g);
  if (!matches) return undefined;
  for (const url of matches) {
    if (isSharePointOrOneDriveUrl(url)) return url;
  }
  return undefined;
}

/**
 * Recognise the various SharePoint / OneDrive host variants that surface in
 * Teams channel messages:
 *   - `<tenant>.sharepoint.com`
 *   - `<tenant>-my.sharepoint.com` (OneDrive for Business)
 *   - `1drv.ms` (shortlinks from personal OneDrive)
 *   - `<tenant>.sharepoint-df.com` (dogfood, defensive)
 */
function isSharePointOrOneDriveUrl(url: string): boolean {
  try {
    const host = new URL(url).host.toLowerCase();
    return (
      host.endsWith('.sharepoint.com') ||
      host.endsWith('.sharepoint-df.com') ||
      host === '1drv.ms' ||
      host.endsWith('.1drv.ms')
    );
  } catch {
    return false;
  }
}

// --- helpers --------------------------------------------------------------

function buildKey(
  prefix: string,
  tenant: string,
  conversationId: string,
  turnIso: string,
  sha256: string,
  ext: string,
): string {
  const safeConv = conversationId.replace(/[^A-Za-z0-9_.:@-]/g, '_').slice(0, 128);
  const shortHash = sha256.slice(0, 16);
  return `${prefix}/${tenant}/${safeConv}/${turnIso}-${shortHash}${ext}`;
}

function inferExtension(name: string, contentType: string): string {
  const m = /\.[A-Za-z0-9]{1,8}$/.exec(name);
  if (m) return m[0].toLowerCase();
  return extFromMime(contentType) ?? '';
}

const MIME_TO_EXT: Readonly<Record<string, string>> = {
  'image/png': '.png',
  'image/jpeg': '.jpg',
  'image/gif': '.gif',
  'image/webp': '.webp',
  'image/svg+xml': '.svg',
  'application/pdf': '.pdf',
  'application/zip': '.zip',
  'text/plain': '.txt',
  'text/csv': '.csv',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
    '.docx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation':
    '.pptx',
};

const EXT_TO_MIME: Readonly<Record<string, string>> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  pdf: 'application/pdf',
  zip: 'application/zip',
  txt: 'text/plain',
  csv: 'text/csv',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

function extFromMime(contentType: string): string | undefined {
  return MIME_TO_EXT[contentType.toLowerCase()];
}

function mimeFromExtension(ext: string): string | undefined {
  const clean = ext.trim().toLowerCase().replace(/^\./, '');
  return EXT_TO_MIME[clean];
}

function hash(bytes: Buffer): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function errMsg(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

function truncateUrl(url: unknown): string {
  if (typeof url !== 'string' || url.length === 0) return '-';
  try {
    const u = new URL(url);
    return `${u.host}${u.pathname.slice(0, 80)}`;
  } catch {
    return url.slice(0, 80);
  }
}

function contentKeys(content: unknown): string {
  if (!content || typeof content !== 'object') return '-';
  const keys = Object.keys(content as Record<string, unknown>).slice(0, 8);
  return keys.length > 0 ? keys.join(',') : '-';
}

function contentPreview(content: unknown): string {
  if (typeof content === 'string') {
    const flat = content.replace(/\s+/g, ' ').trim();
    return flat.length <= 160 ? flat : `${flat.slice(0, 159)}…`;
  }
  if (content && typeof content === 'object') {
    try {
      const json = JSON.stringify(content);
      return json.length <= 160 ? json : `${json.slice(0, 159)}…`;
    } catch {
      return 'obj';
    }
  }
  return '-';
}

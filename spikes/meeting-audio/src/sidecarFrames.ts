// Spike S1' (byte5ai/omadia#584, omadia ADR-0287): the localhost hand-off
// between the headless-Chromium sidecar and the plugin.
//
// SPIKE CODE, not wired. The sidecar (spikes/meeting-audio/sidecar) captures
// the meeting's mixed remote audio, converts it to PCM s16le mono at 24 kHz
// (what transcription@1's realtime provider takes) and streams it, together
// with dominant-speaker changes and participant updates, over one TCP
// connection on 127.0.0.1. Raw audio is never written anywhere (spec 2043
// §6.3); the frames live only in this stream.
//
// Wire format, one frame:
//   magic   4 bytes  "OMAF"
//   version 1 byte   1
//   type    1 byte   1 = audio, 2 = event (UTF-8 JSON)
//   length  4 bytes  payload length, big endian, at most MAX_PAYLOAD_BYTES
//   atMs    4 bytes  offset from the start of capture, big endian
//   payload `length` bytes

export const FRAME_MAGIC = 'OMAF';
export const FRAME_VERSION = 1;
export const FRAME_HEADER_BYTES = 14;
/** One second of 24 kHz s16le mono is 48 000 bytes; a frame is far smaller. */
export const MAX_PAYLOAD_BYTES = 64 * 1024;
export const PCM_SAMPLE_RATE = 24_000;

const TYPE_AUDIO = 1;
const TYPE_EVENT = 2;

/** Events the sidecar reports next to the audio. Ids are ACS raw ids. */
export type SidecarEvent =
  | { readonly kind: 'joined'; readonly selfRawId?: string }
  | { readonly kind: 'floor'; readonly speakerRawId?: string }
  | {
      readonly kind: 'participants';
      readonly participants: readonly { readonly rawId: string; readonly displayName?: string }[];
    }
  | { readonly kind: 'left'; readonly reason: string };

export type SidecarFrame =
  | { readonly type: 'audio'; readonly atMs: number; readonly pcm: Uint8Array }
  | { readonly type: 'event'; readonly atMs: number; readonly event: SidecarEvent };

export class SidecarFrameError extends Error {
  override readonly name = 'SidecarFrameError';
}

export function encodeFrame(frame: SidecarFrame): Uint8Array {
  const payload = frame.type === 'audio' ? frame.pcm : new TextEncoder().encode(JSON.stringify(frame.event));
  if (payload.byteLength > MAX_PAYLOAD_BYTES) throw new SidecarFrameError(`payload too large: ${payload.byteLength}`);
  const out = new Uint8Array(FRAME_HEADER_BYTES + payload.byteLength);
  const view = new DataView(out.buffer);
  out.set(new TextEncoder().encode(FRAME_MAGIC), 0);
  view.setUint8(4, FRAME_VERSION);
  view.setUint8(5, frame.type === 'audio' ? TYPE_AUDIO : TYPE_EVENT);
  view.setUint32(6, payload.byteLength);
  view.setUint32(10, Math.max(0, Math.floor(frame.atMs)));
  out.set(payload, FRAME_HEADER_BYTES);
  return out;
}

const EVENT_KINDS: ReadonlySet<string> = new Set(['joined', 'floor', 'participants', 'left']);

function parseEvent(payload: Uint8Array): SidecarEvent {
  let value: unknown;
  try {
    value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(payload));
  } catch {
    throw new SidecarFrameError('event payload is not UTF-8 JSON');
  }
  if (typeof value !== 'object' || value === null) throw new SidecarFrameError('event is not an object');
  const kind = (value as { kind?: unknown }).kind;
  if (typeof kind !== 'string' || !EVENT_KINDS.has(kind)) throw new SidecarFrameError(`unknown event kind: ${String(kind)}`);
  return value as SidecarEvent;
}

/**
 * Streaming decoder: feed TCP chunks as they come, get whole frames out. A
 * malformed header throws; the caller drops the connection (the sidecar is
 * local and trusted no further than its own process).
 */
export class SidecarFrameDecoder {
  private buffered: Uint8Array = new Uint8Array(0);

  push(chunk: Uint8Array): SidecarFrame[] {
    const joined = new Uint8Array(this.buffered.byteLength + chunk.byteLength);
    joined.set(this.buffered, 0);
    joined.set(chunk, this.buffered.byteLength);
    const frames: SidecarFrame[] = [];
    let offset = 0;
    while (joined.byteLength - offset >= FRAME_HEADER_BYTES) {
      const view = new DataView(joined.buffer, joined.byteOffset + offset);
      const magic = new TextDecoder().decode(joined.subarray(offset, offset + 4));
      if (magic !== FRAME_MAGIC) throw new SidecarFrameError('bad magic');
      if (view.getUint8(4) !== FRAME_VERSION) throw new SidecarFrameError(`unsupported version ${view.getUint8(4)}`);
      const type = view.getUint8(5);
      const length = view.getUint32(6);
      if (length > MAX_PAYLOAD_BYTES) throw new SidecarFrameError(`payload too large: ${length}`);
      if (joined.byteLength - offset < FRAME_HEADER_BYTES + length) break;
      const atMs = view.getUint32(10);
      const payload = joined.slice(offset + FRAME_HEADER_BYTES, offset + FRAME_HEADER_BYTES + length);
      if (type === TYPE_AUDIO) {
        if (length % 2 !== 0) throw new SidecarFrameError('PCM s16le payload has an odd length');
        frames.push({ type: 'audio', atMs, pcm: payload });
      } else if (type === TYPE_EVENT) {
        frames.push({ type: 'event', atMs, event: parseEvent(payload) });
      } else {
        throw new SidecarFrameError(`unknown frame type ${type}`);
      }
      offset += FRAME_HEADER_BYTES + length;
    }
    this.buffered = joined.slice(offset);
    return frames;
  }

  /** Bytes held back waiting for the rest of a frame. */
  get pendingBytes(): number {
    return this.buffered.byteLength;
  }
}

/** Audio duration of a PCM s16le mono payload at {@link PCM_SAMPLE_RATE}. */
export function pcmDurationMs(pcm: Uint8Array): number {
  return (pcm.byteLength / 2 / PCM_SAMPLE_RATE) * 1000;
}

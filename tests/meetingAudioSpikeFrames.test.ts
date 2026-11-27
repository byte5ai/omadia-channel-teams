// Spike S1' (byte5ai/omadia#584): the sidecar → plugin frame stream.
import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  encodeFrame,
  pcmDurationMs,
  SidecarFrameDecoder,
  SidecarFrameError,
  type SidecarFrame,
} from '../spikes/meeting-audio/src/sidecarFrames.js';

const PCM_20MS = new Uint8Array(960).map((_, i) => i % 251);

describe("S1' sidecar frames", () => {
  it('round-trips audio and events', () => {
    const frames: SidecarFrame[] = [
      { type: 'event', atMs: 0, event: { kind: 'joined', selfRawId: '8:orgid:a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d' } },
      { type: 'audio', atMs: 20, pcm: PCM_20MS },
      { type: 'event', atMs: 40, event: { kind: 'floor', speakerRawId: '8:orgid:6f1c2b8e-1d2a-4c3b-9e8f-0a1b2c3d4e5f' } },
    ];
    const decoder = new SidecarFrameDecoder();
    const out = frames.flatMap((f) => decoder.push(encodeFrame(f)));
    assert.deepEqual(out, frames);
    assert.equal(decoder.pendingBytes, 0);
  });

  it('reassembles frames split across TCP chunks and several frames in one chunk', () => {
    const a = encodeFrame({ type: 'audio', atMs: 0, pcm: PCM_20MS });
    const b = encodeFrame({ type: 'event', atMs: 20, event: { kind: 'left', reason: 'meeting_ended' } });
    const all = new Uint8Array([...a, ...b]);
    const decoder = new SidecarFrameDecoder();
    assert.deepEqual(decoder.push(all.subarray(0, 7)), []);
    assert.deepEqual(decoder.push(all.subarray(7, a.byteLength + 3)).length, 1);
    const rest = decoder.push(all.subarray(a.byteLength + 3));
    assert.deepEqual(rest, [{ type: 'event', atMs: 20, event: { kind: 'left', reason: 'meeting_ended' } }]);
  });

  it('rejects malformed input', () => {
    const good = encodeFrame({ type: 'audio', atMs: 0, pcm: PCM_20MS });
    const badMagic = good.slice();
    badMagic[0] = 0x58;
    assert.throws(() => new SidecarFrameDecoder().push(badMagic), SidecarFrameError);

    const oddPcm = encodeFrame({ type: 'audio', atMs: 0, pcm: new Uint8Array(3) });
    assert.throws(() => new SidecarFrameDecoder().push(oddPcm), /odd length/);

    const unknownEvent = encodeFrame({ type: 'event', atMs: 0, event: { kind: 'left', reason: 'x' } });
    const json = new TextEncoder().encode('{"kind":"exfiltrate"}');
    const forged = new Uint8Array(14 + json.byteLength);
    forged.set(unknownEvent.subarray(0, 14));
    new DataView(forged.buffer).setUint32(6, json.byteLength);
    forged.set(json, 14);
    assert.throws(() => new SidecarFrameDecoder().push(forged), /unknown event kind/);

    const huge = good.slice(0, 14);
    new DataView(huge.buffer).setUint32(6, 10_000_000);
    assert.throws(() => new SidecarFrameDecoder().push(huge), /too large/);
    assert.throws(() => encodeFrame({ type: 'audio', atMs: 0, pcm: new Uint8Array(70_000) }), /too large/);
  });

  it('computes the duration of 24 kHz s16le mono', () => {
    assert.equal(pcmDurationMs(PCM_20MS), 20);
  });
});

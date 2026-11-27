// Spike S1' — the page side of the sidecar, bundled for headless Chromium.
//
// Joins a Teams meeting as a Teams user (ACS "Communication as Teams user",
// TeamsCallAgent), muted, takes the mixed remote audio stream, converts it
// to PCM s16le mono 24 kHz in an AudioWorklet (the AudioContext resamples)
// and hands 20 ms frames plus dominant-speaker and participant changes to
// the Node side through the `omadiaEmit` binding. Nothing is stored here.
//
// `selfTest` skips ACS entirely: a 440 Hz oscillator stands in for the
// meeting, which proves the capture path (worklet, resampling, framing)
// headless on the host without any Microsoft resource.

import { CallClient, Features, type TeamsCall } from '@azure/communication-calling';
import {
  AzureCommunicationTokenCredential,
  getIdentifierRawId,
  type CommunicationIdentifierKind,
} from '@azure/communication-common';

export interface PageConfig {
  readonly selfTest: boolean;
  readonly acsToken?: string;
  readonly meetingLink?: string;
  readonly maxSeconds: number;
}

type EmitKind = 'audio' | 'event';

declare global {
  interface Window {
    omadiaEmit(kind: EmitKind, atMs: number, payload: string): Promise<void>;
    omadiaStart(config: PageConfig): Promise<void>;
  }
}

const SAMPLE_RATE = 24_000;
const FRAME_SAMPLES = 480; // 20 ms at 24 kHz

const WORKLET = `
class OmadiaPcm extends AudioWorkletProcessor {
  constructor() { super(); this.buf = new Int16Array(${FRAME_SAMPLES}); this.n = 0; }
  process(inputs) {
    const ch = inputs[0] && inputs[0][0];
    if (!ch) return true;
    for (let i = 0; i < ch.length; i++) {
      const s = Math.max(-1, Math.min(1, ch[i]));
      this.buf[this.n++] = s < 0 ? s * 0x8000 : s * 0x7fff;
      if (this.n === this.buf.length) { this.port.postMessage(this.buf.buffer.slice(0)); this.n = 0; }
    }
    return true;
  }
}
registerProcessor('omadia-pcm', OmadiaPcm);
`;

const started = performance.now();
const elapsed = (): number => Math.round(performance.now() - started);

function toBase64(buffer: ArrayBuffer): string {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  for (let i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i] ?? 0);
  return btoa(binary);
}

function emitEvent(event: Record<string, unknown>): void {
  void window.omadiaEmit('event', elapsed(), JSON.stringify(event));
}

async function captureStream(stream: MediaStream): Promise<AudioContext> {
  const ctx = new AudioContext({ sampleRate: SAMPLE_RATE });
  const url = URL.createObjectURL(new Blob([WORKLET], { type: 'application/javascript' }));
  await ctx.audioWorklet.addModule(url);
  const node = new AudioWorkletNode(ctx, 'omadia-pcm', { channelCount: 1, channelCountMode: 'explicit' });
  node.port.onmessage = (msg: MessageEvent<ArrayBuffer>) => void window.omadiaEmit('audio', elapsed(), toBase64(msg.data));
  ctx.createMediaStreamSource(stream).connect(node);
  // The worklet only pulls while connected to the graph; a muted gain keeps it running without playing.
  const sink = ctx.createGain();
  sink.gain.value = 0;
  node.connect(sink).connect(ctx.destination);
  return ctx;
}

async function selfTest(maxSeconds: number): Promise<void> {
  const source = new AudioContext({ sampleRate: 48_000 });
  const osc = source.createOscillator();
  osc.frequency.value = 440;
  const dest = source.createMediaStreamDestination();
  osc.connect(dest);
  osc.start();
  emitEvent({ kind: 'joined' });
  emitEvent({ kind: 'participants', participants: [{ rawId: '8:orgid:00000000-0000-4000-8000-000000000001' }] });
  emitEvent({ kind: 'floor', speakerRawId: '8:orgid:00000000-0000-4000-8000-000000000001' });
  const ctx = await captureStream(dest.stream);
  await new Promise((resolve) => setTimeout(resolve, maxSeconds * 1000));
  await ctx.close();
  await source.close();
  emitEvent({ kind: 'left', reason: 'self_test_done' });
}

function rawIdOf(id: CommunicationIdentifierKind): string {
  return getIdentifierRawId(id);
}

function reportParticipants(call: TeamsCall): void {
  emitEvent({
    kind: 'participants',
    // displayName is not set for Teams users by the SDK; it is a label at best.
    participants: call.remoteParticipants.map((p) => ({ rawId: rawIdOf(p.identifier), displayName: p.displayName })),
  });
}

async function joinMeeting(config: PageConfig): Promise<void> {
  if (!config.acsToken || !config.meetingLink) throw new Error('acsToken and meetingLink are required');
  const agent = await new CallClient().createTeamsCallAgent(new AzureCommunicationTokenCredential(config.acsToken));
  const call = agent.join({ meetingLink: config.meetingLink }, { audioOptions: { muted: true } });
  let capture: AudioContext | undefined;
  const startCapture = async (): Promise<void> => {
    const stream = call.remoteAudioStreams[0];
    if (capture || !stream) return;
    capture = await captureStream(await stream.getMediaStream());
  };
  const cap = setTimeout(() => void call.hangUp(), config.maxSeconds * 1000);
  call.on('stateChanged', () => {
    if (call.state === 'Connected') {
      emitEvent({ kind: 'joined' });
      reportParticipants(call);
      void startCapture();
    }
    if (call.state === 'Disconnected') {
      clearTimeout(cap);
      void capture?.close();
      emitEvent({ kind: 'left', reason: `ended:${call.callEndReason?.code ?? '?'}/${call.callEndReason?.subCode ?? '?'}` });
    }
  });
  call.on('remoteAudioStreamsUpdated', () => void startCapture());
  call.on('remoteParticipantsUpdated', () => reportParticipants(call));
  const dominant = call.feature(Features.DominantSpeakers);
  dominant.on('dominantSpeakersChanged', () => {
    const head = dominant.dominantSpeakers.speakersList[0];
    emitEvent({ kind: 'floor', ...(head ? { speakerRawId: rawIdOf(head) } : {}) });
  });
}

window.omadiaStart = async (config: PageConfig): Promise<void> => {
  try {
    await (config.selfTest ? selfTest(config.maxSeconds) : joinMeeting(config));
  } catch (err) {
    emitEvent({ kind: 'left', reason: `error:${err instanceof Error ? err.message : String(err)}` });
  }
};

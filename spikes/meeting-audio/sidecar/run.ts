// Spike S1' — the Node side of the sidecar.
//
//   TEAMS_MEETING_AUDIO_SPIKE=1 node out/run.mjs --self-test [--seconds 3]
//   TEAMS_MEETING_AUDIO_SPIKE=1 ACS_TOKEN=… MEETING_LINK=… node out/run.mjs \
//       [--seconds 300] [--sink 127.0.0.1:7777] [--headed]
//
// Launches headless Chromium with fake media devices, serves the bundled page
// on 127.0.0.1, and turns what the page emits into sidecar frames
// (spikes/meeting-audio/src/sidecarFrames.ts). With `--sink` the frames go to a
// TCP listener on loopback (see sink.ts, standing in for the plugin);
// without, they are decoded in process and only counted. Audio is never
// written to disk. Secrets come from the environment, never from argv.

import { createServer, type Server } from 'node:http';
import { readFile } from 'node:fs/promises';
import { connect, type Socket } from 'node:net';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { chromium, type Browser } from 'playwright';

import {
  encodeFrame,
  pcmDurationMs,
  SidecarFrameDecoder,
  type SidecarEvent,
  type SidecarFrame,
} from '../src/sidecarFrames.js';
import { assertMeetingAudioSpikeEnabled } from '../src/teamsUserToken.js';

interface Args {
  readonly selfTest: boolean;
  readonly seconds: number;
  readonly sink?: { readonly host: string; readonly port: number };
  readonly headed: boolean;
}

const LOOPBACK = new Set(['127.0.0.1', 'localhost', '::1']);
const DEFAULT_SELF_TEST_SECONDS = 3;
const DEFAULT_MEETING_SECONDS = 300;
const MAX_SECONDS = 3600;
const PAGE_READY_TIMEOUT_MS = 20_000;
const WALL_CLOCK_GRACE_SECONDS = 30;
const CHROMIUM_ARGS = [
  '--use-fake-ui-for-media-stream',
  '--use-fake-device-for-media-stream',
  '--autoplay-policy=no-user-gesture-required',
  '--disable-dev-shm-usage',
];

function parseArgs(argv: readonly string[]): Args {
  const value = (flag: string): string | undefined => {
    const i = argv.indexOf(flag);
    return i >= 0 ? argv[i + 1] : undefined;
  };
  const selfTest = argv.includes('--self-test');
  const seconds = Number(value('--seconds') ?? (selfTest ? DEFAULT_SELF_TEST_SECONDS : DEFAULT_MEETING_SECONDS));
  if (!Number.isFinite(seconds) || seconds <= 0 || seconds > MAX_SECONDS) throw new Error(`--seconds must be 1..${MAX_SECONDS}`);
  const sinkRaw = value('--sink');
  let sink: Args['sink'];
  if (sinkRaw) {
    const at = sinkRaw.lastIndexOf(':');
    const host = sinkRaw.slice(0, at);
    const port = Number(sinkRaw.slice(at + 1));
    if (!LOOPBACK.has(host) || !Number.isInteger(port)) throw new Error('--sink must be a loopback host:port');
    sink = { host, port };
  }
  return { selfTest, seconds, headed: argv.includes('--headed'), ...(sink ? { sink } : {}) };
}

async function servePage(): Promise<{ server: Server; url: string }> {
  const here = dirname(fileURLToPath(import.meta.url));
  const pageJs = await readFile(join(here, 'page.js'));
  const server = createServer((req, res) => {
    if (req.url === '/page.js') {
      res.writeHead(200, { 'content-type': 'application/javascript' }).end(pageJs);
      return;
    }
    res.writeHead(200, { 'content-type': 'text/html' }).end('<!doctype html><script type="module" src="/page.js"></script>');
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  if (!address || typeof address === 'string') throw new Error('no loopback port');
  // The Calling SDK refuses any origin but https, file: or `localhost` (127.0.0.1 is rejected).
  return { server, url: `http://localhost:${address.port}/` };
}

interface Stats {
  frames: number;
  audioMs: number;
  sumSquares: number;
  samples: number;
  events: SidecarEvent[];
}

function account(stats: Stats, frame: SidecarFrame): void {
  stats.frames += 1;
  if (frame.type === 'event') {
    stats.events.push(frame.event);
    return;
  }
  stats.audioMs += pcmDurationMs(frame.pcm);
  const view = new DataView(frame.pcm.buffer, frame.pcm.byteOffset, frame.pcm.byteLength);
  for (let i = 0; i + 1 < frame.pcm.byteLength; i += 2) {
    const s = view.getInt16(i, true) / 0x8000;
    stats.sumSquares += s * s;
    stats.samples += 1;
  }
}

function rmsDbfs(stats: Stats): number {
  if (stats.samples === 0) return Number.NEGATIVE_INFINITY;
  return Math.round(20 * Math.log10(Math.sqrt(stats.sumSquares / stats.samples)) * 10) / 10;
}

async function openSink(sink: Args['sink']): Promise<Socket | undefined> {
  if (!sink) return undefined;
  const socket = connect(sink.port, sink.host);
  await new Promise<void>((resolve, reject) => socket.once('connect', resolve).once('error', reject));
  return socket;
}

async function main(): Promise<void> {
  assertMeetingAudioSpikeEnabled(process.env);
  const args = parseArgs(process.argv.slice(2));
  const acsToken = process.env.ACS_TOKEN;
  const meetingLink = process.env.MEETING_LINK;
  if (!args.selfTest && (!acsToken || !meetingLink)) throw new Error('ACS_TOKEN and MEETING_LINK are required (or --self-test)');

  const { server, url } = await servePage();
  const socket = await openSink(args.sink);
  const decoder = new SidecarFrameDecoder();
  const stats: Stats = { frames: 0, audioMs: 0, sumSquares: 0, samples: 0, events: [] };
  const started = Date.now();
  let browser: Browser | undefined;
  try {
    browser = await chromium.launch({ headless: !args.headed, args: CHROMIUM_ARGS });
    const page = await browser.newPage();
    let leftResolve: (reason: string) => void = () => undefined;
    const left = new Promise<string>((resolve) => (leftResolve = resolve));
    await page.exposeFunction('omadiaEmit', (kind: 'audio' | 'event', atMs: number, payload: string) => {
      const frame: SidecarFrame =
        kind === 'audio'
          ? { type: 'audio', atMs, pcm: new Uint8Array(Buffer.from(payload, 'base64')) }
          : { type: 'event', atMs, event: JSON.parse(payload) as SidecarEvent };
      const bytes = encodeFrame(frame);
      if (socket) socket.write(bytes);
      for (const decoded of decoder.push(bytes)) account(stats, decoded);
      if (frame.type === 'event' && process.env.SPIKE_EVENTS === '1') console.error(JSON.stringify({ atMs, event: frame.event }));
      if (frame.type === 'event' && frame.event.kind === 'left') leftResolve(frame.event.reason);
    });
    page.on('pageerror', (err) => console.error(`[page] ${err.message}`));
    page.on('console', (msg) => process.env.SPIKE_DEBUG === '1' && console.error(`[console] ${msg.text()}`));
    await page.goto(url);
    await page.waitForFunction(() => typeof window.omadiaStart === 'function', undefined, { timeout: PAGE_READY_TIMEOUT_MS });
    await page.evaluate(
      (config) => window.omadiaStart(config),
      { selfTest: args.selfTest, maxSeconds: args.seconds, ...(acsToken ? { acsToken } : {}), ...(meetingLink ? { meetingLink } : {}) },
    );
    let capTimer: NodeJS.Timeout | undefined;
    const reason = await Promise.race([
      left,
      new Promise<string>((resolve) => {
        capTimer = setTimeout(() => resolve('wall_clock_cap'), (args.seconds + WALL_CLOCK_GRACE_SECONDS) * 1000);
      }),
    ]);
    clearTimeout(capTimer);
    const summary = {
      mode: args.selfTest ? 'self-test' : 'meeting',
      leftReason: reason,
      wallMs: Date.now() - started,
      frames: stats.frames,
      audioSeconds: Math.round(stats.audioMs) / 1000,
      rmsDbfs: rmsDbfs(stats),
      floorChanges: stats.events.filter((e) => e.kind === 'floor').length,
      participantUpdates: stats.events.filter((e) => e.kind === 'participants').length,
      chromium: browser.version(),
    };
    console.log(JSON.stringify(summary));
    const ok = args.selfTest
      ? summary.audioSeconds >= args.seconds * 0.8 && summary.rmsDbfs > -40
      : !reason.startsWith('error:');
    process.exitCode = ok ? 0 : 1;
  } finally {
    socket?.end();
    await browser?.close();
    server.close();
  }
}

main().catch((err: unknown) => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exitCode = 1;
});

// Spike S1' — stand-in for the plugin end of the hand-off: a loopback TCP
// listener that decodes sidecar frames and prints one line per second.
//
//   node out/sink.mjs [port]      (default 7777, binds 127.0.0.1 only)
//
// The audio is counted, never stored.

import { createServer } from 'node:net';

import { pcmDurationMs, SidecarFrameDecoder } from '../src/sidecarFrames.js';

const port = Number(process.argv[2] ?? 7777);

createServer((socket) => {
  const decoder = new SidecarFrameDecoder();
  let audioMs = 0;
  const timer = setInterval(() => console.log(JSON.stringify({ audioSeconds: Math.round(audioMs) / 1000 })), 1000);
  socket.on('data', (chunk: Buffer) => {
    try {
      for (const frame of decoder.push(chunk)) {
        if (frame.type === 'audio') audioMs += pcmDurationMs(frame.pcm);
        else console.log(JSON.stringify({ atMs: frame.atMs, event: frame.event }));
      }
    } catch (err) {
      console.error(`dropping connection: ${err instanceof Error ? err.message : String(err)}`);
      socket.destroy();
    }
  });
  socket.on('close', () => clearInterval(timer));
}).listen(port, '127.0.0.1', () => console.log(`sink on 127.0.0.1:${port}`));

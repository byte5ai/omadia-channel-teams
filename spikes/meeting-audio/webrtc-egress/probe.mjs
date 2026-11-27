#!/usr/bin/env node
// Spike S4' (byte5ai/omadia#584, omadia ADR-0287): can this host reach the
// Teams/ACS media relays the way a WebRTC client must?
//
//   1. UDP 3478: a plain STUN Binding request (RFC 5389) to the Teams relay.
//      A Binding Success proves outbound UDP and returns the host's public
//      mapped address (its NAT egress).
//   2. TCP 443 + TLS: the fallback path the Calling SDK uses when UDP is
//      blocked (worse quality: TCP turns loss into delay).
//
// No credentials, no Microsoft tenant resource, no dependency: run it with
// plain Node on the host to test (Fly machine, Azure VM, laptop):
//   node probe.mjs [--json]
// It sends a few hundred bytes and prints one line per check.

import { createSocket } from 'node:dgram';
import { randomBytes } from 'node:crypto';
import { lookup } from 'node:dns/promises';
import { connect } from 'node:tls';

const RELAYS = ['worldaz.tr.teams.microsoft.com', 'euaz.tr.teams.microsoft.com'];
const STUN_PORT = 3478;
const TLS_PORT = 443;
const TIMEOUT_MS = 3000;
const ATTEMPTS = 3;

const STUN_BINDING_REQUEST = 0x0001;
const STUN_BINDING_SUCCESS = 0x0101;
const STUN_MAGIC_COOKIE = 0x2112a442;
const ATTR_XOR_MAPPED_ADDRESS = 0x0020;

function bindingRequest(transactionId) {
  const msg = Buffer.alloc(20);
  msg.writeUInt16BE(STUN_BINDING_REQUEST, 0);
  msg.writeUInt16BE(0, 2);
  msg.writeUInt32BE(STUN_MAGIC_COOKIE, 4);
  transactionId.copy(msg, 8);
  return msg;
}

/** XOR-MAPPED-ADDRESS (IPv4 only; enough for a NAT check). */
function mappedAddress(msg) {
  let offset = 20;
  while (offset + 4 <= msg.length) {
    const type = msg.readUInt16BE(offset);
    const length = msg.readUInt16BE(offset + 2);
    if (type === ATTR_XOR_MAPPED_ADDRESS && msg[offset + 5] === 0x01) {
      const port = msg.readUInt16BE(offset + 6) ^ (STUN_MAGIC_COOKIE >>> 16);
      const ip = (msg.readUInt32BE(offset + 8) ^ STUN_MAGIC_COOKIE) >>> 0;
      return `${ip >>> 24}.${(ip >>> 16) & 255}.${(ip >>> 8) & 255}.${ip & 255}:${port}`;
    }
    offset += 4 + length + ((4 - (length % 4)) % 4);
  }
  return undefined;
}

function stunOnce(address) {
  return new Promise((resolve) => {
    const socket = createSocket('udp4');
    const transactionId = randomBytes(12);
    const started = Date.now();
    const done = (result) => {
      clearTimeout(timer);
      socket.close();
      resolve(result);
    };
    const timer = setTimeout(() => done({ ok: false, error: 'timeout' }), TIMEOUT_MS);
    socket.on('error', (err) => done({ ok: false, error: err.code ?? err.message }));
    socket.on('message', (msg) => {
      if (msg.length < 20 || !msg.subarray(8, 20).equals(transactionId)) return;
      const type = msg.readUInt16BE(0);
      done(
        type === STUN_BINDING_SUCCESS
          ? { ok: true, rttMs: Date.now() - started, mapped: mappedAddress(msg) }
          : { ok: false, error: `stun type 0x${type.toString(16)}` },
      );
    });
    socket.send(bindingRequest(transactionId), STUN_PORT, address);
  });
}

async function stunProbe(host) {
  const { address } = await lookup(host, { family: 4 });
  let last;
  for (let attempt = 1; attempt <= ATTEMPTS; attempt += 1) {
    last = await stunOnce(address);
    if (last.ok) return { check: 'udp-stun', host, address, attempt, ...last };
  }
  return { check: 'udp-stun', host, address, attempts: ATTEMPTS, ...last };
}

function tlsProbe(host) {
  return new Promise((resolve) => {
    const started = Date.now();
    const socket = connect({ host, port: TLS_PORT, servername: host, timeout: TIMEOUT_MS });
    const done = (result) => {
      socket.destroy();
      resolve({ check: 'tcp-tls-443', host, ...result });
    };
    socket.once('secureConnect', () =>
      done({ ok: true, rttMs: Date.now() - started, cn: socket.getPeerCertificate()?.subject?.CN }),
    );
    socket.once('timeout', () => done({ ok: false, error: 'timeout' }));
    socket.once('error', (err) => done({ ok: false, error: err.code ?? err.message }));
  });
}

async function main() {
  const results = [];
  for (const host of RELAYS) {
    for (const probe of [stunProbe, tlsProbe]) {
      try {
        results.push(await probe(host));
      } catch (err) {
        results.push({ check: probe.name, host, ok: false, error: err.code ?? err.message });
      }
    }
  }
  if (process.argv.includes('--json')) {
    console.log(JSON.stringify({ at: new Date().toISOString(), results }));
  } else {
    for (const r of results) console.log(JSON.stringify(r));
  }
  process.exitCode = results.every((r) => r.ok) ? 0 : 1;
}

await main();

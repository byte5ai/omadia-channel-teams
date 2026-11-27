# Meeting-audio spikes (byte5ai/omadia#584, omadia ADR-0287)

Prototype code for the spikes S1', S4', A-S1 and S3'. Results:
[`docs/meeting-audio-spikes.md`](../../docs/meeting-audio-spikes.md).

None of this is part of the plugin. `tsconfig.json` compiles only `src/`, so nothing here
lands in `dist/` or in the plugin zip. The unit tests run with the plugin's suite:
`tests/meetingAudioSpike*.test.ts`.

| Path | Spike |
|---|---|
| `src/participantIdentity.ts`, `src/speakerAttribution.ts` | S3': participant id to roster, and T1 attribution |
| `src/teamsUserToken.ts` | A-S1: refresh-token grant and the ACS `teamsUser/:exchangeAccessToken` |
| `src/sidecarFrames.ts` | S1': the loopback frame format from the sidecar to the plugin |
| `sidecar/` | S1': headless Chromium with the ACS Calling SDK (Playwright) |
| `webrtc-egress/probe.mjs` | S4': UDP STUN and TLS 443 reachability of the Teams relays |

The token calls and the sidecar refuse to run without `TEAMS_MEETING_AUDIO_SPIKE=1`. The
egress probe is not gated: it sends only an unauthenticated STUN request and a TLS handshake,
and touches no tenant. Secrets come only from the environment. Audio is never written to disk.

## Run

```sh
# S4': no dependencies, no credentials
node spikes/meeting-audio/webrtc-egress/probe.mjs

# S1' self-test: no Microsoft resource; an oscillator stands in for the meeting
cd spikes/meeting-audio/sidecar
npm ci && npx playwright install chromium-headless-shell
npm run self-test

# S1' live (needs Marcel: test user, ACS resource, test meeting)
TEAMS_MEETING_AUDIO_SPIKE=1 ACS_TOKEN=… MEETING_LINK='https://teams.microsoft.com/l/meetup-join/…' \
  node out/run.mjs --seconds 120 [--sink localhost:7777]
node out/sink.mjs 7777   # in a second shell, stands in for the plugin end
```

To get the `ACS_TOKEN` for the live run (A-S1), call `refreshEntraToken` and then
`exchangeForAcsToken` from `src/teamsUserToken.ts`, with:

- the test user's refresh token from a one-time device-code login;
- the public-client app id;
- the user's object id;
- the ACS endpoint and key.

What to record from a live run: the `leftReason`, `audioSeconds` against the wall time,
`floorChanges`, the participant raw ids against the chat roster, and the RSS of the process
tree.

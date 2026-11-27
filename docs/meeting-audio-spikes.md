# Meeting audio: spike results S1', S4', A-S1, S3'

Spikes for byte5ai/omadia#584, gating omadia ADR-0287 (Proposed). The owner approved
these four spikes on 2026-11-27 and nothing else. No production code, no licence
decision, and no tenant users were created. Code: [`spikes/meeting-audio/`](../spikes/meeting-audio/README.md).
The plugin's `src/` and the shipped `dist/` are unchanged.

Labels as in [`teams-meeting-apis.md`](./teams-meeting-apis.md):

- **Docs:** a Microsoft Learn page, fetched 2026-11-27. Sources are listed at the end.
- **Measured:** run on a dev Mac on 2026-11-27.
- **Simulated:** covered by tests.
- **Open:** needs Marcel. These are the steps that need real Microsoft resources or production hosts.

## Summary

| Spike | Question | Result | Still open (needs Marcel) |
|---|---|---|---|
| S1' | Can a headless Chromium join as a Teams user and capture the meeting audio? | The capture path works headless. Measured: 2.98 s of 24 kHz PCM in 3 s, level correct, about 476 MB RSS. The SDK (`@azure/communication-calling` 2.0.0) loads headless and rejects a bad token cleanly. The join API is documented without a preview label. **Raw remote audio is preview.** | Join a real test meeting with a test user's token. |
| S4' | Is WebRTC egress available from Fly and the Azure VM? | Measured from the Mac: UDP STUN on 3478 and TLS on 443 to the Teams relays both answer. Azure's default NSG allows outbound UDP (Docs). Fly documents nothing about outbound UDP. | Run `probe.mjs` on a Fly machine and on the Azure VM. The auto-mode classifier blocked `fly ssh` to production for this session. |
| A-S1 | Can we get a delegated token for the agent's tenant user without an interactive login? | The ACS exchange is implemented and simulated, with an HMAC-signed `teamsUser/:exchangeAccessToken`. Refresh token from a one-time device-code login: **feasible**. ROPC: **rejected**. Entra Agent ID agent user: **preferred, unproven**. | Create a test user (or an agent user), consent the app, and run the exchange once for real. |
| S3' | Can a participant id be mapped to the roster? | Yes, by Entra object id only (`8:orgid:<oid>`). The match returns the roster's own `userRef`, so speech gets the same principal as that person's chat. The SDK sets no `displayName` for Teams users. Simulated. | Check on the real participant list during the S1' live run. |

## S1': headless join and audio capture

**How it works.** The sidecar is in `spikes/meeting-audio/sidecar`:

1. Playwright launches Chromium headless, with fake media devices and `--disable-dev-shm-usage`.
2. It serves the bundled page on `localhost`.
3. The page calls `CallClient().createTeamsCallAgent(token)`, then joins with `join({ meetingLink }, { audioOptions: { muted: true } })`.
4. It reads `call.remoteAudioStreams[0].getMediaStream()`.
5. An `AudioContext({ sampleRate: 24000 })` resamples the audio. An AudioWorklet turns it into 20 ms s16le frames.
6. `Features.DominantSpeakers` reports the floor, and `remoteParticipants` reports who is in the call.
7. Frames go to the plugin over loopback TCP. The format is in `spikes/meeting-audio/src/sidecarFrames.ts`. Nothing is written to disk.

**Findings.**

- **Measured:** a self-test in which an oscillator stands in for the meeting audio, run in Chromium 159 headless shell:
  - 3 s run: 153 frames, 2.98 s of audio, −3.1 dBFS (a sine wave should read −3.0).
  - 10 s run: 9.98 s of audio.
  - The process tree (Node, browser, renderer, audio, GPU) used **about 476 MB RSS** during capture. A real call adds WebRTC on top, so plan for at least 1 GB. Fly community guidance is 2 GB to be safe.
- **Measured:** the SDK refuses any origin other than `https`, `file:` or `localhost`. `http://127.0.0.1` fails with "ACS Web Calling SDK must be used through https, file, or localhost". The sidecar serves on `localhost` for that reason.
- **Measured:** an invalid token fails locally with "Invalid token" before any network join.
- **Docs:** joining a Teams meeting as a Teams user is listed as supported in the JS SDK, and the how-to has no preview banner. Only one TeamsCallAgent is allowed per CallClient. [S1][S2]
- **Docs:** `remoteAudioStreams[].getMediaStream()` is part of the raw media access feature, which is **preview** [S3]. Microsoft describes it as "the mixed incoming remote audio stream … of the top four most dominant speakers" [S4]. That fits attribution T1. It also means a fifth voice in heavy crosstalk is not in the stream at all.
- **Docs:** the SDK sets no `displayName` for Teams users [S2]. Names must come from the roster or from Graph, keyed by object id (see S3').
- **Docs, new since ADR-0287:** Microsoft retires ACS as a standalone offering on **2028-09-30** [S5]:
  - Calling SDK use stays supported together with a Teams-aligned scenario, and Teams meeting interop counts as one.
  - **ACS Chat for Teams interop is retired.** The meeting chat must be written through **Graph** (`POST /chats/{id}/messages`, delegated `ChatMessage.Send`) with a Graph-scoped token of the same user, not through ACS.
  - Server-side WebSocket media streaming is not covered by the interop exception.
- **Docs, pricing:** ACS bills the agent's media at **$0.004 per participant-minute**, lobby time included, with no extra interop fee [S6]. A 60-minute meeting costs about $0.24 for media plus $1.02 for `gpt-live-transcribe` (#584).

**Needs Marcel.** Running the harness against a real meeting needs:

- a test tenant user with a Teams licence (see A-S1);
- an ACS resource;
- a scheduled test meeting the user is invited to.

The command is in the spike README.

## S4': WebRTC egress

**Docs [S7][S8].** Media goes to `20.202.0.0/16`, `52.112.0.0/14` and `52.122.0.0/15` on UDP 3478–3481, with TCP 443 as a fallback. Signalling runs on TCP 443. Audio-only needs about 40 kbps.

When UDP is blocked, the SDK falls back to TCP. Microsoft calls that a last resort because of the quality loss. No inbound port is needed: the client always connects out to the media processor, so NAT type does not matter for meetings.

**Measured on the dev Mac, 2026-11-27.** `spikes/meeting-audio/webrtc-egress/probe.mjs` sends an unauthenticated STUN Binding request and opens a TLS connection:

| Check | `worldaz.tr.teams.microsoft.com` | `euaz.tr.teams.microsoft.com` |
|---|---|---|
| UDP 3478 STUN Binding | success, 23 ms, public mapped address returned | success, 34 ms |
| TCP 443 TLS | ok, 66 ms, `CN=tr.teams.microsoft.com` | ok, 101 ms |

**Azure VM (Docs).** The default NSG rule `AllowInternetOutBound` allows every outbound protocol and port [S9]. VMs created in a private subnet without "default outbound access" need a NAT gateway or a public IP.

**Fly (Docs, partial).** The documentation covers only inbound UDP (`fly-global-services`), which the sidecar does not need. Outbound UDP and the NAT type are undocumented. The egress IPv4 is NATed and "may change"; a static egress IP can be allocated per app and region [S10].

**Needs Marcel.** Run the probe on the target hosts. The auto-mode classifier blocked this session from running commands on the production machine.

```sh
B64=$(base64 -i spikes/meeting-audio/webrtc-egress/probe.mjs | tr -d '\n')
fly ssh console -a odoo-bot-harness -C "sh -c 'echo $B64 | base64 -d > /tmp/p.mjs && node /tmp/p.mjs; rm /tmp/p.mjs'"
az vm run-command invoke -g RG-OMADIA-PROD-WEU -n vm-omadia-prod-weu --command-id RunShellScript \
  --scripts "echo $B64 | base64 -d > /tmp/p.mjs && node /tmp/p.mjs; rm /tmp/p.mjs"
```

How to read the result:

- **Pass:** `udp-stun` is `ok:true` on both hosts.
- **TCP only:** `udp-stun` fails and `tcp-tls-443` is ok. The sidecar works, but with degraded audio. Prefer the Azure VM.

## A-S1: a token for the agent's tenant user without an interactive login

The chain is: an Entra token of the user for `https://auth.msft.communication.azure.com/Teams.ManageCalls` and `…/Teams.ManageChats`, then `POST {acs}/teamsUser/:exchangeAccessToken?api-version=2023-10-01` with `{ token, appId, userId }`.

ACS checks that the token's `appid` and `oid` match the body, and returns an ACS token that expires with the Entra token [S11][S12].

The prototype is `spikes/meeting-audio/src/teamsUserToken.ts` and is **Simulated**. It covers:

- the refresh-token grant, including the rolled refresh token;
- the HMAC-SHA256 request signature;
- error mapping without echoing secrets;
- the `TEAMS_MEETING_AUDIO_SPIKE=1` gate.

How to get the Entra token without a person present:

| Option | Verdict | Why |
|---|---|---|
| ROPC (password grant) | **Rejected** | Microsoft: "do not use"; incompatible with MFA. The MSAL and Azure Identity APIs for it are deprecated. Any security-defaults or Conditional Access MFA requirement blocks it [S13]. |
| One device-code login by an admin when the user is provisioned, then the refresh token in the vault | **Feasible now** | 90-day rolling lifetime, renewed on every use. It is lost on a password reset, on "revoke sessions" and under a sign-in-frequency policy [S14]. Needs a public-client app registration with the two ACS permissions and admin consent. |
| **Entra Agent ID agent user** | **Preferred, unproven** | GA. A user-type identity with no password that cannot sign in interactively. It gets delegated (`idtyp=user`) tokens through its agent identity and the blueprint's credential (FIC, managed identity or certificate). It can hold licences [S15][S16]. Open: whether `exchangeAccessToken` and the Calling SDK accept such a user. Microsoft supports the token flow only through its own libraries or the Auth SDK sidecar ("don't implement the token exchange directly"), and has no Node flow. Microsoft 365 / Teams provisioning of agent users goes through Teams / Agent 365. |

**Needs Marcel.**

1. Create one test user, or an agent user, in a test or byte5 tenant, with an eligible Teams plan.
2. Register a public-client app with the two ACS delegated permissions, and grant admin consent.
3. Sign in once with the device code.
4. Run the exchange (`TEAMS_MEETING_AUDIO_SPIKE=1`).
5. Use the resulting ACS token for the S1' live run.

## S3': participant id to roster

The prototype is `spikes/meeting-audio/src/participantIdentity.ts` and `speakerAttribution.ts`, with 11 tests. **Simulated.**

- **Binding rule.** Only the Entra object id binds. `8:orgid:<oid>` is matched against roster `userRef.id` (kind `teams-aad`) or `externalId`, case-insensitively. The roster's own `userRef` (with its e-mail) is returned, so `channelUserPrincipal` gives the same principal as for that person's chat messages.
- **Names never bind.** The SDK leaves `displayName` empty for Teams users anyway.
- **Other participants:**
  - The agent's own tenant user resolves to `self` and is never a speaker.
  - Other apps resolve to `app`.
  - Anonymous visitors (`8:teamsvisitor:`) and ACS identities resolve to `anonymous`.
  - Dial-in callers resolve to `phone`. The number is not carried.
  - Sovereign-cloud ids (`8:gcch:`, `8:dod:`) never match the public roster.
  - A tenant user who is not on the roster resolves to `tenant_user_unlisted` with its oid. This happens when someone joined the call but not the chat, or when the roster is partial.
- **T1 attribution.** A segment goes to the participant who held the dominant-speaker floor for at least 60 % of it. Otherwise it gets no speaker. Under ADR-0026 an unattributed segment is better than a wrong name.

## Recommendation on the licence question (input for the owner, not a decision)

1. **Buy nothing per agent yet.** One test identity is enough for the live S1' and A-S1 runs. The licence decision follows their result.
2. **If CTE works with an Entra Agent ID agent user, use that** for I1. It has no password to rotate or leak, matches ADR-0012/0134, and can be licensed. Microsoft lists Agent 365 at $15 per user per month, also included in M365 E7. Whether an agent user also needs a separate Teams plan for meeting join is **open**.
3. **Otherwise use a regular cloud-only member user per meeting-enabled agent:**
   - the cheapest plan that contains the Teams service plan (CTE needs any Teams plan; Teams Phone only for PSTN or 1:1 calls to Teams users [S17]);
   - a refresh token from one device-code login, stored in the vault;
   - an exclusion from sign-in-frequency Conditional Access policies (a security decision).
4. **I2 (one shared user) stays the fallback** for tenants that refuse per-agent identities.
5. **Running cost per meeting hour is about $1.26** ($0.24 ACS media + $1.02 transcription), plus the identity's monthly licence.

## Sources (Microsoft Learn unless noted, fetched 2026-11-27)

- [S1] azure/communication-services/concepts/teams-endpoint (updated 2026-10-26)
- [S2] azure/communication-services/how-tos/cte-calling-sdk/manage-calls (updated 2026-09-23)
- [S3] azure/communication-services/quickstarts/voice-video-calling/get-started-raw-media-access (updated 2026-09-23)
- [S4] azure/communication-services/concepts/voice-video-calling/media-access (updated 2026-09-23, preview banner)
- [S5] azure/communication-services/acs-retirement-and-breaking-changes-guide (updated 2026-11-20)
- [S6] azure/communication-services/concepts/pricing/teams-interop-pricing (updated 2026-10-26)
- [S7] azure/communication-services/concepts/voice-video-calling/network-requirements (updated 2026-09-23)
- [S8] microsoft-365/enterprise/urls-and-ip-address-ranges (updated 2026-10-30)
- [S9] azure/virtual-network/network-security-groups-overview (updated 2025-10-23)
- [S10] docs.fly.io/networking/egress-ips and /networking/udp-and-tcp (Fly.io)
- [S11] rest/api/communication/identity/communication-identity/exchange-teams-user-access-token (2026-10-23)
- [S12] azure/communication-services/quickstarts/manage-teams-identity (updated 2026-10-26)
- [S13] entra/identity-platform/v2-oauth-ropc (updated 2026-06-15)
- [S14] entra/identity-platform/refresh-tokens (updated 2026-06-15)
- [S15] entra/agent-id/agent-users and agent-user-oauth-flow (updated 2026-11-17)
- [S16] entra/agent-id/what-is-agent-id-platform (updated 2026-11-17); microsoft.com/microsoft-agent-365 (pricing)
- [S17] azure/communication-services/quickstarts/eligible-teams-licenses (updated 2026-09-23)

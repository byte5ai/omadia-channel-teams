<div align="center">

# @omadia/channel-teams

### Microsoft Teams as a first-class channel for omadia's agents — Bot Framework adapter, Adaptive Cards, roster resolution, signed attachment proxy.

A **Microsoft Teams** channel plugin for [omadia](https://github.com/byte5ai/omadia).
It implements [`@omadia/channel-sdk`](https://github.com/byte5ai/omadia)'s
`ChannelPlugin` contract: CloudAdapter inbound/reply, Adaptive-Card rendering,
roster + @mention resolution, HMAC-signed attachment persistence, and the
`/api/messages` + `/attachments` HTTP routes.

[![License: MIT](https://img.shields.io/badge/License-MIT-black.svg)](LICENSE)
[![TypeScript](https://img.shields.io/badge/built%20with-TypeScript-3178C6.svg?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)

[**omadia**](https://github.com/byte5ai/omadia) · [**M365 connector**](https://github.com/byte5ai/omadia-m365-connector) · [**Website**](https://omadia.ai) · [**Setup**](#setup-one-time)

</div>

---

## How it works

| Concern | Implementation |
|---|---|
| Transport | Microsoft Bot Framework `CloudAdapter`, inbound `/api/messages` + reply |
| Rendering | Adaptive Cards — answer cards, approve/reject cards, Direct-Line buttons, smart follow-up/choice cards |
| Roster | Mention + roster resolution (`teamsRoster.ts`, `teamsMentions.ts`) |
| Attachments | HMAC-signed proxy at `/attachments/<key>`, persisted via `teamsAttachmentStore.ts` |
| Delegation | Every user turn is delegated to the shared Conductor/orchestrator (`chatAgent@^1` capability) |
| Events | Emits `teams.message.posted` / `teams.mention` on the deny-by-default event catalog |
| Shared credentials | Reuses the Azure AD app registration owned by [`@omadia/integration-microsoft365`](https://github.com/byte5ai/omadia-m365-connector) — install that first |

This channel is a pure consumer of the orchestrator's `chatAgent` capability
and optionally the `microsoft365.graph` service (for attachment downloads and
roster lookups); it doesn't publish services of its own.

## Setup (one time)

Teams runs as an Azure Bot reusing the **same Azure AD app** as the Microsoft
365 integration — install [`@omadia/integration-microsoft365`](https://github.com/byte5ai/omadia-m365-connector)
first.

1. [Azure Portal](https://portal.azure.com) → create an **Azure Bot** resource, using the existing app registration (same Microsoft App ID).
2. **Configuration → Messaging endpoint**: `<your-public-base-url>/api/teams/messages`.
3. **Channels** → enable **Microsoft Teams**.
4. Upload the Teams app package (manifest.json + icons as a ZIP) under **Apps → Manage your apps → Upload a custom app**.

Optional knobs (all have safe defaults): OAuth connection name for calendar
SSO (OBO flow), notify-team AAD groupId for bell notifications, attachment
key prefix, notification topic URL, dashboard directory label. See
[`manifest.yaml`](manifest.yaml) for the full field list.

### Optional: resolved names + members on the channels dashboard

The operator channels dashboard shows raw conversation ids (`28:<app-id>`,
`19:…@thread.skype`) until Graph can resolve them. To get group-chat
topics, member display names, and the tenant org name on the catch-all
row, grant the app registration these **application permissions** (admin
consent required): `Chat.Read.All`, `ChatMember.Read.All`,
`TeamMember.Read.All`, `Organization.Read.All`. Without them the plugin
degrades silently to the Bot-Framework-derived labels (one log line per
failure class explains what is missing). Resolution is cached (10 min
TTL) and never blocks the dashboard render.

## Build, typecheck & test

```bash
npm install
npm run typecheck   # tsc --noEmit
npm run build        # tsc
npm test             # esbuild-transpile tests/ → node --test
```

The `@omadia/*` peers (`plugin-api`, `plugin-ui-helpers`, `channel-sdk`,
`orchestrator`, `diagrams`, `integration-microsoft365`) are provided by the
omadia host at runtime. For local typechecking, `tsconfig.json` maps them to
sibling checkouts — `../odoo-bot/middleware/packages/*` and
`../omadia-m365-connector` — so check those out alongside this repo, or
adjust the `paths` entries to your local sources.

## Manifest

See [`manifest.yaml`](manifest.yaml) for the full plugin manifest (event
capabilities, setup fields, permissions, network declarations).

## License

MIT © byte5 GmbH — see [LICENSE](LICENSE).

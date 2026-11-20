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
| Rendering | Adaptive Cards: minimal answer cards (text, discreet AI chip when the kernel asks for it, colleagues' contributions attributed to them, one collapsed "Details" element), `system` notices under their own heading, approve/reject cards, decision and choice cards. Every decoration is declared in the channel SDK's decoration registry; buttons exist only for options the agent proposes, decisions, sign-in and polls |
| Roster | Mention + roster resolution (`teamsRoster.ts`, `teamsMentions.ts`) |
| Attachments | HMAC-signed proxy at `/attachments/<key>`, persisted via `teamsAttachmentStore.ts` |
| Delegation | Every user turn is handed to the kernel's agent runtime: per turn the `channelResolver` decides `bound` (an agent answers) or `reject` (the management layer answers with a `system` notice; there is no fallback agent). The channel only passes text through |
| Addressing | An agent is addressed through its own bot. `/agent <name>` (select), `/agent` (clear the selection) and `/agents` (list) are plain text the kernel's management layer carries out and answers as `system`; the channel hands over the whole message, in a group chat too (no `[Name]:` attribution in front of a command) |
| Selection key | `/agent` is stored per sender and the narrowest key of the turn: the conversation (a channel thread is its own conversation), not the bot-wide catch-all binding; on an agent's own bot (exclusive) the bot's key is the only key |
| Ballots | A card or poll Teams does not show natively (kinds it does not declare, more than 5 questions) is posted by the KERNEL as numbered text, and a matching typed reply becomes the typed `card_answer` (kernel #2086); this plugin renders and parses no ballot itself. It reports the Teams activity id of a sent text as `messageId`, sets `replyToMessageId` only for a message the person really answered (a quoted reply, not the root every message of a channel thread carries) and `addressedToBot` for a mention |
| Typed replies | For an addressed message (mention, the bot's own 1:1) that would start a turn, Teams reports it with `reportConversationMessage` first and waits at most 1.5 s: if the kernel consumed it as the typed answer to an open question, no turn starts (no acknowledgement, no agent reply); a timeout, a failure or a kernel without the method lets the turn run as before |
| Data | `SemanticAnswer.data` (spec §5.13, kernel #2085) is drawn as native tables (at most 50 rows each, then a "weitere Zeilen" line) next to `textWithoutData`; an answer without the fields renders from `text` as before |
| Sender | The routing principal is `user:<email>` once Graph verified the address, else `user:teams-aad:<id>`. `TurnOrigin.principal` stays the Entra object id on purpose: the facilitation engine matches participants by it, and `turnPrincipal` would prefer it over the address |
| Pace | Typing from the start; ONE acknowledgement after 8 s under the answering agent's speaker (its own bot, or with a name line where one bot serves several agents) |
| Events | Emits `teams.message.posted` / `teams.mention` on the deny-by-default event catalog |
| Shared credentials | Reuses the Azure AD app registration owned by [`@omadia/integration-microsoft365`](https://github.com/byte5ai/omadia-m365-connector) — install that first |

This channel is a pure consumer of the agent runtime's `chatAgent` service
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
`TeamMember.Read.All`, `Organization.Read.All`; plus `User.Read.All` for the
**people directory** (Conductor resolves "Eskalation an Teresita Biel" to her
address via Graph `/users` — without it that lookup reports *not configured*
and the Conductor picker keeps its free-text address field). Without them the plugin
degrades silently to the Bot-Framework-derived labels (one log line per
failure class explains what is missing). Resolution is cached (10 min
TTL) and never blocks the dashboard render.

### Optional: facilitators (all messages of a chat)

A facilitator's bot receives every message of a chat it moderates, not only
the ones that mention it. Before it starts it checks three things and says
which one is missing:

1. Its Teams app is installed in the chat.
2. The chat granted the resource-specific permission `ChatMessage.Read.Chat`
   to that app. The grant belongs to the install: an install made from an
   older package does not gain it by itself and has to be upgraded (Graph
   `installedApps/{id}/upgrade` with the full consented set).
3. The bot holds a conversation reference for the chat and its last post
   there did not fail.

The check reads Graph with the **instance's** app registration (the one of the
Microsoft 365 integration), which needs
`TeamsAppInstallation.ReadWriteForChat.All` and
`TeamsAppInstallation.ReadWriteAndConsentForChat.All` (team channels: the
`…ForTeam.All` counterparts). The per-agent bot registrations need no Graph
permission.

A facilitator acknowledges a contribution with a thumbs-up instead of an echo.
The reaction is set by the facilitator's own bot through the Bot Connector
([agent reactions](https://learn.microsoft.com/en-us/microsoftteams/platform/agents-in-teams/agent-reactions))
and needs no additional permission, RSC grant or outbound host: the bot's own
token and the declared Bot Connector hosts (`smba.trafficmanager.net`,
`*.botframework.com`) suffice. Teams allows two reaction changes per second per
bot; a rate-limited reaction is retried and otherwise reported, never posted as
text.

Private reminders are sent by the facilitator's own bot. A person who never
installed that bot's app, or blocked it, gets none; the group then gets a
neutral reminder. See [`docs/teams-meeting-apis.md`](./docs/teams-meeting-apis.md)
for the evidence behind each behaviour and for meetings, and
[`docs/facilitator-live-verification.md`](./docs/facilitator-live-verification.md)
for the owner steps M1 to M6 that are still open.

### Optional: auto-invite agent apps into a team (Graph consent)

With `teams_agent_apps` configured (see [`manifest.yaml`](manifest.yaml)),
the channel installs the listed omadia agent apps into a team automatically
when one of its bots is added there. The install path goes through the
`teamsProvisioner@1` service of the M365 connector and needs two additional
Graph **application permissions** on the shared app registration (admin
consent required — and note that previously granted consent does **not**
stretch to cover newly added scopes: extending an existing app
registration's permissions requires admin consent to be **re-granted** on
that registration, otherwise `/appCatalogs/teamsApps` keeps answering `403`
no matter how often the middleware restarts):

- `AppCatalog.ReadWrite.All` — resolve the agent app in the tenant app
  catalog (`/appCatalogs/teamsApps`)
- `TeamsAppInstallation.ReadWriteForTeam.All` — install the catalog app into
  the target team (`POST /teams/{team-id}/installedApps`)

Grant them in the Azure Portal (**API permissions → Add a permission →
Microsoft Graph → Application permissions**, then **Grant admin consent
for &lt;Tenant&gt;**) or send an admin through the tenant-wide admin-consent
URL (contains only the public client id, never a secret):

```
https://login.microsoftonline.com/<tenant-id>/adminconsent?client_id=<application-client-id>
```

Two field-tested gotchas (from the M365 connector's provisioning work — the
consent lessons live in the connector's `README.md`, section "Renewed admin
consent"; its `docs/teams-provisioner.md` documents the capability contract
itself):

- **Portal/CLI consent sometimes silently fails to apply** (observed with
  `az ad app permission admin-consent`: the command succeeds, Graph keeps
  answering `403`). In that case grant the app roles directly via REST
  `appRoleAssignments`, one call per missing permission, on the app's own
  service principal:

  ```
  POST /servicePrincipals/{app-sp-object-id}/appRoleAssignments
  {
    "principalId": "{app-sp-object-id}",
    "resourceId":  "{graph-sp-object-id}",
    "appRoleId":   "{app-role-id-of-the-missing-permission}"
  }
  ```

  Resolve the Microsoft Graph service principal's object id (`resourceId`)
  via `GET /servicePrincipals(appId='00000003-0000-0000-c000-000000000000')`;
  verify with `GET /servicePrincipals/{app-sp-object-id}/appRoleAssignments`.
- **Restart after consent.** Acquired tokens are cached; newly consented app
  roles only appear in a *fresh* token. Restart the middleware (or wait for
  token expiry) after granting consent — otherwise the `403`s persist even
  though consent is in place.

Without consent the feature degrades gracefully: nothing throws on the
message path — the bot posts a fallback card with per-agent install deep
links (`https://teams.microsoft.com/l/app/<teamsAppId>`). Deep links carry
only public Teams app ids, never credentials.

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

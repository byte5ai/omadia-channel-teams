# Teams APIs for native facilitators: chat, direct messages and meetings

Evaluation for byte5ai/omadia-channel-teams#43 and #44 (epic byte5ai/omadia#2043,
kernel spec §6.1 to §6.3): which Teams API delivers which input or output to a
facilitator, and what this plugin implements in 0.36.0.

## How far to trust this document

There is no Teams test tenant. The findings come from an evaluation on 2026-11-18
that read Microsoft Learn and collected **read-only** evidence from the production
instance in the byte5 tenant (middleware log, read-only SQL, Graph GET with the
instance's own app, directory reads). Nothing was sent into Teams for it. Every
statement carries one label:

| Label | Meaning |
|---|---|
| **Live** | Observed read-only in the byte5 tenant on 2026-11-18. |
| **Docs** | Stated on a Microsoft Learn page fetched that day; cited as [Dn], list at the end. |
| **Simulated** | Implemented here and covered by tests against a simulated Bot Framework, Bot Connector and Graph with payloads in the documented or observed shape. Proves what the plugin does with such a payload, not that Teams sends it. |
| **Open** | Neither observed nor documented. The owner step that settles it (M1 to M6) is named; see [`facilitator-live-verification.md`](./facilitator-live-verification.md). |

## 1. All messages without @mention

| Finding | Evidence |
|---|---|
| With RSC `ChatMessage.Read.Chat` a bot receives "all user messages in the group chat where the app is installed"; with `ChannelMessage.Read.Group` those of the team's channels. | Docs [D1] |
| Non-mention messages arrive today: 97 `drop non-mention … type=groupChat` log lines in 100 minutes against 7 mention turns. | Live |
| Every chat with an omadia agent app shows the grants `ChatMessage.Read.Chat`, `ChatMember.Read.Chat`, `TeamsActivity.Send.Chat` for that app. The age of the chat does not matter, the install does. | Live |
| Channel delivery. The team install shows `ChannelMessage.Read.Group`, but there was no channel traffic to observe. | Docs [D1]; **Open → M1** |
| A 1:1 between two people with an agent app (`19:…@unq.gbl.spaces`, Graph `chatType: oneOnOne`) arrives as `conversationType=groupChat`. | Live |
| Only user messages are delivered. Messages written by other bots never are, also with RSC. | Docs [D1][D5]; Live (after each of 7 bot replies no other bot logged anything) |
| Consent is per install and all-or-nothing. An install made from a package without the permission does not gain it by itself. | Docs [D1][D2]; Live (installs at versions 1.0.1 to 1.3.0 keep the set they were installed with) |
| Tenant switches for app-only RSC exist. Their current value could not be read app-only; in byte5 they are evidently on. For customer tenants this is a setup prerequisite. | Docs [D2]; **Open → M6** |

**Implemented.** The mention gate is lifted only for a bot that holds a message
subscription for that conversation; it then emits every human message as a message
event and starts **no chat turn** there, not for an @mention either: the engine holding
the subscription drives every turn. An empty answer is never posted. Nothing concludes the size of a chat from
`conversationType`. *Simulated.*

## 2. Edits, deletions, reactions

| Finding | Evidence |
|---|---|
| Edit = `messageUpdate` + `channelData.eventType: editMessage`; undelete = `messageUpdate` + `undeleteMessage`; soft delete = `messageDelete` + `softDeleteMessage`. Scope "All". | Docs [D6] |
| That `activity.id` is the id of the original message, and that these events also come for messages the bot only saw through RSC. The samples use the same id; no page states either. | **Open → M2** |
| There is no dedicated edit timestamp; only the activity `timestamp`. | Docs [D6] |
| `messageReaction` "is sent when a user adds or removes reactions to a message, which was sent by your agent". Fields `reactionsAdded` / `reactionsRemoved`, `replyToId` = the message. | Docs [D7] |
| Reactions to **other people's** messages. | **Open → M2** |

**Implemented.** Handlers for all four events; the event is keyed on `activity.id`, the
revision counts up per edit, `occurredAt` is the activity `timestamp`. Reaction events
that arrive are emitted, but **`reactions` is not declared in the capability
snapshot**: only reactions to the bot's own posts are documented, and the snapshot
cannot say "own messages only". An engine may use a reaction to its own post (for
example confirming an echo) and must not build on reactions to participants' messages.
Every such activity is logged with its `eventType` and ids, so M2 is answered by
reading the log. *Simulated.*

## 3. Proactive first-contact direct message

| Finding | Evidence |
|---|---|
| The app must be installed for the user in the personal scope before a proactive message; being a member of a team that has the app also counts. | Docs [D8] |
| Create the conversation with `bot.id`, `members[0].id`, `channelData.tenant.id`. The member id may be the user's Entra object id. A `29:…` id is per bot and cannot be reused across bots. UPN and e-mail never work. | Docs [D8] |
| Service URL: from an inbound activity, or the global `https://smba.trafficmanager.net/teams/` for proactive sends. All 24 stored references use `…/emea/<tenant>/`. | Docs [D8]; Live |
| App not installed for the user → `403` with `ForbiddenOperationException`. Blocked or uninstalled → `403` with `subCode: MessageWritesBlocked` (`errorCode 209`). | Docs [D8]; **Open → M4** for the exact bodies |
| Retry `429`, `412`, `502`, `504` with exponential backoff. Limits per bot per thread: 7 sends/s, 60 per 30 s, 1800 per hour. | Docs [D10] |
| Whether being in a group chat that has the app is enough for a 1:1. | **Open → M4** |
| Neither instance holds a `…ForUser.All` permission, so installing the app for a user on demand, and even checking whether it is installed, is not possible today. | Live |

**Implemented.** `sendToUser(…, { asChannelKey })`: the bot's own stored 1:1, else a
conversation created with the Entra object id. `ForbiddenOperationException` →
`app_not_installed`, `MessageWritesBlocked` → `blocked`, neither is retried. The four
transient statuses are retried at most three times (0.5 s, 1 s, 2 s) and never in a way
that could post twice. Error codes no page names are not matched and end as
`channel_error`. Never another bot. *Simulated.*

**Precondition, not solved here:** the person must have the facilitator's app. Install
on demand would need `TeamsAppInstallation.ReadWriteForUser.All` on the instance's
app; that is an owner decision.

## 4. Several bots in one chat

| Finding | Evidence |
|---|---|
| Each bot installed in a chat receives every human message: 17 messages produced exactly 3 simultaneous log lines each in a chat with four agent apps. | Live |
| Bots never see each other's messages. | Docs [D1]; Live |
| The same app can be installed twice in one chat (organisation and sideloaded). | Live |

**Implemented.** Subscription, events, references and posts are keyed by the receiving
bot. N bots deliver N copies of a message; the kernel stores a message delivered more
than once only once. *Simulated.*

## 5. Readiness check

| Finding | Evidence |
|---|---|
| `GET /chats/{id}/installedApps?$select=id,consentedPermissionSet&$expand=teamsApp(…),teamsAppDefinition($select=version;$expand=bot($select=id))` answers 200; about 35 store apps come back per chat; `consentedPermissionSet` only when selected. | Live |
| `GET /chats/{id}/permissionGrants` answers 200, one row per permission. **`clientAppId` is the Entra application (client) id, `clientId` the service principal object id.** The reference page has them the other way round. | Live (checked against the directory); Docs [D14] is wrong on this |
| `GET /teams/{id}/permissionGrants` and `/installedApps` answer 200. | Live |
| The instance's app holds `TeamsAppInstallation.ReadWriteForChat.All`, `…ReadWriteAndConsentForChat.All` and the team counterparts. The **per-agent bot apps hold no Graph application permissions**. | Live |
| "Can this bot post here" has no read-only answer. | Docs (no such API) |
| `POST /chats/{id}/installedApps/{id}/upgrade` with `consentedPermissionSet` exists; the set must match the app definition exactly, else `400 ResourceSpecificPermissionsMismatch`. Consent semantics of the upgrade are not spelled out. | Docs [D9]; **Open → M5** |

**Implemented.** The check runs with the **instance's** Graph client (the registration
of the Microsoft 365 integration, `microsoft_app_id`), not the facilitator bot's.
Ready = an installation row whose app definition carries the bot **and** a grant with
`clientAppId` = bot app id and `ChatMessage.Read.Chat` (team channel:
`ChannelMessage.Read.Group` on the team, known once the bot received an activity from
the channel). `canPost` = the bot's own conversation reference exists and its last
post there did not fail. Reasons: `no_identity`, `app_not_installed`,
`consent_missing`, `graph_forbidden`, `graph_error`, `graph_unavailable`,
`team_unknown`, `cannot_post`. `repairConsent` performs the Graph upgrade with the set
the installed definition declares. *Simulated; the upgrade and reading the declared
set from `teamsAppDefinition.authorization` are Open → M5.*

## 6. Mentions and returned ids in proactive posts

| Finding | Evidence |
|---|---|
| `<at>Name</at>` in the text plus an `entities` item `{type: "mention", mentioned: {id, name}, text: "<at>Name</at>"}`; the entity text must equal the body text exactly. Users can be mentioned by `29:` id, Entra object id or UPN. | Docs [D5] |
| A send returns a `ResourceResponse` with `id`. If text and attachment go in one activity Teams may split it and then returns no id. | Docs [D6][D16] |
| Whether the returned id equals the id other clients and Graph show for the message. | **Open → M3** |

**Implemented.** Mentions are resolved from the chat roster by Entra object id; a person
not in the chat stays plain text. Text and card are **separate activities**; the outcome
carries `messageId` (the text, or the card when there is no text) and `cardMessageId`.
*Simulated.*

## 7. Meetings

| Capability | Verdict | Requirement and evidence |
|---|---|---|
| Bot in the meeting chat, posts and cards | yes, with condition | `groupChat` scope; added by organiser or presenter. Docs [D17][D18]. One stored reference proves meeting chat activities reach the endpoint as `conversationType=groupChat` (`19:meeting_…@thread.v2`, Graph `chatType: meeting`). Live |
| All meeting chat messages without mention | **Open → M5**, likely yes | Docs only say the RSC of a chat "apply to" its meeting [D2]. Microsoft's own Facilitator app holds exactly this grant in byte5 chats. Live |
| Card answer with a verified answerer | yes, with condition | `Action.Execute` arrives as an `adaptiveCard/action` invoke; `from.id` and `from.aadObjectId` are set by Teams. Anonymous users have no `aadObjectId` but may interact with cards. Docs [D18][D19] |
| Native Teams poll (Polls / Forms) | not possible | No bot API. Docs |
| Roster and meeting details on demand | yes, with condition | RSC `OnlineMeetingParticipant.Read.Chat`; participant lookup only below 350 participants; details also for Meet now, 1:1 and group calls. Docs [D19] |
| Meeting start and end events | yes, with condition | RSC `OnlineMeeting.ReadBasic.Chat`; "supported for scheduled and channel meetings". Docs [D19] |
| Participant join and leave events | yes, with condition | RSC `OnlineMeetingParticipant.Read.Chat` **and** the event subscription in the Developer Portal; scheduled meetings only; the bot must be in the meeting before people join. Docs [D19] |
| In-meeting and targeted notification | yes, with condition | Needs a web view in `validDomains`; targeted needs RSC `OnlineMeetingNotification.Send.Chat`, at most 50 recipients, `29:` ids. Docs [D20] |
| Side panel, stage | yes, with condition | A tab with meeting context; stage needs delegated RSC. Docs [D21] |
| Hand raise | **not possible for a bot** | No bot event. Only delegated client-side RSC in a tab. Docs [D2] |
| Real-time audio and video | **not possible on this stack** | Application-hosted media needs C#/.NET on Windows Server in Azure with a public IP per VM; "C++ and Node.js APIs don't support access to real-time media"; developer preview. Docs [D22][D23] |
| Transcript after the meeting | yes, with condition | RSC `OnlineMeetingTranscript.Read.Chat`. Docs [D2] |

**Scheduled meeting versus ad-hoc call.** Apps are supported in scheduled meetings,
1:1 and group calls, Meet now and webinars; not in instant channel meetings,
shared-channel meetings or end-to-end encrypted calls [D17]. Start/end events exist for
scheduled and channel meetings, join/leave for scheduled meetings only. An ad-hoc call
therefore gives chat and cards and an on-demand roster, and **no lifecycle events**.

**Implemented in 0.36.0 (the baseline).**

1. The meeting chat is the group chat path: same subscription rule, same readiness
   check, posts as the facilitator's own bot.
2. A poll is an Adaptive Card with one `Action.Execute` per option. The click is
   emitted as `card_answer`: `cardAnswer.cardId` = poll id, `cardAnswer.value` =
   `{ optionId }`, `cardAnswer.answererVerified`, `messageId` = the card message,
   `author` = `activity.from`. `answererVerified` is false when Teams states no
   `aadObjectId` (anonymous participant). A poll click never becomes a chat turn; it is acknowledged
   with "Antwort übermittelt." or, when nobody is subscribed, "Diese Abstimmung ist
   nicht mehr aktiv.".
3. The capability snapshot, per conversation:
   - **always:** inputs `text`, `message_edits`, `card_answers`; outputs `text`,
     `mentions`, `cards`, `polls`, `private_messages`. Modality `online` for a meeting
     chat, `chat` otherwise. The roster is available through the kernel's roster
     provider.
   - **not declared yet:** `presence` and meeting lifecycle (see the decision below).
   - **never declared:** `hand_raise`, `transcribed_speech`, `voice`, `screen`,
     `pinned_surface`, and modality `hybrid` (Teams does not tell a bot whether people
     share a room). `reactions` is not declared either, see section 2.

All of this is *Simulated*.

### Decision: the meeting RSC are not added to the app package in 0.36.0

`OnlineMeeting.ReadBasic.Chat`, `OnlineMeetingParticipant.Read.Chat` and
`OnlineMeetingNotification.Send.Chat` are **not** in
`appPackage/manifest.json.template`. Reasons:

- Nothing in this version consumes them. Requesting permissions without a consumer
  widens every agent app for no gain.
- Join and leave events also need an event subscription that is set per app in the
  Developer Portal; whether apps provisioned through Graph can set it is unknown.
- New RSC means a new package version per agent app, and existing installs keep their
  old consent set. Each needs **one Graph upgrade call per (agent app, chat)** with the
  exact consented set, and the consent semantics of that call are unverified (M5).

When they are added (7 RSC declared today, 16 allowed, 10 afterwards):

1. Add the three permissions to the template and publish a new package version per
   agent app to the catalog.
2. For every (agent app, chat):
   `POST /chats/{chat-id}/installedApps/{installation-id}/upgrade` with
   `consentedPermissionSet.resourceSpecificPermissions` equal to the **complete** set of
   the new definition. Needs `TeamsAppInstallation.ReadWriteAndConsentForChat.All`,
   which the instance's app already holds. A set that differs from the definition is
   answered with `400 ResourceSpecificPermissionsMismatch`. `repairConsent` in
   `src/teamsReadiness.ts` does this call for one chat.
3. Set the participant event subscription per app, then add the handlers and declare
   `presence` for scheduled meetings only.

Until then a late joiner is visible when they are **added to the chat** (membership
event) and by their first message. Joining the call through the link without being
added to the chat is invisible.

## Known limits of the implementation

| Limit | Consequence |
|---|---|
| The edit revision the adapter sends is counted in memory. | No consequence: the kernel assigns the stored revision itself (created = 0, each edit or deletion = highest stored + 1) and treats the adapter's number as a hint. |
| Quoted and forwarded messages are recognised by attachment markers (`schema.skype.com/Reply`, `forwardedMessageReference`). | **Open**, from neither a page nor the evaluation. If Teams marks them differently such a message is an ordinary `created` event and the engine cannot tell it is a quote. |
| A team channel is checked against its team only after the bot received an activity from that channel in the running process. | Before that the readiness check answers `team_unknown`. |
| The instance's app registration must be served by this channel (its secret in the vault). | Otherwise the readiness check answers `graph_unavailable`. |

## Sources

Microsoft Learn, fetched 2026-11-18. Numbering follows the evaluation note.

- [D1] <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/channel-messages-with-rsc>
- [D2] <https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/rsc/resource-specific-consent>
- [D5] <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/channel-and-group-conversations>
- [D6] <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/conversation-messages>
- [D7] <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/subscribe-to-conversation-events>
- [D8] <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/send-proactive-messages>
- [D9] <https://learn.microsoft.com/en-us/graph/api/chat-teamsappinstallation-upgrade?view=graph-rest-1.0>
- [D10] <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/rate-limit>
- [D13] <https://learn.microsoft.com/en-us/graph/api/chat-list-installedapps?view=graph-rest-1.0>
- [D14] <https://learn.microsoft.com/en-us/graph/api/resources/resourcespecificpermissiongrant?view=graph-rest-1.0>
- [D15] <https://learn.microsoft.com/en-us/graph/api/chat-list-permissiongrants?view=graph-rest-1.0>
- [D16] <https://learn.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-api-reference?view=azure-bot-service-4.0>
- [D17] <https://learn.microsoft.com/en-us/microsoftteams/platform/apps-in-teams-meetings/teams-apps-in-meetings>
- [D18] <https://learn.microsoft.com/en-us/microsoftteams/platform/apps-in-teams-meetings/build-extensible-conversation-for-meeting-chat> and <https://learn.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-actions>
- [D19] <https://learn.microsoft.com/en-us/microsoftteams/platform/apps-in-teams-meetings/meeting-apps-apis>
- [D20] <https://learn.microsoft.com/en-us/microsoftteams/platform/apps-in-teams-meetings/in-meeting-notification-for-meeting>
- [D21] <https://learn.microsoft.com/en-us/microsoftteams/platform/apps-in-teams-meetings/build-apps-for-teams-meeting-stage>
- [D22] <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/calls-and-meetings/requirements-considerations-application-hosted-media-bots>
- [D23] <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/calls-and-meetings/real-time-media-concepts>

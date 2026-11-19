# Teams APIs for native facilitators: chat, direct messages and meetings

Evaluation for byte5ai/omadia-channel-teams#43 and #44 (epic byte5ai/omadia#2043,
kernel spec §6.1 to §6.3): which Teams API delivers which input or output to a
facilitator, and what this plugin implements in 0.36.x.

## How far to trust this document

There is no Teams test tenant. The findings come from an evaluation on 2026-11-18
that read Microsoft Learn and collected **read-only** evidence from the production
instance in the byte5 tenant (middleware log, read-only SQL, Graph GET with the
instance's own app, directory reads). Nothing was sent into Teams for it. Every
statement carries one label:

| Label | Meaning |
|---|---|
| **Live** | Observed in the byte5 tenant on 2026-11-18: read-only evidence from the running instance, and a live verification the same day (12:17 to 12:41 UTC) with test spaces that had one human member. |
| **Docs** | Stated on a Microsoft Learn page fetched that day; cited as [Dn], list at the end. |
| **Simulated** | Implemented here and covered by tests against a simulated Bot Framework, Bot Connector and Graph with payloads in the documented or observed shape. Proves what the plugin does with such a payload, not that Teams sends it. |
| **Open** | Neither observed nor documented. The owner step that settles it (M1 to M6) is named; see [`facilitator-live-verification.md`](./facilitator-live-verification.md). |

## 1. All messages without @mention

| Finding | Evidence |
|---|---|
| With RSC `ChatMessage.Read.Chat` a bot receives "all user messages in the group chat where the app is installed"; with `ChannelMessage.Read.Group` those of the team's channels. | Docs [D1] |
| Non-mention messages arrive today: 97 `drop non-mention … type=groupChat` log lines in 100 minutes against 7 mention turns. | Live |
| Every chat with an omadia agent app shows the grants `ChatMessage.Read.Chat`, `ChatMember.Read.Chat`, `TeamsActivity.Send.Chat` for that app. The age of the chat does not matter, the install does. | Live |
| Channel delivery: a message without mention in a team channel reached the bot two seconds later (`drop non-mention … type=channel`). The team shows the grants `ChannelMessage.Read.Group` and `TeamsActivity.Send.Group` for the app; an install through the Teams app dialog granted them without a separate consent step. | Docs [D1]; **Live** (M1 done) |
| A 1:1 between two people with an agent app (`19:…@unq.gbl.spaces`, Graph `chatType: oneOnOne`) arrives as `conversationType=groupChat`. | Live |
| Only user messages are delivered. Messages written by other bots never are, also with RSC. | Docs [D1][D5]; Live (after each of 7 bot replies no other bot logged anything) |
| Consent is per install and all-or-nothing. An install made from a package without the permission does not gain it by itself. | Docs [D1][D2]; Live (installs at versions 1.0.1 to 1.3.0 keep the set they were installed with) |
| Tenant switch for RSC: the authorization policy's `permissionGrantPoliciesAssigned` contains `ManagePermissionGrantsForOwnedResource.microsoft-dynamically-managed-permissions-for-chat` and `…-for-team` (the Microsoft-managed default); there is no legacy group-settings override. With that, a normal user-side install grants the RSC permissions. **Prerequisite for customer tenants:** both entries (or the "all application permissions" variant) must be present. | Docs [D2]; **Live** (M6 read) |

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
| The app must be installed for the user before a proactive message. A personal install counts, and so does **membership of a team where the app is installed**. Sharing a **group chat** with the app does **not** count. | Docs [D8]; **Live** (create-conversation: 201 for a personal install and for team membership, 403 for six bots the user only shares group chats with) |
| Create the conversation with `bot.id`, `members[0].id`, `channelData.tenant.id`. The member id may be the user's Entra object id. A `29:…` id is per bot and cannot be reused across bots. UPN and e-mail never work. | Docs [D8] |
| Service URL: from an inbound activity, or the global `https://smba.trafficmanager.net/teams/` for proactive sends. All 24 stored references use `…/emea/<tenant>/`. | Docs [D8]; Live |
| App not installed for the user → `403 {"error":{"code":"BadArgument","message":"Bot is not installed in user's personal scope"}}` on create-conversation. | **Live.** The documentation names `ForbiddenOperationException` instead [D8]; that shape was not seen. |
| User blocked the bot → `403 {"error":{"code":"ConversationBlockedByUser","message":"User blocked the conversation with the bot."}}` on the **send**. Create-conversation still answers `201` while blocked. After unblocking, a send answers `201` again. | **Live.** The documentation names `subCode: MessageWritesBlocked` (`errorCode 209`) instead [D8]; that shape was not seen. |
| Retry `429`, `412`, `502`, `504` with exponential backoff. Limits per bot per thread: 7 sends/s, 60 per 30 s, 1800 per hour. | Docs [D10] |
| Neither instance holds a `…ForUser.All` permission, so installing the app for a user on demand, and even checking whether it is installed, is not possible today. | Live |

**Implemented.** `sendToUser(…, { asChannelKey })`: the bot's own stored 1:1, else a
conversation created with the Entra object id. `403` with "not installed in user's
personal scope" → `app_not_installed`; `403` with `ConversationBlockedByUser` →
`blocked`, also on a first contact, because the block shows on the send. The documented
`ForbiddenOperationException` and `MessageWritesBlocked` are matched as well. Neither
outcome is retried. `BadArgument` alone is not read as "not installed". The four
transient statuses are retried at most three times (0.5 s, 1 s, 2 s) and never in a way
that could post twice. Anything else ends as `channel_error`. Never another bot.
*Simulated with the live response bodies; the product path itself was not run live.*

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

### The roster is one bot's view

*Live* for the cause (first facilitation, 2026-11-18); *Simulated* for the fix.

The member list of a conversation is read through the Bot Connector as ONE bot, and a
bot that is not in the conversation gets none. `getRoster(conversationId, { asChannelKey })`
reads it as the asking agent's own bot (its conversation reference, its credentials) or
answers `undefined`; without `asChannelKey` it is the default bot's view, as before. The
cache is keyed by bot and conversation. For a channel thread the members are asked for
the channel. Each read logs one `[teams-roster] roster …` line with the outcome.

### The member list: status, time and what a member carries

*Simulated* for the status logic. *Docs* and *SDK type* for the fields; which of them a
real tenant fills was **not** captured.

**Status** (kernel spec B §5.5, byte5ai/omadia#2091). `getRoster(conversationId,
{ asChannelKey, fresh })` answers with:

| `status` | Meaning | `participants` | `partial` |
|---|---|---|---|
| `complete` | every page was read without an error | everybody | `false` |
| `partial` | some pages were read, then the read failed | what was read | `true` |
| `unknown` | nothing could be read, or the read came back empty | empty, and **not** an answer | `true` |

- `readAt` is the time the platform answered (ISO 8601 UTC). A cached answer keeps the
  time of its read, so its age is visible.
- `statusDetail` gives the reason when the status is not `complete` (HTTP status and the
  service's text).
- `fresh: true` bypasses the cache. A failing fresh read answers `unknown`; it never
  falls back to the older cached list.
- Cache: a `complete` list is kept five minutes, a `partial` or `unknown` one thirty
  seconds.
- `undefined` still means "this bot holds no reference in the conversation". The kernel
  treats that as unknown.
- The fields are additive. `partial` is unchanged for a kernel that does not know
  `status`.

**Where the list comes from.** The Bot Connector's member list
(`TeamsInfo.getPagedMembers`), read as the asking bot. Graph is not used for it.

| Field of a member | Source | Carried as |
|---|---|---|
| `id` (`29:…`, per bot) | Bot Connector | `externalId` |
| `aadObjectId` | Bot Connector | `userRef.id` (kind `teams-aad`) |
| `name` | Bot Connector | `userRef.displayName` |
| `email` | Bot Connector, when sent | `userRef.email` |
| `userPrincipalName` | Bot Connector, when sent | `userPrincipalName` |

- The SDK type `TeamsChannelAccount` declares `email` and `userPrincipalName` as
  optional. The Microsoft Learn page for the roster ("Get Teams specific context for
  your bot", updated 2026-10-13) lists no member fields and no permission, and does not
  say when the two are present. So: they are passed on whenever the platform sends them,
  and nothing is derived when it does not.
- Paging, from the same page: 50 to 500 members per page for teams and channels; "in
  chats, the full roster is returned by the service".
- The app package requests `ChatMember.Read.Chat` and `ChannelMember.Read.Group`.
  Whether the Bot Connector list needs them is not stated on that page.
- Graph would give an address too: `GET /chats/{id}/members` and
  `GET /teams/{id}/members` show `email` in their response examples (application
  permissions `ChatMember.Read.All` or the RSC `ChatMember.Read.Chat`;
  `TeamMember.Read.Group`, which the app package does **not** request). Neither page
  says whether `email` can be empty. This plugin does not call them for the roster.
- **A participant known only from a message has no address.** `activity.from` carries
  id, Entra object id and name, never a mail address. The first live follow-up stopped
  with `no_addresses` for that reason: an address has to come from the member list, or
  from a directory lookup by Entra object id, which this channel does not do.

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
post there did not fail. Reasons, all from the kernel contract's list: `no_identity`,
`app_not_installed`, `consent_missing`, `package_outdated` (the install is from an
earlier app package: its version `1.<template revision>.<n>` is below the shipped
`TEMPLATE_REVISION`), `graph_forbidden`, `graph_error`, `graph_unavailable` (names the
missing setting or vault secret), `team_unknown`, `cannot_post`. The kernel asks this
check when an engine subscribes and refuses a conversation that is not ready.
`repairConsent` (part of the contract) performs the Graph upgrade with the set the
**published** definition declares; when a step is refused it answers `graph_forbidden`
naming the permission that step needs (`TeamsAppInstallation.ReadForChat.All`,
`AppCatalog.ReadWrite.All` or `TeamsAppInstallation.ReadWriteAndConsentForChat.All`),
and `package_outdated` when the catalog still holds an older package. *Simulated; the
upgrade call itself is Open → M5.*

## 5a. Which conversation id goes where

*Live* for the ids and `channelData` keys (byte5 tenant, 2026-11-18); *Docs* and
*Simulated* for a reply inside a thread and for the post into a thread.

| Conversation | `activity.conversation.id` | Towards the kernel | Towards Graph |
|---|---|---|---|
| Group chat | `19:…@thread.v2` | the same id | `/chats/{id}` |
| Meeting chat | `19:meeting_…@thread.v2` (`conversationType: groupChat`) | the same id | `/chats/{id}` |
| Team channel | `19:…@thread.tacv2;messageid=<root message id>` | the same per-thread id | `/teams/{team.aadGroupId}` |

- **One id towards the kernel.** The turn origin states `activity.conversation.id`
  unchanged (`channelConversationId`). Conversation references, the subscription lookup,
  message events, session events and posts use exactly that id. The session scope is
  **not** that id: for a bot with its own identity it is `teams-<botKey>-<id>`, a storage
  key. An id that starts with `28:<appId>-` is refused by name
  (`unsupported_conversation`).
- **A facilitation in a channel belongs to one thread.** Every root post of a channel is
  its own conversation id. A facilitation started by a mention in a root post covers that
  thread: replies arrive under the same id, the facilitator's posts go into the thread. A
  new root post in the same channel is another conversation and is not covered.
- **Graph gets the team, never the thread.** Installation and grants of a channel live on
  the team. The team comes from `channelData.team.aadGroupId` (present on real channel
  messages) and is stored with the conversation reference, per channel.
- **A thread without its own reference.** When the bot was never addressed in a thread, a
  post for it uses the channel's reference (written when the bot was added to the team)
  with the thread as its conversation.
- **Unverified:** no human reply inside a thread was captured, so "a reply carries the
  thread's id" is from the documentation; the post into a thread and the team consent
  upgrade ran against the simulation only.

## 5b. AI marking per send path

**State after rollout defect D24 (2026-11-19).** *Live* for the observation, *Source* for
the entity shape, *Simulated* for everything this plugin now does. Whether Teams renders
its label for any message of this plugin is **open**.

Three things mark a message, and they were confused:

| | What it is | Who shows it |
|---|---|---|
| Platform label | the `AIGeneratedContent` entity on the message activity | Teams: "AI Generated" next to the agent's name in the message header, tooltip "AI-generated content may be incorrect." |
| Disclosure sentence | the instance's line, folded into the text by the kernel | plain text, always visible |
| Card chip | "✨ AI-generated" inside this plugin's **answer card** | the card itself, always visible |

**What was observed.** With 0.37.0 the facilitator's proactive posts (handshake, reminder,
notice) showed no marking in the Teams web client. "✨ AI-generated" was seen on turn
answers; that text is the answer card's own chip, not Teams' label (which reads "AI
Generated" and sits in the header). So the platform label has **not been seen on any
message of this plugin**, proactive or reply.

**What is known about the cause.**

- Microsoft's Teams SDK sends the entity with `"@id": ""` (`microsoft/teams.ts`,
  `packages/api/src/models/entity/message-entity.ts`: "Must be left blank. This is for Bot
  Framework schema"; `addAiGenerated()` in `activities/activity.ts`; read at commit
  `96ec9329e`, 2026-11-10). This plugin sent it **without** `@id` until 0.38.2. That is
  the one field in which our entity differed from Microsoft's; whether it is why the label
  did not show is not proven.
- Microsoft's page for the label ("AI content labels", updated 2026-10-13,
  `learn.microsoft.com/en-us/microsoftteams/platform/agents-in-teams/integrate-ai/ai-content-labels`)
  gives no JSON, names "personal chats, group chats, and channels", and does not mention
  meeting chats, proactive messages, Adaptive Card messages or app type. The word
  "proactive" does not occur on it.
- By code, a plain-text reply in a turn and a proactive text differ in `inputHint` only.
  No path sets `channelData`. The SDK's `continueConversationAsync` does not rebuild the
  activity; `entities` are sent as built. A captured request is still missing; set
  `TEAMS_LOG_OUTGOING_ACTIVITIES=1` to get one (see below).

**What the plugin does now.** The label entity is attached to every message, in Microsoft's
shape. It counts as a marking only once somebody has **seen** it: the setup switch
`teams_native_ai_label` ("AI marking: rely on the Teams label"), **off by default**.

| Send path | Switch off (default) | Switch on |
|---|---|---|
| Turn answer on a card (also a facilitator's fixed refusal) | card chip; sentence stripped | the same |
| Long turn answer (several plain messages) | sentence at the end of the last part | sentence stripped |
| Blocking choice card | entity only (no chip on that card) | the same |
| Proactive text, poll card, question card, private reminder | `supports.aiLabel` **not** declared, so the kernel adds the sentence to an agent's own posts; the adapter posts the text as handed over | `supports.aiLabel` declared; no sentence |
| Routine delivery, card and plain text | sentence kept | sentence stripped |
| Conductor reminder, approval card | entity only (deterministic kernel text, nothing folded in) | the same |
| `NO_REPLY` (turn, routine), also with the sentence behind it | nothing is posted | nothing is posted |
| Streaming | does not exist | does not exist |

Each row has a test in `tests/teamsAiMarking.test.ts`.

- **Off** can show a message with label **and** sentence, if a client does render the
  label. That is accepted until the label is confirmed: marked twice is a blemish, not
  marked at all is a disclosure failure.
- **Still without a visible marking of their own, in both modes:** the blocking choice
  card, the approval card, the Conductor reminder, and cards without text (poll, question
  card) when the kernel has no text to put the sentence into. They carry the entity
  only. Also unmarked, on purpose: the adapter's fixed texts (acknowledgements, errors),
  which are not model output.

**Capturing an outgoing activity.** With the environment variable
`TEAMS_LOG_OUTGOING_ACTIVITIES=1` every activity a turn sends or updates is logged as
`[teams] outgoing path=turn|proactive|direct|update activity={…}`, exactly as the SDK
hands it to the Bot Connector. The lines contain message text and names, so the flag is
for a diagnosis, not for normal operation.

## 5c. Question cards

*Docs* and *Simulated*. Nothing here was seen in a Teams client; the open points are
the items Q1 to Q13 in `facilitator-live-verification.md`. Design: task note "Design:
questionnaire cards in Teams" (2026-11-19); kernel spec B §6.1 and byte5ai/omadia#2089.

**What a card is.** The engine sends `TargetedMessage.card` (a `QuestionCard`: id,
proposal id and version, title, one to five questions, closing time). The adapter posts
it as an Adaptive Card 1.5 with one submit button (`Action.Execute`).

| Kind | On the card | Typed answer |
|---|---|---|
| `single_choice` | radio buttons up to six options, a dropdown above | `{ optionId }` |
| `multiple_choice` | checkboxes | `{ optionIds }` |
| `ranking` | one dropdown per place, up to six options | `{ ranking }`, best first |
| `consent` | agree / object, and a reason field | `{ decision, reason? }` |
| `assignment` | the people picker of the conversation (or a given list), and "Ich übernehme das" | `{ assignee }`: `self` or a person |
| `text` | a text field; `append` (idea collection) or `replace` | `{ text }` |

`number` and `date` are in the contract and not built; a card with one is refused. The
provider declares `supports.card` with exactly these kinds, five questions and `update`.

**Three views of one card.**

- The question card is what everybody sees. No names, no running tally.
- The own-answer card is returned to the person who clicked and shown to them only, with
  "Ändern" while the card is changeable. It is kept in memory per person; after a
  middleware restart a refresh shows the question card again (the engine still holds the
  answer).
- The closed card replaces the posted message when the engine asks for it
  (`updateCard`): the engine's result lines, no inputs. Only a message that the same
  bot posted as a question card in that conversation is updated; another bot's card, a
  text message or an unknown id is refused. The cards are remembered in memory, so
  after a restart of the Teams channel an update of an earlier card is refused and the
  engine's result text in the chat stands.

**Per-person refresh.** The card carries `refresh` without `userIds`. Teams then refreshes
automatically for conversations of up to 60 members and offers "Refresh Card" above that.
The own-answer card returned to a click does not depend on it.

**Every rule is enforced on the server.** `isRequired`, `errorMessage` and `maxLength`
are set on the inputs as a convenience; whether a client honours them for
`Action.Execute` is not documented. A submission is read again against the card:
unknown option, a place used twice, a missing reason, too long a text, an empty click.
A refused submission emits nothing and comes back as the same card with the person's
values kept and the reason shown, to that person only.

**What a click is checked against.**

| Click | Answer |
|---|---|
| card clicked in another conversation than it was posted into | refused, nothing emitted |
| click on an earlier version of the proposal | the current card with a notice, nothing emitted |
| click after the close update, or after `closesAt` | the closed card / "nicht mehr aktiv", nothing emitted |
| click when no facilitation holds the conversation, or sent to another bot | "nicht mehr aktiv" |
| an author named in the click's data | ignored; the author is `activity.from` |
| a click Teams does not attribute to a person | emitted with `answererVerified: false`, and the person is told |

The card a click carries comes from the client. While this process remembers the card it
posted, that one counts. After a restart the click is read against the card it carries,
and the engine, which knows what it sent, drops what does not match.

**The event.** One `card_answer` per accepted submission: `cardId`, `proposalId`,
`proposalVersion`, `answers[]` (one typed value per answered question),
`answererVerified`; `value` repeats `{ answers }` for readers that only know it.
`messageId` is the card's message id.

**Chat ballots stay.** A typed answer in the chat is as valid as a click; the engine lets
the newest answer of a person win. The 0.36.0 poll (`message.poll`) is unchanged.

**Not built:** a card in a direct message (the targeted send takes text only), the
step-by-step questionnaire, a personal view that survives a restart.

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
| Bot in the meeting chat, posts and cards | yes, with condition | `groupChat` scope; added by organiser or presenter (in the chat: "add agents"). Docs [D17][D18]. Meeting chat activities reach the endpoint as `conversationType=groupChat` (`19:meeting_…@thread.v2`, Graph `chatType: meeting`). **Live** |
| All meeting chat messages without mention | **yes** | A message without mention in the chat of a scheduled meeting reached the bot (`drop non-mention conv=19:meeting_… type=groupChat`). Graph shows the chat as `chatType: meeting` with the grants `ChatMessage.Read.Chat`, `ChatMember.Read.Chat`, `TeamsActivity.Send.Chat`. **Live** (first half of M5 done). Docs only say the RSC of a chat "apply to" its meeting [D2]. |
| Card answer with a verified answerer | yes, with condition | `Action.Execute` arrives as an `adaptiveCard/action` invoke; `from.id` and `from.aadObjectId` are set by Teams. Anonymous users have no `aadObjectId` but may interact with cards. Docs [D18][D19] |
| Native Teams poll (Polls / Forms) | not possible | No bot API. Docs |
| Roster and meeting details on demand | yes, with condition | RSC `OnlineMeetingParticipant.Read.Chat`; participant lookup only below 350 participants; details also for Meet now, 1:1 and group calls. Docs [D19] |
| Meeting start and end events | yes, with condition | RSC `OnlineMeeting.ReadBasic.Chat`; "supported for scheduled and channel meetings". Docs [D19] |
| Participant join and leave events | yes, with condition | RSC `OnlineMeetingParticipant.Read.Chat` **and** the event subscription in the Developer Portal (no manifest field or API documented); scheduled meetings only; the bot must be in the meeting before people join. Docs [D19] |
| In-meeting and targeted notification | possible, **not built** | Needs a web page of the app in `validDomains`; targeted needs RSC `OnlineMeetingNotification.Send.Chat`, at most 50 recipients, `29:` ids only. Docs [D20] |
| Side panel, stage | yes, with condition | A tab with meeting context; stage needs delegated RSC. Docs [D21] |
| Hand raise | **not possible for a bot** | No bot event. Only delegated client-side RSC in a tab. Docs [D2] |
| Real-time audio and video | **not possible on this stack** | Application-hosted media needs C#/.NET on Windows Server in Azure with a public IP per VM; "C++ and Node.js APIs don't support access to real-time media"; developer preview. Docs [D22][D23] |
| Transcript after the meeting | yes, with condition | RSC `OnlineMeetingTranscript.Read.Chat`. Docs [D2] |

**Scheduled meeting versus ad-hoc call.** Apps are supported in scheduled meetings,
1:1 and group calls, Meet now and webinars; not in instant channel meetings,
shared-channel meetings or end-to-end encrypted calls [D17]. Start/end events exist for
scheduled and channel meetings, join/leave for scheduled meetings only. An ad-hoc call
therefore gives chat and cards and an on-demand roster, and **no lifecycle events**.

**Implemented in 0.36.x.**

1. The meeting chat is the group chat path: same subscription rule, same readiness
   check, posts as the facilitator's own bot.
2. A poll is an Adaptive Card with one `Action.Execute` per option. The click is
   emitted as `card_answer`: `cardAnswer.cardId` = poll id, `cardAnswer.value` =
   `{ optionId }`, `cardAnswer.answererVerified`, `messageId` = the card message,
   `author` = `activity.from`. `answererVerified` is false when Teams states no
   `aadObjectId` (anonymous participant). A poll click never becomes a chat turn; it is
   acknowledged with "Antwort übermittelt." or, when nobody is subscribed, "Diese
   Abstimmung ist nicht mehr aktiv.".
3. **Meeting events.** The four `event` activities are handled and passed to the kernel
   as session events through `core.emitSessionEvent` (feature-detected):

   | Teams activity name | Session event | Needs |
   |---|---|---|
   | `application/vnd.microsoft.meetingStart` | `session_started` | RSC `OnlineMeeting.ReadBasic.Chat` (channel meeting: `ChannelMeeting.ReadBasic.Group`) |
   | `application/vnd.microsoft.meetingEnd` | `session_ended` | same |
   | `application/vnd.microsoft.meetingParticipantJoin` | `participant_joined` | RSC `OnlineMeetingParticipant.Read.Chat` **and** the Developer Portal subscription |
   | `application/vnd.microsoft.meetingParticipantLeave` | `participant_left` | same |

   Shape: `{ channelId, channelType, conversationId, sessionRef, kind, participants?,
   occurredAt }`. `sessionRef` is `channelData.meeting.id` (Teams says not to use
   `value.Id` of the event; without a meeting id the meeting chat id is used).
   `participants` are `value.members[].user` with the Entra object id. `occurredAt` is
   `value.StartTime` / `value.EndTime` for start and end, else the activity timestamp.
   The documentation prints the start and end names with a leading space; names are
   trimmed before dispatch. Docs [D19] for names and fields; *Simulated*, nothing
   observed live.
4. **The capability snapshot**, per conversation and bot:
   - **always:** inputs `text`, `message_edits`, `card_answers`; outputs `text`,
     `mentions`, `cards`, `polls`, `private_messages`. Modality `online` for a meeting
     chat or while a meeting or call is running, `chat` otherwise. The roster is
     available through the kernel's roster provider.
   - **`presence`, only while it is proven.** All three must hold, none is assumed:
     a scheduled meeting is running in this chat (its start event said `Scheduled` or
     `Recurring`, or a participant event arrived, which exists for scheduled meetings
     only); a join or leave event has actually arrived for this meeting; and the install
     in this chat holds both `OnlineMeeting.ReadBasic.Chat` and
     `OnlineMeetingParticipant.Read.Chat`, read from Graph. It is withdrawn when the
     meeting ends, and every occurrence has to prove it again.
   - **Meet-now, group calls, 1:1 calls:** the "always" set. Teams sends them no
     participant events.
   - **never declared:** `hand_raise`, `transcribed_speech`, `voice`, `screen`,
     `pinned_surface`, and modality `hybrid` (Teams does not tell a bot whether people
     share a room). `reactions` is not declared either, see section 2.
   - The meeting lifecycle has no capability name in the kernel contract. It reaches the
     kernel as the session events above and as a snapshot pushed at start, at end, and
     at the first participant event of a meeting.

All of this is *Simulated*.

### The meeting permissions in the app package

Since `TEMPLATE_REVISION` 2 the app package requests `OnlineMeeting.ReadBasic.Chat`,
`OnlineMeetingParticipant.Read.Chat`, `OnlineMeetingNotification.Send.Chat` and
`ChannelMeeting.ReadBasic.Group`. All four install into a team and into a chat (*Live*,
2026-11-18). Revision 3 (0.36.2) removes `FileStorageContainer.Selected.Group`, which the
team consent check refuses and nothing used; the package now requests ten RSC permissions
(sixteen are allowed). See "RSC permissions: install proof" in
`facilitator-live-verification.md`.

- **Who has to agree with the template.** Nobody keeps a second list. The M365
  connector's provisioner reads the consented set at install from the **published**
  catalog definition
  (`GET /appCatalogs/teamsApps?$filter=id eq '…'&$expand=appDefinitions($select=id,authorization)`)
  and sends it verbatim; `repairConsent` here uses the same query (plus the definition's
  `version`). A set that differs
  from the definition is answered with `400 ResourceSpecificPermissionsMismatch`, so
  neither side derives it from the template text. The connector needed no change.
- **The package version.** Teams replaces a published package only by a higher version.
  The kernel builds an agent app's version as
  `1.<TEMPLATE_REVISION>.<identity revision>`, reading `appPackage/TEMPLATE_REVISION.txt` of
  the installed plugin; content that is not a non-negative integer fails the package
  build. Raise the file with every change of the template that changes the published
  definition.
- **Existing installs.** Facilitators are new agents; their apps are provisioned from
  the new template, nothing has to be mass-upgraded. An agent app that was installed
  from the earlier package keeps the consent set it was installed with (live: installs
  at versions 1.0.1 to 1.3.0 show exactly the permissions they were installed with).
  That is fine: those agents do not use meeting events and keep working as before. To
  bring one such install to the full set, call `repairConsent` for that chat. It runs
  `POST /chats/{chat-id}/installedApps/{installation-id}/upgrade` with the complete
  `consentedPermissionSet` of the published definition and needs
  `TeamsAppInstallation.ReadWriteAndConsentForChat.All` and `AppCatalog.ReadWrite.All`
  on the instance's app, both held in byte5. The consent semantics of the upgrade call
  are **Open → M5**.

### The participant event subscription is a manual step

Docs [D19]: "A bot can receive the participant events only if subscribed to these events
in Developer Portal." The page gives portal steps only. The app manifest schema has no
property for it (`meetingExtensionDefinition` is mentioned with
`supportsAnonymousGuestUsers` only), and neither a Graph API nor a Developer Portal API
for it is documented. So it **cannot be set for apps provisioned through Graph**, and it
is unknown whether the portal setting survives a catalog update of the package
(**Open → M5**).

Consequences, as implemented: meeting start and end work without it. `presence` is
declared only after a join or leave event has actually been seen for the running
meeting, which is the only proof that the subscription is set. The manual portal step
per agent app is in [`facilitator-live-verification.md`](./facilitator-live-verification.md).

### Targeted in-meeting notification: not built

`POST /v1/meetings/{meetingId}/notification` with `type: targetedMeetingNotification`
needs more than the RSC `OnlineMeetingNotification.Send.Chat` (which the package now
requests): its only content type is a dialog that renders a **web page** of the app
(`content.value.url`, listed in `validDomains`; "if an empty string or no URL is
provided, nothing is rendered"), and recipients must be `29:` ids, Entra object ids are
not accepted. Docs [D20]. This plugin has no meeting web surface, so nothing is sent.
The same holds for the non-targeted in-meeting alert (`externalResourceUrl`). The
facilitator uses chat posts and private messages instead.

### Still not possible

A late joiner who joins the call **without** being added to the chat is visible only
through the participant event, so only in a scheduled meeting with the portal
subscription set. Hand raise has no bot event. Audio needs a .NET media bot (#584).

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

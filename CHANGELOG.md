# Changelog

## 0.38.4

Two findings of the first live card clicks (0.38.3, 2026-11-19).

- **An assignment card with one candidate is a card** (rollout defect D27). It was refused
  with "question 1: at least two options are needed", although "Ich übernehme das" next to
  one named person is a real choice. The rule now: a fixed candidate list needs **one**
  candidate when the card shows the volunteer button, and **two** otherwise. The button
  is shown when `allowVolunteer` is not `false` (it is on by default) **and** the
  assignment is the card's only question; on a questionnaire it is not rendered, so it
  does not count there. An empty list is refused in every case
  (`candidates: 'conversation_members'` is unchanged). Single choice, multiple choice
  and ranking keep their minimum of two options.
- A volunteer click is accepted exactly where the button is shown; a crafted one on a
  questionnaire is refused (before, only `allowVolunteer: false` refused it).
- **The error line of a refused submission survives the refresh.** One live click on
  "Antwort senden" without a selection came back as the plain question card. The server
  did answer that click with the card and "Bitte eine Option wählen." (and always has);
  what replaced it is not proven. The likely cause: the client asks for the person's
  view (`refresh`) after it received a card, and that answer was the plain card. The
  adapter now remembers a person's refused submission for ten minutes and answers the
  refresh with the same card, values and error line; a valid answer or "Ändern" ends
  it. Kept in memory only.
- No contract change; `compat.core` unchanged.

## 0.38.3

Posted question cards and the meeting state survive a restart. The instance restarts on every deploy; a vote
that spanned one left a card with live inputs that could neither be closed nor checked.

- **Stored**: every posted question card is written to the table `teams_question_cards`
  on the graph pool (bot, card id, conversation, message id, proposal id and version,
  the card, closing time, state, the closed view). The registry reads through to it on
  a cache miss, so `updateCard` and clicks work after a restart.
- **A click is read against the stored card only.** The buttons no longer carry the
  card; a card in a click's data is ignored. A card id this bot has no record of
  posting gets "Diese Abstimmung ist nicht mehr aktiv." and emits nothing. (0.38.0 read
  such a click against the card it carried.)
- **The table is created by the plugin** on first use (`CREATE TABLE IF NOT EXISTS`):
  this repository has no migrations, and `teams_conversation_refs` is a kernel
  migration. The DDL is exported as `QUESTION_CARDS_DDL` should the kernel take it over.
- **Pooler-safe**: single autocommit statements only; no transaction blocks, advisory
  locks, LISTEN, prepared-statement names or temp tables.
- **Expiry**: a card is kept 14 days after its closing time (or after it was closed),
  60 days when it has no closing time; expired rows are deleted at most once an hour.
- **Degrades, and says so**: without a pool, without the right to create the table, or
  when a write fails, cards are kept in memory and one log line names the consequence.
- Cards posted by 0.38.0 / 0.38.1 are not in the table: after the update they answer
  "nicht mehr aktiv" and cannot be closed by `updateCard`.
- Still in memory only: a person's own-answer view. After a restart a refresh shows the
  question card again; the engine holds the answer.
- **Meeting state stored too** (table `teams_meeting_states`, same rules): whether a
  meeting runs in a chat is what the bot saw (start, join, end), and since 0.38.2 the
  snapshot says `online` only then. With the state stored, a deploy during a meeting
  no longer turns it into a plain chat, and the first event after a restart continues
  the stored meeting. A stored "running" is believed for twelve hours after its last
  start or join; older than that (an end event that was missed) it is unknown, and
  unknown is reported as `chat`.
- `compat` is unchanged (`core >=1.51.0 <2.0`).

## 0.38.2

Rollout defect D24: the facilitator's proactive posts (handshake, reminder, notice)
showed **no AI marking at all** in Teams. Since 0.36.6 the send providers declared that
Teams labels natively, so the kernel added no disclosure sentence, and the adapter
stripped the sentence elsewhere; but the platform label was never seen on any message of
this plugin. What had been read as "the label" on turn answers is the answer card's own
"✨ AI-generated" chip.

- **Every agent post is visibly marked again.** New setup switch
  `teams_native_ai_label` ("AI marking: rely on the Teams label"), **off by default**.
  Off: the providers do not declare `supports.aiLabel`, so the kernel adds the
  instance's disclosure sentence to an agent's own posts; routine deliveries and long
  plain-text answers keep the sentence. On (only after the label was seen in a client):
  the 0.36.6 behaviour, label only.
- **The label entity now has Microsoft's shape.** It is sent with `"@id": ""`, as
  Microsoft's Teams SDK does for `addAiGenerated()`; this plugin sent it without. That
  is the one difference found; whether it is why the label did not show is not proven.
- **Unchanged:** answer cards keep their own chip and drop the sentence. The entity is
  attached to every message in both modes.
- **Debug capture:** with `TEAMS_LOG_OUTGOING_ACTIVITIES=1` every outgoing activity is
  logged as JSON with the path it took (`turn`, `proactive`, `direct`, `update`).
- **Tests** pin the proactive activity against the reply in a turn, field by field, and
  the entity against Microsoft's SDK shape; one test per send path and mode.
- **Live checklist:** six items (A1 to A6) in `docs/facilitator-live-verification.md`.
- `compat` is unchanged (`core >=1.51.0 <2.0`).

Also in this version (rollout defect D26): **a meeting chat is `chat`, not `online`,
unless a meeting is running.** The capability snapshot reported `online` for every
meeting chat at any time; an engine reads `online` as "a meeting is running". Now
`online` only between a meeting's start (or a participant join) and its end, as this
bot saw it; `chat` before, after, and when the state is unknown. A leave event no
longer marks a meeting as running. Microsoft documents no way to read the running
state on demand; it is kept in memory, so after a restart during a meeting the
snapshot says `chat` until the next join or start.

While the switch is off, a client that does render the label shows label and sentence
on the same message. Still without a visible marking of their own in both modes: the
blocking choice card, the approval card, Conductor reminders, and cards posted without
any text.

## 0.38.1

Two checks against the kernel's final question card contract (byte5ai/omadia#2105).

- **`updateCard` touches only a question card the same bot posted.** The kernel keeps no
  record of posted cards and checks identity and subscription only. The adapter looked
  the message up only to pick the card's language, and would have sent the update for
  any message id: Teams refuses another bot's message, but this bot's own text message
  would have been replaced by a card. Now the id must be a card that this bot posted in
  this conversation; anything else is `unreachable` (`not_permitted`) and no request is
  made. A corrected card must keep its `cardId`.
- **Consequence:** the cards are remembered in memory. After a restart of the Teams
  channel an update of an earlier card is refused, and the engine's result text in the
  chat is what people see.
- **Emitted answers checked against the kernel's reader.** No change was needed: every
  form's `card_answer` passes the kernel's `readCardAnswerFields` unchanged. A test now
  feeds each form through a copy of that function, and compares the copy with the
  kernel source where a kernel checkout with the contract is present.
- `compat` is unchanged (`core >=1.51.0 <2.0`).

## 0.38.0

Question cards: votes and structured input in Teams as Adaptive Cards (kernel spec B
§6.1, byte5ai/omadia#2089; design note "Design: questionnaire cards in Teams"). Additive:
the 0.36.0 poll and chat ballots work as before, and the plugin compiles against the
channel-sdk of the kernel's main branch as well as the newer one. The contract types are
mirrored in `kernel-types.ts` until the kernel publishes them.

- **Six forms on one card type** (`TargetedMessage.card`): single choice, multiple
  choice, ranked vote (one dropdown per place), consent (agree / object with a reason),
  assignment (people picker of the conversation or a given list, plus "Ich übernehme
  das"), free text (`append` for collecting ideas, `replace` for one statement). A
  questionnaire is up to five questions on one card. Adaptive Card 1.5, `Action.Execute`.
- **Declared**: `supports.card` with exactly the six kinds, five questions and `update`.
  `number` and `date` are in the contract and not built; a card with one is refused.
- **Typed answers**: one `card_answer` event per accepted submission with `proposalId`,
  `proposalVersion`, `answers[]` and `answererVerified`. The author is who Teams says
  clicked.
- **Own answer for the clicker only**: the click is answered with a card that shows the
  person their answer and "Ändern". The posted card is not touched: no names, no running
  tally for the group.
- **`updateCard`**: the engine replaces the posted card with its closed form (its result
  lines, no inputs) or with a corrected open card, on the stored message id, as the bot
  that posted it.
- **Every rule on the server**: required answers, unknown options, a place used twice,
  an objection without a reason, too long a text. A refused submission emits nothing and
  comes back as the same card with the values kept and the reason shown.
- **Forged or stale clicks**: another conversation, an earlier proposal version, a
  closed card, a passed closing time, another bot, an author named in the data. None of
  them is recorded.
- **AI label** on the posted card and on every update (one marking per message).
- **Live checklist**: thirteen items for everything Microsoft's pages do not state
  (`docs/facilitator-live-verification.md`, "Question cards").
- `compat` is unchanged (`core >=1.51.0 <2.0`).

Not verified: nothing was clicked in a Teams client. Not built: a card in a direct
message, the step-by-step questionnaire, an own-answer view that survives a restart.

## 0.37.0

Roster status as the kernel spec defines it (spec B §5.5, byte5ai/omadia#2091, issue
#43). Independent of the kernel: every new field is additive, and the plugin compiles
against the channel-sdk of the kernel's main branch (0.3.1) as well as the newer one.

- **Roster status.** `getRoster` answers with `status` (`complete`, `partial`,
  `unknown`), `readAt` (when the platform answered) and, when not complete,
  `statusDetail`. `partial` stays and is true unless the status is `complete`.
- **An unreadable roster is `unknown`, never an empty list.** A failed read used to
  come back as "no members" (and the pages already read were discarded). Now nothing
  read is `unknown`, a read that failed part way is `partial` with what was read, and a
  read that comes back empty is `unknown` too.
- **`fresh: true`** bypasses the cache and reads now. A failing fresh read answers
  `unknown`; it does not fall back to the cached list.
- **Cache**: a `complete` list is kept five minutes as before; a `partial` or `unknown`
  read only thirty seconds. A cached answer keeps the time of its read.
- **Address fields**: `email` and `userPrincipalName` of a member are passed on whenever
  the member list sends them (they already were; now covered by a test and documented).
  Nothing is derived when the platform sends none.
- **Compiles against the older SDK again.** 0.36.6 did not typecheck against the
  kernel's main branch (`supports.aiLabel` was unknown there). The capability
  declarations no longer depend on the installed SDK's type, and
  `scripts/typecheck-against-sdk.mjs` checks the plugin against any SDK checkout.
- **Live run recorded**: `docs/facilitator-live-verification.md` holds the results of
  the run of 2026-11-18 per owner step and scenario, with what was not run.
- `compat` is unchanged (`core >=1.51.0 <2.0`).

Terminology (`docs/terminology.md` of the kernel), user-visible texts only:

| Where | Before | Now |
|---|---|---|
| manifest description | "First-Class-Channel für Omadia" | "Kanal für omadia" |
| setup guide, en | "Multiple bots"; "several Teams bots — one … per omadia agent" | "Several channel identities"; each bot is the **channel identity** of one agent |
| setup guide, de | "Mehrere Bots" | "Mehrere Kanalidentitäten"; jeder Bot ist die **Kanalidentität** eines Agenten |
| field label `teams_bots` | "Teams Bots (JSON)" | "Teams-Kanalidentitäten (Bots, JSON)" (en guide: "Teams channel identities (bots, JSON)") |
| field help `teams_bots` | "Bot-Identitäten" | "Kanalidentitäten (ein Bot je Agent)" |
| Tool-Trace help | "sub-agents" / "Sub-Agenten", "Tools" | "internal helpers" / "interne Helfer", "Werkzeuge" |
| Direct-line help | "consulted specialist" / "konsultiertem Spezialisten" | "consulted agent" / "befragtem Agenten (Rückfrage)" |
| de texts | "User", "operator-freundlich", "Operator-UI", "Endanwender" | "Nutzer", "für Administratoren lesbar", "Administrationsoberfläche", "Nutzer" |
| de texts | "dieser Channel", "der Channel", "Teams-Channel" | "dieser Kanal", "der Kanal", "Teams-Kanal" |
| playbook, de | "delegiert … an den zentralen Orchestrator" | "… an die Agentenlaufzeit des Kerns" |
| README | "delegated to the shared Conductor/orchestrator"; "the orchestrator's `chatAgent` capability" | "handed to the kernel's agent runtime"; "the agent runtime's `chatAgent` service" |

Not renamed: the config key `teams_bots` and the route `/api/teams/<botSlug>/messages`
(identifiers of existing installs), product names ("Azure Bot", "Bot Framework"), and
the button labels on answer cards ("Memory-Auswirkung", "Konsultiert", "Tool-Trace"),
which are card texts in code and belong to a separate change.

## 0.36.6

One AI marking per message (kernel byte5ai/omadia#2098, `@omadia/channel-sdk` 0.3.2). In
Teams that is always the "AI generated" label and never the disclosure sentence.
**Install this before the kernel with #2098 is deployed**: without the declaration below
that kernel adds the sentence to the first own-identity post, next to the label.

- **Declared**: both send providers state `supports.aiLabel: true`, so the kernel adds no
  disclosure sentence to posts and private reminders sent as an agent's own identity.
- **Routine deliveries and approvals** (`publishProactiveSend`): a routine delivery was
  posted with the folded sentence and without the label. The sentence is now stripped and
  the label attached, for the routine card and the plain-text fallback; the approval card
  and Conductor reminders get the label. Moved into `teamsRoutineDelivery.ts`.
- **`NO_REPLY` with a folded sentence** is still dropped, for a routine delivery and for a
  turn answer. With the sentence behind it the sentinel was no longer the last line and
  the delivery would have been posted.
- **Log**: the lines for edit / delete / reaction, `recorded for the engine`, `drop
  non-mention` and meeting events print the full conversation id (they cut it at 24
  characters, which hid the thread part of a channel id).
- `compat` is unchanged (`core >=1.51.0 <2.0`).

| Send path | What arrives from the kernel | What Teams shows | Test (`tests/teamsAiMarking.test.ts` unless named) |
|---|---|---|---|
| Turn answer (card) | text with the sentence folded in, `aiDisclosure` | label; sentence stripped | "a model answer with the folded sentence" |
| Fixed refusal / "not set up" answer of a facilitator | same shape as a model answer | label; sentence stripped | "a fixed refusal of the facilitator" |
| Turn answer, disclosure turned off | text, no `aiDisclosure` | label | "an answer without a disclosure" |
| Long turn answer (several plain messages) | as a turn answer | label on the first part; sentence stripped | "a long answer that is posted as plain text" |
| Blocking choice card | question, options | label | "a blocking choice card" |
| `NO_REPLY` answer, also with the sentence behind it | sentinel | nothing is posted | "a NO_REPLY answer stays silent" |
| Proactive text (`sendToConversation`) | text without the sentence, because `supports.aiLabel` is declared | label | "both providers declare the label" |
| Poll card (`sendToConversation`) | poll | label | same test |
| Private reminder (`sendToUser`) | text without the sentence | label | same test; `teamsOutboundAsOwnBot.test.ts` |
| Routine delivery, card | text with the sentence folded in, `aiDisclosure` | label; sentence stripped from the card body | "a routine delivery as a card" |
| Routine delivery, plain text | same | label; sentence stripped | "a routine delivery as plain text" |
| Conductor reminder (`{ text }`) | text, no `aiDisclosure` | label | "a delivery without aiDisclosure" |
| Approval card | deterministic text | label | "an approval card" |
| `NO_REPLY` routine answer | sentinel | nothing is posted | "a NO_REPLY routine answer is still dropped" |
| Streaming | does not exist: the adapter calls `chat()` and posts the final answer | n/a | n/a |

Not marked: the adapter's own fixed texts and side cards in a turn (card-click
acknowledgements, error texts, recalled-context, slot, consent and routine-list cards).

## 0.36.5

From the first live facilitation in a Teams meeting chat (2026-11-18).

- **Roster as the asking agent's own bot.** The handshake's participant list was empty:
  the roster was looked up without a bot, i.e. as the instance's default bot, which is
  not in a facilitator's chat. `getRoster(conversationId, { asChannelKey })`
  (`@omadia/channel-sdk` 0.3.2) now resolves the conversation reference of exactly that
  bot and reads the members with that bot's credentials; `self` is that bot. A bot
  without a reference in the conversation gets `undefined`, never another bot's view.
  Without `asChannelKey` the behaviour is unchanged.
- **The roster cache is per bot and conversation.** It was per conversation, so one
  bot's refused (empty) read could be served to another bot in the same chat for five
  minutes.
- **Channel threads**: the members are asked for the channel (the bare id), also when a
  post into a thread resolves mentions.
- **One log line on every roster path**: `[teams-roster] roster conv=… as=… fetched
  count=N`, `no reference: …`, or `refused status=… detail=…`.
- **AI label on everything posted as an agent identity.** Posts through
  `sendToConversation` and `sendToUser` (engine posts, polls, private reminders) carried
  no "AI generated" label; only turn answers did. Every such message activity now
  carries the label entity once. No disclosure sentence is added to the text.
- **The turn origin names the speaker** (`TurnOrigin.speaker`, channel-sdk 0.3.2):
  `{ kind: 'teams-aad', id: <Entra object id>, displayName? }` from `activity.from`.
  Omitted without an Entra object id. The kernel builds `callerUserRef` from it, so a
  facilitator can name who started a facilitation. `channelIdentity` is still not set.
- `compat` is unchanged (`core >=1.51.0 <2.0`).

Not verified live: the member read on the bare channel id for a thread.

## 0.36.4

A facilitation could not be started from Teams: the engine asked the readiness check
about `28:<appId>-19:meeting_…@thread.v2`, an id the kernel had cut out of this channel's
session scope (`teams-<botKey>-<conversation id>` for a bot with its own identity). Graph
answered 400 and no conversation reference exists under such an id.

- **The turn origin states the channel-native conversation id** (`channelConversationId`):
  `activity.conversation.id`, unchanged. It is the id this channel uses for conversation
  references, the subscription lookup, message events and posts.
- **A facilitation in a team channel belongs to the thread it was started in.** A channel
  message's conversation id is `19:…@thread.tacv2;messageid=<root>`; that per-thread id is
  used everywhere towards the kernel. Replies in the thread arrive under it, a post for it
  goes into the thread, and a new root post in the channel is another conversation. A
  thread the bot holds no reference of its own for is reached through the channel's
  reference.
- **Graph never sees a thread id.** Readiness and consent repair of a channel address the
  team (`channelData.team.aadGroupId`, or the id stored with the reference; now kept per
  channel, not per thread). A chat is addressed by its plain id.
- **Consent repair works for a team install** as well (`/teams/{id}/installedApps/{id}/upgrade`,
  needs `TeamsAppInstallation.ReadWriteAndConsentForTeam.All`). It answered
  `unsupported_conversation` before.
- **A bot-key-prefixed conversation id is refused by name**: readiness, consent repair and
  the conversation send answer `unsupported_conversation` / `no_binding` with a message
  that names the id and the cause, without asking Graph, and log it once.
- **Tests on real shapes**: the simulated activities carry the ids and `channelData` keys
  captured live on 2026-11-18, and the scenario suite runs a facilitation in a channel
  thread and in a meeting chat with bot-qualified session scopes.
- `compat` is unchanged (`core >=1.51.0 <2.0`).

Not verified live: a human reply inside a channel thread (handled as documented: same
conversation id as the root post), the post into a thread, and the team consent upgrade.

## 0.36.3

A facilitator refused every mention in a team channel or a group chat ("I am a moderator,
not an assistant … add me to the group chat"), before the model ran. The turn origin this
channel sends carried no `conversationType`, and the kernel treats a turn as a group turn
only on a positive `conversationType: 'group'`. Other agents were not affected: only the
facilitator wrapper reads the field.

- **Turn origin states group or direct** (`teamsTurnOrigin`): `channel` and `groupChat`
  (a meeting chat is a `groupChat`) are `group`, `personal` is `direct`. Without a type,
  `conversation.isGroup === true` counts as `group`; anything else stays unstated. Both
  paths that start a turn (inline and the detached card-click continuation) build the
  origin there.
- **Tests**: the origin per conversation kind, checked against a copy of the kernel's
  `isGroupTurn` (compared with the kernel source where a checkout is present), and a
  scenario in which a mention in a group chat, a channel and a meeting chat reaches the
  turn with a group origin.
- `compat` is unchanged (`core >=1.51.0 <2.0`).

## 0.36.2

An agent app built from the template could not be installed into a team: Graph's team
consent check refuses `FileStorageContainer.Selected.Group` (`400`, found by a live
bisect on 2026-11-18). The permission had been in the template since 2026-08-25 and
nothing used it. Chat installs were not affected.

- **Removed** `FileStorageContainer.Selected.Group` from
  `appPackage/manifest.json.template`. The four meeting permissions stay; the template
  requests ten RSC permissions. `appPackage/TEMPLATE_REVISION.txt` is 3.
- **Consent repair follows the published definition**: `repairConsent` now also upgrades
  an install whose version differs from the published one, and sends exactly the set of
  the current definition. It answers `package_outdated` while the catalog still holds an
  app from an earlier template revision. Before, an install from the revision 2 package
  with all chat permissions granted was `not_needed` and would have stayed
  `package_outdated` for good.
- **Existing installs** from 0.36.0 / 0.36.1 (app version `1.2.<n>`) are reported as
  `package_outdated` until the app is published again and the consent repaired.
- **Test**: the template's RSC list is pinned against the permissions that were installed
  live into a team and a chat. A new permission has to be installed once into a test team
  and a test chat first (`docs/facilitator-live-verification.md`, "RSC permissions:
  install proof").
- `compat` is unchanged (`core >=1.51.0 <2.0`).

## 0.36.1

0.36.0 could not be installed and was withdrawn: the kernel's plugin zip ingest refuses
extension-less files, and the package contained `appPackage/TEMPLATE_REVISION`
(`422 zip.forbidden_extension`, on every kernel). 0.36.1 is 0.36.0 with that fixed.

- **Renamed** `appPackage/TEMPLATE_REVISION` to `appPackage/TEMPLATE_REVISION.txt`,
  same content (`2` and a newline). The plugin reads the new name; the old file is no
  longer shipped.
- **Packaging test**: every entry of the zip that `npm run package` builds is checked
  against the kernel ingest's rules (extension allowlist, the `appPackage/` and `ui/`
  exceptions, extension-less names, symlinks, escaping paths).
- `compat` is unchanged (`core >=1.51.0 <2.0`).

## 0.36.0

Native facilitators (byte5ai/omadia#2043; #43 and the baseline of #44). Needs the
kernel with `@omadia/channel-sdk` 0.3.0 for the new behaviour; every kernel member is
feature-detected, on an older kernel nothing changes.

- **Every message in a moderated chat**: a bot that holds a message subscription for
  a conversation (`isAllMessagesSubscribed`) emits each human message as a message
  event with the author Teams states, instead of dropping it at the mention gate. The
  adapter starts **no chat turn** in such a conversation, not for a mention either:
  the engine drives every turn. An empty answer is never posted and never replaced by
  the "no answer came back" text, also when only the AI disclosure sentence or
  follow-up buttons are left of it (the kernel empties the answer of the mention that
  opens a facilitation). The subscription is per bot and conversation: other bots in
  the same chat keep the gate, chats nobody moderates behave as before.
- **Edits, deletions, reactions, card answers**: emitted as `edited` / `deleted` /
  `reaction` / `card_answer`, keyed on the message id, the revision counted per edit.
  Each such activity is logged with its `eventType` and ids.
- **Readiness check** (`teamsReadiness.ts`): runs with the instance's Graph client
  (the Microsoft 365 integration's registration), because the per-agent bot
  registrations hold no Graph permissions. Ready = installation row for the bot's app
  and a grant with `clientAppId` = bot app id and `ChatMessage.Read.Chat` (team:
  `ChannelMessage.Read.Group`); can-post = own conversation reference plus the last
  send result. `repairConsent` upgrades an install that lacks the grant
  (`installedApps/{id}/upgrade` with the full consented set).
- **Restart**: a subscribed conversation needs no new `bot_added` after a middleware
  restart. The subscription is asked from the kernel per message, the roster is
  fetched per call, the conversation reference comes from `teams_conversation_refs`,
  and the team of a channel (needed by the readiness check) is now stored with that
  reference. Without a database pool the references are memory only, as before.
- **Posts as the own bot**: real @mentions resolved from the roster by Entra object
  id; text and card are separate activities, the outcome carries `messageId` and
  `cardMessageId`.
- **Declared send capabilities**: the conversation send provider declares
  `supports: { asChannelKey, mentions, poll }` and the targeted send provider
  `supports: { asChannelKey }`. The kernel only uses a send option a provider declares;
  without it every own-identity post is answered with `no_identity`. Message and
  session events carry the plugin's own catalog id and channel type.
- **Direct messages as the own bot**: `sendToUser(…, { asChannelKey })` uses that
  bot's own 1:1 or creates it by Entra object id (first contact, also after a
  restart). `403 ForbiddenOperationException` → `app_not_installed`, `403` with
  `MessageWritesBlocked` → `blocked`, neither retried; 429 / 412 / 502 / 504 retried
  with backoff. Never another bot, no default-bot fallback. Without `asChannelKey`
  unchanged.
- **Meetings (#44)**: meeting chats follow the same subscription rule; poll card with
  `Action.Execute` (`teamsPollCard.ts`), answers arrive as `card_answer` with
  `answererVerified` false for a participant without an Entra identity. Meeting start,
  end and participant join / leave are passed to the kernel as session events
  (`teamsMeetingSessions.ts`, `core.emitSessionEvent`, feature-detected) and a fresh
  capability snapshot is pushed at start and end. The snapshot
  (`teamsMeetingSource.ts`) declares `presence` only while it is proven: a scheduled
  meeting is running, a join or leave event arrived for it, and the install holds both
  meeting permissions (read from Graph). Meet-now and calls keep chat and cards.
  `reactions`, hand raise and audio are never declared. Targeted in-meeting
  notification is not built (it needs a web page of the app).
- **App package**: the template requests `OnlineMeeting.ReadBasic.Chat`,
  `OnlineMeetingParticipant.Read.Chat`, `OnlineMeetingNotification.Send.Chat` and
  `ChannelMeeting.ReadBasic.Group` (eleven RSC permissions of sixteen allowed).
  `appPackage/TEMPLATE_REVISION` is 2; the kernel builds the agent app version as
  `1.<TEMPLATE_REVISION>.<identity revision>`. An install from an earlier package is
  reported as `package_outdated` by the readiness check. Agent apps installed
  from the earlier package keep their consent set until `repairConsent` upgrades the
  install; that is fine for agents that do not use meeting events. Participant events
  additionally need a manual per-app setting in Developer Portal.
- **Tab configuration page**: adding an agent app to a channel from the Teams app
  dialog opens `/p/channel-teams/tab-config`. For an agent without plugin surfaces the
  picker was empty, so Save never enabled and the install looked failed although it had
  succeeded. The picker now always offers the plugin overview, Save is enabled, and the
  page says that a tab is optional and the agent is already added. The app package
  template is unchanged (`configurableTabs` stays; the tab catalogue needs it).
- **Docs**: `docs/teams-meeting-apis.md` (findings per topic, each labelled live /
  docs / simulated / open, with sources) and `docs/facilitator-live-verification.md`
  (owner steps M1 to M6 and the scenario walk-through).
- **Verification status.** There is no Teams test tenant. The tests run against a
  simulated Bot Framework, Bot Connector and Graph. Confirmed live in the byte5 tenant
  on 2026-11-18, with the builds running that day: delivery without mention in group
  chats, in a team channel and in a scheduled meeting chat; the Graph readiness reads;
  the DM error bodies and that team membership counts as installed; the tenant RSC
  setting. Still open, because they need this build on an instance: edit / delete /
  reaction events under RSC, the returned message id, the meeting events, the portal
  subscription and the consent upgrade (M2, M3, rest of M4 and M5).
- Limits: the edit revision the adapter sends is a hint (the kernel assigns the stored
  one); a team channel can be
  checked only after the bot received an activity from it.

## 0.35.0

Der Plugin-Katalog der Teams-Tabs läuft über den Kernel-Modus
`auth: 'teams-sso'` (byte5ai/omadia#1620) statt über die eigene Token-Prüfung.

- **Router getrennt**: `/p/channel-teams/hub` und `/tab-config` sind
  datenfreie Hüllen mit `auth: 'public'` (`createTeamsUiShellRouter` — die
  Factory bekommt den Katalog gar nicht erst). Der Katalog
  `/p/channel-teams/catalog` ist ein eigener Router mit `auth: 'teams-sso'`
  (`createTeamsCatalogRouter`); der Kernel leitet an das längste lebende
  Präfix weiter, die URL der Seiten bleibt unverändert.
- **Prüfung im Kernel**: RS256 gegen Entra-JWKS, Tenant + App-ID eines Bots
  dieser Installation, `api://`-Audience-Host des Deployments, Teams-/M365-Client,
  delegiertes Token eines Tenant-**Members** (Gäste und App-only abgelehnt),
  Laufzeit. Neu gegenüber 0.31.0: B2B-Gäste sehen den Katalog nicht mehr.
- **Fail-closed**: Erste Middleware des Katalog-Routers antwortet
  `401 {"code":"auth.teams_sso_required"}`, wenn `req.teamsSsoUser` fehlt —
  ein Bearer-Header allein öffnet nichts.
- **Entfernt**: `src/teamsTabAuth.ts` (eigener Verifier) samt Exporten
  (`createTeamsSsoVerifier`, `bearerToken`, `TeamsSsoError`, …) und
  `login.microsoftonline.com` unter `network.outbound` (nur für den
  Schlüssel-Abruf des Verifiers deklariert; der Graph-Client nutzt es nicht
  über `ctx.http`). `createTeamsUiRouter` ist durch die beiden neuen
  Factories ersetzt.
- **Manifest**: `compat.core` `>=1.51.0 <2.0` (plugin-api mit `'teams-sso'`).
  `permissions.public_paths` bleibt `"/p/channel-teams"` — der Katalog liegt
  darunter, die bestehende Freigabe gilt weiter.
- Die `href`s im Katalog kommen weiter aus den `uiRoutes`-Deskriptoren der
  Plugins (`/p/<slug><path>`); Plugins nach #1620 zeigen dort auf ihre
  Teams-Hülle (`/teams/<page>`).
- Tests: `teamsUiRouter.test.ts` (Hüllen ohne Daten, Katalog 401 ohne
  `req.teamsSsoUser` und mit bloßem Bearer, 200 mit Kernel-User, Manifest).
  `teamsTabAuth.test.ts` entfällt mit dem Verifier.

## 0.34.0

- **People directory** (`teamsPeopleDirectory.ts`): registers the kernel's
  `PeopleDirectoryProvider` for `teams` (feature-detected —
  `core.registerPeopleDirectoryProvider?.`), so Conductor human steps can
  resolve a person named in prose to a routable address. Graph
  `GET /users?$search` under the default bot's app registration,
  filtered to enabled `Member` accounts of the bot's tenant; `principalRef` =
  mail lowercased = the channel binding key (mail-less users are skipped — UPN
  never binds).
  Requires the **`User.Read.All` application permission** (admin consent);
  401/403 → `not_configured` (the kernel surfaces a *partial* result, nothing
  breaks). `TeamsGraphResolver.searchUsers()` + `GraphRequestError` (status
  on Graph failures) added.

## 0.33.0

Buttons auf der Antwort-Card per Plugin-Einstellung ein- und ausschaltbar.

- **Sieben neue Boolean-Setup-Felder** (`teams_card_*`, Standard jeweils an):
  `teams_card_tool_trace` (🔧 Tool-Trace), `teams_card_memory_disclosure`
  (🧠 Memory-Auswirkung), `teams_card_privacy_receipt` (🛡 Privacy Guard),
  `teams_card_fresh_check` (🔄 Fresh Check), `teams_card_follow_ups`
  (💡 Nächste Varianten), `teams_card_direct_line` (💬 Direkt mit <Agent>),
  `teams_card_verbosity_steps` (🔽 Kürzer / 🔼 Mehr Details). Ohne gesetzte
  Werte rendert die Card unverändert zu 0.32.0.
- **Was ein Schalter tut**: Button *und* zugehöriger aufklappbarer Abschnitt
  (bzw. Überschrift + ActionSet bei den Folgevorschlägen) entfallen gemeinsam —
  nie ein Button ohne Ziel, nie ein verwaister Container. Dieselben Schalter
  gelten für die Mini-Cards des Long-Answer-Fallbacks (> 25 KB).
  `buildFollowUpsOnlyCard` liefert dafür jetzt `undefined`, wenn nichts zu
  zeigen ist.
- **Bewusst nicht schaltbar**: Rückfrage-, Topic-, Freigabe-, Slot-Picker-,
  Consent-, Routine- und Agenten-Apps-Cards (ihre Buttons beenden einen
  Ablauf) sowie alle Status-Chips (✨ KI-Kennzeichnung nach AI Act Art. 50,
  Prüf-Badge, 🗄 Originaldaten, 🔎 Konsultiert, 💬 Direkte Antwort). Der
  Direct-Line-Schalter entfernt nur die Buttons, der Konsultiert-Chip bleibt.
- **Nur Ausblenden**: Die Daten hinter einem Button (Folgevorschläge,
  Privacy-Receipt, Trace, Memory-Capture) erzeugt der Kernel weiterhin — ein
  Schalter spart keine Tokens.
- **Mechanik**: `src/teamsCardFeatures.ts` (Typ, Defaults, Key-Map, toleranter
  Boolean-Parser für `true/false` und `"true"/"false"`), neues Feld
  `features` auf `BuildAnswerCardInput`, Auflösung einmal in `activate()`
  (der Kernel re-aktiviert das Plugin beim Speichern der Config), Log-Zeile
  mit den deaktivierten Schaltern. **Gilt für neue Antworten**: Click-Handler
  bleiben ungeändert, bereits gepostete Cards behalten ihre Buttons und
  funktionieren weiter. Mögliche Fortsetzung: ein Kernel-Hinweis, der bei
  ausgeschalteten Folgevorschlägen das `suggest_follow_ups`-Tool gar nicht
  erst laufen lässt (spart dann auch Tokens).
- Tests: `teamsCardFeatures.test.ts` (Parser, Resolver, pro Schalter an/aus,
  Byte-Identität ohne Config, alle aus, Fallback-Cards).

## 0.32.0

Kein Fallback mehr auf den Default-Orchestrator ohne Plugin-Zuordnung
(byte5ai/omadia#1598).

- Wenn der Kernel-Resolver für einen Chat **keinen** Agenten liefert (er lehnt
  ab oder wirft), antwortete der Bot bisher über den `chatAgent@1`-Default. Der
  hat keine Plugin-Zuordnung pro Agent und damit die Tools **aller** Plugins.
  Jetzt wird der Turn verweigert, und der Bot antwortet: „Diesem Chat ist kein
  Agent zugeordnet. Bitte wende dich an deinen Administrator."
- Das gilt auch für Karten-Klicks und abgesetzte Turns, die den Agenten später
  erneut auflösen: Liefert der Resolver dort keinen Agenten, wird der Turn
  ebenfalls verweigert.
- Installationen ohne Resolver (Single-Agent, kein Postgres) verhalten sich wie
  bisher.

## 0.31.0

Teams-Tabs laden wieder (Folge von byte5ai/omadia#1596 / #1608): Seit der Kernel
`/p/…`-Routen mit `auth: 'session'` hinter die omadia-Session stellt, antworteten
`/p/channel-teams/hub` und `/tab-config` in Teams mit 401 — Teams lädt sie ohne
unser Session-Cookie.

- **`auth: 'custom'` + `permissions.public_paths: ["/p/channel-teams"]`**: Der
  Router läuft nicht mehr durch das Session-Gate, sondern über den
  Public-Path-Mount des Kernels — erst, wenn der Operator den Pfad freigibt.
- **Eigene Prüfung im Plugin**: `/hub` und `/tab-config` sind jetzt datenfreie
  Hüllen. Den Plugin-Katalog liefert nur der neue JSON-Endpunkt `/catalog`, und
  der antwortet 401 ohne gültiges **Teams-SSO-Token**
  (`authentication.getAuthToken()`, Ressource = `webApplicationInfo` des
  App-Pakets). Geprüft werden RS256-Signatur gegen die Entra-Schlüssel, Issuer
  und `tid` (nur Tenants der konfigurierten Bots), Audience (App-ID eines Bots,
  `api://<host>/<appId>` oder blank), ausgestellt an einen Teams-/M365-Client
  (`azp`/`appid`), Laufzeit sowie ein delegiertes User-Token (`scp` + `oid`;
  App-only-Tokens zählen nicht). Der Tab-Kontext aus
  `app.getContext()` dient nicht zur Authentifizierung.
- **Ohne Abhängigkeit**: `node:crypto` (JWK-Import) + `ctx.http.fetch` (Egress
  über die Allow-List des Kernels). Schlüssel werden pro Tenant eine Stunde
  gecacht; ein unbekanntes `kid` löst höchstens alle 5 Minuten einen neuen
  Abruf aus, nach einem Fehlschlag gilt 60 s Backoff, parallele Anfragen teilen
  sich einen Abruf, und abgerufen wird nur für konfigurierte Tenants.
- **Wer den Katalog sieht**: jede angemeldete Person eines konfigurierten
  Tenants, auch B2B-Gäste. Der Katalog enthält nur Plugin-IDs, Titel und Pfade;
  die verlinkten Seiten prüfen selbst.
- **Manifest**: `login.microsoftonline.com` unter `network.outbound`
  (Schlüssel-Abruf).
- **Voraussetzung**: Die App-Registrierung muss eine API unter
  `api://<middleware-host>/<appId>` mit Scope `access_as_user` freigeben und die
  Teams-Clients vorab autorisieren (Setup-Guide, Abschnitt 5). Ohne das zeigt
  der Hub „Teams sign-in failed“ statt des Katalogs — nie Daten.
- Tests: `teamsTabAuth.test.ts` (v1/v2-Token, fremder Schlüssel, manipulierter
  Payload, `alg: none`, fremder Tenant ohne Schlüssel-Abruf, fremde Audience,
  falscher Issuer, abgelaufen, App-only, fremder Client, `kid`-Refetch-Floor,
  Fehler-Backoff + Single-Flight; Router: Hüllen
  ohne Katalogdaten, `/catalog` 401 ohne/mit gefälschtem Token, 200 mit
  gültigem).

## 0.30.0

Antworten, die Privacy Shield server-seitig gerendert hat, tragen jetzt auch in
Teams den Hinweis auf Originaldaten (Kernel-Seite: byte5ai/omadia #1402/#1407):

- **„🗄 Originaldaten aus dem System“** als Chip in der Statuszeile der
  Antwort-Card, neben „✨ AI-generated“ und dem Verifier-Chip. Gleicher Wortlaut
  wie das `SystemDataBadge` im Web-Chat. Die Werte setzt der Server unverändert
  aus dem Quellsystem ein; der Verifier-Chip bewertet weiter nur den Text des
  Modells drumherum.
- **Quelle**: `SemanticAnswer.answerSource === 'privacy-render'`, das der
  Kernel schon mitliefert, Teams bisher aber ignoriert hat.
- Beim Long-Answer-Fallback (> 25 KB, Text-Chunks statt Card) gibt es wie beim
  Verifier-Chip keine Statuszeile.
- Tests: `teamsCard.test.ts` (Chip mit/ohne Verifier, fehlt bei Modell-Antwort,
  übersteht die Größenstufen).

## 0.29.0

Antwortumfang pro Turn: zwei Buttons auf der Antwort-Card (Phase 3 des
Antwortumfangs; Kernel-Seite in byte5ai/omadia, nach #1369 und #1370):

- **„🔽 Kürzer“ / „🔼 Mehr Details“** stellen dieselbe Frage noch einmal, eine
  Stufe knapper bzw. ausführlicher auf der Skala `tldr · brief · standard ·
  detailed · max` — nur für diesen einen Turn; der nächste Turn läuft wieder
  mit der Agent-/Installationseinstellung. Am Rand der Skala fehlt der
  jeweilige Button.
- **Woher die Stufe kommt**: der Kernel meldet mit jeder Antwort
  `answerVerbosity: { effective, source }`. Ohne diese Angabe (älterer Kernel)
  gibt es keine Buttons; das Plugin liest das Feld strukturell
  (`readAnswerVerbosity`), weil das SDK, gegen das es kompiliert, es noch
  nicht kennt.
- **Mechanik wie Fresh Check**: `Action.Submit` mit `{ type: 'verbosity_step',
  level, originalMessage }`, Klick läuft losgelöst (`runOrchestratorTurnDetached`)
  mit erhaltener Historie und `answerVerbosity: level` im Turn-Input.
- Budget: über Teams' Card-Grenze fliegen zuerst die Größen-Buttons, dann die
  Direct-Line-Buttons. Beim Long-Answer-Fallback (> 25 KB, Card wird durch
  Text-Chunks ersetzt) reiten die Buttons auf der schlanken Direct-Line-Card
  mit — gerade eine `max`-Antwort braucht den Weg zurück zu „Kürzer“.
- Tests: `teamsCard.test.ts` (Buttons, Skalen-Rand, Payload-Parsing, kein
  Rendern ohne Stufe/Frage).

## 0.28.0

Teams liefert die Anwesenden für das mitgliederbezogene Gedächtnis
(byte5ai/omadia#1340, Kontextmodus `members`):

- **Der Roster gilt als vollständig**: `teamsBot.ts` reicht den
  Teilnehmer-Provider mit `completeRoster: true` an den Kernel. Erst damit
  löst der Kernel einen Gruppenchat zu einer Owner-Menge auf; ohne das Flag
  bleibt eine Gruppe „unbekannt“ und sieht kein mitgliederbezogenes Wissen.
- **Kein veralteter Roster als Publikum**: `TeamsRosterProvider.list` nimmt
  ein optionales `maxAgeMs`; für die Audience-Abfrage gilt 10 s statt des
  5-Minuten-Caches. Ein Neuzugang, dessen Membership-Event verloren ging,
  wird so spätestens nach 10 s mitgezählt — und sperrt das Wissen, das nicht
  auch seins ist.
- **Bots sind keine Anwesenden**: Roster-Einträge mit `28:`-ID tragen
  `kind: 'agent'` und zählen nicht zur Owner-Menge.
- **TurnOrigin**: jeder Turn nennt seinen Chat-Kontext (1:1 → `personal`,
  Gruppenchat → Konversation, Team-Kanal → Konversation + Team).
- Tests: `teamsRosterAudience.test.ts` (Cache vs. `maxAgeMs`, Bot-Markierung),
  `teamsTurnOrigin.test.ts`. `botbuilder` ist im Test-Bundle jetzt extern,
  damit ein Stub von `TeamsInfo` dieselbe Instanz trifft wie `dist/`.

## 0.26.2

Der Recheck-Klick nimmt sein Installationsziel nicht mehr aus der Karte
(byte5ai/omadia#1030):

- **`handleTeamsAgentAppsRecheck` leitet Team + Tenant nur noch aus der
  Submit-Activity ab** (`teamsBot.ts`). Der Zweig ist out-of-band — er
  antwortet vor dem Orchestrator, es hat also niemand vorher einen Principal
  etabliert — und `data` einer Adaptive Card ist client-editierbar. Der alte
  Fallback auf die durchgeschleiften `teamId`/`tenantId` machte damit jeden
  handgebauten oder wiedergespielten „🔄 Prüfen"-Klick zu einem frei wählbaren
  Ziel für `installAgentApps`, über Tenant-Grenzen hinweg. Ein Klick, dessen
  Transport kein Team-Scope trägt, wird jetzt abgelehnt (`'refused-unscoped'`
  plus eine erklärende Nachricht statt stiller Wirkungslosigkeit) — dasselbe
  Muster, mit dem `handleApprovalDecision` verweigert, wenn es den
  Antwortenden nicht bestätigen kann.
- **Das Scope steht gar nicht mehr auf der Leitung**: `AgentAppsRecheckValue`
  ist auf den Discriminator zusammengeschrumpft, `parseAgentAppsRecheckValue`
  verwirft jeden weiteren Schlüssel, und `buildAgentAppsResultCard` schleift
  nichts mehr durch (`BuildAgentAppsCardInput` braucht `teamId`/`tenantId`
  nicht länger — sie dienten ausschließlich diesem Round-Trip). Bereits in
  Kanälen liegende Karten tragen die beiden Felder weiter; der Parser ignoriert
  sie, alte Karten funktionieren also unverändert und die Werte werden nur
  nirgends mehr gelesen. `satisfies AgentAppsRecheckValue` an der
  Submit-`data` macht ein Wiedereinführen zum Compilerfehler.

## 0.26.1

Routine-Smart-Card-Klicks tragen ihren Principal (byte5ai/omadia#1029):

- **`actor` am `handleRoutineAction`-Aufruf** (`teamsBot.ts`, `plugin.ts`):
  Der Karten-Pfad wird out-of-band abgearbeitet und erreicht
  `runOrchestratorTurn` nie — also installiert nichts die Routinen-ALS, aus der
  der Kernel sonst `(tenant, userId)` liest, und er handelte ungescoped. Da die
  Karte die Routine-ID trägt, erreichte ein wiedergespielter Payload
  `pause`/`resume`/`trigger_now`/`delete` für jede Zeile; `trigger_now`
  liefert in die `conversationRef` der Routine, konnte also eine Nachricht in
  die Konversation eines fremden Tenants schieben.
  Übergeben wird jetzt derselbe Wert, den `captureRoutineTurn` auf dem
  Orchestrator-Pfad bekommt: Tenant aus `GRAPH_TENANT_ID`, User aus der
  gemeinsamen `userId`-Ableitung in `handleMessage` (`from.aadObjectId`, sonst
  `from.id`). Fehlt eine der beiden Hälften, entfällt `actor` komplett — ein
  halb gefüllter Principal wäre schlechter als der dokumentierte Fallback.
- Optional im Contract (`plugin-api >= 1.11.0`): ein älterer Kernel ignoriert
  den Zusatzschlüssel, ein neuerer bevorzugt ihn gegenüber der ALS. Damit
  bleibt dieses Release auf beiden Kernel-Ständen lauffähig.

## 0.21.0

Auto-Invite für Agenten-Apps (epic byte5ai/omadia#860, Wave W2 — Refs #20 #21
#22):

- **`teams_agent_apps[]` config surface** (`teamsAgentApps.ts`): JSON-Array
  `{agentSlug, teamsAppExternalId, teamsAppId?, displayName?}` — als echtes
  Array (Install-Registry) oder JSON-String (Setup-Wizard). Nur öffentliche
  App-IDs, Inline-Credentials werden hart abgelehnt (teams_bots-Präzedenz).
  Leeres/fehlendes Feld → Feature komplett aus, null Verhaltensänderung für
  Bestandsdeployments.
- **`TeamsAgentInstaller`** (`teamsAgentInstaller.ts`): installiert die
  gelisteten Agenten-Apps via `teamsProvisioner@1` (M365-Connector) in ein
  Team — nie eigene Graph-Calls. Katalog-Auflösung: konfigurierte
  `teamsAppId` → feature-detektiertes `getCatalogApp` (Connector >= 0.3.1,
  nie vorausgesetzt) → typisiertes `not-in-catalog`-Fallback. Fehler-Mapping
  ohne Throw auf dem Message-Pfad: 403 → Fallback mit `missingScopes` +
  Tenant-Negative-Cache (15 min TTL, Scope-Union), 409 `already-existed` →
  Erfolg, 429 → begrenzter Retry mit `Retry-After`-Hint; lange Hints
  (> 30 s) sind non-retryable (Backpressure-Vertrag des Connectors), ein
  gedrosselter Run short-circuitet die restlichen Apps (429 ist
  tenant-weit), Gesamt-Sleep pro Run gedeckelt (60 s Budget).
- **Onboarding-Hook** (`teamsBot.ts`): wird einer unserer Bots in ein Team
  eingeladen (`membersAdded`, Team-Scope mit `aadGroupId`), läuft der
  Installer und postet eine Ergebnis-Card. Frisch installierte Agenten-Bots
  unterdrücken ihr Default-Intro über den In-Memory-Marker (2 min TTL) —
  nur der installierende Bot postet die Summary (`bot_added` wird zu
  `members_added` herabgestuft).
- **Ergebnis-/Fallback-Card** (`teamsCard.ts`): pro App eine Statuszeile,
  unterscheidet explizit „Admin-Zustimmung fehlt" (mit Scope-Liste) von
  „nicht im App-Katalog"; Install-Deep-Links
  `https://teams.microsoft.com/l/app/<teamsAppId>` (nur öffentliche IDs)
  und „🔄 Prüfen"-Re-Check, der den Installer erneut ausführt und die Card
  in-place aktualisiert. Team/Tenant der Prüfung kommen transportseitig aus
  `channelData` — Card-Daten sind client-editierbar und zählen nur als
  Fallback.
- **Wiring** (`plugin.ts`/`manifest.yaml`): `teamsProvisioner@1` wird spät
  über die Service-Registry aufgelöst (bare key, CHANNEL_RESOLVER-Muster)
  und unter `optional_requires:` deklariert — ohne Connector degradiert das
  Feature mit einer Log-Zeile, die Aktivierung bleibt unberührt. Neues
  Setup-Feld `teams_agent_apps` (type string, wie `teams_bots`).
- **Setup-Guide/README**: Consent-Abschnitt für `AppCatalog.ReadWrite.All`
  + `TeamsAppInstallation.ReadWriteForTeam.All` (Admin-Consent-URL, REST-
  `appRoleAssignments`-Fallback, Restart-nach-Consent, Renewed-Consent-
  Regel) in beiden Sprachvarianten.
- Deploy-Hinweis: Katalog-Lookup ohne konfigurierte `teamsAppId` braucht
  den M365-Connector >= 0.3.1 (`getCatalogApp`, feature-detected); ältere
  Connector-Versionen degradieren sauber auf die Fallback-Card.

## 0.20.0

Multi-Bot-Fundament (epic byte5ai/omadia#860, Wave W0a — Refs #15 #16 #17 #18 #19):

- **`teams_bots[]` config surface**: eine Plugin-Installation betreibt mehrere
  Teams-Bots (eine Entra-App-Registration / ein Azure-Bot pro omadia-Agent).
  Jeder Eintrag: `{botSlug, appId, tenantId, appPasswordSecretRef, appType?,
  displayName?}`. Passwörter kommen ausschließlich als Vault-Secret-Referenz —
  inline `appPassword` wird hart abgelehnt. Slugs: `^[a-z0-9][a-z0-9-]{0,62}$`.
- **Legacy-Shim**: ohne `teams_bots` mappen die Scalar-Credentials der
  Microsoft-365-Integration (`microsoft_app_id` / `microsoft_tenant_id` /
  `microsoft_app_password`, App-Typ aus `MICROSOFT_APP_TYPE`, Default
  MultiTenant) unverändert auf `teams_bots[0]` (`botSlug: 'default'`) —
  bestehende Single-Bot-Deployments laufen ohne jede Änderung weiter.
- **Per-Bot-Routing + Aliasse**: `POST /api/teams/:botSlug/messages` dispatcht
  auf den CloudAdapter + Credential-Factory des jeweiligen Bots; unbekannte
  Slugs → 404 (`teams.unknown_bot`), nie ein Fallback auf fremde Credentials.
  `/api/messages` (Live-Legacy-Pfad) und `/api/teams/messages`
  (Manifest-Doku-Pfad) bleiben Aliasse des Default-Bots (`teams_bots[0]`).
- **Per-Bot Conversation-Ref-Persistenz**: `teams_conversation_refs` wird pro
  `(conversation_id, bot_app_id)` geschrieben (bot_app_id lowercase); Casing
  wird überall über den kanonischen `teamsBotKey()`-Helper normalisiert.
  Proaktive Sends mit Bot-Attribution, die auf keinen konfigurierten Bot
  zeigt, schlagen bei >1 Bot laut fehl statt unter fremder Identität
  zuzustellen (Single-Bot: Fallback für App-Registration-Rotation bleibt).
- **Migration-Hinweis**: Die Cross-Bot-Isolation der Conversation-Refs braucht
  die Kernel-Migration 0010 (byte5ai/omadia, `ADD COLUMN bot_app_id TEXT NOT
  NULL DEFAULT ''` + Composite-PK). Auf Prä-0010-Kerneln degradiert der Store
  featuredetektiert auf das alte Single-Key-Verhalten; bei >1 konfiguriertem
  Bot warnt die Aktivierung laut. Der plugin-seitige Backfill ist idempotent
  und nicht-destruktiv (nur `UPDATE`, kein `DELETE`).
- **Doppelantwort-Policy**: Sind mehrere konfigurierte Bots in derselben
  Teams-Conversation installiert, beantwortet jeder Bot die an ihn
  adressierten Activities unabhängig (Teams stellt Channel-Posts nur dem
  @erwähnten Bot zu; Group-Chats je nach Tenant-Policy allen). Das Plugin
  dedupliziert nicht bot-übergreifend — Antwort-Hoheit wird über die
  `/operator/channels`-Bindings pro Bot-Catch-All (`28:<appId>`-Key je Bot)
  gesteuert.
- **App-Package-Template**: versioniertes `appPackage/manifest.json.template`
  (+ Icons + README) mit parametrisierten `developer.*`, `accentColor`,
  `commandLists` und dokumentierter JSON-Escaping-Pflicht; wird via
  `scripts/build-zip.mjs` ins Release-ZIP gestaged, damit der Provisioner
  (byte5ai/omadia#865) das Template aus dem installierten Plugin rendert.
- Fix: strip folded AI disclosure, gate Fresh Check on memoryUsed (#14).

## 0.19.1

- Declare every capability this channel resolves through `ctx.services.get`
  (omadia#838). `chatAgent@^1` stays under `requires:`; the thirteen further
  names go under `optional_requires:`, which grants the same declaration the
  service gate asks for without adding an activation prerequisite. Retires the
  `@omadia/channel-teams` row in `STANDALONE_LEGACY_SERVICE_GRANTS_2026_08_20`.

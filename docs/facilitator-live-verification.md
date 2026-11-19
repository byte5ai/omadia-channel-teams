# Live verification of native facilitators in Teams

Owner checklist for byte5ai/omadia-channel-teams#43 and #44 (epic
byte5ai/omadia#2043, kernel spec §6.2, §6.3 and acceptance criterion 13).

There is no Teams test tenant. The automated tests run against a simulated Bot
Framework, Bot Connector and Graph, and the evaluation of 2026-11-18 collected
read-only evidence from the byte5 tenant (see
[`teams-meeting-apis.md`](./teams-meeting-apis.md)). What neither could settle is listed
here as six owner steps, **M1 to M6**. Each step is one action in Teams; the observation
settles the open point. After that, the scenario walk-through checks the facilitator end
to end.

Record for each run: date, instance, plugin version, Teams client (desktop, web,
mobile) and who ran it.

## Preparation

1. An omadia instance with the kernel of #2043, `@omadia/channel-teams` ≥ 0.36.1 and the
   facilitator engine plugin.
2. Two facilitators with their own Teams identity, here **Mira** and **Jonas**, and one
   ordinary agent with its own identity, here **Otto**.
3. The instance's app registration (the one of the Microsoft 365 integration) is served
   by the Teams channel and holds `TeamsAppInstallation.ReadWriteForChat.All` and
   `…ReadWriteAndConsentForChat.All` (it does in byte5). The facilitator bots' own
   registrations need no Graph permission.
4. Access to the middleware log (`fly logs -a odoo-bot-middleware`). The lines to look
   for start with `[teams]`.

## Owner steps M1 to M6

### M1: channel delivery (done 2026-11-18, confirmed)

Result: the message reached the bot two seconds after the post, logged as
`[teams] drop non-mention conv=19:fdsWodnEDi7J3FI5tgPf1… type=channel`; the team shows
`ChannelMessage.Read.Group` for the app. Kept here so it can be repeated on another
tenant:

- **Do:** in a channel of a team where an agent app is installed, post one message
  without a mention.
- **Expect:** one `[teams] drop non-mention … type=channel` line (or, with a
  facilitation running there, `recorded for the engine, no turn`).
- **Settles:** delivery in channels under `ChannelMessage.Read.Group`, which is
  documented but was never observed.

### M2: edit, delete and reaction under RSC (done 2026-11-18, settled)

Result (meeting chat and channel thread, see "Recorded run" below): edit and delete
arrive for messages the bot only saw through RSC, with the id of the original message. A
reaction arrives only for a message **the bot** posted; a reaction to another person's
message produces no activity at all. `reactions` therefore stays out of the capability
snapshot.

- **Do:** in a test group chat with an agent app, post a message without a mention,
  edit it, react to it with a thumbs-up, delete it. Then react to a message **the bot**
  posted.
- **Expect in the log**, one line each:
  `[teams] messageUpdate eventType=editMessage … id=<id of the original message>`,
  `[teams] messageDelete eventType=softDeleteMessage … id=<same id>`, and for the
  reaction to the bot's post `[teams] messageReaction … replyToId=<id of the bot's post>`.
- **Note:** whether a `messageReaction` line appears for the reaction to **your own**
  message. That is the open question.
- **Settles:** that edit and delete events arrive for messages the bot only saw through
  RSC, that they carry the original message id, and whether reactions to other people's
  messages are delivered. If they are, `reactions` can be declared in the capability
  snapshot.

### M3: returned message id (done 2026-11-18, confirmed)

Result: every proactive post returned an id, and each id equals the id the Teams client
carries for that message (15 posts in the meeting chat, the posts of the channel thread).
"Copy link" itself was not read from the clipboard.

- **Do:** after one facilitator post, take the stored message id of that post and
  compare it with the id in "Copy link" of the same message in Teams.
- **Expect:** the same id.
- **Settles:** that the id the bot gets back is the id everybody else sees.

### M4: first-contact errors (done 2026-11-18, bodies confirmed)

The Bot Connector answers were verified live: not installed =
`403 BadArgument "Bot is not installed in user's personal scope"`, blocked =
`403 ConversationBlockedByUser` on the send. What remains is to see the product path
end to end:

- **Do:** with a colleague who never opened the facilitator's app, let a private
  reminder become due.
- **Expect:** no direct message; log line
  `[teams] direct message as 'Mira' not delivered: app_not_installed`; the group gets
  the neutral reminder.
- **Do:** the colleague installs the app, then blocks the bot. Let the next reminder
  become due.
- **Expect:** `… not delivered: blocked`.
- **Note:** if either line says `channel_error` instead, copy the error text.
- **Known from the live run:** a colleague who only shares a group chat with the app
  gets no direct message (`app_not_installed`); a colleague who is a member of a team
  where the app is installed does.
- **Settles:** that the plugin's own send path produces these outcomes.

### M5: meeting (delivery confirmed 2026-11-18, start and end events confirmed 2026-11-19)

Results so far:

- 2026-11-18: a message without mention in the chat of a scheduled meeting reached the
  bot as `[teams] drop non-mention conv=19:meeting_… type=groupChat`.
- 2026-11-19, test meeting: **the meeting start and end events arrived.** `meetingStart`
  at 05:36:01Z and `meetingEnd` at 06:02:17Z, one
  `[teams] meeting event name=… recorded=true` log line each.
- **Still open:** participant join and leave events (not seen: the app has no "Meeting
  event subscriptions" set in Developer Portal, see below), and the consent upgrade.

- **Do:** schedule a meeting, add an agent app to it, post in the meeting chat without
  a mention.
- **Expect:** a log line with a `19:meeting_…` conversation id (`drop non-mention`, or
  `recorded for the engine, no turn` with a facilitation running).
- **Do:** in a chat whose install lacks `ChatMessage.Read.Chat`, run the consent repair
  (Graph `installedApps/{id}/upgrade` with the full consented set).
- **Expect:** afterwards the grant list of the chat shows the permission and the
  readiness check passes.
- **Do:** with an agent app from the current package (`TEMPLATE_REVISION` 2 or later) in a
  **scheduled** meeting, start and end the meeting.
- **Expect:** two log lines
  `[teams] meeting event name=application/vnd.microsoft.meetingStart … recorded=true`
  and `…meetingEnd … recorded=true`.
- **Do:** set the participant event subscription for that agent app (see "Manual
  Developer Portal step" below), add the app to a new scheduled meeting **before**
  anybody joins, then let two people join and one leave.
- **Expect:** `…meetingParticipantJoin … recorded=true` and
  `…meetingParticipantLeave … recorded=true`, and from then on the facilitator's
  handshake or status mentions that it sees who is in the meeting.
- **Do:** repeat start, join and end in a **Meet-now** call started from a group chat.
- **Expect:** no participant lines. Note whether start and end lines appear; the
  documentation names scheduled and channel meetings only.
- **Do:** publish a new package version of an existing agent app to the catalog and
  check its portal subscription again.
- **Note:** whether the setting is still there.
- **Settles:** mention-free delivery in meeting chats, the consent semantics of the
  upgrade call, the exact event names and fields, and whether the portal subscription
  survives a catalog update.

#### Manual Developer Portal step (per agent app)

Teams sends participant join and leave events only to an app that subscribed to them
in Developer Portal. There is no manifest field and no documented API for it, so it
cannot be set while an agent is provisioned. Once per agent app that should see who is
in a meeting:

1. Open <https://dev.teams.microsoft.com/>, **Apps**, and open the agent's app (import
   it from the organisation's catalog if it is not listed).
2. In the section **Meeting event subscriptions** select **Participant join** and
   **Participant leave**.
3. Select **Save**.
4. Under **Configure → Permissions** check that `OnlineMeetingParticipant.Read.Chat`
   is listed (it is, for packages from `TEMPLATE_REVISION` 2 on).

Without this step everything else works; the facilitator just does not declare that it
sees who joins and leaves.

### M6: tenant RSC switch (done 2026-11-18, read)

Result in byte5: `GET /policies/authorizationPolicy` lists
`ManagePermissionGrantsForOwnedResource.microsoft-dynamically-managed-permissions-for-chat`
and `…-for-team` under `defaultUserRolePermissions.permissionGrantPoliciesAssigned`
(the Microsoft-managed default); `GET /groupSettings` is empty. For a customer tenant,
repeat:

- **Do:** as an administrator, read the resource-specific consent settings for chats
  and for teams (Teams admin center, or Graph as an admin).
- **Expect:** both allow app consent.
- **Settles:** the documented prerequisite for customer tenants. Record the values.

## RSC permissions: install proof

**Rule: every RSC permission the app package template requests must have been installed
once into a test team and once into a test chat before a release ships it.** Graph
accepts any declared set at publish and checks it only at install. A refusal is a `400`
that names the backend request (`IsUserAuthorizedToGrantGroupResourceSpecificPermissionsRequest`)
and no permission, so a refused set can only be found by installing variants.

Procedure for a new permission:

1. Build a throw-away app from a generated agent package: new manifest id, the RSC list
   to test, everything else unchanged. Publish it to the tenant app catalog.
2. Install it into the test team (`POST /teams/{id}/installedApps` with the full declared
   set as `consentedPermissionSet`) and into a test chat
   (`POST /chats/{id}/installedApps`). Both must answer `201`.
3. Read the grants of the team and the chat: the new permission must be listed.
4. Remove the installs and the catalog app.
5. Add a row to the table below and the permission to `LIVE_VERIFIED_RSC` in
   `tests/appPackageTemplate.test.ts`. That test fails for a permission that is in the
   template and not in the list.

### Record: bisect of 2026-11-18 (byte5 tenant, test team)

Found at the first team install of a facilitator's app. All installs were made with the
instance's app credentials and the full declared set; successful ones were removed again.

| Variant | RSC set | Team install |
|---|---|---|
| current package (0.36.1, revision 2) | all 11 | **400** |
| A | 11 minus `ChannelMeeting.ReadBasic.Group` | **400** |
| C | the 7 of the 0.35.0 template | **400** (also refused in the Teams client) |
| D | `ChannelMessage.Read.Group`, `TeamsActivity.Send.Group`, `ChatMessage.Read.Chat`, `TeamsActivity.Send.Chat`, `ChatMember.Read.Chat` | 201 |
| E | D + `ChannelMember.Read.Group` | 201 |
| F | D + `FileStorageContainer.Selected.Group` | **400** |
| G | D + `ChannelMeeting.ReadBasic.Group` | 201 |
| H | D + `OnlineMeeting.ReadBasic.Chat`, `OnlineMeetingParticipant.Read.Chat`, `OnlineMeetingNotification.Send.Chat` | 201 |
| I | all 11 minus `FileStorageContainer.Selected.Group` (the revision 3 set) | **201** |

- **Refused:** `FileStorageContainer.Selected.Group`, in every variant that carried it. It
  had been in the template since 2026-08-25, so no agent app from those packages could be
  installed into a team. Nothing in this plugin, the M365 connector or the kernel used
  it. Removed in revision 3 (0.36.2).
- **Team, verified:** the ten permissions of variant I. With I installed, the team's
  grants for the bot were `ChannelMessage.Read.Group`, `TeamsActivity.Send.Group`,
  `ChannelMember.Read.Group`, `ChannelMeeting.ReadBasic.Group`.
- **Chat, verified:** the revision 2 package (all eleven) installed into a meeting chat
  with `201`. The chat check accepts the three meeting permissions.
- **Not tested:** whether the refusal is specific to this tenant or team (the permission
  belongs to SharePoint Embedded and may need a container type registered in the
  tenant), and an install of the exact revision 3 package through the product.

### Installs from the revision 2 package (0.36.0 / 0.36.1)

Such an install (version `1.2.<n>`) is reported as `package_outdated` once 0.36.2 runs.
In a chat nothing is missing in practice: the removed permission is team-scoped. To clear
it, publish the agent's app again (the kernel builds version `1.3.<n>` with ten
permissions), then run the consent repair: it upgrades the install to the published
definition and sends exactly that definition's set. While the catalog still holds the
revision 2 app, the repair answers `package_outdated` and changes nothing. *Simulated.*

## Scenario walk-through

One section per scenario of `tests/facilitatorScenarios.test.ts`. A scenario counts as
verified only when every "Expected" line was observed.

### 1. Message without mention

1. Create a group chat with three people and add **Mira**.
2. Start a decision with a ranked vote: `@Mira …`.
   - Expected: the handshake; it says what Mira perceives.
3. One person votes with a mention, two without.
   - Expected: every vote is echoed to its author and the status names the right
     missing people.
   - Expected in the log: `recorded for the engine, no turn` for the two votes.
4. In a second chat with Mira but no facilitation, write without a mention.
   - Expected: no reaction; log `drop non-mention`.
5. Edit a vote, then quote somebody's vote and write "Anna agrees too".
   - Expected: the edit is confirmed to its author and counts; nothing is recorded for
     the quoted person. If something is, note the message: the quote markers are an
     open point.

### 2. Missing consent

1. Add Mira to a chat from an old app package without `ChatMessage.Read.Chat`.
2. Try to start a facilitation.
   - Expected: Mira does not start and says the permission is not granted here and that
     the installation has to be upgraded.
3. Run the consent repair, try again.
   - Expected: the facilitation starts.
4. Try in a chat where Mira's app is not installed at all.
   - Expected: it says the app is not installed there.

### 3. First-contact direct message

- See M4. Additionally: a person who already has Mira's app receives the private
  reminder **from Mira** (name and avatar), and the group gets only a neutral reminder
  that names nobody.

### 4. Blocked direct message

- See M4. Additionally: after the person unblocks Mira the next reminder arrives, and
  at no point does another bot write to them.

### 5. Restart

1. Run scenario 1 until two of four people voted, then restart the middleware.
2. The third person votes without a mention.
   - Expected: recorded and echoed.
3. Mira posts on its own (status, reminder).
   - Expected: it arrives, as Mira.
4. Somebody who edited a vote before the restart edits it again.
   - Expected: the latest text counts. (The kernel numbers the revisions; the
     adapter's own count restarting does not matter.)

### 6. Two facilitator bots in one chat

1. Add **Mira** and **Jonas** to one chat and start a facilitation with each.
2. Write one message without a mention.
   - Expected: each handles it for its own facilitation; every post carries the right
     name and avatar.
3. Add **Otto**; write without a mention, then `@Otto` a question.
   - Expected: Otto stays silent on the first and is the only one to answer the second.
   - Expected: no facilitator reacts to anything a bot wrote.

### 7. Scheduled meeting chat

1. Schedule a meeting, add Mira, start a facilitation in the meeting chat.
   - Expected: the readiness check passes; the handshake says Mira reads the chat and
     the polls and does **not** hear the call.
2. During the meeting write a topic into the chat without a mention.
   - Expected: it is collected.
3. Start a call from an ordinary group chat with a running facilitation and write into
   the chat.
   - Expected: the facilitation continues; no lifecycle events are expected for an
     ad-hoc call.

### 8. Poll answers

1. Let Mira run a live poll.
   - Expected: a text message and, as a separate message, a card with one button per
     option.
2. Three people click, at least one on mobile.
   - Expected: each sees "Antwort übermittelt." and gets their answer echoed.
3. One person clicks a different option afterwards.
   - Expected: the engine applies its rule for changed answers.
4. An anonymous participant (joined through the link without signing in) clicks.
   - Expected: the answer is not counted as a verified vote.
5. Click the card after the facilitation ended.
   - Expected: "Diese Abstimmung ist nicht mehr aktiv."

### 9. A participant joining late

1. Add a person to the meeting chat while the facilitation runs.
   - Expected: Mira welcomes them and counts them from now on.
2. They vote without a mention.
   - Expected: recorded for them.
3. In a scheduled meeting with the portal subscription set, a person joins the call
   twenty minutes late.
   - Expected in the log: `…meetingParticipantJoin … recorded=true`; the engine knows
     they arrived late.
4. The same in a Meet-now call, or without the portal subscription.
   - Expected: Mira does not notice until the person writes. This is the stated limit.

## AI marking (0.38.2): to verify in the client

Found on 2026-11-19 (rollout defect D24): the facilitator's proactive posts showed no AI
marking at all. The cause is not proven; see `teams-meeting-apis.md` §5b. 0.38.2 makes
every such post carry the instance's disclosure sentence again and sends the label entity
in Microsoft's shape. What only a Teams client can show:

| # | Do | Settles |
|---|---|---|
| A1 | After the update, let a facilitator post proactively (start a moderation). Look at the handshake in Teams desktop, web and mobile. | That the disclosure sentence is on the agent's first own post in the conversation (it must be: `teams_native_ai_label` is off). Note on which later posts it is absent: the kernel adds it to the first post per conversation |
| A2 | On the same messages, look at the **message header** next to the agent's name. | Whether Teams now shows "AI Generated" there (tooltip "AI-generated content may be incorrect."). 0.38.2 sends the entity with `"@id": ""`, as Microsoft's SDK does. Do not confuse it with "✨ AI-generated" **inside** an answer card: that is the card's own chip |
| A3 | Look at the header of a normal agent answer (a card) and of a very long answer (plain text). | Whether the platform label appears on replies in a turn at all. It has never been confirmed on any message of this plugin |
| A4 | Repeat A2 in a group chat, a channel thread, a meeting chat and a 1:1. | Microsoft's page names personal chats, group chats and channels; meeting chats are not mentioned |
| A5 | Set `TEAMS_LOG_OUTGOING_ACTIVITIES=1`, restart, post once proactively and once as a reply, then unset it. Keep the two `[teams] outgoing path=…` lines. | The exact activities that went to the Bot Connector, for a field-by-field comparison if the label still does not show |
| A6 | Only if A2 shows the label on proactive posts in every scope you use: switch **AI marking: rely on the Teams label** on, and repeat A1. | That the sentence is gone and the label alone marks the post. If the label shows only in some scopes, leave the switch off |

## Question cards (0.38.0): not run yet

Everything below is **unverified**: the question cards were built against Microsoft's
documentation and a simulated transport, and no card of this kind has been clicked in a
Teams client. The poll card of 0.36.0 has not been clicked live either. The server
enforces every rule itself, so none of these points can let a wrong answer through; they
decide whether the cards are pleasant to use or need a change.

Needed: at least three people in one group chat, one channel thread and one meeting
chat; one of them on Teams mobile (iOS or Android); one guest of the tenant; one
anonymous participant in a meeting. For each step note the client (desktop, web, iOS,
Android) and read the `[teams]` log lines and the kernel events.

| # | Do | Settles (what Microsoft's pages do not state) |
|---|---|---|
| Q1 | Post a single-choice card with three options, and one with eight. | That `Input.ChoiceSet` renders as radio buttons (`expanded`) and as a dropdown (`compact`) in Teams, in a group chat, a channel and a meeting chat, on desktop, web and mobile |
| Q2 | Post a multiple-choice card; select two options; send. | That `isMultiSelect` works in Teams and the value arrives as a comma-separated list |
| Q3 | On a card with a required question, click "Antwort senden" without choosing. | Whether the client checks `isRequired` before an `Action.Execute` click. Either the client shows its own message, or the card comes back with "Bitte eine Option wählen."; both are fine, note which. Teams web: the client checks it (see Q17) |
| Q4 | On a free-text card with `maxLength` 200, try to type or paste 300 characters. | Whether Teams enforces `maxLength`. If not, the card must come back with "Bitte höchstens 200 Zeichen." |
| Q5 | On a ranking card, give the same option two places and send. | How a returned card with an error line is displayed: the same card with the values kept and the red line, for the clicker only |
| Q6 | Submit a valid answer. | That the person sees their own answer card with "Ändern", that nobody else's card changes, and that the spelling `application/vnd.microsoft.card.adaptive` of the response type is accepted |
| Q7 | Stop the middleware for a moment while clicking, or delay it. | The invoke timeout: what the client shows when the answer takes too long (the adapter answers without waiting for the engine, so this should not happen in normal use) |
| Q8 | After answering, leave the chat, wait two minutes, come back. Repeat in a conversation with more than 60 members. | The per-person refresh: the own-answer card is shown again in a conversation of up to 60 members; above that the question card is shown and the message menu offers "Refresh Card" |
| Q9 | Let the engine close the step right away, and once more a day later. | That the update of the posted message is visible to everybody, replaces the personal views, and whether an update still works on an old message (no time limit is documented) |
| Q10 | As a guest of the tenant and as an anonymous meeting participant, click a card. | Whether the click carries `from.aadObjectId`. A click without one is recorded with `answererVerified: false` and the person gets "… kann dir nicht sicher zugeordnet werden" |
| Q11 | On an assignment card, search a person in the picker and send; click "Ich übernehme das"; try the picker as a guest. | That the people picker with `scope=currentContext` lists the conversation's members in a group chat, a channel and a meeting chat, and returns the Entra object id |
| Q12 | Post a card with five questions and open it on a phone and in the meeting side panel. | Size and scrolling on the narrowest surfaces; whether five questions are too many |
| Q13 | Click the poll card of 0.36.0 (two to six buttons). | That an `Action.Execute` click arrives at all and shows the acknowledgement; never done live |
| Q14 | Post a card, deploy (or restart the middleware), then click it and let the engine close it. Read the log at activation. | That the table `teams_question_cards` is created on the instance's database through the pooler (no "kept in memory only" line), that the click is accepted and the close update goes through after the restart |
| Q15 | Start a scheduled meeting, deploy during it, then read the stored capability snapshot; end the meeting and read it again. | That `teams_meeting_states` is created, that the snapshot says `online` after the restart and `chat` after the end. Meeting start and end events themselves were seen live on 2026-11-19; a restart between them was not |
| Q16 | Post an assignment card with one named candidate; click "Ich übernehme das" as one person and "Vorschlag senden" with the candidate chosen as another. | That the one-candidate card (0.38.4) is posted as a card and both buttons arrive; "Ich übernehme das" was never clicked live |
| Q17 | On a single-choice card click "Antwort senden" without a selection. | **Settled for Teams web, 2026-11-19 (captured outgoing log, 0.38.3):** `isRequired` is checked client-side. The client shows the input's static `errorMessage` ("Bitte eine Option wählen.") itself and most likely does not send the click. The `invokeResponse` seen at 06:56:47Z answered the automatic per-person `facilitation.card.view` refresh, which follows about one second after a card is shown; it was not a reply to a refused answer. The 0.38.4 assumption (a refresh wipes the error card) is not supported. Desktop and mobile: not checked |
| Q18 | On a ranking card give two ranks the same option and send, with `TEAMS_LOG_OUTGOING_ACTIVITIES=1` (from 0.38.5 the incoming invoke is logged next to its response). | A duplicate rank passes the client's required check, so this is the case that shows a server-side refusal of a choice: the card must come back with the values kept and the reason, and must still show it after the refresh one second later. Not seen live yet |
| Q19 | Post a card with a closing time, set the process kill switch, close the vote (the kernel refuses the card update), release the switch; wait until five to ten minutes after the closing time. On Fly the row `f942876d-…-v1` of 2026-11-19 is such a card already: it is closed by the first sweep after 0.40.2 starts. | That the card in the chat is replaced by its closed form without inputs ("Geschlossen am …", "Diese Abstimmung ist nicht mehr aktiv."), that its row in `teams_question_cards` is `closed`, and that the log has one line `question card '…' was still open past its closing time … and is closed now`. Whether Teams accepts the replacement of a message that old is not known; the row is closed either way (0.40.2). Not seen live yet |

**Server-side validation stays the authority.** What a client checks is a convenience
and differs by client; a click can also be crafted. Every submission is read against the
stored card on the server. Seen live on 2026-11-19: a consent objection without a reason
was refused by the server, and its error line stayed on the card.

Reported by the rollout session for 0.38.3 on 2026-11-19 (real clicks, not re-checked here):
ranking, single choice with "Ändern", a consent objection refused without a reason, an
idea, a questionnaire and the result card update passed. Teams' own "AI Generated"
header showed on no message.

Record the results in a new "Recorded run" section. An item that fails becomes an issue;
the fallback for every form is the chat ballot, which stays valid.

## Delivery modes and private messages (0.39.0): not run yet

Needs a kernel with `conversationDelivery` (K4.3) and a facilitation that reaches
`following_up`. Set `TEAMS_LOG_OUTGOING_ACTIVITIES=1` for P3 to P5. Nothing below was
seen in a Teams client; the tests run on the simulated transport.

| # | Do | Settles |
|---|---|---|
| P1 | In the home chat of a facilitation in follow-up, write a message without a mention, then one that mentions the facilitator. | Log: `drop non-mention` for the first; for the second `recorded for the engine (addressed: mention)` and an ordinary turn. **Count the facilitator's posts for the mention: exactly one** (the engine's or the turn's, not both) |
| P2 | In that chat use "Reply" on a reminder the facilitator posted, without a mention. | That the quote arrives with the reminder's id: log `recorded for the engine (addressed: reply)`, kernel event with `replyToMessageId` = the id the send returned. If the line is missing, capture the activity: the quote shape is then another one than the two the adapter reads |
| P3 | Let the facilitator send a reminder as a direct message; note `messageId` and `conversationId` of the send. In the 1:1, use "Reply" on it and send "erledigt". | Whether a quote in a 1:1 names the quoted message, and whether that id equals the id the Bot Connector returned for the send. Capture the inbound activity (`entities`, `attachments`, `replyToId`) |
| P4 | In the same 1:1 write "erledigt" as an ordinary message. | That it carries no reference (`replyToId` absent, no quote): documented for chats, to confirm for the bot activity |
| P5 | As a participant of a running facilitation, ask the facilitator in its 1:1 "Wie ist der Stand?"; then the same as a person who takes part in nothing. | The kernel's private status path (#2118): the participant gets the status, the other person the fixed refusal. The adapter's part is the author on the turn (`principal` = Entra object id) |
| P6 | Link the chat of a scheduled meeting where the facilitator's app was added without the all-messages permission; then repair the consent and check again. | Criterion 60 live: `consent_missing` for the meeting chat while the home chat stays ready; ready after the upgrade |
| P7 | Click a question card and react to a facilitator post in a chat in follow-up. | That card answers and reactions reach the process in `addressed` mode |

## Recorded run: 2026-11-18, byte5 tenant

The record issue #43 asks for. Source: the rollout log of that day and the evaluation
note; evidence files (handshake texts, database rows, outbox, screenshots) were kept by
the rollout session under `/tmp/rollout/ev` and `/tmp/rollout/shots` on the operator's
machine and are **not** in this repository.

| | |
|---|---|
| Date | 2026-11-18, 17:30 to 19:20 UTC |
| Instance | Fly `odoo-bot-middleware` / `odoo-bot-harness`, kernel `sha-4e490e2f` for the runs that passed |
| Plugins | `@omadia/channel-teams` 0.36.4 (0.36.1 to 0.36.3 for the earlier attempts), `@omadia/agent-facilitator` 1.0.1, `@omadia/integration-microsoft365` 0.9.1 |
| Teams client | Teams web in Chrome |
| Who | Marcel Wege's session, driven by the rollout agent session; the only human participant was Marcel |
| Facilitator | "Mona (Test)" with its own Teams identity, Claude Sonnet 5.5 |
| Where | a test team's General channel and the chat of a scheduled test meeting |

**Limits of this run.** One human only, so nothing about several participants was
exercised (completeness with more than one person, two people editing, a late joiner).
One facilitator bot was active per conversation; "two facilitator bots in one chat" was
not run. No call was joined (the browser's microphone prompt was not granted), so no
meeting event was seen.

### Owner steps

| Step | Result | Evidence |
|---|---|---|
| M1 channel delivery | **Confirmed** | a message without a mention reached the bot in the channel (`drop non-mention … type=channel`) |
| M2 edit | **Confirmed** | `messageUpdate eventType=editMessage`, kernel event `edited` revision 1 with the same message id; the edited ballot counted (meeting chat and channel thread) |
| M2 delete | **Confirmed** | `messageDelete eventType=softDeleteMessage` with the same id, kernel event `deleted` revision 1 |
| M2 reaction to the bot's message | **Confirmed** | `messageReaction`, kernel event `reaction {"name":"like"}` |
| M2 reaction to a person's message | **Not delivered** | no log line, no event. This is the answer to the open question |
| M3 returned message id | **Confirmed** | see M3 above |
| M4 first-contact errors | **Confirmed** | product path: the report to the initiator failed with `app_not_installed` (no personal install) and the facilitator said so in the group |
| M5 delivery in the meeting chat | **Confirmed** | every message without a mention was recorded (`recorded for the engine, no turn`) |
| M5 meeting start and end events | **Confirmed 2026-11-19** | not in this run: test meeting the next day, `meetingStart` 05:36:01Z and `meetingEnd` 06:02:17Z, one `recorded=true` log line each |
| M5 participant join / leave events, portal subscription | **Not seen** | no Developer Portal subscription was set for the app |
| M5 consent upgrade (`repairConsent`) | **Not run** | no install from an earlier package was left to upgrade |
| M6 tenant RSC switch | Read on 2026-11-18 | unchanged |

### Scenarios

| # | Scenario | Result |
|---|---|---|
| 1 | Message without mention | **Passed.** "weiter", "einverstanden", "Status" and the ballot, all without a mention, were stored as kernel events with the author Teams states and answered within about a second |
| 2 | Missing consent | **Not run** as a scenario. Seen instead: a team install was refused by Graph until 0.36.2 removed one RSC permission (see "RSC permissions: install proof") |
| 3 | First-contact direct message | **Passed for the refusal**: `app_not_installed`, said in the group. A successful first-contact DM was **not run** |
| 4 | Blocked direct message | **Not run** |
| 5 | Restart | **Partly.** A running facilitation survived a re-activation of the engine plugin and closed at its deadline. A middleware restart during a facilitation was not run |
| 6 | Two facilitator bots in one chat | **Not run.** A second agent bot in the same chat got the same messages and dropped them at its mention gate, as it should |
| 7 | Scheduled meeting chat | **Passed for the chat**: start by mention, handshake as the facilitator's own bot after 8 s, result and close in the chat, subscription released. Meeting start / end and participant events: not run |
| 8 | Poll answers | **Not run** (the runs voted by chat message) |
| 9 | A participant joining late | **Not run** |
| 10 | Channel thread | **Passed.** The handshake landed as the first reply under the root post; a reply in the thread arrived under the thread id and was recorded; a new root post in the same channel was not taken; the subscription key was the thread id |

### What the run found

- A facilitator could not be started at all until 0.36.3 (the turn origin lacked
  `conversationType`) and 0.36.4 (the kernel derived the conversation id from the
  session scope). Both are fixed and were verified live.
- **The participant list of the handshake was empty**, and with it a reminder reminded
  nobody. Cause on this side: the roster was read as the instance's default bot (fixed
  in 0.36.5) and a failed read came back as an empty list (fixed in 0.37.0: the status
  is then `unknown`). Not verified live since.
- Engine posts carried no AI label (fixed in 0.36.5 and 0.36.6). Not looked at in a
  Teams client since.
- The follow-up meeting ended with `no_addresses`: the participant had no mail address,
  because participants came from message events, which carry none. Addresses can only
  come from the member list (see `teams-meeting-apis.md`, "The member list").
- The log lines for edit, delete and reaction cut the conversation id at 24 characters,
  so the thread part could not be read (fixed in 0.36.6).

## After the run

- Every "Expected" that did not hold becomes an issue in
  `byte5ai/omadia-channel-teams` with the log lines.
- Still open: from M5 the consent upgrade and the participant join / leave events with
  the portal subscription (meeting start and end were seen on 2026-11-19), and the
  scenarios marked "Not run" above:
  missing consent, a successful first-contact DM, a blocked DM, a middleware restart
  during a facilitation, two facilitator bots in one chat, poll answers, a late joiner.
  Everything with more than one human participant is unverified.
- Not verified live since they were fixed: the participant list of the handshake
  (0.36.5, 0.37.0), the AI label on engine posts (0.36.5, 0.36.6).
- In `teams-meeting-apis.md`, change the label of each settled row from "Open" or
  "Simulated" to "Live" with the date.
- Acceptance criterion 13 of the kernel spec and the criteria of #43 and #44 that need
  a real Teams instance stay open until M1 to M6 and the walk-through are done.

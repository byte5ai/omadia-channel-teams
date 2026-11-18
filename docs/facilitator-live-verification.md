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

1. An omadia instance with the kernel of #2043, `@omadia/channel-teams` ≥ 0.36.0 and the
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

### M1: channel delivery

- **Do:** in a channel of a team where an agent app is installed, post one message
  without a mention.
- **Expect:** one `[teams] drop non-mention … type=channel` line (or, with a
  facilitation running there, `recorded for the engine, no turn`).
- **Settles:** delivery in channels under `ChannelMessage.Read.Group`, which is
  documented but was never observed.

### M2: edit, delete and reaction under RSC

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

### M3: returned message id

- **Do:** after one facilitator post, take the stored message id of that post and
  compare it with the id in "Copy link" of the same message in Teams.
- **Expect:** the same id.
- **Settles:** that the id the bot gets back is the id everybody else sees.

### M4: first-contact errors

- **Do:** with a colleague who never opened the facilitator's app, let a private
  reminder become due.
- **Expect:** no direct message; log line
  `[teams] direct message as 'Mira' not delivered: app_not_installed`; the group gets
  the neutral reminder.
- **Do:** the colleague installs the app, then blocks the bot. Let the next reminder
  become due.
- **Expect:** `… not delivered: blocked`.
- **Note:** if either line says `channel_error` instead, copy the error text: the plugin
  matches `ForbiddenOperationException` and `MessageWritesBlocked` only. Also note
  whether a colleague who is in a group chat that has the app, but never opened it, does
  get the message.
- **Settles:** the exact error bodies and the install precondition.

### M5: meeting

- **Do:** schedule a meeting, add an agent app to it, post in the meeting chat without
  a mention.
- **Expect:** a log line with a `19:meeting_…` conversation id (`drop non-mention`, or
  `recorded for the engine, no turn` with a facilitation running).
- **Do:** in a chat whose install lacks `ChatMessage.Read.Chat`, run the consent repair
  (Graph `installedApps/{id}/upgrade` with the full consented set).
- **Expect:** afterwards the grant list of the chat shows the permission and the
  readiness check passes.
- **Later, once the meeting RSC are in the app package:** expect `meetingStart` /
  `meetingEnd` and join/leave activities in a scheduled meeting, and none of them in a
  Meet-now call.
- **Settles:** mention-free delivery in meeting chats and the consent semantics of the
  upgrade call.

### M6: tenant RSC switch

- **Do:** as an administrator, read the resource-specific consent settings for chats
  and for teams (Teams admin center, or Graph as an admin).
- **Expect:** both allow app consent.
- **Settles:** the documented prerequisite for customer tenants. Record the values.

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
3. A person joins the call through the link without being added to the chat.
   - Expected: Mira does not notice. This is the stated limit of the baseline.

## After the run

- Every "Expected" that did not hold becomes an issue in
  `byte5ai/omadia-channel-teams` with the log lines.
- In `teams-meeting-apis.md`, change the label of each settled row from "Open" or
  "Simulated" to "Live" with the date.
- Acceptance criterion 13 of the kernel spec and the criteria of #43 and #44 that need
  a real Teams instance stay open until M1 to M6 and the walk-through are done.

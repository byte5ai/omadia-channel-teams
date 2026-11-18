# Live verification of native facilitators in Teams

Manual checklist for byte5ai/omadia-channel-teams#43 and #44 (epic
byte5ai/omadia#2043, kernel spec §6.2, §6.3 and acceptance criterion 13).

There is no Teams test tenant, so the automated tests
(`tests/facilitatorScenarios.test.ts` and the module tests next to it) run against a
simulated Bot Framework, Bot Connector and Graph. They prove what the plugin does
with a payload in the documented shape. They do **not** prove that Teams sends that
payload. This checklist closes that gap on a real instance. Each section is one
scenario of the automated suite; a scenario counts as verified only when every
"Expected" line was observed.

Record for each run: date, instance, plugin version, Teams client (desktop, web,
mobile) and who ran it.

## Preparation

1. An omadia instance with the kernel of #2043, `@omadia/channel-teams` ≥ 0.36.0 and
   the facilitator engine plugin.
2. Two facilitators with their own Teams identity, here called **Mira** and
   **Jonas**, and one ordinary agent with its own Teams identity, here **Otto**.
   Each has its own app registration, Azure bot and Teams app package.
3. The app packages are generated from `appPackage/manifest.json.template` of this
   version (they request `ChatMessage.Read.Chat`).
4. Each facilitator's app registration has the Graph application permission
   `TeamsAppInstallation.ReadForChat.All` with admin consent. Without it the readiness
   check answers `consent_unverifiable` (that outcome is checked in scenario 2).
5. Four test people: **A** (initiator), **B**, **C**, **D**. D has never used any of
   the three apps.
6. Access to the middleware log. The lines to look for start with `[teams]`.

## Scenario 1: message without mention

1. A creates a group chat with B and C and adds the app **Mira** to it. Teams shows a
   consent dialog for the app's permissions; A accepts.
2. A starts a decision with a ranked vote: `@Mira …`.
   - Expected: Mira answers with the handshake and names what it perceives.
3. A votes with a mention (`@Mira A > B > C`). B votes **without** a mention
   (`B > A > C`). C votes without a mention.
   - Expected: every vote is echoed to its author ("Recorded for you: …").
   - Expected: the status after each vote names the right missing people.
   - Expected in the log: `non-mention recorded, no turn` for B's and C's messages and
     no `drop non-mention` line for this chat.
4. In a second group chat with Mira added but **no** running facilitation, B writes a
   message without a mention.
   - Expected: no reaction, and the log shows `drop non-mention` for that chat.
5. B edits their vote in the first chat (`A > B > C`).
   - Expected: Mira confirms the change to B, and the count uses the edited vote.
   - This is the only check of an assumption the documentation leaves open: that Teams
     sends edit events for a message that did not mention the bot.
6. B deletes a message that was a vote.
   - Expected: the facilitator treats the vote as withdrawn (engine behaviour); in the
     log no error.
7. A replies to B's vote with "quote" and writes "Anna agrees too" style text for C.
   - Expected: nothing is recorded for C. If something is recorded, note the exact
     message: the quote markers the plugin looks for are an assumption.

## Scenario 2: missing consent

1. Create a group chat and add **Mira** using a tenant or chat where resource-specific
   consent for chats is disabled, or with an old app package without
   `ChatMessage.Read.Chat`.
2. A tries to start a facilitation with `@Mira …`.
   - Expected: Mira does not start. It says that the permission
     `ChatMessage.Read.Chat` is not granted in this chat and what to do.
3. Remove `TeamsAppInstallation.ReadForChat.All` from Mira's app registration (or use
   a facilitator without it), restart the middleware, and try again in a chat where
   consent **is** given.
   - Expected: Mira does not start and says that the consent could not be confirmed,
     naming `TeamsAppInstallation.ReadForChat.All`.
4. Restore the permission, restart, try again.
   - Expected: the facilitation starts.
5. Try to start a facilitation in a **team channel**.
   - Expected: Mira does not start and says channels are not supported.

## Scenario 3: first-contact DM

1. In the facilitation of scenario 1, let a reminder become due while C has not
   voted. C has the app Mira available (added to the chat) but never wrote to Mira
   1:1.
   - Expected: C receives a direct message **from Mira** (name and avatar of Mira, not
     of Otto or the default bot).
   - Expected: the group gets only a neutral reminder that does not name C.
2. Repeat with D, who was added to the chat after the app and has never opened it.
   - Expected, one of two outcomes, note which:
     - D receives the DM, or
     - no DM arrives, the log shows the outcome `app_not_installed`, and the group gets
       the neutral reminder. No other bot writes to D.
   - This is where the documented rule "the app must be installed for the user" meets
     reality: whether membership in a chat that has the app is enough is not something
     the simulation can know.

## Scenario 4: blocked DM

1. B opens the 1:1 chat with Mira and blocks the bot (chat menu → Block), or
   uninstalls the app for themselves.
2. Let a reminder for B become due.
   - Expected: no DM. The log shows the outcome `blocked`. The group gets the neutral
     reminder. No other bot writes to B.
   - Note the exact error text in the log: the plugin recognises `MessageWritesBlocked`
     and `ConversationBlockedByUser`. Any other text ends as `channel_error` and has to
     be added to `src/teamsOutboundMessage.ts`.
3. B unblocks Mira. Let the next reminder become due.
   - Expected: the DM arrives again.

## Scenario 5: restart

1. Run scenario 1 until two of four people have voted.
2. Restart the middleware.
3. Without anybody mentioning Mira, the third person votes.
   - Expected: the vote is recorded and echoed, the status is correct.
4. Let Mira post something on its own (status on request, reminder).
   - Expected: the post arrives, as Mira.
5. A person who voted **before** the restart edits their vote for the second time.
   - Expected: the latest text counts. Known limit: the plugin numbers this edit as
     revision 1 again. If the engine keeps the older edit instead, that is the kernel
     gap listed in `docs/teams-meeting-apis.md`.
6. Let a DM reminder become due for someone who received one before the restart.
   - Expected: it arrives in the same 1:1 chat as before.

## Scenario 6: two facilitator bots in one chat

1. Add **Mira** and **Jonas** to the same group chat. Start one facilitation with each.
2. B writes one message without a mention.
   - Expected: each facilitator handles it for its own facilitation; every post shows
     the name and avatar of the bot it belongs to; no post of Mira appears under Jonas
     or the other way round.
3. Add **Otto** (ordinary agent) to the chat. B writes a message without a mention.
   - Expected: Otto stays silent (log: `drop non-mention`), the facilitators record it.
4. `@Otto` a question.
   - Expected: only Otto answers.
5. Run Mira in a second chat at the same time.
   - Expected: status and results never mix between the two chats.

## Scenario 7: scheduled meeting chat

1. A schedules a Teams meeting with B and C and adds the app **Mira** to the meeting
   (meeting details → Apps, or in the meeting chat). Accept the consent dialog.
2. Before the meeting, A starts a facilitation in the meeting chat.
   - Expected: the readiness check passes. If it fails with "Graph does not know this
     chat", note it: that a meeting chat is a chat for
     `/chats/{id}/permissionGrants` is an assumption.
   - Expected: the handshake says Mira reads the chat and the polls and does **not**
     hear the call.
3. Join the meeting. B writes a topic into the meeting chat without a mention.
   - Expected: it is collected.
4. Start a call from an ordinary group chat that has a running facilitation ("Meet
   now" in the chat) and write into the chat during the call.
   - Expected: the facilitation continues as before. Note whether the conversation of
     the call chat is the same chat (assumption in `docs/teams-meeting-apis.md`).

## Scenario 8: poll answers

1. In the meeting chat, let Mira run a live poll.
   - Expected: a card with the question and one button per option, posted by Mira.
2. A, B and C each click an option, on desktop and at least one on mobile.
   - Expected: each answer is recorded for the person who clicked, and echoed to them.
   - Expected: no "Something went wrong" under the card.
3. B clicks a different option afterwards.
   - Expected: the engine applies its rule for changed answers and says so to B.
4. A forwards the card or quotes it and writes "C votes yes".
   - Expected: nothing is recorded for C.

## Scenario 9: a participant joining late

1. During the meeting, A adds D to the meeting (so D joins the meeting chat).
   - Expected: Mira welcomes D and counts D as a participant from now on.
2. D writes a vote without a mention.
   - Expected: recorded for D.
3. A person joins the call through the link **without** being added to the chat.
   - Expected: Mira does not notice. This is the documented limit of the baseline:
     participant events need `OnlineMeetingParticipant.Read.Chat`, which the app
     package does not request.

## After the run

- Every "Expected" that did not hold becomes an issue in
  `byte5ai/omadia-channel-teams` with the log lines.
- Update the labels in `docs/teams-meeting-apis.md`: a row confirmed here is no longer
  "Simulated" or "Assumed".
- Acceptance criterion 13 of the kernel spec and the tenant criteria of #43 and #44
  stay open until this checklist was run completely.

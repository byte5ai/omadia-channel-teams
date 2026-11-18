# Teams meeting APIs for native facilitators

Evaluation for byte5ai/omadia-channel-teams#44 (epic byte5ai/omadia#2043, kernel
spec §6.1 and §6.3): which Teams API delivers which input or output to a
facilitator in a scheduled meeting and in an ad-hoc call, and what this plugin
implements in 0.36.0.

## How far to trust this document

**Nothing in this document was verified against a live Teams tenant.** There is no
Teams test tenant for this project. Every statement carries one of three labels:

| Label | Meaning |
|---|---|
| **Simulated** | Implemented in this plugin and covered by automated tests that feed it activities and error responses in the payload shape Microsoft documents. This proves the plugin handles such a payload. It does not prove Teams sends it in that situation. |
| **Documented** | Stated in Microsoft Learn, read on 2026-11-18. The page is linked. Not exercised by this project. |
| **Assumed** | Stated by neither a test nor a page we read. Treat as a hypothesis until someone checks it. |

The manual checklist in [`facilitator-live-verification.md`](./facilitator-live-verification.md)
is how a **Simulated** or **Assumed** row becomes verified on a real instance.

## Result in one table

Capability names are the ones of the kernel contract (`SessionCapabilitySnapshot`).
"Scheduled" is a meeting created in the calendar, "ad-hoc" is "Meet now" or a call
started from a chat.

### Inputs

| Capability | Teams mechanism | Scheduled | Ad-hoc | In 0.36.0 | Evidence |
|---|---|---|---|---|---|
| `text` | Bot Framework `message` activities in the meeting chat. Without @mention only with RSC `ChatMessage.Read.Chat` granted in that chat. | yes | yes, if the call has a chat the app is installed in | **yes** | Plugin: Simulated. RSC delivers all messages of a group chat: Documented [1]. That a meeting chat counts as such a group chat: **Assumed** ([1] does not mention meeting chats). |
| `message_edits` | `messageUpdate` with `channelData.eventType` `editMessage` / `undeleteMessage`, `messageDelete` with `softDeleteMessage`. | yes | yes | **yes** | Plugin: Simulated. Event shapes: Documented [2]. Whether they arrive for messages that did not mention the bot: **Assumed** ([2] gives the scope as "All" and says nothing about RSC). |
| `reactions` | `messageReaction` with `reactionsAdded` / `reactionsRemoved`, `replyToId` is the message. | partly | partly | **yes, with a limit** | Plugin: Simulated. Teams sends the event only for reactions to messages **the bot itself sent**: Documented [3]. A reaction to another person's message never reaches the bot. |
| `card_answers` | `Action.Submit` on an Adaptive Card arrives as a `message` activity with `value`; `from` is the person who clicked. | yes | yes | **yes** (poll card) | Plugin: Simulated. That `from` is always the clicking person: **Assumed**, it is how Bot Framework activities work but we read no page that states it for cards in meeting chats. |
| `presence` (join / leave) | Bot events `onTeamsMeetingParticipantsJoin` / `Leave`; needs RSC `OnlineMeetingParticipant.Read.Chat` and the participant event subscription of the app. | scheduled only | no | no | Documented [4] ("Participant events are supported only for scheduled meetings"). Not implemented: the RSC permission is not in the app package. |
| `presence` (meeting start / end) | Bot events for meeting start and end; RSC `OnlineMeeting.ReadBasic.Chat`. | yes, also channel meetings | not stated | no | Documented [4]. Not implemented, same reason. |
| `hand_raise` | No bot API or event. | no | no | no | Documented by absence: [4] lists none. |
| `transcribed_speech` (live) | Real-time media bot that joins the call and receives audio; transcription is then our own job. Needs C#/.NET, the `Microsoft.Graph.Communications.Calls.Media` library and a Windows Server VM in Azure with a public IP. | possible | possible | no | Documented [5]. Out of reach for a Node.js plugin; this is #584. |
| `transcribed_speech` (after the meeting) | Graph transcripts, RSC `OnlineMeetingTranscript.Read.Chat` for scheduled private meetings, `CallTranscripts.Read.All` for ad-hoc calls. Available once generated **after** the meeting ends. | after the end | after the end | no | Documented [6]. No live transcript API is described there. |
| `shared_documents` | Files posted in the chat arrive as attachments of the `message` activity. | yes | yes | no | The plugin already stores attachments for agent turns; they are not part of message events. **Assumed** for meeting chats. |

### Outputs

| Capability | Teams mechanism | Scheduled | Ad-hoc | In 0.36.0 | Evidence |
|---|---|---|---|---|---|
| `text` | Proactive post into the meeting chat as the facilitator's own bot. | yes | yes | **yes** | Plugin: Simulated. |
| `mentions` | `<at>` token plus mention entity with the member id from the chat roster. | yes | yes | **yes** | Plugin: Simulated. That Teams notifies the person: **Assumed**. |
| `cards`, `polls` | Adaptive Card with one `Action.Submit` per option, posted into the chat. | yes | yes | **yes** | Plugin: Simulated. |
| `private_messages` | 1:1 conversation of the bot with the person; first contact needs the app installed for that person. | yes | yes | **yes** | Plugin: Simulated. Requirement and the 403 answers: Documented [7]. |
| `pinned_surface` | Meeting tab, side panel, stage (`shareAppContentToStage`), or the in-meeting notification (`OnlineMeetingNotification.Send.Chat`). All need a web surface or extra RSC. | possible | partly | no | Documented [4]. Not implemented. |
| `voice` | Real-time media bot (see `transcribed_speech`, live). | possible | possible | no | Documented [5]. #584. |
| `screen` | Share to stage through the Teams client SDK in a meeting app. | possible | not stated | no | Documented [4]. Not implemented. |

## What 0.36.0 implements

The baseline of the spec: **the chat of the meeting, and cards**.

1. **Facilitator bot in the meeting chat.** A meeting chat reaches the bot as a
   group chat (`conversationType: groupChat`, id `19:meeting_…@thread.v2`), so the
   rule of #43 applies unchanged: the kernel subscribes the facilitator's bot to the
   conversation, the bot then emits every human message as a message event, starts a
   turn only on an @mention, and posts as its own bot. Other bots in the same chat
   keep their mention gate. *Simulated; the conversation type and id format of a
   meeting chat are Assumed.*
2. **Poll cards with authenticated answers.** `buildPollCard` renders a question with
   up to six options. A click is emitted as a `card_answer` message event:
   `cardAnswer.cardId` is the poll id, `cardAnswer.value` is `{ optionId }`,
   `messageId` is the card message, `author` is `activity.from`. The payload of the
   click carries no person, so a name in it cannot count. A poll click never becomes a
   chat turn. *Simulated.*
3. **Capability snapshot.** `createTeamsMeetingSource` is registered as the meeting
   source. For a meeting chat it reports modality `online`, inputs `text`,
   `message_edits`, `reactions`, `card_answers`, outputs `text`, `mentions`, `cards`,
   `polls`, `private_messages`. For a chat the asked bot is not in it reports nothing,
   and the kernel falls back to its chat baseline. A snapshot is pushed once when a bot
   is added to a chat. *Simulated.*

With this snapshot the facilitator's handshake has to say that it reads the chat and
the polls and **does not hear the call** (spec §9 criterion 18).

## What is deliberately not in 0.36.0

- **Presence and late joiners through meeting events.** The app package does not
  request `OnlineMeetingParticipant.Read.Chat`. Adding an RSC permission changes
  every generated app package and needs the app to be updated and consented again in
  each chat [1], so it is a rollout decision and not a silent code change. Until
  then a late joiner is visible in two ways only: the `conversationUpdate` when a
  person is **added to the chat** (already emitted as a membership event), and their
  first message. Joining the call without being added to the chat is invisible.
- **`hybrid` modality.** Teams does not tell a bot whether several people sit in one
  room. The snapshot never claims it.
- **Hand raise.** No API.
- **Audio in and out.** Needs a .NET media bot on Windows Server; tracked in #584.
- **Pinned agenda, timer, parking lot as a surface.** Would need a meeting tab or the
  in-meeting notification permission. The facilitator uses chat posts instead.
- **Team channels.** The readiness check answers `unsupported_conversation` for
  channel conversations: the grant lookup works on `/chats/{id}` only. Channel
  meetings are therefore not supported.

## Ad-hoc calls

For "Meet now" and calls started from a chat the documentation we read is thin:

- Participant events exist for scheduled meetings only [4].
- Meeting details are documented for scheduled meetings, "Meet now", one-on-one and
  group calls [4].
- Whether a call started from an existing group chat keeps using that chat's
  conversation id, so that a facilitation running in the chat simply continues during
  the call: **Assumed**. If it holds, the baseline works in such a call without any
  change, because the bot only ever deals with the chat.
- Whether a "Meet now" meeting created outside a chat has a chat the app can be
  installed in before people join: **Assumed**.

## Known limits of the implementation

| Limit | Consequence |
|---|---|
| Edit revisions are counted in memory per bot, conversation and message. | After a restart of the middleware the next edit of an already edited message is emitted with revision 1 again. The kernel should order revisions by `occurredAt` or assign them itself. |
| Reactions reach the bot only on its own messages [3]. | "Confirm the echo with a reaction" works, because the echo is the bot's message. A reaction to another person's ballot is never seen. |
| Quoted and forwarded messages are recognised by attachment markers (`schema.skype.com/Reply`, `forwardedMessageReference`). | **Assumed** from payloads seen in the field, not from a page. If Teams marks them differently, such a message is emitted as an ordinary `created` event and the engine cannot tell it is a quote. |
| The kernel's `TargetedMessage` has no poll field. | The plugin reads `message.poll` structurally. The kernel contract has to add it before an engine can post a poll through the official type. |
| Consent is read from Graph `GET /chats/{id}/permissionGrants`, which needs `TeamsAppInstallation.ReadForChat.All` (application) or RSC `TeamsAppInstallation.Read.Chat` [8]. Neither is in the app package or documented setup today. | Until one is granted the readiness check answers `consent_unverifiable` and a facilitation does not start. |

## Sources

Microsoft Learn, read 2026-11-18.

1. Receive all messages with RSC: <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/channel-messages-with-rsc>
2. Message edit, undelete and soft delete events: <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/conversation-messages>
3. Conversation events, `messageReaction`: <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/subscribe-to-conversation-events>
4. Meeting apps APIs: <https://learn.microsoft.com/en-us/microsoftteams/platform/apps-in-teams-meetings/meeting-apps-apis>
5. Real-time media bots: <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/calls-and-meetings/real-time-media-concepts> and <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/calls-and-meetings/requirements-considerations-application-hosted-media-bots>
6. Meeting transcripts: <https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/meeting-transcripts/overview-transcripts>
7. Proactive messages: <https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/send-proactive-messages>
8. List permission grants of a chat: <https://learn.microsoft.com/en-us/graph/api/chat-list-permissiongrants>

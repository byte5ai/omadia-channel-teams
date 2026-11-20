# AGENTS.md

Rules for agents working in this repository.

## Rules

- No version bumps in pull requests: do not change `version` in
  `package.json`/`manifest.yaml`. Changes go under `## Unreleased` in
  `CHANGELOG.md`; the coordinator sets the version and releases.
- Gates before a PR: `npm run typecheck`, `npm run build`, `npm test` (Node 22).
- Declare a capability (`supports.*`) only when it is implemented. An
  own-identity operation (`asChannelKey`) goes out as exactly that bot or
  returns `unreachable`, never through another bot.

## Decisions

### Bot reactions (spec 2043 §5.14, omadia ADR 0027)

`react` is implemented and `supports.reactions: true` is declared
(`@omadia/channel-sdk` 0.3.9, `@omadia/plugin-api` 1.82.0).

- Platform route: Bot Connector
  `PUT {serviceUrl}/v3/conversations/{conversationId}/activities/{activityId}/reactions/{reactionType}`,
  no body, with the bot's own token. Sources:
  [Microsoft Learn, agent reactions](https://learn.microsoft.com/en-us/microsoftteams/platform/agents-in-teams/agent-reactions)
  ("agents can attach one or more different reactions to any message in a
  conversation"; two adds/removes per second across all conversations; a
  deleted message or a removed bot fails and must not be retried) and the
  Teams SDK `ReactionClient.add`
  (`microsoft/teams.ts`, `packages/api/src/clients/reaction/reaction.ts`).
- `acknowledge` maps to `like` (thumbs-up).
- botbuilder 4.23 has no method for this route. `src/teamsReactions.ts` makes
  the request inside the bot's proactive turn and signs it with the
  credentials of the Bot Connector client CloudAdapter puts into that turn
  (`turnState[adapter.ConnectorClientKey].credentials.signRequest`), after
  checking those credentials belong to the requested app id. Path ids are
  URL-encoded as botbuilder's connector client does.
- Addressing: `conversationId` and `messageId` exactly as the plugin's
  inbound message event stated them (`activity.conversation.id`,
  `inboundMessageId(activity)`), so a channel-thread message is addressed
  under its thread id.
- Not used: Graph `chatMessage: setReaction` (delegated permissions only).
- No new permission, RSC grant or outbound host.

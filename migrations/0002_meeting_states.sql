-- @omadia/channel-teams: meeting state per bot and conversation.
-- Applied by the kernel before activate() (permissions.sql, ADR-0050, omadia#838).
-- IF NOT EXISTS: installs from before 0.46 created this table at runtime.

CREATE TABLE IF NOT EXISTS teams_meeting_states (
  bot_key         TEXT        NOT NULL,
  conversation_id TEXT        NOT NULL,
  state           JSONB       NOT NULL,
  seen_at         TIMESTAMPTZ NOT NULL,
  expires_at      TIMESTAMPTZ NOT NULL,
  PRIMARY KEY (bot_key, conversation_id)
);

CREATE INDEX IF NOT EXISTS teams_meeting_states_expires_idx ON teams_meeting_states (expires_at);

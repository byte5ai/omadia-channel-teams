-- @omadia/channel-teams: question cards the facilitator posted.
-- Applied by the kernel before activate() (permissions.sql, ADR-0050, omadia#838).
-- IF NOT EXISTS: installs from before 0.46 created this table at runtime.

CREATE TABLE IF NOT EXISTS teams_question_cards (
  bot_app_id       TEXT        NOT NULL,
  card_id          TEXT        NOT NULL,
  conversation_id  TEXT        NOT NULL,
  message_id       TEXT,
  proposal_id      TEXT,
  proposal_version INTEGER,
  card             JSONB       NOT NULL,
  closes_at        TIMESTAMPTZ,
  state            TEXT        NOT NULL DEFAULT 'open',
  closed           JSONB,
  expires_at       TIMESTAMPTZ NOT NULL,
  updated_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
  PRIMARY KEY (bot_app_id, card_id)
);

CREATE INDEX IF NOT EXISTS teams_question_cards_message_idx ON teams_question_cards (bot_app_id, message_id);

CREATE INDEX IF NOT EXISTS teams_question_cards_expires_idx ON teams_question_cards (expires_at);

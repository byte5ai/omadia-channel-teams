# Changelog

## 0.19.1

- Declare every capability this channel resolves through `ctx.services.get`
  (omadia#838). `chatAgent@^1` stays under `requires:`; the thirteen further
  names go under `optional_requires:`, which grants the same declaration the
  service gate asks for without adding an activation prerequisite. Retires the
  `@omadia/channel-teams` row in `STANDALONE_LEGACY_SERVICE_GRANTS_2026_08_20`.

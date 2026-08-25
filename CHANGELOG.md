# Changelog

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

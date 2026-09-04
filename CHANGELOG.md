# Changelog

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

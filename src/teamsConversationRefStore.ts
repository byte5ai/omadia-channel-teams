// #330 field report — the per-conversation ConversationReference cache is the
// ONLY way to open a proactive turn (roster reads, group nudges via
// conversationSend), and it was in-memory: every middleware restart made
// proactive delivery answer `no_binding` until the conversation happened to
// produce a new inbound activity. This store write-throughs the cache into
// `teams_conversation_refs` (kernel graph migration 0009) so references
// survive restarts. Best-effort by design: no pool, a missing table, or a
// failing query degrade to the old cache-only behaviour — never a throw on
// the message path.

import type { ConversationReference } from 'botbuilder';
import type { Pool } from 'pg';

export interface TeamsConversationRefPersistence {
  save(conversationId: string, ref: Partial<ConversationReference>, teamsType?: string): Promise<void>;
  load(conversationId: string): Promise<{ ref: Partial<ConversationReference>; teamsType?: string } | undefined>;
}

/** Bot-Framework service hosts a persisted reference may point proactive
 *  turns at. Defense in depth (review M2): a poisoned DB row must not make
 *  `continueConversationAsync` present bot credentials to an attacker host —
 *  provenance on the WRITE path is JWT-authenticated inbound activities, this
 *  guards the read path. */
function isAllowedServiceUrl(serviceUrl: unknown): boolean {
  if (typeof serviceUrl !== 'string') return false;
  try {
    const host = new URL(serviceUrl).hostname.toLowerCase();
    return host === 'smba.trafficmanager.net' || host.endsWith('.botframework.com');
  } catch {
    return false;
  }
}

export class PgTeamsConversationRefStore implements TeamsConversationRefPersistence {
  constructor(
    private readonly pool: Pool,
    private readonly log?: (msg: string) => void,
  ) {}

  async save(conversationId: string, ref: Partial<ConversationReference>, teamsType?: string): Promise<void> {
    await this.pool.query(
      `INSERT INTO teams_conversation_refs (conversation_id, ref, teams_type, updated_at)
       VALUES ($1, $2, $3, now())
       ON CONFLICT (conversation_id)
       DO UPDATE SET ref = EXCLUDED.ref, teams_type = EXCLUDED.teams_type, updated_at = now()`,
      [conversationId, JSON.stringify(ref), teamsType ?? null],
    );
  }

  async load(conversationId: string): Promise<{ ref: Partial<ConversationReference>; teamsType?: string } | undefined> {
    try {
      const r = await this.pool.query<{ ref: Partial<ConversationReference>; teams_type: string | null }>(
        `SELECT ref, teams_type FROM teams_conversation_refs WHERE conversation_id = $1`,
        [conversationId],
      );
      const row = r.rows[0];
      if (!row) return undefined;
      if (!isAllowedServiceUrl(row.ref.serviceUrl)) {
        this.log?.(`[teams] persisted conversation ref for '${conversationId}' dropped — serviceUrl outside Bot Framework domains`);
        return undefined;
      }
      return { ref: row.ref, ...(row.teams_type ? { teamsType: row.teams_type } : {}) };
    } catch (err) {
      // Pre-migration kernel or transient pool trouble: cache-only behaviour.
      this.log?.(`[teams] conversation-ref load failed: ${err instanceof Error ? err.message : String(err)}`);
      return undefined;
    }
  }
}

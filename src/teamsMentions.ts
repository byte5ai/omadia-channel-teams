import type { ChatParticipant } from './kernel-types.js';

const MENTION_RE = /<at>([^<]+)<\/at>/g;

/** Bot-Framework mention entity shape (subset — what Teams needs). */
export interface MentionEntity {
  type: 'mention';
  text: string;
  mentioned: { id: string; name: string };
}

export interface MentionResolution {
  /** Entities to append to the outgoing Activity's `entities[]`. */
  entities: MentionEntity[];
  /** Display names in the answer text that had no roster match. */
  unresolved: string[];
  /** Whether the answer text contained at least one `<at>…</at>` token. */
  hasMentions: boolean;
}

/**
 * Scan `text` for `<at>Display Name</at>` tokens, match each against
 * `roster` by exact `displayName`, and produce the Mention-Entity list the
 * Bot Framework expects on the outgoing Activity.
 *
 * Duplicate mentions of the same name collapse into ONE entity (Teams
 * still highlights every occurrence of the matching `<at>…</at>` token).
 *
 * Unmatched names land in `unresolved`. The caller decides whether to
 * send anyway, strip the tokens, or abort — we prefer strip + log in the
 * Teams adapter so a typo doesn't ping a random person.
 */
export function resolveMentions(
  text: string,
  roster: readonly ChatParticipant[],
): MentionResolution {
  const entities: MentionEntity[] = [];
  const unresolved: string[] = [];
  const seen = new Set<string>();
  let hasMentions = false;
  for (const match of text.matchAll(MENTION_RE)) {
    hasMentions = true;
    const raw = match[1] ?? '';
    const name = raw.trim();
    if (name.length === 0) continue;
    if (seen.has(name)) continue;
    seen.add(name);
    const person = roster.find((r) => r.displayName === name);
    if (!person) {
      unresolved.push(name);
      continue;
    }
    entities.push({
      type: 'mention',
      text: `<at>${name}</at>`,
      mentioned: { id: person.channelUserId, name: person.displayName },
    });
  }
  return { entities, unresolved, hasMentions };
}

/**
 * Replace every `<at>Name</at>` token with just `Name`. Used when a
 * mention could not be resolved and we downgrade to plain text rather
 * than send a broken `<at>…</at>` literal to Teams.
 */
export function stripMentionTokens(text: string): string {
  return text.replace(MENTION_RE, (_, name: string) => name.trim());
}

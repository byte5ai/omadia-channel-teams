import type {
  PeopleDirectoryEntry,
  PeopleDirectoryLookup,
  PeopleDirectoryProvider,
} from '@omadia/channel-sdk';

import type { TeamsGraphResolver, GraphUser } from './teamsGraphResolver.js';

/**
 * Teams people directory — the kernel's `PeopleDirectoryProvider` for
 * `channelType: 'teams'` (Conductor human-step principal resolution:
 * "Eskalation an Teresita Biel" → her routable address).
 *
 * Backed by Graph `GET /users?$search=…` under the bot's own app
 * registration (client credentials). Needs the **`User.Read.All`
 * application permission** with admin consent — one more than the
 * channels-directory enrichment asks for. Without it Graph answers 403 and
 * this provider reports `unavailable/not_configured`, which the kernel
 * surfaces as a `partial` result: the operator keeps the free-text address
 * field, nothing breaks.
 *
 * `principalRef` is the user's mail (falling back to the UPN), lowercased —
 * the same key the Teams channel hands the kernel as `principalRef` on
 * inbound turns, so a picked person is exactly what reminders and awaits
 * match on. Users with neither mail nor UPN (resource accounts, guests
 * without a routable address) are skipped rather than offered as
 * unroutable entries.
 */
export function createTeamsPeopleDirectoryProvider(opts: {
  readonly resolver: Pick<TeamsGraphResolver, 'searchUsers'>;
  readonly log: (msg: string) => void;
}): PeopleDirectoryProvider {
  return {
    channelType: 'teams',
    async search(query, { limit }): Promise<PeopleDirectoryLookup> {
      const q = query.trim();
      if (q.length === 0) return { outcome: 'resolved', people: [], partial: false };
      try {
        const result = await opts.resolver.searchUsers(q, limit);
        return {
          outcome: 'resolved',
          people: result.users.map(toEntry).filter((e): e is PeopleDirectoryEntry => e !== null),
          partial: result.truncated,
        };
      } catch (err) {
        const status = (err as { status?: unknown }).status;
        const message = err instanceof Error ? err.message : String(err);
        if (status === 401 || status === 403) {
          opts.log(
            `[teams] people directory unavailable (graph ${String(status)}) — grant the User.Read.All application permission (admin consent) to resolve people by name`,
          );
          return { outcome: 'unavailable', code: 'not_configured', message };
        }
        return { outcome: 'unavailable', code: 'source_error', message };
      }
    },
  };
}

function toEntry(user: GraphUser): PeopleDirectoryEntry | null {
  // Inbound Teams turns bind by the user's MAIL (else the native id) — never by a UPN
  // guess — so only mail-bearing users are routable; the rest are skipped, not offered.
  const ref = (user.mail ?? '').trim().toLowerCase();
  if (ref.length === 0 || /\s/.test(ref)) return null;
  const displayName = (user.displayName ?? '').trim();
  return {
    principalRef: ref,
    displayName: displayName.length > 0 ? displayName : ref,
    channelUserId: user.id,
    ...(user.jobTitle && user.jobTitle.trim() ? { hint: user.jobTitle.trim() } : {}),
  };
}

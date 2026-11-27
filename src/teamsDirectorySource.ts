import type { DirectoryDeltaResult, DirectoryPersonChange, DirectorySource } from '@omadia/channel-sdk';

import { GraphRequestError, type GraphDeltaUser, type TeamsGraphResolver } from './teamsGraphResolver.js';

/**
 * ADR-0300 (omadia#2335) — the tenant's directory of people for the kernel.
 *
 * The kernel keeps colleagues of the Entra tenant (not accounts) as the anchor
 * for channel identities and per-person tokens, and holds no Graph
 * credentials of its own. This source reads Graph `/users/delta` under the
 * default bot's app registration (`User.Read.All`, the same grant the people
 * directory needs); the kernel stores the people, runs the schedule and
 * deprovisions leavers.
 *
 * Members only: a guest's mail lives in another tenant, so a guest is reported
 * as removed (a member turned guest leaves the directory). Graph 410 means the
 * delta token expired → `resync`; 401/403 means the grant is missing.
 */
export function createTeamsDirectorySource(opts: {
  readonly tenantId: string;
  readonly resolver: Pick<TeamsGraphResolver, 'usersDelta'>;
  readonly log: (msg: string) => void;
}): DirectorySource {
  return {
    source: 'entra',
    tenantId: opts.tenantId,
    async delta(cursor?: string): Promise<DirectoryDeltaResult> {
      try {
        const page = await opts.resolver.usersDelta(cursor);
        return {
          outcome: 'page',
          page: {
            changes: page.users.map(toChange),
            ...(page.nextLink ? { nextCursor: page.nextLink } : {}),
            ...(page.deltaLink ? { deltaCursor: page.deltaLink } : {}),
          },
        };
      } catch (err) {
        const status = err instanceof GraphRequestError ? err.status : undefined;
        if (status === 410) return { outcome: 'resync' };
        const message = err instanceof Error ? err.message : String(err);
        if (status === 401 || status === 403) {
          opts.log(
            `[teams] directory sync unavailable (graph ${String(status)}) — grant the User.Read.All application permission (admin consent)`,
          );
        }
        return { outcome: 'unavailable', message };
      }
    },
  };
}

function toChange(user: GraphDeltaUser): DirectoryPersonChange {
  if (user.removed || user.userType === 'Guest') return { externalId: user.id, removed: true };
  return {
    externalId: user.id,
    ...(user.displayName !== undefined ? { displayName: user.displayName } : {}),
    ...(user.userPrincipalName !== undefined ? { upn: user.userPrincipalName } : {}),
    ...(user.mail !== undefined ? { mail: user.mail } : {}),
    ...(user.accountEnabled !== undefined ? { accountEnabled: user.accountEnabled } : {}),
  };
}

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { createTeamsDirectorySource, TeamsGraphResolver } from '@omadia/channel-teams';

// ADR-0300 (omadia#2335): Graph /users/delta pages → kernel directory changes;
// guests and @removed users leave, 410 asks for a full read, foreign links are refused.

function jsonResponse(status: number, body: unknown): Response {
  return { ok: status >= 200 && status < 300, status, json: async () => body } as unknown as Response;
}

const TOKEN = { match: 'login.microsoftonline.com', body: { access_token: 'tok-1', expires_in: 3600 } };
const TENANT = '11111111-2222-4333-8444-555555555555';

function sourceWith(routes: Array<{ match: string; status?: number; body: unknown }>, calls: string[] = []) {
  const resolver = new TeamsGraphResolver({
    tenantId: TENANT,
    clientId: 'app',
    clientSecret: 'secret',
    log: () => undefined,
    fetchImpl: (async (input: RequestInfo | URL) => {
      const url = String(input);
      calls.push(url);
      for (const r of [TOKEN, ...routes]) if (url.includes(r.match)) return jsonResponse(r.status ?? 200, r.body);
      return jsonResponse(404, {});
    }) as typeof fetch,
  });
  return createTeamsDirectorySource({ tenantId: TENANT, resolver, log: () => undefined });
}

describe('teams directory source (ADR-0300)', () => {
  it('maps a full-read page with a next link', async () => {
    const calls: string[] = [];
    const source = sourceWith(
      [
        {
          match: '/users/delta',
          body: {
            value: [
              { id: 'a', displayName: 'Anna', mail: 'anna@contoso.com', userPrincipalName: 'anna@contoso.com', accountEnabled: true, userType: 'Member' },
              { id: 'g', displayName: 'Guest', userType: 'Guest' },
              { id: 'r', '@removed': { reason: 'deleted' } },
            ],
            '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users/delta?$skiptoken=x',
          },
        },
      ],
      calls,
    );
    assert.equal(source.tenantId, TENANT);
    const result = await source.delta();
    assert.equal(result.outcome, 'page');
    if (result.outcome !== 'page') return;
    assert.deepEqual(result.page.changes, [
      { externalId: 'a', displayName: 'Anna', upn: 'anna@contoso.com', mail: 'anna@contoso.com', accountEnabled: true },
      { externalId: 'g', removed: true },
      { externalId: 'r', removed: true },
    ]);
    assert.equal(result.page.nextCursor, 'https://graph.microsoft.com/v1.0/users/delta?$skiptoken=x');
    assert.ok(calls.some((u) => u.includes('/users/delta?$select=')));
  });

  it('hands out the delta link on the last page', async () => {
    const source = sourceWith([
      { match: '/users/delta', body: { value: [], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/users/delta?$deltatoken=d' } },
    ]);
    const result = await source.delta('https://graph.microsoft.com/v1.0/users/delta?$deltatoken=c');
    assert.equal(result.outcome, 'page');
    if (result.outcome === 'page') assert.equal(result.page.deltaCursor, 'https://graph.microsoft.com/v1.0/users/delta?$deltatoken=d');
  });

  it('asks for a full read on 410 and is unavailable without the grant', async () => {
    assert.deepEqual(await sourceWith([{ match: '/users/delta', status: 410, body: {} }]).delta('https://graph.microsoft.com/v1.0/users/delta?$deltatoken=old'), {
      outcome: 'resync',
    });
    const denied = await sourceWith([{ match: '/users/delta', status: 403, body: {} }]).delta();
    assert.equal(denied.outcome, 'unavailable');
  });

  it('never follows a stored link to another host', async () => {
    const calls: string[] = [];
    const result = await sourceWith([], calls).delta('https://evil.example/users/delta');
    assert.equal(result.outcome, 'unavailable');
    assert.ok(!calls.some((u) => u.includes('evil.example')));
  });
});

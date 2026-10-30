import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { createTeamsPeopleDirectoryProvider, GraphRequestError, TeamsGraphResolver } from '@omadia/channel-teams';

// People directory for Conductor human steps: Graph /users $search → routable
// principal refs; permission failures degrade to not_configured, never throw.

interface FetchCall {
  url: string;
  init?: RequestInit;
}

function jsonResponse(status: number, body: unknown): Response {
  return { ok: status >= 200 && status < 300, status, json: async () => body } as unknown as Response;
}

function mockFetch(routes: Array<{ match: string; status?: number; body: unknown }>, calls: FetchCall[]): typeof fetch {
  return (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    calls.push({ url, ...(init ? { init } : {}) });
    for (const r of routes) if (url.includes(r.match)) return jsonResponse(r.status ?? 200, r.body);
    return jsonResponse(404, {});
  }) as typeof fetch;
}

const TOKEN_ROUTE = { match: 'login.microsoftonline.com', body: { access_token: 'tok-1', expires_in: 3600 } };

function resolverWith(routes: Array<{ match: string; status?: number; body: unknown }>, calls: FetchCall[] = []): TeamsGraphResolver {
  return new TeamsGraphResolver({
    tenantId: 'tenant',
    clientId: 'app',
    clientSecret: 'secret',
    log: () => undefined,
    fetchImpl: mockFetch([TOKEN_ROUTE, ...routes], calls),
  });
}

const USERS = {
  value: [
    { id: 'u1', displayName: 'Teresita Biel', mail: 'Teresita.Biel@Byte5.de', userPrincipalName: 'tbiel@byte5.onmicrosoft.com', jobTitle: 'Finance' },
    { id: 'u2', displayName: 'Tero Noaddress' },
    { id: 'u3', displayName: 'UPN Only', userPrincipalName: 'upn.only@byte5.de' },
    { nope: true },
  ],
};

describe('TeamsGraphResolver.searchUsers', () => {
  it('queries /users with $search on name/mail/UPN, ConsistencyLevel eventual, bounded $top', async () => {
    const calls: FetchCall[] = [];
    const resolver = resolverWith([{ match: '/users?', body: USERS }], calls);
    const res = await resolver.searchUsers('Teresita "Biel"', 999);
    const call = calls.find((c) => c.url.includes('/users?'));
    assert.ok(call);
    assert.match(call.url, /\$search=%22displayName%3ATeresita%20%20Biel%22%20OR%20%22mail%3A/);
    assert.match(call.url, /\$top=50$/);
    assert.match(call.url, /\$filter=accountEnabled%20eq%20true%20and%20userType%20eq%20'Member'&\$count=true/);
    assert.equal((call.init?.headers as Record<string, string>)['ConsistencyLevel'], 'eventual');
    assert.equal(res.users.length, 3); // the malformed entry is skipped
    assert.equal(res.truncated, false);
  });

  it('flags truncation when Graph returns a next page and throws GraphRequestError with the status', async () => {
    const resolver = resolverWith([{ match: '/users?', body: { ...USERS, '@odata.nextLink': 'https://graph/next' } }]);
    assert.equal((await resolver.searchUsers('t', 5)).truncated, true);
    const denied = resolverWith([{ match: '/users?', status: 403, body: {} }]);
    await assert.rejects(denied.searchUsers('t', 5), (err: unknown) => err instanceof GraphRequestError && err.status === 403);
  });
});

describe('createTeamsPeopleDirectoryProvider', () => {
  it('a 200 without a value array is a GraphRequestError (source_error downstream), not an empty directory', async () => {
    const resolver = resolverWith([{ match: '/users?', body: { odd: true } }]);
    await assert.rejects(resolver.searchUsers('t', 5), (err: unknown) => err instanceof GraphRequestError && err.status === 0);
    const provider = createTeamsPeopleDirectoryProvider({ resolver, log: () => undefined });
    const lookup = await provider.search('t', { limit: 5 });
    assert.equal(lookup.outcome === 'unavailable' ? lookup.code : '', 'source_error');
  });

  it('maps users to routable entries: mail lowercased as principalRef, mail-less users skipped (UPN never binds)', async () => {
    const provider = createTeamsPeopleDirectoryProvider({ resolver: resolverWith([{ match: '/users?', body: USERS }]), log: () => undefined });
    assert.equal(provider.channelType, 'teams');
    const lookup = await provider.search('ter', { limit: 10 });
    assert.equal(lookup.outcome, 'resolved');
    if (lookup.outcome !== 'resolved') return;
    assert.equal(lookup.partial, false);
    assert.deepEqual(lookup.people, [
      { principalRef: 'teresita.biel@byte5.de', displayName: 'Teresita Biel', channelUserId: 'u1', hint: 'Finance' },
    ]);
  });

  it('reports not_configured on 401/403, source_error otherwise, and never throws', async () => {
    const logs: string[] = [];
    const denied = createTeamsPeopleDirectoryProvider({ resolver: resolverWith([{ match: '/users?', status: 403, body: {} }]), log: (m) => logs.push(m) });
    const l1 = await denied.search('ter', { limit: 5 });
    assert.equal(l1.outcome, 'unavailable');
    assert.equal(l1.outcome === 'unavailable' ? l1.code : '', 'not_configured');
    assert.match(logs[0] ?? '', /User\.Read\.All/);

    const broken = createTeamsPeopleDirectoryProvider({ resolver: resolverWith([{ match: '/users?', status: 503, body: {} }]), log: () => undefined });
    const l2 = await broken.search('ter', { limit: 5 });
    assert.equal(l2.outcome === 'unavailable' ? l2.code : '', 'source_error');

    const empty = await broken.search('   ', { limit: 5 });
    assert.deepEqual(empty, { outcome: 'resolved', people: [], partial: false });
  });

  it('passes truncation through as partial', async () => {
    const provider = createTeamsPeopleDirectoryProvider({
      resolver: resolverWith([{ match: '/users?', body: { ...USERS, '@odata.nextLink': 'x' } }]),
      log: () => undefined,
    });
    const lookup = await provider.search('t', { limit: 2 });
    assert.equal(lookup.outcome === 'resolved' ? lookup.partial : false, true);
  });
});

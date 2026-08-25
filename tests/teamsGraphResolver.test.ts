import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  buildTeamsChannelKeyDirectory,
  TeamsConversationObserver,
  TeamsGraphResolver,
} from '@omadia/channel-teams';
import type { TeamsGraphResolver as ResolverType } from '@omadia/channel-teams';

// Graph name/member resolution for the channels directory: resolver
// (token reuse, caps, permission degrade, TTL), observer capture of the
// Graph-relevant ids, and the directory merge that keeps listKeys()
// synchronous-fast.

interface FetchCall {
  url: string;
  init?: RequestInit;
}

function jsonResponse(status: number, body: unknown): Response {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
  } as unknown as Response;
}

function mockFetch(
  routes: Array<{ match: string; status?: number; body: unknown }>,
  calls: FetchCall[],
): typeof fetch {
  return (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    calls.push({ url, ...(init ? { init } : {}) });
    for (const r of routes) {
      if (url.includes(r.match)) {
        return jsonResponse(r.status ?? 200, r.body);
      }
    }
    return jsonResponse(404, {});
  }) as typeof fetch;
}

async function settled(): Promise<void> {
  for (let i = 0; i < 10; i += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

const TOKEN_ROUTE = {
  match: 'login.microsoftonline.com',
  body: { access_token: 'tok-1', expires_in: 3600 },
};

function makeResolver(opts: {
  routes: Array<{ match: string; status?: number; body: unknown }>;
  calls: FetchCall[];
  logs?: string[];
  now?: () => number;
}): TeamsGraphResolver {
  return new TeamsGraphResolver({
    tenantId: 'tenant-1',
    clientId: 'app-1',
    clientSecret: 'secret-1',
    log: (m) => opts.logs?.push(m),
    fetchImpl: mockFetch(opts.routes, opts.calls),
    ...(opts.now ? { now: opts.now } : {}),
  });
}

describe('TeamsGraphResolver', () => {
  it('resolves group-chat topic + members, caps names, reuses the token', async () => {
    const calls: FetchCall[] = [];
    const members = Array.from({ length: 12 }, (_, i) => ({
      displayName: `Member ${String(i + 1)}`,
    }));
    const resolver = makeResolver({
      calls,
      routes: [
        TOKEN_ROUTE,
        { match: '/members', body: { value: members } },
        { match: '/chats/', body: { topic: 'Projekt Phoenix' } },
      ],
    });

    resolver.prime({
      conversationId: '19:abc@thread.v2',
      conversationType: 'groupChat',
    });
    await settled();

    const resolved = resolver.get('19:abc@thread.v2');
    assert.ok(resolved);
    assert.equal(resolved.label, 'Teams · Projekt Phoenix');
    assert.equal(resolved.members?.length, 8);
    assert.equal(resolved.memberCount, 12);

    // Second conversation: token endpoint must NOT be hit again.
    resolver.prime({
      conversationId: '19:def@thread.v2',
      conversationType: 'groupChat',
    });
    await settled();
    const tokenCalls = calls.filter((c) =>
      c.url.includes('login.microsoftonline.com'),
    );
    assert.equal(tokenCalls.length, 1);
  });

  it('degrades on 403, caches the negative result, logs once', async () => {
    const calls: FetchCall[] = [];
    const logs: string[] = [];
    const resolver = makeResolver({
      calls,
      logs,
      routes: [
        TOKEN_ROUTE,
        { match: '/chats/', status: 403, body: {} },
        { match: '/members', status: 403, body: {} },
      ],
    });

    resolver.prime({
      conversationId: '19:abc@thread.v2',
      conversationType: 'groupChat',
    });
    await settled();
    assert.equal(resolver.get('19:abc@thread.v2'), undefined);

    const graphCallsAfterFirst = calls.length;
    // Within the TTL a re-prime must not refetch (negative cache).
    resolver.prime({
      conversationId: '19:abc@thread.v2',
      conversationType: 'groupChat',
    });
    await settled();
    assert.equal(calls.length, graphCallsAfterFirst);
    assert.equal(logs.length, 1);
    assert.match(logs[0]!, /graph resolve resolution unavailable/);
  });

  it('re-resolves after the TTL lapses', async () => {
    const calls: FetchCall[] = [];
    let clock = 1_000_000;
    const resolver = makeResolver({
      calls,
      now: () => clock,
      routes: [
        TOKEN_ROUTE,
        { match: '/members', body: { value: [{ displayName: 'Alice' }] } },
        { match: '/chats/', body: { topic: 'T' } },
      ],
    });

    resolver.prime({
      conversationId: '19:abc@thread.v2',
      conversationType: 'groupChat',
    });
    await settled();
    const afterFirst = calls.length;

    clock += 11 * 60 * 1000; // beyond the 10-min TTL
    resolver.prime({
      conversationId: '19:abc@thread.v2',
      conversationType: 'groupChat',
    });
    await settled();
    assert.ok(calls.length > afterFirst);
  });

  it('resolves team members via the aad group id for channel conversations', async () => {
    const calls: FetchCall[] = [];
    const resolver = makeResolver({
      calls,
      routes: [
        TOKEN_ROUTE,
        {
          match: '/teams/aad-group-1/members',
          body: { value: [{ displayName: 'Alice' }, { displayName: 'Bob' }] },
        },
      ],
    });

    resolver.prime({
      conversationId: '19:chan@thread.tacv2',
      conversationType: 'channel',
      teamAadGroupId: 'aad-group-1',
    });
    await settled();
    const resolved = resolver.get('19:chan@thread.tacv2');
    assert.deepEqual(resolved?.members, ['Alice', 'Bob']);
    assert.equal(resolved?.memberCount, 2);
    assert.equal(resolved?.label, undefined);
  });

  it('never calls Graph for @thread.skype group chats (Graph answers 400 for them)', async () => {
    const calls: FetchCall[] = [];
    const resolver = makeResolver({
      calls,
      routes: [TOKEN_ROUTE],
    });
    resolver.prime({
      conversationId: '19:9cdb0cd5c91049ad99667e30b0795de6@thread.skype',
      conversationType: 'groupChat',
    });
    await settled();
    assert.equal(calls.length, 0);
    assert.equal(
      resolver.get('19:9cdb0cd5c91049ad99667e30b0795de6@thread.skype'),
      undefined,
    );
  });

  it('still resolves @thread.v2 group chats via Graph', async () => {
    const calls: FetchCall[] = [];
    const resolver = makeResolver({
      calls,
      routes: [
        TOKEN_ROUTE,
        { match: '/members', body: { value: [{ displayName: 'Alice' }] } },
        { match: '/chats/', body: { topic: 'Planung' } },
      ],
    });
    resolver.prime({
      conversationId: '19:abc123@thread.v2',
      conversationType: 'groupChat',
    });
    await settled();
    assert.equal(resolver.get('19:abc123@thread.v2')?.label, 'Teams · Planung');
  });

  it('resolves the tenant org name for the catch-all', async () => {
    const calls: FetchCall[] = [];
    const resolver = makeResolver({
      calls,
      routes: [
        TOKEN_ROUTE,
        { match: '/organization', body: { value: [{ displayName: 'byte5' }] } },
      ],
    });
    resolver.primeOrg();
    await settled();
    assert.equal(resolver.getOrgName(), 'byte5');
  });
});

describe('TeamsConversationObserver Graph fields', () => {
  function turn(activity: Record<string, unknown>): TurnContext {
    return { activity } as unknown as TurnContext;
  }

  it('captures teamAadGroupId, peer name for DMs, and fires onObserved', () => {
    const observedIds: string[] = [];
    const observer = new TeamsConversationObserver((conv) =>
      observedIds.push(conv.conversationId),
    );

    observer.observe(
      turn({
        conversation: { id: '19:chan@thread.tacv2', conversationType: 'channel' },
        channelData: {
          channel: { name: 'general' },
          team: { name: 'Phoenix', aadGroupId: 'aad-1' },
        },
      }),
    );
    observer.observe(
      turn({
        conversation: { id: 'a:dm1', conversationType: 'personal' },
        from: { id: 'u1', name: 'Alice Adams' },
      }),
    );

    const list = observer.list();
    const chan = list.find((c) => c.conversationId === '19:chan@thread.tacv2');
    assert.equal(chan?.teamAadGroupId, 'aad-1');
    const dm = list.find((c) => c.conversationId === 'a:dm1');
    assert.equal(dm?.peerName, 'Alice Adams');
    assert.equal(dm?.label, 'Teams · DM · Alice Adams');
    assert.deepEqual(observedIds, ['19:chan@thread.tacv2', 'a:dm1']);
  });

  it('enriches group conversations with the Bot-Framework roster and keeps it across re-observes', async () => {
    const fetched: string[] = [];
    const observer = new TeamsConversationObserver(undefined, async (ctx) => {
      const id = (ctx as unknown as { activity: { conversation: { id: string } } })
        .activity.conversation.id;
      fetched.push(id);
      return { names: ['Alice Adams', 'Bob Brown'], count: 2 };
    });

    const groupTurn = turn({
      conversation: {
        id: '19:9cdb@thread.skype',
        conversationType: 'groupChat',
      },
    });
    observer.observe(groupTurn);
    await new Promise((resolve) => setTimeout(resolve, 0));

    let conv = observer.list().find((c) => c.conversationId === '19:9cdb@thread.skype');
    assert.deepEqual(conv?.members, ['Alice Adams', 'Bob Brown']);
    assert.equal(conv?.memberCount, 2);

    // Re-observe must not drop the roster while the refresh is in flight.
    observer.observe(groupTurn);
    conv = observer.list().find((c) => c.conversationId === '19:9cdb@thread.skype');
    assert.deepEqual(conv?.members, ['Alice Adams', 'Bob Brown']);

    // Personal chats never trigger a roster fetch.
    observer.observe(
      turn({
        conversation: { id: 'a:dm1', conversationType: 'personal' },
        from: { id: 'u1', name: 'Alice Adams' },
      }),
    );
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.deepEqual(
      fetched.filter((id) => id === 'a:dm1'),
      [],
    );
  });

  it('uses conversation.name as the group-chat label when Teams sends it', () => {
    const observer = new TeamsConversationObserver();
    observer.observe(
      turn({
        conversation: {
          id: '19:9cdb@thread.skype',
          conversationType: 'groupChat',
          name: 'Projekt Phoenix',
        },
      }),
    );
    const conv = observer
      .list()
      .find((c) => c.conversationId === '19:9cdb@thread.skype');
    assert.equal(conv?.label, 'Teams · Projekt Phoenix');
  });

  it('derives a members label for unnamed group chats and keeps it across re-observes', async () => {
    const observer = new TeamsConversationObserver(undefined, async () => ({
      names: [
        'Teresita Biel - byte5',
        'Christian Köhler - byte5',
        'Christian Wendler - byte5',
        'Marcel Wege - byte5',
      ],
      count: 4,
    }));
    const groupTurn = turn({
      conversation: { id: '19:9cdb@thread.skype', conversationType: 'groupChat' },
    });
    observer.observe(groupTurn);
    await new Promise((resolve) => setTimeout(resolve, 0));

    let conv = observer
      .list()
      .find((c) => c.conversationId === '19:9cdb@thread.skype');
    assert.equal(
      conv?.label,
      'Teams · Group chat: Teresita, Christian, Christian +1',
    );

    // Re-observe (roster refresh in flight) must not regress the label
    // to the opaque-id fallback.
    observer.observe(groupTurn);
    conv = observer
      .list()
      .find((c) => c.conversationId === '19:9cdb@thread.skype');
    assert.equal(
      conv?.label,
      'Teams · Group chat: Teresita, Christian, Christian +1',
    );
  });

  it('a real group name wins over the members-derived label', async () => {
    const observer = new TeamsConversationObserver(undefined, async () => ({
      names: ['Alice'],
      count: 1,
    }));
    observer.observe(
      turn({
        conversation: {
          id: '19:9cdb@thread.skype',
          conversationType: 'groupChat',
          name: 'Projekt Phoenix',
        },
      }),
    );
    await new Promise((resolve) => setTimeout(resolve, 0));
    const conv = observer
      .list()
      .find((c) => c.conversationId === '19:9cdb@thread.skype');
    assert.equal(conv?.label, 'Teams · Projekt Phoenix');
    assert.deepEqual(conv?.members, ['Alice']);
  });
});

describe('buildTeamsChannelKeyDirectory Graph merge', () => {
  function fakeResolver(opts: {
    byId?: Record<
      string,
      { label?: string; members?: readonly string[]; memberCount?: number }
    >;
    orgName?: string;
  }): ResolverType {
    return {
      prime: () => undefined,
      primeOrg: () => undefined,
      get: (id: string) => opts.byId?.[id],
      getOrgName: () => opts.orgName,
    } as unknown as ResolverType;
  }

  function observerWith(
    convs: Array<{ conversationId: string; conversationType?: string }>,
  ): TeamsConversationObserver {
    const observer = new TeamsConversationObserver();
    for (const c of convs) {
      observer.observe({
        activity: {
          conversation: {
            id: c.conversationId,
            ...(c.conversationType
              ? { conversationType: c.conversationType }
              : {}),
          },
        },
      } as unknown as TurnContext);
    }
    return observer;
  }

  it('merges resolved labels/members and prefers the org name for the catch-all', async () => {
    const directory = buildTeamsChannelKeyDirectory({
      microsoftAppId: '737c6ddd-6d4e-4599-8dc3-260281ea906e',
      microsoftTenantId: 'tenant-1',
      conversationObserver: observerWith([
        { conversationId: '19:abc@thread.v2', conversationType: 'groupChat' },
      ]),
      graphResolver: fakeResolver({
        orgName: 'byte5',
        byId: {
          '19:abc@thread.v2': {
            label: 'Teams · Projekt Phoenix',
            members: ['Alice', 'Bob'],
            memberCount: 5,
          },
        },
      }),
    });

    const keys = await directory.listKeys();
    const catchAll = keys[0]!;
    assert.equal(catchAll.label, 'Teams · byte5 (all)');
    const chat = keys.find((k) => k.key === '19:abc@thread.v2');
    assert.equal(chat?.label, 'Teams · Projekt Phoenix');
    assert.deepEqual(chat?.members, ['Alice', 'Bob']);
    assert.equal(chat?.memberCount, 5);
  });

  it('falls back to observer roster members when Graph resolved nothing, capped at 8', async () => {
    const observer = new TeamsConversationObserver();
    observer.observe({
      activity: {
        conversation: {
          id: '19:9cdb@thread.skype',
          conversationType: 'groupChat',
        },
      },
    } as unknown as TurnContext);
    observer.noteMembers('19:9cdb@thread.skype', {
      names: Array.from({ length: 12 }, (_, i) => `Roster ${String(i + 1)}`),
      count: 12,
    });

    const directory = buildTeamsChannelKeyDirectory({
      microsoftAppId: 'app-1',
      microsoftTenantId: 'tenant-1',
      conversationObserver: observer,
      graphResolver: fakeResolver({}),
    });
    const keys = await directory.listKeys();
    const chat = keys.find((k) => k.key === '19:9cdb@thread.skype');
    assert.equal(chat?.members?.length, 8);
    assert.equal(chat?.memberCount, 12);
  });

  it('prefers Graph members over the observer roster', async () => {
    const observer = new TeamsConversationObserver();
    observer.observe({
      activity: {
        conversation: { id: '19:abc@thread.v2', conversationType: 'groupChat' },
      },
    } as unknown as TurnContext);
    observer.noteMembers('19:abc@thread.v2', {
      names: ['Roster Only'],
      count: 1,
    });

    const directory = buildTeamsChannelKeyDirectory({
      microsoftAppId: 'app-1',
      microsoftTenantId: 'tenant-1',
      conversationObserver: observer,
      graphResolver: fakeResolver({
        byId: {
          '19:abc@thread.v2': { members: ['Graph Member'], memberCount: 3 },
        },
      }),
    });
    const keys = await directory.listKeys();
    const chat = keys.find((k) => k.key === '19:abc@thread.v2');
    assert.deepEqual(chat?.members, ['Graph Member']);
    assert.equal(chat?.memberCount, 3);
  });

  it('keeps observer labels when nothing is resolved and lets displayLabel win', async () => {
    const directory = buildTeamsChannelKeyDirectory({
      microsoftAppId: 'app-1',
      microsoftTenantId: 'tenant-1',
      displayLabel: 'Production · Marketing',
      conversationObserver: observerWith([
        { conversationId: '19:abc@thread.v2', conversationType: 'groupChat' },
      ]),
      graphResolver: fakeResolver({ orgName: 'byte5' }),
    });

    const keys = await directory.listKeys();
    assert.equal(keys[0]!.label, 'Production · Marketing');
    const chat = keys.find((k) => k.key === '19:abc@thread.v2');
    assert.match(chat!.label, /Group chat/);
    assert.equal(chat!.members, undefined);
  });
});

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { TeamsAttachmentStore, graphMessageLookup } from '@omadia/channel-teams';

// Live finding 2026-11-21: at a facilitation start in a channel thread the
// attachment lookup asked Graph `/chats/{id}/messages/{id}` with the channel
// thread id and got 404 "Invalid ThreadId". A channel message is read under
// its team; without a team id or a connector that has the route, nothing is
// asked at all.

const CHANNEL = '19:abc@thread.tacv2';
const THREAD = `${CHANNEL};messageid=1700000000000`;
const ATTACHMENT = { contentType: 'reference', contentUrl: 'https://contoso.sharepoint.com/f.pdf', name: 'f.pdf' };

function fakeGraph(withChannelRoute = true) {
  const calls: Array<{ route: string; args: unknown }> = [];
  const client: Record<string, unknown> = {
    fetchChatMessage: async (chatId: string, messageId: string) => {
      calls.push({ route: 'chat', args: { chatId, messageId } });
      return { id: messageId, attachments: [ATTACHMENT] };
    },
  };
  if (withChannelRoute) {
    client['fetchChannelMessage'] = async (input: unknown) => {
      calls.push({ route: 'channel', args: input });
      return { id: 'x', attachments: [ATTACHMENT] };
    };
  }
  return { client, calls };
}

function store(client: Record<string, unknown>, logs: string[], now?: () => number) {
  return new TeamsAttachmentStore({
    tigris: {} as never,
    pool: {} as never,
    tenant: 't',
    graphClient: client as never,
    log: (m) => logs.push(m),
    ...(now ? { now } : {}),
  });
}

function failingChannelGraph(status: number) {
  const { client, calls } = fakeGraph();
  client['fetchChannelMessage'] = async (input: unknown) => {
    calls.push({ route: 'channel', args: input });
    throw Object.assign(new Error(`graph channelMessage ${String(status)} channel=c msg=m body=secret text`), {
      status,
      ...(status === 403 ? { code: 'not_permitted' } : {}),
    });
  };
  return { client, calls };
}

describe('attachment lookup through Graph: chat vs. channel thread', () => {
  it('reads a reply in a channel thread under its team, never as a chat', async () => {
    const { client, calls } = fakeGraph();
    const logs: string[] = [];
    const found = await store(client, logs).discoverFromGraphMessage(THREAD, '1700000000999', 'team-guid');
    assert.deepEqual(calls, [
      { route: 'channel', args: { teamId: 'team-guid', channelId: CHANNEL, rootMessageId: '1700000000000', messageId: '1700000000999' } },
    ]);
    assert.equal(found.length, 1);
  });

  it('reads the thread root itself with root = message', () => {
    const { client } = fakeGraph();
    const lookup = graphMessageLookup(client as never, THREAD, '1700000000000', 'team-guid');
    assert.equal(lookup.kind, 'fetch');
  });

  it('skips a channel message without a team id instead of asking Graph a wrong path', async () => {
    const { client, calls } = fakeGraph();
    const logs: string[] = [];
    assert.deepEqual(await store(client, logs).discoverFromGraphMessage(THREAD, '1700000000999'), []);
    assert.deepEqual(calls, []);
    assert.ok(logs.some((l) => l.includes('graph-discover skip') && l.includes('without team id')));
  });

  it('skips a channel message when the connector has no channel route', async () => {
    const { client, calls } = fakeGraph(false);
    const logs: string[] = [];
    assert.deepEqual(await store(client, logs).discoverFromGraphMessage(THREAD, '1', 'team-guid'), []);
    assert.deepEqual(calls, []);
    assert.ok(logs.some((l) => l.includes('no channel message lookup')));
  });

  it('keeps the chat route for chats', async () => {
    const { client, calls } = fakeGraph();
    await store(client, []).discoverFromGraphMessage('19:x@thread.v2', '42');
    assert.deepEqual(calls, [{ route: 'chat', args: { chatId: '19:x@thread.v2', messageId: '42' } }]);
  });

  it('a 403 is "not permitted": one log line, then the team is skipped for an hour', async () => {
    const { client, calls } = failingChannelGraph(403);
    const logs: string[] = [];
    let t = 1_000_000;
    const s = store(client, logs, () => t);
    assert.deepEqual(await s.discoverFromGraphMessage(THREAD, '1', 'team-guid'), []);
    assert.deepEqual(await s.discoverFromGraphMessage(THREAD, '2', 'team-guid'), []);
    assert.equal(calls.length, 1);
    const lines = logs.filter((l) => l.includes('graph-discover'));
    assert.equal(lines.length, 1);
    assert.match(lines[0] ?? '', /not permitted team=team-guid channel=19:abc@thread\.tacv2/);
    assert.ok(!lines[0]?.includes('secret text'));
    // another team is still asked
    await s.discoverFromGraphMessage(THREAD, '3', 'other-team');
    assert.equal(calls.length, 2);
    // after the hour the team is asked again
    t += 60 * 60 * 1000;
    await s.discoverFromGraphMessage(THREAD, '4', 'team-guid');
    assert.equal(calls.length, 3);
  });

  it('other errors keep the FAIL log and are not remembered', async () => {
    const { client, calls } = failingChannelGraph(500);
    const logs: string[] = [];
    const s = store(client, logs);
    await s.discoverFromGraphMessage(THREAD, '1', 'team-guid');
    await s.discoverFromGraphMessage(THREAD, '2', 'team-guid');
    assert.equal(calls.length, 2);
    assert.equal(logs.filter((l) => l.includes('graph-discover FAIL')).length, 2);
  });
});

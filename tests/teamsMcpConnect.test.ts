import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  buildMcpConnectCard,
  connectBaseUrl,
  createTeamsUiShellRouter,
  mcpConnectTaskModule,
  readMcpConnectValue,
  MCP_CONNECT_VALUE_TYPE,
} from '@omadia/channel-teams';

// omadia#2312 (ADR-0301): a Connect card opens the connect dialog via task/fetch; no
// link is posted, the dialog URL is built from the deployment's own https origin.

describe('teams MCP connect (omadia#2312)', () => {
  it('the card opens a task/fetch dialog and carries no URL', () => {
    const card = buildMcpConnectCard({ serverId: 'srv-1', serverName: 'Notion' });
    const content = card.content as { actions: Array<{ data: Record<string, unknown> }> };
    assert.deepEqual(content.actions[0]?.data, {
      msteams: { type: 'task/fetch' },
      type: MCP_CONNECT_VALUE_TYPE,
      serverId: 'srv-1',
      serverName: 'Notion',
    });
    assert.ok(!JSON.stringify(card).includes('http'));
  });

  it('reads only its own dialog data', () => {
    assert.equal(readMcpConnectValue({ type: 'other' }), undefined);
    assert.equal(readMcpConnectValue({ type: MCP_CONNECT_VALUE_TYPE, serverId: '' }), undefined);
    assert.deepEqual(readMcpConnectValue({ type: MCP_CONNECT_VALUE_TYPE, serverId: 'srv-1' }), {
      type: MCP_CONNECT_VALUE_TYPE,
      serverId: 'srv-1',
      serverName: 'srv-1',
    });
  });

  it('builds the dialog on the https origin only', () => {
    assert.equal(connectBaseUrl({ PUBLIC_BASE_URL: 'http://localhost:3979' }), undefined);
    assert.equal(connectBaseUrl({ TEAMS_PUBLIC_BASE_URL: 'https://bot.example/x', PUBLIC_BASE_URL: 'https://other' }), 'https://bot.example');
    const task = mcpConnectTaskModule({ type: MCP_CONNECT_VALUE_TYPE, serverId: 'a b', serverName: 'N&M' }, 'https://bot.example') as {
      task: { value: { url: string } };
    };
    assert.equal(task.task.value.url, 'https://bot.example/p/channel-teams/connect?server=a%20b&name=N%26M');
  });

  it('serves the connect page as a data-free shell', async () => {
    const router = createTeamsUiShellRouter();
    const layer = (router as unknown as { stack: Array<{ route?: { path: string } }> }).stack.find(
      (l) => l.route?.path === '/connect',
    );
    assert.ok(layer, 'GET /connect is registered');
  });
});

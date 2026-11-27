import { CardFactory, type Attachment } from 'botbuilder';

/**
 * byte5ai/omadia#2312 (omadia ADR-0301) — a Connect action for an MCP server a
 * turn could not use because the person has not connected it
 * (`SemanticAnswer.mcpConnectRequired`). The button opens a Teams dialog
 * (`task/fetch`) on the connect page; the kernel binds the connection to the
 * person's Teams SSO token. No link that could be forwarded is posted.
 */
export const MCP_CONNECT_VALUE_TYPE = 'omadia.mcp.connect';

export interface McpConnectRequest {
  readonly serverId: string;
  readonly serverName: string;
}

export interface McpConnectValue {
  readonly type: typeof MCP_CONNECT_VALUE_TYPE;
  readonly serverId: string;
  readonly serverName: string;
}

const MAX_NAME_LENGTH = 120;

export function buildMcpConnectCard(request: McpConnectRequest): Attachment {
  const name = request.serverName.slice(0, MAX_NAME_LENGTH);
  return CardFactory.adaptiveCard({
    type: 'AdaptiveCard',
    version: '1.5',
    body: [{ type: 'TextBlock', wrap: true, text: `🔒 ${name}` }],
    actions: [
      {
        type: 'Action.Submit',
        title: 'Verbinden',
        data: {
          msteams: { type: 'task/fetch' },
          type: MCP_CONNECT_VALUE_TYPE,
          serverId: request.serverId,
          serverName: name,
        } satisfies McpConnectValue & { msteams: { type: string } },
      },
    ],
  });
}

/** The `task/fetch` data of a Connect button, or `undefined` for any other dialog. */
export function readMcpConnectValue(data: unknown): McpConnectValue | undefined {
  if (data === null || typeof data !== 'object') return undefined;
  const d = data as Record<string, unknown>;
  if (d['type'] !== MCP_CONNECT_VALUE_TYPE) return undefined;
  const serverId = typeof d['serverId'] === 'string' ? d['serverId'].trim() : '';
  if (serverId === '' || serverId.length > 200) return undefined;
  const serverName = typeof d['serverName'] === 'string' && d['serverName'].trim() ? d['serverName'].trim() : serverId;
  return { type: MCP_CONNECT_VALUE_TYPE, serverId, serverName: serverName.slice(0, MAX_NAME_LENGTH) };
}

/**
 * The deployment's public origin, where the connect page lives
 * (`/p/channel-teams/connect`) — the same host the Teams app package lists as
 * valid domain. `undefined` when none is configured: no Connect card then.
 */
export function connectBaseUrl(env: Readonly<Record<string, string | undefined>> = process.env): string | undefined {
  for (const key of ['TEAMS_PUBLIC_BASE_URL', 'PUBLIC_BASE_URL', 'ATTACHMENT_PUBLIC_BASE_URL']) {
    const raw = env[key]?.trim();
    if (!raw) continue;
    try {
      const url = new URL(raw);
      if (url.protocol === 'https:') return url.origin;
    } catch {
      // not a URL — try the next one
    }
  }
  return undefined;
}

/** The `task/fetch` answer: the connect page in a Teams dialog. */
export function mcpConnectTaskModule(value: McpConnectValue, baseUrl: string): Record<string, unknown> {
  const url = `${baseUrl}/p/channel-teams/connect?server=${encodeURIComponent(value.serverId)}&name=${encodeURIComponent(value.serverName)}`;
  return {
    task: {
      type: 'continue',
      value: { title: value.serverName, url, fallbackUrl: url, height: 320, width: 480 },
    },
  };
}

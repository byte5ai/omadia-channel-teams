// Simulated Bot Framework transport for the native-facilitator tests
// (byte5ai/omadia#2043, issues #43 / #44).
//
// There is no Teams test tenant. The activities below follow the payload
// shapes Microsoft documents for the Teams channel (field names, id prefixes,
// `channelData.eventType`, entity lists); the values are invented. A test
// that passes here proves the plugin handles a payload of that SHAPE. It does
// not prove Teams sends that payload in a given situation. That part is
// checked by hand with docs/facilitator-live-verification.md.

import type { Activity, TurnContext } from 'botbuilder';

import { TeamsBot, TeamsConversationMessages } from '@omadia/channel-teams';
import type { ConversationMessageEvent } from '../../src/kernel-types.js';

export const TENANT_ID = '72f988bf-86f1-41af-91ab-2d7cd011db47';
export const SERVICE_URL = 'https://smba.trafficmanager.net/emea/';
export const GROUP_CHAT_ID = '19:b8a1c0f2d3e44f5a9b6c7d8e9f0a1b2c@thread.v2';
export const OTHER_GROUP_CHAT_ID = '19:0f9e8d7c6b5a4f3e2d1c0b9a8f7e6d5c@thread.v2';
export const MEETING_CHAT_ID = '19:meeting_MjdhNjM4YzUtYzExZi00OTFkLWE5YjctNTVkYTQ0ZGVlM2Jh@thread.v2';

export interface SimBot {
  appId: string;
  name: string;
  key: string;
}

export const MIRA: SimBot = {
  appId: '3d78d742-eefb-4fb2-bae5-3687f24c46fc',
  name: 'Mira',
  key: '28:3d78d742-eefb-4fb2-bae5-3687f24c46fc',
};
export const JONAS: SimBot = {
  appId: '9c1b5a7e-2f4d-4c8a-b6e1-0a3f7d5c2b91',
  name: 'Jonas',
  key: '28:9c1b5a7e-2f4d-4c8a-b6e1-0a3f7d5c2b91',
};

export interface SimUser {
  id: string;
  aadObjectId: string;
  name: string;
}

export const ANNA: SimUser = {
  id: '29:1fY2xQmB7kq8ZVnD0eJpA3sLhRtUwXc5GiOy6MbN4vE9KrTz',
  aadObjectId: '5b1e2c3d-4f5a-4b6c-8d7e-9f0a1b2c3d4e',
  name: 'Anna Köhler',
};
export const BERT: SimUser = {
  id: '29:1Qw3ErT5yU7iO9pA2sD4fG6hJ8kL0zX1cV3bN5mM7qW9eR2t',
  aadObjectId: '6c2f3d4e-5a6b-4c7d-9e8f-0a1b2c3d4e5f',
  name: 'Bert Maier',
};
export const CARLA: SimUser = {
  id: '29:1Zx2Cv3Bn4Mm5Qq6Ww7Ee8Rr9Tt0Yy1Uu2Ii3Oo4Pp5Aa6Ss7',
  aadObjectId: '7d3a4e5f-6b7c-4d8e-8f9a-1b2c3d4e5f6a',
  name: 'Carla Voss',
};

let nextMessageId = 1_731_930_000_000;
/** Teams message ids are epoch-millisecond strings. */
export function newMessageId(): string {
  nextMessageId += 1;
  return String(nextMessageId);
}

interface BaseInput {
  bot: SimBot;
  from: SimUser;
  conversationId?: string;
  conversationType?: 'groupChat' | 'personal' | 'channel';
  id?: string;
  timestamp?: string;
}

function baseActivity(input: BaseInput): Partial<Activity> {
  const conversationType = input.conversationType ?? 'groupChat';
  return {
    id: input.id ?? newMessageId(),
    timestamp: new Date(input.timestamp ?? '2026-11-18T09:15:00.000Z'),
    serviceUrl: SERVICE_URL,
    channelId: 'msteams',
    from: { id: input.from.id, name: input.from.name, aadObjectId: input.from.aadObjectId },
    conversation: {
      ...(conversationType === 'personal' ? {} : { isGroup: true }),
      conversationType,
      tenantId: TENANT_ID,
      id: input.conversationId ?? GROUP_CHAT_ID,
    } as Activity['conversation'],
    recipient: { id: `28:${input.bot.appId}`, name: input.bot.name },
    channelData: { tenant: { id: TENANT_ID } },
  };
}

/** A chat message. `mention` adds the `<at>` token and the mention entity the
 *  way the Teams client does when a person picks the bot from the @ list. */
export function messageActivity(input: BaseInput & { text: string; mention?: boolean }): Partial<Activity> {
  const base = baseActivity(input);
  const text = input.mention ? `<at>${input.bot.name}</at> ${input.text}` : input.text;
  return {
    ...base,
    type: 'message',
    textFormat: 'plain',
    locale: 'de-DE',
    text,
    entities: [
      ...(input.mention
        ? [{ type: 'mention', text: `<at>${input.bot.name}</at>`, mentioned: { id: `28:${input.bot.appId}`, name: input.bot.name } }]
        : []),
      { type: 'clientInfo', locale: 'de-DE', country: 'DE', platform: 'Web', timezone: 'Europe/Berlin' },
    ],
  };
}

/** A reply that quotes an earlier message: Teams adds an HTML attachment
 *  whose blockquote names the quoted message id and its author. */
export function quotedReplyActivity(
  input: BaseInput & { text: string; quotedMessageId: string; quotedAuthor: SimUser; quotedText: string },
): Partial<Activity> {
  const message = messageActivity(input);
  const quoteHtml =
    `<blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="${input.quotedMessageId}">` +
    `<strong itemprop="mri" itemid="${input.quotedAuthor.id}">${input.quotedAuthor.name}</strong>` +
    `<span itemprop="time" itemid="${input.quotedMessageId}"></span>` +
    `<p itemprop="preview">${input.quotedText}</p></blockquote><p>${input.text}</p>`;
  return { ...message, attachments: [{ contentType: 'text/html', content: quoteHtml }] };
}

/** `messageUpdate` with `eventType: editMessage` — the edited message keeps
 *  its id and carries the new text. */
export function editActivity(input: BaseInput & { id: string; text: string }): Partial<Activity> {
  const base = baseActivity(input);
  return {
    ...base,
    type: 'messageUpdate',
    text: input.text,
    textFormat: 'plain',
    channelData: { eventType: 'editMessage', tenant: { id: TENANT_ID } },
  };
}

/** `messageDelete` with `eventType: softDeleteMessage`. */
export function deleteActivity(input: BaseInput & { id: string }): Partial<Activity> {
  return {
    ...baseActivity(input),
    type: 'messageDelete',
    channelData: { eventType: 'softDeleteMessage', tenant: { id: TENANT_ID } },
  };
}

/** `messageUpdate` with `eventType: undeleteMessage`. */
export function undeleteActivity(input: BaseInput & { id: string; text: string }): Partial<Activity> {
  return {
    ...baseActivity(input),
    type: 'messageUpdate',
    text: input.text,
    channelData: { eventType: 'undeleteMessage', tenant: { id: TENANT_ID } },
  };
}

/** `messageReaction`: `replyToId` is the message reacted to. */
export function reactionActivity(
  input: BaseInput & { targetMessageId: string; added?: string[]; removed?: string[] },
): Partial<Activity> {
  return {
    ...baseActivity(input),
    type: 'messageReaction',
    replyToId: input.targetMessageId,
    ...(input.added ? { reactionsAdded: input.added.map((type) => ({ type })) } : {}),
    ...(input.removed ? { reactionsRemoved: input.removed.map((type) => ({ type })) } : {}),
  };
}

/** An Action.Submit click: a message activity without text, `value` is the
 *  card's `data`, `replyToId` the card message. */
export function cardSubmitActivity(
  input: BaseInput & { value: unknown; cardMessageId: string },
): Partial<Activity> {
  return {
    ...baseActivity(input),
    type: 'message',
    replyToId: input.cardMessageId,
    value: input.value,
    entities: [{ type: 'clientInfo', locale: 'de-DE', country: 'DE', platform: 'Web', timezone: 'Europe/Berlin' }],
  };
}

/** An `Action.Execute` click: an `adaptiveCard/action` invoke whose
 *  `value.action` carries verb and data; `replyToId` is the card message. */
export function cardExecuteActivity(
  input: BaseInput & { verb: string; data: unknown; cardMessageId: string; anonymous?: boolean },
): Partial<Activity> {
  const base = baseActivity(input);
  return {
    ...base,
    // An anonymous meeting participant has no Entra object id.
    ...(input.anonymous ? { from: { id: `8:teamsvisitor:${input.from.aadObjectId.slice(0, 8)}`, name: `${input.from.name} (Gast)` } } : {}),
    type: 'invoke',
    name: 'adaptiveCard/action',
    replyToId: input.cardMessageId,
    value: { action: { type: 'Action.Execute', verb: input.verb, data: input.data }, trigger: 'manual' },
  };
}

/** `conversationUpdate` for a person added to the chat. */
export function membersAddedActivity(input: BaseInput & { added: SimUser[] }): Partial<Activity> {
  return {
    ...baseActivity(input),
    type: 'conversationUpdate',
    membersAdded: input.added.map((member) => ({ id: member.id, aadObjectId: member.aadObjectId, name: member.name })),
    channelData: { eventType: 'teamMemberAdded', tenant: { id: TENANT_ID } },
  };
}

/** The kernel side of the subscription: which bot moderates which chat. */
export class SimulatedKernel {
  readonly events: ConversationMessageEvent[] = [];
  private readonly subscriptions = new Set<string>();

  subscribe(conversationId: string, bot: SimBot): void {
    this.subscriptions.add(`${conversationId}|${bot.key}`);
  }

  unsubscribe(conversationId: string, bot: SimBot): void {
    this.subscriptions.delete(`${conversationId}|${bot.key}`);
  }

  isSubscribed(conversationId: string, channelKey: string): boolean {
    return this.subscriptions.has(`${conversationId}|${channelKey}`);
  }

  eventsFor(bot: SimBot): ConversationMessageEvent[] {
    return this.events.filter((event) => event.receivedByChannelKey === bot.key);
  }
}

export interface SimulatedBot {
  bot: TeamsBot;
  /** One entry per turn the agent was asked to answer. */
  turns: string[];
  /** Everything the bot sent back into the conversation. */
  sent: unknown[];
  /** Deliver one activity as the Bot Framework adapter would. */
  deliver(activity: Partial<Activity>): Promise<void>;
}

/**
 * A real `TeamsBot` (every bot of the deployment shares one instance, as in
 * production) wired to the simulated kernel. Calling it again with the same
 * kernel models a middleware restart: new process state, same subscriptions.
 */
export function simulatedBot(
  kernel: SimulatedKernel,
  group?: {
    captureConversationReference: (context: TurnContext) => void;
    emitMembershipEvent: (event: Record<string, unknown>) => void;
    /** What the agent answers; '' models the kernel's empty safety-net answer. */
    answerText?: string;
  },
): SimulatedBot {
  const turns: string[] = [];
  const sent: unknown[] = [];
  const history = new Map<string, unknown[]>();
  const bot = new (TeamsBot as unknown as new (...args: unknown[]) => TeamsBot)(
    {
      chat: async (input: { userMessage: string }) => {
        turns.push(input.userMessage);
        return { text: group?.answerText ?? 'ok' };
      },
    },
    {
      get: (scope: string) => history.get(scope) ?? [],
      append: () => {},
      resetTurns: () => {},
      markPending: () => {},
      getPending: () => undefined,
      clearPending: () => {},
      size: () => 0,
      clear: () => {},
    },
    undefined,
    {
      run: async (_value: unknown, fn: () => Promise<unknown>) => fn(),
      enter: () => {},
      runWithChatParticipants: async (_provider: unknown, fn: () => Promise<unknown>) => fn(),
      current: () => undefined,
      currentTurnId: () => undefined,
      currentTurnDate: () => '2026-11-18',
    },
    // attachmentStore … resolveConductorAwait: not part of these scenarios.
    ...Array.from({ length: 12 }, () => undefined),
    group,
  );
  bot.attachConversationMessages(
    new TeamsConversationMessages({
      isAllMessagesSubscribed: (conversationId, channelKey) => kernel.isSubscribed(conversationId, channelKey),
      emit: (event) => {
        kernel.events.push(event);
      },
    }),
  );
  const deliver = async (activity: Partial<Activity>): Promise<void> => {
    const context = {
      activity,
      sendActivity: async (outgoing: unknown) => {
        sent.push(outgoing);
        return { id: newMessageId() };
      },
      sendActivities: async (outgoing: unknown[]) => {
        sent.push(...outgoing);
        return outgoing.map(() => ({ id: newMessageId() }));
      },
      updateActivity: async () => undefined,
      deleteActivity: async () => undefined,
      turnState: new Map<unknown, unknown>(),
    } as unknown as TurnContext;
    await bot.run(context);
  };
  return { bot, turns, sent, deliver };
}

// ---------------------------------------------------------------------------
// Simulated Microsoft Graph + Bot Connector
// ---------------------------------------------------------------------------

export const GRAPH_RESOURCE_APP_ID = '00000003-0000-0000-c000-000000000000';

/** A Graph error body as the service returns it. */
export function graphError(code: string, message: string): Record<string, unknown> {
  return {
    error: {
      code,
      message,
      innerError: { date: '2026-11-18T09:15:00', 'request-id': 'b0c9a2f4-0000-4000-8000-000000000001' },
    },
  };
}

interface SimInstallation {
  bot: SimBot;
  id: string;
  version: string;
  /** RSC the installed app definition declares. */
  declared: string[];
  /** RSC consented for this install (what the grants show). */
  consented: string[];
}

/** The RSC set of this repo's app package template. */
export const TEMPLATE_RSC = [
  'ChannelMessage.Read.Group',
  'ChatMessage.Read.Chat',
  'TeamsActivity.Send.Group',
  'TeamsActivity.Send.Chat',
  'FileStorageContainer.Selected.Group',
  'ChatMember.Read.Chat',
  'ChannelMember.Read.Group',
];

/** What a chat install consents to today (live finding, byte5 tenant). */
export const CHAT_RSC = ['ChatMessage.Read.Chat', 'ChatMember.Read.Chat', 'TeamsActivity.Send.Chat'];

/**
 * `fetch` for `TeamsGraphResolver`: the Entra token endpoint and, per chat or
 * team, `installedApps` (list, single row, `upgrade`) and `permissionGrants`.
 * Row shapes follow the responses read from the byte5 tenant on 2026-11-18:
 * one grant row per permission with `clientAppId` = the Entra application id
 * and `clientId` = the service principal object id; an installation row with
 * `teamsAppDefinition.bot.id` and, when selected, `consentedPermissionSet`.
 */
export class SimulatedGraph {
  readonly requests: string[] = [];
  readonly upgrades: { scope: string; installationId: string; permissions: string[] }[] = [];
  private readonly installs = new Map<string, SimInstallation[]>();
  private readonly failures = new Map<string, { status: number; body: Record<string, unknown> }>();

  /** Install `bot`'s app in a chat (`scope` = chat id) or a team (team id). */
  install(scope: string, bot: SimBot, opts: { consented?: string[]; declared?: string[]; version?: string } = {}): void {
    const list = (this.installs.get(scope) ?? []).filter((entry) => entry.bot.appId !== bot.appId);
    list.push({
      bot,
      id: `inst-${bot.appId.slice(0, 8)}-${Buffer.from(scope).toString('base64url').slice(0, 12)}`,
      version: opts.version ?? '1.3.0',
      declared: opts.declared ?? TEMPLATE_RSC,
      consented: opts.consented ?? CHAT_RSC,
    });
    this.installs.set(scope, list);
  }

  /** Make `installedApps` or `permissionGrants` of a scope fail. */
  fail(scope: string, resource: 'installedApps' | 'permissionGrants', status: number, code: string, message: string): void {
    this.failures.set(`${scope}|${resource}`, { status, body: graphError(code, message) });
  }

  private installationRow(entry: SimInstallation, withDefinition: boolean): Record<string, unknown> {
    return {
      id: entry.id,
      consentedPermissionSet: {
        resourceSpecificPermissions: entry.consented.map((permissionValue) => ({ permissionValue, permissionType: 'application' })),
      },
      teamsApp: {
        id: `catalog-${entry.bot.appId.slice(0, 8)}`,
        displayName: entry.bot.name,
        externalId: `manifest-${entry.bot.appId.slice(0, 8)}`,
        distributionMethod: 'organization',
      },
      teamsAppDefinition: {
        version: entry.version,
        bot: { id: entry.bot.appId },
        ...(withDefinition
          ? {
              authorization: {
                requiredPermissionSet: {
                  resourceSpecificPermissions: entry.declared.map((permissionValue) => ({ permissionValue, permissionType: 'application' })),
                },
              },
            }
          : {}),
      },
    };
  }

  readonly fetch: typeof fetch = async (input, init) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.toString() : input.url;
    const method = init?.method ?? 'GET';
    this.requests.push(`${method} ${url}`);
    if (url.startsWith('https://login.microsoftonline.com/')) {
      return json(200, { token_type: 'Bearer', expires_in: 3599, access_token: 'simulated-token' });
    }
    const match = /^https:\/\/graph\.microsoft\.com\/v1\.0\/(chats|teams)\/([^/]+)\/(installedApps|permissionGrants)(?:\/([^/?]+))?(\/upgrade)?(?:\?.*)?$/.exec(url);
    if (!match) return json(404, graphError('NotFound', 'NotFound'));
    const scope = decodeURIComponent(match[2] ?? '');
    const resource = match[3] as 'installedApps' | 'permissionGrants';
    const failure = this.failures.get(`${scope}|${resource}`);
    if (failure) return json(failure.status, failure.body);
    const installs = this.installs.get(scope);
    if (!installs) return json(404, graphError('NotFound', 'NotFound'));

    if (resource === 'permissionGrants') {
      return json(200, {
        '@odata.context': `https://graph.microsoft.com/v1.0/$metadata#${match[1] ?? ''}('${match[2] ?? ''}')/permissionGrants`,
        value: installs.flatMap((entry) =>
          entry.consented.map((permission) => ({
            id: `grant-${entry.bot.appId.slice(0, 8)}-${permission}`,
            clientAppId: entry.bot.appId,
            clientId: `sp-object-${entry.bot.appId.slice(0, 8)}`,
            resourceAppId: GRAPH_RESOURCE_APP_ID,
            permissionType: 'Application',
            permission,
          })),
        ),
      });
    }
    const installationId = match[4] ? decodeURIComponent(match[4]) : undefined;
    if (!installationId) {
      return json(200, { value: installs.map((entry) => this.installationRow(entry, false)) });
    }
    const entry = installs.find((candidate) => candidate.id === installationId);
    if (!entry) return json(404, graphError('NotFound', 'NotFound'));
    if (!match[5]) return json(200, this.installationRow(entry, true));
    if (method !== 'POST') return json(405, graphError('MethodNotAllowed', 'MethodNotAllowed'));
    const body = JSON.parse(String(init?.body ?? '{}')) as {
      consentedPermissionSet?: { resourceSpecificPermissions?: { permissionValue: string }[] };
    };
    const requested = (body.consentedPermissionSet?.resourceSpecificPermissions ?? []).map((p) => p.permissionValue);
    if ([...requested].sort().join() !== [...entry.declared].sort().join()) {
      return json(400, graphError('BadRequest', 'ResourceSpecificPermissionsMismatch: the consented permission set does not match the app definition.'));
    }
    entry.consented = [...entry.declared];
    this.upgrades.push({ scope, installationId, permissions: requested });
    return new Response(null, { status: 204 });
  };
}

function json(status: number, body: unknown): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

/** An error as the Bot Connector client raises it: `statusCode` plus the
 *  service's `{ error: { code, message } }` body. */
export function connectorError(statusCode: number, code: string, message: string): Error {
  return Object.assign(new Error(message), {
    statusCode,
    code,
    body: { error: { code, message } },
    details: { error: { code, message } },
  });
}

/** The conversation reference Bot Framework derives from an inbound activity. */
export function referenceFor(bot: SimBot, conversationId: string, from: SimUser = ANNA): Record<string, unknown> {
  return {
    activityId: newMessageId(),
    user: { id: from.id, name: from.name, aadObjectId: from.aadObjectId },
    bot: { id: bot.key, name: bot.name },
    conversation: { isGroup: true, conversationType: 'groupChat', tenantId: TENANT_ID, id: conversationId },
    channelId: 'msteams',
    locale: 'de-DE',
    serviceUrl: SERVICE_URL,
  };
}

/** Teams' answer when the user blocked the bot or removed its app: 403 with
 *  `subCode: MessageWritesBlocked` in the inner error. */
export function messageWritesBlockedError(): Error {
  const body = {
    error: {
      code: 'Forbidden',
      message: '',
      innerHttpError: {
        statusCode: 403,
        body: { subCode: 'MessageWritesBlocked', details: 'Thread is blocked from message writes.', errorCode: 209, message: '' },
      },
    },
  };
  return Object.assign(new Error('Forbidden'), { statusCode: 403, code: 'Forbidden', body, details: body });
}

/** Teams' answer to a first 1:1 message when the app is not installed in the
 *  user's personal scope: 403 `ForbiddenOperationException`. */
export function botNotInstalledError(): Error {
  return connectorError(403, 'ForbiddenOperationException', "Bot is not installed in user's personal scope");
}

export interface SentActivity {
  /** `28:<appId>` of the bot the activity went out as. */
  asBot: string;
  conversationId: string;
  activity: Record<string, unknown>;
  id: string;
}

/**
 * The outbound half of the Bot Connector: every proactive post and every
 * opened 1:1, with the bot it went out as. `failDirect` makes the next 1:1
 * to a user fail the way Teams does.
 */
export class SimulatedConnector {
  readonly sent: SentActivity[] = [];
  readonly opened: { asBot: string; userId: string; serviceUrl?: string }[] = [];
  private readonly directFailures = new Map<string, Error>();

  /** `times` limits the failure to that many attempts (a transient error). */
  failDirect(bot: SimBot, user: SimUser, err: Error, times = Number.POSITIVE_INFINITY): void {
    this.directFailures.set(`${bot.appId}|${user.aadObjectId}`, err);
    this.remainingFailures.set(`${bot.appId}|${user.aadObjectId}`, times);
  }

  private readonly remainingFailures = new Map<string, number>();

  private takeFailure(key: string): Error | undefined {
    const failure = this.directFailures.get(key);
    const remaining = this.remainingFailures.get(key) ?? 0;
    if (!failure || remaining <= 0) return undefined;
    this.remainingFailures.set(key, remaining - 1);
    return failure;
  }

  private context(asBot: string, conversationId: string): TurnContext {
    return {
      activity: { conversation: { id: conversationId }, recipient: { id: asBot }, serviceUrl: SERVICE_URL },
      sendActivity: async (activity: Record<string, unknown>) => {
        const id = newMessageId();
        this.sent.push({ asBot, conversationId, activity, id });
        return { id };
      },
    } as unknown as TurnContext;
  }

  readonly sendProactive = async (
    reference: { bot?: { id?: string }; conversation?: { id?: string } },
    build: (turnContext: TurnContext) => Promise<void>,
    opts?: { botAppId?: string },
  ): Promise<void> => {
    const asBot = opts?.botAppId ? `28:${opts.botAppId}` : (reference.bot?.id ?? '28:default');
    const key = [...this.directFailures.keys()].find(
      (candidate) => reference.conversation?.id === directConversationId(candidate),
    );
    const failure = key ? this.takeFailure(key) : undefined;
    if (failure) throw failure;
    await build(this.context(asBot, reference.conversation?.id ?? ''));
  };

  readonly openDirectConversation = async (
    botAppId: string,
    target: { userId: string; serviceUrl?: string },
    build: (turnContext: TurnContext) => Promise<void>,
  ): Promise<void> => {
    this.opened.push({ asBot: `28:${botAppId}`, userId: target.userId, ...(target.serviceUrl ? { serviceUrl: target.serviceUrl } : {}) });
    const failure = this.takeFailure(`${botAppId}|${target.userId}`);
    if (failure) throw failure;
    await build(this.context(`28:${botAppId}`, directConversationId(`${botAppId}|${target.userId}`)));
  };
}

/** Teams 1:1 conversation ids are opaque `a:…` strings, stable per bot+user. */
export function directConversationId(botAndUser: string): string {
  return `a:1${Buffer.from(botAndUser).toString('base64url').slice(0, 40)}`;
}

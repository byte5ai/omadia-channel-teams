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
import type { TeamsMeetingSessions } from '@omadia/channel-teams';
import type { ConversationMessageEvent } from '../../src/kernel-types.js';

export const TENANT_ID = '72f988bf-86f1-41af-91ab-2d7cd011db47';
export const SERVICE_URL = 'https://smba.trafficmanager.net/emea/';
export const GROUP_CHAT_ID = '19:b8a1c0f2d3e44f5a9b6c7d8e9f0a1b2c@thread.v2';
export const OTHER_GROUP_CHAT_ID = '19:0f9e8d7c6b5a4f3e2d1c0b9a8f7e6d5c@thread.v2';
/** A team channel as Teams names it, and the AAD group id of its team. In
 *  the "General" channel `team.id`, `teamsTeamId` and `channel.id` are all
 *  this thread id (live, 2026-11-18). */
export const CHANNEL_ID = '19:fdsWodnEDi7J3FI5tgPf1_4KxrURUShMFpNJrmcN8ns1@thread.tacv2';
export const TEAM_AAD_GROUP_ID = 'c6031502-025c-4297-92c6-b734e17e46c5';

/**
 * The conversation id of a channel message: the channel id plus the ROOT
 * message of its thread. Live: every root post is its own conversation id,
 * and `messageid` equals the activity id of the root post. A reply in the
 * thread is documented to arrive under the same id; that was not captured.
 */
export function channelThreadId(rootMessageId: string, channelId: string = CHANNEL_ID): string {
  return `${channelId};messageid=${rootMessageId}`;
}

/**
 * `channelData` as captured live (keys only where the content was not read):
 *   channel message      teamsChannelId, teamsTeamId, channel, team, tenant, replyChain
 *   chat message         tenant, chat, source
 *   meeting chat message tenant, chat, source, meeting
 * `team.aadGroupId` IS present on channel messages; `team.name` is not.
 * The contents of `replyChain`, `chat` and `source` were not captured, so
 * they are empty objects here: present as keys, with no invented values.
 */
function channelDataFor(conversationId: string, conversationType: string): Record<string, unknown> {
  const tenant = { id: TENANT_ID };
  if (conversationType === 'channel') {
    const channelId = conversationId.split(';messageid=')[0] ?? conversationId;
    return {
      teamsChannelId: channelId,
      teamsTeamId: CHANNEL_ID,
      channel: { id: channelId },
      team: { id: CHANNEL_ID, aadGroupId: TEAM_AAD_GROUP_ID },
      tenant,
      replyChain: {},
    };
  }
  if (conversationType === 'personal') return { tenant };
  return {
    tenant,
    chat: {},
    source: {},
    ...(conversationId.startsWith('19:meeting_') ? { meeting: { id: MEETING_ID } } : {}),
  };
}

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
    channelData: channelDataFor(input.conversationId ?? GROUP_CHAT_ID, conversationType),
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
    // Live: a mention message carries its HTML once more as an attachment.
    ...(input.mention
      ? {
          attachments: [
            {
              contentType: 'text/html',
              content: `<p><span itemtype="http://schema.skype.com/Mention" itemscope="" itemid="0">${input.bot.name}</span>&nbsp;${input.text}</p>`,
            },
          ],
        }
      : {}),
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

export const MEETING_ID = 'MCMxOTptZWV0aW5nX01qZGhOak00WXpVdFl6RXhaaTAwT1RGa0xXRTVZamN0TlRWa1lUUTBaR1ZsTTJKaEB0aHJlYWQudjIjMA==';

/**
 * Meeting start / end as Microsoft Learn shows them ("meeting-apps-apis"):
 * an `event` activity whose `value` carries `Id`, `JoinUrl`, `Title`,
 * `MeetingType`, `StartTime`, `EndTime` in that casing. The page's samples
 * print the name with a LEADING SPACE; `paddedName` reproduces that.
 */
export function meetingLifecycleActivity(
  input: BaseInput & { phase: 'start' | 'end'; meetingType?: string; at?: string; paddedName?: boolean; meetingId?: string },
): Partial<Activity> {
  const at = input.at ?? '2026-11-18T10:00:00.000Z';
  const name = input.phase === 'start' ? 'application/vnd.microsoft.meetingStart' : 'application/vnd.microsoft.meetingEnd';
  return {
    ...baseActivity({ ...input, conversationId: input.conversationId ?? MEETING_CHAT_ID, timestamp: at }),
    type: 'event',
    name: input.paddedName ? ` ${name}` : name,
    value: {
      Id: 'value-id-not-to-be-used',
      JoinUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_x%40thread.v2/0',
      Title: 'Budgetrunde',
      MeetingType: input.meetingType ?? 'Scheduled',
      StartTime: input.phase === 'start' ? at : null,
      EndTime: input.phase === 'end' ? at : null,
    },
    channelData: { tenant: { id: TENANT_ID }, meeting: { id: input.meetingId ?? MEETING_ID } },
  };
}

/** Participant join / leave as the same page shows them: `value.members[]`
 *  with `user` (`id`, `aadObjectId`, `name`, `tenantId`, `objectId`) and
 *  `meeting` (`inMeeting`, on join also `role`). */
export function meetingParticipantActivity(
  input: BaseInput & { phase: 'join' | 'leave'; people: SimUser[]; at?: string; meetingId?: string },
): Partial<Activity> {
  return {
    ...baseActivity({ ...input, conversationId: input.conversationId ?? MEETING_CHAT_ID, timestamp: input.at ?? '2026-11-18T10:05:00.000Z' }),
    type: 'event',
    name:
      input.phase === 'join'
        ? 'application/vnd.microsoft.meetingParticipantJoin'
        : 'application/vnd.microsoft.meetingParticipantLeave',
    value: {
      members: input.people.map((person) => ({
        user: {
          tenantId: TENANT_ID,
          objectId: person.aadObjectId,
          id: `${person.id} `,
          name: person.name,
          aadObjectId: ` ${person.aadObjectId} `,
        },
        meeting: input.phase === 'join' ? { inMeeting: true, role: 'Attendee' } : { inMeeting: false },
      })),
    },
    channelData: { tenant: { id: TENANT_ID }, meeting: { id: input.meetingId ?? MEETING_ID } },
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
  private readonly subscriptions = new Map<string, 'all' | 'addressed'>();
  /** `false` models a kernel before channel-sdk 0.3.5: it offers
   *  `isAllMessagesSubscribed` and no `conversationDelivery`. */
  offersDeliveryModes = true;

  /** `mode`: what the kernel's `conversationDelivery` answers for this bot
   *  in this conversation (#2089). `addressed` is a process in follow-up. */
  subscribe(conversationId: string, bot: SimBot, mode: 'all' | 'addressed' = 'all'): void {
    this.subscriptions.set(`${conversationId}|${bot.key}`, mode);
  }

  unsubscribe(conversationId: string, bot: SimBot): void {
    this.subscriptions.delete(`${conversationId}|${bot.key}`);
  }

  /** The kernel's `isAllMessagesSubscribed`: `true` exactly for `all`. */
  isSubscribed(conversationId: string, channelKey: string): boolean {
    return this.subscriptions.get(`${conversationId}|${channelKey}`) === 'all';
  }

  deliveryMode(conversationId: string, channelKey: string): 'all' | 'addressed' | 'none' {
    return this.subscriptions.get(`${conversationId}|${channelKey}`) ?? 'none';
  }

  eventsFor(bot: SimBot): ConversationMessageEvent[] {
    return this.events.filter((event) => event.receivedByChannelKey === bot.key);
  }
}

/** What `channelResolver.resolve` is asked for one key (#2055). */
export interface SimulatedResolveInput {
  channelType: string;
  channelKey: string;
  conversationId: string;
  principal?: string;
  locale?: string;
}

export interface SimulatedBot {
  bot: TeamsBot;
  /** One entry per turn the agent was asked to answer. */
  turns: string[];
  /** The `origin` each turn was handed to the agent with, in turn order. */
  origins: unknown[];
  /** The session scope of each turn, in turn order. */
  scopes: string[];
  /** Everything the bot sent back into the conversation. */
  sent: unknown[];
  /** What the conversation history was given, in order. */
  appended: { userMessage: string; assistantAnswer: string }[];
  /** What the bot asked the management layer, one entry per probed key. */
  resolveCalls: SimulatedResolveInput[];
  /** The agent the simulated management layer hands the text to. */
  agent: { chat(input: { userMessage: string; origin?: unknown; sessionScope: string }): Promise<Record<string, unknown>> };
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
    /** The whole answer object, for what rides along with an emptied text. */
    answer?: Record<string, unknown>;
    meetingSessions?: TeamsMeetingSessions;
    /** Every bot is an agent's OWN provisioned identity, as every
     *  facilitator's is: the session scope is then bot-qualified. */
    ownIdentity?: boolean;
    /** The management layer's decision per resolved key (#2055). Wins over
     *  `ownIdentity`. Receives what the bot asked: key, principal, locale. */
    resolve?: (input: SimulatedResolveInput) => Record<string, unknown>;
    /** Earlier turns of the conversation (what the history store holds). */
    priorTurns?: { userMessage: string; assistantAnswer: string }[];
    /** A topic detector; it is consulted only when there are prior turns. */
    topicDetector?: { classify(input: { userMessage: string; history: unknown[] }): Promise<Record<string, unknown>> };
    /** Verified e-mail address per Entra object id (Graph `getUserMail`). */
    emailByAad?: Record<string, string>;
    /** The agent the simulated management layer hands the text to; default: answers `answer`. */
    chat?: (input: { userMessage: string; origin?: unknown; sessionScope: string }) => Promise<Record<string, unknown>>;
  },
): SimulatedBot {
  const turns: string[] = [];
  const origins: unknown[] = [];
  const scopes: string[] = [];
  const sent: unknown[] = [];
  const history = new Map<string, unknown[]>();
  const appended: { userMessage: string; assistantAnswer: string }[] = [];
  const resolveCalls: SimulatedResolveInput[] = [];
  const agent = {
    chat: async (input: { userMessage: string; origin?: unknown; sessionScope: string }) => {
      turns.push(input.userMessage);
      origins.push(input.origin);
      scopes.push(input.sessionScope);
      if (group?.chat) return group.chat(input);
      return group?.answer ?? { text: group?.answerText ?? 'ok' };
    },
  };
  const bot = new (TeamsBot as unknown as new (...args: unknown[]) => TeamsBot)(
    agent,
    {
      get: (scope: string) => history.get(scope) ?? group?.priorTurns ?? [],
      append: (_scope: string, turn: { userMessage: string; assistantAnswer: string }) => {
        appended.push(turn);
      },
      resetTurns: () => {},
      markPending: () => {},
      getPending: () => undefined,
      clearPending: () => {},
      size: () => 0,
      clear: () => {},
    },
    group?.topicDetector,
    {
      run: async (_value: unknown, fn: () => Promise<unknown>) => fn(),
      enter: () => {},
      runWithChatParticipants: async (_provider: unknown, fn: () => Promise<unknown>) => fn(),
      current: () => undefined,
      currentTurnId: () => undefined,
      currentTurnDate: () => '2026-11-18',
    },
    // attachmentStore … buildRoutineListSmartCardAttachment: not part of these scenarios.
    ...Array.from({ length: 7 }, () => undefined),
    // resolveChatAgentForActivity: the kernel's answer for a bot that IS one
    // agent (`exclusive`), which is what makes the session scope bot-qualified.
    group?.resolve
      ? (input: SimulatedResolveInput) => {
          resolveCalls.push(input);
          return group.resolve?.(input);
        }
      : group?.ownIdentity
        ? (input: SimulatedResolveInput) => {
            resolveCalls.push(input);
            return input.channelKey.startsWith('28:')
              ? { decision: 'bound', chatAgent: agent, exclusive: true }
              : { decision: 'reject', chatAgent: agent };
          }
        : undefined,
    // conversationObserver, emitConductorEvent: not part of these scenarios.
    undefined,
    undefined,
    // resolveEmailByAad: Graph's verified address for an Entra object id.
    group?.emailByAad
      ? async (aadObjectId: string) => group.emailByAad?.[aadObjectId] ?? null
      : undefined,
    // resolveConductorAwait: not part of these scenarios.
    undefined,
    group,
  );
  bot.attachConversationMessages(
    new TeamsConversationMessages({
      isAllMessagesSubscribed: (conversationId, channelKey) => kernel.isSubscribed(conversationId, channelKey),
      ...(kernel.offersDeliveryModes
        ? { conversationDelivery: (conversationId: string, channelKey: string) => kernel.deliveryMode(conversationId, channelKey) }
        : {}),
      emit: (event) => {
        kernel.events.push(event);
      },
    }),
  );
  if (group?.meetingSessions) bot.attachMeetingSessions(group.meetingSessions);
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
  return { bot, turns, origins, scopes, sent, resolveCalls, appended, agent, deliver };
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
  /** RSC consented for this install (what the grants show). */
  consented: string[];
}

/** The RSC set of the app package template before the meeting permissions
 *  (revisions 0 and 1). It still carried `FileStorageContainer.Selected.Group`,
 *  which no team accepts at install. */
export const LEGACY_TEMPLATE_RSC = [
  'ChannelMessage.Read.Group',
  'ChatMessage.Read.Chat',
  'TeamsActivity.Send.Group',
  'TeamsActivity.Send.Chat',
  'FileStorageContainer.Selected.Group',
  'ChatMember.Read.Chat',
  'ChannelMember.Read.Group',
];

const MEETING_RSC = [
  'OnlineMeeting.ReadBasic.Chat',
  'OnlineMeetingParticipant.Read.Chat',
  'OnlineMeetingNotification.Send.Chat',
  'ChannelMeeting.ReadBasic.Group',
];

/** The template of 0.36.0 / 0.36.1 (TEMPLATE_REVISION 2): eleven permissions. */
export const REVISION_2_TEMPLATE_RSC = [...LEGACY_TEMPLATE_RSC, ...MEETING_RSC];

/** The RSC set of this repo's app package template (TEMPLATE_REVISION 3):
 *  revision 2 without `FileStorageContainer.Selected.Group`. */
export const TEMPLATE_RSC = REVISION_2_TEMPLATE_RSC.filter(
  (permission) => permission !== 'FileStorageContainer.Selected.Group',
);

/** What a chat install from the earlier package shows (live, byte5 tenant):
 *  the chat-scoped permissions only. */
export const LEGACY_CHAT_RSC = ['ChatMessage.Read.Chat', 'ChatMember.Read.Chat', 'TeamsActivity.Send.Chat'];

/** The chat-scoped permissions of the current template. */
export const CHAT_RSC = TEMPLATE_RSC.filter((permission) => permission.endsWith('.Chat'));

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
  /** What each bot's app declares in the catalog right now. */
  private readonly published = new Map<string, { declared: string[]; version: string }>();

  /** Publish a (new) app definition for `bot` to the tenant catalog. */
  publish(bot: SimBot, declared: string[], version: string): void {
    this.published.set(bot.appId, { declared, version });
  }

  private catalogOf(bot: SimBot): { declared: string[]; version: string } {
    return this.published.get(bot.appId) ?? { declared: TEMPLATE_RSC, version: '1.3.0' };
  }

  private readonly failures = new Map<string, { status: number; body: Record<string, unknown> }>();

  /** Install `bot`'s app in a chat (`scope` = chat id) or a team (team id). */
  install(scope: string, bot: SimBot, opts: { consented?: string[]; version?: string } = {}): void {
    const list = (this.installs.get(scope) ?? []).filter((entry) => entry.bot.appId !== bot.appId);
    list.push({
      bot,
      id: `inst-${bot.appId.slice(0, 8)}-${Buffer.from(scope).toString('base64url').slice(0, 12)}`,
      version: opts.version ?? this.catalogOf(bot).version,
      consented: opts.consented ?? CHAT_RSC,
    });
    this.installs.set(scope, list);
  }

  /** Make `installedApps` or `permissionGrants` of a scope fail. */
  fail(
    scope: string,
    resource: 'installedApps' | 'permissionGrants' | 'upgrade' | 'catalog',
    status: number,
    code: string,
    message: string,
  ): void {
    this.failures.set(`${scope}|${resource}`, { status, body: graphError(code, message) });
  }

  private installationRow(entry: SimInstallation): Record<string, unknown> {
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
      teamsAppDefinition: { version: entry.version, bot: { id: entry.bot.appId } },
    };
  }

  readonly fetch: typeof fetch = async (input, init) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.toString() : input.url;
    const method = init?.method ?? 'GET';
    this.requests.push(`${method} ${url}`);
    if (url.startsWith('https://login.microsoftonline.com/')) {
      return json(200, { token_type: 'Bearer', expires_in: 3599, access_token: 'simulated-token' });
    }
    if (url.startsWith('https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?')) {
      const catalogFailure = [...this.failures.entries()].find(([key]) => key.endsWith('|catalog'))?.[1];
      if (catalogFailure) return json(catalogFailure.status, catalogFailure.body);
      const wanted = /id eq '([^']+)'/.exec(decodeURIComponent(url))?.[1];
      const bots = [...this.installs.values()].flat().map((entry) => entry.bot);
      const bot = bots.find((candidate) => `catalog-${candidate.appId.slice(0, 8)}` === wanted);
      if (!bot) return json(200, { value: [] });
      const catalog = this.catalogOf(bot);
      return json(200, {
        value: [
          {
            id: wanted,
            appDefinitions: [
              {
                id: `def-${bot.appId.slice(0, 8)}-${catalog.version}`,
                version: catalog.version,
                authorization: {
                  requiredPermissionSet: {
                    resourceSpecificPermissions: catalog.declared.map((permissionValue) => ({ permissionValue, permissionType: 'Application' })),
                  },
                },
              },
            ],
          },
        ],
      });
    }
    const match = /^https:\/\/graph\.microsoft\.com\/v1\.0\/(chats|teams)\/([^/]+)\/(installedApps|permissionGrants)(?:\/([^/?]+))?(\/upgrade)?(?:\?.*)?$/.exec(url);
    if (!match) return json(404, graphError('NotFound', 'NotFound'));
    const scope = decodeURIComponent(match[2] ?? '');
    const resource = match[3] as 'installedApps' | 'permissionGrants';
    const failure = this.failures.get(`${scope}|${match[5] ? 'upgrade' : resource}`);
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
      return json(200, { value: installs.map((entry) => this.installationRow(entry)) });
    }
    const entry = installs.find((candidate) => candidate.id === installationId);
    if (!entry) return json(404, graphError('NotFound', 'NotFound'));
    if (!match[5]) return json(200, this.installationRow(entry));
    if (method !== 'POST') return json(405, graphError('MethodNotAllowed', 'MethodNotAllowed'));
    const body = JSON.parse(String(init?.body ?? '{}')) as {
      consentedPermissionSet?: { resourceSpecificPermissions?: { permissionValue: string }[] };
    };
    const requested = (body.consentedPermissionSet?.resourceSpecificPermissions ?? []).map((p) => p.permissionValue);
    // The upgrade moves the install to the published definition; the
    // consented set must equal what that definition declares.
    const catalog = this.catalogOf(entry.bot);
    if ([...requested].sort().join() !== [...catalog.declared].sort().join()) {
      return json(400, graphError('BadRequest', 'ResourceSpecificPermissionsMismatch: the consented permission set does not match the app definition.'));
    }
    const applicable = match[1] === 'chats' ? '.Chat' : '.Group';
    entry.consented = catalog.declared.filter((permission) => permission.endsWith(applicable));
    entry.version = catalog.version;
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

/**
 * LIVE (byte5 tenant, 2026-11-18): the answer to a SEND into a 1:1 the user
 * blocked. Creating the conversation still answers 201 while blocked.
 */
export function conversationBlockedError(): Error {
  return connectorError(403, 'ConversationBlockedByUser', 'User blocked the conversation with the bot.');
}

/**
 * LIVE (byte5 tenant, 2026-11-18): the answer to create-conversation when
 * the app is not installed for the user. The code is `BadArgument`, not the
 * `ForbiddenOperationException` the documentation names.
 */
export function botNotInstalledError(): Error {
  return connectorError(403, 'BadArgument', "Bot is not installed in user's personal scope");
}

/** DOCUMENTED shape for a blocked or uninstalled bot: 403 with
 *  `subCode: MessageWritesBlocked` in the inner error. Not seen live. */
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

/** DOCUMENTED shape for "not installed": 403 ForbiddenOperationException.
 *  Not seen live. */
export function documentedNotInstalledError(): Error {
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
  /** Every `PUT …/activities/{id}`: a posted message replaced. */
  readonly updated: SentActivity[] = [];
  /** Make the next message update fail (deleted or too old a message). */
  updateFailure: Error | undefined;
  readonly opened: { asBot: string; userId: string; serviceUrl?: string }[] = [];
  private readonly directFailures = new Map<string, Error>();

  /** The user blocked the bot: creating the 1:1 still works, the SEND fails. */
  block(bot: SimBot, user: SimUser, err: Error = conversationBlockedError()): void {
    this.sendFailures.set(directConversationId(`${bot.appId}|${user.aadObjectId}`), err);
  }

  unblock(bot: SimBot, user: SimUser): void {
    this.sendFailures.delete(directConversationId(`${bot.appId}|${user.aadObjectId}`));
  }

  private readonly sendFailures = new Map<string, Error>();

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
        const blocked = this.sendFailures.get(conversationId);
        if (blocked) throw blocked;
        const id = newMessageId();
        this.sent.push({ asBot, conversationId, activity, id });
        return { id };
      },
      updateActivity: async (activity: Record<string, unknown>) => {
        if (this.updateFailure) throw this.updateFailure;
        this.updated.push({ asBot, conversationId, activity, id: String(activity['id'] ?? '') });
        return { id: activity['id'] };
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

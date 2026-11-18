// Readiness check for native facilitators (byte5ai/omadia#2043 §6.2, issue
// #43): can THIS bot moderate THIS conversation?
//
//   installed           the bot's Teams app is in the chat
//   canReadAllMessages  RSC `ChatMessage.Read.Chat` is granted to the app in
//                       this chat (requesting it in the manifest is not
//                       enough; the grant exists per chat)
//   canPost             the bot can reach the chat through the Bot Connector
//
// Two sources answer:
//   - Graph `GET /chats/{id}/permissionGrants` lists the RSC grants of the
//     chat. A grant for the bot's app proves installation and names the
//     granted permissions.
//   - A Bot Connector roster read (`TeamsInfo.getPagedMembers`, one member)
//     on the bot's own conversation reference proves the bot is a member and
//     its credentials work, without anything appearing in the chat.
//
// Nothing is assumed: what cannot be confirmed is reported as missing, with
// the reason and what to do. The provider returns and never throws.

import { TeamsInfo } from 'botbuilder';
import type { ConversationReference } from 'botbuilder';

import type {
  ChannelReadinessIssue,
  ChannelReadinessProvider,
  ChannelReadinessReport,
} from './kernel-types.js';
import type { TeamsProactiveSend } from './messagesRouter.js';
import { parseTeamsBotKey, TEAMS_BOT_KEY_PREFIX } from './teamsBotIdentity.js';
import type { TeamsConversationReferenceCache } from './teamsGroupPrimitives.js';

export const CHAT_READ_ALL_RSC_PERMISSION = 'ChatMessage.Read.Chat';
/** Least-privileged tenant-wide application permission for listing a chat's
 *  grants (Graph "List permissionGrants of a chat"). */
export const GRANT_READ_GRAPH_PERMISSION = 'TeamsAppInstallation.ReadForChat.All';
/** The RSC permission that allows the same call for the one chat. */
export const GRANT_READ_RSC_PERMISSION = 'TeamsAppInstallation.Read.Chat';

const HTTP_UNAUTHORIZED = 401;
const HTTP_FORBIDDEN = 403;
const HTTP_NOT_FOUND = 404;

/** The Graph surface the check needs; `TeamsGraphResolver` satisfies it. */
export interface TeamsReadinessGraph {
  getJson(path: string): Promise<Record<string, unknown>>;
}

export interface TeamsReadinessDeps {
  /** The configured bot for an app id (lowercase), with its Graph client. */
  botFor(botAppId: string): { displayName: string; graph: TeamsReadinessGraph } | undefined;
  refs: Pick<TeamsConversationReferenceCache, 'getOrLoad'>;
  /** Read one roster member through the Bot Connector as this bot; rejects
   *  when the bot is not in the conversation or cannot authenticate. */
  probeMembership(reference: Partial<ConversationReference>, botAppId: string): Promise<void>;
}

/** The default membership probe: a roster read of page size 1. */
export function createTeamsMembershipProbe(
  sendProactive: TeamsProactiveSend,
): TeamsReadinessDeps['probeMembership'] {
  return async (reference, botAppId) => {
    await sendProactive(
      reference,
      async (turnContext) => {
        await TeamsInfo.getPagedMembers(turnContext, 1);
      },
      { botAppId },
    );
  };
}

type GrantLookup =
  | { state: 'read'; permissions: readonly string[] }
  | { state: 'refused'; status: number }
  | { state: 'chat_unknown' }
  | { state: 'failed'; detail: string };

function statusOf(err: unknown): number | undefined {
  const status = (err as { status?: unknown; statusCode?: unknown } | null)?.status
    ?? (err as { statusCode?: unknown } | null)?.statusCode;
  return typeof status === 'number' ? status : undefined;
}

function describeError(err: unknown): string {
  const status = statusOf(err);
  const text = err instanceof Error ? err.message : String(err);
  return status !== undefined ? `HTTP ${String(status)}: ${text}` : text;
}

/** Channel conversations (`…@thread.tacv2`, classic `…@thread.skype`) hang
 *  off a team, not a chat; `/chats/{id}` does not know them. */
function isTeamChannel(conversationId: string, teamsType: string | undefined): boolean {
  if (teamsType === 'channel') return true;
  if (teamsType === 'groupChat' || teamsType === 'personal') return false;
  return conversationId.includes('@thread.tacv2') || conversationId.includes('@thread.skype');
}

async function readGrants(
  graph: TeamsReadinessGraph,
  conversationId: string,
  botAppId: string,
): Promise<GrantLookup> {
  try {
    const json = await graph.getJson(`/chats/${encodeURIComponent(conversationId)}/permissionGrants`);
    const rows = Array.isArray(json['value']) ? (json['value'] as unknown[]) : [];
    const permissions = rows.flatMap((row) => {
      if (!row || typeof row !== 'object') return [];
      const grant = row as Record<string, unknown>;
      // The Graph reference describes `clientAppId` and `clientId` with the
      // app id and the service principal id in swapped roles; whichever of
      // the two carries the bot's app id identifies the grant as ours.
      const isOurs = ['clientAppId', 'clientId'].some(
        (key) => typeof grant[key] === 'string' && (grant[key] as string).toLowerCase() === botAppId,
      );
      const permission = grant['permission'];
      const isApplication = grant['permissionType'] === undefined || grant['permissionType'] === 'Application';
      return isOurs && isApplication && typeof permission === 'string' ? [permission] : [];
    });
    return { state: 'read', permissions };
  } catch (err) {
    const status = statusOf(err);
    if (status === HTTP_UNAUTHORIZED || status === HTTP_FORBIDDEN) return { state: 'refused', status };
    if (status === HTTP_NOT_FOUND) return { state: 'chat_unknown' };
    return { state: 'failed', detail: describeError(err) };
  }
}

function report(
  flags: Pick<ChannelReadinessReport, 'installed' | 'canReadAllMessages' | 'canPost'>,
  missing: readonly ChannelReadinessIssue[],
): ChannelReadinessReport {
  return { ...flags, ready: missing.length === 0, missing };
}

function consentIssue(grants: GrantLookup, bot: string, isInstalled: boolean): ChannelReadinessIssue {
  switch (grants.state) {
    case 'read':
      return {
        code: 'consent_missing',
        message: isInstalled
          ? `'${bot}' is in this chat, but the permission ${CHAT_READ_ALL_RSC_PERMISSION} is not granted to it here. It only receives messages that @mention it. Remove the app from the chat and add it again with a package that requests ${CHAT_READ_ALL_RSC_PERMISSION}; if adding fails or the permission is still missing, the tenant does not allow resource-specific consent for chats and an administrator has to enable it.`
          : `the permission ${CHAT_READ_ALL_RSC_PERMISSION} is not granted in this chat because the app of '${bot}' is not installed there.`,
      };
    case 'refused':
      return {
        code: 'consent_unverifiable',
        message: `Microsoft Graph refused to list the permission grants of this chat (HTTP ${String(grants.status)}), so the consent for ${CHAT_READ_ALL_RSC_PERMISSION} cannot be confirmed. Grant the app registration of '${bot}' the application permission ${GRANT_READ_GRAPH_PERMISSION} with admin consent, or add the resource-specific permission ${GRANT_READ_RSC_PERMISSION} to its app package.`,
      };
    case 'chat_unknown':
      return {
        code: 'consent_unverifiable',
        message: `Microsoft Graph does not know this chat (HTTP 404), so the consent for ${CHAT_READ_ALL_RSC_PERMISSION} cannot be confirmed. The id is not a chat the tenant of '${bot}' can see.`,
      };
    case 'failed':
      return {
        code: 'consent_unverifiable',
        message: `the permission grants of this chat could not be read (${grants.detail}); the consent for ${CHAT_READ_ALL_RSC_PERMISSION} is unconfirmed. Try again.`,
      };
  }
}

export function createTeamsReadinessProvider(deps: TeamsReadinessDeps): ChannelReadinessProvider {
  return {
    channelType: 'teams',
    async check(conversationId, opts): Promise<ChannelReadinessReport> {
      const none = { installed: false, canReadAllMessages: false, canPost: false };
      const botAppId = parseTeamsBotKey(opts.asChannelKey);
      if (!botAppId) {
        return report(none, [
          {
            code: 'no_identity',
            message: `'${opts.asChannelKey}' is not a Teams bot identity (expected '${TEAMS_BOT_KEY_PREFIX}<appId>'). A facilitator needs its own Teams bot.`,
          },
        ]);
      }
      const bot = deps.botFor(botAppId);
      if (!bot) {
        return report(none, [
          {
            code: 'no_identity',
            message: 'this bot identity is not configured in the Teams channel (teams_bots[]). A facilitator needs its own Teams bot; no other bot is used in its place.',
          },
        ]);
      }

      // The bot's OWN reference only: another bot's reference says nothing
      // about this one and must never be probed or posted through.
      const cached = await deps.refs.getOrLoad(conversationId, botAppId).catch(() => undefined);
      const ownerAppId = cached?.ref.bot?.id ? parseTeamsBotKey(cached.ref.bot.id) : undefined;
      const reference = cached && (ownerAppId === undefined || ownerAppId === botAppId) ? cached : undefined;

      if (isTeamChannel(conversationId, reference?.teamsType)) {
        return report(none, [
          {
            code: 'unsupported_conversation',
            message: 'this conversation is a team channel. Readiness can only be confirmed for group chats and meeting chats; facilitation in channels is not supported yet.',
          },
        ]);
      }

      let probeError: string | undefined;
      let isMember = false;
      if (reference) {
        try {
          await deps.probeMembership(reference.ref, botAppId);
          isMember = true;
        } catch (err) {
          probeError = describeError(err);
        }
      }

      const isPersonal = reference?.teamsType === 'personal';
      // A 1:1 chat delivers every message to the bot; there is no grant to read.
      const grants: GrantLookup = isPersonal
        ? { state: 'read', permissions: [CHAT_READ_ALL_RSC_PERMISSION] }
        : await readGrants(bot.graph, conversationId, botAppId);

      const hasGrant = grants.state === 'read' && grants.permissions.length > 0;
      const installed = isMember || hasGrant;
      const canReadAllMessages =
        grants.state === 'read' && grants.permissions.includes(CHAT_READ_ALL_RSC_PERMISSION);
      const canPost = isMember;

      const missing: ChannelReadinessIssue[] = [];
      if (!installed) {
        missing.push({
          code: 'app_not_installed',
          message:
            grants.state === 'read'
              ? `the app of '${bot.displayName}' is not installed in this chat. Add it to the chat (Apps, or @mention it and choose "Add").`
              : `the installation of '${bot.displayName}' in this chat could not be confirmed: no activity from this chat has reached the bot and Microsoft Graph did not answer. Add the app to the chat.`,
        });
      }
      if (!canReadAllMessages) missing.push(consentIssue(grants, bot.displayName, installed));
      if (!canPost) {
        missing.push({
          code: 'cannot_post',
          message: reference
            ? `'${bot.displayName}' cannot reach this chat through the Bot Connector (${probeError ?? 'unknown error'}). It was probably removed from the chat, or its app registration credentials are no longer valid.`
            : `'${bot.displayName}' has no conversation reference for this chat: no activity from it has reached the bot yet. ${installed ? 'Mention the bot once in the chat, or remove and add the app again.' : 'Add the app to the chat.'}`,
        });
      }
      return report({ installed, canReadAllMessages, canPost }, missing);
    },
  };
}

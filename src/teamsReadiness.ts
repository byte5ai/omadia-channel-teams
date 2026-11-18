// Readiness check for native facilitators (byte5ai/omadia#2043 §6.2, issue
// #43): can THIS bot moderate THIS conversation?
//
//   installed           an installation row for the bot's Teams app exists in
//                       the chat (or team)
//   canReadAllMessages  a resource-specific grant for the bot's app with
//                       `ChatMessage.Read.Chat` (team: `ChannelMessage.Read.Group`)
//                       exists there. Requesting it in the manifest is not
//                       enough; the grant exists per install.
//   canPost             the bot holds its own conversation reference and its
//                       last post there, if any, did not fail. No GET answers
//                       this; the first real send is the evidence.
//
// WHO ASKS GRAPH. The per-agent bot registrations hold no Graph application
// permissions (verified in the byte5 tenant, 2026-11-18). The calls therefore
// run with the INSTANCE's Graph client: the app registration of the M365
// integration, which holds `TeamsAppInstallation.ReadWrite…ForChat.All` /
// `…ForTeam.All`. That identity only reads; nothing is posted as it.
//
// WHICH ID IS WHICH. In a `resourceSpecificPermissionGrant`, `clientAppId` is
// the Entra application (client) id, i.e. the bot id, and `clientId` is the
// service principal object id. The Graph reference page describes them the
// other way round; the live values were checked against the directory.
//
// The provider returns what is missing and never throws.

import type {
  ChannelReadinessIssue,
  ChannelReadinessProvider,
  ChannelReadinessReport,
} from './kernel-types.js';
import { parseTeamsBotKey, TEAMS_BOT_KEY_PREFIX } from './teamsBotIdentity.js';
import type { TeamsConversationReferenceCache } from './teamsGroupPrimitives.js';

export const CHAT_READ_ALL_RSC_PERMISSION = 'ChatMessage.Read.Chat';
export const CHANNEL_READ_ALL_RSC_PERMISSION = 'ChannelMessage.Read.Group';
/** Least-privileged application permissions for the two reads. */
export const CHAT_INSTALL_READ_GRAPH_PERMISSION = 'TeamsAppInstallation.ReadForChat.All';
export const TEAM_INSTALL_READ_GRAPH_PERMISSION = 'TeamsAppInstallation.ReadForTeam.All';
/** What the consent upgrade needs. */
export const CHAT_CONSENT_GRAPH_PERMISSION = 'TeamsAppInstallation.ReadWriteAndConsentForChat.All';

const HTTP_UNAUTHORIZED = 401;
const HTTP_FORBIDDEN = 403;
const HTTP_NOT_FOUND = 404;
/** A chat lists every tenant-wide store app (about 35 rows in the byte5
 *  tenant); a handful of pages is more than any real chat has. */
const MAX_INSTALLED_APP_PAGES = 5;
const INSTALLED_APPS_QUERY =
  '$select=id,consentedPermissionSet' +
  '&$expand=teamsApp($select=id,displayName,externalId,distributionMethod),' +
  'teamsAppDefinition($select=version;$expand=bot($select=id))';

/** The Graph surface the check needs; `TeamsGraphResolver` satisfies it. */
export interface TeamsReadinessGraph {
  getJson(path: string): Promise<Record<string, unknown>>;
  postJson?(path: string, body: unknown): Promise<void>;
}

export interface TeamsReadinessDeps {
  /** Display name of a CONFIGURED bot (app id lowercase), else `undefined`. */
  botLabel(botAppId: string): string | undefined;
  /** The instance's Graph client (the M365 integration's registration).
   *  `undefined` when its credentials are not available to this plugin. */
  graph: TeamsReadinessGraph | undefined;
  refs: Pick<TeamsConversationReferenceCache, 'getOrLoad' | 'teamIdFor' | 'lastPostResult'>;
}

/** Where Graph keeps the installation: a chat, or the team of a channel. */
type Scope =
  | { kind: 'chat'; path: string; permission: string; readPermission: string }
  | { kind: 'team'; path: string; permission: string; readPermission: string };

interface Installation {
  id: string;
  consented: readonly string[];
  version?: string;
}

type Lookup<T> =
  | { state: 'read'; value: T }
  | { state: 'forbidden'; status: number }
  | { state: 'unknown_resource' }
  | { state: 'failed'; detail: string };

function statusOf(err: unknown): number | undefined {
  const record = (err ?? {}) as { status?: unknown; statusCode?: unknown };
  const status = record.status ?? record.statusCode;
  return typeof status === 'number' ? status : undefined;
}

function describeError(err: unknown): string {
  const status = statusOf(err);
  const text = err instanceof Error ? err.message : String(err);
  return status !== undefined && status > 0 ? `HTTP ${String(status)}: ${text}` : text;
}

async function lookup<T>(read: () => Promise<T>): Promise<Lookup<T>> {
  try {
    return { state: 'read', value: await read() };
  } catch (err) {
    const status = statusOf(err);
    if (status === HTTP_UNAUTHORIZED || status === HTTP_FORBIDDEN) return { state: 'forbidden', status };
    if (status === HTTP_NOT_FOUND) return { state: 'unknown_resource' };
    return { state: 'failed', detail: describeError(err) };
  }
}

function rows(json: Record<string, unknown>): Record<string, unknown>[] {
  const value = json['value'];
  return Array.isArray(value)
    ? value.filter((row): row is Record<string, unknown> => Boolean(row) && typeof row === 'object')
    : [];
}

function permissionNames(consentedPermissionSet: unknown): string[] {
  const list = (consentedPermissionSet as { resourceSpecificPermissions?: unknown } | null)
    ?.resourceSpecificPermissions;
  if (!Array.isArray(list)) return [];
  return list.flatMap((entry) => {
    const value = (entry as { permissionValue?: unknown } | null)?.permissionValue;
    return typeof value === 'string' ? [value] : [];
  });
}

/** The installation row whose app definition carries this bot. */
async function findInstallation(
  graph: TeamsReadinessGraph,
  scope: Scope,
  botAppId: string,
): Promise<Installation | undefined> {
  let next: string | undefined = `${scope.path}/installedApps?${INSTALLED_APPS_QUERY}`;
  for (let page = 0; next !== undefined && page < MAX_INSTALLED_APP_PAGES; page += 1) {
    const json: Record<string, unknown> = await graph.getJson(next);
    for (const row of rows(json)) {
      const definition = row['teamsAppDefinition'] as { version?: unknown; bot?: { id?: unknown } } | undefined;
      const botId = definition?.bot?.id;
      if (typeof botId !== 'string' || botId.toLowerCase() !== botAppId) continue;
      if (typeof row['id'] !== 'string') continue;
      return {
        id: row['id'],
        consented: permissionNames(row['consentedPermissionSet']),
        ...(typeof definition?.version === 'string' ? { version: definition.version } : {}),
      };
    }
    const link = json['@odata.nextLink'];
    next = typeof link === 'string' ? link : undefined;
  }
  return undefined;
}

/** Application grants of this bot's app in the chat or team. */
async function readGrants(graph: TeamsReadinessGraph, scope: Scope, botAppId: string): Promise<string[]> {
  const json = await graph.getJson(`${scope.path}/permissionGrants`);
  return rows(json).flatMap((grant) => {
    const clientAppId = grant['clientAppId'];
    const permission = grant['permission'];
    const isOurs = typeof clientAppId === 'string' && clientAppId.toLowerCase() === botAppId;
    const isApplication = grant['permissionType'] === undefined || grant['permissionType'] === 'Application';
    return isOurs && isApplication && typeof permission === 'string' ? [permission] : [];
  });
}

function graphIssue(what: string, result: Exclude<Lookup<unknown>, { state: 'read' }>, scope: Scope): ChannelReadinessIssue {
  if (result.state === 'forbidden') {
    return {
      code: 'graph_forbidden',
      message: `Microsoft Graph refused to list the ${what} of this ${scope.kind} (HTTP ${String(result.status)}). The instance's app registration (the one of the Microsoft 365 integration) needs the application permission ${scope.readPermission} with admin consent.`,
    };
  }
  if (result.state === 'unknown_resource') {
    return {
      code: 'graph_error',
      message: `Microsoft Graph does not know this ${scope.kind} (HTTP 404), so its ${what} cannot be read.`,
    };
  }
  return { code: 'graph_error', message: `the ${what} of this ${scope.kind} could not be read (${result.detail}). Try again.` };
}

function report(
  flags: Pick<ChannelReadinessReport, 'installed' | 'canReadAllMessages' | 'canPost'>,
  missing: readonly ChannelReadinessIssue[],
): ChannelReadinessReport {
  return { ...flags, ready: missing.length === 0, missing };
}

const NOTHING = { installed: false, canReadAllMessages: false, canPost: false } as const;

/** The outcome of {@link TeamsReadinessProvider.repairConsent}. */
export type ConsentRepairOutcome =
  | { outcome: 'upgraded'; installationId: string; permissions: readonly string[] }
  | { outcome: 'not_needed' }
  | { outcome: 'failed'; code: string; message: string };

/** The kernel contract plus the fix for a missing grant. The extra method is
 *  a superset: a kernel that does not know it never calls it. */
export interface TeamsReadinessProvider extends ChannelReadinessProvider {
  /**
   * Fix `consent_missing` for an app that IS installed: Graph
   * `POST /chats/{id}/installedApps/{installation}/upgrade` with the full
   * `consentedPermissionSet` the app definition declares. The set must match
   * the definition exactly, otherwise Graph answers 400
   * `ResourceSpecificPermissionsMismatch`; it is therefore read from the
   * installed definition, not from this repo's template.
   */
  repairConsent(conversationId: string, opts: { asChannelKey: string }): Promise<ConsentRepairOutcome>;
}

export function createTeamsReadinessProvider(deps: TeamsReadinessDeps): TeamsReadinessProvider {
  /**
   * A conversation is a team channel only on positive evidence: Teams said
   * `channel`, or the id has the channel suffix. `conversationType` is NOT
   * used to tell chats apart: a 1:1 between two people with an agent app
   * arrives as `groupChat`, and old group chats carry `@thread.skype`.
   */
  const scopeFor = (conversationId: string, teamsType: string | undefined): Scope | { missingTeamId: true } => {
    const isChannel = teamsType === 'channel' || conversationId.includes('@thread.tacv2');
    if (!isChannel) {
      return {
        kind: 'chat',
        path: `/chats/${encodeURIComponent(conversationId)}`,
        permission: CHAT_READ_ALL_RSC_PERMISSION,
        readPermission: CHAT_INSTALL_READ_GRAPH_PERMISSION,
      };
    }
    const teamId = deps.refs.teamIdFor(conversationId);
    if (!teamId) return { missingTeamId: true };
    return {
      kind: 'team',
      path: `/teams/${encodeURIComponent(teamId)}`,
      permission: CHANNEL_READ_ALL_RSC_PERMISSION,
      readPermission: TEAM_INSTALL_READ_GRAPH_PERMISSION,
    };
  };

  const identify = (asChannelKey: string): { botAppId: string; label: string } | ChannelReadinessIssue => {
    const botAppId = parseTeamsBotKey(asChannelKey);
    if (!botAppId) {
      return {
        code: 'no_identity',
        message: `'${asChannelKey}' is not a Teams bot identity (expected '${TEAMS_BOT_KEY_PREFIX}<appId>'). A facilitator needs its own Teams bot.`,
      };
    }
    const label = deps.botLabel(botAppId);
    if (label === undefined) {
      return {
        code: 'no_identity',
        message: 'this bot identity is not configured in the Teams channel (teams_bots[]). A facilitator needs its own Teams bot; no other bot is used in its place.',
      };
    }
    return { botAppId, label };
  };

  /** The bot's OWN reference only: another bot's reference says nothing
   *  about this one and must never be posted through. */
  const ownReference = async (conversationId: string, botAppId: string) => {
    const cached = await deps.refs.getOrLoad(conversationId, botAppId).catch(() => undefined);
    const ownerAppId = cached?.ref.bot?.id ? parseTeamsBotKey(cached.ref.bot.id) : undefined;
    return cached && (ownerAppId === undefined || ownerAppId === botAppId) ? cached : undefined;
  };

  return {
    channelType: 'teams',

    async check(conversationId, opts): Promise<ChannelReadinessReport> {
      const identity = identify(opts.asChannelKey);
      if ('code' in identity) return report(NOTHING, [identity]);
      const { botAppId, label } = identity;

      const reference = await ownReference(conversationId, botAppId);
      const lastPost = deps.refs.lastPostResult(conversationId, botAppId);
      const canPost = reference !== undefined && lastPost?.ok !== false;
      const postIssue: ChannelReadinessIssue | undefined = canPost
        ? undefined
        : {
            code: 'cannot_post',
            message: reference
              ? `the last post of '${label}' into this conversation failed (${lastPost?.detail ?? 'unknown error'}). It was probably removed, or its app registration credentials are no longer valid.`
              : `'${label}' has no conversation reference here: no activity from this conversation has reached the bot yet. Mention the bot once, or remove and add its app again.`,
          };

      // The bot's own 1:1 with a person: every message reaches the bot and
      // there is no grant to read.
      if (reference?.teamsType === 'personal') {
        return report({ installed: true, canReadAllMessages: true, canPost }, postIssue ? [postIssue] : []);
      }

      const scope = scopeFor(conversationId, reference?.teamsType);
      if ('missingTeamId' in scope) {
        return report({ ...NOTHING, canPost }, [
          {
            code: 'team_unknown',
            message: `this is a team channel and '${label}' has not received an activity from it since the last restart, so the team is not known. Mention the bot once in the channel and check again.`,
          },
          ...(postIssue ? [postIssue] : []),
        ]);
      }
      if (!deps.graph) {
        return report({ ...NOTHING, canPost }, [
          {
            code: 'graph_unavailable',
            message: 'the credentials of the instance\'s app registration (Microsoft 365 integration) are not available to the Teams channel, so installation and consent cannot be read.',
          },
          ...(postIssue ? [postIssue] : []),
        ]);
      }
      const graph = deps.graph;

      const [installation, grants] = await Promise.all([
        lookup(() => findInstallation(graph, scope, botAppId)),
        lookup(() => readGrants(graph, scope, botAppId)),
      ]);

      const installed = installation.state === 'read' && installation.value !== undefined;
      // Ready needs BOTH rows: a grant without an installation row is a
      // leftover, an installation without the grant hears mentions only.
      const canReadAllMessages =
        installed && grants.state === 'read' && grants.value.includes(scope.permission);

      const missing: ChannelReadinessIssue[] = [];
      if (installation.state !== 'read') {
        missing.push(graphIssue('installed apps', installation, scope));
      } else if (!installed) {
        missing.push({
          code: 'app_not_installed',
          message: `the app of '${label}' is not installed in this ${scope.kind}. Add it there.`,
        });
      }
      if (grants.state !== 'read') {
        if (installation.state === 'read') missing.push(graphIssue('permission grants', grants, scope));
      } else if (installed && !canReadAllMessages) {
        missing.push({
          code: 'consent_missing',
          message: `'${label}' is installed in this ${scope.kind}${installation.state === 'read' && installation.value?.version ? ` (app version ${installation.value.version})` : ''}, but the permission ${scope.permission} is not granted to it here, so it only receives messages that @mention it. Fix: upgrade the installation with the full consented permission set (Graph installedApps/{id}/upgrade, needs ${CHAT_CONSENT_GRAPH_PERMISSION} on the instance's app registration). If the tenant blocks resource-specific consent for apps, an administrator has to allow it first.`,
        });
      }
      if (postIssue) missing.push(postIssue);
      return report({ installed, canReadAllMessages, canPost }, missing);
    },

    async repairConsent(conversationId, opts): Promise<ConsentRepairOutcome> {
      const identity = identify(opts.asChannelKey);
      if ('code' in identity) return { outcome: 'failed', code: identity.code, message: identity.message };
      const graph = deps.graph;
      if (!graph?.postJson) {
        return { outcome: 'failed', code: 'graph_unavailable', message: 'no Graph client that can write is available to the Teams channel.' };
      }
      const reference = await ownReference(conversationId, identity.botAppId);
      const scope = scopeFor(conversationId, reference?.teamsType);
      if ('missingTeamId' in scope || scope.kind !== 'chat') {
        return { outcome: 'failed', code: 'unsupported_conversation', message: 'the consent upgrade is implemented for chats only.' };
      }
      try {
        const installation = await findInstallation(graph, scope, identity.botAppId);
        if (!installation) {
          return { outcome: 'failed', code: 'app_not_installed', message: `the app of '${identity.label}' is not installed in this chat; there is nothing to upgrade.` };
        }
        const granted = await readGrants(graph, scope, identity.botAppId);
        if (granted.includes(scope.permission)) return { outcome: 'not_needed' };
        // The set the INSTALLED definition declares; Graph rejects any other.
        const detail = await graph.getJson(
          `${scope.path}/installedApps/${encodeURIComponent(installation.id)}?$expand=teamsAppDefinition`,
        );
        const declared = (
          detail['teamsAppDefinition'] as { authorization?: { requiredPermissionSet?: unknown } } | undefined
        )?.authorization?.requiredPermissionSet as { resourceSpecificPermissions?: unknown } | undefined;
        const declaredList = Array.isArray(declared?.resourceSpecificPermissions)
          ? (declared.resourceSpecificPermissions as { permissionValue?: unknown; permissionType?: unknown }[])
          : [];
        const resourceSpecificPermissions = declaredList.flatMap((entry) =>
          typeof entry.permissionValue === 'string'
            ? [{ permissionValue: entry.permissionValue, permissionType: typeof entry.permissionType === 'string' ? entry.permissionType : 'Application' }]
            : [],
        );
        const names = resourceSpecificPermissions.map((entry) => entry.permissionValue);
        if (!names.includes(scope.permission)) {
          return {
            outcome: 'failed',
            code: 'package_outdated',
            message: `the installed app version of '${identity.label}' does not declare ${scope.permission}; publish a current app package to the catalog first, then upgrade.`,
          };
        }
        await graph.postJson(`${scope.path}/installedApps/${encodeURIComponent(installation.id)}/upgrade`, {
          consentedPermissionSet: { resourceSpecificPermissions },
        });
        return { outcome: 'upgraded', installationId: installation.id, permissions: names };
      } catch (err) {
        const status = statusOf(err);
        return {
          outcome: 'failed',
          code: status === HTTP_UNAUTHORIZED || status === HTTP_FORBIDDEN ? 'graph_forbidden' : 'graph_error',
          message: describeError(err),
        };
      }
    },
  };
}

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
  ChannelConsentRepairOutcome,
  ChannelReadinessIssue,
  ChannelReadinessProvider,
  ChannelReadinessReport,
} from './kernel-types.js';
import { templateRevisionOfAppVersion, type TeamsAppTemplateInfo } from './teamsAppTemplate.js';
import { parseTeamsBotKey, TEAMS_BOT_KEY_PREFIX } from './teamsBotIdentity.js';
import {
  bareConversationId,
  botKeyPrefixedConversationMessage,
  isBotKeyPrefixedConversationId,
  isChannelConversationId,
  reportBotKeyPrefixedConversationId,
} from './teamsConversationIds.js';
import type { TeamsConversationReferenceCache } from './teamsGroupPrimitives.js';

export const CHAT_READ_ALL_RSC_PERMISSION = 'ChatMessage.Read.Chat';
export const CHANNEL_READ_ALL_RSC_PERMISSION = 'ChannelMessage.Read.Group';
/** Least-privileged application permissions for the two reads. */
export const CHAT_INSTALL_READ_GRAPH_PERMISSION = 'TeamsAppInstallation.ReadForChat.All';
export const TEAM_INSTALL_READ_GRAPH_PERMISSION = 'TeamsAppInstallation.ReadForTeam.All';
/** What the consent upgrade needs. */
export const CHAT_CONSENT_GRAPH_PERMISSION = 'TeamsAppInstallation.ReadWriteAndConsentForChat.All';
export const TEAM_CONSENT_GRAPH_PERMISSION = 'TeamsAppInstallation.ReadWriteAndConsentForTeam.All';
/** What reading the published app definition needs. */
export const APP_CATALOG_GRAPH_PERMISSION = 'AppCatalog.ReadWrite.All';

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
  refs: Pick<TeamsConversationReferenceCache, 'getOrLoad' | 'resolveTeamId' | 'lastPostResult'>;
  /** Why `graph` is absent, naming the exact setting or secret. The kernel
   *  refuses a facilitation that is not ready, so this message is what an
   *  operator has to act on. */
  graphUnavailableReason?: string;
  /** The shipped app package template; absent when it could not be read, in
   *  which case no install is ever called outdated. */
  template?: TeamsAppTemplateInfo;
}

/** Where Graph keeps the installation: a chat, or the team of a channel. */
interface Scope {
  kind: 'chat' | 'team';
  /** Graph path of the chat or the team. Never a thread id, never a session
   *  scope fragment. */
  path: string;
  /** The RSC permission that lets the bot read every message here. */
  permission: string;
  /** Graph application permission to read installs and grants. */
  readPermission: string;
  /** Graph application permission for the consent upgrade. */
  consentPermission: string;
  /** Suffix of the RSC permissions that apply to this scope. */
  rscSuffix: string;
}

interface Installation {
  id: string;
  /** Catalog id of the app (`teamsApp.id`). */
  catalogAppId?: string;
  consented: readonly string[];
  version?: string;
}

/** Permissions Graph can grant in a chat; `.Group` ones apply to a team. */
const CHAT_SCOPE_SUFFIX = '.Chat';
const TEAM_SCOPE_SUFFIX = '.Group';

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
      const catalogAppId = (row['teamsApp'] as { id?: unknown } | undefined)?.id;
      return {
        id: row['id'],
        ...(typeof catalogAppId === 'string' ? { catalogAppId } : {}),
        consented: permissionNames(row['consentedPermissionSet']),
        ...(typeof definition?.version === 'string' ? { version: definition.version } : {}),
      };
    }
    const link = json['@odata.nextLink'];
    next = typeof link === 'string' ? link : undefined;
  }
  return undefined;
}

/**
 * The version and the RSC set of the PUBLISHED definition of a catalog app:
 * `GET /appCatalogs/teamsApps?$filter=id eq '…'&$expand=appDefinitions($select=id,version,authorization)`.
 * This is the query Graph's install reference names for the set an install
 * or upgrade must consent to, and the one the M365 connector's provisioner
 * uses at install; an upgrade moves the install to that definition, so its
 * set is the one that has to match.
 */
async function publishedDefinition(graph: TeamsReadinessGraph, catalogAppId: string): Promise<PublishedDefinition> {
  const filter = encodeURIComponent(`id eq '${catalogAppId.replace(/'/g, "''")}'`);
  const expand = encodeURIComponent('appDefinitions($select=id,version,authorization)');
  const json = await graph.getJson(`/appCatalogs/teamsApps?$filter=${filter}&$expand=${expand}`);
  const definitions: CatalogDefinition[] = rows(json)
    // Re-checked client-side: a response that ignores the filter must not
    // make us consent to another app's permissions.
    .filter((app) => app['id'] === catalogAppId && Array.isArray(app['appDefinitions']))
    .flatMap((app) => app['appDefinitions'] as CatalogDefinition[])
    .filter((definition) => Boolean(definition) && typeof definition === 'object');
  // The upgrade moves the install to the CURRENT definition, and Graph wants
  // exactly its set: a permission an earlier definition declared and the
  // current one dropped must not be sent. Organization apps carry one
  // definition; should there be several, the highest version is the current.
  const current = definitions.reduce<CatalogDefinition | undefined>(
    (best, candidate) => (best === undefined || compareAppVersions(candidate.version, best.version) > 0 ? candidate : best),
    undefined,
  );
  const list = (current?.authorization?.requiredPermissionSet as { resourceSpecificPermissions?: unknown } | undefined)
    ?.resourceSpecificPermissions;
  const permissions: { permissionValue: string; permissionType: string }[] = [];
  const seen = new Set<string>();
  for (const entry of (Array.isArray(list) ? list : []) as { permissionValue?: unknown; permissionType?: unknown }[]) {
    if (typeof entry?.permissionValue !== 'string' || typeof entry.permissionType !== 'string') continue;
    const permissionType = entry.permissionType.toLowerCase();
    if (permissionType !== 'application' && permissionType !== 'delegated') continue;
    const key = `${permissionType}:${entry.permissionValue}`;
    if (seen.has(key)) continue;
    seen.add(key);
    permissions.push({ permissionValue: entry.permissionValue, permissionType });
  }
  return { permissions, ...(typeof current?.version === 'string' ? { version: current.version } : {}) };
}

interface CatalogDefinition {
  version?: unknown;
  authorization?: { requiredPermissionSet?: unknown };
}

interface PublishedDefinition {
  permissions: { permissionValue: string; permissionType: string }[];
  /** `teamsAppDefinition.version` of the published definition, when stated. */
  version?: string;
}

/** Orders dotted numeric app versions; a missing or non-numeric one is lowest. */
function compareAppVersions(a: unknown, b: unknown): number {
  const parts = (value: unknown): number[] | undefined =>
    typeof value === 'string' && /^\d+(\.\d+)*$/.test(value.trim()) ? value.trim().split('.').map(Number) : undefined;
  const left = parts(a);
  const right = parts(b);
  if (!left || !right) return left ? 1 : right ? -1 : 0;
  for (let i = 0; i < Math.max(left.length, right.length); i += 1) {
    const diff = (left[i] ?? 0) - (right[i] ?? 0);
    if (diff !== 0) return diff;
  }
  return 0;
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

/** The outcome of {@link TeamsReadinessProvider.repairConsent}: the kernel
 *  contract's `ChannelConsentRepairOutcome`. */
export type ConsentRepairOutcome = ChannelConsentRepairOutcome;

/** The kernel contract with `repairConsent` always present, plus one method
 *  the meeting source uses. */
export interface TeamsReadinessProvider extends ChannelReadinessProvider {
  /**
   * Bring an install to the full consent set of the published app: Graph
   * `POST /chats/{id}/installedApps/{installation}/upgrade` with the complete
   * `consentedPermissionSet` the definition declares. Fixes `consent_missing`
   * and gives an install from an earlier package the permissions added since
   * (the meeting RSC). The set must match the definition exactly, otherwise
   * Graph answers 400 `ResourceSpecificPermissionsMismatch`; it is therefore
   * read from the catalog, never from this repo's template.
   */
  repairConsent(conversationId: string, opts: { asChannelKey: string }): Promise<ConsentRepairOutcome>;
  /**
   * The RSC permissions granted to this bot's app in a chat right now, read
   * from Graph. `undefined` when that cannot be read (no Graph client, a
   * refusal, a channel, an unknown identity): the caller must then treat
   * every permission as not granted.
   */
  grantedPermissions(conversationId: string, opts: { asChannelKey: string }): Promise<readonly string[] | undefined>;
}

export function createTeamsReadinessProvider(deps: TeamsReadinessDeps): TeamsReadinessProvider {
  /**
   * A conversation is a team channel only on positive evidence: Teams said
   * `channel`, or the id has the channel suffix. `conversationType` is NOT
   * used to tell chats apart: a 1:1 between two people with an agent app
   * arrives as `groupChat`, and old group chats carry `@thread.skype`.
   */
  const scopeFor = async (
    conversationId: string,
    teamsType: string | undefined,
    botAppId: string,
  ): Promise<Scope | { missingTeamId: true }> => {
    const isChannel = teamsType === 'channel' || isChannelConversationId(conversationId);
    if (!isChannel) {
      return {
        kind: 'chat',
        // The plain chat id. A chat id has no thread part; cut one anyway
        // rather than hand Graph something it answers with a bare 400.
        path: `/chats/${encodeURIComponent(bareConversationId(conversationId))}`,
        permission: CHAT_READ_ALL_RSC_PERMISSION,
        readPermission: CHAT_INSTALL_READ_GRAPH_PERMISSION,
        consentPermission: CHAT_CONSENT_GRAPH_PERMISSION,
        rscSuffix: CHAT_SCOPE_SUFFIX,
      };
    }
    // A channel, or one thread of it (`…;messageid=<root>`): Graph keeps the
    // installation and the grants on the TEAM. The thread id is this
    // plugin's and the kernel's key only.
    const teamId = await deps.refs.resolveTeamId(conversationId, botAppId);
    if (!teamId) return { missingTeamId: true };
    return {
      kind: 'team',
      path: `/teams/${encodeURIComponent(teamId)}`,
      permission: CHANNEL_READ_ALL_RSC_PERMISSION,
      readPermission: TEAM_INSTALL_READ_GRAPH_PERMISSION,
      consentPermission: TEAM_CONSENT_GRAPH_PERMISSION,
      rscSuffix: TEAM_SCOPE_SUFFIX,
    };
  };

  /** #2043 — `28:<appId>-<id>` is a fragment of this plugin's session scope,
   *  not a conversation id. Refused by name: Graph would answer a bare 400
   *  and no reference exists under such an id. */
  const callerBug = (conversationId: string, where: string): ChannelReadinessIssue | undefined => {
    if (!isBotKeyPrefixedConversationId(conversationId)) return undefined;
    reportBotKeyPrefixedConversationId(conversationId, where);
    return { code: 'unsupported_conversation', message: botKeyPrefixedConversationMessage(conversationId) };
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
      const wrongId = callerBug(conversationId, 'readiness check');
      if (wrongId) return report(NOTHING, [wrongId]);

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

      const scope = await scopeFor(conversationId, reference?.teamsType, botAppId);
      if ('missingTeamId' in scope) {
        return report({ ...NOTHING, canPost }, [
          {
            code: 'team_unknown',
            message: `this is a team channel and '${label}' has no stored activity from it that names the team, so the team is not known. Mention the bot once in the channel and check again.`,
          },
          ...(postIssue ? [postIssue] : []),
        ]);
      }
      if (!deps.graph) {
        return report({ ...NOTHING, canPost }, [
          {
            code: 'graph_unavailable',
            message:
              deps.graphUnavailableReason ??
              'the credentials of the instance\'s app registration (Microsoft 365 integration) are not available to the Teams channel, so installation and consent cannot be read.',
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
          message: `'${label}' is installed in this ${scope.kind}${installation.state === 'read' && installation.value?.version ? ` (app version ${installation.value.version})` : ''}, but the permission ${scope.permission} is not granted to it here, so it only receives messages that @mention it. Fix: upgrade the installation with the full consented permission set (Graph installedApps/{id}/upgrade, needs ${scope.consentPermission} on the instance's app registration). If the tenant blocks resource-specific consent for apps, an administrator has to allow it first.`,
        });
      }
      // An install made from an earlier app package: the chat can be read,
      // but the permission set differs from the shipped template (revision 2
      // added the meeting permissions, revision 3 dropped one that no team
      // accepts). The kernel's fix is `repairConsent`.
      const installedRevision =
        installation.state === 'read' ? templateRevisionOfAppVersion(installation.value?.version) : undefined;
      if (
        canReadAllMessages &&
        deps.template &&
        installedRevision !== undefined &&
        installedRevision < deps.template.revision
      ) {
        missing.push({
          code: 'package_outdated',
          message: `'${label}' is installed in this ${scope.kind} from an earlier app package (version ${installation.state === 'read' ? (installation.value?.version ?? '?') : '?'}, template revision ${String(installedRevision)}; current is ${String(deps.template.revision)}). The permission set of the app changed since. Fix: repair the consent, which upgrades the installation to the published package.`,
        });
      }
      if (postIssue) missing.push(postIssue);
      return report({ installed, canReadAllMessages, canPost }, missing);
    },

    async grantedPermissions(conversationId, opts): Promise<readonly string[] | undefined> {
      const identity = identify(opts.asChannelKey);
      if ('code' in identity || !deps.graph || callerBug(conversationId, 'granted permissions')) return undefined;
      const reference = await ownReference(conversationId, identity.botAppId);
      const scope = await scopeFor(conversationId, reference?.teamsType, identity.botAppId);
      if ('missingTeamId' in scope || scope.kind !== 'chat') return undefined;
      const graph = deps.graph;
      const grants = await lookup(() => readGrants(graph, scope, identity.botAppId));
      return grants.state === 'read' ? grants.value : undefined;
    },

    async repairConsent(conversationId, opts): Promise<ConsentRepairOutcome> {
      const identity = identify(opts.asChannelKey);
      if ('code' in identity) return { outcome: 'failed', code: identity.code, message: identity.message };
      const graph = deps.graph;
      if (!graph?.postJson) {
        return {
          outcome: 'failed',
          code: 'graph_unavailable',
          message: deps.graphUnavailableReason ?? 'no Graph client that can write is available to the Teams channel.',
        };
      }
      const wrongId = callerBug(conversationId, 'consent repair');
      if (wrongId) return { outcome: 'failed', code: 'unsupported_conversation', message: wrongId.message };
      const reference = await ownReference(conversationId, identity.botAppId);
      const scope = await scopeFor(conversationId, reference?.teamsType, identity.botAppId);
      if ('missingTeamId' in scope) {
        return {
          outcome: 'failed',
          code: 'team_unknown',
          message: `this is a team channel and '${identity.label}' has no stored activity from it that names the team, so there is no team to upgrade the installation in. Mention the bot once in the channel and repair again.`,
        };
      }
      // Each Graph step names the permission it needs: "forbidden" alone
      // would not tell an operator what to grant.
      const step = async <T>(read: () => Promise<T>, what: string, permission: string): Promise<{ value: T } | ConsentRepairOutcome> => {
        try {
          return { value: await read() };
        } catch (err) {
          const status = statusOf(err);
          if (status === HTTP_UNAUTHORIZED || status === HTTP_FORBIDDEN) {
            return {
              outcome: 'failed',
              code: 'graph_forbidden',
              message: `Microsoft Graph refused to ${what} (HTTP ${String(status)}). The instance's app registration (the one of the Microsoft 365 integration) needs the application permission ${permission} with admin consent.`,
            };
          }
          return { outcome: 'failed', code: 'graph_error', message: `could not ${what}: ${describeError(err)}` };
        }
      };

      const found = await step(
        () => findInstallation(graph, scope, identity.botAppId),
        `list the apps installed in this ${scope.kind}`,
        scope.readPermission,
      );
      if ('outcome' in found) return found;
      const installation = found.value;
      if (!installation) {
        return { outcome: 'failed', code: 'app_not_installed', message: `the app of '${identity.label}' is not installed in this ${scope.kind}; there is nothing to upgrade.` };
      }
      if (!installation.catalogAppId) {
        return { outcome: 'failed', code: 'graph_error', message: 'the installation row carries no catalog app id; the declared permission set cannot be read.' };
      }
      const catalogAppId = installation.catalogAppId;
      // The set of the PUBLISHED definition: the upgrade moves the install
      // to it, and Graph compares the consented set against it.
      const declared = await step(
        () => publishedDefinition(graph, catalogAppId),
        'read the published app definition from the tenant app catalog',
        APP_CATALOG_GRAPH_PERMISSION,
      );
      if ('outcome' in declared) return declared;
      const resourceSpecificPermissions = declared.value.permissions;
      const publishedVersion = declared.value.version;
      const names = resourceSpecificPermissions.map((entry) => entry.permissionValue);
      // What the published app has to declare before an upgrade is worth it:
      // the read permission, and every permission of the shipped template
      // that applies to this scope (chat or team).
      const expected = [
        scope.permission,
        ...(deps.template?.permissions.filter((name) => name.endsWith(scope.rscSuffix)) ?? []),
      ];
      const undeclared = [...new Set(expected)].filter((name) => !names.includes(name));
      if (undeclared.length > 0) {
        return {
          outcome: 'failed',
          code: 'package_outdated',
          message: `the published app of '${identity.label}' does not declare ${undeclared.join(', ')}. Publish the current app package${deps.template ? ` (template revision ${String(deps.template.revision)})` : ''} to the tenant app catalog first, then repair again.`,
        };
      }
      // The published app itself is from an earlier template: nothing new
      // would be declared, and a permission the template dropped since would
      // stay. (The `undeclared` check above cannot see a removal.)
      const publishedRevision = templateRevisionOfAppVersion(publishedVersion);
      if (deps.template && publishedRevision !== undefined && publishedRevision < deps.template.revision) {
        return {
          outcome: 'failed',
          code: 'package_outdated',
          message: `the published app of '${identity.label}' (version ${publishedVersion ?? '?'}) is built from template revision ${String(publishedRevision)}; current is ${String(deps.template.revision)}. Publish the current app package to the tenant app catalog first, then repair again.`,
        };
      }
      // Needed while any permission of this scope the definition declares is
      // not granted here (an older install is brought to the FULL set, not
      // only to the one permission the readiness check looks at), and while
      // the install is on another version than the published definition:
      // that is how an install loses a permission the template dropped.
      const granted = await step(
        () => readGrants(graph, scope, identity.botAppId),
        `list the permission grants of this ${scope.kind}`,
        scope.readPermission,
      );
      if ('outcome' in granted) return granted;
      const isComplete = names
        .filter((name) => name.endsWith(scope.rscSuffix))
        .every((name) => granted.value.includes(name));
      const isOnPublishedVersion =
        installation.version === undefined || publishedVersion === undefined || installation.version === publishedVersion;
      if (isComplete && isOnPublishedVersion) return { outcome: 'not_needed' };
      const upgrade = graph.postJson.bind(graph);
      const upgraded = await step(
        () =>
          upgrade(`${scope.path}/installedApps/${encodeURIComponent(installation.id)}/upgrade`, {
            consentedPermissionSet: { resourceSpecificPermissions },
          }),
        `upgrade the app installation in this ${scope.kind}`,
        scope.consentPermission,
      );
      if ('outcome' in upgraded) return upgraded;
      return { outcome: 'upgraded', installationId: installation.id, permissions: names };
    },
  };
}

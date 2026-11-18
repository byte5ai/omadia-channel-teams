import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  APP_CATALOG_GRAPH_PERMISSION,
  CHAT_CONSENT_GRAPH_PERMISSION,
  CHAT_INSTALL_READ_GRAPH_PERMISSION,
  CHAT_READ_ALL_RSC_PERMISSION,
  createTeamsReadinessProvider,
  TeamsConversationReferenceCache,
  TeamsGraphResolver,
} from '@omadia/channel-teams';

import {
  ANNA,
  CHAT_RSC,
  GROUP_CHAT_ID,
  JONAS,
  LEGACY_CHAT_RSC,
  LEGACY_TEMPLATE_RSC,
  REVISION_2_TEMPLATE_RSC,
  messageActivity,
  MIRA,
  SimulatedGraph,
  TEMPLATE_RSC,
  TENANT_ID,
  type SimBot,
} from './support/teamsSimulation.js';

// byte5ai/omadia#2043 §6.2, issue #43 — readiness: installation row, RSC
// grant, can post. Graph is simulated with the row shapes read from the byte5
// tenant; the calls run with the INSTANCE's Graph client, never a bot's own.

const INSTANCE_APP_ID = '737c6ddd-0000-4000-8000-000000000001';
const TEAM_ID = '0f1e2d3c-4b5a-4687-9a8b-7c6d5e4f3a2b';
const CHANNEL_ID = '19:5f0c1a2b3c4d4e5f8a9b0c1d2e3f4a5b@thread.tacv2';

interface Harness {
  graph: SimulatedGraph;
  refs: TeamsConversationReferenceCache;
  provider: ReturnType<typeof createTeamsReadinessProvider>;
  check(bot: SimBot, conversationId?: string): ReturnType<ReturnType<typeof createTeamsReadinessProvider>['check']>;
}

function harness(opts: { configured?: SimBot[]; withGraph?: boolean; reason?: string } = {}): Harness {
  const configured = opts.configured ?? [MIRA, JONAS];
  const graph = new SimulatedGraph();
  const refs = new TeamsConversationReferenceCache();
  const provider = createTeamsReadinessProvider({
    botLabel: (botAppId) => configured.find((candidate) => candidate.appId === botAppId)?.name,
    graph:
      opts.withGraph === false
        ? undefined
        : new TeamsGraphResolver({
            tenantId: TENANT_ID,
            clientId: INSTANCE_APP_ID,
            clientSecret: 'simulated',
            log: () => {},
            fetchImpl: graph.fetch,
          }),
    refs,
    ...(opts.reason ? { graphUnavailableReason: opts.reason } : {}),
    // The shipped template: revision 3 (meeting permissions, no file permission).
    template: { revision: 3, permissions: TEMPLATE_RSC },
  });
  return {
    graph,
    refs,
    provider,
    check: (bot, conversationId = GROUP_CHAT_ID) => provider.check(conversationId, { asChannelKey: bot.key }),
  };
}

/** The bot has seen an activity from the conversation, so it holds a reference. */
function seen(h: Harness, bot: SimBot, conversationId = GROUP_CHAT_ID, extra: Record<string, unknown> = {}): void {
  h.refs.capture({
    activity: { ...messageActivity({ bot, from: ANNA, text: 'hi', conversationId }), ...extra },
  } as unknown as TurnContext);
}

describe('#43 readiness provider', () => {
  it('is ready with an installation row, the RSC grant and an own conversation reference', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA);

    const report = await h.check(MIRA);

    assert.deepEqual(report, { installed: true, canReadAllMessages: true, canPost: true, ready: true, missing: [] });
  });

  it('asks Graph as the instance, never as the facilitator bot', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA);

    await h.check(MIRA);

    const tokenRequests = h.graph.requests.filter((r) => r.includes('login.microsoftonline.com'));
    assert.equal(tokenRequests.length, 1);
    const chat = encodeURIComponent(GROUP_CHAT_ID);
    assert.ok(h.graph.requests.some((r) => r.startsWith(`GET https://graph.microsoft.com/v1.0/chats/${chat}/installedApps?$select=id,consentedPermissionSet&$expand=`)));
    assert.ok(h.graph.requests.includes(`GET https://graph.microsoft.com/v1.0/chats/${chat}/permissionGrants`));
  });

  it('names the missing consent and the fix when the install predates the RSC permission', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: ['TeamsActivity.Send.Chat'], version: '1.0.1' });

    const report = await h.check(MIRA);

    assert.deepEqual(
      { installed: report.installed, canReadAllMessages: report.canReadAllMessages, canPost: report.canPost, ready: report.ready },
      { installed: true, canReadAllMessages: false, canPost: true, ready: false },
    );
    assert.deepEqual(report.missing.map((m) => m.code), ['consent_missing']);
    const message = report.missing[0]?.message ?? '';
    assert.match(message, /app version 1\.0\.1/);
    assert.match(message, /ChatMessage\.Read\.Chat is not granted/);
    assert.match(message, /installedApps\/\{id\}\/upgrade/);
  });

  it('matches the grant on clientAppId only: another app\'s grant does not count', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: [] });
    h.graph.install(GROUP_CHAT_ID, JONAS);

    const report = await h.check(MIRA);

    assert.equal(report.installed, true);
    assert.equal(report.canReadAllMessages, false);
  });

  it('reports the app as not installed when the chat has no row for it', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, JONAS);

    const report = await h.check(MIRA);

    assert.equal(report.installed, false);
    assert.equal(report.canReadAllMessages, false);
    assert.deepEqual(report.missing.map((m) => m.code), ['app_not_installed']);
    assert.match(report.missing[0]?.message ?? '', /'Mira' is not installed in this chat/);
  });

  it('answers graph_forbidden with the permission the instance app needs', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA);
    h.graph.fail(GROUP_CHAT_ID, 'installedApps', 403, 'Forbidden', 'Missing role permissions on the request.');
    h.graph.fail(GROUP_CHAT_ID, 'permissionGrants', 403, 'Forbidden', 'Missing role permissions on the request.');

    const report = await h.check(MIRA);

    assert.equal(report.installed, false, 'unconfirmed is never reported as installed');
    assert.equal(report.canReadAllMessages, false);
    assert.deepEqual(report.missing.map((m) => m.code), ['graph_forbidden']);
    assert.ok(report.missing[0]?.message.includes(CHAT_INSTALL_READ_GRAPH_PERMISSION));
    assert.match(report.missing[0]?.message ?? '', /instance's app registration/);
  });

  it('answers graph_unavailable with exactly which secret or setting is missing', async () => {
    const reason =
      "the app password of the Microsoft 365 integration's app registration is not in the vault of the Teams channel under the secret name 'microsoft_app_password'";
    const h = harness({ withGraph: false, reason });
    seen(h, MIRA);

    const report = await h.check(MIRA);

    assert.deepEqual(report.missing.map((m) => m.code), ['graph_unavailable']);
    assert.equal(report.missing[0]?.message, reason);
    assert.deepEqual(h.graph.requests, []);
    const repair = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });
    assert.deepEqual(repair, { outcome: 'failed', code: 'graph_unavailable', message: reason });
  });

  it('names an install from an earlier app package as package_outdated', async () => {
    const h = harness();
    seen(h, MIRA);
    // Version scheme 1.<template revision>.<identity revision>; versions from
    // before the scheme (1.0.n) are revision 0.
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: LEGACY_CHAT_RSC, version: '1.0.4' });

    const report = await h.check(MIRA);

    assert.equal(report.ready, false);
    assert.deepEqual(
      { installed: report.installed, canReadAllMessages: report.canReadAllMessages, canPost: report.canPost },
      { installed: true, canReadAllMessages: true, canPost: true },
    );
    assert.deepEqual(report.missing.map((m) => m.code), ['package_outdated']);
    assert.match(report.missing[0]?.message ?? '', /version 1\.0\.4, template revision 0; current is 3/);
    assert.match(report.missing[0]?.message ?? '', /repair the consent/);
  });

  it('does not call an install outdated when its version says nothing about the template', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA, { version: '2024.11' });

    assert.equal((await h.check(MIRA)).ready, true);
  });

  it('uses only readiness codes the kernel contract documents', async () => {
    const documented = [
      'no_identity', 'app_not_installed', 'consent_missing', 'cannot_post', 'graph_forbidden',
      'graph_error', 'graph_unavailable', 'team_unknown', 'unsupported_conversation', 'package_outdated',
    ];
    const seenCodes = new Set<string>();
    const collect = async (run: (h: Harness) => Promise<unknown>): Promise<void> => {
      const h = harness();
      const result = (await run(h)) as { missing?: { code: string }[]; code?: string };
      for (const issue of result.missing ?? []) seenCodes.add(issue.code);
      if (result.code) seenCodes.add(result.code);
    };
    await collect((h) => h.check(MIRA, CHANNEL_ID));
    await collect((h) => h.provider.check(GROUP_CHAT_ID, { asChannelKey: '29:x' }));
    await collect((h) => h.check(MIRA));
    await collect((h) => {
      seen(h, MIRA);
      h.graph.install(GROUP_CHAT_ID, JONAS);
      return h.check(MIRA);
    });
    await collect((h) => {
      seen(h, MIRA);
      h.graph.install(GROUP_CHAT_ID, MIRA, { consented: [], version: '1.0.1' });
      return h.check(MIRA);
    });
    await collect((h) => {
      seen(h, MIRA);
      h.graph.install(GROUP_CHAT_ID, MIRA);
      h.graph.fail(GROUP_CHAT_ID, 'installedApps', 403, 'Forbidden', 'x');
      return h.check(MIRA);
    });
    await collect((h) => h.provider.repairConsent(CHANNEL_ID, { asChannelKey: MIRA.key }));
    for (const code of seenCodes) assert.ok(documented.includes(code), `undocumented readiness code: ${code}`);
    assert.ok(seenCodes.size >= 6);
  });

  it('takes "can post" from the own reference and the last send, not from a GET', async () => {
    const h = harness();
    h.graph.install(GROUP_CHAT_ID, MIRA);

    const noReference = await h.check(MIRA);
    assert.equal(noReference.canPost, false);
    assert.deepEqual(noReference.missing.map((m) => m.code), ['cannot_post']);
    assert.match(noReference.missing[0]?.message ?? '', /no conversation reference/);

    seen(h, MIRA);
    assert.equal((await h.check(MIRA)).ready, true, 'a reference and no failed send');

    h.refs.recordPostResult(GROUP_CHAT_ID, MIRA.appId, false, 'HTTP 403 BotNotInConversationRoster');
    const failed = await h.check(MIRA);
    assert.equal(failed.canPost, false);
    assert.match(failed.missing[0]?.message ?? '', /last post of 'Mira' .* failed \(HTTP 403/);

    h.refs.recordPostResult(GROUP_CHAT_ID, MIRA.appId, true);
    assert.equal((await h.check(MIRA)).canPost, true);
  });

  it('checks each bot on its own and never through another bot\'s reference', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA);
    h.graph.install(GROUP_CHAT_ID, JONAS);

    assert.equal((await h.check(MIRA)).ready, true);
    const jonas = await h.check(JONAS);
    assert.deepEqual(
      { installed: jonas.installed, canReadAllMessages: jonas.canReadAllMessages, canPost: jonas.canPost },
      { installed: true, canReadAllMessages: true, canPost: false },
    );
    assert.deepEqual(jonas.missing.map((m) => m.code), ['cannot_post']);
  });

  it('refuses an identity that is not a configured bot, without asking Graph', async () => {
    const h = harness({ configured: [MIRA] });
    assert.deepEqual((await h.check(JONAS)).missing.map((m) => m.code), ['no_identity']);
    assert.deepEqual(
      (await h.provider.check(GROUP_CHAT_ID, { asChannelKey: '29:user' })).missing.map((m) => m.code),
      ['no_identity'],
    );
    assert.deepEqual(h.graph.requests, []);
  });

  it('does not trust conversationType or the id suffix to call something a channel', async () => {
    const h = harness();
    // A 1:1 between two people with an agent app arrives as groupChat; old
    // group chats carry @thread.skype. Both are chats for Graph.
    for (const chatId of ['19:5b1e2c3d_6c2f3d4e@unq.gbl.spaces', '19:9cdb0a1b2c3d4e5f@thread.skype']) {
      seen(h, MIRA, chatId);
      h.graph.install(chatId, MIRA);
      assert.equal((await h.check(MIRA, chatId)).ready, true, chatId);
    }
  });

  it('checks a team channel against the team: installation plus ChannelMessage.Read.Group', async () => {
    const h = harness();
    seen(h, MIRA, CHANNEL_ID, {
      conversation: { id: CHANNEL_ID, conversationType: 'channel', tenantId: TENANT_ID, isGroup: true },
      channelData: { team: { id: '19:general@thread.tacv2', aadGroupId: TEAM_ID }, tenant: { id: TENANT_ID } },
    });
    h.graph.install(TEAM_ID, MIRA, { consented: ['ChannelMessage.Read.Group', 'TeamsActivity.Send.Group'] });

    assert.equal((await h.check(MIRA, CHANNEL_ID)).ready, true);
    assert.ok(h.graph.requests.includes(`GET https://graph.microsoft.com/v1.0/teams/${TEAM_ID}/permissionGrants`));

    h.graph.install(TEAM_ID, MIRA, { consented: ['TeamsActivity.Send.Group'] });
    assert.deepEqual((await h.check(MIRA, CHANNEL_ID)).missing.map((m) => m.code), ['consent_missing']);
  });

  it('still knows the team of a channel after a restart, before any new activity', async () => {
    // The reference table is all that survives a restart. A subscribed
    // channel must stay ready from it alone: no new `bot_added`, no message.
    const rows = new Map<string, { ref: Record<string, unknown>; teamsType?: string; botAppId?: string }>();
    const table = {
      save: async (conversationId: string, ref: unknown, teamsType?: string, botAppId?: string): Promise<void> => {
        // Through JSON, as the database column stores it.
        rows.set(`${botAppId ?? ''}|${conversationId}`, {
          ref: JSON.parse(JSON.stringify(ref)) as Record<string, unknown>,
          ...(teamsType ? { teamsType } : {}),
          ...(botAppId ? { botAppId } : {}),
        });
      },
      load: async (conversationId: string, botAppId?: string) => rows.get(`${botAppId ?? ''}|${conversationId}`),
    };
    const before = harness();
    before.refs.attachPersistence(table as never);
    seen(before, MIRA, CHANNEL_ID, {
      conversation: { id: CHANNEL_ID, conversationType: 'channel', tenantId: TENANT_ID, isGroup: true },
      channelData: { team: { id: '19:general@thread.tacv2', aadGroupId: TEAM_ID }, tenant: { id: TENANT_ID } },
    });
    await new Promise((resolve) => setImmediate(resolve));

    const after = harness();
    after.refs.attachPersistence(table as never);
    after.graph.install(TEAM_ID, MIRA, { consented: ['ChannelMessage.Read.Group', 'TeamsActivity.Send.Group'] });

    assert.equal(after.refs.teamIdFor(CHANNEL_ID), undefined, 'nothing is in memory in the new process');
    const report = await after.check(MIRA, CHANNEL_ID);

    assert.deepEqual(report, { installed: true, canReadAllMessages: true, canPost: true, ready: true, missing: [] });
    assert.ok(after.graph.requests.includes(`GET https://graph.microsoft.com/v1.0/teams/${TEAM_ID}/permissionGrants`));
    // The team id rides in the stored row only; the reference handed to the
    // Bot Framework is a plain one.
    const loaded = await after.refs.getOrLoad(CHANNEL_ID, MIRA.appId);
    assert.equal(Object.keys(loaded?.ref ?? {}).some((key) => key.startsWith('omadia')), false);
  });

  it('ignores a stored team id that is not an AAD group id', async () => {
    const h = harness();
    h.refs.attachPersistence({
      save: async () => {},
      load: async () => ({
        ref: {
          serviceUrl: 'https://smba.trafficmanager.net/emea/',
          conversation: { id: CHANNEL_ID },
          bot: { id: `28:${MIRA.appId}` },
          omadiaTeamAadGroupId: '../users',
        },
        teamsType: 'channel',
        botAppId: MIRA.appId,
      }),
    } as never);

    const report = await h.check(MIRA, CHANNEL_ID);

    assert.deepEqual(report.missing.map((m) => m.code), ['team_unknown']);
    assert.deepEqual(h.graph.requests, []);
  });

  it('says the team is unknown for a channel it has not heard from', async () => {
    const h = harness();
    const report = await h.check(MIRA, CHANNEL_ID);
    assert.deepEqual(report.missing.map((m) => m.code), ['team_unknown', 'cannot_post']);
    assert.deepEqual(h.graph.requests, []);
  });

  it('treats the bot\'s own 1:1 as readable without a grant lookup', async () => {
    const h = harness();
    const conversationId = 'a:1dRsHCobZ1AxURzY05Dc_personal';
    h.refs.capture({
      activity: messageActivity({ bot: MIRA, from: ANNA, text: 'hi', conversationId, conversationType: 'personal' }),
    } as unknown as TurnContext);

    assert.equal((await h.check(MIRA, conversationId)).ready, true);
    assert.deepEqual(h.graph.requests, []);
  });

  it('never throws: a transport failure is a named issue', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA);
    h.graph.fail(GROUP_CHAT_ID, 'permissionGrants', 503, 'ServiceUnavailable', 'The service is temporarily unavailable.');

    const report = await h.check(MIRA);

    assert.equal(report.installed, true);
    assert.equal(report.canReadAllMessages, false);
    assert.deepEqual(report.missing.map((m) => m.code), ['graph_error']);
  });
});

describe('#43 consent repair through the Graph upgrade', () => {
  it('upgrades the installation with the full set the published definition declares', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: ['TeamsActivity.Send.Chat'], version: '1.0.1' });
    assert.equal((await h.check(MIRA)).canReadAllMessages, false);

    const outcome = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'upgraded');
    assert.deepEqual(h.graph.upgrades.map((u) => u.permissions), [TEMPLATE_RSC]);
    assert.equal((await h.check(MIRA)).ready, true);
  });

  it('brings an older install of the same app to the full set, meeting permissions included', async () => {
    const h = harness();
    seen(h, MIRA);
    // Installed from the package before the meeting permissions.
    h.graph.publish(MIRA, LEGACY_TEMPLATE_RSC, '1.0.4');
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: LEGACY_CHAT_RSC });
    assert.deepEqual((await h.check(MIRA)).missing.map((m) => m.code), ['package_outdated']);

    // The catalog still holds the old package: nothing to upgrade to.
    const early = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });
    assert.equal(early.outcome === 'failed' ? early.code : '', 'package_outdated');
    assert.match(early.outcome === 'failed' ? early.message : '', /does not declare OnlineMeeting\.ReadBasic\.Chat/);
    assert.match(early.outcome === 'failed' ? early.message : '', /template revision 3/);
    assert.deepEqual(h.graph.upgrades, []);

    // The new package version is published to the catalog.
    h.graph.publish(MIRA, TEMPLATE_RSC, '1.3.4');
    const outcome = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'upgraded');
    assert.deepEqual(h.graph.upgrades.at(-1)?.permissions, TEMPLATE_RSC, 'the complete set, not only the new permissions');
    const chat = encodeURIComponent(GROUP_CHAT_ID);
    assert.ok(h.graph.requests.some((r) => r.startsWith(`POST https://graph.microsoft.com/v1.0/chats/${chat}/installedApps/`) && r.endsWith('/upgrade')));
    assert.equal((await h.check(MIRA)).ready, true, 'the upgraded install is at the current revision');
    assert.deepEqual(await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key }), { outcome: 'not_needed' });
  });

  it('brings an install from the revision 2 package to the definition without the file permission', async () => {
    // 0.36.1 shipped revision 2 with eleven permissions. Revision 3 drops
    // FileStorageContainer.Selected.Group. In a chat the grants do not change
    // (the dropped permission is team-scoped), so only the version tells the
    // install from a current one.
    const h = harness();
    seen(h, MIRA);
    h.graph.publish(MIRA, REVISION_2_TEMPLATE_RSC, '1.2.1');
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: CHAT_RSC });

    const before = await h.check(MIRA);
    assert.deepEqual(before.missing.map((m) => m.code), ['package_outdated']);
    assert.match(before.missing[0]?.message ?? '', /version 1\.2\.1, template revision 2; current is 3/);

    // The catalog still holds revision 2: every chat permission is declared
    // and granted, and still there is nothing to upgrade to yet.
    const early = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });
    assert.equal(early.outcome === 'failed' ? early.code : '', 'package_outdated');
    assert.match(early.outcome === 'failed' ? early.message : '', /version 1\.2\.1\) is built from template revision 2; current is 3/);
    assert.deepEqual(h.graph.upgrades, []);

    // Revision 3 is published: ten permissions.
    h.graph.publish(MIRA, TEMPLATE_RSC, '1.3.1');
    const outcome = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'upgraded');
    const sent = h.graph.upgrades.at(-1)?.permissions ?? [];
    assert.deepEqual(sent, TEMPLATE_RSC, 'exactly the set of the published definition');
    assert.equal(sent.length, 10);
    assert.ok(!sent.includes('FileStorageContainer.Selected.Group'), 'the removed permission is not consented again');
    assert.equal((await h.check(MIRA)).ready, true);
    assert.deepEqual(await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key }), { outcome: 'not_needed' });
  });

  it('says which Graph permission is missing when a repair step is refused', async () => {
    const cases: ['installedApps' | 'catalog' | 'permissionGrants' | 'upgrade', string][] = [
      ['installedApps', CHAT_INSTALL_READ_GRAPH_PERMISSION],
      ['catalog', APP_CATALOG_GRAPH_PERMISSION],
      ['permissionGrants', CHAT_INSTALL_READ_GRAPH_PERMISSION],
      ['upgrade', CHAT_CONSENT_GRAPH_PERMISSION],
    ];
    for (const [resource, permission] of cases) {
      const h = harness();
      seen(h, MIRA);
      h.graph.install(GROUP_CHAT_ID, MIRA, { consented: ['TeamsActivity.Send.Chat'] });
      h.graph.fail(GROUP_CHAT_ID, resource, 403, 'Forbidden', 'Missing role permissions on the request.');

      const outcome = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });

      assert.equal(outcome.outcome === 'failed' ? outcome.code : '', 'graph_forbidden', resource);
      assert.ok(outcome.outcome === 'failed' && outcome.message.includes(permission), `${resource}: ${permission}`);
      assert.deepEqual(h.graph.upgrades, []);
    }
  });

  it('reports a mismatch Graph answers as graph_error with the service text', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: ['TeamsActivity.Send.Chat'] });
    h.graph.fail(GROUP_CHAT_ID, 'upgrade', 400, 'BadRequest', 'ResourceSpecificPermissionsMismatch');

    const outcome = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome === 'failed' ? outcome.code : '', 'graph_error');
    assert.match(outcome.outcome === 'failed' ? outcome.message : '', /ResourceSpecificPermissionsMismatch/);
  });

  it('does nothing when every chat permission of the definition is granted', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: CHAT_RSC });

    assert.deepEqual(await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key }), { outcome: 'not_needed' });
    assert.deepEqual(h.graph.upgrades, []);
  });

  it('refuses when the published app does not declare the permission', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.publish(MIRA, ['TeamsActivity.Send.Chat'], '1.0.0');
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: [] });

    const outcome = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome === 'failed' ? outcome.code : '', 'package_outdated');
    assert.deepEqual(h.graph.upgrades, []);
  });

  it('reports a missing installation and a forbidden upgrade as outcomes', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, JONAS);
    const absent = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });
    assert.equal(absent.outcome === 'failed' ? absent.code : '', 'app_not_installed');

    h.graph.fail(GROUP_CHAT_ID, 'installedApps', 403, 'Forbidden', 'Missing role permissions on the request.');
    const forbidden = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });
    assert.equal(forbidden.outcome === 'failed' ? forbidden.code : '', 'graph_forbidden');
  });
});

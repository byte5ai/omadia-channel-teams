import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
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

function harness(opts: { configured?: SimBot[]; withGraph?: boolean } = {}): Harness {
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

  it('answers graph_unavailable when the instance credentials are not served here', async () => {
    const h = harness({ withGraph: false });
    seen(h, MIRA);

    const report = await h.check(MIRA);

    assert.deepEqual(report.missing.map((m) => m.code), ['graph_unavailable']);
    assert.deepEqual(h.graph.requests, []);
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
  it('upgrades the installation with the full set the app definition declares', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: ['TeamsActivity.Send.Chat'], declared: TEMPLATE_RSC });
    assert.equal((await h.check(MIRA)).canReadAllMessages, false);

    const outcome = await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key });

    assert.equal(outcome.outcome, 'upgraded');
    assert.deepEqual(h.graph.upgrades.map((u) => u.permissions), [TEMPLATE_RSC]);
    assert.equal((await h.check(MIRA)).ready, true);
  });

  it('does nothing when the grant is already there', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: CHAT_RSC });

    assert.deepEqual(await h.provider.repairConsent(GROUP_CHAT_ID, { asChannelKey: MIRA.key }), { outcome: 'not_needed' });
    assert.deepEqual(h.graph.upgrades, []);
  });

  it('refuses when the installed app version does not declare the permission', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.install(GROUP_CHAT_ID, MIRA, { consented: [], declared: ['TeamsActivity.Send.Chat'], version: '1.0.0' });

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

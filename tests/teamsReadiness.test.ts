import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  CHAT_READ_ALL_RSC_PERMISSION,
  createTeamsReadinessProvider,
  GRANT_READ_GRAPH_PERMISSION,
  TeamsConversationReferenceCache,
  TeamsGraphResolver,
} from '@omadia/channel-teams';

import {
  ANNA,
  connectorError,
  GROUP_CHAT_ID,
  JONAS,
  messageActivity,
  MIRA,
  rscGrant,
  SimulatedGraph,
  TENANT_ID,
  type SimBot,
} from './support/teamsSimulation.js';

// byte5ai/omadia#2043 §6.2, issue #43 — readiness: installed, RSC consent,
// can post. Graph and the Bot Connector are simulated.

interface Harness {
  graph: SimulatedGraph;
  refs: TeamsConversationReferenceCache;
  probes: string[];
  failProbeWith(err: Error | undefined): void;
  check(bot: SimBot, conversationId?: string): ReturnType<ReturnType<typeof createTeamsReadinessProvider>['check']>;
}

function harness(configured: SimBot[] = [MIRA, JONAS]): Harness {
  const graph = new SimulatedGraph();
  const refs = new TeamsConversationReferenceCache();
  const probes: string[] = [];
  let probeFailure: Error | undefined;
  const resolvers = new Map<string, TeamsGraphResolver>();
  const provider = createTeamsReadinessProvider({
    botFor: (botAppId) => {
      const bot = configured.find((candidate) => candidate.appId === botAppId);
      if (!bot) return undefined;
      const resolver =
        resolvers.get(botAppId) ??
        new TeamsGraphResolver({
          tenantId: TENANT_ID,
          clientId: bot.appId,
          clientSecret: 'simulated',
          log: () => {},
          fetchImpl: graph.fetch,
        });
      resolvers.set(botAppId, resolver);
      return { displayName: bot.name, graph: resolver };
    },
    refs,
    probeMembership: async (_reference, botAppId) => {
      probes.push(botAppId);
      if (probeFailure) throw probeFailure;
    },
  });
  return {
    graph,
    refs,
    probes,
    failProbeWith: (err) => {
      probeFailure = err;
    },
    check: (bot, conversationId = GROUP_CHAT_ID) => provider.check(conversationId, { asChannelKey: bot.key }),
  };
}

/** The bot has seen an activity from the chat, so it holds a reference. */
function seen(h: Harness, bot: SimBot, conversationId = GROUP_CHAT_ID): void {
  h.refs.capture({ activity: messageActivity({ bot, from: ANNA, text: 'hi', conversationId }) } as unknown as TurnContext);
}

describe('#43 readiness provider', () => {
  it('is ready when the app is installed, the RSC permission is granted and the bot reaches the chat', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.setGrants(GROUP_CHAT_ID, [
      rscGrant(MIRA, CHAT_READ_ALL_RSC_PERMISSION),
      rscGrant(MIRA, 'TeamsActivity.Send.Chat'),
    ]);

    const report = await h.check(MIRA);

    assert.deepEqual(report, { installed: true, canReadAllMessages: true, canPost: true, ready: true, missing: [] });
    assert.deepEqual(h.probes, [MIRA.appId]);
  });

  it('names the missing consent when the app is installed without the RSC permission', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.setGrants(GROUP_CHAT_ID, [rscGrant(MIRA, 'TeamsActivity.Send.Chat')]);

    const report = await h.check(MIRA);

    assert.equal(report.ready, false);
    assert.equal(report.installed, true);
    assert.equal(report.canPost, true);
    assert.equal(report.canReadAllMessages, false);
    assert.deepEqual(report.missing.map((m) => m.code), ['consent_missing']);
    assert.match(report.missing[0]?.message ?? '', /ChatMessage\.Read\.Chat is not granted/);
    assert.match(report.missing[0]?.message ?? '', /only receives messages that @mention it/);
  });

  it('recognises its grant whichever id field carries the app id', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.setGrants(GROUP_CHAT_ID, [
      { ...rscGrant(MIRA, CHAT_READ_ALL_RSC_PERMISSION), clientAppId: 'sp-of-mira', clientId: MIRA.appId.toUpperCase() },
    ]);

    assert.equal((await h.check(MIRA)).ready, true);
  });

  it('ignores a delegated grant of the same name', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.setGrants(GROUP_CHAT_ID, [{ ...rscGrant(MIRA, CHAT_READ_ALL_RSC_PERMISSION), permissionType: 'Delegated' }]);

    assert.equal((await h.check(MIRA)).canReadAllMessages, false);
  });

  it('does not count another app\'s grant in the same chat', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.setGrants(GROUP_CHAT_ID, [rscGrant(JONAS, CHAT_READ_ALL_RSC_PERMISSION)]);

    const report = await h.check(MIRA);

    assert.equal(report.canReadAllMessages, false);
    assert.deepEqual(report.missing.map((m) => m.code), ['consent_missing']);
  });

  it('reports the app as not installed when there is neither a grant nor a reference', async () => {
    const h = harness();
    h.graph.setGrants(GROUP_CHAT_ID, []);

    const report = await h.check(MIRA);

    assert.deepEqual(
      { installed: report.installed, canReadAllMessages: report.canReadAllMessages, canPost: report.canPost },
      { installed: false, canReadAllMessages: false, canPost: false },
    );
    assert.deepEqual(report.missing.map((m) => m.code), ['app_not_installed', 'consent_missing', 'cannot_post']);
    assert.match(report.missing[0]?.message ?? '', /'Mira' is not installed in this chat/);
    assert.deepEqual(h.probes, [], 'nothing is probed without the bot\'s own reference');
  });

  it('says the consent is unverifiable when Graph refuses the grant list', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.setGrantsError(GROUP_CHAT_ID, 403, 'Forbidden', 'Missing role permissions on the request.');

    const report = await h.check(MIRA);

    assert.equal(report.installed, true, 'the roster read proves membership');
    assert.equal(report.canPost, true);
    assert.equal(report.canReadAllMessages, false, 'unconfirmed consent is never reported as granted');
    assert.deepEqual(report.missing.map((m) => m.code), ['consent_unverifiable']);
    assert.ok(report.missing[0]?.message.includes(GRANT_READ_GRAPH_PERMISSION));
  });

  it('reports cannot_post when the bot was removed from the chat', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.setGrants(GROUP_CHAT_ID, []);
    h.failProbeWith(connectorError(403, 'BotNotInConversationRoster', 'The bot is not part of the conversation roster.'));

    const report = await h.check(MIRA);

    assert.equal(report.installed, false);
    assert.equal(report.canPost, false);
    const cannotPost = report.missing.find((m) => m.code === 'cannot_post');
    assert.match(cannotPost?.message ?? '', /HTTP 403: The bot is not part of the conversation roster/);
  });

  it('checks each bot on its own: Mira ready, Jonas not, in the same chat', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.setGrants(GROUP_CHAT_ID, [rscGrant(MIRA, CHAT_READ_ALL_RSC_PERMISSION)]);

    assert.equal((await h.check(MIRA)).ready, true);
    const jonas = await h.check(JONAS);
    assert.equal(jonas.ready, false);
    assert.deepEqual(jonas.missing.map((m) => m.code), ['app_not_installed', 'consent_missing', 'cannot_post']);
    assert.deepEqual(h.probes, [MIRA.appId], 'Jonas is never probed through Mira\'s reference');
  });

  it('refuses an identity that is not a configured bot', async () => {
    const h = harness([MIRA]);
    const unknown = await h.check(JONAS);
    assert.deepEqual(unknown.missing.map((m) => m.code), ['no_identity']);
    const notABot = await createTeamsReadinessProvider({
      botFor: () => undefined,
      refs: h.refs,
      probeMembership: async () => {},
    }).check(GROUP_CHAT_ID, { asChannelKey: '29:user' });
    assert.deepEqual(notABot.missing.map((m) => m.code), ['no_identity']);
    assert.deepEqual(h.graph.requests, [], 'no Graph call for an unknown identity');
  });

  it('does not claim readiness for a team channel', async () => {
    const h = harness();
    const report = await h.check(MIRA, '19:5f0c1a2b3c4d4e5f8a9b0c1d2e3f4a5b@thread.tacv2');
    assert.equal(report.ready, false);
    assert.deepEqual(report.missing.map((m) => m.code), ['unsupported_conversation']);
  });

  it('treats a 1:1 chat as readable without a grant lookup', async () => {
    const h = harness();
    const conversationId = 'a:1dRsHCobZ1AxURzY05Dc_personal';
    seen(h, MIRA, conversationId);
    h.refs.capture({
      activity: messageActivity({ bot: MIRA, from: ANNA, text: 'hi', conversationId, conversationType: 'personal' }),
    } as unknown as TurnContext);

    const report = await h.check(MIRA, conversationId);

    assert.equal(report.ready, true);
    assert.deepEqual(h.graph.requests, []);
  });

  it('never throws: a transport failure is reported as unverifiable', async () => {
    const h = harness();
    seen(h, MIRA);
    h.graph.setGrantsError(GROUP_CHAT_ID, 503, 'ServiceUnavailable', 'The service is temporarily unavailable.');

    const report = await h.check(MIRA);

    assert.deepEqual(report.missing.map((m) => m.code), ['consent_unverifiable']);
    assert.match(report.missing[0]?.message ?? '', /could not be read \(HTTP 503/);
  });
});

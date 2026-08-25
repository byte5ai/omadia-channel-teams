import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// #860 W0a — per-bot proactive dispatch. Runtime values via the package entry
// (messagesRouter value-imports CJS botbuilder — see messagesRouter.multibot.test.ts),
// type-only imports relative (erased by esbuild).
import { createTeamsRouter } from '@omadia/channel-teams';
import type { ConversationReference } from 'botbuilder';
import type { TeamsBotIdentity } from '../src/teamsBotIdentity.js';
import type { TeamsBotRuntime } from '../src/messagesRouter.js';
import type { TeamsBot } from '../src/teamsBot.js';

const fakeBot = { run: async () => {} } as unknown as TeamsBot;

const APP_ID_A = 'aaaaaaaa-1111-4111-8111-aaaaaaaaaaaa';
const APP_ID_B = 'bbbbbbbb-2222-4222-8222-bbbbbbbbbbbb';
const ALLOWED_SERVICE_URL = 'https://smba.trafficmanager.net/emea/';

function identity(botSlug: string, appId: string): TeamsBotIdentity {
  return {
    botSlug,
    appId,
    tenantId: 'tenant-common',
    appPasswordSecretRef: `${botSlug}_password`,
    appType: 'SingleTenant',
    displayName: botSlug,
  };
}

interface RecordedContinuation {
  botAppId: string;
  reference: Partial<ConversationReference>;
}

/** Replace a runtime adapter's `continueConversationAsync` with a recorder —
 *  the real one would try to fetch a BF token over the network. The bound
 *  `sendProactive` looks the method up on the instance at call time, so the
 *  long-lived adapter object stays shared while the continuation is faked. */
function recordContinuations(runtime: TeamsBotRuntime): RecordedContinuation[] {
  const calls: RecordedContinuation[] = [];
  (runtime.adapter as unknown as {
    continueConversationAsync: (botAppId: string, reference: Partial<ConversationReference>, logic: unknown) => Promise<void>;
  }).continueConversationAsync = async (botAppId, reference) => {
    calls.push({ botAppId, reference });
  };
  return calls;
}

function makeRouter() {
  const artifacts = createTeamsRouter({
    bot: fakeBot,
    bots: [
      { identity: identity('hr', APP_ID_A), appPassword: 'secret-a' },
      { identity: identity('accounting', APP_ID_B), appPassword: 'secret-b' },
    ],
  });
  const [runtimeA, runtimeB] = artifacts.botRuntimes;
  assert.ok(runtimeA && runtimeB);
  return {
    artifacts,
    callsA: recordContinuations(runtimeA),
    callsB: recordContinuations(runtimeB),
  };
}

function refOwnedBy(appId: string): Partial<ConversationReference> {
  return {
    serviceUrl: ALLOWED_SERVICE_URL,
    bot: { id: `28:${appId}`, name: 'bot' },
    conversation: { id: 'conv-1' } as ConversationReference['conversation'],
  };
}

const noopBuild = async () => {};

describe('sendProactive — per-bot dispatch (#860 W0a)', () => {
  it("continues the conversation with the OWNING bot's adapter and app id (ref.bot.id = 28:<appId>)", async () => {
    const { artifacts, callsA, callsB } = makeRouter();
    await artifacts.sendProactive(refOwnedBy(APP_ID_B), noopBuild);
    assert.equal(callsA.length, 0);
    assert.equal(callsB.length, 1);
    assert.equal(callsB[0]?.botAppId, APP_ID_B);
  });

  it('explicit opts.botAppId wins over the reference bot attribution', async () => {
    const { artifacts, callsA, callsB } = makeRouter();
    await artifacts.sendProactive(refOwnedBy(APP_ID_B), noopBuild, { botAppId: APP_ID_A });
    assert.equal(callsA.length, 1);
    assert.equal(callsA[0]?.botAppId, APP_ID_A);
    assert.equal(callsB.length, 0);
  });

  it('resolves the owning bot case-insensitively', async () => {
    const { artifacts, callsB } = makeRouter();
    await artifacts.sendProactive(refOwnedBy(APP_ID_B.toUpperCase()), noopBuild);
    assert.equal(callsB.length, 1);
    // continueConversationAsync gets the configured casing, not the ref's.
    assert.equal(callsB[0]?.botAppId, APP_ID_B);
  });

  it('a reference without bot attribution falls back to the default bot (legacy rows)', async () => {
    const { artifacts, callsA, callsB } = makeRouter();
    await artifacts.sendProactive(
      { serviceUrl: ALLOWED_SERVICE_URL, conversation: { id: 'conv-2' } as ConversationReference['conversation'] },
      noopBuild,
    );
    assert.equal(callsA.length, 1);
    assert.equal(callsA[0]?.botAppId, APP_ID_A);
    assert.equal(callsB.length, 0);
  });

  it('an UNKNOWN bot attribution with more than one bot configured REFUSES instead of crossing identities', async () => {
    // Cross-bot isolation (#860 W0a review): delivering another bot's
    // conversation under the default bot's credentials is the exact leak
    // this unit exists to prevent — fail loudly, appId-free message.
    const { artifacts, callsA, callsB } = makeRouter();
    await assert.rejects(
      artifacts.sendProactive(refOwnedBy('99999999-0000-4000-8000-000000000000'), noopBuild),
      (err: Error) => {
        assert.match(err.message, /not configured in teams_bots/);
        assert.ok(!err.message.includes('99999999'), 'error message must not leak the appId');
        return true;
      },
    );
    assert.equal(callsA.length, 0);
    assert.equal(callsB.length, 0);
  });

  it('an UNKNOWN bot attribution with a SINGLE configured bot falls back to it (app-registration rotation)', async () => {
    const artifacts = createTeamsRouter({
      bot: fakeBot,
      bots: [{ identity: identity('hr', APP_ID_A), appPassword: 'secret-a' }],
    });
    const calls = recordContinuations(artifacts.defaultBotRuntime);
    await artifacts.sendProactive(refOwnedBy('99999999-0000-4000-8000-000000000000'), noopBuild);
    assert.equal(calls.length, 1);
    assert.equal(calls[0]?.botAppId, APP_ID_A);
  });

  it('a user id (29:…) in bot.id is not bot attribution — default bot serves it', async () => {
    const { artifacts, callsA } = makeRouter();
    await artifacts.sendProactive(
      { serviceUrl: ALLOWED_SERVICE_URL, bot: { id: '29:user', name: 'user' } },
      noopBuild,
    );
    assert.equal(callsA.length, 1);
  });

  it('refuses a poisoned serviceUrl before ANY bot credentials are presented (isAllowedServiceUrl semantics)', async () => {
    const { artifacts, callsA, callsB } = makeRouter();
    await assert.rejects(
      artifacts.sendProactive(
        { ...refOwnedBy(APP_ID_B), serviceUrl: 'https://evil.example.com/emea/' },
        noopBuild,
      ),
      /serviceUrl outside Bot Framework domains/,
    );
    assert.equal(callsA.length, 0);
    assert.equal(callsB.length, 0);
  });

  it("per-runtime sendProactive is pre-bound to that bot — no cross-bot dispatch even for another bot's ref", async () => {
    const { artifacts, callsA, callsB } = makeRouter();
    const [runtimeA] = artifacts.botRuntimes;
    assert.ok(runtimeA);
    await runtimeA.sendProactive(refOwnedBy(APP_ID_B), noopBuild);
    // The bound sender authenticates as ITS bot; dispatch-by-ref is the
    // top-level sendProactive's job.
    assert.equal(callsA.length, 1);
    assert.equal(callsA[0]?.botAppId, APP_ID_A);
    assert.equal(callsB.length, 0);
  });

  it('legacy scalar deps keep the old behaviour: single adapter, deps.appId', async () => {
    const artifacts = createTeamsRouter({
      bot: fakeBot,
      appId: 'legacy-app-id',
      appPassword: 'legacy-secret',
      appType: 'MultiTenant',
    });
    const calls = recordContinuations(artifacts.defaultBotRuntime);
    await artifacts.sendProactive(
      { serviceUrl: ALLOWED_SERVICE_URL, conversation: { id: 'conv-3' } as ConversationReference['conversation'] },
      noopBuild,
    );
    assert.equal(calls.length, 1);
    assert.equal(calls[0]?.botAppId, 'legacy-app-id');
  });
});

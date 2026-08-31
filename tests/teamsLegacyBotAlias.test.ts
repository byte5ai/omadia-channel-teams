import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// Regression: a multi-bot deployment 401s every bot that is not `teams_bots[0]`
// on the no-slug aliases.
//
// Field incident (byte5 tenant, 2026-08-31): the deployment's ORIGINAL Azure
// bot — registered on the microsoft365 integration's app registration and
// still installed in a group chat — kept POSTing to `/api/messages`, the
// endpoint it was registered with. The moment the first agent was provisioned
// `teams_bots[]` became non-empty, `/messages` silently rebound from that bot
// to `teams_bots[0]`, and every turn died with
//
//   AuthenticationError: Unauthorized. Invalid AppId passed on token: <legacy appId>
//
// because the Bot-Framework audience claim (the legacy bot's appId) no longer
// matched the credential factory serving that path. Two promises in the code
// were broken at once:
//
//   1. messagesRouter's route comment — "no messaging endpoint registered on
//      an existing Azure bot breaks" — the alias serves ONE bot, so it breaks
//      every other bot pointed at it.
//   2. plugin.ts's shim comment — "existing single-bot deployments keep
//      working with zero changes" — the legacy scalar identity is DISCARDED
//      as soon as `teams_bots[]` is set, so its credentials are gone.
//
// The fix is per-bot on both counts: the aliases dispatch on the activity's
// own bot attribution (`recipient.id` = `28:<appId>`), and the legacy scalar
// identity is MERGED into `teams_bots[]` instead of being an either/or.
import {
  createTeamsRouter,
  legacyTeamsBotFromScalars,
  mergeLegacyTeamsBot,
  LEGACY_TEAMS_BOT_SLUG,
} from '@omadia/channel-teams';
import type { TeamsBotIdentity } from '../src/teamsBotIdentity.js';
import type { TeamsBotRuntime, TeamsRouterArtifacts } from '../src/messagesRouter.js';
import type { TeamsBot } from '../src/teamsBot.js';

const fakeBot = { run: async () => {} } as unknown as TeamsBot;

/** `teams_bots[0]` — the provisioned agent that became the default. */
const APP_ID_DEFAULT = 'aaaaaaaa-1111-4111-8111-aaaaaaaaaaaa';
/** A second provisioned agent. */
const APP_ID_SECOND = 'bbbbbbbb-2222-4222-8222-bbbbbbbbbbbb';
/** The deployment's ORIGINAL bot — the one that owns `/api/messages`. */
const APP_ID_LEGACY = 'cccccccc-3333-4333-8333-cccccccccccc';

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

/** Replace a runtime adapter's `process` with a recorder — the real one would
 *  run Bot-Framework JWT authentication against the request. */
function recordProcessedTurns(runtime: TeamsBotRuntime): number[] {
  const calls: number[] = [];
  (runtime.adapter as unknown as {
    process: (req: unknown, res: { end: () => void }, logic: unknown) => Promise<void>;
  }).process = async (_req, res) => {
    calls.push(1);
    res.end();
  };
  return calls;
}

type DispatchResult =
  | { kind: 'handled'; status: number; body?: unknown }
  | { kind: 'fallthrough' };

/** Drive the Router directly as the middleware function it is. `body` is the
 *  parsed activity the kernel's JSON body parser has already put on the
 *  request by the time the channel router runs. */
function dispatch(
  artifacts: TeamsRouterArtifacts,
  method: string,
  url: string,
  body?: unknown,
): Promise<DispatchResult> {
  return new Promise((resolve, reject) => {
    const req = { method, url, headers: {}, body };
    const res = {
      statusCode: 200,
      setHeader: () => res,
      status: (code: number) => {
        res.statusCode = code;
        return res;
      },
      json: (b: unknown) => {
        resolve({ kind: 'handled', status: res.statusCode, body: b });
        return res;
      },
      end: () => {
        resolve({ kind: 'handled', status: res.statusCode });
        return res;
      },
    };
    const next = (err?: unknown) => {
      if (err) reject(err instanceof Error ? err : new Error(String(err)));
      else resolve({ kind: 'fallthrough' });
    };
    (artifacts.router as unknown as (rq: unknown, rs: unknown, nx: unknown) => void)(
      req,
      res,
      next,
    );
  });
}

/** Three bots: the provisioned default, a second agent, and the legacy bot
 *  whose Azure registration still points at the no-slug alias. */
function makeFixture() {
  const artifacts = createTeamsRouter({
    bot: fakeBot,
    bots: [
      { identity: identity('hr', APP_ID_DEFAULT), appPassword: 'secret-default' },
      { identity: identity('accounting', APP_ID_SECOND), appPassword: 'secret-second' },
      { identity: identity(LEGACY_TEAMS_BOT_SLUG, APP_ID_LEGACY), appPassword: 'secret-legacy' },
    ],
  });
  const [rDefault, rSecond, rLegacy] = artifacts.botRuntimes;
  assert.ok(rDefault && rSecond && rLegacy);
  return {
    artifacts,
    turnsDefault: recordProcessedTurns(rDefault),
    turnsSecond: recordProcessedTurns(rSecond),
    turnsLegacy: recordProcessedTurns(rLegacy),
  };
}

/** A Teams activity as the channel receives it — `recipient` is the BOT the
 *  activity is addressed to, which is exactly the identity whose credentials
 *  must authenticate it. */
function activityFor(appId: string): unknown {
  return {
    type: 'message',
    text: 'hallo',
    recipient: { id: `28:${appId}`, name: 'a bot' },
    conversation: { id: '19:whatever@thread.skype' },
  };
}

describe('no-slug aliases dispatch per bot, not to a global default', () => {
  it("routes /messages to the bot the activity is ADDRESSED to, not to teams_bots[0]", async () => {
    const { artifacts, turnsDefault, turnsSecond, turnsLegacy } = makeFixture();

    const result = await dispatch(
      artifacts,
      'POST',
      '/messages',
      activityFor(APP_ID_LEGACY),
    );

    assert.deepEqual(result, { kind: 'handled', status: 200 });
    // The legacy bot's own credential factory must see this turn. Before the
    // fix the default bot's adapter got it and the Bot Framework rejected the
    // token with "Invalid AppId passed on token".
    assert.equal(turnsLegacy.length, 1, 'legacy bot must serve its own turn');
    assert.equal(turnsDefault.length, 0, 'default bot must NOT see another bot’s turn');
    assert.equal(turnsSecond.length, 0);
  });

  it('routes /teams/messages (documented alias) the same way', async () => {
    const { artifacts, turnsDefault, turnsSecond, turnsLegacy } = makeFixture();

    const result = await dispatch(
      artifacts,
      'POST',
      '/teams/messages',
      activityFor(APP_ID_SECOND),
    );

    assert.deepEqual(result, { kind: 'handled', status: 200 });
    assert.equal(turnsSecond.length, 1);
    assert.equal(turnsDefault.length, 0);
    assert.equal(turnsLegacy.length, 0);
  });

  it('matches the recipient app id case-insensitively (Azure vs operator casing)', async () => {
    const { artifacts, turnsDefault, turnsLegacy } = makeFixture();

    await dispatch(
      artifacts,
      'POST',
      '/messages',
      activityFor(APP_ID_LEGACY.toUpperCase()),
    );

    assert.equal(turnsLegacy.length, 1);
    assert.equal(turnsDefault.length, 0);
  });

  it('falls back to the default bot when the activity carries no bot attribution', async () => {
    const { artifacts, turnsDefault, turnsSecond, turnsLegacy } = makeFixture();

    // No body at all (the pre-existing contract), and a body without a
    // recipient — both keep the historical default-bot behaviour.
    await dispatch(artifacts, 'POST', '/messages');
    await dispatch(artifacts, 'POST', '/messages', { type: 'message' });
    await dispatch(artifacts, 'POST', '/messages', { recipient: { id: '' } });

    assert.equal(turnsDefault.length, 3);
    assert.equal(turnsSecond.length, 0);
    assert.equal(turnsLegacy.length, 0);
  });

  it('falls back to the default bot for an UNKNOWN attribution (app-registration rotation)', async () => {
    const { artifacts, turnsDefault, turnsLegacy } = makeFixture();

    // An appId nobody is configured under: the default bot's adapter answers
    // and the Bot Framework returns its own 401. Refusing here instead would
    // turn a credential-rotation window into a silent black hole.
    await dispatch(
      artifacts,
      'POST',
      '/messages',
      activityFor('dddddddd-4444-4444-8444-dddddddddddd'),
    );

    assert.equal(turnsDefault.length, 1);
    assert.equal(turnsLegacy.length, 0);
  });

  it('ignores a NON-bot recipient id (`29:` user keys never select an adapter)', async () => {
    const { artifacts, turnsDefault, turnsLegacy } = makeFixture();

    await dispatch(artifacts, 'POST', '/messages', {
      type: 'message',
      recipient: { id: '29:some-user-key' },
    });

    assert.equal(turnsDefault.length, 1);
    assert.equal(turnsLegacy.length, 0);
  });

  it('leaves the per-slug route untouched — it still ignores the body', async () => {
    const { artifacts, turnsDefault, turnsSecond, turnsLegacy } = makeFixture();

    // The slug IS the attribution on this route; a mismatching recipient must
    // not redirect the turn to another bot's credentials.
    await dispatch(
      artifacts,
      'POST',
      '/teams/accounting/messages',
      activityFor(APP_ID_LEGACY),
    );

    assert.equal(turnsSecond.length, 1);
    assert.equal(turnsLegacy.length, 0);
    assert.equal(turnsDefault.length, 0);
  });
});

describe('legacy scalar identity is merged into teams_bots[], not discarded', () => {
  const legacy = legacyTeamsBotFromScalars({
    appId: APP_ID_LEGACY,
    tenantId: 'tenant-legacy',
  });

  it('appends the legacy bot so a provisioned deployment keeps serving it', () => {
    const configured = [identity('hr', APP_ID_DEFAULT), identity('accounting', APP_ID_SECOND)];

    const merged = mergeLegacyTeamsBot(configured, legacy);

    assert.equal(merged.length, 3);
    // Appended, never prepended: `teams_bots[0]` stays the documented default
    // bot (channel-directory label, proactive-send fallback).
    assert.equal(merged[0]?.appId, APP_ID_DEFAULT);
    assert.equal(merged[2]?.appId, APP_ID_LEGACY);
  });

  it('is a no-op when the legacy appId is already a configured bot', () => {
    const configured = [identity('hr', APP_ID_LEGACY), identity('accounting', APP_ID_SECOND)];

    const merged = mergeLegacyTeamsBot(configured, legacy);

    assert.deepEqual(merged, configured);
  });

  it('matches an already-configured legacy appId case-insensitively', () => {
    const configured = [identity('hr', APP_ID_LEGACY.toUpperCase())];

    assert.deepEqual(mergeLegacyTeamsBot(configured, legacy), configured);
  });

  it('does not append when the operator already took the legacy slug', () => {
    // Appending would throw `duplicate botSlug` at router construction and
    // take the whole channel down — the operator's explicit config wins.
    const configured = [identity(LEGACY_TEAMS_BOT_SLUG, APP_ID_DEFAULT)];

    assert.deepEqual(mergeLegacyTeamsBot(configured, legacy), configured);
  });

  it('returns the configured list unchanged when there is no legacy identity', () => {
    const configured = [identity('hr', APP_ID_DEFAULT)];

    assert.deepEqual(mergeLegacyTeamsBot(configured, undefined), configured);
  });

  it('yields the legacy bot alone when nothing is configured (single-bot deployment)', () => {
    assert.deepEqual(mergeLegacyTeamsBot([], legacy), [legacy]);
  });
});

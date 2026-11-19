import { afterEach, beforeEach, describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { TeamsInfo, type TurnContext } from 'botbuilder';

import {
  createTeamsRosterAdapter,
  TeamsConversationReferenceCache,
  TeamsRosterProvider,
} from '@omadia/channel-teams';

import {
  ANNA,
  BERT,
  CARLA,
  GROUP_CHAT_ID,
  messageActivity,
  MIRA,
  SimulatedConnector,
  type SimUser,
} from './support/teamsSimulation.js';

/**
 * Kernel spec B §5.5, byte5ai/omadia#2091, issue #43 — ROSTER STATUS.
 *
 * A roster answer says how complete it is (`complete`, `partial`, `unknown`)
 * and when it was read; `fresh: true` bypasses the cache. An unreadable
 * roster is `unknown`, never an empty list: the first live facilitation
 * listed no participants because a failed read came back as "nobody".
 *
 * The Bot Connector's member list is simulated page by page; the real
 * `TeamsRosterProvider` and the real SDK adapter run on top of it.
 */

const original = TeamsInfo.getPagedMembers;
const originalError = console.error;

type Page = { members: Record<string, unknown>[]; continuationToken?: string } | Error;

function member(user: SimUser, extra: Record<string, unknown> = {}): Record<string, unknown> {
  return { id: user.id, aadObjectId: user.aadObjectId, name: user.name, ...extra };
}

function refused(statusCode: number, message: string): Error {
  return Object.assign(new Error(message), { statusCode });
}

/** Each call to the member list takes the next page from `script`; a script
 *  entry that is an `Error` is thrown. The last entry repeats. */
function memberList(script: Page[]): { calls: () => number; set(next: Page[]): void } {
  let pages = script;
  let index = 0;
  let calls = 0;
  (TeamsInfo as unknown as { getPagedMembers: unknown }).getPagedMembers = async () => {
    calls += 1;
    const page = pages[Math.min(index, pages.length - 1)];
    index += 1;
    if (page instanceof Error) throw page;
    return page;
  };
  return {
    calls: () => calls,
    set(next) {
      pages = next;
      index = 0;
    },
  };
}

function harness(): {
  clock: { now: number };
  adapter: ReturnType<typeof createTeamsRosterAdapter>;
  ask(opts?: { fresh?: boolean }): ReturnType<ReturnType<typeof createTeamsRosterAdapter>['getRoster']>;
} {
  const clock = { now: Date.parse('2026-11-19T09:00:00.000Z') };
  const refs = new TeamsConversationReferenceCache();
  refs.capture({ activity: messageActivity({ bot: MIRA, from: ANNA, text: 'hi' }) } as unknown as TurnContext);
  const connector = new SimulatedConnector();
  const adapter = createTeamsRosterAdapter({
    refs,
    roster: new TeamsRosterProvider({ now: () => clock.now }),
    sendProactive: connector.sendProactive,
  });
  return { clock, adapter, ask: (opts) => adapter.getRoster(GROUP_CHAT_ID, { asChannelKey: MIRA.key, ...opts }) };
}

const names = (roster: { participants: readonly { userRef: { displayName?: string } }[] } | undefined): (string | undefined)[] =>
  (roster?.participants ?? []).map((participant) => participant.userRef.displayName);

const MINUTE = 60_000;

describe('#43 / #2091 roster status', () => {
  beforeEach(() => {
    console.error = () => {};
  });

  afterEach(() => {
    console.error = originalError;
    (TeamsInfo as unknown as { getPagedMembers: unknown }).getPagedMembers = original;
  });

  it('complete: every page read; states when it was read', async () => {
    memberList([{ members: [member(ANNA)], continuationToken: 'p2' }, { members: [member(BERT)] }]);
    const h = harness();

    const roster = await h.ask();

    assert.equal(roster?.status, 'complete');
    assert.equal(roster?.partial, false);
    assert.deepEqual(names(roster), [ANNA.name, BERT.name]);
    assert.equal(roster?.readAt, '2026-11-19T09:00:00.000Z');
    assert.equal(roster?.statusDetail, undefined);
  });

  it('unknown: a failing roster call is never an empty list presented as the roster', async () => {
    memberList([refused(403, 'Forbidden: the app lacks ChatMember.Read.Chat in this chat')]);
    const h = harness();

    const roster = await h.ask();

    assert.equal(roster?.status, 'unknown');
    assert.equal(roster?.partial, true, 'the older flag says the same to a kernel that does not know `status`');
    assert.deepEqual(roster?.participants, []);
    assert.equal(roster?.statusDetail, 'HTTP 403: Forbidden: the app lacks ChatMember.Read.Chat in this chat');
    assert.equal(roster?.readAt, '2026-11-19T09:00:00.000Z', 'the time of the attempt');
    assert.equal(roster?.self?.userRef.id, MIRA.key, 'who asked is still known');
  });

  it('partial: pages read before a failure are kept and marked, not discarded', async () => {
    memberList([{ members: [member(ANNA), member(BERT)], continuationToken: 'p2' }, refused(429, 'Too many requests')]);
    const h = harness();

    const roster = await h.ask();

    assert.equal(roster?.status, 'partial');
    assert.equal(roster?.partial, true);
    assert.deepEqual(names(roster), [ANNA.name, BERT.name], 'what was read');
    assert.equal(roster?.statusDetail, 'HTTP 429: Too many requests');
  });

  it('unknown: a read that comes back without any member is not "nobody is here"', async () => {
    memberList([{ members: [] }]);
    const h = harness();

    const roster = await h.ask();

    assert.equal(roster?.status, 'unknown');
    assert.equal(roster?.partial, true);
    assert.match(roster?.statusDetail ?? '', /no members/);
  });

  it('a cached answer keeps the time of its read', async () => {
    const list = memberList([{ members: [member(ANNA)] }]);
    const h = harness();

    const first = await h.ask();
    h.clock.now += 2 * MINUTE;
    const second = await h.ask();

    assert.equal(list.calls(), 1, 'within five minutes the cache answers');
    assert.equal(second?.status, 'complete');
    assert.equal(second?.readAt, first?.readAt, 'the age is visible: readAt is the original read');
  });

  it('fresh: a join within the cache lifetime is seen by the fresh query', async () => {
    const list = memberList([{ members: [member(ANNA), member(BERT)] }]);
    const h = harness();
    assert.deepEqual(names(await h.ask()), [ANNA.name, BERT.name]);

    // Carla joins one minute later; the membership event was missed.
    h.clock.now += MINUTE;
    list.set([{ members: [member(ANNA), member(BERT), member(CARLA)] }]);

    const cached = await h.ask();
    const fresh = await h.ask({ fresh: true });

    assert.deepEqual(names(cached), [ANNA.name, BERT.name], 'the cache does not know her yet');
    assert.deepEqual(names(fresh), [ANNA.name, BERT.name, CARLA.name]);
    assert.equal(fresh?.status, 'complete');
    assert.equal(fresh?.readAt, '2026-11-19T09:01:00.000Z', 'read now');
    assert.equal(list.calls(), 2);
    assert.deepEqual(names(await h.ask()), [ANNA.name, BERT.name, CARLA.name], 'and the cache holds the fresh answer afterwards');
  });

  it('fresh: a failing fresh query answers unknown and does not fall back to the cached list', async () => {
    const list = memberList([{ members: [member(ANNA), member(BERT)] }]);
    const h = harness();
    await h.ask();

    h.clock.now += MINUTE;
    list.set([refused(503, 'Service unavailable')]);
    const fresh = await h.ask({ fresh: true });

    assert.equal(fresh?.status, 'unknown', 'the engine must not decide on an old list it asked to refresh');
    assert.deepEqual(fresh?.participants, []);
    assert.equal(fresh?.statusDetail, 'HTTP 503: Service unavailable');
  });

  it('a partial or unknown read is kept 30 seconds, not five minutes', async () => {
    const list = memberList([refused(503, 'Service unavailable')]);
    const h = harness();
    assert.equal((await h.ask())?.status, 'unknown');

    // Still failing-quiet inside 30 s: the Bot Connector is not hammered.
    h.clock.now += 20_000;
    list.set([{ members: [member(ANNA)] }]);
    assert.equal((await h.ask())?.status, 'unknown');
    assert.equal(list.calls(), 1);

    // After 30 s the next ordinary query reads again.
    h.clock.now += 15_000;
    const later = await h.ask();
    assert.equal(later?.status, 'complete');
    assert.deepEqual(names(later), [ANNA.name]);
    assert.equal(list.calls(), 2);
  });

  it('carries the address fields the member list returns, and invents none', async () => {
    memberList([
      {
        members: [
          member(ANNA, { email: 'anna.koehler@example.org', userPrincipalName: 'anna.koehler@example.org' }),
          // A guest: the platform sends a UPN in the #EXT# form and no email.
          member(BERT, { userPrincipalName: 'bert.maier_partner.example#EXT#@example.onmicrosoft.com' }),
          // Nothing but id and name.
          member(CARLA),
        ],
      },
    ]);
    const h = harness();

    const roster = await h.ask();

    assert.deepEqual(
      roster?.participants.map((p) => [p.userRef.displayName, (p.userRef as { email?: string }).email, p.userPrincipalName]),
      [
        [ANNA.name, 'anna.koehler@example.org', 'anna.koehler@example.org'],
        [BERT.name, undefined, 'bert.maier_partner.example#EXT#@example.onmicrosoft.com'],
        [CARLA.name, undefined, null],
      ],
    );
    assert.deepEqual(roster?.participants.map((p) => p.userRef.id), [ANNA.aadObjectId, BERT.aadObjectId, CARLA.aadObjectId]);
  });

  it('callers that need a list or nothing still get nothing for an incomplete read', async () => {
    // Mention resolution and the room's audience for member-scoped memory
    // read "empty" as "unknown room"; they must never work on a subset.
    memberList([{ members: [member(ANNA)], continuationToken: 'p2' }, refused(429, 'Too many requests')]);
    const provider = new TeamsRosterProvider();
    const context = { activity: { conversation: { id: GROUP_CHAT_ID }, recipient: { id: MIRA.key } } } as unknown as TurnContext;

    assert.deepEqual(await provider.list(context), []);
    const read = await provider.read(context);
    assert.equal(read.status, 'partial');
    assert.equal(read.participants.length, 1);
    assert.equal(read.fromCache, true);
  });
});

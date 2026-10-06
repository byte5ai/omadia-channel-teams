/**
 * The roster as the room's audience for member-scoped memory: a short-lived
 * read refetches, and bots are marked so the kernel does not count them as
 * people present.
 */
import { afterEach, describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { TeamsInfo, type TurnContext } from 'botbuilder';

import { TeamsRosterProvider } from '@omadia/channel-teams';

const original = TeamsInfo.getPagedMembers;

function stubMembers(pages: Array<Array<Record<string, unknown>>>): { calls: () => number } {
  let calls = 0;
  (TeamsInfo as unknown as { getPagedMembers: unknown }).getPagedMembers = async () => {
    const members = pages[Math.min(calls, pages.length - 1)];
    calls += 1;
    return { members };
  };
  return { calls: () => calls };
}

const context = { activity: { conversation: { id: 'conv-1' } } } as unknown as TurnContext;

const marcel = { id: '29:marcel', aadObjectId: 'aad-marcel', name: 'Marcel' };
const chris = { id: '29:chris', aadObjectId: 'aad-chris', name: 'Chris' };

describe('TeamsRosterProvider — the roster as audience', () => {
  afterEach(() => {
    (TeamsInfo as unknown as { getPagedMembers: unknown }).getPagedMembers = original;
  });

  it('serves the cache within the TTL when no max age is given', async () => {
    const stub = stubMembers([[marcel], [marcel, chris]]);
    const roster = new TeamsRosterProvider();
    await roster.list(context);
    const second = await roster.list(context);
    assert.equal(stub.calls(), 1);
    assert.equal(second.length, 1);
  });

  it('refetches once the entry is older than maxAgeMs, so a newcomer is seen', async () => {
    const stub = stubMembers([[marcel], [marcel, chris]]);
    const roster = new TeamsRosterProvider();
    await roster.list(context);
    const second = await roster.list(context, { maxAgeMs: 0 });
    assert.equal(stub.calls(), 2);
    assert.deepEqual(
      second.map((p) => p.aadObjectId),
      ['aad-marcel', 'aad-chris'],
    );
  });

  it('marks bot ids (28:) as agents and leaves people unmarked', async () => {
    stubMembers([[marcel, { id: '28:bot-app', name: 'Omadia' }]]);
    const participants = await new TeamsRosterProvider().list(context);
    assert.equal(participants.find((p) => p.channelUserId === '29:marcel')?.kind, undefined);
    assert.equal(participants.find((p) => p.channelUserId === '28:bot-app')?.kind, 'agent');
  });
});

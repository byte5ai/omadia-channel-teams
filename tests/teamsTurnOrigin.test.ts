import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { teamsSessionScope, teamsTurnOrigin } from '@omadia/channel-teams';

/**
 * W5 memory-ACL (#870 §4, Teams row) — the Teams channel states WHERE a turn
 * came from so the kernel can scope `/memories/` to that chat context.
 *
 * ## What this suite is actually guarding
 *
 * The kernel's `memoryAxesForOrigin` decides which memory tiers a turn reaches;
 * `middleware/test/memoryAxesForOrigin.test.ts` proves that decision against the
 * §2 table. It proves it for *literal shapes written by hand in that file*. This
 * suite is the other half of the seam: that the shapes this plugin really emits
 * are those shapes. Neither suite alone catches a producer that builds a correct
 * origin for the wrong Teams context.
 *
 * So the assertions are deliberately about the mapping, not about the tiers:
 *
 *  - a 1:1 chat produces a `personal` scope       → user tier
 *  - a group chat produces a conversation scope   → channel tier, no team
 *  - a team channel adds the `team` container     → channel tier + team tier
 *
 * The kernel-side consequence of each is one row of the §2 table and is asserted
 * there.
 *
 * ## Why the fail-closed cases assert an `unscoped` scope rather than "no origin"
 *
 * The design puts the fail-closed decision in exactly one place — the SDK — and
 * the producer's job is to state the context truthfully rather than to
 * pre-judge it. A turn with no usable conversation therefore still gets an
 * origin; what makes it harmless is that its scope is one the SDK refuses to
 * derive a context from.
 */
describe('W5 teamsTurnOrigin — 1:1 chat', () => {
  it('names the PERSON, not the chat window', () => {
    // The user tier has to follow the human: the same person writing to the same
    // agent must reach one tree. `sessionScope` cannot say that — it stays
    // `teams-<conversationId>` on purpose (#575 D7).
    const origin = teamsTurnOrigin(
      { conversation: { conversationType: 'personal' } },
      'teams-a:1on1@thread.tacv2',
      'aad-oid-1',
    );

    assert.deepEqual(origin.scope, { kind: 'personal', userId: 'aad-oid-1' });
    assert.equal(origin.container, undefined);
  });

  it('reaches the same tier from a second 1:1 conversation with the same person', () => {
    const first = teamsTurnOrigin(
      { conversation: { conversationType: 'personal' } },
      'teams-conv-1',
      'aad-oid-1',
    );
    const second = teamsTurnOrigin(
      { conversation: { conversationType: 'personal' } },
      'teams-conv-2',
      'aad-oid-1',
    );

    assert.deepEqual(first.scope, second.scope);
  });

  it('keeps two people in their own 1:1 trees', () => {
    const alice = teamsTurnOrigin(
      { conversation: { conversationType: 'personal' } },
      'teams-conv-1',
      'aad-oid-alice',
    );
    const bob = teamsTurnOrigin(
      { conversation: { conversationType: 'personal' } },
      'teams-conv-2',
      'aad-oid-bob',
    );

    assert.notDeepEqual(alice.scope, bob.scope);
  });

  it('falls back to the conversation scope when the user cannot be named', () => {
    // NOT a hole — a Teams 1:1 conversation id is already per-person, so the
    // turn stays isolated; it just cannot be recognised as the same person from
    // a second chat window. The alternative, `personal:`, would be one bucket
    // every anonymous 1:1 turn shares.
    const origin = teamsTurnOrigin(
      { conversation: { conversationType: 'personal' } },
      'teams-conv-1',
      undefined,
    );

    assert.deepEqual(origin.scope, { kind: 'conversation', conversationId: 'teams-conv-1' });
    assert.equal(origin.principal, undefined);
  });

  it('treats a whitespace-only user id as unnamed', () => {
    const origin = teamsTurnOrigin(
      { conversation: { conversationType: 'personal' } },
      'teams-conv-1',
      '   ',
    );

    assert.equal(origin.scope.kind, 'conversation');
  });
});

describe('W5 teamsTurnOrigin — group chat', () => {
  it('produces a conversation scope and NO container', () => {
    // A group chat has no enclosing Team, so `channelData.team` is absent and the
    // turn must reach the channel tier alone.
    const origin = teamsTurnOrigin(
      { conversation: { conversationType: 'groupChat' }, channelData: { tenant: { id: 't-1' } } },
      'teams-19:group@thread.v2',
      'aad-oid-1',
    );

    assert.deepEqual(origin.scope, {
      kind: 'conversation',
      conversationId: 'teams-19:group@thread.v2',
    });
    assert.equal(origin.container, undefined);
  });

  it('does not turn a group chat into a personal one just because a user is named', () => {
    const origin = teamsTurnOrigin(
      { conversation: { conversationType: 'groupChat' } },
      'teams-19:group@thread.v2',
      'aad-oid-1',
    );

    assert.equal(origin.scope.kind, 'conversation');
  });
});

describe('W5 teamsTurnOrigin — team channel', () => {
  it('carries the enclosing team as the container', () => {
    const origin = teamsTurnOrigin(
      {
        conversation: { conversationType: 'channel' },
        channelData: {
          team: { id: '19:team-a@thread.tacv2', name: 'Team A' },
          channel: { id: '19:general@thread.tacv2' },
          tenant: { id: 't-1' },
        },
      },
      'teams-19:general@thread.tacv2',
      'aad-oid-1',
    );

    assert.deepEqual(origin.container, { kind: 'team', id: '19:team-a@thread.tacv2' });
    assert.deepEqual(origin.scope, {
      kind: 'conversation',
      conversationId: 'teams-19:general@thread.tacv2',
    });
  });

  it('gives two channels of one team the same container and different scopes', () => {
    // The pair the design exists for, stated on the producer side: shared team
    // tier, separate channel tiers.
    const channelData = { team: { id: '19:team-a@thread.tacv2' } };
    const general = teamsTurnOrigin(
      { conversation: { conversationType: 'channel' }, channelData },
      'teams-19:general@thread.tacv2',
      'aad-oid-1',
    );
    const random = teamsTurnOrigin(
      { conversation: { conversationType: 'channel' }, channelData },
      'teams-19:random@thread.tacv2',
      'aad-oid-1',
    );

    assert.deepEqual(general.container, random.container);
    assert.notDeepEqual(general.scope, random.scope);
  });

  it('gives two teams different containers', () => {
    const a = teamsTurnOrigin(
      {
        conversation: { conversationType: 'channel' },
        channelData: { team: { id: '19:team-a@thread.tacv2' } },
      },
      'teams-19:general-a@thread.tacv2',
      'aad-oid-1',
    );
    const b = teamsTurnOrigin(
      {
        conversation: { conversationType: 'channel' },
        channelData: { team: { id: '19:team-b@thread.tacv2' } },
      },
      'teams-19:general-b@thread.tacv2',
      'aad-oid-1',
    );

    assert.notDeepEqual(a.container, b.container);
  });
});

describe('W5 teamsTurnOrigin — the container is read defensively', () => {
  // `channelData` is untyped wire data and it decides which memory tree a turn
  // may write to. Every malformed shape must resolve to "no container" rather
  // than to a container keyed on garbage that they would all then share.
  const malformed: ReadonlyArray<readonly [string, unknown]> = [
    ['absent', undefined],
    ['null', null],
    ['a string', 'team'],
    ['no team key', { tenant: { id: 't-1' } }],
    ['a null team', { team: null }],
    ['a non-object team', { team: '19:team-a' }],
    ['a team with no id', { team: { name: 'Team A' } }],
    ['a numeric team id', { team: { id: 42 } }],
    ['a blank team id', { team: { id: '' } }],
    ['a whitespace team id', { team: { id: '  ' } }],
  ];

  for (const [label, channelData] of malformed) {
    it(`drops a container for channelData that is ${label}`, () => {
      const origin = teamsTurnOrigin(
        { conversation: { conversationType: 'channel' }, channelData },
        'teams-19:general@thread.tacv2',
        'aad-oid-1',
      );

      assert.equal(origin.container, undefined);
    });
  }
});

describe('W5 teamsTurnOrigin — channel type and principal', () => {
  it('always declares itself as `teams`', () => {
    // The SDK's allowlist is keyed on this token; a different spelling would
    // silently make every Teams turn context-free.
    const origin = teamsTurnOrigin({}, 'teams-conv-1', 'aad-oid-1');
    assert.equal(origin.channelType, 'teams');
  });

  it('carries the speaker as a user principal', () => {
    const origin = teamsTurnOrigin(
      { conversation: { conversationType: 'channel' } },
      'teams-conv-1',
      'AAD-OID-1',
    );

    // Canonicalised by the SDK — the principal is an identity, not a spelling.
    assert.deepEqual(origin.principal, { kind: 'user', userId: 'aad-oid-1' });
  });

  it('omits the principal rather than inventing an unmatched one', () => {
    const origin = teamsTurnOrigin(
      { conversation: { conversationType: 'channel' } },
      'teams-conv-1',
      undefined,
    );

    assert.equal(origin.principal, undefined);
  });
});

describe('W5 teamsTurnOrigin — fail-closed inputs', () => {
  it('emits a scope the SDK refuses when the session scope is empty', () => {
    const origin = teamsTurnOrigin({ conversation: { conversationType: 'channel' } }, '', 'u-1');
    assert.deepEqual(origin.scope, { kind: 'unscoped', reason: 'absent' });
  });

  it('does not resurrect the shared bucket', () => {
    // `teams-unknown` is a known shared token; the SDK classifies it as unscoped
    // and the turn falls back to the agent-private tier of today.
    const origin = teamsTurnOrigin({}, 'teams-unknown', 'u-1');
    assert.equal(origin.scope.kind, 'unscoped');
  });

  it('a turn with no usable conversation id stays out of the shared bucket', () => {
    // `teamsSessionScope` already refuses to build `teams-` from a blank id
    // (#575 D7); the origin inherits that, so two such turns never share a tier.
    const first = teamsTurnOrigin({}, teamsSessionScope({ id: 'act-1' }), 'u-1');
    const second = teamsTurnOrigin({}, teamsSessionScope({ id: 'act-2' }), 'u-1');

    assert.notDeepEqual(first.scope, second.scope);
  });
});

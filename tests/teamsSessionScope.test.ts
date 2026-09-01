import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { teamsSessionScope } from '@omadia/channel-teams';

/**
 * #575 D7 — the Teams channel is on the SDK's typed scope resolver.
 *
 * Two things have to hold at once, and they pull against each other:
 *
 *  1. **Nothing moves.** A Teams conversation that resolves normally must keep
 *     the exact scope string it had, or its knowledge-graph partition is
 *     orphaned and the conversation loses its memory on deploy.
 *  2. **The shared bucket is gone.** A Teams activity with no usable
 *     conversation id used to land in the single literal `teams-unknown`, which
 *     every such caller shared — the `'http-default'` hole in production Teams.
 */
describe('#575 D7 — teamsSessionScope', () => {
  it('a normal conversation keeps its exact scope string — no partition moves', () => {
    assert.equal(
      teamsSessionScope({ conversation: { id: '19:abc@thread.tacv2' }, id: 'act-1' }),
      'teams-19:abc@thread.tacv2',
    );
  });

  it('the scope does not depend on the activity id when the conversation resolves', () => {
    const a = teamsSessionScope({ conversation: { id: 'conv-1' }, id: 'act-1' });
    const b = teamsSessionScope({ conversation: { id: 'conv-1' }, id: 'act-2' });
    assert.equal(a, b);
    assert.equal(a, 'teams-conv-1');
  });

  it('a MISSING conversation id no longer produces the shared bucket', () => {
    const scope = teamsSessionScope({ id: 'act-1' });
    assert.notEqual(scope, 'teams-unknown');
    assert.ok(scope.length > 0);
  });

  it('two callers who both hit the gap get two different scopes', () => {
    // This is the whole point: before, both were `teams-unknown` and shared one
    // conversation history and one graph partition.
    const a = teamsSessionScope({ id: 'act-1' });
    const b = teamsSessionScope({ id: 'act-2' });
    assert.notEqual(a, b);
  });

  it('a RETRY of the same activity resolves to the same scope', () => {
    // Isolation must not cost continuity: Bot Framework redelivers activities.
    assert.equal(teamsSessionScope({ id: 'act-1' }), teamsSessionScope({ id: 'act-1' }));
  });

  it('a literal conversation id of "unknown" is also un-shared', () => {
    // The other way the gap reaches production: Teams supplies the string
    // rather than omitting the field, which used to build `teams-unknown` too.
    const scope = teamsSessionScope({ conversation: { id: 'unknown' }, id: 'act-9' });
    assert.notEqual(scope, 'teams-unknown');
    assert.notEqual(scope, teamsSessionScope({ conversation: { id: 'unknown' }, id: 'act-8' }));
  });

  it('an EMPTY conversation id is un-shared too — the case the SDK cannot catch', () => {
    // Worth stating why this one has its own test. `teams-unknown` is already
    // in the SDK's `SHARED_SCOPE_TOKENS`, so the resolver would neutralise it
    // even if this producer handed the literal over. An empty id is different:
    // it would build `teams-`, which is NOT a known shared token, so the SDK
    // would classify it as a perfectly ordinary conversation and every caller
    // with a blank id would share it. Refusing to build a scope from a falsy
    // id is what actually closes that hole — verified by mutation: passing
    // `teams-${id ?? 'unknown'}` unconditionally leaves every OTHER assertion
    // in this file green.
    const a = teamsSessionScope({ conversation: { id: '' }, id: 'act-1' });
    const b = teamsSessionScope({ conversation: { id: '' }, id: 'act-2' });
    assert.notEqual(a, 'teams-');
    assert.notEqual(a, b);
  });

  it('with neither a conversation id nor an activity id, isolation still holds', () => {
    // Continuity is genuinely unrecoverable here; isolation is not negotiable.
    assert.notEqual(teamsSessionScope({}), teamsSessionScope({}));
  });
});

/**
 * Several bots in one chat must not share one history.
 *
 * The scope keys the conversation history, and the history was per
 * CONVERSATION. Every bot in a group chat therefore received the other bots'
 * replies as its own prior assistant turns — and continued them. Measured in
 * production: a bot that had just joined a chat logged `pool=0` (no memory of
 * its own) and `history=9` in the same turn, then answered "I am Karen" — the
 * name of the bot that had been speaking there. Minutes later the mirror
 * image, with Karen quoting the other bot's scripture.
 *
 * An identity line in the system prompt cannot win that: one sentence against
 * nine turns of transcript is persuasion, not structure.
 */
describe('per-bot conversation scope', () => {
  const CONV = '19:abc@thread.tacv2';
  const MESSIAS = '28:19ad2729-f7d3-4099-9d2a-7da1230c9533';
  const KAREN = '28:3d78d742-eefb-4fb2-bae5-3687f24c46fc';

  it('two provisioned bots in ONE chat get two different scopes', () => {
    const a = teamsSessionScope({ conversation: { id: CONV }, id: 'act-1' }, MESSIAS);
    const b = teamsSessionScope({ conversation: { id: CONV }, id: 'act-1' }, KAREN);
    assert.notEqual(a, b);
    assert.equal(a, `teams-${MESSIAS}-${CONV}`);
    assert.equal(b, `teams-${KAREN}-${CONV}`);
  });

  it('the same bot in the same chat is stable across turns', () => {
    // Isolation must not cost continuity: a bot still has to recognise its own
    // conversation on the next message.
    assert.equal(
      teamsSessionScope({ conversation: { id: CONV }, id: 'act-1' }, MESSIAS),
      teamsSessionScope({ conversation: { id: CONV }, id: 'act-9' }, MESSIAS),
    );
  });

  it('NOTHING moves when no provisioned identity resolved', () => {
    // The load-bearing half. Re-spelling a scope moves every partition behind
    // it, so a single-bot deployment — and any turn the platform could not
    // attribute to a provisioned bot — has to keep the byte-identical string
    // it had before this existed.
    assert.equal(
      teamsSessionScope({ conversation: { id: CONV }, id: 'act-1' }),
      'teams-19:abc@thread.tacv2',
    );
    assert.equal(
      teamsSessionScope({ conversation: { id: CONV }, id: 'act-1' }, undefined),
      'teams-19:abc@thread.tacv2',
    );
  });

  it('a bot key cannot rescue a conversation with no id', () => {
    // The unshared-scope guarantee outranks the qualification: without a
    // conversation there is nothing to be continuous with, and two such turns
    // must still not share a bucket.
    const a = teamsSessionScope({ conversation: { id: '' }, id: 'act-1' }, MESSIAS);
    const b = teamsSessionScope({ conversation: { id: '' }, id: 'act-2' }, MESSIAS);
    assert.notEqual(a, b);
  });
});

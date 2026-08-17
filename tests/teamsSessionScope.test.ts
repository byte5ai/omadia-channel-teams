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

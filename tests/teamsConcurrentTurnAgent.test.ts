import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
// Pure precedence logic — the relative import is bundled by the test
// transpiler and needs no built `dist/`, matching `agentForTurn.test.ts`.
import { pickChatAgentForTurn } from '../src/agentForTurn.js';

/**
 * TWO BOTS, ONE INSTANCE, OVERLAPPING TURNS.
 *
 * Teams delivers the SAME group message to EVERY bot installed in the chat,
 * and the plugin builds exactly one `TeamsBot`. So several `handleMessage`
 * calls run against one object and interleave across every `await`.
 *
 * The resolved agent used to live on that object. Its own doc-comment said the
 * quiet part: it holds "under normal load (one event-loop tick per Bot
 * Framework invocation)". A multi-bot group chat is not that load. The second
 * resolution overwrote the first before it was read, and whichever bot
 * resolved last answered for both — with its own permissions, under the other
 * one's name.
 *
 * Observed in production, in two log lines from the same turn:
 *
 *     route  28:3d78d742… → slug=hr        (Karen's bot, resolved correctly)
 *     scope  messias::teams-28-3d78d742…   (the messias agent ran it)
 *
 * These tests pin the property that makes that impossible: resolving is a
 * FUNCTION OF ITS INPUT and nothing else. Interleave two resolutions in any
 * order, and each still yields its own agent.
 */

const KAREN_BOT = '28:3d78d742-eefb-4fb2-bae5-3687f24c46fc';
const MESSIAS_BOT = '28:19ad2729-f7d3-4099-9d2a-7da1230c9533';

/** The platform's answer for each bot key, as `identityForChannel` gives it. */
function resolverFor(): (key: string) => {
  decision: 'bound' | 'fallback' | 'reject';
  chatAgent?: { slug: string };
  exclusive?: boolean;
} {
  const byKey: Record<string, string> = {
    [KAREN_BOT]: 'hr',
    [MESSIAS_BOT]: 'messias',
  };
  return (key: string) => {
    const slug = byKey[key];
    return slug
      ? { decision: 'bound' as const, chatAgent: { slug }, exclusive: true }
      : { decision: 'reject' as const };
  };
}

describe('two bots resolving in one process', () => {
  it('each bot key yields its OWN agent, whatever order they resolve in', () => {
    const resolve = resolverFor();
    const pick = (botKey: string) =>
      pickChatAgentForTurn({
        conversationId: '19:group@thread.skype',
        botKey,
        resolve: resolve as never,
      }) as unknown as { slug: string } | undefined;

    // Karen first, then Messias.
    const a1 = pick(KAREN_BOT);
    const a2 = pick(MESSIAS_BOT);
    assert.equal(a1?.slug, 'hr');
    assert.equal(a2?.slug, 'messias');

    // Messias first, then Karen — the order that used to decide the winner.
    const b1 = pick(MESSIAS_BOT);
    const b2 = pick(KAREN_BOT);
    assert.equal(b1?.slug, 'messias');
    assert.equal(b2?.slug, 'hr');
  });

  it('resolving one bot does not change what the other resolves to', () => {
    // The regression in one assertion: a second resolution must leave the
    // first one's answer untouched. When the answer lived on the shared
    // instance, this is exactly what stopped being true.
    const resolve = resolverFor();
    const karen = pickChatAgentForTurn({
      botKey: KAREN_BOT,
      resolve: resolve as never,
    }) as unknown as { slug: string } | undefined;

    pickChatAgentForTurn({ botKey: MESSIAS_BOT, resolve: resolve as never });

    assert.equal(
      karen?.slug,
      'hr',
      'Karen’s resolved agent must survive Messias resolving after her',
    );
  });
});

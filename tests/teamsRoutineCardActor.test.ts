import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// Runtime value via the package entry — `teamsBot.ts` value-imports CJS
// `botbuilder`, so the built `dist/` is what the test must exercise (same
// reason `teamsProactiveSendPerBot.test.ts` imports `createTeamsRouter` this
// way). Types come from the relative source and are erased by esbuild.
import { TeamsBot } from '@omadia/channel-teams';
import type { TurnContext } from 'botbuilder';

/**
 * byte5ai/omadia#1029 — THE CARD PATH HAS TO NAME ITS PRINCIPAL.
 *
 * `manage_routine`, the operator router and the runner are scoped by
 * `(tenant, userId)`. The routine smart-card buttons were the one remaining
 * door onto the same four mutations with no principal at all: the card carries
 * the routine id, so a replayed or hand-crafted `value` payload reached
 * `pause` / `resume` / `trigger_now` / `delete` for any row. `trigger_now` is
 * the worst of the four, because it delivers into the routine's OWN
 * `conversationRef` — a message pushed into someone else's conversation.
 *
 * Why the kernel could not just read its own per-turn context: this branch is
 * dispatched out-of-band and `return`s long before `runOrchestratorTurn`,
 * which is the single place calling `captureRoutineTurn`. Nothing installs the
 * ALS on this path, so the kernel has nothing to read and falls back to acting
 * unscoped.
 *
 * WHY THESE TESTS DRIVE THE REAL HANDLER
 * --------------------------------------
 * The first attempt at the kernel side of this fix passed while production
 * broke, because its test wrapped the call in a context wrapper production
 * never applies. So these tests construct the real `TeamsBot`, push a real
 * card-click activity through `run()`, and assert on the object the bot hands
 * to the kernel's own callback. Nothing about the principal is stubbed: if the
 * call site stops passing `actor`, the first test goes red.
 */

const TENANT = 'tenant-11111111-2222-4333-8444-555555555555';
const AAD_OBJECT_ID = 'aad-99999999-8888-4777-8666-555555555555';
const CHANNEL_USER_ID = '29:channel-scoped-user-id';
const ROUTINE_ID = 'routine-abcdef01-2345-4678-89ab-cdef01234567';

/** What the bot handed to the kernel, one entry per click. */
interface RecordedAction {
  action: string;
  id: string;
  actor?: { tenant: string; userId: string };
}

/**
 * A `TeamsBot` wired with nothing but the two things this path touches: the
 * tenant id and the routine-action callback. Every other constructor slot is
 * an optional the card branch never reaches, so it stays undefined — which is
 * also the honest shape of a routines-only deployment.
 */
function makeBot(opts: { tenantId?: string }): {
  bot: TeamsBot;
  recorded: RecordedAction[];
} {
  const recorded: RecordedAction[] = [];
  const bot = new (TeamsBot as unknown as new (...args: unknown[]) => TeamsBot)(
    // defaultOrchestrator — never invoked on the card path.
    { chat: async () => ({ text: '' }) },
    // history
    { load: async () => [], append: async () => {} },
    // topicDetector
    undefined,
    // turnContext (kernel ALS module)
    { run: async (_v: unknown, fn: () => Promise<unknown>) => fn(), enter: () => {} },
    // attachmentStore, rosterProvider, ssoConnectionName
    undefined,
    undefined,
    undefined,
    // tenantId
    opts.tenantId,
    // captureRoutineTurn — deliberately NOT called on this path; if the fix
    // ever tried to route through the ALS instead, this would stay silent and
    // the actor assertions below would fail.
    undefined,
    // handleRoutineAction — the callback under test.
    async (input: RecordedAction) => {
      recorded.push(input);
      return 'ok';
    },
  );
  return { bot, recorded };
}

/** A card click, as Teams delivers it: `value` payload, no message text. */
function cardClickContext(from: Record<string, unknown> | undefined): {
  context: TurnContext;
  sent: string[];
} {
  const sent: string[] = [];
  const context = {
    activity: {
      type: 'message',
      // `personal` keeps the mention-gate out of the way; a card action would
      // pass it anyway (`isCardAction`), so this is convenience, not a crutch.
      conversation: { id: '19:personal@thread.skype', conversationType: 'personal' },
      recipient: { id: '28:11111111-2222-4333-8444-555555555555' },
      from,
      value: { kind: 'routine.action', action: 'pause', id: ROUTINE_ID },
    },
    sendActivity: async (text: unknown) => {
      sent.push(typeof text === 'string' ? text : JSON.stringify(text));
      return undefined;
    },
    turnState: new Map<unknown, unknown>(),
  } as unknown as TurnContext;
  return { context, sent };
}

describe('#1029 routine card click carries its actor', () => {
  it('passes tenant + the aad object id the orchestrator path also uses', async () => {
    const { bot, recorded } = makeBot({ tenantId: TENANT });
    const { context, sent } = cardClickContext({
      id: CHANNEL_USER_ID,
      aadObjectId: AAD_OBJECT_ID,
      name: 'Silvio Lange',
    });

    await bot.run(context);

    assert.equal(recorded.length, 1, 'the kernel callback ran exactly once');
    const call = recorded[0];
    assert.ok(call);
    assert.equal(call.action, 'pause');
    assert.equal(call.id, ROUTINE_ID);
    assert.deepEqual(
      call.actor,
      { tenant: TENANT, userId: AAD_OBJECT_ID },
      'actor names the clicking principal, not the card',
    );
    // The ack still reaches the chat — scoping must not cost the confirmation.
    assert.deepEqual(sent, ['ok']);
  });

  it('falls back to `from.id` exactly as the orchestrator path does', async () => {
    // `handleMessage` computes `from.aadObjectId ?? from.id` once and every
    // branch uses it, including `captureRoutineTurn`. Deriving only from
    // `aadObjectId` here would mean the two doors disagree for a user Teams
    // reports without one — and once the kernel makes `actor` required, that
    // disagreement refuses a legitimate click.
    const { bot, recorded } = makeBot({ tenantId: TENANT });
    const { context } = cardClickContext({ id: CHANNEL_USER_ID, name: 'No AAD id' });

    await bot.run(context);

    assert.equal(recorded[0]?.actor?.userId, CHANNEL_USER_ID);
    assert.equal(recorded[0]?.actor?.tenant, TENANT);
  });

  it('omits actor entirely when the tenant is unknown', async () => {
    // Both halves or neither. A tenant-less actor would scope the mutation to
    // a partial principal, which is worse than the kernel's documented
    // unscoped fallback — that one at least logs and counts itself.
    const { bot, recorded } = makeBot({ tenantId: undefined });
    const { context } = cardClickContext({
      id: CHANNEL_USER_ID,
      aadObjectId: AAD_OBJECT_ID,
    });

    await bot.run(context);

    assert.equal(recorded.length, 1);
    assert.equal(
      Object.hasOwn(recorded[0] as object, 'actor'),
      false,
      'no half-filled principal, and no `actor: undefined` key either',
    );
  });

  it('omits actor entirely when Teams names no user', async () => {
    const { bot, recorded } = makeBot({ tenantId: TENANT });
    const { context } = cardClickContext(undefined);

    await bot.run(context);

    assert.equal(recorded.length, 1);
    assert.equal(Object.hasOwn(recorded[0] as object, 'actor'), false);
  });
});

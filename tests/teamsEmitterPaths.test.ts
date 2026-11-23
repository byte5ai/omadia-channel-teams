import { strict as assert } from 'node:assert';
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

// Spec 2042 criterion 5 — every user-visible message names its speaker. In
// Teams the speaker is the bot that posts (an agent's own bot, or the shared
// bot with a name line / the "System" card, INV-4), so what has to be
// enumerated is every place that posts. A new post path cannot appear without
// this table changing: the author then decides which speaker it has.

const SRC = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');
const SEND_CALL = /\b(sendActivity|sendActivities|updateActivity|createConversationAsync)\(/;

/** Post sites per source file (comments and type declarations excluded). */
const EXPECTED_SITES: Readonly<Record<string, number>> = {
  // The answer of an agent: `sendAnswer` and its helpers (card, long-answer
  // text, blocking choice/slot/consent cards), the recall card, the routine
  // list card; the pace acknowledgement and typing live in teamsPace.ts.
  // The rest are the channel's own notices (no agent, retired button,
  // approval acknowledgement, failure line): speaker `system` by origin.
  'teamsBot.ts': 31,
  'teamsPace.ts': 2,
  // The kernel's seams: posts as an agent's own bot or the shared bot,
  // stamped by the kernel (`TargetedMessage.speaker`).
  'teamsGroupPrimitives.ts': 7,
  'teamsRoutineDelivery.ts': 1,
  'teamsQuestionCardSweep.ts': 1,
  // The adapter's own failure line and the first-contact conversation.
  'messagesRouter.ts': 2,
};

function sites(): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const file of readdirSync(SRC).filter((f) => f.endsWith('.ts'))) {
    const lines = readFileSync(join(SRC, file), 'utf8').split('\n');
    for (const line of lines) {
      const t = line.trim();
      if (t.startsWith('*') || t.startsWith('//') || t.startsWith('/*')) continue;
      if (/\):\s*Promise</.test(t)) continue; // a type declaration, not a call
      if (SEND_CALL.test(t)) counts[file] = (counts[file] ?? 0) + 1;
    }
  }
  return counts;
}

describe('emitter paths (criterion 5)', () => {
  it('posts into a conversation from exactly the places this table lists', () => {
    assert.deepEqual(
      sites(),
      EXPECTED_SITES,
      'a post path was added or removed: give it a speaker (agent bot, name line or system) and update the table',
    );
  });

  it('an agent\'s answer leaves through one function: sendAnswer', () => {
    const bot = readFileSync(join(SRC, 'teamsBot.ts'), 'utf8');
    const calls = bot.match(/\bsendAnswer\(/g)?.length ?? 0;
    // the definition, the refusal notice (a system answer) and the one call of
    // `deliver` in runOrchestratorTurn (the agent's answer)
    assert.equal(calls, 3);
  });
});

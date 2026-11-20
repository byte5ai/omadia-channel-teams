import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { PaceClock } from '@omadia/channel-sdk';
import { runPacedTurn } from '@omadia/channel-sdk';
import { paceAckText } from '@omadia/channel-sdk';
import {
  RETIRED_CARD_ACTION_TEXT,
  SYSTEM_NOTICE_LABEL,
  TEAMS_PACE_PROFILE,
  teamsPaceHooks,
} from '@omadia/channel-teams';
import type { TurnContext } from 'botbuilder';

import {
  ANNA,
  cardSubmitActivity,
  messageActivity,
  MIRA,
  SimulatedKernel,
  simulatedBot,
} from './support/teamsSimulation.js';

/**
 * byte5ai/omadia-channel-teams#42 — what a turn posts into Teams: the system
 * notice, the colleagues' contributions, `/agent <name>` passed through
 * untouched, a click on a retired button, and the pace of a long turn.
 */

type Sent = Record<string, unknown>;

async function deliverAnswer(
  answer: Record<string, unknown>,
  text = 'hallo',
  conversationType: 'groupChat' | 'personal' = 'groupChat',
): Promise<{ posted: Sent[]; turns: string[] }> {
  const sim = simulatedBot(new SimulatedKernel(), {
    captureConversationReference: () => {},
    emitMembershipEvent: () => {},
    answer,
  });
  await sim.deliver(
    messageActivity({
      bot: MIRA,
      from: ANNA,
      text,
      mention: conversationType === 'groupChat',
      conversationType,
      ...(conversationType === 'personal' ? { conversationId: 'a:1personalchat' } : {}),
    }),
  );
  return { posted: (sim.sent as Sent[]).filter((activity) => activity['type'] !== 'typing'), turns: sim.turns };
}

const labelCount = (activity: Sent | undefined): number =>
  ((activity?.['entities'] as { additionalType?: string[] }[] | undefined) ?? []).filter((e) =>
    e.additionalType?.includes('AIGeneratedContent'),
  ).length;

const NOTICE = 'Mira ist deaktiviert und kann gerade nicht antworten.';
const systemAnswer = {
  text: NOTICE,
  speaker: { kind: 'system' },
  systemNotice: { code: 'agent_disabled', text: NOTICE, locale: 'de' },
  // Even if the kernel attached a disclosure, a notice is not AI output.
  aiDisclosure: { text: 'KI-Hinweis.', level: 'standard', locale: 'de', source: 'default' },
};
const JONAS = { kind: 'agent', agentId: 'a-jonas', slug: 'jonas', displayName: 'Jonas' };
const MIRA_SPEAKER = { kind: 'agent', agentId: 'a-mira', slug: 'mira', displayName: 'Mira' };

describe('system notice posted into Teams', () => {
  it('is a card under the heading "System", with no AI chip and no AI label', async () => {
    const { posted } = await deliverAnswer(systemAnswer);
    assert.equal(posted.length, 1);
    const json = JSON.stringify(posted[0]);
    assert.ok(json.includes(`"${SYSTEM_NOTICE_LABEL}"`));
    assert.ok(json.includes(NOTICE));
    assert.ok(!json.includes('✨ AI-generated'));
    assert.equal(labelCount(posted[0]), 0);
  });

  it('a system notice too long for a card still says System first', async () => {
    const long = `${NOTICE} ${'Details. '.repeat(4000)}`;
    const { posted } = await deliverAnswer({ ...systemAnswer, text: long, systemNotice: { code: 'agent_disabled', text: long, locale: 'de' } });
    assert.ok(posted.length > 1);
    assert.ok(String(posted[0]?.['text']).startsWith(`${SYSTEM_NOTICE_LABEL}: Mira ist deaktiviert`));
    assert.equal(labelCount(posted[0]), 0);
  });
});

describe('contributions posted into Teams', () => {
  const contributions = [
    { id: 'c1', speaker: JONAS, text: 'Jonas sagt: Q3 ist stabil.', status: 'delivered' },
    { id: 'c2', speaker: { ...JONAS, displayName: 'Theo' }, text: 'Theo konnte nicht antworten.', status: 'failed' },
  ];

  it('on the card: attributed to the consulted agent, text unchanged, failure shown', async () => {
    const { posted } = await deliverAnswer({ text: 'Mein Teil.', speaker: MIRA_SPEAKER, contributions });
    assert.equal(posted.length, 1);
    const json = JSON.stringify(posted[0]);
    assert.ok(json.includes('Jonas (konsultiert von Mira)'));
    assert.ok(json.includes('Jonas sagt: Q3 ist stabil.'));
    assert.ok(json.includes('Theo (konsultiert von Mira) – fehlgeschlagen'));
    assert.ok(json.includes('Mein Teil.'));
  });

  it('on the long-answer path: the same attribution in plain text', async () => {
    const long = Array.from({ length: 700 }, (_, n) => `Absatz ${String(n)}: ${'lorem ipsum '.repeat(4)}`).join('\n\n');
    const { posted } = await deliverAnswer({ text: long, speaker: MIRA_SPEAKER, contributions });
    const all = posted.map((activity) => String(activity['text'] ?? '')).join('\n');
    assert.ok(all.includes('Jonas (konsultiert von Mira):\nJonas sagt: Q3 ist stabil.'));
    assert.ok(all.includes('Theo (konsultiert von Mira) – fehlgeschlagen:\nTheo konnte nicht antworten.'));
  });

  it('an answer that is only colleagues\' contributions is still posted', async () => {
    const { posted } = await deliverAnswer({ text: '', speaker: MIRA_SPEAKER, contributions });
    assert.equal(posted.length, 1);
    assert.ok(JSON.stringify(posted[0]).includes('Jonas sagt: Q3 ist stabil.'));
  });
});

describe('what the user typed reaches the kernel as it is', () => {
  it('`/agent <name>` in a 1:1 chat is passed through unchanged: the kernel reads it, not the channel', async () => {
    const { turns } = await deliverAnswer({ text: 'ok' }, '/agent Jonas', 'personal');
    assert.deepEqual(turns, ['/agent Jonas']);
  });

  it('in a group chat the existing speaker attribution (#330) stands in front of it; the text itself is untouched', async () => {
    const { turns } = await deliverAnswer({ text: 'ok' }, '/agent Jonas');
    assert.deepEqual(turns, ['[Anna Köhler]: /agent Jonas']);
  });
});

describe('a click on a button of an older card', () => {
  it('gets one honest line instead of "no text found" or a phantom question', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
    });
    for (const type of ['fresh_check', 'verbosity_step']) {
      await sim.deliver(cardSubmitActivity({ bot: MIRA, from: ANNA, value: { type, originalMessage: 'frage', level: 'brief' }, cardMessageId: '1731930000001' }));
    }
    const texts = sim.sent
      .filter((a) => (a as Sent)['type'] !== 'typing')
      .map((a) => (typeof a === 'string' ? a : (a as Sent)['text']));
    assert.deepEqual(texts, [RETIRED_CARD_ACTION_TEXT, RETIRED_CARD_ACTION_TEXT]);
    assert.deepEqual(sim.turns, [], 'no turn was started');
  });
});

/** A manual clock, as in the kernel's channelPace test: timers fire only when time is advanced. */
function manualClock() {
  let now = 0;
  const timers: { at: number; fn: () => void; every?: number; id: number; done?: boolean }[] = [];
  let nextId = 1;
  const clock: PaceClock = {
    now: () => now,
    setTimeout: (fn, ms) => {
      const t = { at: now + ms, fn, id: nextId++ };
      timers.push(t);
      return t.id;
    },
    clearTimeout: (h) => {
      const t = timers.find((x) => x.id === h);
      if (t) t.done = true;
    },
    setInterval: (fn, ms) => {
      const t = { at: now + ms, fn, every: ms, id: nextId++ };
      timers.push(t);
      return t.id;
    },
    clearInterval: (h) => {
      const t = timers.find((x) => x.id === h);
      if (t) t.done = true;
    },
  };
  const advance = async (ms: number): Promise<void> => {
    const target = now + ms;
    for (;;) {
      const due = timers.filter((t) => !t.done && t.at <= target).sort((a, b) => a.at - b.at)[0];
      if (!due) break;
      now = due.at;
      if (due.every !== undefined) due.at += due.every;
      else due.done = true;
      due.fn();
      await Promise.resolve();
    }
    now = target;
    await Promise.resolve();
  };
  return { clock, advance };
}

describe('pace of a Teams turn (spec 2042 §5.8, criterion 10)', () => {
  function fakeContext(): { context: TurnContext; sent: unknown[] } {
    const sent: unknown[] = [];
    const context = {
      sendActivity: async (activity: unknown) => {
        sent.push(activity);
        return { id: 'x' };
      },
    } as unknown as TurnContext;
    return { context, sent };
  }

  it('typing from the start and renewed, ONE acknowledgement after the threshold, no chatter', async () => {
    const { clock, advance } = manualClock();
    const { context, sent } = fakeContext();
    let release: (v: string) => void = () => undefined;
    const pending = runPacedTurn(
      () => new Promise<string>((resolve) => (release = resolve)),
      TEAMS_PACE_PROFILE,
      teamsPaceHooks(context),
      { clock, locale: 'de' },
    );
    await advance(0);
    assert.deepEqual(sent, [{ type: 'typing' }], 'typing at the start');
    await advance(TEAMS_PACE_PROFILE.ackThresholdMs - 1);
    assert.equal(sent.filter((a) => typeof a === 'string').length, 0, 'no acknowledgement before the threshold');
    await advance(60_000);
    const texts = sent.filter((a) => typeof a === 'string');
    assert.deepEqual(texts, [paceAckText('de')], 'exactly one acknowledgement');
    assert.ok(sent.filter((a) => (a as { type?: string }).type === 'typing').length > 2, 'typing is renewed');
    release('fertig');
    const outcome = await pending;
    assert.equal(outcome.acknowledged, true);
    assert.equal(outcome.result, 'fertig');
    const before = sent.length;
    await advance(60_000);
    assert.equal(sent.length, before, 'nothing more is sent once the work is done');
  });

  it('a normal answer gets no acknowledgement at all', async () => {
    const { clock } = manualClock();
    const { context, sent } = fakeContext();
    const outcome = await runPacedTurn(async () => 'schnell', TEAMS_PACE_PROFILE, teamsPaceHooks(context), { clock });
    assert.equal(outcome.acknowledged, false);
    assert.equal(sent.filter((a) => typeof a === 'string').length, 0);
  });

  it('a failing typing or acknowledgement send never fails the turn', async () => {
    const { clock, advance } = manualClock();
    const context = {
      sendActivity: async () => {
        throw new Error('connector down');
      },
    } as unknown as TurnContext;
    let release: (v: string) => void = () => undefined;
    const pending = runPacedTurn(() => new Promise<string>((resolve) => (release = resolve)), TEAMS_PACE_PROFILE, teamsPaceHooks(context), { clock });
    await advance(TEAMS_PACE_PROFILE.ackThresholdMs + 1);
    release('ok');
    assert.equal((await pending).result, 'ok');
  });
});

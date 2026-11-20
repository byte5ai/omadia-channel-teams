import { afterEach, describe, it, mock } from 'node:test';
import { strict as assert } from 'node:assert';

import { paceAckText } from '@omadia/channel-sdk';
import { SYSTEM_NOTICE_LABEL } from '@omadia/channel-teams';

import { NO_AGENT_FOR_CONVERSATION_TEXT } from '../src/teamsBot.js';

import {
  ANNA,
  BERT,
  cardSubmitActivity,
  messageActivity,
  MIRA,
  SimulatedKernel,
  simulatedBot,
  type SimulatedResolveInput,
} from './support/teamsSimulation.js';

/**
 * byte5ai/omadia#2055 / spec 2042 §5.3, INV-7 — Teams against the management
 * layer's contract: `reject` carries the management layer's ChatAgent and a
 * `system` notice, `/agent` and `/agents` reach that ChatAgent as the whole
 * message, the sender is the one principal form of every channel, and the
 * agent that will answer is known (and acknowledges long work) before it does.
 */

type Sent = Record<string, unknown>;

const MIRA_SPEAKER = { kind: 'agent', agentId: 'a-mira', slug: 'mira', displayName: 'Mira' };
const JONAS_SPEAKER = { kind: 'agent', agentId: 'a-jonas', slug: 'jonas', displayName: 'Jonas' };
const NOTICE_TEXT = 'Mira ist nicht verfügbar. Bitte wende dich an deinen Administrator.';
const NOTICE = { code: 'agent_not_available', text: NOTICE_TEXT, locale: 'de' };
const systemAnswer = (text: string, code = 'agent_selected') => ({
  text,
  speaker: { kind: 'system' },
  systemNotice: { code, text, locale: 'de' },
});

const posted = (sent: unknown[]): Sent[] => (sent as Sent[]).filter((activity) => activity['type'] !== 'typing');
const labelCount = (activity: Sent | undefined): number =>
  ((activity?.['entities'] as { additionalType?: string[] }[] | undefined) ?? []).filter((e) =>
    e.additionalType?.includes('AIGeneratedContent'),
  ).length;

describe('reject: the management layer answers, as system', () => {
  const rejectWith = (extra: Record<string, unknown>) => () => ({ decision: 'reject', ...extra });

  it('hands the text to the management layer\'s ChatAgent and posts what it answers as a System card', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: () => ({ decision: 'reject', notice: NOTICE, chatAgent: sim.agent }),
      chat: async () => systemAnswer(NOTICE_TEXT, 'agent_not_available'),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'wie ist der Stand?', mention: true }));

    assert.deepEqual(sim.turns, ['wie ist der Stand?'], 'the text as it is, without the [Name]: attribution of an agent turn');
    const [card, ...rest] = posted(sim.sent);
    assert.equal(rest.length, 0);
    const json = JSON.stringify(card);
    assert.ok(json.includes(`"${SYSTEM_NOTICE_LABEL}"`));
    assert.ok(json.includes(NOTICE_TEXT));
    assert.ok(!json.includes('✨ AI-generated'));
    assert.equal(labelCount(card), 0, 'a notice of the management layer is not AI output');
  });

  it('a notice without a ChatAgent (a platform that sent only the notice) is posted as it is, and no agent runs', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: rejectWith({ notice: NOTICE }),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    assert.deepEqual(sim.turns, []);
    assert.ok(JSON.stringify(posted(sim.sent)[0]).includes(NOTICE_TEXT));
  });

  it('answers a rejected turn as system even when the platform\'s ChatAgent answered with agent prose', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: () => ({ decision: 'reject', notice: NOTICE, chatAgent: sim.agent }),
      chat: async () => ({ text: 'Ich helfe gern!', speaker: MIRA_SPEAKER }),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    const json = JSON.stringify(posted(sim.sent));
    assert.ok(json.includes(NOTICE_TEXT));
    assert.ok(!json.includes('Ich helfe gern'));
  });

  it('stores and starts nothing for a sender no agent serves: no history, no typing indicator', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: () => ({ decision: 'reject', notice: NOTICE, chatAgent: sim.agent }),
      chat: async () => systemAnswer(NOTICE_TEXT, 'agent_not_available'),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    assert.deepEqual(sim.appended, [], 'a refusal is not conversation content');
    assert.equal((sim.sent as Sent[]).filter((a) => a['type'] === 'typing').length, 0, 'no typing: there is no agent working');
  });

  it('fails closed with the fixed text when the platform gave neither an agent nor a notice', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: rejectWith({}),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    assert.deepEqual(sim.sent, [NO_AGENT_FOR_CONVERSATION_TEXT]);
    assert.deepEqual(sim.turns, []);
  });

  it('fails closed with the fixed text when the resolver throws, never with the default bundle', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: () => {
        throw new Error('registry down');
      },
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    assert.deepEqual(sim.sent, [NO_AGENT_FOR_CONVERSATION_TEXT]);
    assert.deepEqual(sim.turns, []);
  });

  it('a card click of a refused sender gets the notice, not a turn', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: rejectWith({ notice: NOTICE }),
    });

    await sim.deliver(
      cardSubmitActivity({ bot: MIRA, from: ANNA, value: { type: 'follow_up', prompt: 'zeig mehr' }, cardMessageId: '1731930000001' }),
    );

    assert.deepEqual(sim.turns, []);
    assert.ok(JSON.stringify(posted(sim.sent)[0]).includes(NOTICE_TEXT));
  });
});

describe('/agent and /agents reach the management layer as the whole message', () => {
  const boundTo = (speaker: unknown) => (input: SimulatedResolveInput) => ({
    decision: 'bound',
    chatAgent: undefined as unknown,
    speaker,
    exclusive: input.channelKey.startsWith('28:'),
  });

  function groupBot(chat?: () => Promise<Record<string, unknown>>) {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: (input) => ({ ...boundTo(MIRA_SPEAKER)(input), chatAgent: sim.agent }),
      ...(chat ? { chat } : {}),
    });
    return sim;
  }

  for (const command of ['/agent Jonas', '/agent', '/agents', '/agent "Jonas Meyer"']) {
    it(`"${command}" in a group chat carries no [Name]: attribution in front of it`, async () => {
      const sim = groupBot(async () => systemAnswer('ok'));

      await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: command, mention: true }));

      assert.deepEqual(sim.turns, [command], 'the whole message, as typed');
    });
  }

  it('ordinary text in a group chat keeps the attribution; /agent inside a sentence is ordinary text', async () => {
    const sim = groupBot();

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'nutze /agent Jonas bitte', mention: true }));

    assert.deepEqual(sim.turns, ['[Anna Köhler]: hallo', '[Anna Köhler]: nutze /agent Jonas bitte']);
  });

  it('the confirmation is a System card, not an answer of the agent, and not conversation history', async () => {
    const sim = groupBot(async () => systemAnswer('Jonas ist jetzt dein Ansprechpartner in diesem Chat.'));

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: '/agent Jonas', mention: true }));

    const [card, ...rest] = posted(sim.sent);
    assert.equal(rest.length, 0);
    const json = JSON.stringify(card);
    assert.ok(json.includes(`"${SYSTEM_NOTICE_LABEL}"`));
    assert.ok(json.includes('Jonas ist jetzt dein Ansprechpartner'));
    assert.equal(labelCount(card), 0);
    assert.deepEqual(sim.appended, []);
  });

  it('a command typed into a running conversation is not read as a topic break', async () => {
    const classified: string[] = [];
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      priorTurns: [{ userMessage: 'Umsatz Q3?', assistantAnswer: '1,2 Mio.' }],
      topicDetector: {
        classify: async ({ userMessage }) => {
          classified.push(userMessage);
          return { decision: 'ask', reason: 'test' };
        },
      },
      resolve: () => ({ decision: 'bound', chatAgent: sim.agent, speaker: MIRA_SPEAKER }),
      chat: async () => systemAnswer('ok'),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: '/agents', mention: true }));

    assert.deepEqual(classified, []);
    assert.deepEqual(sim.turns, ['/agents']);
  });

  it('in a 1:1 chat the text is the command as typed', async () => {
    const sim = groupBot(async () => systemAnswer('ok'));

    await sim.deliver(
      messageActivity({ bot: MIRA, from: ANNA, text: '/agents', conversationType: 'personal', conversationId: 'a:1personalchat' }),
    );

    assert.deepEqual(sim.turns, ['/agents']);
  });
});

describe('the sender principal is the one form of every channel', () => {
  const ANNA_EMAIL = 'Anna.Koehler@example.com';
  const resolveBound = () => ({ decision: 'bound', chatAgent: undefined as unknown, speaker: MIRA_SPEAKER });

  it('is user:<email> once Graph verified the address, lowercase, on every probed key', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      emailByAad: { [ANNA.aadObjectId]: ANNA_EMAIL },
      resolve: () => ({ ...resolveBound(), chatAgent: sim.agent }),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    assert.ok(sim.resolveCalls.length >= 1);
    for (const call of sim.resolveCalls) assert.equal(call.principal, 'user:anna.koehler@example.com');
    assert.deepEqual(
      new Set(sim.resolveCalls.map((call) => call.channelKey)).has(MIRA.key),
      true,
      'the bot key is probed',
    );
    assert.equal(sim.resolveCalls[0]?.channelType, 'teams');
  });

  it('is user:teams-aad:<id> when no address is verified', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: () => ({ ...resolveBound(), chatAgent: sim.agent }),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: BERT, text: 'hallo', mention: true }));

    for (const call of sim.resolveCalls) assert.equal(call.principal, `user:teams-aad:${BERT.aadObjectId}`);
  });

  it('passes the notice language of the activity (de-DE → de)', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: () => ({ ...resolveBound(), chatAgent: sim.agent }),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    for (const call of sim.resolveCalls) assert.equal(call.locale, 'de');
  });

  it('passes en for an English activity locale, and nothing for a language the notices do not know', async () => {
    const run = async (locale: string) => {
      const sim = simulatedBot(new SimulatedKernel(), {
        captureConversationReference: () => {},
        emitMembershipEvent: () => {},
        resolve: () => ({ ...resolveBound(), chatAgent: sim.agent }),
      });
      await sim.deliver({ ...messageActivity({ bot: MIRA, from: ANNA, text: 'hello', mention: true }), locale });
      return sim.resolveCalls.map((call) => call.locale);
    };

    assert.deepEqual([...new Set(await run('en-US'))], ['en']);
    assert.deepEqual([...new Set(await run('fr-FR'))], [undefined]);
  });

  it('is not asked for a message nobody addressed (no Graph lookup for a dropped message)', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: () => ({ ...resolveBound(), chatAgent: sim.agent }),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'unter uns' }));

    assert.deepEqual(sim.resolveCalls, []);
  });

  it('leaves TurnOrigin.principal as the Entra object id the facilitation engine matches participants by', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      emailByAad: { [ANNA.aadObjectId]: ANNA_EMAIL },
      resolve: () => ({ ...resolveBound(), chatAgent: sim.agent }),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    assert.deepEqual((sim.origins[0] as { principal?: unknown }).principal, { kind: 'user', userId: ANNA.aadObjectId });
  });
});

describe('the speaker is known before the answer', () => {
  afterEach(() => mock.timers.reset());

  /** One bot, two agents: a binding serves Jonas, a later selection serves Mira. */
  function sharedBot(chat: () => Promise<Record<string, unknown>>) {
    let speaker: unknown = JONAS_SPEAKER;
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: () => ({ decision: 'bound', chatAgent: sim.agent, speaker }),
      chat,
    });
    return {
      sim,
      speak: (next: unknown) => {
        speaker = next;
      },
    };
  }

  it('the pace acknowledgement is the SDK template, once, under the speaker\'s name line when one bot serves several agents', async () => {
    mock.timers.enable({ apis: ['setTimeout', 'setInterval'] });
    let slow = false;
    const { sim, speak } = sharedBot(async () => {
      if (slow) {
        mock.timers.tick(9_000);
        await Promise.resolve();
      }
      return { text: 'Fertig.', speaker: slow ? MIRA_SPEAKER : JONAS_SPEAKER };
    });
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'erste frage', mention: true }));
    speak(MIRA_SPEAKER);
    slow = true;

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'zweite frage', mention: true }));

    const acks = (sim.sent as unknown[]).filter((a) => typeof a === 'string');
    assert.deepEqual(acks, [`Mira: ${paceAckText('de')}`], 'one acknowledgement, in the new speaker\'s name');
  });

  it('on the agent\'s own bot the acknowledgement is the bare template: the bot is the speaker', async () => {
    mock.timers.enable({ apis: ['setTimeout', 'setInterval'] });
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: (input) =>
        input.channelKey.startsWith('28:')
          ? { decision: 'bound', chatAgent: sim.agent, exclusive: true, speaker: MIRA_SPEAKER }
          : { decision: 'reject' },
      chat: async () => {
        mock.timers.tick(9_000);
        await Promise.resolve();
        return { text: 'Fertig.', speaker: MIRA_SPEAKER };
      },
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'lange arbeit', mention: true }));

    assert.deepEqual((sim.sent as unknown[]).filter((a) => typeof a === 'string'), [paceAckText('de')]);
  });

  it('a name line stands over the first answer after the speaker changed, not before', async () => {
    let current: unknown = JONAS_SPEAKER;
    const { sim, speak } = sharedBot(async () => ({ text: 'Antwort.', speaker: current }));
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'eins', mention: true }));
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'zwei', mention: true }));
    current = MIRA_SPEAKER;
    speak(MIRA_SPEAKER);
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'drei', mention: true }));

    const cards = posted(sim.sent).map((activity) => JSON.stringify(activity));
    assert.equal(cards.length, 3);
    assert.ok(!cards[0]?.includes('"Jonas:"') && !cards[1]?.includes('"Jonas:"'), 'one agent in the conversation: no name line');
    assert.ok(cards[2]?.includes('"Mira:"'), 'the speaker changed: the name line says who speaks now');
  });

  it('on the agent\'s own bot no name line ever appears, the bot is the speaker', async () => {
    let current: unknown = JONAS_SPEAKER;
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: (input) =>
        input.channelKey.startsWith('28:')
          ? { decision: 'bound', chatAgent: sim.agent, exclusive: true, speaker: current }
          : { decision: 'reject' },
      chat: async () => ({ text: 'Antwort.', speaker: current }),
    });
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'eins', mention: true }));
    current = MIRA_SPEAKER;
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'zwei', mention: true }));

    for (const activity of posted(sim.sent)) {
      const json = JSON.stringify(activity);
      assert.ok(!json.includes('"Mira:"') && !json.includes('"Jonas:"'));
    }
  });

  it('the answer states its speaker from the routing decision when the kernel left it out', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: () => ({ decision: 'bound', chatAgent: sim.agent, exclusive: true, speaker: MIRA_SPEAKER }),
      answer: {
        text: 'Mein Teil.',
        contributions: [{ id: 'c1', speaker: JONAS_SPEAKER, text: 'Jonas sagt: Q3 ist stabil.', status: 'delivered' }],
      },
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'frage', mention: true }));

    assert.ok(JSON.stringify(posted(sim.sent)[0]).includes('Jonas (konsultiert von Mira)'));
  });
});

describe('/agent is stored under the narrowest key and the singleton never answers behind a registry', () => {
  it('a bot-wide catch-all answers text, but /agent reaches the conversation\'s management layer', async () => {
    const answeredBy: string[] = [];
    const agentFor = (name: string) => ({
      chat: async (input: { userMessage: string }) => {
        answeredBy.push(`${name}:${input.userMessage}`);
        return { text: 'ok', speaker: MIRA_SPEAKER };
      },
    });
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: (input) =>
        input.channelKey.startsWith('28:')
          ? { decision: 'bound', chatAgent: agentFor('bot'), speaker: MIRA_SPEAKER }
          : { decision: 'reject', chatAgent: agentFor('conversation'), notice: NOTICE },
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: '/agent Jonas', mention: true }));
    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    assert.deepEqual(answeredBy, ['conversation:/agent Jonas', 'bot:[Anna Köhler]: hallo']);
  });

  it('a bound decision without any ChatAgent is refused, not answered by the host\'s singleton', async () => {
    const sim = simulatedBot(new SimulatedKernel(), {
      captureConversationReference: () => {},
      emitMembershipEvent: () => {},
      resolve: () => ({ decision: 'bound', notice: undefined }),
    });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    assert.deepEqual(sim.sent, [NO_AGENT_FOR_CONVERSATION_TEXT]);
    assert.deepEqual(sim.turns, [], 'the singleton agent did not run');
  });

  it('without the resolver service (a host with one agent) the singleton answers', async () => {
    const sim = simulatedBot(new SimulatedKernel(), { captureConversationReference: () => {}, emitMembershipEvent: () => {} });

    await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));

    assert.deepEqual(sim.turns, ['[Anna Köhler]: hallo']);
  });
});

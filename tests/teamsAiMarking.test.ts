import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  createTeamsConversationSendAdapter,
  createTeamsRoutineDelivery,
  createTeamsTargetedSendAdapter,
  PROACTIVE_POSTS_CARRY_AI_LABEL,
  TeamsConversationReferenceCache,
} from '@omadia/channel-teams';

import {
  ANNA,
  BERT,
  GROUP_CHAT_ID,
  messageActivity,
  MIRA,
  referenceFor,
  SimulatedConnector,
  SimulatedKernel,
  simulatedBot,
} from './support/teamsSimulation.js';

/**
 * byte5ai/omadia#2043, kernel #2098 — ONE AI MARKING PER MESSAGE.
 *
 * The rule: every message under an agent's name gets the channel's AI label
 * where the channel has one, the instance's disclosure sentence otherwise;
 * once, never both. Teams has a label (the `AIGeneratedContent` entity), so
 * in Teams it is always the label and never the sentence.
 *
 * One test per send path of this adapter. The table in
 * `docs/teams-meeting-apis.md` ("AI marking per send path") names them.
 */

const DISCLOSURE = 'Diese Antwort wurde von einer KI erzeugt.';
const folded = (text: string): string => `${text}\n\n${DISCLOSURE}`;

type Sent = Record<string, unknown>;

function labelsOf(activity: Sent | undefined): unknown[] {
  const entities = (activity?.['entities'] as { additionalType?: string[] }[] | undefined) ?? [];
  return entities.filter((entity) => entity.additionalType?.includes('AIGeneratedContent'));
}

/** Exactly one label on the activity, and the sentence nowhere in it. */
function assertLabelOnly(activity: Sent | undefined, what: string): void {
  assert.ok(activity, `${what}: nothing was posted`);
  assert.equal(labelsOf(activity).length, 1, `${what}: exactly one AI label`);
  assert.ok(!JSON.stringify(activity).includes(DISCLOSURE), `${what}: the disclosure sentence must not be posted`);
}

async function turnAnswer(answer: Record<string, unknown>): Promise<Sent[]> {
  const sim = simulatedBot(new SimulatedKernel(), {
    captureConversationReference: () => {},
    emitMembershipEvent: () => {},
    answer,
  });
  await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));
  return (sim.sent as Sent[]).filter((activity) => activity['type'] !== 'typing');
}

describe('#2043 AI marking: turn answers', () => {
  it('a model answer with the folded sentence: label, sentence stripped', async () => {
    const posted = await turnAnswer({ text: folded('Die Antwort ist 42.'), aiDisclosure: { text: DISCLOSURE } });

    assert.equal(posted.length, 1);
    assertLabelOnly(posted[0], 'turn answer');
    assert.ok(JSON.stringify(posted[0]).includes('Die Antwort ist 42.'));
  });

  it('a fixed refusal of the facilitator, which now carries aiDisclosure too: label, sentence stripped', async () => {
    const refusal = 'Ich bin Mira, Moderator und kein Assistent: Fragen beantworten kann ich nicht.';
    const posted = await turnAnswer({ text: folded(refusal), aiDisclosure: { text: DISCLOSURE } });

    assert.equal(posted.length, 1);
    assertLabelOnly(posted[0], 'fixed refusal');
    assert.ok(JSON.stringify(posted[0]).includes('Moderator und kein Assistent'));
  });

  it('an answer without a disclosure (the operator turned it off): the label all the same', async () => {
    const posted = await turnAnswer({ text: 'Die Antwort ist 42.' });

    assert.equal(posted.length, 1);
    assertLabelOnly(posted[0], 'turn answer without disclosure');
  });

  it('a long answer that is posted as plain text: label on the first part, sentence stripped', async () => {
    const long = Array.from({ length: 700 }, (_, n) => `Absatz ${String(n)}: ${'lorem ipsum '.repeat(4)}`).join('\n\n');
    assert.ok(long.length > 25_000, 'longer than one Teams message');
    const posted = await turnAnswer({ text: folded(long), aiDisclosure: { text: DISCLOSURE } });

    assert.ok(posted.length > 1, 'split into several messages');
    assert.equal(labelsOf(posted[0]).length, 1, 'the label is on the first part');
    assert.equal(posted.flatMap((activity) => labelsOf(activity)).length, 1, 'and only there');
    assert.ok(!JSON.stringify(posted).includes(DISCLOSURE));
  });

  it('a NO_REPLY answer stays silent, also with the sentence folded in behind it', async () => {
    assert.deepEqual(await turnAnswer({ text: 'NO_REPLY' }), []);
    assert.deepEqual(await turnAnswer({ text: folded('NO_REPLY'), aiDisclosure: { text: DISCLOSURE } }), []);
  });

  it('a blocking choice card: label', async () => {
    const posted = await turnAnswer({
      text: folded('Welche meinst du?'),
      aiDisclosure: { text: DISCLOSURE },
      interactive: { kind: 'choice', question: 'Welche Option?', options: [{ label: 'A', value: 'a' }, { label: 'B', value: 'b' }] },
    });

    assert.equal(posted.length, 1);
    assertLabelOnly(posted[0], 'choice card');
  });
});

describe('#2043 AI marking: posts as the agent\'s own identity', () => {
  function providers(): {
    connector: SimulatedConnector;
    conversation: ReturnType<typeof createTeamsConversationSendAdapter>;
    targeted: ReturnType<typeof createTeamsTargetedSendAdapter>;
  } {
    const connector = new SimulatedConnector();
    const refs = new TeamsConversationReferenceCache();
    refs.capture({ activity: messageActivity({ bot: MIRA, from: ANNA, text: 'hi' }) } as unknown as TurnContext);
    return {
      connector,
      conversation: createTeamsConversationSendAdapter({ refs, sendProactive: connector.sendProactive }),
      targeted: createTeamsTargetedSendAdapter({
        sendProactive: connector.sendProactive,
        direct: { refs, openDirectConversation: connector.openDirectConversation, botLabel: () => MIRA.name },
      }),
    };
  }

  it('both providers declare the label, and what they post carries it: text, poll card, private reminder', async () => {
    const p = providers();
    // The kernel leaves the disclosure sentence out BECAUSE of this
    // declaration, so the text arrives without it.
    assert.equal(p.conversation.supports.aiLabel, true);
    assert.equal(p.targeted.supports?.aiLabel, true);
    assert.equal(PROACTIVE_POSTS_CARRY_AI_LABEL, true);

    await p.conversation.sendToConversation(
      GROUP_CHAT_ID,
      { text: 'Bitte wählt.', poll: { pollId: 'p1', question: 'Welche Option?', options: [{ id: 'a', label: 'A' }, { id: 'b', label: 'B' }] } },
      { asChannelKey: MIRA.key },
    );
    await p.targeted.sendToUser(
      { principalId: `user:${BERT.aadObjectId}`, userRef: { kind: 'teams-aad', id: BERT.aadObjectId } },
      { text: 'Deine Stimme fehlt noch.' },
      { asChannelKey: MIRA.key },
    );

    assert.equal(p.connector.sent.length, 3, 'text, poll card, private reminder');
    const [text, card, reminder] = p.connector.sent.map((sent) => sent.activity);
    assertLabelOnly(text, 'proactive text');
    assertLabelOnly(card, 'poll card');
    assertLabelOnly(reminder, 'private reminder');
    assert.equal(text?.['text'], 'Bitte wählt.', 'the text goes out as the kernel handed it over');
  });
});

describe('#2043 AI marking: deliveries outside a turn (routines, approvals)', () => {
  const ROUTINE = { id: 'routine-1', name: 'Morgenbericht', cron: '0 8 * * 1-5' };
  const reference = referenceFor(MIRA, GROUP_CHAT_ID);

  function delivery(withCard: boolean): { connector: SimulatedConnector; send: ReturnType<typeof createTeamsRoutineDelivery> } {
    const connector = new SimulatedConnector();
    return {
      connector,
      send: createTeamsRoutineDelivery({
        sendProactive: connector.sendProactive,
        ...(withCard
          ? {
              buildRoutineSmartCardAttachment: (input: { routine: { name: string }; body: string }) => ({
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: { type: 'AdaptiveCard', body: [{ type: 'TextBlock', text: `${input.routine.name}: ${input.body}` }] },
              }),
            }
          : {}),
      }),
    };
  }

  it('a routine delivery as a card: label, sentence stripped from the card body', async () => {
    const { connector, send } = delivery(true);

    await send(reference, { text: folded('Heute stehen drei Termine an.'), aiDisclosure: { text: DISCLOSURE } } as never, ROUTINE as never);

    assert.equal(connector.sent.length, 1);
    assertLabelOnly(connector.sent[0]?.activity, 'routine card');
    assert.ok(JSON.stringify(connector.sent[0]?.activity).includes('Morgenbericht: Heute stehen drei Termine an.'));
  });

  it('a routine delivery as plain text: label, sentence stripped', async () => {
    const { connector, send } = delivery(false);

    await send(reference, { text: folded('Heute stehen drei Termine an.'), aiDisclosure: { text: DISCLOSURE } } as never, ROUTINE as never);

    assertLabelOnly(connector.sent[0]?.activity, 'routine text');
    assert.equal(connector.sent[0]?.activity['text'], 'Heute stehen drei Termine an.');
  });

  it('a delivery without aiDisclosure (disclosure turned off, or a Conductor reminder): the label all the same', async () => {
    const { connector, send } = delivery(false);

    await send(reference, { text: 'Erinnerung: die Freigabe steht noch aus.' } as never, undefined as never);

    assertLabelOnly(connector.sent[0]?.activity, 'reminder text');
    assert.equal(connector.sent[0]?.activity['text'], 'Erinnerung: die Freigabe steht noch aus.');
  });

  it('an approval card: label only', async () => {
    const { connector, send } = delivery(true);

    await send(
      reference,
      {
        text: 'Freigabe nötig',
        approval: { awaitId: 'await-1', question: 'Angebot 4711 freigeben?', workflowName: 'Angebote', stepLabel: 'Freigabe', quorum: 'any' },
      } as never,
      undefined as never,
    );

    assert.equal(connector.sent.length, 1);
    assertLabelOnly(connector.sent[0]?.activity, 'approval card');
    assert.ok(JSON.stringify(connector.sent[0]?.activity).includes('Angebot 4711 freigeben?'));
  });

  it('a NO_REPLY routine answer is still dropped before anything is posted', async () => {
    const { connector, send } = delivery(true);

    await send(reference, { text: 'NO_REPLY', aiDisclosure: { text: DISCLOSURE } } as never, ROUTINE as never);
    await send(reference, { text: folded('NO_REPLY'), aiDisclosure: { text: DISCLOSURE } } as never, ROUTINE as never);

    assert.deepEqual(connector.sent, []);
  });
});

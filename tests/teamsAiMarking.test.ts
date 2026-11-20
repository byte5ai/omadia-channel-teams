import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { TurnContext } from 'botbuilder';

import {
  createTeamsConversationSendAdapter,
  createTeamsRoutineDelivery,
  createTeamsTargetedSendAdapter,
  isOutgoingLogEnabled,
  incomingInvokeLogLine,
  logOutgoingActivity,
  PROACTIVE_POSTS_CARRY_AI_LABEL,
  resolveNativeAiLabel,
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
 * byte5ai/omadia#2043, kernel #2098, rollout defect D24 — EVERY MESSAGE OF AN
 * AGENT CARRIES A VISIBLE AI MARKING.
 *
 * Three things exist: the platform label (the `AIGeneratedContent` entity,
 * shown by Teams as "AI Generated" next to the agent's name), the instance's
 * disclosure sentence (folded into the text by the kernel), and the
 * "✨ AI-generated" chip the answer card draws itself.
 *
 * D24: with the label declared and the sentence stripped everywhere, the
 * facilitator's proactive posts showed no marking in the client. The label
 * that had been seen on turn answers was the card's own chip. So the label
 * counts only once it was SEEN in a client (`teams_native_ai_label`, default
 * off). Off: the sentence stays wherever no card chip is shown. On: label
 * only, as 0.36.6 did.
 *
 * One test per send path and mode; the table in `docs/teams-meeting-apis.md`
 * ("AI marking per send path") names them. What a Teams client RENDERS is
 * not verified by any of these tests.
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

/** Exactly one label entity, and the sentence exactly once, at the end. */
function assertLabelAndSentence(activities: readonly (Sent | undefined)[], what: string): void {
  const json = JSON.stringify(activities);
  assert.equal(activities.flatMap((activity) => labelsOf(activity)).length, 1, `${what}: exactly one AI label entity`);
  assert.equal(json.split(DISCLOSURE).length - 1, 1, `${what}: the disclosure sentence exactly once`);
}

/** The card's own chip. */
const hasChip = (activity: Sent | undefined): boolean => JSON.stringify(activity).includes('✨ AI-generated');

async function turnAnswer(answer: Record<string, unknown>, mode: { nativeAiLabel?: boolean } = {}): Promise<Sent[]> {
  const sim = simulatedBot(new SimulatedKernel(), {
    captureConversationReference: () => {},
    emitMembershipEvent: () => {},
    answer,
  });
  sim.bot.setNativeAiLabel(mode.nativeAiLabel === true);
  await sim.deliver(messageActivity({ bot: MIRA, from: ANNA, text: 'hallo', mention: true }));
  return (sim.sent as Sent[]).filter((activity) => activity['type'] !== 'typing');
}

describe('#2043 AI marking: turn answers', () => {
  it('a model answer on a card: the card\'s own chip and the label entity, sentence stripped, in both modes', async () => {
    for (const nativeAiLabel of [false, true]) {
      const posted = await turnAnswer({ text: folded('Die Antwort ist 42.'), aiDisclosure: { text: DISCLOSURE } }, { nativeAiLabel });

      assert.equal(posted.length, 1);
      assertLabelOnly(posted[0], 'turn answer');
      assert.ok(hasChip(posted[0]), 'the chip is the marking that is visible whatever Teams does with the entity');
      assert.ok(JSON.stringify(posted[0]).includes('Die Antwort ist 42.'));
    }
  });

  it('a fixed refusal of the facilitator, which now carries aiDisclosure too: label, sentence stripped', async () => {
    const refusal = 'Ich bin Mira, Moderator und kein Assistent: Fragen beantworten kann ich nicht.';
    const posted = await turnAnswer({ text: folded(refusal), aiDisclosure: { text: DISCLOSURE } });

    assert.equal(posted.length, 1);
    assertLabelOnly(posted[0], 'fixed refusal');
    assert.ok(JSON.stringify(posted[0]).includes('Moderator und kein Assistent'));
  });

  it('an answer without a disclosure (the operator turned it off): no chip, no label entity (0.41.0)', async () => {
    const posted = await turnAnswer({ text: 'Die Antwort ist 42.' });

    assert.equal(posted.length, 1);
    // 0.40.1 — the chip stands in for the kernel's sentence on a card. The
    // kernel disclosed nothing, so the channel draws no marker of its own.
    // 0.41.0 — and the activity-level label follows the chip (kernel #2087:
    // channels render their native label only when a marker is due).
    assert.equal(hasChip(posted[0]), false);
    assert.equal(labelsOf(posted[0]).length, 0, 'operator off: no AI label either');
    assert.ok(!JSON.stringify(posted[0]).includes(DISCLOSURE));
    assert.ok(JSON.stringify(posted[0]).includes('Die Antwort ist 42.'), 'the answer itself is unchanged');
  });

  it('the frequency setting (#2087): a marker that is not due leaves neither chip nor label', async () => {
    const notDue = { text: DISCLOSURE, markers: { textLine: false, chip: false } };
    const due = { text: DISCLOSURE, markers: { textLine: false, chip: true } };

    const [quiet] = await turnAnswer({ text: 'Die Antwort ist 42.', aiDisclosure: notDue });
    assert.equal(hasChip(quiet), false);
    assert.equal(labelsOf(quiet).length, 0);
    assert.ok(!JSON.stringify(quiet).includes(DISCLOSURE));

    const [marked] = await turnAnswer({ text: 'Die Antwort ist 42.', aiDisclosure: due });
    assert.equal(hasChip(marked), true);
    assert.equal(labelsOf(marked).length, 1);

    // A long answer that is not due carries no sentence either, whatever the label mode.
    const long = Array.from({ length: 700 }, (_, n) => `Absatz ${String(n)}: ${'lorem ipsum '.repeat(4)}`).join('\n\n');
    for (const nativeAiLabel of [false, true]) {
      const posted = await turnAnswer({ text: long, aiDisclosure: notDue }, { nativeAiLabel });
      assert.ok(posted.length > 1);
      assert.ok(!JSON.stringify(posted).includes(DISCLOSURE));
      assert.equal(posted.flatMap((activity) => labelsOf(activity)).length, 0);
    }
  });

  it('the chip follows the kernel\'s disclosure for every kind of turn answer: refusal and private status too', async () => {
    const refusal = 'Ich bin Mira, Moderator und kein Assistent: Fragen beantworten kann ich nicht.';
    const status = 'Stand: Schritt 2 von 3, eine Stimme fehlt.';
    for (const text of [refusal, status]) {
      const [disclosed] = await turnAnswer({ text: folded(text), aiDisclosure: { text: DISCLOSURE } });
      const [undisclosed] = await turnAnswer({ text });
      assert.equal(hasChip(disclosed), true, text);
      assert.equal(hasChip(undisclosed), false, text);
    }
  });

  it('a long answer posted as plain text has no chip: it keeps the sentence unless the label is confirmed', async () => {
    const long = Array.from({ length: 700 }, (_, n) => `Absatz ${String(n)}: ${'lorem ipsum '.repeat(4)}`).join('\n\n');
    assert.ok(long.length > 25_000, 'longer than one Teams message');
    const answer = { text: folded(long), aiDisclosure: { text: DISCLOSURE } };

    const unconfirmed = await turnAnswer(answer);
    assert.ok(unconfirmed.length > 1, 'split into several messages');
    assert.equal(hasChip(unconfirmed[0]), false);
    assert.equal(labelsOf(unconfirmed[0]).length, 1, 'the label entity is on the first part');
    assertLabelAndSentence(unconfirmed, 'long answer, label not confirmed');
    assert.ok(String(unconfirmed.at(-1)?.['text']).endsWith(DISCLOSURE), 'the sentence closes the last part');

    const confirmed = await turnAnswer(answer, { nativeAiLabel: true });
    assert.equal(confirmed.flatMap((activity) => labelsOf(activity)).length, 1);
    assert.ok(!JSON.stringify(confirmed).includes(DISCLOSURE), 'label only once it was seen in a client');
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
  function providers(mode: { nativeAiLabel?: boolean } = {}): {
    connector: SimulatedConnector;
    conversation: ReturnType<typeof createTeamsConversationSendAdapter>;
    targeted: ReturnType<typeof createTeamsTargetedSendAdapter>;
  } {
    const connector = new SimulatedConnector();
    const refs = new TeamsConversationReferenceCache();
    refs.capture({ activity: messageActivity({ bot: MIRA, from: ANNA, text: 'hi' }) } as unknown as TurnContext);
    return {
      connector,
      conversation: createTeamsConversationSendAdapter({
        refs,
        sendProactive: connector.sendProactive,
        ...(mode.nativeAiLabel ? { nativeAiLabel: true } : {}),
      }),
      targeted: createTeamsTargetedSendAdapter({
        sendProactive: connector.sendProactive,
        direct: { refs, openDirectConversation: connector.openDirectConversation, botLabel: () => MIRA.name },
        ...(mode.nativeAiLabel ? { nativeAiLabel: true } : {}),
      }),
    };
  }

  it('by default the providers do NOT declare the label, so the kernel adds the disclosure sentence', () => {
    // D24: declaring it made the kernel leave the sentence out while nobody
    // had seen Teams render the label on a proactive post.
    const p = providers();
    assert.equal(p.conversation.supports.aiLabel, false);
    assert.equal(p.targeted.supports?.aiLabel, false);
  });

  it('with the label confirmed both providers declare it', () => {
    const p = providers({ nativeAiLabel: true });
    assert.equal(p.conversation.supports.aiLabel, true);
    assert.equal(p.targeted.supports?.aiLabel, true);
  });

  it('a text the kernel hands over with the sentence goes out unchanged, with the label entity', async () => {
    const p = providers();

    await p.conversation.sendToConversation(GROUP_CHAT_ID, { text: folded('Die Abstimmung läuft.') }, { asChannelKey: MIRA.key });

    const sent = p.connector.sent[0]?.activity;
    assert.equal(sent?.['text'], folded('Die Abstimmung läuft.'), 'the adapter neither adds nor removes the sentence on this path');
    assert.equal(labelsOf(sent).length, 1);
  });

  it('what the providers post carries the label entity in both modes: text, poll card, private reminder', async () => {
    const p = providers({ nativeAiLabel: true });
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

  function delivery(
    withCard: boolean,
    mode: { nativeAiLabel?: boolean } = {},
  ): { connector: SimulatedConnector; send: ReturnType<typeof createTeamsRoutineDelivery> } {
    const connector = new SimulatedConnector();
    return {
      connector,
      send: createTeamsRoutineDelivery({
        sendProactive: connector.sendProactive,
        ...(mode.nativeAiLabel ? { nativeAiLabel: true } : {}),
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

  it('a routine delivery keeps the sentence by default, as a card and as plain text', async () => {
    const message = { text: folded('Heute stehen drei Termine an.'), aiDisclosure: { text: DISCLOSURE } } as never;
    const asCard = delivery(true);
    const asText = delivery(false);

    await asCard.send(reference, message, ROUTINE as never);
    await asText.send(reference, message, ROUTINE as never);

    assert.equal(asCard.connector.sent.length, 1);
    assertLabelAndSentence([asCard.connector.sent[0]?.activity], 'routine card');
    assert.ok(JSON.stringify(asCard.connector.sent[0]?.activity).includes('Morgenbericht: Heute stehen drei Termine an.'));
    assertLabelAndSentence([asText.connector.sent[0]?.activity], 'routine text');
    assert.equal(asText.connector.sent[0]?.activity['text'], folded('Heute stehen drei Termine an.'));
  });

  it('with the label confirmed a routine delivery drops the sentence, as a card and as plain text', async () => {
    const message = { text: folded('Heute stehen drei Termine an.'), aiDisclosure: { text: DISCLOSURE } } as never;
    const asCard = delivery(true, { nativeAiLabel: true });
    const asText = delivery(false, { nativeAiLabel: true });

    await asCard.send(reference, message, ROUTINE as never);
    await asText.send(reference, message, ROUTINE as never);

    assertLabelOnly(asCard.connector.sent[0]?.activity, 'routine card');
    assert.ok(JSON.stringify(asCard.connector.sent[0]?.activity).includes('Morgenbericht: Heute stehen drei Termine an.'));
    assertLabelOnly(asText.connector.sent[0]?.activity, 'routine text');
    assert.equal(asText.connector.sent[0]?.activity['text'], 'Heute stehen drei Termine an.');
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

describe('D24 the outgoing activity of a proactive post against a reply in a turn', () => {
  // The entity Microsoft's Teams SDK sends for `addAiGenerated()` (microsoft/
  // teams.ts at 96ec9329e: `activity.ts`, `models/entity/message-entity.ts`,
  // and its own test `activity.spec.ts`). `@id` is typed as required there:
  // "Must be left blank. This is for Bot Framework schema."
  const MICROSOFT_SDK_AI_ENTITY = {
    type: 'https://schema.org/Message',
    '@type': 'Message',
    '@context': 'https://schema.org',
    '@id': '',
    additionalType: ['AIGeneratedContent'],
  };

  async function proactiveText(): Promise<Sent> {
    const connector = new SimulatedConnector();
    const refs = new TeamsConversationReferenceCache();
    refs.capture({ activity: messageActivity({ bot: MIRA, from: ANNA, text: 'hi' }) } as unknown as TurnContext);
    const provider = createTeamsConversationSendAdapter({ refs, sendProactive: connector.sendProactive });
    await provider.sendToConversation(GROUP_CHAT_ID, { text: 'Die Abstimmung läuft.' }, { asChannelKey: MIRA.key });
    return connector.sent[0]?.activity as Sent;
  }

  it('both paths carry the same label entity, and it is the one Microsoft\'s SDK sends', async () => {
    const proactive = await proactiveText();
    const aiDisclosure = { text: DISCLOSURE };
    const [reply] = await turnAnswer({ text: 'Die Antwort ist 42.', aiDisclosure });
    const [longReply] = await turnAnswer({ text: 'Absatz. '.repeat(4000), aiDisclosure });

    for (const [name, activity] of [['proactive text', proactive], ['card reply', reply], ['plain-text reply', longReply]] as const) {
      assert.deepEqual(labelsOf(activity), [MICROSOFT_SDK_AI_ENTITY], name);
    }
    // Until 0.38.2 the entity went out without `@id`.
    assert.equal(Object.hasOwn(labelsOf(proactive)[0] as object, '@id'), true);
  });

  it('the proactive text activity, field by field', async () => {
    const proactive = await proactiveText();

    assert.deepEqual(proactive, {
      type: 'message',
      text: 'Die Abstimmung läuft.',
      entities: [MICROSOFT_SDK_AI_ENTITY],
    });
  });

  it('what a reply in a turn has and a proactive post has not', async () => {
    const proactive = await proactiveText();
    const [reply] = await turnAnswer({ text: folded('Die Antwort ist 42.'), aiDisclosure: { text: DISCLOSURE } });
    const [longReply] = await turnAnswer({ text: 'Absatz. '.repeat(4000), aiDisclosure: { text: DISCLOSURE } });

    // A card reply: an Adaptive Card with this plugin's own chip. This is
    // what was read as "the label" on turn answers in the client.
    assert.deepEqual(Object.keys(reply ?? {}).sort(), ['attachmentLayout', 'attachments', 'entities', 'inputHint', 'type']);
    assert.equal(hasChip(reply), true);
    // A plain-text reply and a proactive text differ in `inputHint` only:
    // nothing in the reply path sets a field the label could depend on.
    assert.deepEqual(Object.keys(longReply ?? {}).sort(), ['entities', 'inputHint', 'text', 'type']);
    assert.deepEqual(Object.keys(proactive).sort(), ['entities', 'text', 'type']);
    assert.equal(hasChip(longReply), false);
    assert.equal(hasChip(proactive), false);
    for (const activity of [reply, longReply, proactive]) {
      assert.equal(Object.hasOwn(activity ?? {}, 'channelData'), false, 'no path sets channelData (feedbackLoop, streamInfo)');
    }
  });
});

describe('D24 the outgoing activity log', () => {
  it('is off unless TEAMS_LOG_OUTGOING_ACTIVITIES is set', () => {
    assert.equal(isOutgoingLogEnabled({}), false);
    assert.equal(isOutgoingLogEnabled({ TEAMS_LOG_OUTGOING_ACTIVITIES: '0' }), false);
    assert.equal(isOutgoingLogEnabled({ TEAMS_LOG_OUTGOING_ACTIVITIES: '1' }), true);
    assert.equal(isOutgoingLogEnabled({ TEAMS_LOG_OUTGOING_ACTIVITIES: 'TRUE' }), true);
  });

  it('prints the whole activity as JSON with the path it took', () => {
    const lines: string[] = [];
    const original = console.error;
    console.error = (...args: unknown[]) => {
      lines.push(args.map(String).join(' '));
    };
    try {
      logOutgoingActivity('proactive', { type: 'message', text: 'x', entities: [{ '@id': '', additionalType: ['AIGeneratedContent'] }] });
    } finally {
      console.error = original;
    }
    assert.deepEqual(lines, [
      '[teams] outgoing path=proactive activity={"type":"message","text":"x","entities":[{"@id":"","additionalType":["AIGeneratedContent"]}]}',
    ]);
  });

  it('the switch reads as off unless it is explicitly on', () => {
    assert.equal(resolveNativeAiLabel(() => undefined), false);
    assert.equal(resolveNativeAiLabel(() => 'false'), false);
    assert.equal(resolveNativeAiLabel(() => 'true'), true);
    assert.equal(resolveNativeAiLabel(() => true), true);
  });
});

describe('the incoming invoke next to its response (0.38.5)', () => {
  const click = (data: unknown, verb = 'facilitation.card.answer') => ({
    type: 'invoke',
    name: 'adaptiveCard/action',
    id: 'f:1',
    replyToId: '1700000000001',
    conversation: { id: '19:abc@thread.v2' },
    from: { id: '29:x', aadObjectId: 'aad-1' },
    value: { action: { type: 'Action.Execute', verb, data } },
  });

  it('one line with name, verb, ids and the value', () => {
    assert.equal(
      incomingInvokeLogLine(click({ type: 'facilitation.card', cardId: 'c1', q0: 'b' }) as never),
      '[teams] incoming invoke name=adaptiveCard/action verb=facilitation.card.answer id=f:1 replyToId=1700000000001' +
        ' conversation=19:abc@thread.v2 from=aad-1' +
        ' value={"action":{"type":"Action.Execute","verb":"facilitation.card.answer","data":{"type":"facilitation.card","cardId":"c1","q0":"b"}}}',
    );
  });

  it('cuts free text after 200 characters, also inside data delivered as a JSON string', () => {
    const long = 'x'.repeat(250);
    for (const data of [{ q0: long }, JSON.stringify({ q0: long })]) {
      const line = incomingInvokeLogLine(click(data) as never) ?? '';
      assert.ok(line.includes(`"q0":"${'x'.repeat(200)}…(+50 chars)"`), line);
      assert.ok(!line.includes('x'.repeat(201)));
    }
  });

  it('is only for invokes', () => {
    assert.equal(incomingInvokeLogLine({ type: 'message', text: 'hi' } as never), undefined);
    assert.equal(incomingInvokeLogLine(undefined), undefined);
  });
});

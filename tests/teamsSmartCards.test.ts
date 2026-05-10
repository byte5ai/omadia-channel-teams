import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import {
  buildAnswerCard,
  buildChoiceAskCard,
  buildFollowUpsOnlyCard,
  CHOICE_ASK_VALUE_TYPE,
  FOLLOW_UP_VALUE_TYPE,
  parseChoiceAskValue,
  parseFollowUpValue,
} from '@omadia/channel-teams';

type Action = {
  type?: string;
  title?: string;
  style?: string;
  data?: { type?: string; label?: string; value?: string; prompt?: string };
};

function actions(card: ReturnType<typeof buildAnswerCard>): Action[] {
  const content = card.content as { actions?: Action[] };
  return content.actions ?? [];
}

function body(card: ReturnType<typeof buildAnswerCard>): unknown[] {
  const content = card.content as { body?: unknown[] };
  return content.body ?? [];
}

/**
 * Follow-up actions now live inside a body-level ActionSet (not the card's
 * `actions[]`), so they stay visually separated from Tool-Trace /
 * Fresh-Check. This helper flattens them back out for assertions.
 */
function followUpActions(card: ReturnType<typeof buildAnswerCard>): Action[] {
  const result: Action[] = [];
  for (const item of body(card)) {
    if (
      typeof item === 'object' &&
      item !== null &&
      (item as { type?: string }).type === 'ActionSet'
    ) {
      const set = item as { actions?: Action[] };
      if (Array.isArray(set.actions)) result.push(...set.actions);
    }
  }
  return result;
}

describe('parseChoiceAskValue', () => {
  it('accepts a valid payload', () => {
    const parsed = parseChoiceAskValue({
      type: CHOICE_ASK_VALUE_TYPE,
      label: 'Sales',
      value: 'sales_module',
    });
    assert.deepEqual(parsed, {
      type: CHOICE_ASK_VALUE_TYPE,
      label: 'Sales',
      value: 'sales_module',
    });
  });

  it('rejects wrong type marker', () => {
    assert.equal(
      parseChoiceAskValue({ type: 'nope', label: 'A', value: 'a' }),
      undefined,
    );
  });

  it('rejects missing label', () => {
    assert.equal(
      parseChoiceAskValue({ type: CHOICE_ASK_VALUE_TYPE, value: 'a' }),
      undefined,
    );
  });

  it('rejects empty value', () => {
    assert.equal(
      parseChoiceAskValue({
        type: CHOICE_ASK_VALUE_TYPE,
        label: 'A',
        value: '',
      }),
      undefined,
    );
  });

  it('rejects non-object input', () => {
    assert.equal(parseChoiceAskValue(undefined), undefined);
    assert.equal(parseChoiceAskValue('x'), undefined);
    assert.equal(parseChoiceAskValue(42), undefined);
  });
});

describe('parseFollowUpValue', () => {
  it('accepts a valid payload', () => {
    const parsed = parseFollowUpValue({
      type: FOLLOW_UP_VALUE_TYPE,
      label: 'Vorjahr',
      prompt: 'Top 5 Kunden nach Umsatz 2025',
    });
    assert.deepEqual(parsed, {
      type: FOLLOW_UP_VALUE_TYPE,
      label: 'Vorjahr',
      prompt: 'Top 5 Kunden nach Umsatz 2025',
    });
  });

  it('rejects wrong type marker', () => {
    assert.equal(
      parseFollowUpValue({ type: 'nope', label: 'A', prompt: 'p' }),
      undefined,
    );
  });

  it('rejects missing prompt', () => {
    assert.equal(
      parseFollowUpValue({ type: FOLLOW_UP_VALUE_TYPE, label: 'A' }),
      undefined,
    );
  });

  it('rejects empty label', () => {
    assert.equal(
      parseFollowUpValue({
        type: FOLLOW_UP_VALUE_TYPE,
        label: '',
        prompt: 'p',
      }),
      undefined,
    );
  });
});

describe('buildChoiceAskCard', () => {
  it('renders the question + 2 action buttons', () => {
    const card = buildChoiceAskCard({
      question: 'Welches Modul?',
      options: [
        { label: 'Sales', value: 'sales' },
        { label: 'POS', value: 'pos' },
      ],
    });
    const acts = actions(card);
    assert.equal(acts.length, 2);
    assert.equal(acts[0]!.title, 'Sales');
    assert.equal(acts[0]!.style, 'positive');
    assert.equal(acts[1]!.title, 'POS');
    assert.equal(acts[1]!.style, undefined);
    assert.equal(acts[0]!.data?.type, CHOICE_ASK_VALUE_TYPE);
    assert.equal(acts[0]!.data?.label, 'Sales');
    assert.equal(acts[0]!.data?.value, 'sales');
  });

  it('renders rationale as subtle text when provided', () => {
    const card = buildChoiceAskCard({
      question: 'Welches Modul?',
      rationale: 'Beide tracken Umsatz.',
      options: [
        { label: 'A', value: 'a' },
        { label: 'B', value: 'b' },
      ],
    });
    const items = body(card);
    const texts = items
      .filter((i): i is { type: string; text?: string } => {
        if (typeof i !== 'object' || i === null) return false;
        return (i as { type?: unknown }).type === 'TextBlock';
      })
      .map((i) => i.text ?? '');
    assert.ok(texts.some((t) => t.includes('Beide tracken Umsatz.')));
  });

  it('supports up to 4 options (spec max)', () => {
    const card = buildChoiceAskCard({
      question: 'Wähle',
      options: [
        { label: 'A', value: 'a' },
        { label: 'B', value: 'b' },
        { label: 'C', value: 'c' },
        { label: 'D', value: 'd' },
      ],
    });
    assert.equal(actions(card).length, 4);
  });

  it('truncates long value to 500 chars in round-trip payload', () => {
    const longVal = 'x'.repeat(600);
    const card = buildChoiceAskCard({
      question: 'Wähle',
      options: [
        { label: 'A', value: longVal },
        { label: 'B', value: 'b' },
      ],
    });
    assert.equal(actions(card)[0]!.data?.value?.length, 500);
  });
});

describe('buildAnswerCard — follow-up options', () => {
  it('adds one ActionSet entry per follow-up option in body', () => {
    const card = buildAnswerCard({
      answer: 'Hier die Top 5.',
      followUpOptions: [
        { label: 'Vorjahr', prompt: 'Top 5 Kunden 2025' },
        { label: 'Nach DB', prompt: 'Top 5 Kunden nach DB YTD 2026' },
      ],
    });
    const followUps = followUpActions(card);
    assert.equal(followUps.length, 2);
    assert.equal(followUps[0]!.title, 'Vorjahr');
    assert.equal(followUps[0]!.data?.prompt, 'Top 5 Kunden 2025');
    assert.equal(followUps[0]!.data?.label, 'Vorjahr');
    assert.equal(followUps[0]!.style, 'positive');
  });

  it('caps follow-ups at 4 even if more are passed in', () => {
    const card = buildAnswerCard({
      answer: 'Top 5',
      followUpOptions: [
        { label: 'A', prompt: 'Frage A' },
        { label: 'B', prompt: 'Frage B' },
        { label: 'C', prompt: 'Frage C' },
        { label: 'D', prompt: 'Frage D' },
        { label: 'E', prompt: 'Frage E' },
      ],
    });
    assert.equal(followUpActions(card).length, 4);
  });

  it('follow-ups sit in body ActionSet, NOT in card actions[]', () => {
    const card = buildAnswerCard({
      answer: 'Antwort',
      originalUserMessage: 'Originalfrage',
      followUpOptions: [
        { label: 'A', prompt: 'Frage A' },
        { label: 'B', prompt: 'Frage B' },
      ],
    });
    // Fresh-Check lives in card actions[]
    const cardActionTitles = actions(card).map((a) => a.title ?? '');
    assert.ok(cardActionTitles.some((t) => t.includes('Fresh Check')));
    // Follow-Ups must NOT leak into card actions[]
    const followUpsInCardActions = actions(card).filter(
      (a) => a.data?.type === FOLLOW_UP_VALUE_TYPE,
    );
    assert.equal(followUpsInCardActions.length, 0);
    // But follow-ups are reachable via body ActionSet
    const followUps = followUpActions(card);
    assert.equal(followUps.length, 2);
    assert.equal(followUps[0]!.title, 'A');
  });

  it('renders a Nächste Varianten header before the ActionSet', () => {
    const card = buildAnswerCard({
      answer: 'Antwort',
      followUpOptions: [
        { label: 'A', prompt: 'Frage A' },
        { label: 'B', prompt: 'Frage B' },
      ],
    });
    const texts = body(card)
      .filter((i): i is { type: string; text?: string } => {
        if (typeof i !== 'object' || i === null) return false;
        return (i as { type?: unknown }).type === 'TextBlock';
      })
      .map((i) => i.text ?? '');
    assert.ok(texts.some((t) => t.includes('Nächste Varianten')));
  });

  it('no ActionSet when no follow-ups', () => {
    const card = buildAnswerCard({ answer: 'Hi' });
    assert.equal(followUpActions(card).length, 0);
  });

  it('truncates long prompt to 500 chars in round-trip payload', () => {
    const card = buildAnswerCard({
      answer: 'Top 5',
      followUpOptions: [
        { label: 'Lang', prompt: 'x'.repeat(600) },
        { label: 'Kurz', prompt: 'ok' },
      ],
    });
    const first = followUpActions(card).find(
      (a) => a.data?.type === FOLLOW_UP_VALUE_TYPE,
    );
    assert.equal(first!.data?.prompt?.length, 500);
  });
});

describe('buildFollowUpsOnlyCard', () => {
  it('renders a standalone card with only follow-up actions', () => {
    const card = buildFollowUpsOnlyCard([
      { label: 'A', prompt: 'Frage A' },
      { label: 'B', prompt: 'Frage B' },
    ]);
    const acts = actions(card);
    assert.equal(acts.length, 2);
    assert.equal(acts[0]!.data?.type, FOLLOW_UP_VALUE_TYPE);
    const items = body(card);
    assert.ok(items.length >= 1);
  });
});

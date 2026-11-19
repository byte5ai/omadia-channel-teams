// The kernel's check of a `card_answer` event's typed part, COPIED from
// `middleware/src/channels/questionCard.ts` (byte5ai/omadia#2089, PR #2105).
//
// The kernel refuses the whole event when `answers` is not in the contract's
// form (unknown kind, missing value, two answers for one question, an id
// outside the id rule). Nothing in this repo had run an emitted answer
// through that check. The two functions below are the kernel's, verbatim; a
// test compares them with the kernel source when a kernel checkout sits next
// to this repo, so the copy cannot drift unnoticed.

import type { CardAnswerValue } from '../../src/kernel-types.js';

const CARD_ID_PATTERN = /^[A-Za-z0-9_-]{1,64}$/;
const MAX_CARD_QUESTIONS = 10;
const MAX_OPTIONS = 25;
const MAX_ANSWERS = MAX_CARD_QUESTIONS;
const MAX_ANSWER_TEXT = 10_000;
const MAX_REF = 512;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

type Read<T> = { readonly value: T } | { readonly error: string };

const isId = (raw: unknown): raw is string => typeof raw === 'string' && CARD_ID_PATTERN.test(raw);
const isText = (raw: unknown, max: number): raw is string =>
  typeof raw === 'string' && raw.trim().length > 0 && raw.length <= max;
const isOptionalText = (raw: unknown, max: number): raw is string | undefined =>
  raw === undefined || (typeof raw === 'string' && raw.length <= max);
const isCount = (raw: unknown): raw is number => typeof raw === 'number' && Number.isInteger(raw) && raw >= 0;
const isObject = (raw: unknown): raw is Record<string, unknown> =>
  raw !== null && typeof raw === 'object' && !Array.isArray(raw);

// --- copied verbatim: readAnswer -------------------------------------------
function readAnswer(raw: unknown): CardAnswerValue | undefined {
  if (!isObject(raw) || !isId(raw['questionId'])) return undefined;
  const questionId = raw['questionId'];
  const ids = (list: unknown): list is string[] =>
    Array.isArray(list) && list.length <= MAX_OPTIONS && list.every(isId) && new Set(list).size === list.length;
  switch (raw['kind']) {
    case 'single_choice':
      return isId(raw['optionId']) ? { questionId, kind: 'single_choice', optionId: raw['optionId'] } : undefined;
    case 'multiple_choice':
      return ids(raw['optionIds']) ? { questionId, kind: 'multiple_choice', optionIds: [...raw['optionIds']] } : undefined;
    case 'ranking':
      return ids(raw['ranking']) && raw['ranking'].length > 0
        ? { questionId, kind: 'ranking', ranking: [...raw['ranking']] }
        : undefined;
    case 'consent': {
      const { decision, reason } = raw;
      if ((decision !== 'agree' && decision !== 'object') || !isOptionalText(reason, MAX_ANSWER_TEXT)) return undefined;
      return { questionId, kind: 'consent', decision, ...(reason ? { reason } : {}) };
    }
    case 'assignment': {
      const { assignee } = raw;
      if (assignee === 'self') return { questionId, kind: 'assignment', assignee };
      if (!isObject(assignee) || !isText(assignee['kind'], 64) || !isText(assignee['id'], MAX_REF)) return undefined;
      const displayName = assignee['displayName'];
      return {
        questionId,
        kind: 'assignment',
        assignee: {
          kind: assignee['kind'],
          id: assignee['id'],
          ...(isText(displayName, 256) ? { displayName } : {}),
        },
      };
    }
    case 'text':
      return typeof raw['text'] === 'string' && raw['text'].length <= MAX_ANSWER_TEXT
        ? { questionId, kind: 'text', text: raw['text'] }
        : undefined;
    case 'number':
      return typeof raw['value'] === 'number' && Number.isFinite(raw['value'])
        ? { questionId, kind: 'number', value: raw['value'] }
        : undefined;
    case 'date':
      return typeof raw['value'] === 'string' && ISO_DATE.test(raw['value']) && !Number.isNaN(Date.parse(raw['value']))
        ? { questionId, kind: 'date', value: raw['value'] }
        : undefined;
    default:
      return undefined;
  }
}

// --- copied verbatim: readCardAnswerFields ----------------------------------
export function readCardAnswerFields(answer: {
  proposalId?: unknown;
  proposalVersion?: unknown;
  answers?: unknown;
}): Read<{ proposalId?: string; proposalVersion?: number; answers?: CardAnswerValue[] }> {
  const { proposalId, proposalVersion, answers } = answer;
  if (proposalId !== undefined && !isText(proposalId, MAX_REF)) return { error: 'cardAnswer.proposalId' };
  if (proposalVersion !== undefined && !isCount(proposalVersion)) return { error: 'cardAnswer.proposalVersion' };
  let typed: CardAnswerValue[] | undefined;
  if (answers !== undefined) {
    if (!Array.isArray(answers) || answers.length > MAX_ANSWERS) return { error: 'cardAnswer.answers' };
    typed = [];
    for (const raw of answers as unknown[]) {
      const one = readAnswer(raw);
      if (!one || typed.some((a) => a.questionId === one.questionId)) return { error: 'cardAnswer.answers' };
      typed.push(one);
    }
  }
  return {
    value: {
      ...(proposalId !== undefined ? { proposalId } : {}),
      ...(proposalVersion !== undefined ? { proposalVersion } : {}),
      ...(typed !== undefined ? { answers: typed } : {}),
    },
  };
}

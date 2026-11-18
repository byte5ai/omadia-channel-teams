// Outbound helpers for proactive posts and DMs (byte5ai/omadia#2043 §6.2):
// real @mentions, the optional poll card, and the mapping of Bot Connector
// errors to the kernel's delivery outcomes.

import type { Activity } from 'botbuilder';

import type { FacilitatorDeliveryOutcome, TargetedMention, TargetedPoll } from './kernel-types.js';
import type { MentionEntity } from './teamsMentions.js';
import { buildPollCard, validatePollInput, type BuildPollCardInput } from './teamsPollCard.js';

const HTTP_FORBIDDEN = 403;
const HTTP_NOT_FOUND = 404;

/** The two 403 answers Teams documents for a 1:1 the bot may not write to
 *  (send-proactive-messages). Other codes seen in the wild are not on any
 *  page we read and are deliberately not matched. */
const BLOCKED_MARKER = 'messagewritesblocked';
const NOT_INSTALLED_MARKER = 'forbiddenoperationexception';
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Statuses Teams asks bots to retry with exponential backoff (rate-limit). */
const RETRYABLE_STATUSES: readonly number[] = [429, 412, 502, 504];
const MAX_SEND_ATTEMPTS = 4;
const RETRY_BASE_DELAY_MS = 500;

/** An Entra object id. A Teams user id (`29:…`) is per bot and cannot open a
 *  conversation for another bot, so only this form is used for first contact. */
export function isEntraObjectId(raw: string | undefined): raw is string {
  return typeof raw === 'string' && UUID_PATTERN.test(raw);
}

function connectorStatus(err: unknown): number | undefined {
  const record = (err ?? {}) as Record<string, unknown>;
  const status = record['statusCode'] ?? record['status'];
  return typeof status === 'number' ? status : undefined;
}

export interface RetryOptions {
  /** Injected in tests; real time otherwise. */
  sleep?: (ms: number) => Promise<void>;
}

/**
 * Run one Bot Connector call, retrying 429 / 412 / 502 / 504 with exponential
 * backoff (0.5 s, 1 s, 2 s). Everything else, in particular every 403, is
 * thrown at once: a blocked or uninstalled bot must not knock again.
 */
export async function retryTransient<T>(
  call: () => Promise<T>,
  opts: RetryOptions = {},
  /** Return false to stop retrying (e.g. once part of the work is done). */
  mayRetry: () => boolean = () => true,
): Promise<T> {
  const sleep = opts.sleep ?? ((ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms)));
  for (let attempt = 1; ; attempt += 1) {
    try {
      return await call();
    } catch (err) {
      const status = connectorStatus(err);
      const isTransient = status !== undefined && RETRYABLE_STATUSES.includes(status);
      if (attempt >= MAX_SEND_ATTEMPTS || !isTransient || !mayRetry()) throw err;
      await sleep(RETRY_BASE_DELAY_MS * 2 ** (attempt - 1));
    }
  }
}

export interface RenderedMentions {
  text: string;
  entities: MentionEntity[];
  /** Display texts that stayed plain because the person is not in the chat. */
  unresolved: string[];
}

/**
 * Turn each `mention.text` occurring in `text` into a Teams mention. Teams
 * only notifies a person for an `<at>` token backed by a mention entity whose
 * `mentioned.id` is a member of the conversation; `resolveMemberId` maps the
 * kernel's user id to that member id and returns `undefined` for somebody who
 * is not in the chat. Such a mention stays plain text.
 */
export function renderMentions(
  text: string,
  mentions: readonly TargetedMention[] | undefined,
  resolveMemberId: (userId: string) => string | undefined,
): RenderedMentions {
  let rendered = text;
  const entities: MentionEntity[] = [];
  const unresolved: string[] = [];
  const seen = new Set<string>();
  for (const mention of mentions ?? []) {
    const display = mention.text;
    if (display.length === 0 || seen.has(display) || !rendered.includes(display)) continue;
    seen.add(display);
    const memberId = resolveMemberId(mention.user.id);
    if (!memberId) {
      unresolved.push(display);
      continue;
    }
    const name = display.replace(/^@/, '').replace(/[<>]/g, '').trim();
    if (name.length === 0) continue;
    const token = `<at>${name}</at>`;
    rendered = rendered.split(display).join(token);
    entities.push({ type: 'mention', text: token, mentioned: { id: memberId, name } });
  }
  return { text: rendered, entities, unresolved };
}

/**
 * The poll of a message (`TargetedMessage.poll`), checked once more before
 * anything is posted. The kernel validates it too; the adapter still refuses
 * a poll it cannot render rather than post half of one. Returns `undefined`
 * when the message carries no poll, and the reason when it cannot be rendered.
 */
export function readPollRequest(
  poll: TargetedPoll | undefined,
): { poll: BuildPollCardInput } | { problem: string } | undefined {
  if (poll === undefined || poll === null) return undefined;
  const candidate: BuildPollCardInput = {
    pollId: typeof poll.pollId === 'string' ? poll.pollId : '',
    question: typeof poll.question === 'string' ? poll.question : '',
    options: (Array.isArray(poll.options) ? poll.options : []).map((option) => ({
      id: typeof option?.id === 'string' ? option.id : '',
      label: typeof option?.label === 'string' ? option.label : '',
    })),
    ...(typeof poll.hint === 'string' ? { hint: poll.hint } : {}),
  };
  const problem = validatePollInput(candidate);
  return problem ? { problem } : { poll: candidate };
}

/**
 * The outgoing activities of one post: the text (with mentions) and the card
 * are SEPARATE activities. Teams may split an activity that carries both and
 * then returns no message id; sent separately, each returns its own id.
 */
export function buildOutgoingActivity(input: { text: string; entities?: readonly MentionEntity[] }): Partial<Activity> {
  return {
    type: 'message',
    text: input.text,
    ...(input.entities && input.entities.length > 0 ? { entities: [...input.entities] } : {}),
  };
}

export function buildPollActivity(poll: BuildPollCardInput): Partial<Activity> {
  return { type: 'message', attachments: [buildPollCard(poll)] };
}

function errorHaystack(err: unknown): string {
  const record = (err ?? {}) as Record<string, unknown>;
  const parts: unknown[] = [record['code'], record['message'], record['body'], record['details']];
  return parts
    .map((part) => {
      if (typeof part === 'string') return part;
      if (part === undefined || part === null) return '';
      try {
        return JSON.stringify(part);
      } catch {
        return '';
      }
    })
    .join(' ')
    .toLowerCase();
}

/**
 * Map a Bot Connector failure of a 1:1 send to the kernel's outcome.
 *   403 + subCode MessageWritesBlocked  → blocked
 *   403 + ForbiddenOperationException   → app_not_installed
 * Neither is ever retried. Teams answers MessageWritesBlocked both when the
 * user blocked the bot and when they uninstalled it; the response does not
 * say which, and both mean "remind in the group instead".
 * A 404 is reported as user_not_found; that reading is ours, no page states it.
 */
export function classifyDirectSendError(
  err: unknown,
  botLabel: string,
): Extract<FacilitatorDeliveryOutcome, { outcome: 'unreachable' }> {
  const status = connectorStatus(err);
  const haystack = errorHaystack(err);
  const detail = err instanceof Error ? err.message : String(err);
  if (status === HTTP_FORBIDDEN && haystack.includes(BLOCKED_MARKER)) {
    return {
      outcome: 'unreachable',
      code: 'blocked',
      message: `the user blocked '${botLabel}' or removed its app; Teams refuses direct messages from it (${detail})`,
    };
  }
  if (status === HTTP_FORBIDDEN && haystack.includes(NOT_INSTALLED_MARKER)) {
    return {
      outcome: 'unreachable',
      code: 'app_not_installed',
      message: `the app of '${botLabel}' is not installed for this user, so a first direct message is not possible (${detail})`,
    };
  }
  if (status === HTTP_NOT_FOUND) {
    return {
      outcome: 'unreachable',
      code: 'user_not_found',
      message: `Teams does not know this user or conversation for '${botLabel}' (${detail})`,
    };
  }
  return { outcome: 'unreachable', code: 'channel_error', message: detail };
}

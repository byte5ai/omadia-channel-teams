// Outbound helpers for proactive posts and DMs (byte5ai/omadia#2043 §6.2):
// real @mentions, the optional poll card, and the mapping of Bot Connector
// errors to the kernel's delivery outcomes.

import type { Activity } from 'botbuilder';

import type { FacilitatorDeliveryOutcome, TargetedMention } from './kernel-types.js';
import type { MentionEntity } from './teamsMentions.js';
import { buildPollCard, validatePollInput, type BuildPollCardInput } from './teamsPollCard.js';

const HTTP_FORBIDDEN = 403;
const HTTP_NOT_FOUND = 404;

/** Markers Teams documents for a 1:1 the bot may not write to. */
const BLOCKED_MARKERS = ['messagewritesblocked', 'conversationblockedbyuser'];
const NOT_INSTALLED_MARKERS = ['forbiddenoperationexception', 'botnotinconversationroster', 'not installed'];

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
 * The poll a message asks for. The kernel's `TargetedMessage` has no field
 * for it yet, so it is read structurally: `{ pollId, question, options }`.
 * Returns `undefined` when the message carries no poll, and the reason when
 * it carries one that cannot be rendered.
 */
export function readPollRequest(message: unknown): { poll: BuildPollCardInput } | { problem: string } | undefined {
  const raw = (message as { poll?: unknown } | null)?.poll;
  if (raw === undefined || raw === null) return undefined;
  if (typeof raw !== 'object') return { problem: 'poll must be an object' };
  const candidate = raw as Record<string, unknown>;
  const options = Array.isArray(candidate['options']) ? candidate['options'] : [];
  const poll: BuildPollCardInput = {
    pollId: typeof candidate['pollId'] === 'string' ? candidate['pollId'] : '',
    question: typeof candidate['question'] === 'string' ? candidate['question'] : '',
    options: options.map((option) => {
      const o = (option ?? {}) as Record<string, unknown>;
      return {
        id: typeof o['id'] === 'string' ? o['id'] : '',
        label: typeof o['label'] === 'string' ? o['label'] : '',
      };
    }),
    ...(typeof candidate['hint'] === 'string' ? { hint: candidate['hint'] } : {}),
  };
  const problem = validatePollInput(poll);
  return problem ? { problem } : { poll };
}

/** The outgoing message activity: text with mentions, plus the poll card. */
export function buildOutgoingActivity(input: {
  text: string;
  entities?: readonly MentionEntity[];
  poll?: BuildPollCardInput;
}): Partial<Activity> {
  return {
    type: 'message',
    text: input.text,
    ...(input.entities && input.entities.length > 0 ? { entities: [...input.entities] } : {}),
    ...(input.poll ? { attachments: [buildPollCard(input.poll)] } : {}),
  };
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
 *   403 + MessageWritesBlocked / ConversationBlockedByUser → blocked
 *   403 + ForbiddenOperationException / BotNotInConversationRoster
 *                                                    → app_not_installed
 *   404                                              → user_not_found
 * Teams also answers MessageWritesBlocked when the user UNINSTALLED the bot;
 * the two cannot be told apart from the response, and both mean "do not
 * retry, remind in the group instead".
 */
export function classifyDirectSendError(
  err: unknown,
  botLabel: string,
): Extract<FacilitatorDeliveryOutcome, { outcome: 'unreachable' }> {
  const record = (err ?? {}) as Record<string, unknown>;
  const status = typeof record['statusCode'] === 'number' ? record['statusCode'] : record['status'];
  const haystack = errorHaystack(err);
  const detail = err instanceof Error ? err.message : String(err);
  if (status === HTTP_FORBIDDEN && BLOCKED_MARKERS.some((marker) => haystack.includes(marker))) {
    return {
      outcome: 'unreachable',
      code: 'blocked',
      message: `the user blocked '${botLabel}' or removed its app; Teams refuses direct messages from it (${detail})`,
    };
  }
  if (status === HTTP_FORBIDDEN && NOT_INSTALLED_MARKERS.some((marker) => haystack.includes(marker))) {
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

// The incoming side of the outgoing-activity log (`TEAMS_LOG_OUTGOING_ACTIVITIES`).
//
// The log of what the bot SENDS shows an `invokeResponse`, and nothing about
// the click it answers. On the first live card checks that left one question
// open for a day: was a response the reply to a refused answer, or to the
// refresh a client sends by itself about a second after it shows a card? With
// the incoming invoke on the line before its response, both sides are there.

import type { Activity } from 'botbuilder';

/** Longer strings in an invoke's value are cut: the log is for the shape of a
 *  click (verb, input ids, chosen options), not for what people wrote. */
export const INVOKE_LOG_MAX_TEXT = 200;
const MAX_DEPTH = 8;

function shortened(text: string): string {
  return text.length <= INVOKE_LOG_MAX_TEXT
    ? text
    : `${text.slice(0, INVOKE_LOG_MAX_TEXT)}…(+${String(text.length - INVOKE_LOG_MAX_TEXT)} chars)`;
}

function parsedObject(text: string): unknown {
  // Some clients deliver an action's `data` as a JSON string.
  if (!text.startsWith('{')) return undefined;
  try {
    return JSON.parse(text) as unknown;
  } catch {
    return undefined;
  }
}

/** A copy of an invoke's value with every long string cut. */
export function invokeValueForLog(value: unknown, depth = 0): unknown {
  if (typeof value === 'string') {
    const parsed = depth < MAX_DEPTH ? parsedObject(value) : undefined;
    return parsed !== undefined ? invokeValueForLog(parsed, depth + 1) : shortened(value);
  }
  if (value === null || typeof value !== 'object') return value;
  if (depth >= MAX_DEPTH) return '(nested too deep)';
  if (Array.isArray(value)) return value.map((item) => invokeValueForLog(item, depth + 1));
  return Object.fromEntries(
    Object.entries(value as Record<string, unknown>).map(([key, item]) => [key, invokeValueForLog(item, depth + 1)]),
  );
}

/** The log line for one incoming invoke; `undefined` for any other activity. */
export function incomingInvokeLogLine(activity: Partial<Activity> | undefined): string | undefined {
  if (!activity || activity.type !== 'invoke') return undefined;
  const value = invokeValueForLog(activity.value);
  const verb = (value as { action?: { verb?: unknown } } | null)?.action?.verb;
  let json: string;
  try {
    json = JSON.stringify(value) ?? 'null';
  } catch {
    json = '"(not serialisable)"';
  }
  return (
    `[teams] incoming invoke name=${String(activity.name ?? '')} verb=${typeof verb === 'string' ? verb : ''}` +
    ` id=${String(activity.id ?? '')} replyToId=${String(activity.replyToId ?? '')}` +
    ` conversation=${String(activity.conversation?.id ?? '')} from=${String(activity.from?.aadObjectId ?? activity.from?.id ?? '')}` +
    ` value=${json}`
  );
}

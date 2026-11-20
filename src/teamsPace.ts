/**
 * Pace of a Teams turn (spec 2042 §5.8 "Pace", criterion 10): the typing
 * indicator from the start of the work, ONE short acknowledgement when the work
 * clearly takes longer than a normal answer, no progress chatter. The timing is
 * the channel SDK's (`runPacedTurn`); this module is what is Teams-specific.
 *
 * Teams clears the typing indicator after a few seconds, so it is renewed every
 * {@link TEAMS_TYPING_INTERVAL_MS}. The profile names no reply window: the
 * inline turn is answered through the live turn context, and a turn that
 * outlives it (a detached card click) already continues proactively. See the
 * report of byte5ai/omadia-channel-teams#42 for the open question on a hard
 * Bot Framework window.
 */

import { ActivityTypes, type TurnContext } from 'botbuilder';
import { DEFAULT_PACE_PROFILE, type ChannelPaceProfile, type PaceHooks } from '@omadia/channel-sdk';

export const TEAMS_TYPING_INTERVAL_MS = 5_000;

export const TEAMS_PACE_PROFILE: ChannelPaceProfile = Object.freeze({
  ackThresholdMs: DEFAULT_PACE_PROFILE.ackThresholdMs,
  typingIntervalMs: TEAMS_TYPING_INTERVAL_MS,
});

/** Typing and the one acknowledgement, posted through the turn's own context (the agent's bot). */
export function teamsPaceHooks(context: TurnContext): PaceHooks {
  return {
    sendTyping: async () => {
      await context.sendActivity({ type: ActivityTypes.Typing });
    },
    sendAck: async (text) => {
      await context.sendActivity(text);
    },
    onHookError: (hook, error) => {
      console.warn(`[teams] pace ${hook} failed:`, error instanceof Error ? error.message : error);
    },
  };
}

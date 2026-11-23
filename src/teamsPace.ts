/**
 * Pace of a Teams turn (spec 2042 §5.8 "Pace", criterion 10): the typing
 * indicator from the start of the work, ONE short acknowledgement when the work
 * clearly takes longer than a normal answer, no progress chatter. The timing is
 * the channel SDK's (`runPacedTurn`); this module is what is Teams-specific.
 *
 * Teams clears the typing indicator after a few seconds, so it is renewed every
 * {@link TEAMS_TYPING_INTERVAL_MS}.
 *
 * Reply window (byte5ai/omadia-channel-teams#42, decided): the Bot Framework
 * answers an inbound activity's HTTP request within about 15 seconds, and the
 * live turn context is only good for that request. A card click is therefore
 * already detached and continues proactively. A message turn that outlives
 * {@link TEAMS_REPLY_WINDOW_MS} is delivered as a follow-up on a proactive
 * continuation of the same conversation instead of through the live context
 * (spec 2042 §6 criterion 10: never a silent timeout). The acknowledgement
 * threshold stays below the window, so the person has heard from the agent
 * before the follow-up arrives.
 */

import { ActivityTypes, type TurnContext } from 'botbuilder';
import { DEFAULT_PACE_PROFILE, type ChannelPaceProfile, type PaceHooks } from '@omadia/channel-sdk';

export const TEAMS_TYPING_INTERVAL_MS = 5_000;
/** The Bot Framework's reply window for an inbound activity. */
export const TEAMS_REPLY_WINDOW_MS = 15_000;

export const TEAMS_PACE_PROFILE: ChannelPaceProfile = Object.freeze({
  ackThresholdMs: DEFAULT_PACE_PROFILE.ackThresholdMs,
  typingIntervalMs: TEAMS_TYPING_INTERVAL_MS,
  replyWindowMs: TEAMS_REPLY_WINDOW_MS,
});

/**
 * Typing and the one acknowledgement, posted through the turn's own context
 * (the agent's bot). `speakerLine` ("Mira:") goes in front of the
 * acknowledgement where one bot serves several agents and the speaker would be
 * unclear; on an agent's own bot the bot's name already says it.
 */
export function teamsPaceHooks(context: TurnContext, speakerLine?: string): PaceHooks {
  return {
    sendTyping: async () => {
      await context.sendActivity({ type: ActivityTypes.Typing });
    },
    sendAck: async (text) => {
      await context.sendActivity(speakerLine ? `${speakerLine} ${text}` : text);
    },
    onHookError: (hook, error) => {
      console.warn(`[teams] pace ${hook} failed:`, error instanceof Error ? error.message : error);
    },
  };
}

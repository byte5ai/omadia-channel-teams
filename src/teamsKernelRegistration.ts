// What this plugin hands to the kernel for native facilitators (#2043), in
// one place so it can be tested on the very objects that get registered.
//
// The kernel does not take a provider's word for granted:
//   - a send option is used only when the provider DECLARES it (`supports`);
//     without `asChannelKey: true` an own-identity send never reaches the
//     provider and is answered with `no_identity`;
//   - an emitted message or session event is accepted only when its
//     `channelId` and `channelType` belong to the plugin that emits it.

import type { CoreApi } from '@omadia/channel-sdk';

import type { TeamsOpenDirectConversation, TeamsProactiveSend } from './messagesRouter.js';
import {
  createTeamsConversationSendAdapter,
  createTeamsTargetedSendAdapter,
  type TeamsConversationReferenceCache,
} from './teamsGroupPrimitives.js';
import type { TeamsRosterProvider } from './teamsRoster.js';

/** The `channel_bindings.channel_type` every provider and event of this
 *  plugin carries. */
export const TEAMS_CHANNEL_TYPE = 'teams';

/** An event as the kernel accepts it from this plugin: its own catalog id and
 *  its own channel type, whatever the emitting module filled in. */
export function stampChannel<T extends { channelId: string; channelType: string }>(event: T, channelId: string): T {
  return { ...event, channelId, channelType: TEAMS_CHANNEL_TYPE };
}

export interface TeamsSendProviderDeps {
  refs: TeamsConversationReferenceCache;
  sendProactive: TeamsProactiveSend;
  roster: Pick<TeamsRosterProvider, 'list'>;
  openDirectConversation: TeamsOpenDirectConversation;
  /** Display name of a CONFIGURED bot (app id lowercase), else `undefined`. */
  botLabel(botAppId: string): string | undefined;
}

/**
 * Build and register the two send providers. Both are wired with everything
 * their declared capabilities need (per-bot references and first contact for
 * DMs, the roster for mentions), so what they declare is what they do.
 */
export function registerTeamsSendProviders(
  core: Pick<CoreApi, 'registerTargetedSendProvider' | 'registerConversationSendProvider'>,
  channelId: string,
  deps: TeamsSendProviderDeps,
): {
  targeted: ReturnType<typeof createTeamsTargetedSendAdapter>;
  conversation: ReturnType<typeof createTeamsConversationSendAdapter>;
} {
  const targeted = createTeamsTargetedSendAdapter({
    sendProactive: deps.sendProactive,
    // A DM "as bot X" (asChannelKey): per-bot references and first contact,
    // never another bot.
    direct: {
      refs: deps.refs,
      openDirectConversation: deps.openDirectConversation,
      botLabel: deps.botLabel,
    },
  });
  const conversation = createTeamsConversationSendAdapter({
    refs: deps.refs,
    sendProactive: deps.sendProactive,
    roster: deps.roster,
  });
  core.registerTargetedSendProvider?.(channelId, targeted);
  core.registerConversationSendProvider?.(channelId, conversation);
  return { targeted, conversation };
}

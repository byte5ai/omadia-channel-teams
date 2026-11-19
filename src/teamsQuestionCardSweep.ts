// Overdue question cards close themselves (0.40.2).
//
// WHY. A posted card is closed by the engine's `updateCard`. That call can
// fail to arrive for reasons this channel never sees: the kernel refuses the
// action before it reaches the channel (an operator's kill switch, a spent
// budget), or the engine is gone. The engine does not retry a card update.
// The card then stays in the chat with live inputs and its row stays `open`
// (rollout pass 6, observation O2).
//
// WHAT THE CHANNEL KNOWS ON ITS OWN is the card's closing time. A card that
// is still open some minutes past it is replaced by its closed form, without
// a result (the engine posts its result as text anyway), and its row is
// closed. A later `updateCard` of the engine still works and brings the
// result onto the card.
//
// The row is closed even when Teams refuses the replacement (the message was
// deleted, the bot was removed): nothing can be answered on that card any
// more, and the next sweep must not try again for ever.

import type { Activity } from 'botbuilder';

import type { TeamsProactiveSend } from './messagesRouter.js';
import { parseTeamsBotKey } from './teamsBotIdentity.js';
import { aiLabelEntity } from './teamsCard.js';
import type { TeamsConversationReferenceCache } from './teamsGroupPrimitives.js';
import { retryTransient, type RetryOptions } from './teamsOutboundMessage.js';
import { buildClosedCardContent, toCardAttachment, type ClosedCardView } from './teamsQuestionCardRender.js';
import type { TeamsQuestionCards } from './teamsQuestionCards.js';

/** How often the plugin looks for overdue cards. */
export const OVERDUE_SWEEP_INTERVAL_MS = 5 * 60 * 1000;

/** The message that replaces a posted card once it is closed. */
export function buildClosedCardActivity(closed: ClosedCardView): Partial<Activity> {
  return {
    type: 'message',
    attachments: [toCardAttachment(buildClosedCardContent(closed))],
    entities: [aiLabelEntity()],
  };
}

export interface OverdueCardSweepDeps {
  cards: TeamsQuestionCards;
  refs: Pick<TeamsConversationReferenceCache, 'getOrLoad'>;
  sendProactive: TeamsProactiveSend;
  retry?: RetryOptions;
  log?: (message: string) => void;
}

/**
 * One sweep: every card past its closing time that nobody closed is replaced
 * in Teams by its closed form and closed in the registry. Returns how many
 * cards were closed. Never throws.
 */
export function createOverdueCardSweep(deps: OverdueCardSweepDeps): () => Promise<number> {
  const log = deps.log ?? ((message: string) => console.error(message));
  return async () => {
    try {
      return await deps.cards.closeOverdue(async (posted, closed) => {
        if (!posted.messageId) throw new Error('the id of the posted message is not known');
        const cached = await deps.refs.getOrLoad(posted.conversationId, posted.botAppId);
        const ownerAppId = cached?.ref.bot?.id ? parseTeamsBotKey(cached.ref.bot.id) : undefined;
        if (!cached || (ownerAppId && ownerAppId !== posted.botAppId)) {
          throw new Error(`bot '${posted.botAppId}' has no Teams conversation reference for that conversation`);
        }
        const activity = { ...buildClosedCardActivity(closed), id: posted.messageId };
        await deps.sendProactive(
          cached.ref,
          async (turnContext) => {
            await retryTransient(() => turnContext.updateActivity(activity), deps.retry);
          },
          { botAppId: posted.botAppId },
        );
      }, log);
    } catch (err) {
      log(`[teams] closing overdue question cards failed: ${err instanceof Error ? err.message : String(err)}`);
      return 0;
    }
  };
}

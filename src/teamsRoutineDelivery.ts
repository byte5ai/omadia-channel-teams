// Deliveries the kernel pushes into a Teams conversation outside a chat
// turn: a routine's scheduled answer, a Conductor approval request, a
// Conductor reminder. Registered with `routinesIntegration.publishProactiveSend`.
//
// AI MARKING (#643, #2043). Everything posted under an agent's name carries
// Teams' "AI generated" label, once, and never the disclosure sentence:
//   - A routine delivery is an answer of the agent. The kernel hands it over
//     with the instance's disclosure line folded into `text` and named in
//     `aiDisclosure`. The line is stripped and the label attached, exactly as
//     for a turn answer. Without `aiDisclosure` (the operator turned the
//     disclosure off) there is nothing to strip; the label is attached all
//     the same.
//   - An approval card is deterministic kernel text with nothing folded in:
//     label only.

import type { Activity, Attachment, ConversationReference, TurnContext } from 'botbuilder';
import { isNoReply, logNoReplyDrop } from '@omadia/channel-sdk';
import type { RoutinesIntegration } from '@omadia/plugin-api';

import type { TeamsProactiveSend } from './messagesRouter.js';
import { aiLabelEntity, buildApprovalCard, stripFoldedAiDisclosure } from './teamsCard.js';

type ProactiveSender = Parameters<RoutinesIntegration['publishProactiveSend']>[1];
type RoutineMessage = Parameters<ProactiveSender>[1];
type RoutineInfo = Parameters<ProactiveSender>[2];

export interface TeamsRoutineDeliveryDeps {
  sendProactive: TeamsProactiveSend;
  /** The kernel's routine card (pause / delete inline); absent → plain text. */
  buildRoutineSmartCardAttachment?: (input: {
    routine: { id: string; name: string; cron: string };
    body: string;
  }) => { contentType: string; content: unknown };
}

/** The delivery's text without the disclosure line the kernel folded in. */
function deliveryBody(message: RoutineMessage): string {
  return stripFoldedAiDisclosure(message.text, (message as { aiDisclosure?: { text?: string } }).aiDisclosure?.text);
}

/** The one activity a delivery posts. Exported for the marking tests. */
export function buildRoutineDeliveryActivity(
  message: RoutineMessage,
  routine: RoutineInfo,
  buildRoutineSmartCardAttachment: TeamsRoutineDeliveryDeps['buildRoutineSmartCardAttachment'],
): Partial<Activity> {
  const entities = [aiLabelEntity()];
  // Conductor human-await request: the approve/reject card (shows WHAT is
  // being approved and the workflow's current step). A click resolves the
  // await in-process via the conductorAwaitResolver service.
  if (message.approval) {
    return { type: 'message', attachments: [buildApprovalCard(message.approval)], entities };
  }
  const body = deliveryBody(message);
  // With routine metadata and the integration's card builder: the card, so
  // the user sees this is a scheduled delivery and can pause or delete it
  // inline. Otherwise plain text.
  if (routine && buildRoutineSmartCardAttachment) {
    return {
      type: 'message',
      attachments: [buildRoutineSmartCardAttachment({ routine, body }) as Attachment],
      entities,
    };
  }
  return { type: 'message', text: body, entities };
}

export function createTeamsRoutineDelivery(deps: TeamsRoutineDeliveryDeps): ProactiveSender {
  return async (conversationRef, message, routine) => {
    // Routines fire on a cron without a user question; the sentinel is the
    // agent's way to say "nothing to report this run". The delivery is
    // dropped (the run history still records the turn) and the chat stays
    // quiet.
    // Checked on the text without a folded disclosure line as well: with the
    // line appended the sentinel is no longer the last line.
    if (isNoReply(message) || isNoReply({ text: deliveryBody(message) })) {
      logNoReplyDrop('teams', {
        trigger: 'routine',
        ...(routine?.id ? { routineId: routine.id } : {}),
        ...(routine?.name ? { routineName: routine.name } : {}),
      });
      return;
    }
    await deps.sendProactive(conversationRef as Partial<ConversationReference>, async (turnContext: TurnContext) => {
      await turnContext.sendActivity(
        buildRoutineDeliveryActivity(message, routine, deps.buildRoutineSmartCardAttachment),
      );
    });
  };
}

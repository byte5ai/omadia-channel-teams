import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
} from 'botbuilder';
import type {
  ConversationReference,
  Activity,
  TurnContext,
} from 'botbuilder';
import { Router } from 'express';
import type { Request, Response } from 'express';
import type { TeamsBot } from './teamsBot.js';

export interface TeamsRouterDeps {
  bot: TeamsBot;
  appId: string;
  appPassword: string;
  appType: 'MultiTenant' | 'SingleTenant' | 'UserAssignedMSI';
  appTenantId?: string;
}

/**
 * Channel-native proactive send. Picks up a previously-captured
 * `ConversationReference` (taken by `TurnContext.getConversationReference`
 * during an inbound turn — typically stored on a routine row) and pushes
 * an outbound activity into that conversation. The routines runner uses
 * it to deliver a scheduled agent answer.
 */
export type TeamsProactiveSend = (
  reference: Partial<ConversationReference>,
  build: (turnContext: TurnContext) => Promise<void>,
) => Promise<void>;

export interface TeamsRouterArtifacts {
  router: Router;
  /** Long-lived CloudAdapter — share with proactive-send callers so the BF
   *  authentication factory + token cache are reused. */
  adapter: CloudAdapter;
  /** Convenience wrapper: `continueConversationAsync` bound to this
   *  adapter + the bot's app id. Pass an async builder that calls
   *  `turnContext.sendActivity({...})`. */
  sendProactive: TeamsProactiveSend;
}

export function createTeamsRouter(deps: TeamsRouterDeps): TeamsRouterArtifacts {
  const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: deps.appId,
    MicrosoftAppPassword: deps.appPassword,
    MicrosoftAppType: deps.appType,
    MicrosoftAppTenantId: deps.appTenantId,
  });

  const botFrameworkAuthentication = createBotFrameworkAuthenticationFromConfiguration(
    null,
    credentialsFactory,
  );

  const adapter = new CloudAdapter(botFrameworkAuthentication);

  adapter.onTurnError = async (context, error) => {
    console.error('[teams] onTurnError:', error);
    try {
      await context.sendActivity('Es gab einen internen Fehler. Bitte erneut versuchen.');
    } catch (sendErr) {
      console.error('[teams] onTurnError: could not send error activity', sendErr);
    }
  };

  const router = Router();
  router.post('/messages', (req: Request, res: Response) => {
    // CloudAdapter handles BF authentication and calls the bot's turn handler.
    adapter.process(req, res, (context) => deps.bot.run(context));
  });

  const sendProactive: TeamsProactiveSend = async (reference, build) => {
    await adapter.continueConversationAsync(
      deps.appId,
      reference as Partial<Activity>,
      build,
    );
  };

  return { router, adapter, sendProactive };
}

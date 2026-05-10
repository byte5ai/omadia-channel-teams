import {
  ActivityTypes,
  CardFactory,
  MessageFactory,
  TeamsActivityHandler,
  TurnContext,
  type SigninStateVerificationQuery,
} from 'botbuilder';
import type {
  ChatAgent,
  ConversationHistoryStore,
  FollowUpOption,
  OutgoingAttachment,
  OutgoingChoiceCard,
  OutgoingSlotPicker,
  RunTracePayload,
  TopicDetector,
  TurnContextModule,
  VerifierBadge,
} from './kernel-types.js';
import type {
  PersistedAttachment,
  TeamsAttachmentStore,
} from './teamsAttachmentStore.js';
import {
  aiLabelEntity,
  buildAnswerCard,
  buildChoiceAskCard,
  buildFollowUpsOnlyCard,
  buildSlotPickerCard,
  buildTopicAskCard,
  parseBookSlotValue,
  parseChoiceAskValue,
  parseFollowUpValue,
  parseFreshCheckValue,
  parseRoutineCardActionValue,
  parseRoutineListFilterValue,
  parseTopicDecisionValue,
} from './teamsCard.js';
import type { TeamsRosterProvider } from './teamsRoster.js';
import {
  resolveMentions,
  stripMentionTokens,
  type MentionEntity,
} from './teamsMentions.js';

// Teams accepts long messages but the UI truncates around 28k chars. We leave a safety
// margin and split on paragraph boundaries so tables aren't cut in half.
const TEAMS_MAX_MESSAGE_CHARS = 25_000;

// Frequency of typing indicators during long-running orchestrator calls. The Teams
// client dims the "typing" animation after ~10 s of silence, so we refresh faster.
const TYPING_INTERVAL_MS = 5_000;

/**
 * Teams bot that delegates every turn to the orchestrator. No session stickiness —
 * each incoming message becomes a fresh orchestrator chat. Conversation continuity
 * comes from the persistent memory store, not from a stored session.
 */
export class TeamsBot extends TeamsActivityHandler {
  constructor(
    private readonly orchestrator: ChatAgent,
    private readonly history: ConversationHistoryStore,
    private readonly topicDetector: TopicDetector | undefined,
    /**
     * Kernel-owned `turnContext` AsyncLocalStorage module. Injected by the
     * plugin's `activate()` so the Teams adapter and the orchestrator share
     * the same storage instance — critical for `runWithChatParticipants`
     * handing the roster provider down into the orchestrator's own `run()`.
     */
    private readonly turnContext: TurnContextModule,
    private readonly attachmentStore?: TeamsAttachmentStore,
    private readonly rosterProvider?: TeamsRosterProvider,
    /**
     * Bot Framework OAuth Connection Name. When set, every inbound turn
     * attempts a silent `UserTokenClient.getUserToken` lookup; the resulting
     * JWT is threaded as `ssoAssertion` so the calendar tools can OBO-exchange
     * for delegated Graph scopes. Missing connection OR missing token → the
     * assertion stays undefined and calendar tools return `sso_unavailable`.
     */
    private readonly ssoConnectionName?: string,
    /**
     * Tenant id passed into the routines per-turn context. Comes from the
     * kernel's `GRAPH_TENANT_ID` (same key the verifier-store + graph use)
     * so routines stay aligned with the rest of the user-scoped data.
     */
    private readonly tenantId?: string,
    /**
     * Optional kernel hook called once per inbound turn, BEFORE the
     * chatAgent is invoked. The kernel installs the routines per-turn
     * AsyncLocalStorage from this hook so the `manage_routine` tool can
     * attribute `create` to the right user and capture the conversation
     * reference for proactive delivery. Undefined when routines is off.
     */
    private readonly captureRoutineTurn?: (info: {
      tenant: string;
      userId: string;
      conversationRef: unknown;
    }) => void,
    /**
     * Optional kernel-owned dispatcher for routine smart-card actions.
     * Invoked when the bot detects an inbound Activity whose `value`
     * matches `{kind: 'routine.action', action: 'pause'|'delete', id}`.
     * Returns a short user-facing confirmation string the bot renders
     * back into the chat.
     */
    private readonly handleRoutineAction?: (input: {
      action: 'pause' | 'resume' | 'trigger_now' | 'delete';
      id: string;
    }) => Promise<string>,
    /**
     * Builder for the routine LIST smart card sidecar (rendered after the
     * agent's prose answer when `result.interactive.kind === 'routine_list'`).
     * Kernel-supplied; undefined → list-card path skipped silently.
     */
    private readonly buildRoutineListSmartCardAttachment?: (input: {
      filter: 'all' | 'active' | 'paused';
      totals: { all: number; active: number; paused: number };
      routines: Array<{
        id: string;
        name: string;
        cron: string;
        prompt: string;
        status: 'active' | 'paused';
        lastRunAt: string | null;
        lastRunStatus: 'ok' | 'error' | 'timeout' | null;
      }>;
    }) => { contentType: string; content: unknown },
  ) {
    super();

    this.onMessage(async (context, next) => {
      await this.handleMessage(context);
      await next();
    });
  }

  private async handleMessage(context: TurnContext): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    const sessionScope = `teams-${conversationId}`;
    const from = context.activity.from;
    const userId =
      from?.aadObjectId ??
      (typeof from?.id === 'string' ? from.id : undefined);
    // Hard entry-point log — stderr so Fly can't silently drop it. Tells
    // us a Teams message actually reached the bot process, independent
    // of every downstream log that might get filtered under load.
    console.error(
      `[teams] inbound conv=${conversationId} aad=${from?.aadObjectId ?? '-'} type=${context.activity.type ?? '-'} hasText=${String(Boolean(context.activity.text))} hasAttachments=${String((context.activity.attachments?.length ?? 0) > 0)}`,
    );

    // Persist any file / image attachments BEFORE delegating to the
    // orchestrator — we want to thread the stored metadata (storage key +
    // signed URL + filename) into the user message so Claude can decide
    // whether to treat it as a brand asset (Logo → /memories/_brand/logo.md)
    // or reference it in a later diagram. Failures are swallowed by the
    // store itself and yield an empty list, so the bot still replies.
    const persistedAttachments = await this.persistAttachmentsIfAny(
      context,
      conversationId,
      userId,
    );

    // Adaptive-card button click: the activity carries a structured `value`.
    // Resolve any pending topic decision *before* we treat it like a normal
    // text message, otherwise the user would see "Ich habe keinen Text…".
    const decision = parseTopicDecisionValue(context.activity.value);
    if (decision) {
      await this.handleTopicDecision(context, sessionScope, userId, decision);
      return;
    }

    // "🔄 Fresh Check" — user clicked the button on a previous answer card
    // to re-run the same question while bypassing memory + context block.
    const freshCheck = parseFreshCheckValue(context.activity.value);
    if (freshCheck) {
      await this.handleFreshCheck(context, sessionScope, userId, freshCheck.originalMessage);
      return;
    }

    // `ask_user_choice` option click — blocking clarification resolved.
    // The chosen `value` is injected as the next user turn so the LLM can
    // re-run with the disambiguated input.
    const choiceAsk = parseChoiceAskValue(context.activity.value);
    if (choiceAsk) {
      await this.handleChoiceAsk(context, sessionScope, userId, choiceAsk);
      return;
    }

    // `suggest_follow_ups` refinement click — non-blocking. The full
    // `prompt` (not just the label) is sent as the next user turn; the
    // LLM treats it as a fresh self-contained question.
    const followUp = parseFollowUpValue(context.activity.value);
    if (followUp) {
      await this.handleFollowUp(context, sessionScope, userId, followUp);
      return;
    }

    // Slot-picker click from `find_free_slots` sidecar card. The synthetic
    // user message carries the opaque slotId so the LLM can call
    // `book_meeting({slotId, subject})` on the next turn with the subject
    // inferred from prior conversation context.
    const bookSlot = parseBookSlotValue(context.activity.value);
    if (bookSlot) {
      await this.handleBookSlotClick(context, sessionScope, userId, bookSlot);
      return;
    }

    // Routine smart-card click (Pausieren / Aktivieren / Löschen — either
    // on a cron-triggered delivery or on a routine row inside the list
    // smart-card). Dispatched out-of-band — never reaches the orchestrator
    // because the action is purely lifecycle, not a conversation turn.
    const routineAction = parseRoutineCardActionValue(context.activity.value);
    if (routineAction && this.handleRoutineAction) {
      try {
        const ack = await this.handleRoutineAction({
          action: routineAction.action,
          id: routineAction.id,
        });
        await context.sendActivity(ack);
      } catch (err) {
        const detail = err instanceof Error ? err.message : String(err);
        const verb =
          routineAction.action === 'pause'
            ? 'pausieren'
            : routineAction.action === 'resume'
              ? 'aktivieren'
              : routineAction.action === 'trigger_now'
                ? 'jetzt auslösen'
                : 'löschen';
        await context.sendActivity(
          `Konnte die Routine nicht ${verb}: ${detail}`,
        );
      }
      return;
    }

    // Filter-pill click on the routine list smart-card. Synthesize a
    // German user message that the agent recognises and dispatches via
    // `manage_routine.list` with the chosen filter. Same pattern as the
    // follow-up / choice-ask buttons: the synthetic user-turn is
    // chronologically visible in the transcript so the user sees the
    // intent was registered.
    const routineFilter = parseRoutineListFilterValue(context.activity.value);
    if (routineFilter) {
      const phrase =
        routineFilter.filter === 'all'
          ? 'Zeig mir alle meine Routinen.'
          : routineFilter.filter === 'active'
            ? 'Zeig mir nur die aktiven Routinen.'
            : 'Zeig mir nur die pausierten Routinen.';
      this.history.clearPending(sessionScope);
      const priorTurns = this.history.get(sessionScope);
      await this.runOrchestratorTurn(context, {
        conversationId,
        sessionScope,
        userId,
        userMessage: phrase,
        priorTurns,
      });
      return;
    }

    const rawUserMessage = extractUserMessage(context);

    // Magic-code intercept for the bypass sign-in flow. When the user pastes
    // the 6-digit code from the BF token-service page into the chat, Bot
    // Framework does NOT intercept it automatically (that would require the
    // stock OAuthCard we deliberately sidestepped). We call `getUserToken`
    // with the code as the magic-code param — success stores the token
    // against user+connection so subsequent `extractSsoAssertion` calls
    // succeed silently.
    const channelUserId = context.activity.from?.id;
    if (this.ssoConnectionName && /^\d{6}$/.test(rawUserMessage) && channelUserId) {
      const exchanged = await this.tryExchangeMagicCode(
        context,
        channelUserId,
        rawUserMessage,
      );
      if (exchanged) {
        console.error(
          `[teams] magic-code exchange ok conv=${conversationId} user=${userId}`,
        );
        await context.sendActivity(
          'Kalender-Zugriff freigegeben ✓ — stell deine Terminfrage einfach nochmal, dann buche ich direkt.',
        );
        return;
      }
      console.error(
        `[teams] magic-code exchange failed conv=${conversationId} — continuing as normal message`,
      );
    }
    // Empty text + no attachment → ask for a real question. Empty text
    // WITH attachments is treated as "hier ein File, merk dir das" — we
    // let the orchestrator see the attachment-hint and decide how to
    // react (usually: store as brand asset + short acknowledgement).
    if (rawUserMessage.length === 0 && persistedAttachments.length === 0) {
      await context.sendActivity(
        'Ich habe keinen Text in deiner Nachricht gefunden. Schreib mir ruhig eine Frage — z. B. "Wer sind unsere umsatzstärksten Kunden?".',
      );
      return;
    }
    const attachmentHint = this.buildAttachmentHint(persistedAttachments);
    const userMessage = attachmentHint
      ? `${rawUserMessage || '(keine Texteingabe — nur eine Datei)'}${attachmentHint}`
      : rawUserMessage;

    // A fresh typed message while a topic-ask was pending resolves that
    // ambiguity implicitly — the user didn't click the buttons, they moved
    // on. We clear the pending flag so the bot doesn't keep asking.
    this.history.clearPending(sessionScope);

    // Topic detection runs before the orchestrator so we don't pay a full
    // agent turn just to hand the user a clarification card. The detector
    // is optional — if unavailable (sidecar down, config off), we fall
    // through to the previous behaviour.
    const priorTurns = this.history.get(sessionScope);
    let effectiveHistory = priorTurns;
    if (this.topicDetector && priorTurns.length > 0) {
      try {
        const verdict = await this.topicDetector.classify({
          userMessage,
          history: priorTurns,
        });
        // stderr so Fly's log aggregator doesn't silently drop it — this is
        // the only signal that tells us which branch fired per turn, and we
        // saw other stdout-INFO lines vanish under load.
        console.error(
          `[topic] decision=${verdict.decision} reason=${verdict.reason}` +
            (verdict.similarity !== undefined
              ? ` sim=${verdict.similarity.toFixed(3)}`
              : '') +
            (verdict.classifier ? ` cls=${verdict.classifier}` : '') +
            ` conv=${conversationId}`,
        );
        if (verdict.decision === 'reset') {
          this.history.resetTurns(sessionScope);
          effectiveHistory = [];
        } else if (verdict.decision === 'ask') {
          const lastTurn = priorTurns.at(-1);
          const summary = lastTurn
            ? `Deine letzte Frage: "${truncate(lastTurn.userMessage, 180)}"`
            : 'Der vorherige Chat-Kontext liegt vor.';
          await sendTopicAsk(context, summary, userMessage);
          this.history.markPending(sessionScope, {
            userMessage,
            askedAt: Date.now(),
          });
          console.error(
            `[topic] asked user conv=${conversationId} msg="${truncate(userMessage, 60)}"`,
          );
          return;
        }
      } catch (err) {
        console.error(
          '[topic] classify failed — continuing with full history:',
          err instanceof Error ? err.message : err,
        );
      }
    }

    await this.runOrchestratorTurn(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage,
      priorTurns: effectiveHistory,
    });
  }

  /**
   * Downloads + persists attached files to Tigris + Neon, then returns the
   * list of successfully stored metadata records. Swallows all errors
   * internally (they land in the store's logger) and returns `[]` — the
   * Teams reply continues in every case.
   *
   * We await this (rather than fire-and-forget) so the caller can thread
   * the metadata into the orchestrator's user message. Typical logo upload
   * is sub-second; 25 MB cap + 30 s download timeout bounds the worst case.
   */
  private async persistAttachmentsIfAny(
    context: TurnContext,
    conversationId: string,
    userId: string | undefined,
  ): Promise<PersistedAttachment[]> {
    const activity = context.activity;
    const originalAttachments = activity.attachments ?? [];
    const turnTime = activity.timestamp
      ? new Date(activity.timestamp)
      : undefined;

    // Channel messages often put the file reference OUTSIDE `attachments[]`
    // — in `channelData`, `entities[]`, or as a bare URL in the message text.
    // Probe all of them, synthesise pseudo-attachments so the store's
    // existing resolver + Graph-download path can handle them uniformly.
    const synthesised = this.extractOutOfBandAttachments(activity);

    // Group-chat fallback: Teams does NOT forward file attachments to
    // the bot Activity even with RSC `ChatMessage.Read.Chat` granted
    // (empirical observation, 2026-04-20). We use the same RSC permission
    // to pull the message via Graph, which *does* surface attachments.
    // Runs only when we have message + chat ids AND the store has a
    // graph client configured.
    const messageId =
      typeof activity.id === 'string' && activity.id.length > 0
        ? activity.id
        : undefined;
    let graphSynthesised: Array<{
      contentType: string;
      contentUrl: string;
      name: string;
    }> = [];
    if (this.attachmentStore && messageId && conversationId !== 'unknown') {
      graphSynthesised = await this.attachmentStore.discoverFromGraphMessage(
        conversationId,
        messageId,
      );
    }

    // Always log, even on empty — the 0-case tells us "a message arrived
    // that we could NOT find a file reference in", which is exactly what
    // we need to triage channel upload failures.
    const channelDataKeys = activity.channelData
      ? Object.keys(activity.channelData as Record<string, unknown>)
          .slice(0, 8)
          .join(',')
      : '-';
    const entityTypes = Array.isArray(activity.entities)
      ? activity.entities
          .map((e) => String(e?.type ?? '?'))
          .slice(0, 8)
          .join(',')
      : '-';
    // stderr: Fly's log aggregator drops stdout INFO under load; this
    // diag line is the only signal that tells us what shape the inbound
    // Teams activity has, so it MUST make it through. Logged
    // unconditionally (even when attachmentStore is disabled) so we can
    // triage inbound shape independently of the store wiring.
    console.error(
      `[teams] activity conv=${conversationId} attachments=${String(originalAttachments.length)} entities=${String(activity.entities?.length ?? 0)} entityTypes=${entityTypes} channelDataKeys=${channelDataKeys} synthesised=${String(synthesised.length)} graphSynthesised=${String(graphSynthesised.length)} textLen=${String(activity.text?.length ?? 0)} storeReady=${String(Boolean(this.attachmentStore))}`,
    );

    if (!this.attachmentStore) return [];

    const allAttachments = [
      ...originalAttachments,
      ...synthesised,
      ...graphSynthesised,
    ];
    if (allAttachments.length === 0) return [];

    try {
      return await this.attachmentStore.persistTurn({
        conversationId,
        ...(userId ? { userId } : {}),
        ...(turnTime ? { turnTime } : {}),
        attachments: allAttachments,
      });
    } catch (err) {
      console.error(
        '[teams] attachment persistence threw (swallowed):',
        err instanceof Error ? err.message : err,
      );
      return [];
    }
  }

  /**
   * Gather file references that aren't in `activity.attachments[]`:
   *   1. SharePoint / OneDrive URLs in the message text (plain text or
   *      extracted from HTML `<a href>` / `<img src>`).
   *   2. Likely the same URLs surfaced in `entities[]`.
   *   3. Teams-specific `channelData.attachments` when present.
   *
   * Each hit becomes a minimal `Attachment`-shaped object with
   * `contentUrl = <sharing-url>`, so the store's existing SharePoint
   * resolver can route it through the Graph client.
   */
  private extractOutOfBandAttachments(activity: {
    text?: string;
    entities?: unknown[];
    channelData?: unknown;
  }): Array<{ contentType: string; contentUrl: string; name: string }> {
    const out: Array<{ contentType: string; contentUrl: string; name: string }> = [];
    const seen = new Set<string>();

    const push = (url: string, nameHint?: string): void => {
      if (seen.has(url)) return;
      if (!isSharePointLike(url)) return;
      seen.add(url);
      const name = nameHint && nameHint.trim().length > 0
        ? nameHint.trim()
        : fileNameFromUrl(url) ?? 'sharepoint-file';
      out.push({ contentType: '', contentUrl: url, name });
    };

    // 1) scan message.text (Teams sometimes delivers HTML, sometimes plain)
    const text = activity.text ?? '';
    if (text) {
      for (const m of text.matchAll(/https?:\/\/[^\s"'<>]+/g)) {
        const url = m[0];
        push(url);
      }
    }

    // 2) scan channelData (shape varies; dump generically)
    if (activity.channelData && typeof activity.channelData === 'object') {
      for (const url of collectStringsDeep(activity.channelData)) {
        if (typeof url === 'string') push(url);
      }
    }

    // 3) entities[]
    if (Array.isArray(activity.entities)) {
      for (const ent of activity.entities) {
        if (ent && typeof ent === 'object') {
          for (const url of collectStringsDeep(ent)) {
            if (typeof url === 'string') push(url);
          }
        }
      }
    }

    return out;
  }

  /**
   * Produce a short system-style block that the orchestrator gets appended
   * to the user's message. Tells Claude exactly which files arrived in
   * this turn, their signed URLs (if any), and the memory convention for
   * promoting them to persistent brand assets. Returns '' when nothing to
   * announce.
   */
  private buildAttachmentHint(
    persisted: readonly PersistedAttachment[],
  ): string {
    if (persisted.length === 0) return '';
    const lines: string[] = [
      '',
      '---',
      `[attachments-info] ${String(persisted.length)} Datei(en) in diesem Turn hochgeladen + persistiert:`,
    ];
    for (const a of persisted) {
      const sizeKb = Math.round(a.sizeBytes / 1024);
      lines.push(
        `- ${a.fileName} (${a.contentType}, ${String(sizeKb)} KB) · storage_key=${a.storageKey}${a.signedUrl ? ` · signed_url=${a.signedUrl}` : ''}`,
      );
    }
    lines.push(
      '',
      'Wenn der User diese Datei als **Brand-Asset** markiert ("das ist unser Logo", "unser Firmenlogo", "das Team-Banner"), schreibe/aktualisiere die Memory-Datei `/memories/_brand/<asset-name>.md` (z.B. `/memories/_brand/logo.md`) mit folgenden Feldern:',
      '```yaml',
      'storage_key: <storage_key von oben>',
      'signed_url: <signed_url von oben>',
      'file_name: <Dateiname>',
      'content_type: <MIME-Type>',
      'uploaded_at: <heutiges Datum>',
      'asset_role: logo | banner | favicon | …',
      '```',
      'Die signed_url ist 7 Tage gültig — wenn du sie später in einem Diagramm brauchst und sie abgelaufen ist, frag nach oder lass sie vom Middleware via storage_key neu signieren.',
      'Wenn der User die Datei NICHT als Asset markiert (nur "hier ein Screenshot", "schau mal drauf"), NICHT ins Brand-Memory schreiben — nur inhaltlich antworten.',
    );
    return lines.join('\n');
  }

  private async handleTopicDecision(
    context: TurnContext,
    sessionScope: string,
    userId: string | undefined,
    decision: { choice: 'continue' | 'reset'; originalMessage: string },
  ): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    const pending = this.history.getPending(sessionScope);
    // The card carries the original user message so we can proceed even if
    // the bot restarted in the meantime. `pending` just confirms we were in
    // the "asking" state; we don't hard-require it.
    if (!pending) {
      console.warn(
        `[topic] decision received but no pending entry for scope=${sessionScope}`,
      );
    }
    this.history.clearPending(sessionScope);
    if (decision.choice === 'reset') {
      this.history.resetTurns(sessionScope);
      console.error(
        `[topic] user chose RESET conv=${conversationId} msg="${truncate(decision.originalMessage, 60)}"`,
      );
    } else {
      console.error(
        `[topic] user chose CONTINUE conv=${conversationId} msg="${truncate(decision.originalMessage, 60)}"`,
      );
    }
    const priorTurns =
      decision.choice === 'reset' ? [] : this.history.get(sessionScope);
    await this.runOrchestratorTurn(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage: decision.originalMessage,
      priorTurns,
    });
  }

  /**
   * User clicked the "🔄 Fresh Check" button on a previous answer card.
   * Re-runs the original question while bypassing the memory-read
   * convention AND the FTS context block — the orchestrator is told
   * explicitly to treat the turn as isolated. We intentionally do NOT
   * pass `priorTurns` so the in-memory conversation history can't
   * re-contaminate the turn either.
   */
  private async handleFreshCheck(
    context: TurnContext,
    sessionScope: string,
    userId: string | undefined,
    originalMessage: string,
  ): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    console.error(
      `[teams] fresh-check conv=${conversationId} msg="${truncate(originalMessage, 80)}"`,
    );
    // Empty priorTurns array — the point of fresh-check is strict isolation.
    const priorTurns: ReturnType<ConversationHistoryStore['get']> = [];
    await this.runOrchestratorTurn(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage: originalMessage,
      priorTurns,
      freshCheck: true,
    });
  }

  /**
   * User clicked one of the option buttons on an `ask_user_choice` card.
   * The chosen `value` is sent as the next user turn with the full chat
   * history preserved — the LLM then has the disambiguated input and can
   * answer normally.
   */
  private async handleChoiceAsk(
    context: TurnContext,
    sessionScope: string,
    userId: string | undefined,
    choice: { label: string; value: string },
  ): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    console.error(
      `[teams] choice-ask click conv=${conversationId} label="${truncate(choice.label, 40)}" value="${truncate(choice.value, 80)}"`,
    );
    this.history.clearPending(sessionScope);
    const priorTurns = this.history.get(sessionScope);
    await this.runOrchestratorTurn(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage: choice.value,
      priorTurns,
    });
  }

  /**
   * User clicked a follow-up refinement button on a previous answer. The
   * option's full `prompt` becomes the next user turn — no free-text
   * extraction from the label needed because the LLM pre-composed the
   * self-contained question when invoking `suggest_follow_ups`.
   */
  private async handleFollowUp(
    context: TurnContext,
    sessionScope: string,
    userId: string | undefined,
    followUp: { label: string; prompt: string },
  ): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    console.error(
      `[teams] follow-up click conv=${conversationId} label="${truncate(followUp.label, 40)}" prompt="${truncate(followUp.prompt, 80)}"`,
    );
    this.history.clearPending(sessionScope);
    const priorTurns = this.history.get(sessionScope);
    await this.runOrchestratorTurn(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage: followUp.prompt,
      priorTurns,
    });
  }

  /**
   * Silent SSO path: Teams exchanges the user's token for our resource
   * without a popup. Fires when the user had previously consented and the
   * cached token was still valid. We acknowledge so the user knows the
   * original turn can be retried; the next message will find the token
   * silently via `extractSsoAssertion`.
   */
  protected override async handleTeamsSigninTokenExchange(
    context: TurnContext,
    _query: SigninStateVerificationQuery,
  ): Promise<void> {
    await this.ackConsentAndPromptRetry(context);
  }

  /**
   * Magic-code / popup path: user completed the interactive sign-in. Same
   * ack as the silent path — we don't auto-retry because we never stored
   * the original message, and asking the user to retype is clearer than a
   * phantom replay.
   */
  protected override async handleTeamsSigninVerifyState(
    context: TurnContext,
    _query: SigninStateVerificationQuery,
  ): Promise<void> {
    await this.ackConsentAndPromptRetry(context);
  }

  private async ackConsentAndPromptRetry(context: TurnContext): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    console.error(`[teams] consent granted conv=${conversationId}`);
    await context.sendActivity(
      'Kalender-Zugriff freigegeben ✓ — stell deine Terminfrage einfach nochmal, dann buche ich direkt.',
    );
  }

  /**
   * Exchange a user-supplied 6-digit magic code from the Bot Framework
   * sign-in redirect page. Returns `true` when a token is now bound to
   * (user, connection), `false` otherwise (wrong code / expired flow /
   * adapter shape mismatch). Swallows errors — the caller falls back to
   * treating the text as a normal message.
   */
  private async tryExchangeMagicCode(
    context: TurnContext,
    userId: string,
    magicCode: string,
  ): Promise<boolean> {
    try {
      const adapter = context.adapter as unknown as { UserTokenClientKey?: symbol };
      if (!adapter.UserTokenClientKey || !this.ssoConnectionName) return false;
      const client = context.turnState.get(adapter.UserTokenClientKey) as
        | {
            getUserToken?: (
              userId: string,
              connectionName: string,
              channelId: string,
              magicCode?: string,
            ) => Promise<{ token?: string } | undefined>;
          }
        | undefined;
      if (!client?.getUserToken) return false;
      const channelId = context.activity.channelId ?? 'msteams';
      const result = await client.getUserToken(
        userId,
        this.ssoConnectionName,
        channelId,
        magicCode,
      );
      return Boolean(result?.token);
    } catch (err) {
      console.error(
        '[teams] magic-code exchange threw:',
        err instanceof Error ? err.message : err,
      );
      return false;
    }
  }

  /**
   * User clicked a slot button on a `find_free_slots` sidecar card. We
   * synthesize a user message that carries the opaque `slotId` so the
   * LLM recognises the booking intent and invokes `book_meeting(slotId, …)`.
   * The synthetic message also surfaces the human-readable label so the
   * transcript stays meaningful.
   */
  private async handleBookSlotClick(
    context: TurnContext,
    sessionScope: string,
    userId: string | undefined,
    bookSlot: { slotId: string; label: string },
  ): Promise<void> {
    const conversationId = context.activity.conversation?.id ?? 'unknown';
    console.error(
      `[teams] book-slot click conv=${conversationId} slotId=${bookSlot.slotId} label="${truncate(bookSlot.label, 40)}"`,
    );
    this.history.clearPending(sessionScope);
    const priorTurns = this.history.get(sessionScope);
    const synthMessage =
      `Bitte diesen Termin buchen: ${bookSlot.label}. ` +
      `Nutze book_meeting mit slotId="${bookSlot.slotId}" und leite einen sinnvollen Betreff aus dem vorherigen Kontext ab.`;
    await this.runOrchestratorTurn(context, {
      conversationId,
      sessionScope,
      userId,
      userMessage: synthMessage,
      priorTurns,
    });
  }

  private async runOrchestratorTurn(
    context: TurnContext,
    input: {
      conversationId: string;
      sessionScope: string;
      userId: string | undefined;
      userMessage: string;
      priorTurns: ReturnType<ConversationHistoryStore['get']>;
      /** When true, the orchestrator skips context retrieval + memory read. */
      freshCheck?: boolean;
    },
  ): Promise<void> {
    // Scope-local binding so the ALS wrapper can forward the roster accessor
    // to the orchestrator (which surfaces `get_chat_participants` to Claude)
    // and sendAnswer (which resolves `<at>…</at>` tokens back into Mention
    // entities). 1:1 conversations skip ALS entirely — no roster, no mentions.
    const isGroupConversation =
      context.activity.conversation?.conversationType !== 'personal';
    const rosterProvider =
      this.rosterProvider && isGroupConversation ? this.rosterProvider : undefined;

    const run = async (): Promise<void> => {
      const stopTyping = startTypingLoop(context);
      const ssoAssertion = this.ssoConnectionName
        ? await extractSsoAssertion(context, this.ssoConnectionName)
        : undefined;
      console.log(
        `[teams] turn start conv=${input.conversationId} user=${input.userId ?? 'anon'} history=${String(input.priorTurns.length)} freshCheck=${String(Boolean(input.freshCheck))} roster=${rosterProvider ? 'on' : 'off'} sso=${ssoAssertion ? 'yes' : 'no'}`,
      );
      try {
        const userTimeZone =
          (context.activity as { localTimezone?: string }).localTimezone ??
          'Europe/Berlin';
        // Routines-feature hook: install the kernel's per-turn ALS so the
        // `manage_routine` tool can attribute `create` to this (tenant,
        // user) and capture the channel-native delivery handle. Cheap +
        // synchronous; no-op when the kernel didn't wire the callback.
        if (this.captureRoutineTurn && input.userId && this.tenantId) {
          try {
            const conversationRef =
              TurnContext.getConversationReference(context.activity);
            this.captureRoutineTurn({
              tenant: this.tenantId,
              userId: input.userId,
              conversationRef,
            });
          } catch (refErr) {
            // Non-fatal: the rest of the turn proceeds normally; manage_routine
            // create will surface a friendly error if invoked this turn.
            console.error('[teams] routines turn-context capture failed:', refErr);
          }
        }
        const result = await this.orchestrator.chat({
          userMessage: input.userMessage,
          sessionScope: input.sessionScope,
          ...(input.userId ? { userId: input.userId } : {}),
          ...(input.priorTurns.length > 0 ? { priorTurns: input.priorTurns } : {}),
          ...(input.freshCheck ? { freshCheck: true } : {}),
          ...(ssoAssertion ? { ssoAssertion } : {}),
          userTimeZone,
        });
        // Post-S+7.5: ChatAgent.chat returns the SDK's SemanticAnswer. We
        // narrow the discriminated `interactive` union into the legacy
        // pendingUserChoice / pendingSlotCard slots the Teams renderer
        // consumes. `runTrace` is no longer exposed — the trace panel reads
        // `undefined` and degrades gracefully (no section rendered).
        const pendingChoice =
          result.interactive?.kind === 'choice' ? result.interactive : undefined;
        const pendingSlots =
          result.interactive?.kind === 'slots' ? result.interactive : undefined;
        const pendingRoutineList =
          result.interactive?.kind === 'routine_list'
            ? result.interactive
            : undefined;
        await sendAnswer(
          context,
          result.text,
          undefined,
          result.attachments,
          result.verifier,
          input.userMessage,
          rosterProvider,
          pendingChoice,
          result.followUps,
          pendingSlots,
          result.oauthConsentPending ? this.ssoConnectionName : undefined,
        );

        // Routine-list smart card (sidecar). Rendered AFTER the agent's
        // prose answer so the user sees the narration first ("Hier sind
        // deine 3 Routinen") and the actionable card right below.
        if (pendingRoutineList && this.buildRoutineListSmartCardAttachment) {
          const att = this.buildRoutineListSmartCardAttachment({
            filter: pendingRoutineList.filter,
            totals: pendingRoutineList.totals,
            routines: pendingRoutineList.routines,
          });
          await context.sendActivity(MessageFactory.attachment(att));
        }

        // History append skipped for blocking `ask_user_choice` turns —
        // the "answer" is just the clarification prompt, and appending it
        // would pollute the verbatim tail with a non-substantive turn.
        // Follow-up turns (suggest_follow_ups) DO get appended, since the
        // answer is a real response.
        if (!pendingChoice) {
          this.history.append(input.sessionScope, {
            userMessage: input.userMessage,
            assistantAnswer: result.text,
            at: Date.now(),
          });
        }

        const attachmentCount = result.attachments?.length ?? 0;
        console.log(
          `[teams] turn done (attach=${attachmentCount}, history=${String(input.priorTurns.length + 1)}, conv=${input.conversationId}, user=${input.userId ?? 'anon'})`,
        );
      } catch (err) {
        console.error('[teams] orchestrator failure:', err);
        const detail = err instanceof Error ? err.message : String(err);
        await context.sendActivity(
          `Entschuldigung, beim Verarbeiten deiner Anfrage ist ein Fehler aufgetreten: ${detail}. Versuch es gleich nochmal oder stell die Frage anders.`,
        );
      } finally {
        stopTyping();
      }
    };

    if (rosterProvider) {
      await this.turnContext.runWithChatParticipants(
        () => rosterProvider.list(context),
        run,
      );
    } else {
      await run();
    }
  }
}

async function sendTopicAsk(
  context: TurnContext,
  previousSummary: string,
  originalMessage: string,
): Promise<void> {
  const card = buildTopicAskCard({ previousSummary, originalMessage });
  const activity = MessageFactory.attachment(card);
  await context.sendActivity(activity);
}

function truncate(value: string, max: number): string {
  const single = value.replace(/\s+/g, ' ').trim();
  if (single.length <= max) return single;
  return `${single.slice(0, max - 1)}…`;
}

function extractUserMessage(context: TurnContext): string {
  const cleaned = TurnContext.removeRecipientMention(context.activity) ?? '';
  // Teams sends message body twice (plain text + html attachment) for @mentions;
  // removeRecipientMention cleans the plain text, the html duplicate is on the
  // `attachments` array and is ignored here.
  return cleaned.replace(/\s+/g, ' ').trim();
}

function startTypingLoop(context: TurnContext): () => void {
  const sendTyping = (): void => {
    context.sendActivity({ type: ActivityTypes.Typing }).catch((err) => {
      console.warn('[teams] typing send failed:', err instanceof Error ? err.message : err);
    });
  };
  sendTyping();
  const handle = setInterval(sendTyping, TYPING_INTERVAL_MS);
  return () => clearInterval(handle);
}

/**
 * Sends the final turn reply:
 * - Short-answer path (≤ TEAMS_MAX_MESSAGE_CHARS): one Adaptive Card with the
 *   AI label, the answer and — if present — a collapsed tool-trace section.
 *   ToggleVisibility keeps the trace user-local so a group chat stays quiet.
 * - Long-answer fallback: paragraph-split plain text (Adaptive Cards can't
 *   be chunked). AI-label entity is attached to the first message.
 */
async function sendAnswer(
  context: TurnContext,
  text: string,
  runTrace: RunTracePayload | undefined,
  attachments: OutgoingAttachment[] | undefined,
  verifier: VerifierBadge | undefined,
  originalUserMessage: string | undefined,
  rosterProvider: TeamsRosterProvider | undefined,
  pendingUserChoice: OutgoingChoiceCard | undefined,
  followUpOptions: FollowUpOption[] | undefined,
  pendingSlotCard: OutgoingSlotPicker | undefined,
  consentConnectionName: string | undefined,
): Promise<void> {
  // Blocking clarification — render a standalone Choice-Card instead of
  // an answer. The follow-up / fresh-check actions aren't meaningful here
  // since there's no concrete answer yet; attachments are ignored because
  // the orchestrator dropped them in this branch (see drainFollowUps).
  if (pendingUserChoice) {
    const card = buildChoiceAskCard({
      question: pendingUserChoice.question,
      ...(pendingUserChoice.rationale
        ? { rationale: pendingUserChoice.rationale }
        : {}),
      options: pendingUserChoice.options,
    });
    const activity = MessageFactory.attachment(card);
    activity.entities = [...(activity.entities ?? []), aiLabelEntity()];
    await context.sendActivity(activity);
    return;
  }

  const trimmed = text.trim();
  const hasAttachments = (attachments?.length ?? 0) > 0;
  // Images are the whole point when diagrams come back — always use the
  // Adaptive-Card path even if the accompanying text is empty.
  if (trimmed.length === 0 && !hasAttachments) {
    await context.sendActivity('Es ist keine Antwort zurückgekommen. Bitte versuch es erneut.');
    return;
  }

  // Resolve `<at>Display Name</at>` tokens against the roster BEFORE the
  // card/plain-text branching. Resolved mentions get threaded into the
  // Adaptive Card's `msteams.entities` so Teams renders the blue @-pill
  // directly inside the card — no separate plain-text activity on top.
  // Unresolved names get stripped (typo never pings a random person).
  const mentionPrep = await prepareMentions(
    context,
    trimmed.length === 0 && hasAttachments ? '' : trimmed,
    rosterProvider,
  );
  const resolvedText = mentionPrep.text;
  const mentionEntities =
    mentionPrep.kind === 'mentioned' ? mentionPrep.entities : [];

  const answerForCard = resolvedText.length === 0 && hasAttachments
    ? 'Hier das angeforderte Diagramm:'
    : resolvedText;

  if (answerForCard.length <= TEAMS_MAX_MESSAGE_CHARS) {
    const card = buildAnswerCard({
      answer: answerForCard,
      ...(runTrace ? { runTrace } : {}),
      ...(attachments ? { attachments } : {}),
      ...(verifier ? { verifier } : {}),
      ...(originalUserMessage ? { originalUserMessage } : {}),
      ...(mentionEntities.length > 0 ? { mentions: mentionEntities } : {}),
      ...(followUpOptions && followUpOptions.length > 0
        ? { followUpOptions }
        : {}),
    });
    const activity = MessageFactory.attachment(card);
    // Activity-level entities: AI label + mention entities for Teams'
    // notification pipeline. The card also carries the mentions via
    // `msteams.entities` — that's what drives the in-card pill rendering.
    activity.entities = [
      ...(activity.entities ?? []),
      ...mentionEntities,
      aiLabelEntity(),
    ];
    await context.sendActivity(activity);
    if (pendingSlotCard && pendingSlotCard.slots.length > 0) {
      const slotCard = buildSlotPickerCard({
        question: pendingSlotCard.question,
        slots: pendingSlotCard.slots.map((s) => ({
          slotId: s.slotId,
          label: s.label,
          confidence: s.confidence,
        })),
      });
      await context.sendActivity(MessageFactory.attachment(slotCard));
    }
    if (consentConnectionName) {
      const consentCard = await buildConsentSignInCard(context, consentConnectionName);
      await context.sendActivity(MessageFactory.attachment(consentCard));
    }
    return;
  }

  const chunks = splitAtParagraphs(resolvedText, TEAMS_MAX_MESSAGE_CHARS);
  for (let i = 0; i < chunks.length; i++) {
    const chunk = chunks[i];
    if (chunk === undefined) continue;
    const activity = MessageFactory.text(chunk);
    // AI label + mentions on the first chunk only — Teams uses the first
    // message in a split as the canonical one for notifications and labels.
    if (i === 0) {
      activity.entities = [
        ...(activity.entities ?? []),
        ...mentionEntities,
        aiLabelEntity(),
      ];
    }
    await context.sendActivity(activity);
  }

  // Long-answer fallback ate the card, so follow-up buttons need their own
  // mini-card appended after the chunks. Rare path (>25 KB answers); worth
  // the extra activity to keep refinements working.
  if (followUpOptions && followUpOptions.length > 0) {
    const card = buildFollowUpsOnlyCard(followUpOptions);
    await context.sendActivity(MessageFactory.attachment(card));
  }
  if (pendingSlotCard && pendingSlotCard.slots.length > 0) {
    const slotCard = buildSlotPickerCard({
      question: pendingSlotCard.question,
      slots: pendingSlotCard.slots.map((s) => ({
        slotId: s.slotId,
        label: s.label,
        confidence: s.confidence,
      })),
    });
    await context.sendActivity(MessageFactory.attachment(slotCard));
  }
  if (consentConnectionName) {
    const consentCard = await buildConsentSignInCard(context, consentConnectionName);
    await context.sendActivity(MessageFactory.attachment(consentCard));
  }
}

/**
 * Sidecar sign-in card rendered when a calendar tool reported missing consent.
 *
 * Why not `CardFactory.oauthCard`: Teams' OAuthCard flow asks the client to
 * look up the sign-in URL via Bot Framework Token Service internally, which
 * sporadically fails under SingleTenant configurations with a generic
 * "Something went wrong" that gives zero server-side signal. Generating the
 * URL here via `UserTokenClient.getSignInResource` sidesteps that entire
 * handshake — we hand Teams a ready-to-open URL. The redirect from AAD still
 * lands at `token.botframework.com/.auth/web/redirect`, so the subsequent
 * `signin/verifyState` invoke on the bot side works exactly as before.
 *
 * On any failure to resolve the URL (no UserTokenClient in turnState, token
 * service down) we fall back to the stock OAuthCard — worst case we're back
 * to the old failure mode rather than a crashed turn.
 */
async function buildConsentSignInCard(
  context: TurnContext,
  connectionName: string,
): Promise<ReturnType<typeof CardFactory.adaptiveCard> | ReturnType<typeof CardFactory.oauthCard>> {
  const signInUrl = await tryGetSignInUrl(context, connectionName);
  const text =
    'Ich brauche einmalig deine Zustimmung, um auf deinen M365-Kalender zugreifen zu dürfen (freie Slots finden + Termine buchen). Klick auf den Button, bestätige im Tab — dann läuft es im Hintergrund. Zurück in Teams kannst du die Frage direkt nochmal stellen.';
  if (!signInUrl) {
    console.error('[teams] sign-in URL lookup failed — falling back to OAuthCard');
    return CardFactory.oauthCard(connectionName, 'Kalender freigeben', text);
  }
  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body: [
      {
        type: 'TextBlock',
        text: '🔐 Kalender-Zugriff erforderlich',
        weight: 'Bolder',
        size: 'Medium',
      },
      { type: 'TextBlock', text, wrap: true },
    ],
    actions: [
      {
        type: 'Action.OpenUrl',
        title: 'Kalender freigeben',
        url: signInUrl,
        style: 'positive',
      },
    ],
  });
}

async function tryGetSignInUrl(
  context: TurnContext,
  connectionName: string,
): Promise<string | undefined> {
  try {
    const adapter = context.adapter as unknown as { UserTokenClientKey?: symbol };
    if (!adapter.UserTokenClientKey) return undefined;
    const client = context.turnState.get(adapter.UserTokenClientKey) as
      | {
          getSignInResource?: (
            connectionName: string,
            activity: unknown,
            finalRedirect?: string,
          ) => Promise<{ signInLink?: string } | undefined>;
        }
      | undefined;
    if (!client?.getSignInResource) return undefined;
    const resource = await client.getSignInResource(connectionName, context.activity);
    return resource?.signInLink;
  } catch (err) {
    console.error(
      '[teams] getSignInResource failed:',
      err instanceof Error ? err.message : err,
    );
    return undefined;
  }
}

/**
 * Pull a silent Teams SSO assertion for the current user via the Bot
 * Framework `UserTokenClient` that CloudAdapter installs into turnState.
 * Returns `undefined` when:
 *   - the connection is missing/misconfigured,
 *   - the user hasn't consented yet (first-ever call returns no token),
 *   - the adapter shape differs (older BotFrameworkAdapter without the key).
 *
 * The caller treats every `undefined` identically: calendar tools stay
 * dormant and surface `sso_unavailable`. Errors are logged (stderr so Fly
 * keeps them under load) but never thrown — a Graph hiccup mustn't break
 * a regular non-calendar turn.
 */
async function extractSsoAssertion(
  context: TurnContext,
  connectionName: string,
): Promise<string | undefined> {
  try {
    const adapter = context.adapter as unknown as {
      UserTokenClientKey?: symbol;
    };
    if (!adapter.UserTokenClientKey) return undefined;
    const client = context.turnState.get(adapter.UserTokenClientKey) as
      | {
          getUserToken?: (
            userId: string,
            connectionName: string,
            channelId: string,
            magicCode?: string,
          ) => Promise<{ token?: string } | undefined>;
        }
      | undefined;
    if (!client?.getUserToken) return undefined;
    const fromId = context.activity.from?.id;
    const channelId = context.activity.channelId ?? 'msteams';
    if (!fromId) return undefined;
    const res = await client.getUserToken(fromId, connectionName, channelId);
    return res?.token;
  } catch (err) {
    console.error(
      '[teams] sso token extraction failed (continuing without):',
      err instanceof Error ? err.message : err,
    );
    return undefined;
  }
}

/**
 * Inspect the finished answer for `<at>Display Name</at>` tokens, attempt
 * to resolve them via the roster, and decide on the Activity shape:
 *
 * - `kind: 'plain'` — no tokens, or roster unavailable and tokens were
 *   stripped to plain text. Callers stay on the Adaptive-Card path.
 * - `kind: 'mentioned'` — at least one token resolved; send plain-text
 *   activity with matching Mention entities.
 *
 * Unresolved names are stripped and logged — we never forward a raw
 * `<at>…</at>` literal to Teams.
 */
async function prepareMentions(
  context: TurnContext,
  text: string,
  rosterProvider: TeamsRosterProvider | undefined,
): Promise<
  | { kind: 'plain'; text: string }
  | { kind: 'mentioned'; text: string; entities: MentionEntity[] }
> {
  if (text.length === 0 || !rosterProvider) {
    return { kind: 'plain', text: text.length === 0 ? text : stripMentionTokens(text) };
  }
  if (!text.includes('<at>')) {
    return { kind: 'plain', text };
  }
  let roster;
  try {
    roster = await rosterProvider.list(context);
  } catch (err) {
    console.error(
      '[teams] roster fetch failed — stripping mention tokens:',
      err instanceof Error ? err.message : err,
    );
    return { kind: 'plain', text: stripMentionTokens(text) };
  }
  const resolution = resolveMentions(text, roster);
  if (resolution.unresolved.length > 0) {
    console.error(
      `[teams] unresolved mentions stripped: ${resolution.unresolved.join(', ')}`,
    );
  }
  if (resolution.entities.length === 0) {
    return { kind: 'plain', text: stripMentionTokens(text) };
  }
  // Keep the resolved `<at>Name</at>` tokens intact (Teams renders the pill
  // around them), strip only the unresolved ones.
  let outText = text;
  for (const name of resolution.unresolved) {
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    outText = outText.replace(
      new RegExp(`<at>${escaped}</at>`, 'g'),
      name,
    );
  }
  return {
    kind: 'mentioned',
    text: outText,
    entities: resolution.entities,
  };
}

function splitAtParagraphs(text: string, maxLen: number): string[] {
  const chunks: string[] = [];
  let remaining = text;
  while (remaining.length > maxLen) {
    const window = remaining.slice(0, maxLen);
    const lastParagraph = window.lastIndexOf('\n\n');
    // Prefer paragraph break; fall back to line break; fall back to hard cut.
    const lastLine = window.lastIndexOf('\n');
    const cutAt =
      lastParagraph > maxLen / 2 ? lastParagraph : lastLine > maxLen / 2 ? lastLine : maxLen;
    chunks.push(remaining.slice(0, cutAt).trimEnd());
    remaining = remaining.slice(cutAt).trimStart();
  }
  if (remaining.length > 0) chunks.push(remaining);
  return chunks;
}

function isSharePointLike(url: string): boolean {
  try {
    const host = new URL(url).host.toLowerCase();
    return (
      host.endsWith('.sharepoint.com') ||
      host.endsWith('.sharepoint-df.com') ||
      host === '1drv.ms' ||
      host.endsWith('.1drv.ms')
    );
  } catch {
    return false;
  }
}

function fileNameFromUrl(url: string): string | undefined {
  try {
    const u = new URL(url);
    const last = u.pathname.split('/').filter(Boolean).pop();
    if (!last) return undefined;
    const decoded = decodeURIComponent(last);
    return decoded.length > 200 ? decoded.slice(0, 200) : decoded;
  } catch {
    return undefined;
  }
}

/**
 * Flat-walk any object / array, yielding every string value encountered.
 * Used to scavenge Teams-specific channelData + entities for SharePoint
 * URLs without hard-coding each schema variant (which Microsoft changes
 * regularly).
 */
function* collectStringsDeep(value: unknown): Iterable<string> {
  if (typeof value === 'string') {
    yield value;
    return;
  }
  if (Array.isArray(value)) {
    for (const v of value) yield* collectStringsDeep(v);
    return;
  }
  if (value && typeof value === 'object') {
    for (const v of Object.values(value)) yield* collectStringsDeep(v);
  }
}

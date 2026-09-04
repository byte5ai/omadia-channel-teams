import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// #860 W2 (issue #21) — auto-invite onboarding hook. Runtime values via the
// package entry (teamsBot.ts value-imports CJS botbuilder — same rule as
// teamsMultiBotRouting.test.ts); type-only imports relative (erased).
import {
  handleTeamsAgentAppsRecheck,
  runTeamsAutoInviteHook,
  shouldSuppressAutoInstallIntro,
  teamsTeamScopeFromActivity,
  AGENT_APPS_RECHECK_UNSCOPED_MESSAGE,
  AGENT_APPS_RECHECK_VALUE_TYPE,
} from '@omadia/channel-teams';
import type { Activity, Attachment } from 'botbuilder';
import type {
  TeamsAutoInviteDeps,
  TeamsAutoInviteTurnContext,
} from '../src/teamsBot.js';
import type {
  AgentAppInstallOutcome,
  InstallAgentAppsRequest,
} from '../src/teamsAgentInstaller.js';

const BOT_ID = '28:aaaaaaaa-1111-2222-3333-444444444444';
const TEAM_GROUP_ID = 'eeeeeeee-1111-2222-3333-ffffffffffff';
const TENANT_ID = 'cccccccc-4444-5555-6666-dddddddddddd';
const THREAD_ID = '19:team-thread@thread.tacv2';

const INSTALLED_OUTCOME: AgentAppInstallOutcome = {
  kind: 'installed',
  agentSlug: 'odoo-hr',
  displayName: 'Odoo HR',
  teamsAppId: 'catalog-hr',
  outcome: 'created',
};

/** membersAdded conversationUpdate in a TEAM scope, our bot added. */
function teamActivity(over: Record<string, unknown> = {}): Activity {
  return {
    type: 'conversationUpdate',
    membersAdded: [{ id: BOT_ID, name: 'Omadia' }],
    recipient: { id: BOT_ID, name: 'Omadia' },
    from: { id: '29:user-1', name: 'Alice' },
    conversation: {
      id: THREAD_ID,
      conversationType: 'channel',
      tenantId: TENANT_ID,
    },
    channelData: {
      team: { id: THREAD_ID, aadGroupId: TEAM_GROUP_ID },
      tenant: { id: TENANT_ID },
    },
    ...over,
  } as unknown as Activity;
}

function fakeContext(activity: Activity): {
  context: TeamsAutoInviteTurnContext;
  sent: Partial<Activity>[];
  updated: Partial<Activity>[];
  failUpdates: () => void;
} {
  const sent: Partial<Activity>[] = [];
  const updated: Partial<Activity>[] = [];
  let updateError: Error | undefined;
  const context: TeamsAutoInviteTurnContext = {
    activity,
    sendActivity: async (a) => {
      sent.push(a);
      return undefined;
    },
    updateActivity: async (a) => {
      if (updateError) throw updateError;
      updated.push(a);
      return undefined;
    },
  };
  return {
    context,
    sent,
    updated,
    failUpdates: () => {
      updateError = new Error('update rejected');
    },
  };
}

function fakeDeps(opts?: {
  markerFresh?: boolean;
  outcomes?: readonly AgentAppInstallOutcome[];
}): { deps: TeamsAutoInviteDeps; installCalls: InstallAgentAppsRequest[] } {
  const installCalls: InstallAgentAppsRequest[] = [];
  const deps: TeamsAutoInviteDeps = {
    installAgentApps: async (request) => {
      installCalls.push(request);
      const outcomes = opts?.outcomes ?? [INSTALLED_OUTCOME];
      return {
        teamId: request.teamId,
        outcomes,
        autoInstallMarked: outcomes.some(
          (o) => o.kind === 'installed' && o.outcome === 'created',
        ),
      };
    },
    probeAutoInstallMarker: () => opts?.markerFresh ?? false,
  };
  return { deps, installCalls };
}

function firstAttachment(sent: readonly Partial<Activity>[]): Attachment {
  const attachment = sent[0]?.attachments?.[0];
  assert.ok(attachment, 'expected a card attachment');
  return attachment;
}

describe('runTeamsAutoInviteHook — trigger conditions', () => {
  it('runs the installer and posts the result card when OUR bot joins a team', async () => {
    const { deps, installCalls } = fakeDeps();
    const { context, sent } = fakeContext(teamActivity());

    const outcome = await runTeamsAutoInviteHook(deps, context);

    assert.equal(outcome, 'posted-result-card');
    assert.deepEqual(installCalls, [
      { teamId: TEAM_GROUP_ID, tenantId: TENANT_ID },
    ]);
    assert.equal(sent.length, 1);
    assert.equal(
      firstAttachment(sent).contentType,
      'application/vnd.microsoft.card.adaptive',
    );
  });

  it('skips when the added member is a human, not this bot', async () => {
    const { deps, installCalls } = fakeDeps();
    const { context, sent } = fakeContext(
      teamActivity({ membersAdded: [{ id: '29:user-2', name: 'Bob' }] }),
    );

    assert.equal(await runTeamsAutoInviteHook(deps, context), 'skipped');
    assert.deepEqual(installCalls, []);
    assert.deepEqual(sent, []);
  });

  it('skips outside a team scope (no channelData.team)', async () => {
    const { deps, installCalls } = fakeDeps();
    const { context } = fakeContext(
      teamActivity({
        conversation: {
          id: '19:groupchat@thread.skype',
          conversationType: 'groupChat',
          tenantId: TENANT_ID,
        },
        channelData: { tenant: { id: TENANT_ID } },
      }),
    );

    assert.equal(await runTeamsAutoInviteHook(deps, context), 'skipped');
    assert.deepEqual(installCalls, []);
  });

  it('skips when the event carries no aadGroupId — only the Graph group id can drive installToTeam', async () => {
    const { deps, installCalls } = fakeDeps();
    const { context } = fakeContext(
      teamActivity({
        channelData: { team: { id: THREAD_ID }, tenant: { id: TENANT_ID } },
      }),
    );

    assert.equal(await runTeamsAutoInviteHook(deps, context), 'skipped');
    assert.deepEqual(installCalls, []);
  });

  it('suppresses when the auto-install marker is fresh — no second run, no second card', async () => {
    const { deps, installCalls } = fakeDeps({ markerFresh: true });
    const { context, sent } = fakeContext(teamActivity());

    assert.equal(await runTeamsAutoInviteHook(deps, context), 'suppressed');
    assert.deepEqual(installCalls, []);
    assert.deepEqual(sent, []);
  });

  it('posts nothing when the installer answers zero outcomes (feature effectively off)', async () => {
    const { deps } = fakeDeps({ outcomes: [] });
    const { context, sent } = fakeContext(teamActivity());

    assert.equal(await runTeamsAutoInviteHook(deps, context), 'skipped');
    assert.deepEqual(sent, []);
  });
});

describe('teamsTeamScopeFromActivity', () => {
  it('extracts the Graph group id + tenant id from a team activity', () => {
    assert.deepEqual(teamsTeamScopeFromActivity(teamActivity()), {
      teamId: TEAM_GROUP_ID,
      tenantId: TENANT_ID,
    });
  });

  it('falls back to conversation.tenantId when channelData has no tenant', () => {
    const activity = teamActivity({
      channelData: { team: { id: THREAD_ID, aadGroupId: TEAM_GROUP_ID } },
    });
    assert.deepEqual(teamsTeamScopeFromActivity(activity), {
      teamId: TEAM_GROUP_ID,
      tenantId: TENANT_ID,
    });
  });

  it('returns undefined for personal scope', () => {
    const activity = teamActivity({
      conversation: {
        id: 'a:1personal',
        conversationType: 'personal',
        tenantId: TENANT_ID,
      },
      channelData: { tenant: { id: TENANT_ID } },
    });
    assert.equal(teamsTeamScopeFromActivity(activity), undefined);
  });
});

describe('shouldSuppressAutoInstallIntro', () => {
  it('true only for: probe wired + marker fresh + this bot added + team scope', () => {
    const probe = (teamId: string): boolean => teamId === TEAM_GROUP_ID;
    assert.equal(shouldSuppressAutoInstallIntro(probe, teamActivity()), true);
  });

  it('false without an auto-invite probe (feature off)', () => {
    assert.equal(
      shouldSuppressAutoInstallIntro(undefined, teamActivity()),
      false,
    );
  });

  it('false when the marker is cold', () => {
    assert.equal(
      shouldSuppressAutoInstallIntro(() => false, teamActivity()),
      false,
    );
  });

  it('false when the added member is not this bot', () => {
    const activity = teamActivity({
      membersAdded: [{ id: '29:user-2', name: 'Bob' }],
    });
    assert.equal(shouldSuppressAutoInstallIntro(() => true, activity), false);
  });

  it('false outside a team scope', () => {
    const activity = teamActivity({
      channelData: { tenant: { id: TENANT_ID } },
    });
    assert.equal(shouldSuppressAutoInstallIntro(() => true, activity), false);
  });
});

describe('handleTeamsAgentAppsRecheck', () => {
  /**
   * The "Prüfen" submit activity — a message carrying replyToId. Its
   * `value` still models what a REAL client sends, tampering included:
   * the handler takes no payload argument at all any more (#1030), so a
   * forged `teamId`/`tenantId` in here has no route to the installer.
   */
  function recheckActivity(over: Record<string, unknown> = {}): Activity {
    return teamActivity({
      type: 'message',
      membersAdded: undefined,
      replyToId: 'card-activity-1',
      value: { type: AGENT_APPS_RECHECK_VALUE_TYPE },
      ...over,
    });
  }

  it('re-runs the installer and UPDATES the existing card in place', async () => {
    const { deps, installCalls } = fakeDeps();
    const { context, sent, updated } = fakeContext(recheckActivity());

    assert.equal(
      await handleTeamsAgentAppsRecheck(deps, context),
      'updated-card',
    );

    assert.deepEqual(installCalls, [
      { teamId: TEAM_GROUP_ID, tenantId: TENANT_ID },
    ]);
    assert.equal(updated.length, 1);
    assert.equal(updated[0]?.id, 'card-activity-1');
    assert.ok(updated[0]?.attachments?.[0]);
    assert.deepEqual(sent, []);
  });

  it('#1030 — installs into the ACTIVITY team even when the payload names a foreign team and tenant', async () => {
    const { deps, installCalls } = fakeDeps();
    const { context } = fakeContext(
      recheckActivity({
        value: {
          type: AGENT_APPS_RECHECK_VALUE_TYPE,
          teamId: 'ffffffff-0000-0000-0000-000000000bad',
          tenantId: 'ffffffff-0000-0000-0000-00000000dead',
        },
      }),
    );

    await handleTeamsAgentAppsRecheck(deps, context);

    assert.deepEqual(installCalls, [
      { teamId: TEAM_GROUP_ID, tenantId: TENANT_ID },
    ]);
  });

  it('#1030 — REFUSES a click whose activity names no team scope, instead of trusting the payload', async () => {
    const { deps, installCalls } = fakeDeps();
    const { context, sent, updated } = fakeContext(
      recheckActivity({
        // A replayed / hand-crafted click: personal scope on the wire,
        // a team named only in the client-editable card data.
        conversation: {
          id: 'a:1personal',
          conversationType: 'personal',
          tenantId: TENANT_ID,
        },
        channelData: { tenant: { id: TENANT_ID } },
        value: {
          type: AGENT_APPS_RECHECK_VALUE_TYPE,
          teamId: 'ffffffff-0000-0000-0000-000000000bad',
          tenantId: 'ffffffff-0000-0000-0000-00000000dead',
        },
      }),
    );

    assert.equal(
      await handleTeamsAgentAppsRecheck(deps, context),
      'refused-unscoped',
    );

    // The whole point: no install ran, against ANY team.
    assert.deepEqual(installCalls, []);
    assert.deepEqual(updated, []);
    assert.equal(sent.length, 1);
    assert.equal(sent[0]?.text, AGENT_APPS_RECHECK_UNSCOPED_MESSAGE);
    assert.equal(sent[0]?.attachments, undefined);
  });

  it('#1030 — refuses a team-scope click that carries no Graph group id', async () => {
    const { deps, installCalls } = fakeDeps();
    const { context, sent } = fakeContext(
      recheckActivity({
        // `channelData.team.id` is the `19:…` thread id, which Graph's
        // /teams/{id}/installedApps will not accept — and the payload is
        // no longer allowed to supply the group id instead.
        channelData: { team: { id: THREAD_ID }, tenant: { id: TENANT_ID } },
        value: {
          type: AGENT_APPS_RECHECK_VALUE_TYPE,
          teamId: TEAM_GROUP_ID,
          tenantId: TENANT_ID,
        },
      }),
    );

    assert.equal(
      await handleTeamsAgentAppsRecheck(deps, context),
      'refused-unscoped',
    );
    assert.deepEqual(installCalls, []);
    assert.equal(sent[0]?.text, AGENT_APPS_RECHECK_UNSCOPED_MESSAGE);
  });

  it('posts a fresh card when the activity update is rejected', async () => {
    const { deps } = fakeDeps();
    const { context, sent, failUpdates } = fakeContext(recheckActivity());
    failUpdates();

    assert.equal(
      await handleTeamsAgentAppsRecheck(deps, context),
      'posted-card',
    );
    assert.equal(sent.length, 1);
    assert.ok(firstAttachment(sent));
  });

  it('posts a fresh card when the submit carries no replyToId', async () => {
    const { deps } = fakeDeps();
    const { context, sent, updated } = fakeContext(
      recheckActivity({ replyToId: undefined }),
    );

    assert.equal(
      await handleTeamsAgentAppsRecheck(deps, context),
      'posted-card',
    );
    assert.deepEqual(updated, []);
    assert.equal(sent.length, 1);
  });
});

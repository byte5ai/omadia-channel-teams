import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// #860 W0a — legacy scalar-credential shim: existing single-bot deployments
// (microsoft_app_id / microsoft_tenant_id / vault microsoft_app_password /
// env MICROSOFT_APP_TYPE) map onto teams_bots[0] with zero config changes.
import {
  getDefaultTeamsBot,
  LEGACY_TEAMS_BOT_DISPLAY_NAME,
  LEGACY_TEAMS_BOT_SECRET_REF,
  LEGACY_TEAMS_BOT_SLUG,
  legacyTeamsBotFromScalars,
  teamsBotKey,
} from '@omadia/channel-teams';
import type { LegacyTeamsBotScalars } from '../src/teamsBotsConfig.js';

const SCALARS: LegacyTeamsBotScalars = {
  appId: '11111111-aaaa-bbbb-cccc-222222222222',
  tenantId: 'tenant-legacy',
};

describe('legacyTeamsBotFromScalars — teams_bots[0] mapping', () => {
  it('maps the scalar credentials onto a complete bot identity', () => {
    const bot = legacyTeamsBotFromScalars(SCALARS);
    assert.equal(bot.appId, SCALARS.appId);
    assert.equal(bot.tenantId, SCALARS.tenantId);
    assert.equal(bot.appPasswordSecretRef, 'microsoft_app_password');
    assert.equal(bot.appPasswordSecretRef, LEGACY_TEAMS_BOT_SECRET_REF);
  });

  it('is teams_bots[0] — the default bot the /api/messages aliases serve', () => {
    const bot = legacyTeamsBotFromScalars(SCALARS);
    assert.equal(getDefaultTeamsBot([bot]), bot);
  });

  it('uses a deterministic botSlug so route alias + ref-store backfill agree across restarts', () => {
    const first = legacyTeamsBotFromScalars(SCALARS);
    const second = legacyTeamsBotFromScalars(SCALARS);
    assert.equal(first.botSlug, LEGACY_TEAMS_BOT_SLUG);
    assert.equal(first.botSlug, second.botSlug);
    // The directory catch-all key derives from the same identity.
    assert.equal(teamsBotKey(first.appId), `28:${SCALARS.appId}`);
  });
});

describe('legacyTeamsBotFromScalars — MICROSOFT_APP_TYPE (shimmed bot ONLY)', () => {
  it('defaults to MultiTenant when the env knob is unset — existing deployments were provisioned under that default', () => {
    assert.equal(legacyTeamsBotFromScalars(SCALARS).appType, 'MultiTenant');
    assert.equal(
      legacyTeamsBotFromScalars({ ...SCALARS, appTypeEnv: undefined }).appType,
      'MultiTenant',
    );
  });

  it('defaults to MultiTenant for unknown values (historical readTeamsConfigFromEnv semantics)', () => {
    assert.equal(
      legacyTeamsBotFromScalars({ ...SCALARS, appTypeEnv: 'PublicClient' }).appType,
      'MultiTenant',
    );
    assert.equal(
      legacyTeamsBotFromScalars({ ...SCALARS, appTypeEnv: '' }).appType,
      'MultiTenant',
    );
  });

  it('honours an explicit SingleTenant / UserAssignedMSI', () => {
    assert.equal(
      legacyTeamsBotFromScalars({ ...SCALARS, appTypeEnv: 'SingleTenant' }).appType,
      'SingleTenant',
    );
    assert.equal(
      legacyTeamsBotFromScalars({ ...SCALARS, appTypeEnv: 'UserAssignedMSI' }).appType,
      'UserAssignedMSI',
    );
  });
});

describe('legacyTeamsBotFromScalars — displayName', () => {
  it('reuses the operator label (teams_directory_label) when present', () => {
    const bot = legacyTeamsBotFromScalars({
      ...SCALARS,
      displayName: 'Teams · Production · Marketing',
    });
    assert.equal(bot.displayName, 'Teams · Production · Marketing');
  });

  it('falls back to a stable default name when no label is set', () => {
    assert.equal(
      legacyTeamsBotFromScalars(SCALARS).displayName,
      LEGACY_TEAMS_BOT_DISPLAY_NAME,
    );
    assert.equal(
      legacyTeamsBotFromScalars({ ...SCALARS, displayName: '   ' }).displayName,
      LEGACY_TEAMS_BOT_DISPLAY_NAME,
    );
  });
});

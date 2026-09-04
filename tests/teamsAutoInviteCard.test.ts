import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

// #860 W2 (issue #21) — auto-invite result / fallback card. Runtime values
// via the package entry (teamsCard.ts value-imports CJS botbuilder — same
// rule as teamsCard.test.ts); type-only imports relative (erased).
import {
  buildAgentAppsResultCard,
  parseAgentAppsRecheckValue,
  AGENT_APPS_RECHECK_VALUE_TYPE,
  TEAMS_CARD_CONTENT_TYPE,
} from '@omadia/channel-teams';
import type { AgentAppInstallOutcome } from '../src/teamsAgentInstaller.js';

const TEAM_ID = 'eeeeeeee-1111-2222-3333-ffffffffffff';
const TENANT_ID = 'cccccccc-4444-5555-6666-dddddddddddd';

const HR_INSTALLED: AgentAppInstallOutcome = {
  kind: 'installed',
  agentSlug: 'odoo-hr',
  displayName: 'Odoo HR',
  teamsAppId: 'catalog-hr',
  outcome: 'created',
};

const ACCOUNTING_EXISTED: AgentAppInstallOutcome = {
  kind: 'installed',
  agentSlug: 'odoo-accounting',
  displayName: 'Odoo Accounting',
  teamsAppId: 'catalog-accounting',
  outcome: 'already-existed',
};

const CONSENT_FALLBACK: AgentAppInstallOutcome = {
  kind: 'fallback',
  agentSlug: 'odoo-hr',
  displayName: 'Odoo HR',
  reason: 'consent-missing',
  teamsAppExternalId: 'ext-hr',
  teamsAppId: 'catalog-hr',
  missingScopes: ['TeamsAppInstallation.ReadWriteForTeam.All'],
};

const CATALOG_FALLBACK: AgentAppInstallOutcome = {
  kind: 'fallback',
  agentSlug: 'odoo-accounting',
  displayName: 'Odoo Accounting',
  reason: 'not-in-catalog',
  teamsAppExternalId: 'ext-accounting',
};

interface CardShape {
  body: Array<Record<string, unknown>>;
  actions?: Array<Record<string, unknown>>;
}

function build(outcomes: readonly AgentAppInstallOutcome[]): CardShape {
  const attachment = buildAgentAppsResultCard({ outcomes });
  assert.equal(attachment.contentType, TEAMS_CARD_CONTENT_TYPE);
  return attachment.content as CardShape;
}

/** Every TextBlock text of the card body, flattened through containers. */
function bodyTexts(card: CardShape): string[] {
  const texts: string[] = [];
  const walk = (items: Array<Record<string, unknown>>): void => {
    for (const item of items) {
      if (typeof item['text'] === 'string') texts.push(item['text']);
      if (Array.isArray(item['items'])) {
        walk(item['items'] as Array<Record<string, unknown>>);
      }
    }
  };
  walk(card.body);
  return texts;
}

describe('buildAgentAppsResultCard — snapshot shape', () => {
  it('all installed: one ✅ line per app, no actions at all', () => {
    const card = build([HR_INSTALLED, ACCOUNTING_EXISTED]);
    const texts = bodyTexts(card);

    assert.equal(texts[0], '🤝 Agenten-Apps für dieses Team');
    assert.ok(texts.includes('✅ **Odoo HR** — installiert'));
    assert.ok(
      texts.includes('✅ **Odoo Accounting** — war bereits installiert'),
    );
    assert.equal(card.actions, undefined);
  });

  it('consent-missing: scope list, deep-link OpenUrl and Prüfen submit', () => {
    const card = build([CONSENT_FALLBACK]);
    const texts = bodyTexts(card);

    assert.ok(texts.includes('⚠️ **Odoo HR** — Admin-Zustimmung fehlt'));
    assert.ok(
      texts.some((t) =>
        t.includes('`TeamsAppInstallation.ReadWriteForTeam.All`'),
      ),
      'missing scopes are card material',
    );

    const actions = card.actions ?? [];
    const openUrl = actions.find((a) => a['type'] === 'Action.OpenUrl');
    assert.ok(openUrl, 'expected an install deep link');
    assert.equal(
      openUrl['url'],
      'https://teams.microsoft.com/l/app/catalog-hr',
    );
    assert.equal(openUrl['title'], 'Odoo HR öffnen');

    const submit = actions.find((a) => a['type'] === 'Action.Submit');
    assert.ok(submit, 'expected the Prüfen re-check');
    assert.equal(submit['title'], '🔄 Prüfen');
    const data = submit['data'] as Record<string, unknown>;
    assert.equal(data['type'], AGENT_APPS_RECHECK_VALUE_TYPE);
    // #1030 — the card must NOT round-trip an install target. Card `data`
    // is client-editable, so anything named here is a target the clicker
    // chooses; the handler derives team + tenant from the submit
    // activity's channelData instead.
    assert.equal(data['teamId'], undefined);
    assert.equal(data['tenantId'], undefined);
    assert.deepEqual(Object.keys(data).sort(), ['msteams', 'type']);
    // Teams quirk guard — without messageBack the client shows
    // "Something went wrong" on slow responses.
    assert.equal(
      (data['msteams'] as Record<string, unknown>)['type'],
      'messageBack',
    );
  });

  it('consent-cached renders like consent-missing (same operator action)', () => {
    const card = build([{ ...CONSENT_FALLBACK, reason: 'consent-cached' }]);
    assert.ok(
      bodyTexts(card).includes('⚠️ **Odoo HR** — Admin-Zustimmung fehlt'),
    );
  });

  it('not-in-catalog is distinguished from consent-missing and has no deep link', () => {
    const card = build([CATALOG_FALLBACK]);
    const texts = bodyTexts(card);

    assert.ok(
      texts.includes(
        '⚠️ **Odoo Accounting** — nicht im Teams-App-Katalog gefunden',
      ),
    );
    assert.ok(
      texts.some((t) => t.includes('Teams-App-Katalog der Organisation')),
      'expected the catalog-upload hint',
    );
    assert.ok(
      !texts.some((t) => t.includes('Admin-Zustimmung fehlt')),
      'no consent wording without a consent fallback',
    );
    const actions = card.actions ?? [];
    assert.equal(
      actions.find((a) => a['type'] === 'Action.OpenUrl'),
      undefined,
      'no deep link without a catalog id',
    );
    assert.ok(actions.find((a) => a['type'] === 'Action.Submit'));
  });

  it('failed outcomes: throttled ⏳ vs error ❌ wording', () => {
    const card = build([
      {
        kind: 'failed',
        agentSlug: 'odoo-hr',
        displayName: 'Odoo HR',
        reason: 'throttled',
        retryAfterSeconds: 60,
        message: 'provisioning_throttled',
      },
      {
        kind: 'failed',
        agentSlug: 'odoo-accounting',
        displayName: 'Odoo Accounting',
        reason: 'error',
        message: 'boom',
      },
    ]);
    const texts = bodyTexts(card);

    assert.ok(
      texts.includes(
        '⏳ **Odoo HR** — Microsoft drosselt gerade, bitte später erneut prüfen',
      ),
    );
    assert.ok(
      texts.includes('❌ **Odoo Accounting** — Installation fehlgeschlagen'),
    );
    // The raw error message never leaks into the card.
    assert.ok(!texts.some((t) => t.includes('boom')));
  });

  it('mixed run keeps configured order and offers exactly one Prüfen', () => {
    const card = build([HR_INSTALLED, CATALOG_FALLBACK]);
    const lines = bodyTexts(card).filter(
      (t) => t.startsWith('✅') || t.startsWith('⚠️'),
    );
    assert.equal(lines[0]?.includes('Odoo HR'), true);
    assert.equal(lines[1]?.includes('Odoo Accounting'), true);
    const submits = (card.actions ?? []).filter(
      (a) => a['type'] === 'Action.Submit',
    );
    assert.equal(submits.length, 1);
  });
});

describe('parseAgentAppsRecheckValue', () => {
  const valid = { type: AGENT_APPS_RECHECK_VALUE_TYPE };

  it('accepts the object form', () => {
    assert.deepEqual(parseAgentAppsRecheckValue({ ...valid }), valid);
  });

  it('accepts the JSON-string form some Teams clients deliver', () => {
    assert.deepEqual(parseAgentAppsRecheckValue(JSON.stringify(valid)), valid);
  });

  it('#1030 — drops every payload key but the discriminator', () => {
    // Cards posted BEFORE this change still carry teamId/tenantId, so the
    // parser must keep accepting them — and must not carry them forward,
    // which is what kept them out of the install path in the first place.
    assert.deepEqual(
      parseAgentAppsRecheckValue({
        type: AGENT_APPS_RECHECK_VALUE_TYPE,
        teamId: 'ffffffff-0000-0000-0000-000000000bad',
        tenantId: 'ffffffff-0000-0000-0000-00000000dead',
      }),
      valid,
    );
  });

  it('rejects foreign submit types and malformed payloads', () => {
    assert.equal(parseAgentAppsRecheckValue(undefined), undefined);
    assert.equal(parseAgentAppsRecheckValue('not json{'), undefined);
    assert.equal(parseAgentAppsRecheckValue(42), undefined);
    assert.equal(
      parseAgentAppsRecheckValue({ ...valid, type: 'fresh_check' }),
      undefined,
    );
    assert.equal(parseAgentAppsRecheckValue({}), undefined);
  });
});

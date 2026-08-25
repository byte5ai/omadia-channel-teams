import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * #19 (epic byte5ai/omadia#860, W0) — versioned Teams appPackage template.
 *
 * The Teams app package used to be an unversioned, hand-cut artifact in the
 * middleware deployment. This template versions it and is the input to the
 * provisioner's per-agent manifest generator. Two things must hold:
 *
 *  1. **Nothing moves for single-bot deployments.** Rendered with the legacy
 *     values, the template must reproduce today's shipped manifest exactly —
 *     same fields, same values, same RSC grants.
 *  2. **Everything per-bot is a placeholder.** In particular the Teams app id
 *     (`id`) and the bot's Entra app id (`bots[].botId` / `webApplicationInfo`)
 *     are DIFFERENT GUIDs — mapping both to one placeholder breaks the upload.
 */

// Works from tests/ and from the transpiled .test-build/ — both sit directly
// under the package root.
const pkgRoot = join(dirname(fileURLToPath(import.meta.url)), '..');
const appPackageDir = join(pkgRoot, 'appPackage');
const templatePath = join(appPackageDir, 'manifest.json.template');

/** Every placeholder the template may use. Anything else is a typo. */
const KNOWN_PLACEHOLDERS = [
  'VERSION',
  'APP_ID',
  'BOT_ID',
  'NAME_SHORT',
  'NAME_FULL',
  'DESCRIPTION',
  'DESCRIPTION_FULL',
  'ACCENT_COLOR',
  'TAB_BASE_URL',
  'VALID_DOMAINS',
  'MIDDLEWARE_HOST',
];

/** The six placeholders the spec names explicitly. */
const SPEC_REQUIRED_PLACEHOLDERS = [
  'APP_ID',
  'BOT_ID',
  'NAME_SHORT',
  'NAME_FULL',
  'DESCRIPTION',
  'VALID_DOMAINS',
];

/**
 * Substitute `{{NAME}}` tokens. String values are inserted verbatim into the
 * surrounding JSON string literal; VALID_DOMAINS is a raw JSON array and is
 * inserted where the template deliberately leaves the value unquoted.
 */
function render(template: string, values: Record<string, string | string[]>): string {
  return template.replace(/\{\{([A-Z0-9_]+)\}\}/g, (_match, name: string) => {
    const value = values[name];
    if (value === undefined) throw new Error(`no value for placeholder {{${name}}}`);
    return Array.isArray(value) ? JSON.stringify(value) : value;
  });
}

/** The values of today's shipped single-bot manifest (middleware v1.3.0). */
const LEGACY_VALUES: Record<string, string | string[]> = {
  VERSION: '1.3.0',
  APP_ID: 'bc0bd6cf-7037-4c7c-9e29-de86c4b77177',
  BOT_ID: '737c6ddd-6d4e-4599-8dc3-260281ea906e',
  NAME_SHORT: 'omadia-agent',
  NAME_FULL: 'omadia-agent — byte5 Assistent',
  DESCRIPTION: 'Assistent für Odoo, HR, Confluence & Diagramme.',
  DESCRIPTION_FULL:
    'omadia-agent verbindet Microsoft Teams mit einem Claude-Agent, der über spezialisierte Fach-Agenten auf eure byte5-Systeme zugreift: Odoo-17 (Accounting + HR), Confluence-Playbook, Diagramm-Rendering. Stellt Fragen auf Deutsch — der Bot delegiert an den richtigen Agenten, verifiziert die Antwort gegen die Quelle und merkt sich wiederkehrende Konventionen über Sessions hinweg.',
  ACCENT_COLOR: '#714B67',
  TAB_BASE_URL: 'https://odoo-bot-harness.fly.dev/p/channel-teams',
  VALID_DOMAINS: ['odoo-bot-middleware.fly.dev', 'odoo-bot-harness.fly.dev'],
  MIDDLEWARE_HOST: 'odoo-bot-middleware.fly.dev',
};

const RSC_GRANTS = [
  'ChannelMessage.Read.Group',
  'ChatMessage.Read.Chat',
  'TeamsActivity.Send.Group',
  'TeamsActivity.Send.Chat',
  'FileStorageContainer.Selected.Group',
  'ChatMember.Read.Chat',
  'ChannelMember.Read.Group',
];

describe('#19 W0 — appPackage/manifest.json.template', () => {
  const template = readFileSync(templatePath, 'utf8');

  it('contains the six spec-required placeholders', () => {
    for (const name of SPEC_REQUIRED_PLACEHOLDERS) {
      assert.ok(template.includes(`{{${name}}}`), `missing placeholder {{${name}}}`);
    }
  });

  it('uses only known placeholders — no typos, no undocumented tokens', () => {
    const used = [...template.matchAll(/\{\{([A-Z0-9_]+)\}\}/g)].map((m) => m[1]);
    assert.ok(used.length > 0, 'template has no placeholders at all');
    for (const name of used) {
      assert.ok(KNOWN_PLACEHOLDERS.includes(name), `unknown placeholder {{${name}}}`);
    }
  });

  it('every known placeholder actually appears (dead entries drift)', () => {
    for (const name of KNOWN_PLACEHOLDERS) {
      assert.ok(template.includes(`{{${name}}}`), `documented placeholder {{${name}}} unused`);
    }
  });

  it('rendered with legacy values it parses as JSON with zero leftover tokens', () => {
    const rendered = render(template, LEGACY_VALUES);
    assert.ok(!/\{\{[A-Z0-9_]+\}\}/.test(rendered), 'unsubstituted placeholder left behind');
    JSON.parse(rendered); // throws on invalid JSON
  });

  it('rendered with legacy values it reproduces today\'s shipped manifest', () => {
    const manifest = JSON.parse(render(template, LEGACY_VALUES));

    // Pinned, never parameterized.
    assert.equal(
      manifest.$schema,
      'https://developer.microsoft.com/json-schemas/teams/v1.17/MicrosoftTeams.schema.json',
    );
    assert.equal(manifest.manifestVersion, '1.17');

    // Identity — APP_ID and BOT_ID are DIFFERENT GUIDs, mapped to different fields.
    assert.equal(manifest.version, '1.3.0');
    assert.equal(manifest.id, LEGACY_VALUES.APP_ID);
    assert.equal(manifest.bots[0].botId, LEGACY_VALUES.BOT_ID);
    assert.notEqual(manifest.id, manifest.bots[0].botId);

    // Naming + description.
    assert.deepEqual(manifest.name, {
      short: 'omadia-agent',
      full: 'omadia-agent — byte5 Assistent',
    });
    assert.equal(manifest.description.short, LEGACY_VALUES.DESCRIPTION);
    assert.equal(manifest.description.full, LEGACY_VALUES.DESCRIPTION_FULL);
    assert.equal(manifest.accentColor, '#714B67');

    // Developer block stays literal.
    assert.deepEqual(manifest.developer, {
      name: 'byte5',
      websiteUrl: 'https://omadia.ai',
      privacyUrl: 'https://omadia.ai/privacy',
      termsOfUseUrl: 'https://omadia.ai/terms',
    });

    // Bot shape + the German default command list survive rendering.
    assert.equal(manifest.bots.length, 1);
    assert.deepEqual(manifest.bots[0].scopes, ['personal', 'team', 'groupChat']);
    assert.equal(manifest.bots[0].supportsFiles, true);
    assert.equal(manifest.bots[0].isNotificationOnly, false);
    const commands = manifest.bots[0].commandLists[0].commands.map((c: { title: string }) => c.title);
    assert.deepEqual(commands, ['Umsatz', 'Offene Posten', 'Urlaub', 'Diagramm', 'Playbook', 'Termin']);

    // Icons reference the sibling files by name.
    assert.deepEqual(manifest.icons, { color: 'color.png', outline: 'outline.png' });

    // Web-ui surfaces (src/uiRouter.ts) — tab URLs derive from TAB_BASE_URL.
    assert.equal(
      manifest.staticTabs[0].contentUrl,
      'https://odoo-bot-harness.fly.dev/p/channel-teams/hub',
    );
    assert.equal(manifest.staticTabs[0].websiteUrl, manifest.staticTabs[0].contentUrl);
    assert.equal(
      manifest.configurableTabs[0].configurationUrl,
      'https://odoo-bot-harness.fly.dev/p/channel-teams/tab-config',
    );

    // Security boundary: iframe origins + SSO audience.
    assert.deepEqual(manifest.validDomains, [
      'odoo-bot-middleware.fly.dev',
      'odoo-bot-harness.fly.dev',
    ]);
    assert.deepEqual(manifest.webApplicationInfo, {
      id: LEGACY_VALUES.BOT_ID,
      resource: `api://odoo-bot-middleware.fly.dev/${LEGACY_VALUES.BOT_ID}`,
    });

    // RSC grants ship verbatim — a generator must never drop these.
    const grants = manifest.authorization.permissions.resourceSpecific.map(
      (g: { type: string; name: string }) => g.name,
    );
    assert.deepEqual(grants, RSC_GRANTS);
    for (const g of manifest.authorization.permissions.resourceSpecific) {
      assert.equal(g.type, 'Application');
    }

    assert.deepEqual(manifest.permissions, ['identity', 'messageTeamMembers']);
    assert.equal(manifest.activities.activityTypes[0].type, 'pluginEvent');
  });

  it('a second bot identity renders to a distinct, well-formed manifest', () => {
    const manifest = JSON.parse(
      render(template, {
        ...LEGACY_VALUES,
        VERSION: '1.0.0',
        APP_ID: '11111111-2222-3333-4444-555555555555',
        BOT_ID: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
        NAME_SHORT: 'sales-agent',
        NAME_FULL: 'sales-agent — Omadia',
        DESCRIPTION: 'Sales assistant.',
        DESCRIPTION_FULL: 'Sales assistant for the demo tenant.',
        VALID_DOMAINS: ['example-middleware.fly.dev'],
        MIDDLEWARE_HOST: 'example-middleware.fly.dev',
        TAB_BASE_URL: 'https://example-harness.fly.dev/p/channel-teams',
      }),
    );
    assert.equal(manifest.id, '11111111-2222-3333-4444-555555555555');
    assert.equal(manifest.bots[0].botId, 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee');
    assert.equal(manifest.webApplicationInfo.id, 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee');
    assert.equal(
      manifest.webApplicationInfo.resource,
      'api://example-middleware.fly.dev/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
    );
    assert.deepEqual(manifest.validDomains, ['example-middleware.fly.dev']);
    // The pinned schema + RSC grants are identical for every generated bot.
    assert.equal(manifest.manifestVersion, '1.17');
    assert.equal(manifest.authorization.permissions.resourceSpecific.length, RSC_GRANTS.length);
  });
});

describe('#19 W0 — appPackage icons + README', () => {
  const PNG_MAGIC = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

  it('ships color.png and outline.png as real PNGs', () => {
    for (const name of ['color.png', 'outline.png']) {
      const buf = readFileSync(join(appPackageDir, name));
      assert.ok(buf.length > PNG_MAGIC.length, `${name} is empty`);
      assert.deepEqual(buf.subarray(0, PNG_MAGIC.length), PNG_MAGIC, `${name} is not a PNG`);
    }
  });

  it('README documents every placeholder the template uses', () => {
    const readme = readFileSync(join(appPackageDir, 'README.md'), 'utf8');
    for (const name of KNOWN_PLACEHOLDERS) {
      assert.ok(readme.includes(`{{${name}}}`), `README does not document {{${name}}}`);
    }
  });
});

/**
 * #860 W0a — the template must SHIP, not just exist. The release ZIP is what
 * the provisioner installs, so a template that lives in the repo but not in
 * `out/<id>-<version>.zip` is still the unversioned manual artifact this wave
 * retires. These tests hold the packaging script to that, plus the two-file
 * version invariant its drift guard enforces.
 */
describe('#860 W0a — release artifact ships the appPackage template', () => {
  const pkgJson = JSON.parse(readFileSync(join(pkgRoot, 'package.json'), 'utf8')) as {
    name: string;
    version: string;
  };

  it('package.json and manifest.yaml carry the same semver (drift-guard input)', () => {
    const manifestYaml = readFileSync(join(pkgRoot, 'manifest.yaml'), 'utf8');
    // Same expression build-zip.mjs uses to read identity.version.
    const manifestVersion = manifestYaml.match(/^\s{2}version:\s*["']?([^"'\s]+)/m)?.[1];
    assert.match(pkgJson.version, /^\d+\.\d+\.\d+$/, 'package.json version is not plain semver');
    assert.equal(
      manifestVersion,
      pkgJson.version,
      `version drift: package.json says ${pkgJson.version}, manifest.yaml says ${manifestVersion}`,
    );
  });

  it('build-zip.mjs stages appPackage/ as a REQUIRED dir, not an optional one', () => {
    const script = readFileSync(join(pkgRoot, 'scripts', 'build-zip.mjs'), 'utf8');
    assert.match(
      script,
      /REQUIRED_DIRS\s*=\s*\[[^\]]*'appPackage'/,
      'appPackage missing from REQUIRED_DIRS — the template would silently drop out of the ZIP',
    );
  });

  it('npm run package produces a flat artifact containing the template', () => {
    // Runs the real packaging script (which also exercises the drift guard).
    // Requires `npm run build` first — same precondition as the whole suite.
    const res = spawnSync(process.execPath, [join(pkgRoot, 'scripts', 'build-zip.mjs')], {
      cwd: pkgRoot,
      encoding: 'utf8',
    });
    assert.equal(res.status, 0, `build-zip.mjs failed:\n${res.stdout}\n${res.stderr}`);

    const safeName = pkgJson.name.replace(/^@/, '').replace(/\//g, '-');
    const stageDir = join(pkgRoot, 'out', `${safeName}-${pkgJson.version}-stage`);
    const zipPath = join(pkgRoot, 'out', `${safeName}-${pkgJson.version}.zip`);
    assert.ok(existsSync(zipPath), `missing release ZIP: ${zipPath}`);

    // The stage dir mirrors the archive root (createFlatZip archives its
    // CONTENTS), so asserting on it asserts the ZIP layout: FLAT, with the
    // template directory alongside manifest.yaml and dist/.
    assert.ok(existsSync(join(stageDir, 'manifest.yaml')), 'stage is not flat: manifest.yaml not at root');
    assert.ok(existsSync(join(stageDir, 'dist', 'plugin.js')), 'stage is missing dist/plugin.js');
    for (const rel of ['manifest.json.template', 'color.png', 'outline.png', 'README.md']) {
      assert.ok(
        existsSync(join(stageDir, 'appPackage', rel)),
        `release artifact is missing appPackage/${rel}`,
      );
    }

    // Held invariants from the packaging script: no node_modules, and the
    // staged package.json ships without devDependencies (their file:../ paths
    // describe one machine's directory layout, not the artifact's).
    assert.ok(!existsSync(join(stageDir, 'node_modules')), 'node_modules leaked into the stage');
    const stagedPkg = JSON.parse(readFileSync(join(stageDir, 'package.json'), 'utf8')) as {
      version: string;
      devDependencies?: Record<string, string>;
    };
    assert.equal(stagedPkg.devDependencies, undefined, 'devDependencies not stripped');
    assert.equal(stagedPkg.version, pkgJson.version);
  });
});

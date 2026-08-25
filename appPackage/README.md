# Teams app package template (epic byte5ai/omadia#860, W0)

Versioned template for the Microsoft Teams **app package** (`manifest.json` +
icons) that used to live as an unversioned, hand-maintained artifact in the
middleware deployment. It is the basis for the provisioner's manifest
generator: the middleware's app-package ZIP generator (byte5ai/omadia#865)
renders one package per agent identity by substituting the placeholders below
and zipping the result together with a per-agent `color.png` / `outline.png`.

Rendered with the legacy single-bot values (kept in
`tests/appPackageTemplate.test.ts`, never in this README), the template
reproduces the previously shipped manifest exactly — no breaking change for
existing single-bot deployments. The test suite proves that.

## Placeholders

All placeholders use `{{NAME}}` syntax. Every one substitutes a **JSON string
value** (the generator must not add quotes) — except `{{VALID_DOMAINS}}` and
`{{COMMAND_LISTS}}`, which substitute **raw JSON** (an array of domain strings
and a Teams `commandLists` array respectively; the raw template is therefore
not valid JSON until rendered).

### ⚠️ Escaping rule (mandatory)

String placeholders are interpolated into JSON string literals. The generator
**must JSON-escape every substituted string value** — e.g. substitute
`JSON.stringify(value).slice(1, -1)`, or render into a parsed object instead
of doing string substitution. Without escaping, a routine value like a bot
name containing `"` produces an unparseable `manifest.json` (Teams upload
fails), and a crafted value could inject additional JSON keys. Raw-JSON
placeholders (`{{VALID_DOMAINS}}`, `{{COMMAND_LISTS}}`) must be produced by
serialising a validated array with `JSON.stringify` — never by concatenating
user-supplied text.

| Placeholder | Manifest field(s) | Default / notes |
|---|---|---|
| `{{VERSION}}` | `version` | per-package semver, e.g. `1.0.0` |
| `{{APP_ID}}` | `id` — the **Teams app id** | no default — one GUID per app package |
| `{{BOT_ID}}` | `bots[0].botId` **and** `webApplicationInfo.id` — the bot's **Entra app id** | no default — the `teams_bots[]` entry's `appId` |
| `{{NAME_SHORT}}` | `name.short` (≤ 30 chars) | no default |
| `{{NAME_FULL}}` | `name.full` (≤ 100 chars) | no default |
| `{{DESCRIPTION}}` | `description.short` (≤ 80 chars) | no default |
| `{{DESCRIPTION_FULL}}` | `description.full` (≤ 4000 chars) | no default |
| `{{ACCENT_COLOR}}` | `accentColor` | default `#714B67` |
| `{{DEVELOPER_NAME}}` | `developer.name` | default `byte5` |
| `{{DEVELOPER_WEBSITE_URL}}` | `developer.websiteUrl` | default `https://omadia.ai` |
| `{{DEVELOPER_PRIVACY_URL}}` | `developer.privacyUrl` | default `https://omadia.ai/privacy` |
| `{{DEVELOPER_TERMS_URL}}` | `developer.termsOfUseUrl` | default `https://omadia.ai/terms` |
| `{{COMMAND_LISTS}}` | `bots[0].commandLists` — **raw JSON array** | optional; default `[]` (agent-specific command lists are generator input) |
| `{{TAB_BASE_URL}}` | `staticTabs[0].contentUrl` / `.websiteUrl` (`…/hub`), `configurableTabs[0].configurationUrl` (`…/tab-config`) — the plugin's web-ui surfaces from `src/uiRouter.ts` | deployment-specific, e.g. `https://middleware.example.com/p/channel-teams` |
| `{{VALID_DOMAINS}}` | `validDomains` — **raw JSON array** | deployment-specific, e.g. `["middleware.example.com"]` — every host Teams may iframe |
| `{{MIDDLEWARE_HOST}}` | host part of `webApplicationInfo.resource` (`api://<host>/<botId>`) | deployment-specific, e.g. `middleware.example.com` |

### ⚠️ `{{APP_ID}}` is NOT `{{BOT_ID}}`

The top-level `id` (Teams app id) and `bots[0].botId` (Entra app id of the bot
registration) are **different GUIDs**. Mixing them up makes the Teams upload
fail. `webApplicationInfo.id` must equal the bot's Entra app id, i.e.
`{{BOT_ID}}`, and `webApplicationInfo.resource` must be
`api://{{MIDDLEWARE_HOST}}/{{BOT_ID}}`.

## Security-relevant fields a generator must never drop

- `validDomains` and `webApplicationInfo` bound the origins Teams will iframe
  and the SSO audience. Substitute them; never emit an empty array or omit them.
- `authorization.permissions.resourceSpecific` carries the app's
  resource-specific consent (RSC) grants. They ship verbatim in every rendered
  package — the template is their single source of truth; do not edit, filter,
  or extend them at render time.
- `$schema` and `"manifestVersion": "1.17"` are pinned in the template and are
  not parameterizable.

## Icons

`color.png` (192×192) and `outline.png` (32×32, transparent) are the default
omadia icons, sized to the Teams requirements. A per-agent package replaces
both files with the agent's own avatar assets (same names, same dimensions);
`icons.color` / `icons.outline` in the manifest reference them by filename and
never change.

## Package layout

A rendered app package ZIP contains exactly three files at the archive root:

```
manifest.json   (rendered from manifest.json.template)
color.png
outline.png
```

This directory ships inside the published plugin ZIP (`scripts/build-zip.mjs`
stages `appPackage/` at the archive root), so the middleware-side provisioner
can consume the template from the installed plugin — no repo checkout needed.

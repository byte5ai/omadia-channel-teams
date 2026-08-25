# Teams app package template (epic byte5ai/omadia#860, W0)

Versioned template for the Microsoft Teams **app package** (`manifest.json` +
icons) that used to live as an unversioned, hand-maintained artifact in the
middleware deployment (`middleware/appPackage/manifest.json`, v1.3.0). It is the
basis for the provisioner's manifest generator: the middleware's app-package ZIP
generator (byte5ai/omadia#865) renders one package per agent identity by
substituting the placeholders below and zipping the result together with a
per-agent `color.png` / `outline.png`.

Rendered with the legacy single-bot values (see the table), the template
reproduces today's shipped manifest exactly — no breaking change for existing
single-bot deployments. `tests/appPackageTemplate.test.ts` proves that.

## Placeholders

All placeholders use `{{NAME}}` syntax. Every one substitutes a **JSON string
value** (the generator must not add quotes) — except `{{VALID_DOMAINS}}`, which
substitutes a **raw JSON array** of domain strings (the template deliberately
writes it unquoted, so the raw template is not valid JSON until rendered).

| Placeholder | Manifest field(s) | Legacy single-bot value |
|---|---|---|
| `{{VERSION}}` | `version` | `1.3.0` |
| `{{APP_ID}}` | `id` — the **Teams app id** | `bc0bd6cf-7037-4c7c-9e29-de86c4b77177` |
| `{{BOT_ID}}` | `bots[0].botId` **and** `webApplicationInfo.id` — the bot's **Entra app id** | `737c6ddd-6d4e-4599-8dc3-260281ea906e` |
| `{{NAME_SHORT}}` | `name.short` (≤ 30 chars) | `omadia-agent` |
| `{{NAME_FULL}}` | `name.full` (≤ 100 chars) | `omadia-agent — byte5 Assistent` |
| `{{DESCRIPTION}}` | `description.short` (≤ 80 chars) | see `tests/appPackageTemplate.test.ts` |
| `{{DESCRIPTION_FULL}}` | `description.full` (≤ 4000 chars) | see `tests/appPackageTemplate.test.ts` |
| `{{ACCENT_COLOR}}` | `accentColor` | `#714B67` |
| `{{TAB_BASE_URL}}` | `staticTabs[0].contentUrl` / `.websiteUrl` (`…/hub`), `configurableTabs[0].configurationUrl` (`…/tab-config`) — the plugin's web-ui surfaces from `src/uiRouter.ts` | `https://odoo-bot-harness.fly.dev/p/channel-teams` |
| `{{VALID_DOMAINS}}` | `validDomains` — **raw JSON array** | `["odoo-bot-middleware.fly.dev","odoo-bot-harness.fly.dev"]` |
| `{{MIDDLEWARE_HOST}}` | host part of `webApplicationInfo.resource` (`api://<host>/<botId>`) | `odoo-bot-middleware.fly.dev` |

### ⚠️ `{{APP_ID}}` is NOT `{{BOT_ID}}`

The top-level `id` (Teams app id) and `bots[0].botId` (Entra app id of the bot
registration) are **different GUIDs** in the shipped manifest. Mixing them up
makes the Teams upload fail. `webApplicationInfo.id` must equal the bot's Entra
app id, i.e. `{{BOT_ID}}`, and `webApplicationInfo.resource` must be
`api://{{MIDDLEWARE_HOST}}/{{BOT_ID}}`.

## Security-relevant fields a generator must never drop

- `validDomains` and `webApplicationInfo` bound the origins Teams will iframe
  and the SSO audience. Substitute them; never emit an empty array or omit them.
- `authorization.permissions.resourceSpecific` carries the RSC grants
  (`ChannelMessage.Read.Group`, `ChatMessage.Read.Chat`,
  `TeamsActivity.Send.Group`, `TeamsActivity.Send.Chat`,
  `FileStorageContainer.Selected.Group`, `ChatMember.Read.Chat`,
  `ChannelMember.Read.Group`). They ship verbatim in every rendered package.
- `$schema` and `"manifestVersion": "1.17"` are pinned in the template and are
  not parameterizable.

## Deliberately literal (not placeholders)

- `developer.*` — byte5 / omadia.ai. Identical for every generated package.
- `bots[0].commandLists` — the German default command list. A per-agent
  generator that wants agent-specific commands should render the template,
  `JSON.parse` the result, and replace `bots[0].commandLists` on the parsed
  object rather than growing the placeholder surface.

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

This directory is part of the repository, not of the published plugin ZIP
(`scripts/build-zip.mjs` does not stage it); the middleware-side generator
consumes it from the repo.

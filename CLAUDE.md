# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

This is a **Mintlify documentation site** (based on the Mintlify Starter Kit). It contains no application source code — it is entirely docs content (`.mdx`) plus configuration. There is no build step, test suite, or package.json; pages are authored in MDX and deployed by Mintlify's hosted platform.

## Commands

The Mintlify CLI (`mint`) is the only tool. It is installed globally, not as a project dependency:

```bash
npm i -g mint          # install the CLI (requires Node.js 19+)
mint dev               # preview locally at http://localhost:3000
mint dev --port 3333   # preview on a custom port
mint broken-links      # validate all internal links — run this before pushing
mint update            # update the CLI if the local preview drifts from production
```

There is no lint or test command. `mint broken-links` is the closest thing to a check and should be run after changing navigation, filenames, or cross-page links.

## Deployment

Deployment is automatic via the Mintlify GitHub app: changes merged to the **default branch (`main`)** deploy to production. Do not expect a CI build in-repo.

## Architecture & key conventions

- **`docs.json` is the source of truth for navigation and site config.** Pages do not appear in the site unless they are listed under `navigation.tabs[].groups[].pages` in `docs.json`. Page references there are paths relative to the repo root *without* the `.mdx` extension (e.g. `essentials/settings`, `api-reference/endpoint/get`). When you add, rename, move, or delete a page, you must update `docs.json` to match — this is the most common source of broken builds. `docs.json` also controls theme, colors, logo, navbar, footer, and the `contextual` AI options.

- **Every content page is `.mdx` with YAML frontmatter.** Standard frontmatter keys are `title`, `description`, and optionally `icon`. Pages use Mintlify components (`<Note>`, `<Warning>`, `<Info>`, `<Steps>`/`<Step>`, `<Card>`, `<Frame>`, `<Accordion>`/`<AccordionGroup>`, etc.) — match the components and tone of neighboring pages.

- **Content directories** (each maps to a navigation group in `docs.json`):
  - `essentials/` — how-to/customization guides (settings, navigation, markdown, code, images, snippets).
  - `ai-tools/` — guides for Cursor, Claude Code, Windsurf.
  - `api-reference/` — API docs driven by `api-reference/openapi.json`. Endpoint pages (`api-reference/endpoint/*.mdx`) are nearly empty: they declare an `openapi` frontmatter key (e.g. `openapi: 'GET /plants'`) that references an operation in the OpenAPI spec, and Mintlify renders the playground from that. To change endpoint docs, edit `openapi.json`, not just the MDX.
  - `snippets/` — reusable content. **Any file in `snippets/` is treated as a snippet and is NOT rendered as a standalone page.** Snippets are imported into other pages via `import X from '/snippets/...'` and used as components (also support exported variables and arrow-function components). Use snippets to keep repeated content (DRY) in sync.
  - `images/`, `logo/`, plus `favicon.svg` — static assets referenced by absolute paths (e.g. `/images/hero-dark.png`, `/logo/dark.svg`).

- **Asset references are absolute from the repo root**, e.g. `src="/images/checks-passed.png"`. Light/dark variants are common (`hero-light.png`/`hero-dark.png`, `logo/light.svg`/`logo/dark.svg`).

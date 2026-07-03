# AGENTS.md

Guide for coding agents working on **onedrive-scraper**: a Node (TypeScript)
library + CLI that reads folders and downloads files from OneDrive/SharePoint
using a browser with a persisted session (Playwright).

## What it does and how

It does not use the Graph API. Instead:

1. It opens a **visible browser** for login (including MFA) and **persists the
   session** in `sessionDir` (`.onedrive-session/`).
2. With the session active, it lists and downloads using **two strategies**:
   - **REST** (primary): calls SharePoint's internal REST API
     (`_api/web/GetFolderByServerRelativeUrl(...)/Files|Folders`,
     `GetFileByServerRelativeUrl(...)/$value`) via the browser cookies.
     Fast, reliable, gives exact sizes.
   - **DOM** (fallback): parses the web grid and downloads via clicks. More
     fragile; only used if REST fails (with the `auto` strategy).

## Code layout

| File | Responsibility |
|------|----------------|
| [src/index.ts](src/index.ts) | Public library exports. |
| [src/OneDriveScraper.ts](src/OneDriveScraper.ts) | Main class. FS-like API (`list`, `stat`, `exists`, `readFile`, `downloadFile`) + high-level (`getTree`, `download`). Dispatches REST↔DOM. |
| [src/browser.ts](src/browser.ts) | `BrowserSession`: Playwright persistent context, interactive login, auth detection. |
| [src/rest.ts](src/rest.ts) | `RestClient`: listing and download via the internal REST API. |
| [src/dom.ts](src/dom.ts) | DOM fallback. Selectors in `DOM_SELECTORS`. |
| [src/filters.ts](src/filters.ts) | Globs (`*`,`**`,`?`,`{a,b}`), RegExp, functions; include/exclude by name or path. |
| [src/tree.ts](src/tree.ts) | Text/JSON serialization, `flattenTree`, `formatBytes`, ordering. |
| [src/paths.ts](src/paths.ts) | URL and server-relative path parsing. |
| [src/cli.ts](src/cli.ts) | CLI with commander (`login`, `list`, `tree`, `download`). |

Path convention in the public API: **relative to the configured root folder**,
with `/` separator; the root is `""` or `"."`.

## Commands

```bash
npm install            # deps + Chromium browser (postinstall)
npm run build          # tsc -> dist/
npm run cli -- --help  # CLI via tsx, no build
npm run example        # examples/basic.ts
npm run test:unit      # pure tests (no browser)
npm run test:live      # integration (needs a persisted session)
npm test               # everything
```

After code changes: `npm run build` and `npm test`.

## Tests

- Native `node:test` runner via `tsx` (no extra deps).
- **Unit**: `test/paths|filters|tree|unit.test.ts`. Always runnable.
- **Integration**: `test/live*.test.ts`, `test/cli.test.ts`. They verify the
  known contents of the test folder (`PruebaOneDriveScrapper`: `Doc1.docx`
  9452 B, `Doc2.xlsx` 6624 B, `Folder/Doc3.pptx` 28041 B, `Folder/Doc4.md`
  306 B). If there is no valid session they **skip**, they do not fail.
- Integration tests run **serially** (`--test-concurrency=1`): they share the
  browser profile and Playwright locks the `userDataDir`.

To create/refresh the session: `npm run cli -- login --url "<folderUrl>"`.

## Gotchas (learned the hard way)

- **Session**: `.onedrive-session/` contains credentials. It is in `.gitignore`;
  **never** commit it or delete it without warning. Tests/CLI reuse it.
- **Headless viewport**: in headless mode you must force a wide viewport
  (1600×1000). With `viewport:null` Chromium uses 800×600 and OneDrive switches
  to a compact layout that breaks the DOM selectors. See [src/browser.ts](src/browser.ts).
- **`page.evaluate` and `__name`**: do not use named function expressions inside
  `page.evaluate` (`const foo = () => …`). tsx/esbuild (`keepNames`) injects a
  `__name` helper that does not exist in the browser context → `ReferenceError`.
  Use inline anonymous functions. See `scrollGridToEnd` in [src/dom.ts](src/dom.ts).
- **DOM selectors** (Spartan UI, 2026): row
  `[role="row"][data-automationid^="row-"]` (exclude `row-header`); name in
  `field-LinkFilename`; type via `field-DocIcon img[alt]` (files carry the
  extension in `alt`, folders have no `img`); download with `downloadCommand`.
  If Microsoft changes the web, re-inspect with
  `node scripts/inspect.mjs "<folderUrl>"` and adjust `DOM_SELECTORS`.
- **Interactive login**: requires a visible browser (`headed: true`, the
  default). In `headless` it only works if a valid persisted session exists.
- **ESM + NodeNext**: imports use the `.js` extension in the TS code
  (e.g. `import { x } from "./paths.js"`), even though the file is `.ts`.

## Current scope

Read and download only. Upload and sharing are **not** implemented, but the
architecture (separate REST/DOM layers) allows for them.

## Documentation

- [README.md](README.md) — overview, installation, quick start.
- [docs/API.md](docs/API.md) — API reference with examples.
- [docs/CLI.md](docs/CLI.md) — CLI reference.
- [onedrive-scrapper.md](onedrive-scrapper.md) — original specification.

# onedrive-scraper

Node (TypeScript) library and CLI to **read folders and download files from
OneDrive/SharePoint** using a browser with a persisted session.

It does not use the Graph API and requires no app registration. Instead:

1. It opens a **visible browser** so you can authenticate (including the
   multi-factor step).
2. It **persists the session** to disk, so subsequent runs do not require you to
   authenticate again.
3. Once authenticated, it lists and downloads files using the **SharePoint
   internal REST API** (the same one the OneDrive web app uses), with a
   **fallback to DOM scraping** if REST is unavailable.

Designed as a reusable library for multiple Node projects, with a
filesystem-like API (`list`, `stat`, `readFile`, `downloadFile`) and high-level
operations (`getTree`, `download`).

For now it only reads and downloads; the architecture is ready to add upload and
sharing in the future.

## Installation

From GitHub (project repository):

```bash
npm install github:codeurjc/onedrive-scraper
# or pinning a version/tag:
npm install github:codeurjc/onedrive-scraper#v0.1.0
```

The `postinstall` step automatically installs the Playwright browser (Chromium);
if needed, manually:

```bash
npx playwright install chromium
```

The package exposes the `onedrive-scraper` binary (CLI) and the
`onedrive-scraper` module (library).

## Quick start (library)

```ts
import { OneDriveScraper, treeToText } from "onedrive-scraper";

const scraper = new OneDriveScraper({
  folderUrl:
    "https://urjc-my.sharepoint.com/my?id=%2Fpersonal%2Fmicael%5Fgallego%5Furjc%5Fes%2FDocuments%2FPruebaOneDriveScrapper",
  sessionDir: ".onedrive-session", // where the session is persisted
});

await scraper.connect(); // first time: opens the browser to sign in

// Explore like a filesystem
for (const entry of await scraper.list()) {
  console.log(entry.type, entry.name);
}

// Full structure as text
const tree = await scraper.getTree();
console.log(treeToText(tree, { showSize: true }));

// Download only .docx and .md, recreating the folder structure
await scraper.download("", {
  destDir: "./downloads",
  files: { include: ["*.docx", "*.md"] },
});

await scraper.close();
```

## Quick start (CLI)

```bash
# Authenticate once (persists the session)
onedrive-scraper login --url "https://urjc-my.sharepoint.com/my?id=..."

# List the root folder
onedrive-scraper list --url "https://urjc-my.sharepoint.com/my?id=..."

# Export the structure to text and JSON
onedrive-scraper tree --url "..." --out-txt structure.txt --out-json structure.json

# Download only .docx and .md, recreating folders
onedrive-scraper download --url "..." --dest ./downloads --include "*.docx" "*.md"
```

## Specifying the target folder

You can use any of these forms:

- **`folderUrl`** (recommended): the URL you copy from the browser. Its `id`
  query parameter holds the *server-relative* path, from which the site and root
  folder are derived.
- **`siteUrl` + `rootPath`**: the site base (e.g.
  `https://org-my.sharepoint.com/personal/user_org_es`) and the server-relative
  path of the root folder.

All paths in the methods are **relative to that root folder** (the root is `""`).

## Session persistence

The browser profile (cookies, tokens) is stored in `sessionDir`
(`./.onedrive-session` by default). **Do not commit it to version control**
(it is already in `.gitignore`): it contains session credentials.

## Strategies

| Strategy | Description |
|----------|-------------|
| `auto` (default) | Internal REST, falling back to DOM if it fails. |
| `rest` | REST only. Faster and more robust. |
| `dom` | DOM scraping only. Useful if REST is blocked; more fragile. |

## Documentation

- [docs/API.md](docs/API.md) — full API reference with examples.
- [docs/CLI.md](docs/CLI.md) — full CLI command reference.

## Skill for coding agents (Claude Code)

The repository includes a **skill** that teaches a coding agent (Claude Code)
how to use this CLI/library to download files from OneDrive. Install it in your
Node project and the agent will know when and how to use the tool.

Skills are `SKILL.md` files inside `.claude/skills/<name>/`. You can install it
with **project scope** (that repo only) or **user scope** (all your projects).

**Project scope** — inside your project:

```bash
mkdir -p .claude/skills/onedrive-download
curl -fsSL \
  https://raw.githubusercontent.com/codeurjc/onedrive-scraper/main/.claude/skills/onedrive-download/SKILL.md \
  -o .claude/skills/onedrive-download/SKILL.md
```

**User scope** — available in all your projects:

```bash
mkdir -p ~/.claude/skills/onedrive-download
curl -fsSL \
  https://raw.githubusercontent.com/codeurjc/onedrive-scraper/main/.claude/skills/onedrive-download/SKILL.md \
  -o ~/.claude/skills/onedrive-download/SKILL.md
```

Alternatives: copy the file after cloning the repo, or (if you already installed
the package) from `node_modules`:

```bash
git clone https://github.com/codeurjc/onedrive-scraper
cp -r onedrive-scraper/.claude/skills/onedrive-download .claude/skills/

# or, with the package already installed as a dependency:
mkdir -p .claude/skills/onedrive-download
cp node_modules/onedrive-scraper/.claude/skills/onedrive-download/SKILL.md \
   .claude/skills/onedrive-download/
```

Once installed, in an interactive Claude Code session you can invoke it with
`/onedrive-download`, and the agent will use it on its own when the task fits
(e.g. "download the files from this OneDrive folder"). Then follow the recipes in
[docs/CLI.md](docs/CLI.md) / [docs/API.md](docs/API.md).

> Note: for the `node_modules` variant to work, the package must include
> `.claude` in the published `files` (it does — see [package.json](package.json)).

## Development

```bash
npm install
npm run build          # compiles to dist/
npm run cli -- --help  # runs the CLI with tsx (no build)
npm run example        # runs examples/basic.ts
```

## Tests

Uses Node's built-in runner (`node:test`) via `tsx`, with no extra dependencies.

```bash
npm run test:unit   # pure logic (paths, filters, tree) — no browser needed
npm run test:live   # integration against the test folder (uses the persisted session)
npm test            # everything
```

**Unit** (no browser needed):

- [test/paths.test.ts](test/paths.test.ts) — URL and server-relative path parsing (personal/sites/teams).
- [test/filters.test.ts](test/filters.test.ts) — globs (`*`, `**`, `?`, `{a,b}`), RegExp, functions, include/exclude by name or path.
- [test/tree.test.ts](test/tree.test.ts) — ordering (folders first), text, JSON, `flattenTree`, `formatBytes`.
- [test/unit.test.ts](test/unit.test.ts) — base cases with the real folder URL.

**Integration** (use the persisted session; verify the known contents of
`PruebaOneDriveScrapper`: `Doc1.docx`, `Doc2.xlsx`, `Folder/Doc3.pptx`,
`Folder/Doc4.md`):

- [test/live.test.ts](test/live.test.ts) — listing, tree, filter and download with **both strategies** (REST and DOM); exact sizes.
- [test/live-advanced.test.ts](test/live-advanced.test.ts) — `stat`/`exists`, **folder** filters (include/exclude), `maxDepth`, **path** filter, `downloadFile`, and overwrite/skip.
- [test/cli.test.ts](test/cli.test.ts) — runs the CLI (`list --json`, `tree --out-json`, filters, `download --include`) as a child process.

If there is no valid persisted session, the integration tests are **skipped**
instead of failing; create the session once with
`npm run cli -- login --url "..."`. Integration tests run serially
(`--test-concurrency=1`) because they share the browser profile.

The DOM selectors ([src/dom.ts](src/dom.ts)) were tuned against the real
OneDrive UI. If Microsoft changes the markup, you can re-inspect it with:

```bash
node scripts/inspect.mjs "https://<tenant>-my.sharepoint.com/my?id=..."
```

which opens the browser, dumps the REST listing and the per-row DOM details
(name cell text, icon `alt`, size cell and the download button) so you can
re-tune `DOM_SELECTORS`.

## Notes and limitations

- **Interactive login** requires a visible browser (`headed: true`, the
  default). In `headless` mode it only works if a valid persisted session
  already exists.
- The **DOM fallback** depends on OneDrive's (changing, localized) web markup.
  The selectors are grouped in `DOM_SELECTORS` ([src/dom.ts](src/dom.ts)) so they
  can be re-tuned for a specific tenant without touching the logic.
- Upload and sharing are **not implemented yet** (the architecture allows for
  them).

## License

MIT

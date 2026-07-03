---
name: onedrive-download
description: >-
  Download files (or list/export the structure) from a OneDrive/SharePoint
  folder using this project's CLI or library (onedrive-scraper). Use whenever
  the task involves fetching, downloading, listing, or mirroring files from a
  OneDrive or SharePoint folder/URL (e.g. "download the files from this OneDrive
  folder", "list a SharePoint folder", "mirror a *-my.sharepoint.com URL",
  "export the folder tree to json"). Covers authentication (persisted browser
  session), filters, and structure export.
---

# Download files from OneDrive with onedrive-scraper

This library reads and downloads from OneDrive/SharePoint using a browser with a
persisted session (Playwright). It does not use the Graph API. `auto` strategy:
internal REST (reliable) with a DOM fallback.

## When to use CLI vs library

- **CLI**: one-off actions or shell scripts (download, list, export structure).
  Fastest to use.
- **Library**: when integrating the download into Node/TS code and you need the
  entries as data (`OneDriveEntry[]`, tree, programmatic filters).

## Step 0 — Setup (once)

```bash
npm install        # deps + Chromium (postinstall)
# The first operation opens a browser to sign in (MFA).
# The session is saved to ./.onedrive-session and reused afterwards.
```

If you need to force login ahead of time:

```bash
npm run cli -- login --url "<folderUrl>"
```

**Important note for agents**: login is **interactive** and needs a visible
browser. In an environment **without a display / headless with no prior
session** you cannot authenticate: ask the user to run `login` once. If
`.onedrive-session/` already exists and is valid, everything works headless.

## Getting the `folderUrl`

It is the URL shown in the browser when you open the folder, with the `id`
parameter. Example:

```
https://<tenant>-my.sharepoint.com/my?id=%2Fpersonal%2F<user>%2FDocuments%2F<Folder>
```

Pass it verbatim in quotes. Alternatively `--site <url> --root <path>`.

## CLI recipes

In development use `npm run cli --`; once installed, use `onedrive-scraper`.

```bash
# List the root folder
npm run cli -- list --url "<folderUrl>"

# Export the structure to text and JSON (without downloading content)
npm run cli -- tree --url "<folderUrl>" --out-txt structure.txt --out-json structure.json

# Download EVERYTHING recreating the folder structure
npm run cli -- download --url "<folderUrl>" --dest ./downloads

# Download only certain types (globs)
npm run cli -- download --url "<folderUrl>" --dest ./downloads --include "*.docx" "*.pdf"

# Download flat (no folder recreation) and overwriting
npm run cli -- download --url "<folderUrl>" --dest ./out --no-reconstruct --overwrite

# Filter by full PATH (not just name) and skip folders
npm run cli -- download --url "<folderUrl>" --include "Reports/**" --match-path \
  --exclude-folder "Temp*" "~*"

# Limit depth
npm run cli -- download --url "<folderUrl>" --depth 2
```

Key options: `--include/--exclude` (files), `--include-folder/--exclude-folder`
(folders), `--match-path` (match against the full relative path),
`--no-reconstruct`, `--overwrite`, `--depth`, `--strategy auto|rest|dom`,
`--session <dir>`. Full reference: [docs/CLI.md](../../../docs/CLI.md).

`download` output: to stderr, `↓` downloaded · `=` skipped · `✗` error; exit
code `1` if any error occurred.

## Library recipes

```ts
import { OneDriveScraper, treeToText, treeToJson } from "onedrive-scraper";
// In this repo, during development: import ... from "../src/index.js" (with tsx)

const scraper = new OneDriveScraper({
  folderUrl: "<folderUrl>",
  sessionDir: ".onedrive-session",
});
await scraper.connect();               // interactive login the 1st time
try {
  // Explore like a filesystem (paths relative to the root, "" = root)
  const entries = await scraper.list();          // readdir
  const info = await scraper.stat("Folder/Doc3.pptx");
  const exists = await scraper.exists("Doc1.docx");

  // Full structure (with filters) -> text / JSON
  const tree = await scraper.getTree("", {
    files:   { include: ["*.docx", "*.md"] },
    folders: { exclude: "Temp*" },
    maxDepth: Infinity,
  });
  console.log(treeToText(tree, { showSize: true }));

  // Download with filters, recreating the structure
  await scraper.download("", {
    destDir: "./downloads",
    reconstruct: true,          // false = everything flat in destDir
    overwrite: false,           // true = overwrite existing
    files:   { include: "*.pdf" },
    folders: { exclude: "**/Archive" },       // use target 'path' with rules
    onProgress: (p) => console.log(p.index + 1, "/", p.total, p.status, p.entry.path),
  });

  // A single file
  await scraper.downloadFile("Doc1.docx", "./downloads/Doc1.docx");
  const buf = await scraper.readFile("Folder/Doc4.md"); // Buffer in memory
} finally {
  await scraper.close();        // always! flushes the session to disk
}
```

Filters: a pattern = glob (`*`,`**`,`?`,`{a,b}`), `RegExp`, or a function
`(value, entry) => boolean`. By default it matches the **name**; to match the
**path** use `{ pattern, target: "path" }`. Reference: [docs/API.md](../../../docs/API.md).

## Verify

After downloading, confirm the result (do not assume success):

```bash
ls -R ./downloads          # expected files present
```

With the library, check the array returned by `download()`:
`status === "error"` on any item indicates a failure (see `.error`).

## Common errors

- *"Interactive login did not complete"* / a browser opens and nobody
  authenticates → there is no valid session; ask the user to run `login` on a
  machine with a display.
- *"Not authenticated and running headless"* → you passed `--headless`/`headed:false`
  with no prior session. Run the login with a visible browser first.
- REST fails on a restricted tenant → use `--strategy dom` (slower) or leave
  `auto` (it falls back on its own).
- Do not mix credentials: `.onedrive-session/` holds the session; do not commit
  it to git or delete it.

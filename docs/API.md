# API Reference

All paths passed to the methods are **relative to the configured root folder**
and use `/` as the separator. The root is `""` (or `"."`).

- [`new OneDriveScraper(options)`](#new-onedrivescraperoptions)
- [Lifecycle](#lifecycle): `connect`, `close`, `page`
- [Filesystem-like API](#filesystem-like-api): `list`, `stat`, `exists`, `readFile`, `downloadFile`
- [High-level operations](#high-level-operations): `getTree`, `download`
- [Filters](#filters)
- [Tree serialization](#tree-serialization): `treeToText`, `treeToJson`, `flattenTree`
- [Utilities](#utilities)
- [Types](#types)

---

## `new OneDriveScraper(options)`

Creates an instance. It does not open the browser yet (that is done by
`connect()`).

```ts
import { OneDriveScraper } from "onedrive-scraper";

const scraper = new OneDriveScraper({
  folderUrl:
    "https://urjc-my.sharepoint.com/my?id=%2Fpersonal%2Fmicael%5Fgallego%5Furjc%5Fes%2FDocuments%2FPruebaOneDriveScrapper",
  sessionDir: ".onedrive-session",
});
```

### `OneDriveScraperOptions`

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `folderUrl` | `string` | — | Folder URL. Its `id` query parameter is parsed to derive the site and root. |
| `siteUrl` | `string` | — | Site base, e.g. `https://org-my.sharepoint.com/personal/user_org_es`. |
| `rootPath` | `string` | — | Server-relative path of the root folder. |
| `sessionDir` | `string` | `".onedrive-session"` | Directory where the browser profile is persisted. |
| `headed` | `boolean` | `true` | Show the browser window (required for MFA login). |
| `strategy` | `"auto" \| "rest" \| "dom"` | `"auto"` | Scraping strategy. |
| `loginTimeout` | `number` | `300000` | Max time (ms) to wait for interactive login. |
| `locale` | `string` | `"en-US"` | Browser locale. |
| `timeout` | `number` | `30000` | Per-operation/navigation timeout (ms). |

You must provide **`folderUrl`** or **`siteUrl` (+ `rootPath`)**.

---

## Lifecycle

### `connect(): Promise<void>`

Launches the browser and ensures the session is authenticated. The first time it
opens a window for you to sign in (including MFA); subsequent runs reuse the
persisted session. **Must be called before any operation.**

### `close(): Promise<void>`

Closes the browser and flushes the persisted profile to disk. Always call it
when done (ideally in a `finally`).

### `page`

Getter returning the active Playwright `Page`, in case you need custom scraping.
Throws if `connect()` has not been called yet.

---

## Filesystem-like API

### `list(relativePath = ""): Promise<OneDriveEntry[]>`

Lists the immediate children of a folder (like `readdir`). Folders come first,
then alphabetically.

```ts
const entries = await scraper.list();          // root
const sub = await scraper.list("Folder");      // subfolder

for (const e of entries) {
  console.log(e.type, e.name, e.size ?? "");
}
```

### `stat(relativePath): Promise<OneDriveEntry>`

Metadata for a single entry (like `stat`). Throws if it does not exist.

```ts
const info = await scraper.stat("Folder/Doc3.pptx");
console.log(info.size, info.modified);
```

### `exists(relativePath): Promise<boolean>`

```ts
if (await scraper.exists("Folder/Doc4.md")) { /* ... */ }
```

### `readFile(relativePath): Promise<Buffer>`

Downloads a file's content into memory.

```ts
const buf = await scraper.readFile("Doc1.docx");
console.log(buf.length, "bytes");
```

### `downloadFile(relativePath, localPath): Promise<void>`

Downloads a file to a local path (creating the necessary folders).

```ts
await scraper.downloadFile("Folder/Doc3.pptx", "./out/Doc3.pptx");
```

---

## High-level operations

### `getTree(relativePath = "", options?): Promise<OneDriveTreeNode>`

Recursively builds a folder tree, applying filters. A folder rejected by the
folder filter is **neither listed nor descended into**.

```ts
const tree = await scraper.getTree("", {
  files:   { include: ["*.docx", "*.md"] },
  folders: { exclude: "Archive*" },
  maxDepth: Infinity,
});
```

`TreeOptions`: `files?`, `folders?` (see [Filters](#filters)) and `maxDepth?`
(number, default `Infinity`).

Combine it with [`treeToText`/`treeToJson`](#tree-serialization) to export the
structure:

```ts
import { treeToText, treeToJson } from "onedrive-scraper";
import { writeFile } from "node:fs/promises";

const tree = await scraper.getTree();
await writeFile("structure.txt", treeToText(tree, { showSize: true }));
await writeFile("structure.json", JSON.stringify(treeToJson(tree), null, 2));
```

### `download(relativePath = "", options?): Promise<DownloadProgress[]>`

Downloads every file that passes the filters. Recreates the remote folder
structure on disk unless `reconstruct` is `false` (in which case everything is
downloaded flat into `destDir`).

```ts
const results = await scraper.download("", {
  destDir: "./downloads",
  reconstruct: true,       // recreate folders (default)
  overwrite: false,        // skip files that already exist (default)
  files:   { include: "*.pdf" },
  folders: { exclude: "**/Temp" },
  onProgress: (p) => console.log(`${p.index + 1}/${p.total}`, p.status, p.entry.path),
});
```

`DownloadOptions`: `destDir?` (default the cwd), `reconstruct?` (`true`),
`overwrite?` (`false`), `maxDepth?`, `onProgress?`, plus the `files?` /
`folders?` filters.

Returns an array of `DownloadProgress` with `status: "downloaded" | "skipped" |
"error"` per file.

---

## Filters

Filters are applied to **files** (`files`) and **folders** (`folders`)
separately. An item passes if it matches at least one `include` pattern (or there
are none) and matches no `exclude` pattern.

A pattern can be:

- a **glob** (`string`): `*.docx`, `Doc?.pptx`, `**/Report*`, `*.{docx,md}`.
- a **`RegExp`**: `/\.docx$/i`.
- a **function** `(value, entry) => boolean`.

By default patterns are matched against the item's **name**. To match against the
full **path** relative to the root, use the rule form with `target: "path"`:

```ts
await scraper.getTree("", {
  files: {
    include: [
      "*.docx",                                   // by name
      { pattern: "Folder/**", target: "path" },   // by full path
      /\.md$/i,                                    // RegExp
      (name) => name.startsWith("Report"),        // function
    ],
    exclude: "~$*",
  },
});
```

You can also compile a reusable filter:

```ts
import { compileFilter } from "onedrive-scraper";

const f = compileFilter({ include: ["*.docx", "*.md"], exclude: "draft*" });
f.test({ name: "Doc1.docx", path: "Doc1.docx", type: "file", serverRelativeUrl: "" });
```

---

## Tree serialization

### `treeToText(node, options?): string`

Renders the tree as an indented ASCII listing. `options.showSize` appends each
file's size.

```
PruebaOneDriveScrapper/
├── Folder/
│   ├── Doc3.pptx (4.9 KB)
│   └── Doc4.md (100 B)
├── Doc1.docx (12.1 KB)
└── Doc2.xlsx (2 KB)
```

### `treeToJson(node): SerializedTreeNode`

Converts the tree to a `JSON.stringify`-serializable object (dates as ISO).

### `flattenTree(node): OneDriveEntry[]`

Flattens the tree into a list of entries (folders first, then files).

---

## Utilities

```ts
import {
  resolveLocation,          // options -> { origin, siteUrl, rootServerRelativeUrl, ... }
  serverRelativeFromUrl,    // extracts the server-relative `id` from a URL
  normalizeServerRelative,  // normalizes a server-relative path
  globToRegExp,             // glob -> RegExp
  formatBytes,              // 12345 -> "12.1 KB"
  sortEntries,              // comparator (folders first, then alphabetical)
} from "onedrive-scraper";
```

---

## Types

### `OneDriveEntry`

```ts
interface OneDriveEntry {
  name: string;              // "Doc1.docx"
  path: string;              // relative to the root: "Folder/Doc3.pptx" (root = "")
  serverRelativeUrl: string; // "/personal/.../Documents/Root/Folder/Doc3.pptx"
  type: "file" | "folder";
  size?: number;             // bytes (files)
  modified?: Date;
  childCount?: number;       // number of children (folders)
}
```

### `OneDriveTreeNode`

`OneDriveEntry` plus `children?: OneDriveTreeNode[]` (present on folders).

### `DownloadProgress`

```ts
interface DownloadProgress {
  entry: OneDriveEntry;
  localPath: string;
  status: "downloaded" | "skipped" | "error";
  index: number;
  total: number;
  error?: Error;
  bytes?: number;
}
```

See [src/types.ts](../src/types.ts) for the rest of the exported types.

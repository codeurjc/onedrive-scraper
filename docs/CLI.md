# CLI Reference

The CLI exposes the library's high-level operations from the command line.

```bash
onedrive-scraper <command> [options] [path]
```

During development, without building:

```bash
npm run cli -- <command> [options]
```

The `path` (optional) is **relative to the configured root folder**. If omitted,
the root is used.

> Progress/status messages are printed to **stderr**; data output (listings,
> trees) to **stdout**, so it can be redirected with `>`.

## Common options (all commands)

| Option | Description | Default |
|--------|-------------|---------|
| `-u, --url <url>` | OneDrive/SharePoint folder URL (its `id` is parsed). | — |
| `-s, --site <url>` | Site base, e.g. `https://org-my.sharepoint.com/personal/user_org_es`. | — |
| `-r, --root <path>` | Server-relative path of the root folder. | — |
| `--session <dir>` | Persisted session directory. | `.onedrive-session` |
| `--strategy <mode>` | `auto`, `rest` or `dom`. | `auto` |
| `--headless` | Run the browser without a window (only with an existing session). | (no) |
| `--locale <locale>` | Browser locale. | `en-US` |
| `--timeout <ms>` | Per-operation timeout. | `30000` |

You must provide **`--url`** or **`--site` (+ `--root`)**.

## Filter options (`tree` and `download` commands)

| Option | Description |
|--------|-------------|
| `--include <glob...>` | Include only files matching these glob(s). |
| `--exclude <glob...>` | Exclude files matching these glob(s). |
| `--include-folder <glob...>` | Explore only folders matching these glob(s). |
| `--exclude-folder <glob...>` | Skip folders matching these glob(s). |
| `--match-path` | Match filters against the **full path** instead of the name. |
| `--depth <n>` | Maximum recursion depth. |

Globs support `*`, `**`, `?` and `{a,b}`. With multiple values, repeat the option
or list them space-separated: `--include "*.docx" "*.md"`.

---

## `login`

Opens the browser to authenticate and persist the session. Useful to sign in
once before automating the rest.

```bash
onedrive-scraper login --url "https://urjc-my.sharepoint.com/my?id=%2Fpersonal%2F..."
```

---

## `list [path]`

Lists the immediate children of a folder.

| Option | Description |
|--------|-------------|
| `--json` | Output JSON instead of a table. |

```bash
# Table for the root
onedrive-scraper list --url "..."

# JSON for a subfolder
onedrive-scraper list Folder --url "..." --json
```

Table output (`d` = folder, `-` = file):

```
d  Folder
-  Doc1.docx                                12.1 KB
-  Doc2.xlsx                                2 KB
```

---

## `tree [path]`

Prints or exports the folder/file structure as text and/or JSON.

| Option | Description |
|--------|-------------|
| `--out-txt <file>` | Write the tree as text to a file. |
| `--out-json <file>` | Write the tree as JSON to a file. |
| `--show-size` | Show file sizes in the text output. |

If neither `--out-txt` nor `--out-json` is given, it prints the text tree to
stdout.

```bash
# Print the tree with sizes
onedrive-scraper tree --url "..." --show-size

# Export to text and JSON at once
onedrive-scraper tree --url "..." --out-txt structure.txt --out-json structure.json

# Only the structure of .docx and .md, up to 2 levels
onedrive-scraper tree --url "..." --include "*.docx" "*.md" --depth 2
```

---

## `download [path]`

Downloads the files that pass the filters.

| Option | Description | Default |
|--------|-------------|---------|
| `-d, --dest <dir>` | Local destination directory. | `downloads` |
| `--no-reconstruct` | Download everything flat, without recreating the folder structure. | (reconstructs) |
| `--overwrite` | Overwrite files that already exist locally. | (skips) |

```bash
# Download everything recreating the structure
onedrive-scraper download --url "..." --dest ./downloads

# Only .docx and .md
onedrive-scraper download --url "..." --dest ./downloads --include "*.docx" "*.md"

# Everything under "Folder/" (path filter), flat and overwriting
onedrive-scraper download --url "..." --include "Folder/**" --match-path --no-reconstruct --overwrite

# Excluding temporary folders
onedrive-scraper download --url "..." --exclude-folder "Temp*" "~*"
```

Progress (to stderr):

```
[1/4] ↓ Doc1.docx (12.1 KB)
[2/4] ↓ Doc2.xlsx (2 KB)
[3/4] = Folder/Doc3.pptx
[4/4] ↓ Folder/Doc4.md (100 B)
Done: 3 downloaded, 1 skipped, 0 failed → ./downloads
```

`↓` downloaded · `=` skipped (already existed) · `✗` error. If any error occurs,
the command exits with status code `1`.

---

## End-to-end example

```bash
# 1. Authenticate (once)
onedrive-scraper login \
  --url "https://urjc-my.sharepoint.com/my?id=%2Fpersonal%2Fmicael%5Fgallego%5Furjc%5Fes%2FDocuments%2FPruebaOneDriveScrapper"

# 2. Inspect and inventory the structure
onedrive-scraper tree --url "https://urjc-my.sharepoint.com/my?id=..." \
  --out-txt inventory.txt --out-json inventory.json

# 3. Download only the office documents
onedrive-scraper download --url "https://urjc-my.sharepoint.com/my?id=..." \
  --dest ./downloads --include "*.docx" "*.xlsx" "*.pptx"
```

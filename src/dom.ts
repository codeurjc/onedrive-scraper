/**
 * DOM fallback strategy.
 *
 * Used only when the REST strategy is unavailable (e.g. a tenant that blocks
 * the internal REST endpoints). It drives the OneDrive web UI directly: reads
 * the file list from the rendered grid and triggers real browser downloads.
 *
 * The selectors below were tuned against the current ("Spartan") OneDrive web
 * UI. They are grouped in {@link DOM_SELECTORS} so they can be re-tuned for a
 * different tenant / future UI without touching the logic.
 *
 * Observed markup (urjc-my.sharepoint.com, 2026):
 * - Data rows:   `div[role="row"][data-automationid="row-<n>"]` (header is `row-header`).
 * - Name cell:   `[data-automationid="field-LinkFilename"]` (its text is the item name).
 * - Icon cell:   `[data-automationid="field-DocIcon"] img` — `alt` is the file
 *                extension (".docx") for files, and absent for folders.
 * - Size cell:   `[data-automationid="field-FileSizeDisplay"]` — a byte size for
 *                files, an item count ("2 elementos") for folders.
 * - Selection:   `input[data-automationid="selection-checkbox"]` inside the row.
 * - Download:    toolbar `[data-automationid="downloadCommand"]`.
 */

import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { Page } from "playwright";
import type { OneDriveEntry } from "./types.js";
import {
  basename,
  normalizeServerRelative,
  relativeFromServerRelative,
  resolveServerRelative,
} from "./paths.js";

/** Selectors used to read the OneDrive web UI. Tune per tenant if the UI changes. */
export const DOM_SELECTORS = {
  /** Container that appears once the file grid has rendered. */
  grid: '[role="grid"]',
  /** A single data row (excludes the column header row). */
  row: '[role="row"][data-automationid^="row-"]:not([data-automationid="row-header"])',
  /** Cell holding the item name. */
  nameCell: '[data-automationid="field-LinkFilename"]',
  /** Icon image inside a row; its `alt` is the file extension for files. */
  iconImg: '[data-automationid="field-DocIcon"] img',
  /** Per-row selection checkbox. */
  selectionCheckbox: 'input[data-automationid="selection-checkbox"]',
  /** Toolbar "Download" command. */
  downloadButton:
    '[data-automationid="downloadCommand"], button[name="Download"], button[aria-label*="descarg" i]',
};

/**
 * List the immediate children of a folder by reading the rendered grid.
 * Type is detected from the icon's `alt` extension (files) vs its absence
 * (folders). Sizes are not read from the DOM (localized/unreliable); prefer the
 * REST strategy when accurate sizes are needed.
 */
export async function listChildrenViaDom(
  page: Page,
  siteUrl: string,
  folderServerRelative: string,
  rootServerRelative: string
): Promise<OneDriveEntry[]> {
  await navigateToFolder(page, siteUrl, folderServerRelative);
  await waitForGrid(page);
  await scrollGridToEnd(page);

  const raw = await page.evaluate((sel: typeof DOM_SELECTORS) => {
    const rows = Array.from(document.querySelectorAll(sel.row));
    const seen = new Set<string>();
    const out: { name: string; isFolder: boolean }[] = [];
    for (const row of rows) {
      const nameCell = row.querySelector(sel.nameCell);
      const name = (nameCell?.textContent || "").trim();
      if (!name || seen.has(name)) continue;
      seen.add(name);
      const img = row.querySelector(sel.iconImg);
      const alt = (img?.getAttribute("alt") || "").trim();
      // Files expose the extension as the icon alt (".docx"); folders do not.
      const isFile = /^\.[a-z0-9]+$/i.test(alt);
      out.push({ name, isFolder: !isFile });
    }
    return out;
  }, DOM_SELECTORS);

  return raw.map((r) => {
    const serverRelativeUrl = resolveServerRelative(folderServerRelative, r.name);
    return {
      name: r.name,
      path: relativeFromServerRelative(rootServerRelative, serverRelativeUrl),
      serverRelativeUrl,
      type: r.isFolder ? "folder" : "file",
    } satisfies OneDriveEntry;
  });
}

/**
 * Download a file by navigating to its parent folder, selecting the row, and
 * triggering the toolbar download. Returns the file bytes.
 */
export async function downloadFileViaDom(
  page: Page,
  siteUrl: string,
  fileServerRelative: string
): Promise<Buffer> {
  const normalized = normalizeServerRelative(fileServerRelative);
  const parent = normalized.split("/").slice(0, -1).join("/");
  const name = basename(normalized);

  await navigateToFolder(page, siteUrl, parent);
  await waitForGrid(page);
  await scrollGridToEnd(page);

  // Locate the row whose name cell holds this file name and select it.
  const nameCell = page.locator(DOM_SELECTORS.nameCell, { hasText: name });
  const row = page.locator(DOM_SELECTORS.row).filter({ has: nameCell }).first();
  await row.scrollIntoViewIfNeeded();
  await row.hover();
  await row.locator(DOM_SELECTORS.selectionCheckbox).click({ force: true });

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator(DOM_SELECTORS.downloadButton).first().click(),
  ]);

  const dir = await mkdtemp(join(tmpdir(), "onedrive-dom-"));
  const tmpPath = join(dir, download.suggestedFilename() || name);
  try {
    await download.saveAs(tmpPath);
    return await readFile(tmpPath);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

/** Server-relative path currently shown in the grid, from the URL `id` param. */
function currentFolderId(page: Page): string | null {
  try {
    const id = new URL(page.url()).searchParams.get("id");
    return id ? normalizeServerRelative(id) : null;
  } catch {
    return null;
  }
}

async function navigateToFolder(
  page: Page,
  siteUrl: string,
  folderServerRelative: string
): Promise<void> {
  const target = normalizeServerRelative(folderServerRelative);
  if (currentFolderId(page) === target) return;
  const url = `${siteUrl}/_layouts/15/onedrive.aspx?id=${encodeURIComponent(target)}`;
  await page.goto(url, { waitUntil: "domcontentloaded" });
}

async function waitForGrid(page: Page): Promise<void> {
  await page.waitForSelector(DOM_SELECTORS.grid, { state: "visible" });
  await page.waitForSelector(DOM_SELECTORS.row, { state: "visible" }).catch(() => {
    // Empty folder: no rows ever appear. That's fine.
  });
}

/** Scroll the virtualized grid to force all rows to render. */
async function scrollGridToEnd(page: Page): Promise<void> {
  await page.evaluate(async (gridSel: string) => {
    const grid =
      (document.querySelector(gridSel) as HTMLElement | null) ||
      document.scrollingElement;
    if (!grid) return;
    let previous = -1;
    for (let i = 0; i < 50; i++) {
      grid.scrollTop = grid.scrollHeight;
      // Inline (anonymous) to avoid transpilers injecting a __name helper that
      // does not exist inside the browser evaluation context.
      await new Promise((resolve) => setTimeout(resolve, 250));
      if (grid.scrollHeight === previous) break;
      previous = grid.scrollHeight;
    }
    grid.scrollTop = 0;
  }, DOM_SELECTORS.grid);
}

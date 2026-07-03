/**
 * Maintenance utility: inspect the OneDrive DOM to (re-)tune the selectors in
 * src/dom.ts if Microsoft changes the web UI.
 *
 * Usage:
 *   node scripts/inspect.mjs "<folderUrl>"
 *
 * - Opens a headed browser (persisted session) so you can sign in if needed.
 * - Prints, for the target folder, the REST listing and the precise per-row
 *   DOM details used by the DOM fallback (name cell text, icon alt, size cell).
 *
 * The folder URL is the one you copy from the browser address bar, e.g.
 *   https://<tenant>-my.sharepoint.com/my?id=%2Fpersonal%2F...%2FYourFolder
 */
import { chromium } from "playwright";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const HERE = dirname(fileURLToPath(import.meta.url));
const SESSION_DIR = join(HERE, "..", ".onedrive-session");
const log = (...a) => console.error("[inspect]", ...a);

const folderUrl = process.argv[2] || process.env.ONEDRIVE_FOLDER_URL;
if (!folderUrl) {
  console.error('Usage: node scripts/inspect.mjs "<folderUrl>"');
  process.exit(2);
}

const u = new URL(folderUrl);
const origin = u.origin;
const id = decodeURIComponent(u.searchParams.get("id") || "");
if (!id) {
  console.error("The URL has no `id` parameter (server-relative folder path).");
  process.exit(2);
}
const segs = id.split("/").filter(Boolean);
const site = `${origin}/${segs.slice(0, 2).join("/")}`;
const root = "/" + segs.join("/");

function odata(p) {
  return `'${encodeURIComponent(p.replace(/'/g, "''"))}'`;
}

async function restList(request, serverRelative) {
  const base = `${site}/_api/web/GetFolderByServerRelativeUrl(${odata(serverRelative)})`;
  const get = async (u) => {
    const r = await request.get(u, { headers: { Accept: "application/json;odata=nometadata" } });
    return r.ok() ? await r.json() : { error: `${r.status()} ${r.statusText()}` };
  };
  return {
    folders: await get(`${base}/Folders?$select=Name,ServerRelativeUrl,ItemCount`),
    files: await get(`${base}/Files?$select=Name,ServerRelativeUrl,Length`),
  };
}

const ctx = await chromium.launchPersistentContext(SESSION_DIR, {
  headless: false,
  locale: "es-ES",
  viewport: null,
  args: ["--start-maximized"],
});
const page = ctx.pages()[0] ?? (await ctx.newPage());
ctx.setDefaultTimeout(60000);

log("Site:", site);
log("Root:", root);
log("Navigating — sign in if prompted…");
await page.goto(`${site}/_layouts/15/onedrive.aspx?id=${encodeURIComponent(root)}`, {
  waitUntil: "domcontentloaded",
});

const deadline = Date.now() + 420000;
while (Date.now() < deadline) {
  const r = await ctx.request.get(`${site}/_api/web/currentuser?$select=LoginName`, {
    headers: { Accept: "application/json;odata=nometadata" },
  });
  if (r.ok()) { log("Authenticated:", (await r.json()).LoginName); break; }
  await page.waitForTimeout(3000);
}

log("REST listing:", JSON.stringify(await restList(ctx.request, root), null, 2));

await page.waitForSelector('[role="row"]', { timeout: 30000 }).catch(() => log("No rows visible."));
await page.waitForTimeout(2000);

const rows = await page.evaluate(() => {
  const rows = Array.from(
    document.querySelectorAll('[role="row"][data-automationid^="row-"]:not([data-automationid="row-header"])')
  );
  return rows.map((row) => {
    const nameCell = row.querySelector('[data-automationid="field-LinkFilename"]');
    const img = row.querySelector('[data-automationid="field-DocIcon"] img');
    const sizeCell = row.querySelector('[data-automationid="field-FileSizeDisplay"]');
    return {
      rowAuto: row.getAttribute("data-automationid"),
      nameCellText: (nameCell?.textContent || "").trim(),
      iconAlt: img?.getAttribute("alt") ?? null,
      sizeCellText: (sizeCell?.textContent || "").trim(),
    };
  });
});
log("DOM rows:", JSON.stringify(rows, null, 2));

const dl = await page.evaluate(() =>
  Array.from(document.querySelectorAll("button,[role='menuitem']"))
    .map((b) => ({ auto: b.getAttribute("data-automationid"), text: (b.textContent || "").trim().slice(0, 30) }))
    .filter((b) => b.auto && /download/i.test(b.auto))
);
log("Download command candidates:", JSON.stringify(dl, null, 2));

await ctx.close();
log("Done. Session persisted at", SESSION_DIR);

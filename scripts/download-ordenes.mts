/**
 * One-off: download the "Orden del día" PDF of every Consejo de Gobierno from
 * the two URJC OneDrive folders and build a JSON manifest linking each local
 * PDF to its OneDrive origin (so extra documents can be fetched later).
 */
import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { OneDriveScraper } from "../src/index.js";
import type { OneDriveEntry } from "../src/index.js";

const ORIGIN = "https://urjc-my.sharepoint.com";
const DEST = "/home/mica/git/catedra-ebtc/consejos_de_gobierno";

const SITES = [
  {
    key: "etsii_director",
    rootServerRelative:
      "/personal/etsii_director_urjc_es/Documents/Públicos/Docs_Consejo_Gobierno",
  },
  {
    key: "dpto_die_director",
    rootServerRelative:
      "/personal/dpto_die_director_urjc_es/Documents/Dpto.die/Consejo Gobierno",
  },
];

/**
 * True if the file name looks like an agenda. Handles the many observed
 * spellings: "Orden del día ...", "Orden_del_dia_0", "Ordendeldía",
 * "Ordendeldia_6", and the Latin-1-encoded "Orden%20del%20d%EDa_1"
 * (%ED can't be decoded, so match it via the spaced form).
 */
function isAgenda(name: string): boolean {
  if (!/\.pdf$/i.test(name)) return false;
  const spaced = name.replace(/%20/gi, " ").replace(/_/g, " ").toLowerCase();
  const compact = spaced.replace(/\s+/g, "");
  return /^orden\s+del\b/.test(spaced) || /^ordendeld[ií]?a/.test(compact);
}

/** Build a browsable folderUrl usable with the library/CLI (--url). */
function folderUrl(serverRelative: string): string {
  return `${ORIGIN}/shared?id=${encodeURIComponent(serverRelative)}`;
}

/** Filesystem-safe base name from a council folder name. */
function safeName(s: string): string {
  return s.replace(/[\/\\]/g, "-").replace(/\s+/g, " ").trim();
}

interface ManifestEntry {
  date: string | null;
  councilFolder: string;
  site: string;
  pdfLocalPath: string;
  pdfFileName: string;
  originalName: string;
  onedriveFileServerRelativeUrl: string;
  onedriveFolderServerRelativeUrl: string;
  onedriveFolderUrl: string;
  sizeBytes: number | null;
}

/** Extract YYYY-MM-DD from a folder name like "2020_01_31" / "2016_9_30". */
function extractDate(folder: string): string | null {
  const m = folder.match(/(\d{4})[_-](\d{1,2})[_-](\d{1,2})/);
  if (!m) return null;
  const [, y, mo, d] = m;
  return `${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
}

async function main(): Promise<void> {
  await mkdir(DEST, { recursive: true });
  const manifest: ManifestEntry[] = [];
  const missing: { site: string; councilFolder: string; children: string[] }[] = [];
  const errors: { site: string; councilFolder?: string; error: string }[] = [];
  const usedNames = new Set<string>();

  for (const site of SITES) {
    const scraper = new OneDriveScraper({
      folderUrl: folderUrl(site.rootServerRelative),
      sessionDir: ".onedrive-session",
      headed: false,
      // dpto site allows REST; etsii is a cross-user shared folder where REST
      // is 403 — "auto" falls back to DOM there.
      strategy: "auto",
    });
    await scraper.connect();
    try {
      const councils = (await scraper.list("")).filter((e) => e.type === "folder");
      console.error(`[${site.key}] ${councils.length} council folders`);

      for (const council of councils) {
        try {
          let children = await scraper.list(council.path);
          let matches = children.filter((c) => c.type === "file" && isAgenda(c.name));

          // Fallback: look one level down (agenda-named files only) if none at top.
          if (matches.length === 0) {
            for (const sub of children.filter((c) => c.type === "folder")) {
              const subChildren = await scraper.list(sub.path);
              matches.push(
                ...subChildren.filter((c) => c.type === "file" && isAgenda(c.name))
              );
            }
          }

          if (matches.length === 0) {
            missing.push({
              site: site.key,
              councilFolder: council.name,
              children: children.map((c) => c.name),
            });
            console.error(`  ✗ no agenda: ${council.name}`);
            continue;
          }

          // Primary = shortest name (canonical); extras keep a distinguishing suffix.
          matches.sort((a, b) => a.name.length - b.name.length);
          for (let i = 0; i < matches.length; i++) {
            const m = matches[i];
            let base = safeName(council.name);
            if (i > 0) {
              const stem = safeName(m.name.replace(/\.pdf$/i, "")).slice(0, 60);
              base = `${base} - ${stem}`;
            }
            let fileName = `${base}.pdf`;
            let n = 2;
            while (usedNames.has(fileName)) fileName = `${base} (${n++}).pdf`;
            usedNames.add(fileName);

            const localPath = join(DEST, fileName);
            // Download by exact server-relative URL (no decode round-trip).
            const bytes: Buffer = await (scraper as any).downloadBytesAt(
              m.serverRelativeUrl
            );
            await writeFile(localPath, bytes);

            manifest.push({
              date: extractDate(council.name),
              councilFolder: council.name,
              site: site.key,
              pdfLocalPath: localPath,
              pdfFileName: fileName,
              originalName: m.name,
              onedriveFileServerRelativeUrl: m.serverRelativeUrl,
              onedriveFolderServerRelativeUrl: council.serverRelativeUrl,
              onedriveFolderUrl: folderUrl(council.serverRelativeUrl),
              sizeBytes: bytes.length,
            });
            console.error(
              `  ↓ ${council.name} -> ${fileName} (${bytes.length} B)`
            );
          }
        } catch (err) {
          errors.push({
            site: site.key,
            councilFolder: council.name,
            error: String((err as Error)?.message ?? err),
          });
          console.error(`  ! error ${council.name}: ${(err as Error)?.message}`);
        }
      }
    } catch (err) {
      errors.push({ site: site.key, error: String((err as Error)?.message ?? err) });
    } finally {
      await scraper.close();
    }
  }

  manifest.sort((a, b) => (a.date ?? "").localeCompare(b.date ?? ""));
  const manifestPath = join(DEST, "consejos_de_gobierno.json");
  await writeFile(manifestPath, JSON.stringify(manifest, null, 2) + "\n");

  console.error("\n===== SUMMARY =====");
  console.error(`Downloaded: ${manifest.length}`);
  console.error(`Councils without agenda: ${missing.length}`);
  missing.forEach((m) => console.error(`  - [${m.site}] ${m.councilFolder}`));
  console.error(`Errors: ${errors.length}`);
  errors.forEach((e) => console.error(`  - [${e.site}] ${e.councilFolder ?? ""}: ${e.error}`));
  console.error(`Manifest: ${manifestPath}`);

  // Persist the diagnostics next to the manifest for review.
  await writeFile(
    join(DEST, "_download_report.json"),
    JSON.stringify({ downloaded: manifest.length, missing, errors }, null, 2) + "\n"
  );
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

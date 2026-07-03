#!/usr/bin/env node
/**
 * Command-line interface for onedrive-scraper.
 *
 * Run `onedrive-scraper --help` (or `npm run cli -- --help` during dev) for
 * usage. See docs/CLI.md for the full reference.
 */

import { writeFile, mkdir } from "node:fs/promises";
import { dirname } from "node:path";
import { Command, Option } from "commander";

import { OneDriveScraper } from "./OneDriveScraper.js";
import { treeToText, treeToJson, formatBytes } from "./tree.js";
import type {
  FilterInput,
  OneDriveScraperOptions,
  Strategy,
  WalkFilters,
} from "./types.js";

interface CommonOpts {
  url?: string;
  site?: string;
  root?: string;
  session?: string;
  strategy?: Strategy;
  headless?: boolean;
  locale?: string;
  timeout?: string;
}

interface FilterOpts {
  include?: string[];
  exclude?: string[];
  includeFolder?: string[];
  excludeFolder?: string[];
  matchPath?: boolean;
  depth?: string;
}

function scraperOptions(opts: CommonOpts): OneDriveScraperOptions {
  if (!opts.url && !opts.site) {
    throw new Error(
      "Provide a target: --url <folderUrl> (recommended) or --site <siteUrl> [--root <path>]."
    );
  }
  return {
    folderUrl: opts.url,
    siteUrl: opts.site,
    rootPath: opts.root,
    sessionDir: opts.session,
    strategy: opts.strategy,
    headed: opts.headless ? false : true,
    locale: opts.locale,
    timeout: opts.timeout ? Number(opts.timeout) : undefined,
  };
}

function toFilterInput(
  patterns: string[] | undefined,
  matchPath: boolean | undefined
): FilterInput | undefined {
  if (!patterns || patterns.length === 0) return undefined;
  if (!matchPath) return patterns;
  return patterns.map((pattern) => ({ pattern, target: "path" as const }));
}

function buildFilters(opts: FilterOpts): WalkFilters {
  return {
    files: {
      include: toFilterInput(opts.include, opts.matchPath),
      exclude: toFilterInput(opts.exclude, opts.matchPath),
    },
    folders: {
      include: toFilterInput(opts.includeFolder, opts.matchPath),
      exclude: toFilterInput(opts.excludeFolder, opts.matchPath),
    },
  };
}

function depthOf(opts: FilterOpts): number | undefined {
  if (opts.depth === undefined) return undefined;
  const n = Number(opts.depth);
  return Number.isFinite(n) ? n : undefined;
}

/** Attach the target/session options shared by every command. */
function withCommon(cmd: Command): Command {
  return cmd
    .option("-u, --url <url>", "OneDrive/SharePoint folder URL (parses its `id` param)")
    .option("-s, --site <url>", "site base URL, e.g. https://org-my.sharepoint.com/personal/user_org_es")
    .option("-r, --root <path>", "server-relative root folder path")
    .option("--session <dir>", "directory for the persisted browser session", ".onedrive-session")
    .addOption(
      new Option("--strategy <mode>", "scraping strategy")
        .choices(["auto", "rest", "dom"])
        .default("auto")
    )
    .option("--headless", "run the browser headless (only works with an already-persisted session)", false)
    .option("--locale <locale>", "browser locale", "en-US")
    .option("--timeout <ms>", "per-operation timeout in ms", "30000");
}

/** Attach filter + depth options. */
function withFilters(cmd: Command): Command {
  return cmd
    .option("--include <glob...>", "only include files matching these glob(s)")
    .option("--exclude <glob...>", "exclude files matching these glob(s)")
    .option("--include-folder <glob...>", "only explore folders matching these glob(s)")
    .option("--exclude-folder <glob...>", "skip folders matching these glob(s)")
    .option("--match-path", "match filters against the full relative path instead of the name", false)
    .option("--depth <n>", "maximum recursion depth");
}

async function withScraper<T>(
  common: CommonOpts,
  fn: (scraper: OneDriveScraper) => Promise<T>
): Promise<T> {
  const scraper = new OneDriveScraper(scraperOptions(common));
  await scraper.connect();
  try {
    return await fn(scraper);
  } finally {
    await scraper.close();
  }
}

const program = new Command();

program
  .name("onedrive-scraper")
  .description(
    "Read OneDrive/SharePoint folders and download files via a persisted authenticated browser session."
  )
  .version("0.1.0");

// login -----------------------------------------------------------------------
withCommon(
  program
    .command("login")
    .description("Open a browser to authenticate and persist the session")
).action(async (opts: CommonOpts) => {
  await withScraper(opts, async () => {
    console.error("Authenticated. Session persisted at:", opts.session);
  });
});

// list ------------------------------------------------------------------------
withCommon(
  program
    .command("list [path]")
    .description("List the immediate children of a folder (relative to root)")
    .option("--json", "output JSON instead of a table", false)
).action(async (path: string | undefined, opts: CommonOpts & { json?: boolean }) => {
  await withScraper(opts, async (scraper) => {
    const entries = await scraper.list(path ?? "");
    if (opts.json) {
      console.log(
        JSON.stringify(
          entries.map((e) => ({ ...e, modified: e.modified?.toISOString() })),
          null,
          2
        )
      );
      return;
    }
    for (const e of entries) {
      const kind = e.type === "folder" ? "d" : "-";
      const size = e.type === "file" && e.size !== undefined ? formatBytes(e.size) : "";
      console.log(`${kind}  ${e.name.padEnd(40)} ${size}`);
    }
  });
});

// tree ------------------------------------------------------------------------
withFilters(
  withCommon(
    program
      .command("tree [path]")
      .description("Print or export the folder structure as text and/or JSON")
      .option("--out-txt <file>", "write the tree as text to a file")
      .option("--out-json <file>", "write the tree as JSON to a file")
      .option("--show-size", "show file sizes in the text output", false)
  )
).action(
  async (
    path: string | undefined,
    opts: CommonOpts &
      FilterOpts & { outTxt?: string; outJson?: string; showSize?: boolean }
  ) => {
    await withScraper(opts, async (scraper) => {
      const tree = await scraper.getTree(path ?? "", {
        ...buildFilters(opts),
        maxDepth: depthOf(opts),
      });
      const text = treeToText(tree, { showSize: opts.showSize });
      const json = JSON.stringify(treeToJson(tree), null, 2);

      if (opts.outTxt) {
        await writeOut(opts.outTxt, text + "\n");
        console.error("Wrote", opts.outTxt);
      }
      if (opts.outJson) {
        await writeOut(opts.outJson, json + "\n");
        console.error("Wrote", opts.outJson);
      }
      if (!opts.outTxt && !opts.outJson) {
        console.log(text);
      }
    });
  }
);

// download --------------------------------------------------------------------
withFilters(
  withCommon(
    program
      .command("download [path]")
      .description("Download files under a folder (with optional filters)")
      .option("-d, --dest <dir>", "local destination directory", "downloads")
      .option("--no-reconstruct", "download all files flat instead of recreating the folder tree")
      .option("--overwrite", "overwrite files that already exist locally", false)
  )
).action(
  async (
    path: string | undefined,
    opts: CommonOpts &
      FilterOpts & { dest: string; reconstruct: boolean; overwrite?: boolean }
  ) => {
    await withScraper(opts, async (scraper) => {
      const results = await scraper.download(path ?? "", {
        ...buildFilters(opts),
        destDir: opts.dest,
        reconstruct: opts.reconstruct,
        overwrite: opts.overwrite,
        maxDepth: depthOf(opts),
        onProgress: (p) => {
          const tag =
            p.status === "downloaded" ? "↓" : p.status === "skipped" ? "=" : "✗";
          const extra =
            p.status === "downloaded" && p.bytes !== undefined
              ? ` (${formatBytes(p.bytes)})`
              : p.status === "error"
                ? ` — ${p.error?.message}`
                : "";
          console.error(`[${p.index + 1}/${p.total}] ${tag} ${p.entry.path}${extra}`);
        },
      });
      const ok = results.filter((r) => r.status === "downloaded").length;
      const skipped = results.filter((r) => r.status === "skipped").length;
      const failed = results.filter((r) => r.status === "error").length;
      console.error(
        `Done: ${ok} downloaded, ${skipped} skipped, ${failed} failed → ${opts.dest}`
      );
      if (failed > 0) process.exitCode = 1;
    });
  }
);

async function writeOut(file: string, content: string): Promise<void> {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, content, "utf8");
}

program.parseAsync(process.argv).catch((error: unknown) => {
  console.error("Error:", error instanceof Error ? error.message : error);
  process.exit(1);
});

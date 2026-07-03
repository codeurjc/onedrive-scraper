/**
 * High-level OneDrive scraper.
 *
 * Presents a filesystem-like API (`list`, `stat`, `exists`, `readFile`,
 * `downloadFile`) plus high-level operations (`getTree`, `download`) over a
 * OneDrive/SharePoint folder, using a persisted authenticated browser session.
 *
 * All paths passed to the public methods are **relative to the configured root
 * folder** and use `/` separators (the root is `""` or `"."`).
 */

import { mkdir, writeFile, access } from "node:fs/promises";
import { constants as FS } from "node:fs";
import { dirname, join } from "node:path";

import { BrowserSession } from "./browser.js";
import { RestClient } from "./rest.js";
import { downloadFileViaDom, listChildrenViaDom } from "./dom.js";
import { compileFilter } from "./filters.js";
import { sortEntries } from "./tree.js";
import {
  basename,
  relativeFromServerRelative,
  resolveLocation,
  resolveServerRelative,
} from "./paths.js";
import type {
  DownloadOptions,
  DownloadProgress,
  OneDriveEntry,
  OneDriveScraperOptions,
  OneDriveTreeNode,
  ResolvedLocation,
  Strategy,
  TreeOptions,
  WalkFilters,
} from "./types.js";

const DEFAULTS = {
  sessionDir: ".onedrive-session",
  headed: true,
  strategy: "auto" as Strategy,
  loginTimeout: 300_000,
  locale: "en-US",
  timeout: 30_000,
};

export class OneDriveScraper {
  readonly location: ResolvedLocation;
  private readonly session: BrowserSession;
  private readonly strategy: Strategy;
  private rest?: RestClient;
  private connected = false;

  constructor(private readonly options: OneDriveScraperOptions) {
    this.location = resolveLocation(options);
    this.strategy = options.strategy ?? DEFAULTS.strategy;
    this.session = new BrowserSession({
      sessionDir: options.sessionDir ?? DEFAULTS.sessionDir,
      origin: this.location.origin,
      siteUrl: this.location.siteUrl,
      headed: options.headed ?? DEFAULTS.headed,
      locale: options.locale ?? DEFAULTS.locale,
      timeout: options.timeout ?? DEFAULTS.timeout,
      loginTimeout: options.loginTimeout ?? DEFAULTS.loginTimeout,
    });
  }

  /**
   * Launch the browser and ensure the session is authenticated (driving an
   * interactive login the first time). Must be called before any operation.
   */
  async connect(): Promise<void> {
    if (this.connected) return;
    await this.session.ensureAuthenticated();
    this.rest = new RestClient(
      this.session.page,
      this.location.siteUrl,
      this.options.timeout ?? DEFAULTS.timeout
    );
    this.connected = true;
  }

  /** Close the browser, flushing the persisted profile to disk. */
  async close(): Promise<void> {
    await this.session.close();
    this.connected = false;
  }

  /** The active Playwright page (advanced use / custom scraping). */
  get page() {
    return this.session.page;
  }

  // ---------------------------------------------------------------------------
  // Filesystem-like API
  // ---------------------------------------------------------------------------

  /** List the immediate children of a folder (relative to root). Like `readdir`. */
  async list(relativePath = ""): Promise<OneDriveEntry[]> {
    this.assertConnected();
    const folder = resolveServerRelative(
      this.location.rootServerRelativeUrl,
      relativePath
    );
    const entries = await this.listChildrenAt(folder);
    return entries.sort(sortEntries);
  }

  /** Get metadata for a single entry (relative to root). Like `stat`. */
  async stat(relativePath: string): Promise<OneDriveEntry> {
    this.assertConnected();
    let clean = relativePath.replace(/\\/g, "/").trim();
    if (clean === "." || clean === "./") clean = "";
    clean = clean.replace(/^\.\//, "").replace(/^\/+|\/+$/g, "");
    if (!clean) {
      return {
        name: basename(this.location.rootServerRelativeUrl),
        path: "",
        serverRelativeUrl: this.location.rootServerRelativeUrl,
        type: "folder",
      };
    }
    const parent = clean.split("/").slice(0, -1).join("/");
    const name = clean.split("/").pop()!;
    const siblings = await this.list(parent);
    const found = siblings.find((e) => e.name === name);
    if (!found) {
      throw new Error(`Entry not found: ${relativePath}`);
    }
    return found;
  }

  /** Whether an entry exists at the given relative path. */
  async exists(relativePath: string): Promise<boolean> {
    try {
      await this.stat(relativePath);
      return true;
    } catch {
      return false;
    }
  }

  /** Read a file's bytes into a Buffer. Like `readFile`. */
  async readFile(relativePath: string): Promise<Buffer> {
    this.assertConnected();
    const file = resolveServerRelative(
      this.location.rootServerRelativeUrl,
      relativePath
    );
    return this.downloadBytesAt(file);
  }

  /** Download a single file to a local path. */
  async downloadFile(relativePath: string, localPath: string): Promise<void> {
    const bytes = await this.readFile(relativePath);
    await mkdir(dirname(localPath), { recursive: true });
    await writeFile(localPath, bytes);
  }

  // ---------------------------------------------------------------------------
  // High-level operations
  // ---------------------------------------------------------------------------

  /**
   * Recursively build a tree of a folder, applying file/folder filters.
   * Folders rejected by the folder filter are neither listed nor descended.
   */
  async getTree(
    relativePath = "",
    options: TreeOptions = {}
  ): Promise<OneDriveTreeNode> {
    this.assertConnected();
    const fileFilter = compileFilter(options.files);
    const folderFilter = compileFilter(options.folders);
    const maxDepth = options.maxDepth ?? Infinity;

    const rootServerRelative = resolveServerRelative(
      this.location.rootServerRelativeUrl,
      relativePath
    );
    const root: OneDriveTreeNode = {
      name: basename(rootServerRelative),
      path: relativeFromServerRelative(
        this.location.rootServerRelativeUrl,
        rootServerRelative
      ),
      serverRelativeUrl: rootServerRelative,
      type: "folder",
      children: [],
    };

    const build = async (
      node: OneDriveTreeNode,
      depth: number
    ): Promise<void> => {
      if (depth > maxDepth) return;
      const entries = await this.listChildrenAt(node.serverRelativeUrl);
      const children: OneDriveTreeNode[] = [];
      for (const entry of entries.sort(sortEntries)) {
        if (entry.type === "folder") {
          if (!folderFilter.test(entry)) continue;
          const child: OneDriveTreeNode = { ...entry, children: [] };
          if (depth < maxDepth) await build(child, depth + 1);
          children.push(child);
        } else if (fileFilter.test(entry)) {
          children.push({ ...entry });
        }
      }
      node.children = children;
    };

    await build(root, 1);
    return root;
  }

  /**
   * Download all files under a folder that pass the filters. Recreates the
   * remote folder structure on disk unless `reconstruct` is `false`.
   */
  async download(
    relativePath = "",
    options: DownloadOptions = {}
  ): Promise<DownloadProgress[]> {
    this.assertConnected();
    const destDir = options.destDir ?? process.cwd();
    const reconstruct = options.reconstruct ?? true;
    const overwrite = options.overwrite ?? false;

    const tree = await this.getTree(relativePath, {
      files: options.files,
      folders: options.folders,
      maxDepth: options.maxDepth,
    });

    const files = collectFiles(tree);
    const rootRelative = tree.path;
    const results: DownloadProgress[] = [];

    for (let i = 0; i < files.length; i++) {
      const entry = files[i];
      const relForDisk = reconstruct
        ? stripPrefix(entry.path, rootRelative)
        : entry.name;
      const localPath = join(destDir, relForDisk);
      const progress: DownloadProgress = {
        entry,
        localPath,
        status: "downloaded",
        index: i,
        total: files.length,
      };
      try {
        if (!overwrite && (await fileExists(localPath))) {
          progress.status = "skipped";
        } else {
          const bytes = await this.downloadBytesAt(entry.serverRelativeUrl);
          await mkdir(dirname(localPath), { recursive: true });
          await writeFile(localPath, bytes);
          progress.bytes = bytes.length;
        }
      } catch (error) {
        progress.status = "error";
        progress.error = error instanceof Error ? error : new Error(String(error));
      }
      options.onProgress?.(progress);
      results.push(progress);
    }
    return results;
  }

  // ---------------------------------------------------------------------------
  // Strategy dispatch
  // ---------------------------------------------------------------------------

  private async listChildrenAt(
    folderServerRelative: string
  ): Promise<OneDriveEntry[]> {
    if (this.strategy === "dom") {
      return listChildrenViaDom(
        this.session.page,
        this.location.siteUrl,
        folderServerRelative,
        this.location.rootServerRelativeUrl
      );
    }
    try {
      return await this.rest!.listChildren(
        folderServerRelative,
        this.location.rootServerRelativeUrl
      );
    } catch (error) {
      if (this.strategy === "rest") throw error;
      return listChildrenViaDom(
        this.session.page,
        this.location.siteUrl,
        folderServerRelative,
        this.location.rootServerRelativeUrl
      );
    }
  }

  private async downloadBytesAt(fileServerRelative: string): Promise<Buffer> {
    if (this.strategy === "dom") {
      return downloadFileViaDom(
        this.session.page,
        this.location.siteUrl,
        fileServerRelative
      );
    }
    try {
      return await this.rest!.downloadFile(fileServerRelative);
    } catch (error) {
      if (this.strategy === "rest") throw error;
      return downloadFileViaDom(
        this.session.page,
        this.location.siteUrl,
        fileServerRelative
      );
    }
  }

  private assertConnected(): void {
    if (!this.connected) {
      throw new Error("Not connected. Call connect() before using the scraper.");
    }
  }
}

function collectFiles(node: OneDriveTreeNode): OneDriveEntry[] {
  const files: OneDriveEntry[] = [];
  const walk = (n: OneDriveTreeNode) => {
    for (const child of n.children ?? []) {
      if (child.type === "folder") walk(child);
      else files.push(child);
    }
  };
  walk(node);
  return files;
}

function stripPrefix(path: string, prefix: string): string {
  if (!prefix) return path;
  if (path === prefix) return basename(path);
  if (path.startsWith(prefix + "/")) return path.slice(prefix.length + 1);
  return path;
}

async function fileExists(path: string): Promise<boolean> {
  try {
    await access(path, FS.F_OK);
    return true;
  } catch {
    return false;
  }
}

export type {
  WalkFilters,
  DownloadOptions,
  DownloadProgress,
  TreeOptions,
} from "./types.js";

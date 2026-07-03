/**
 * onedrive-scraper — read OneDrive/SharePoint folders and download files using
 * a persisted, interactively-authenticated browser session.
 *
 * @example
 * ```ts
 * import { OneDriveScraper } from "onedrive-scraper";
 *
 * const scraper = new OneDriveScraper({
 *   folderUrl:
 *     "https://org-my.sharepoint.com/my?id=%2Fpersonal%2Fuser_org_es%2FDocuments%2FRoot",
 * });
 * await scraper.connect();               // opens a browser for MFA login the first time
 * const entries = await scraper.list();  // like readdir
 * await scraper.download("", { destDir: "./downloads", files: { include: "*.docx" } });
 * await scraper.close();
 * ```
 */

export { OneDriveScraper } from "./OneDriveScraper.js";

export {
  treeToText,
  treeToJson,
  flattenTree,
  sortEntries,
  formatBytes,
  type SerializedTreeNode,
  type TreeTextOptions,
} from "./tree.js";

export {
  compileFilter,
  globToRegExp,
  type CompiledFilter,
} from "./filters.js";

export {
  resolveLocation,
  serverRelativeFromUrl,
  normalizeServerRelative,
} from "./paths.js";

export type {
  OneDriveScraperOptions,
  OneDriveEntry,
  OneDriveTreeNode,
  EntryType,
  Strategy,
  FilterInput,
  FilterOptions,
  FilterPattern,
  FilterRule,
  FilterTarget,
  WalkFilters,
  TreeOptions,
  DownloadOptions,
  DownloadProgress,
  ResolvedLocation,
} from "./types.js";

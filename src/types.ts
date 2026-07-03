/**
 * Shared type definitions for the OneDrive scraper library.
 */

/** Kind of a filesystem-like entry. */
export type EntryType = "file" | "folder";

/**
 * A single entry (file or folder) in a OneDrive/SharePoint folder.
 * Modelled after a filesystem `Dirent`/`Stats` so the API feels familiar.
 */
export interface OneDriveEntry {
  /** Base name, e.g. `Doc1.docx`. */
  name: string;
  /**
   * Path relative to the configured root folder, using `/` separators.
   * The root itself is the empty string `""`.
   * Example: `Folder/Doc3.pptx`.
   */
  path: string;
  /**
   * Full server-relative URL as SharePoint knows it.
   * Example: `/personal/user_org_es/Documents/Root/Folder/Doc3.pptx`.
   */
  serverRelativeUrl: string;
  /** `"file"` or `"folder"`. */
  type: EntryType;
  /** Size in bytes. Files only. */
  size?: number;
  /** Last-modified timestamp when available. */
  modified?: Date;
  /** Number of children. Folders only. */
  childCount?: number;
}

/** A tree node: an entry plus, for folders, its children. */
export interface OneDriveTreeNode extends OneDriveEntry {
  /** Present on folders. Undefined on files. */
  children?: OneDriveTreeNode[];
}

/** Which string a filter pattern is matched against. */
export type FilterTarget = "name" | "path";

/**
 * A filter pattern. Can be:
 * - a glob string (`*.docx`, `Folder/**`),
 * - a `RegExp`,
 * - a predicate `(value, entry) => boolean`.
 */
export type FilterPattern =
  | string
  | RegExp
  | ((value: string, entry: OneDriveEntry) => boolean);

/** A pattern together with the target it matches against. */
export interface FilterRule {
  pattern: FilterPattern;
  /** `"name"` (default) matches the base name; `"path"` matches the full relative path. */
  target?: FilterTarget;
}

/** A single pattern, a rule, or an array of either. */
export type FilterInput =
  | FilterPattern
  | FilterRule
  | Array<FilterPattern | FilterRule>;

/** Include/exclude filters. An entry passes when it matches `include` (if any) and no `exclude`. */
export interface FilterOptions {
  include?: FilterInput;
  exclude?: FilterInput;
}

/** Filters applied while walking a tree. */
export interface WalkFilters {
  /** Filters applied to files. */
  files?: FilterOptions;
  /** Filters applied to folders (a rejected folder is not descended into). */
  folders?: FilterOptions;
}

/** Strategy used to talk to OneDrive. */
export type Strategy = "auto" | "rest" | "dom";

/** Options for constructing a {@link OneDriveScraper}. */
export interface OneDriveScraperOptions {
  /**
   * Directory used to persist the browser profile (cookies, tokens, cache) so
   * the user does not have to re-authenticate every run.
   * Default: `./.onedrive-session`.
   */
  sessionDir?: string;
  /**
   * A OneDrive/SharePoint folder URL. The `id` query parameter is parsed to
   * derive the site URL and the root folder. Mutually usable with
   * {@link siteUrl}/{@link rootPath}.
   */
  folderUrl?: string;
  /**
   * Explicit site (web) base URL, e.g.
   * `https://org-my.sharepoint.com/personal/user_org_es`.
   */
  siteUrl?: string;
  /**
   * Server-relative root path, e.g.
   * `/personal/user_org_es/Documents/Root`.
   */
  rootPath?: string;
  /** Show the browser window. Required for interactive MFA login. Default `true`. */
  headed?: boolean;
  /** Preferred strategy. `"auto"` = REST, falling back to DOM. Default `"auto"`. */
  strategy?: Strategy;
  /** Max time (ms) to wait for the user to complete interactive login. Default `300000`. */
  loginTimeout?: number;
  /** Browser locale. Default `"es-ES"`. */
  locale?: string;
  /** Per-request/navigation timeout (ms). Default `30000`. */
  timeout?: number;
}

/** Options for {@link OneDriveScraper.getTree} / walking. */
export interface TreeOptions extends WalkFilters {
  /** Max recursion depth. `Infinity` (default) walks the whole subtree. */
  maxDepth?: number;
}

/** Options for {@link OneDriveScraper.download}. */
export interface DownloadOptions extends WalkFilters {
  /** Local destination directory. Default: current working directory. */
  destDir?: string;
  /**
   * Recreate the remote folder structure on disk. When `false`, all matching
   * files are downloaded flat into `destDir`. Default `true`.
   */
  reconstruct?: boolean;
  /** Overwrite files that already exist locally. Default `false` (skip). */
  overwrite?: boolean;
  /** Max recursion depth. Default `Infinity`. */
  maxDepth?: number;
  /** Called after each file is downloaded (or skipped). */
  onProgress?: (info: DownloadProgress) => void;
}

/** Progress information reported during a bulk download. */
export interface DownloadProgress {
  entry: OneDriveEntry;
  localPath: string;
  status: "downloaded" | "skipped" | "error";
  index: number;
  total: number;
  error?: Error;
  bytes?: number;
}

/** Resolved location derived from options (internal but exported for advanced use). */
export interface ResolvedLocation {
  /** Origin, e.g. `https://org-my.sharepoint.com`. */
  origin: string;
  /** Site base URL, e.g. `https://org-my.sharepoint.com/personal/user_org_es`. */
  siteUrl: string;
  /** Site server-relative URL, e.g. `/personal/user_org_es`. */
  siteServerRelativeUrl: string;
  /** Root folder server-relative URL. */
  rootServerRelativeUrl: string;
}

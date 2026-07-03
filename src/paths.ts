/**
 * URL and server-relative-path helpers.
 *
 * OneDrive personal sites live under `/personal/<user>` and SharePoint team
 * sites under `/sites/<name>` (or `/teams/<name>`). The "web" (site) is the
 * first two path segments; everything below is content.
 */

import type { OneDriveScraperOptions, ResolvedLocation } from "./types.js";

const WEB_PREFIXES = ["personal", "sites", "teams"];

/**
 * Decode percent-escapes, tolerating strings that are already decoded and may
 * contain a stray `%` (SharePoint REST returns `ServerRelativeUrl` decoded, so
 * a name like `50%.pdf` would otherwise make `decodeURIComponent` throw).
 */
function safeDecode(s: string): string {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

/** Normalise a server-relative path: leading slash, no trailing slash, collapse `//`. */
export function normalizeServerRelative(path: string): string {
  let p = safeDecode(path.trim());
  p = p.replace(/\\/g, "/").replace(/\/+/g, "/");
  if (!p.startsWith("/")) p = "/" + p;
  if (p.length > 1 && p.endsWith("/")) p = p.slice(0, -1);
  return p;
}

/**
 * Given a full server-relative content path, return the server-relative URL of
 * the site (web) that contains it, e.g.
 * `/personal/user_org_es/Documents/Root` -> `/personal/user_org_es`.
 */
export function siteServerRelativeFromPath(serverRelative: string): string {
  const segments = normalizeServerRelative(serverRelative)
    .split("/")
    .filter(Boolean);
  if (segments.length >= 2 && WEB_PREFIXES.includes(segments[0].toLowerCase())) {
    return "/" + segments.slice(0, 2).join("/");
  }
  // Root-level site.
  return "";
}

/**
 * Extract the `id` query parameter (the server-relative folder path) from a
 * OneDrive/SharePoint web URL such as
 * `https://org-my.sharepoint.com/my?id=%2Fpersonal%2F...%2FRoot`.
 * Returns `null` if there is no usable `id`.
 */
export function serverRelativeFromUrl(url: string): string | null {
  const u = new URL(url);
  const id = u.searchParams.get("id");
  if (id) return normalizeServerRelative(id);
  return null;
}

/**
 * Resolve site URL and root folder from the supplied options. Accepts either a
 * `folderUrl` (whose `id` param is parsed) or explicit `siteUrl` + `rootPath`.
 */
export function resolveLocation(
  options: OneDriveScraperOptions
): ResolvedLocation {
  let origin: string | undefined;
  let siteServerRelativeUrl: string | undefined;
  let rootServerRelativeUrl: string | undefined;

  if (options.folderUrl) {
    const u = new URL(options.folderUrl);
    origin = u.origin;
    const fromId = serverRelativeFromUrl(options.folderUrl);
    if (fromId) {
      rootServerRelativeUrl = fromId;
      siteServerRelativeUrl = siteServerRelativeFromPath(fromId);
    } else {
      // No id param — treat the pathname as the site.
      siteServerRelativeUrl = siteServerRelativeFromPath(u.pathname);
    }
  }

  if (options.siteUrl) {
    const u = new URL(options.siteUrl);
    origin = u.origin;
    siteServerRelativeUrl = normalizeServerRelative(u.pathname) === "/"
      ? ""
      : normalizeServerRelative(u.pathname);
  }

  if (options.rootPath) {
    rootServerRelativeUrl = normalizeServerRelative(options.rootPath);
    if (siteServerRelativeUrl === undefined) {
      siteServerRelativeUrl = siteServerRelativeFromPath(rootServerRelativeUrl);
    }
  }

  if (!origin) {
    throw new Error(
      "Cannot resolve OneDrive location: provide `folderUrl` or `siteUrl` so the origin is known."
    );
  }
  if (rootServerRelativeUrl === undefined) {
    throw new Error(
      "Cannot resolve root folder: provide `folderUrl` (with an `id` param) or `rootPath`."
    );
  }
  if (siteServerRelativeUrl === undefined) {
    siteServerRelativeUrl = siteServerRelativeFromPath(rootServerRelativeUrl);
  }

  return {
    origin,
    siteServerRelativeUrl,
    siteUrl: origin + siteServerRelativeUrl,
    rootServerRelativeUrl,
  };
}

/**
 * Join the root server-relative URL with a path relative to the root.
 * `relative === ""` returns the root itself.
 */
export function resolveServerRelative(
  rootServerRelativeUrl: string,
  relative: string
): string {
  let rel = relative.replace(/\\/g, "/").trim();
  if (rel === "." || rel === "./") rel = "";
  rel = rel.replace(/^\.\//, "").replace(/^\/+/, "").replace(/\/+$/, "");
  if (!rel) return normalizeServerRelative(rootServerRelativeUrl);
  return normalizeServerRelative(rootServerRelativeUrl + "/" + rel);
}

/**
 * Compute a path relative to the root from a full server-relative URL.
 * Returns `""` for the root itself.
 */
export function relativeFromServerRelative(
  rootServerRelativeUrl: string,
  serverRelativeUrl: string
): string {
  const root = normalizeServerRelative(rootServerRelativeUrl);
  const full = normalizeServerRelative(serverRelativeUrl);
  if (full === root) return "";
  if (full.startsWith(root + "/")) return full.slice(root.length + 1);
  // Not under root — fall back to the base name.
  return full.split("/").filter(Boolean).pop() ?? "";
}

/** Escape a server-relative path for use inside an OData string literal in a URL. */
export function odataPathLiteral(serverRelative: string): string {
  const doubled = normalizeServerRelative(serverRelative).replace(/'/g, "''");
  return `'${encodeURIComponent(doubled)}'`;
}

/** Base name of a server-relative path. */
export function basename(serverRelative: string): string {
  return normalizeServerRelative(serverRelative).split("/").filter(Boolean).pop() ?? "";
}

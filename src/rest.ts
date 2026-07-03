/**
 * SharePoint REST access via the authenticated browser session.
 *
 * This is the primary (and most robust) strategy. It uses the browser
 * context's `APIRequestContext`, which shares the session cookies established
 * during interactive login, to call the same internal REST endpoints the
 * OneDrive web UI uses. No app registration / Graph API access is required.
 *
 * Endpoints used (all read-only):
 * - `_api/web/GetFolderByServerRelativeUrl('<path>')/Folders`
 * - `_api/web/GetFolderByServerRelativeUrl('<path>')/Files`
 * - `_api/web/GetFileByServerRelativeUrl('<path>')/$value`
 */

import type { APIRequestContext, Page } from "playwright";
import type { OneDriveEntry } from "./types.js";
import {
  odataPathLiteral,
  relativeFromServerRelative,
  normalizeServerRelative,
} from "./paths.js";

interface RawFile {
  Name: string;
  ServerRelativeUrl: string;
  Length?: string | number;
  TimeLastModified?: string;
}

interface RawFolder {
  Name: string;
  ServerRelativeUrl: string;
  ItemCount?: number;
  TimeLastModified?: string;
}

const JSON_HEADERS = { Accept: "application/json;odata=nometadata" };

/** Thin client over the SharePoint REST API using the browser session cookies. */
export class RestClient {
  constructor(
    private readonly page: Page,
    private readonly siteUrl: string,
    private readonly timeout: number
  ) {}

  private get request(): APIRequestContext {
    return this.page.context().request;
  }

  /** List the immediate children (files + folders) of a folder. */
  async listChildren(
    folderServerRelative: string,
    rootServerRelative: string
  ): Promise<OneDriveEntry[]> {
    const literal = odataPathLiteral(folderServerRelative);
    const base = `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl(${literal})`;

    const [folders, files] = await Promise.all([
      this.getJson<{ value: RawFolder[] }>(
        `${base}/Folders?$select=Name,ServerRelativeUrl,ItemCount,TimeLastModified`
      ),
      this.getJson<{ value: RawFile[] }>(
        `${base}/Files?$select=Name,ServerRelativeUrl,Length,TimeLastModified`
      ),
    ]);

    const entries: OneDriveEntry[] = [];

    for (const f of folders.value ?? []) {
      // Skip the implicit "Forms" system folder that appears at library roots.
      if (f.Name === "Forms") continue;
      const serverRelativeUrl = normalizeServerRelative(f.ServerRelativeUrl);
      entries.push({
        name: f.Name,
        path: relativeFromServerRelative(rootServerRelative, serverRelativeUrl),
        serverRelativeUrl,
        type: "folder",
        childCount: f.ItemCount,
        modified: f.TimeLastModified ? new Date(f.TimeLastModified) : undefined,
      });
    }

    for (const f of files.value ?? []) {
      const serverRelativeUrl = normalizeServerRelative(f.ServerRelativeUrl);
      entries.push({
        name: f.Name,
        path: relativeFromServerRelative(rootServerRelative, serverRelativeUrl),
        serverRelativeUrl,
        type: "file",
        size: f.Length !== undefined ? Number(f.Length) : undefined,
        modified: f.TimeLastModified ? new Date(f.TimeLastModified) : undefined,
      });
    }

    return entries;
  }

  /** Download a file's raw bytes. */
  async downloadFile(fileServerRelative: string): Promise<Buffer> {
    const literal = odataPathLiteral(fileServerRelative);
    // GetFileByServerRelativePath(decodedurl=...) is the modern endpoint: it
    // handles names with special characters (including literal `%`) that the
    // legacy GetFileByServerRelativeUrl mangles and returns 404 for.
    const url = `${this.siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl=${literal})/$value`;
    const res = await this.request.get(url, { timeout: this.timeout });
    if (!res.ok()) {
      throw new Error(
        `REST download failed for ${fileServerRelative}: ${res.status()} ${res.statusText()}`
      );
    }
    return res.body();
  }

  private async getJson<T>(url: string): Promise<T> {
    const res = await this.request.get(url, {
      headers: JSON_HEADERS,
      timeout: this.timeout,
    });
    if (!res.ok()) {
      throw new Error(`REST GET failed: ${res.status()} ${res.statusText()} — ${url}`);
    }
    return (await res.json()) as T;
  }
}

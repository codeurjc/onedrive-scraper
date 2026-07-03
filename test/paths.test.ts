/**
 * Deep unit tests for the URL / server-relative path helpers.
 */
import { test } from "node:test";
import assert from "node:assert/strict";

import {
  normalizeServerRelative,
  siteServerRelativeFromPath,
  serverRelativeFromUrl,
  resolveLocation,
  resolveServerRelative,
  relativeFromServerRelative,
  odataPathLiteral,
  basename,
} from "../src/paths.js";

test("normalizeServerRelative: slashes, backslashes, trailing, encoded", () => {
  assert.equal(normalizeServerRelative("a/b/c"), "/a/b/c");
  assert.equal(normalizeServerRelative("/a/b/"), "/a/b");
  assert.equal(normalizeServerRelative("//a//b//"), "/a/b");
  assert.equal(normalizeServerRelative("a\\b\\c"), "/a/b/c");
  assert.equal(normalizeServerRelative("%2Fa%2Fb"), "/a/b");
  assert.equal(normalizeServerRelative("/"), "/");
  assert.equal(normalizeServerRelative("  /a/b  "), "/a/b");
});

test("siteServerRelativeFromPath: personal, sites, teams, root-level", () => {
  assert.equal(
    siteServerRelativeFromPath("/personal/u_org_es/Documents/X/Y"),
    "/personal/u_org_es"
  );
  assert.equal(
    siteServerRelativeFromPath("/sites/Team/Shared Documents/X"),
    "/sites/Team"
  );
  assert.equal(siteServerRelativeFromPath("/teams/T/Docs"), "/teams/T");
  assert.equal(siteServerRelativeFromPath("/foo/bar"), "");
  assert.equal(siteServerRelativeFromPath("/Documents/X"), "");
});

test("serverRelativeFromUrl: decodes the id param, null when absent", () => {
  assert.equal(
    serverRelativeFromUrl(
      "https://org-my.sharepoint.com/my?id=%2Fpersonal%2Fu_org_es%2FDocuments%2FRoot"
    ),
    "/personal/u_org_es/Documents/Root"
  );
  assert.equal(
    serverRelativeFromUrl("https://org-my.sharepoint.com/personal/u/_layouts/15/onedrive.aspx"),
    null
  );
});

test("resolveLocation: SharePoint team-site URL", () => {
  const loc = resolveLocation({
    folderUrl:
      "https://org.sharepoint.com/sites/Team/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FTeam%2FShared%20Documents%2FReports",
  });
  assert.equal(loc.origin, "https://org.sharepoint.com");
  assert.equal(loc.siteUrl, "https://org.sharepoint.com/sites/Team");
  assert.equal(loc.rootServerRelativeUrl, "/sites/Team/Shared Documents/Reports");
});

test("resolveLocation: throws with insufficient options", () => {
  assert.throws(() => resolveLocation({}), /origin/);
  assert.throws(
    () => resolveLocation({ siteUrl: "https://org-my.sharepoint.com/personal/u" }),
    /root folder/
  );
});

test("resolveServerRelative: joins relative paths under the root", () => {
  assert.equal(resolveServerRelative("/a/b", ""), "/a/b");
  assert.equal(resolveServerRelative("/a/b", "."), "/a/b");
  assert.equal(resolveServerRelative("/a/b", "c/d"), "/a/b/c/d");
  assert.equal(resolveServerRelative("/a/b", "/c/"), "/a/b/c");
  assert.equal(resolveServerRelative("/a/b", "c\\d"), "/a/b/c/d");
});

test("relativeFromServerRelative: strips the root prefix", () => {
  assert.equal(relativeFromServerRelative("/a/b", "/a/b"), "");
  assert.equal(relativeFromServerRelative("/a/b", "/a/b/c/d"), "c/d");
  // Not under root -> falls back to base name.
  assert.equal(relativeFromServerRelative("/a/b", "/x/y/z.txt"), "z.txt");
});

test("odataPathLiteral: doubles quotes and URL-encodes", () => {
  assert.equal(odataPathLiteral("/a/b"), "'%2Fa%2Fb'");
  assert.equal(odataPathLiteral("/a/O'Brien"), "'%2Fa%2FO''Brien'");
});

test("basename", () => {
  assert.equal(basename("/a/b/c.txt"), "c.txt");
  assert.equal(basename("/a/"), "a");
  assert.equal(basename("/"), "");
});

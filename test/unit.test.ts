/**
 * Pure unit tests — no browser / network required.
 * Run with:  npm run test:unit
 */
import { test } from "node:test";
import assert from "node:assert/strict";

import {
  resolveLocation,
  serverRelativeFromUrl,
  normalizeServerRelative,
} from "../src/paths.js";
import { globToRegExp, compileFilter } from "../src/filters.js";
import { treeToText, treeToJson, flattenTree, formatBytes } from "../src/tree.js";
import type { OneDriveTreeNode } from "../src/types.js";
import { TEST_FOLDER_URL } from "./expected.js";

test("resolveLocation parses a folder URL into site + root", () => {
  const loc = resolveLocation({ folderUrl: TEST_FOLDER_URL });
  assert.equal(loc.origin, "https://urjc-my.sharepoint.com");
  assert.equal(
    loc.siteUrl,
    "https://urjc-my.sharepoint.com/personal/micael_gallego_urjc_es"
  );
  assert.equal(loc.siteServerRelativeUrl, "/personal/micael_gallego_urjc_es");
  assert.equal(
    loc.rootServerRelativeUrl,
    "/personal/micael_gallego_urjc_es/Documents/PruebaOneDriveScrapper"
  );
});

test("resolveLocation accepts siteUrl + rootPath equivalently", () => {
  const a = resolveLocation({ folderUrl: TEST_FOLDER_URL });
  const b = resolveLocation({
    siteUrl: "https://urjc-my.sharepoint.com/personal/micael_gallego_urjc_es",
    rootPath:
      "/personal/micael_gallego_urjc_es/Documents/PruebaOneDriveScrapper",
  });
  assert.deepEqual(b, a);
});

test("serverRelativeFromUrl / normalizeServerRelative", () => {
  assert.equal(
    serverRelativeFromUrl(TEST_FOLDER_URL),
    "/personal/micael_gallego_urjc_es/Documents/PruebaOneDriveScrapper"
  );
  assert.equal(normalizeServerRelative("a/b/"), "/a/b");
  assert.equal(normalizeServerRelative("//a//b//"), "/a/b");
});

test("globToRegExp supports *, **, ?, {a,b}", () => {
  assert.ok(globToRegExp("*.docx").test("Doc1.docx"));
  assert.ok(!globToRegExp("*.docx").test("Doc2.xlsx"));
  assert.ok(globToRegExp("**/Doc3.pptx").test("Folder/Doc3.pptx"));
  assert.ok(globToRegExp("*.{docx,md}").test("Doc4.md"));
  assert.ok(globToRegExp("Doc?.docx").test("Doc1.docx"));
  assert.ok(!globToRegExp("*.docx").test("Folder/Doc1.docx")); // * excludes '/'
});

test("compileFilter include/exclude by name and by path", () => {
  const e = (name: string, path: string) =>
    ({ name, path, type: "file", serverRelativeUrl: "" }) as const;

  const byName = compileFilter({ include: ["*.docx", "*.md"] });
  assert.ok(byName.test(e("Doc1.docx", "Doc1.docx")));
  assert.ok(byName.test(e("Doc4.md", "Folder/Doc4.md")));
  assert.ok(!byName.test(e("Doc2.xlsx", "Doc2.xlsx")));

  const byPath = compileFilter({
    include: [{ pattern: "Folder/**", target: "path" }],
  });
  assert.ok(byPath.test(e("Doc3.pptx", "Folder/Doc3.pptx")));
  assert.ok(!byPath.test(e("Doc1.docx", "Doc1.docx")));

  const excl = compileFilter({ exclude: "*.xlsx" });
  assert.ok(!excl.test(e("Doc2.xlsx", "Doc2.xlsx")));
  assert.ok(excl.test(e("Doc1.docx", "Doc1.docx")));
});

test("tree serialization to text and JSON", () => {
  const tree: OneDriveTreeNode = {
    name: "PruebaOneDriveScrapper",
    path: "",
    type: "folder",
    serverRelativeUrl: "/x",
    children: [
      { name: "Doc1.docx", path: "Doc1.docx", type: "file", serverRelativeUrl: "/x/Doc1.docx", size: 9452 },
      {
        name: "Folder",
        path: "Folder",
        type: "folder",
        serverRelativeUrl: "/x/Folder",
        children: [
          { name: "Doc4.md", path: "Folder/Doc4.md", type: "file", serverRelativeUrl: "/x/Folder/Doc4.md", size: 306 },
        ],
      },
    ],
  };

  const text = treeToText(tree, { showSize: true });
  assert.match(text, /PruebaOneDriveScrapper\//);
  assert.match(text, /Folder\//);
  assert.match(text, /Doc4\.md \(306 B\)/);

  const json = treeToJson(tree);
  assert.equal(json.children?.length, 2);
  const folder = json.children?.find((c) => c.name === "Folder");
  assert.equal(folder?.children?.length, 1);

  assert.equal(flattenTree(tree).length, 3);
});

test("formatBytes", () => {
  assert.equal(formatBytes(306), "306 B");
  assert.equal(formatBytes(9452), "9.2 KB");
});

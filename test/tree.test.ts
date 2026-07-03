/**
 * Deep unit tests for tree serialization and formatting.
 */
import { test } from "node:test";
import assert from "node:assert/strict";

import {
  treeToText,
  treeToJson,
  flattenTree,
  sortEntries,
  formatBytes,
} from "../src/tree.js";
import type { OneDriveTreeNode } from "../src/types.js";

function sampleTree(): OneDriveTreeNode {
  return {
    name: "Root",
    path: "",
    type: "folder",
    serverRelativeUrl: "/x",
    children: [
      { name: "b.txt", path: "b.txt", type: "file", serverRelativeUrl: "/x/b.txt", size: 10 },
      {
        name: "Alpha",
        path: "Alpha",
        type: "folder",
        serverRelativeUrl: "/x/Alpha",
        childCount: 1,
        children: [
          { name: "deep.md", path: "Alpha/deep.md", type: "file", serverRelativeUrl: "/x/Alpha/deep.md", size: 5, modified: new Date("2026-01-02T03:04:05Z") },
        ],
      },
      { name: "a.txt", path: "a.txt", type: "file", serverRelativeUrl: "/x/a.txt", size: 20 },
    ],
  };
}

test("sortEntries: folders before files, then case-insensitive alpha", () => {
  const items = [
    { name: "zeta.txt", path: "", type: "file" as const, serverRelativeUrl: "" },
    { name: "Beta", path: "", type: "folder" as const, serverRelativeUrl: "" },
    { name: "alpha.txt", path: "", type: "file" as const, serverRelativeUrl: "" },
    { name: "Gamma", path: "", type: "folder" as const, serverRelativeUrl: "" },
  ];
  const sorted = [...items].sort(sortEntries).map((e) => e.name);
  assert.deepEqual(sorted, ["Beta", "Gamma", "alpha.txt", "zeta.txt"]);
});

test("treeToText: folders first, tree branches, sizes optional", () => {
  const text = treeToText(sampleTree(), { showSize: true });
  const lines = text.split("\n");
  assert.equal(lines[0], "Root/");
  // Alpha (folder) must come before the files a.txt / b.txt.
  assert.equal(lines[1], "├── Alpha/");
  assert.equal(lines[2], "│   └── deep.md (5 B)");
  assert.deepEqual(lines.slice(3), ["├── a.txt (20 B)", "└── b.txt (10 B)"]);
});

test("treeToText: without showSize, no size annotations", () => {
  const text = treeToText(sampleTree());
  assert.ok(!/\(\d+ B\)/.test(text));
});

test("treeToJson: dates to ISO, childCount kept, recursion", () => {
  const json = treeToJson(sampleTree());
  assert.equal(json.type, "folder");
  const alpha = json.children?.find((c) => c.name === "Alpha");
  assert.equal(alpha?.childCount, 1);
  assert.equal(alpha?.children?.[0].modified, "2026-01-02T03:04:05.000Z");
  // A file node has no `children` key.
  const btxt = json.children?.find((c) => c.name === "b.txt");
  assert.equal(btxt && "children" in btxt, false);
});

test("flattenTree: depth-first, folders before files", () => {
  const flat = flattenTree(sampleTree()).map((e) => e.path);
  assert.deepEqual(flat, ["Alpha", "Alpha/deep.md", "a.txt", "b.txt"]);
});

test("formatBytes: units and rounding", () => {
  assert.equal(formatBytes(0), "0 B");
  assert.equal(formatBytes(1023), "1023 B");
  assert.equal(formatBytes(1024), "1 KB");
  assert.equal(formatBytes(1536), "1.5 KB");
  assert.equal(formatBytes(9452), "9.2 KB");
  assert.equal(formatBytes(1048576), "1 MB");
  assert.equal(formatBytes(1073741824), "1 GB");
  assert.equal(formatBytes(Infinity), "Infinity");
});

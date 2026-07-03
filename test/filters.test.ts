/**
 * Deep unit tests for globbing and filter compilation.
 */
import { test } from "node:test";
import assert from "node:assert/strict";

import { globToRegExp, compileFilter } from "../src/filters.js";
import type { OneDriveEntry } from "../src/types.js";

const file = (name: string, path = name): OneDriveEntry => ({
  name,
  path,
  type: "file",
  serverRelativeUrl: "/x/" + path,
});
const folder = (name: string, path = name): OneDriveEntry => ({
  name,
  path,
  type: "folder",
  serverRelativeUrl: "/x/" + path,
});

test("globToRegExp: single star does not cross '/'", () => {
  assert.ok(globToRegExp("*.md").test("a.md"));
  assert.ok(!globToRegExp("*.md").test("dir/a.md"));
});

test("globToRegExp: double star crosses '/' and matches empty", () => {
  assert.ok(globToRegExp("**/a.md").test("a.md"));
  assert.ok(globToRegExp("**/a.md").test("x/y/a.md"));
  assert.ok(globToRegExp("src/**/a.md").test("src/x/y/a.md"));
  assert.ok(globToRegExp("src/**").test("src/x/y"));
});

test("globToRegExp: '?' matches one non-separator char", () => {
  assert.ok(globToRegExp("Doc?.md").test("Doc1.md"));
  assert.ok(!globToRegExp("Doc?.md").test("Doc12.md"));
  assert.ok(!globToRegExp("a?b").test("a/b"));
});

test("globToRegExp: brace alternation", () => {
  const re = globToRegExp("*.{docx,md,pptx}");
  assert.ok(re.test("a.docx"));
  assert.ok(re.test("a.md"));
  assert.ok(re.test("a.pptx"));
  assert.ok(!re.test("a.xlsx"));
});

test("globToRegExp: escapes regex specials and is case-insensitive", () => {
  assert.ok(globToRegExp("a.b+c").test("A.B+C"));
  assert.ok(!globToRegExp("a.b").test("axb")); // '.' is literal, not 'any char'
});

test("compileFilter: empty options match everything", () => {
  const f = compileFilter();
  assert.ok(f.test(file("anything.zip")));
});

test("compileFilter: include with RegExp and function", () => {
  const byRegExp = compileFilter({ include: /\.md$/i });
  assert.ok(byRegExp.test(file("A.MD")));
  assert.ok(!byRegExp.test(file("a.txt")));

  const byFn = compileFilter({
    include: (name) => name.startsWith("Report"),
  });
  assert.ok(byFn.test(file("Report-2026.pdf")));
  assert.ok(!byFn.test(file("Summary.pdf")));
});

test("compileFilter: exclude wins over include", () => {
  const f = compileFilter({ include: "*", exclude: "*.tmp" });
  assert.ok(f.test(file("a.txt")));
  assert.ok(!f.test(file("a.tmp")));
});

test("compileFilter: mixed array of patterns and rules", () => {
  const f = compileFilter({
    include: ["*.docx", { pattern: "Reports/**", target: "path" }, /\.md$/i],
  });
  assert.ok(f.test(file("a.docx")));
  assert.ok(f.test(file("x.md", "sub/x.md")));
  assert.ok(f.test(file("q.pdf", "Reports/q.pdf")));
  assert.ok(!f.test(file("q.pdf", "Other/q.pdf")));
});

test("compileFilter: target 'path' matches full relative path", () => {
  const f = compileFilter({ include: [{ pattern: "Folder/**", target: "path" }] });
  assert.ok(f.test(file("Doc3.pptx", "Folder/Doc3.pptx")));
  assert.ok(!f.test(file("Doc1.docx", "Doc1.docx")));
});

test("compileFilter: works on folder entries too", () => {
  const f = compileFilter({ exclude: ["Temp*", "~*"] });
  assert.ok(f.test(folder("Reports")));
  assert.ok(!f.test(folder("TempCache")));
  assert.ok(!f.test(folder("~drafts")));
});

test("compileFilter: predicate receives the value and the entry", () => {
  const seen: Array<[string, string]> = [];
  const f = compileFilter({
    include: [{ pattern: (value, entry) => { seen.push([value, entry.type]); return true; }, target: "path" }],
  });
  f.test(file("a.md", "sub/a.md"));
  assert.deepEqual(seen, [["sub/a.md", "file"]]);
});

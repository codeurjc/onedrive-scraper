/**
 * CLI integration tests. Runs the CLI as a child process against the test
 * folder (headless, reusing the persisted session). Skips if there is no valid
 * session (the CLI exits non-zero when it cannot authenticate headless).
 *
 * Run with:  npm run test:live
 */
import { describe, it, before } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { readFile, rm, mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

import { TEST_FOLDER_URL, SESSION_DIR } from "./expected.js";

const execFileAsync = promisify(execFile);
const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = join(HERE, "..");
const CLI = join(REPO_ROOT, "src", "cli.ts");

/** Run the CLI via tsx. Returns { stdout, stderr }. Throws on non-zero exit. */
function runCli(args: string[]) {
  return execFileAsync(
    process.execPath,
    ["--import", "tsx", CLI, ...args, "--headless", "--session", SESSION_DIR],
    { cwd: REPO_ROOT, timeout: 120_000, maxBuffer: 20 * 1024 * 1024 }
  );
}

describe("CLI", () => {
  let available = true;
  let rootJson: Array<{ name: string; type: string }> = [];

  before(async () => {
    try {
      const { stdout } = await runCli(["list", "--url", TEST_FOLDER_URL, "--json"]);
      rootJson = JSON.parse(stdout);
    } catch (e) {
      available = false;
    }
  });

  it("list --json returns the root entries", (t) => {
    if (!available) {
      t.skip("no valid session / CLI could not authenticate headless");
      return;
    }
    const byName = Object.fromEntries(rootJson.map((e) => [e.name, e.type]));
    assert.equal(rootJson.length, 3);
    assert.equal(byName["Folder"], "folder");
    assert.equal(byName["Doc1.docx"], "file");
    assert.equal(byName["Doc2.xlsx"], "file");
  });

  it("tree --out-json writes a valid structure to disk", (t) => {
    if (!available) {
      t.skip("no valid session");
      return;
    }
    return (async () => {
      const dir = await mkdtemp(join(tmpdir(), "onedrive-cli-"));
      try {
        const out = join(dir, "tree.json");
        await runCli(["tree", "--url", TEST_FOLDER_URL, "--out-json", out]);
        const tree = JSON.parse(await readFile(out, "utf8"));
        assert.equal(tree.type, "folder");
        const names = (tree.children ?? []).map((c: { name: string }) => c.name).sort();
        assert.deepEqual(names, ["Doc1.docx", "Doc2.xlsx", "Folder"]);
        const folder = tree.children.find((c: { name: string }) => c.name === "Folder");
        const subNames = folder.children.map((c: { name: string }) => c.name).sort();
        assert.deepEqual(subNames, ["Doc3.pptx", "Doc4.md"]);
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    })();
  });

  it("tree with a file filter (--include '*.md')", (t) => {
    if (!available) {
      t.skip("no valid session");
      return;
    }
    return (async () => {
      const dir = await mkdtemp(join(tmpdir(), "onedrive-cli-"));
      try {
        const out = join(dir, "md.json");
        await runCli(["tree", "--url", TEST_FOLDER_URL, "--include", "*.md", "--out-json", out]);
        const tree = JSON.parse(await readFile(out, "utf8"));
        const rootFiles = (tree.children ?? []).filter((c: { type: string }) => c.type === "file");
        assert.equal(rootFiles.length, 0);
        const folder = tree.children.find((c: { name: string }) => c.name === "Folder");
        assert.deepEqual(
          folder.children.map((c: { name: string }) => c.name),
          ["Doc4.md"]
        );
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    })();
  });

  it("download with --include downloads only matching files", (t) => {
    if (!available) {
      t.skip("no valid session");
      return;
    }
    return (async () => {
      const dir = await mkdtemp(join(tmpdir(), "onedrive-cli-dl-"));
      try {
        await runCli(["download", "--url", TEST_FOLDER_URL, "--dest", dir, "--include", "*.md"]);
        const { stat } = await import("node:fs/promises");
        const st = await stat(join(dir, "Folder", "Doc4.md"));
        assert.equal(st.size, 306);
        // A non-matching file must not have been downloaded.
        await assert.rejects(() => stat(join(dir, "Doc1.docx")));
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    })();
  });
});

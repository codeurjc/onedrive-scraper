/**
 * Advanced live integration tests (REST strategy): stat/exists, folder filters,
 * maxDepth, path-target filters, overwrite/skip behaviour and single-file
 * download. Skips if there is no valid persisted session.
 *
 * Run with:  npm run test:live
 */
import { describe, it, before, after } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, stat as fsStat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { OneDriveScraper } from "../src/index.js";
import { TEST_FOLDER_URL, SESSION_DIR } from "./expected.js";

describe("advanced operations (rest)", () => {
  let scraper: OneDriveScraper;
  let available = true;
  let reason = "";

  before(async () => {
    scraper = new OneDriveScraper({
      folderUrl: TEST_FOLDER_URL,
      sessionDir: SESSION_DIR,
      headed: false,
      strategy: "rest",
    });
    try {
      await scraper.connect();
    } catch (e) {
      available = false;
      reason = `no valid session (${(e as Error).message})`;
    }
  });

  after(async () => {
    await scraper?.close().catch(() => {});
  });

  const guard = (t: { skip: (m?: string) => void }) => {
    if (!available) {
      t.skip(reason);
      return false;
    }
    return true;
  };

  it("stat() on the root, a file and a folder", (t) => {
    if (!guard(t)) return;
    return (async () => {
      const root = await scraper.stat("");
      assert.equal(root.type, "folder");
      assert.equal(root.name, "PruebaOneDriveScrapper");

      const rootDot = await scraper.stat("."); // '.' also means root
      assert.equal(rootDot.serverRelativeUrl, root.serverRelativeUrl);

      const doc = await scraper.stat("Doc1.docx");
      assert.equal(doc.type, "file");
      assert.equal(doc.size, 9452);

      const folder = await scraper.stat("Folder");
      assert.equal(folder.type, "folder");

      const nested = await scraper.stat("Folder/Doc4.md");
      assert.equal(nested.size, 306);
    })();
  });

  it("stat() throws for a missing entry", (t) => {
    if (!guard(t)) return;
    return assert.rejects(() => scraper.stat("does-not-exist.txt"), /not found/i);
  });

  it("exists() true/false", (t) => {
    if (!guard(t)) return;
    return (async () => {
      assert.equal(await scraper.exists("Folder/Doc4.md"), true);
      assert.equal(await scraper.exists("Folder"), true);
      assert.equal(await scraper.exists("nope.zip"), false);
      assert.equal(await scraper.exists("Folder/nope.zip"), false);
    })();
  });

  it("folder filter (exclude) prunes the subtree", (t) => {
    if (!guard(t)) return;
    return (async () => {
      const tree = await scraper.getTree("", { folders: { exclude: "Folder" } });
      const names = (tree.children ?? []).map((c) => c.name).sort();
      assert.deepEqual(names, ["Doc1.docx", "Doc2.xlsx"]); // no Folder
    })();
  });

  it("folder filter (include) keeps only matching folders", (t) => {
    if (!guard(t)) return;
    return (async () => {
      const tree = await scraper.getTree("", { folders: { include: "NoSuchFolder" } });
      // Only files remain at root; the non-matching Folder is dropped.
      const names = (tree.children ?? []).map((c) => c.name).sort();
      assert.deepEqual(names, ["Doc1.docx", "Doc2.xlsx"]);
    })();
  });

  it("maxDepth=1 lists the root but does not descend", (t) => {
    if (!guard(t)) return;
    return (async () => {
      const tree = await scraper.getTree("", { maxDepth: 1 });
      const folder = tree.children?.find((c) => c.name === "Folder");
      assert.ok(folder);
      assert.deepEqual(folder!.children, []); // present but not expanded
    })();
  });

  it("path-target file filter keeps only files under Folder/", (t) => {
    if (!guard(t)) return;
    return (async () => {
      const tree = await scraper.getTree("", {
        files: { include: [{ pattern: "Folder/**", target: "path" }] },
      });
      const rootFiles = (tree.children ?? []).filter((c) => c.type === "file");
      assert.equal(rootFiles.length, 0); // Doc1/Doc2 excluded (not under Folder/)
      const folder = tree.children?.find((c) => c.name === "Folder");
      const subNames = (folder?.children ?? []).map((c) => c.name).sort();
      assert.deepEqual(subNames, ["Doc3.pptx", "Doc4.md"]);
    })();
  });

  it("getTree from a subfolder path", (t) => {
    if (!guard(t)) return;
    return (async () => {
      const tree = await scraper.getTree("Folder");
      assert.equal(tree.name, "Folder");
      const names = (tree.children ?? []).map((c) => c.name).sort();
      assert.deepEqual(names, ["Doc3.pptx", "Doc4.md"]);
    })();
  });

  it("downloadFile() writes a single file to a path", (t) => {
    if (!guard(t)) return;
    return (async () => {
      const dir = await mkdtemp(join(tmpdir(), "onedrive-one-"));
      try {
        const dest = join(dir, "nested", "copy.xlsx");
        await scraper.downloadFile("Doc2.xlsx", dest);
        const st = await fsStat(dest);
        assert.equal(st.size, 6624);
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    })();
  });

  it("download() skips existing files unless overwrite is set", (t) => {
    if (!guard(t)) return;
    return (async () => {
      const dir = await mkdtemp(join(tmpdir(), "onedrive-ow-"));
      try {
        const opts = { destDir: dir, files: { include: "Doc1.docx" } };
        const first = await scraper.download("", opts);
        assert.equal(first.length, 1);
        assert.equal(first[0].status, "downloaded");

        const second = await scraper.download("", opts);
        assert.equal(second[0].status, "skipped");

        const third = await scraper.download("", { ...opts, overwrite: true });
        assert.equal(third[0].status, "downloaded");
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    })();
  });
});

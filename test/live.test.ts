/**
 * Live integration tests against the shared test folder.
 *
 * These require a valid persisted session (run `npm run cli -- login ...` or
 * the example once to create ./.onedrive-session). If the session is missing or
 * expired, the tests SKIP rather than fail.
 *
 * Run with:  npm run test:live
 */
import { describe, it, before, after } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { OneDriveScraper } from "../src/index.js";
import type { Strategy } from "../src/types.js";
import {
  TEST_FOLDER_URL,
  SESSION_DIR,
  EXPECTED_ROOT,
  EXPECTED_SUBFOLDER,
  EXPECTED_FILES,
} from "./expected.js";

const CONNECT_TIMEOUT = 60_000;

function makeScraper(strategy: Strategy) {
  return new OneDriveScraper({
    folderUrl: TEST_FOLDER_URL,
    sessionDir: SESSION_DIR,
    headed: false,
    strategy,
  });
}

/** A describe block that runs the same content checks for a given strategy. */
function strategySuite(strategy: Strategy, opts: { checkSizes: boolean }) {
  describe(`strategy: ${strategy}`, () => {
    let scraper: OneDriveScraper;
    let available = true;
    let reason = "";

    before(async () => {
      scraper = makeScraper(strategy);
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

    it("lists the root folder contents", (t) => {
      if (!guard(t)) return;
      return (async () => {
        const entries = await scraper.list();
        assert.equal(entries.length, EXPECTED_ROOT.length);
        for (const exp of EXPECTED_ROOT) {
          const found = entries.find((e) => e.name === exp.name);
          assert.ok(found, `missing ${exp.name}`);
          assert.equal(found!.type, exp.type, `type of ${exp.name}`);
          if (opts.checkSizes && exp.type === "file") {
            assert.equal(found!.size, exp.size, `size of ${exp.name}`);
          }
        }
      })();
    });

    it("lists the subfolder contents", (t) => {
      if (!guard(t)) return;
      return (async () => {
        const entries = await scraper.list("Folder");
        assert.equal(entries.length, EXPECTED_SUBFOLDER.length);
        for (const exp of EXPECTED_SUBFOLDER) {
          const found = entries.find((e) => e.name === exp.name);
          assert.ok(found, `missing ${exp.name}`);
          assert.equal(found!.type, exp.type);
          if (opts.checkSizes) assert.equal(found!.size, exp.size);
        }
      })();
    });

    it("builds the full tree", (t) => {
      if (!guard(t)) return;
      return (async () => {
        const tree = await scraper.getTree();
        const names = (tree.children ?? []).map((c) => c.name).sort();
        assert.deepEqual(names, ["Doc1.docx", "Doc2.xlsx", "Folder"]);
        const folder = tree.children?.find((c) => c.name === "Folder");
        const subNames = (folder?.children ?? []).map((c) => c.name).sort();
        assert.deepEqual(subNames, ["Doc3.pptx", "Doc4.md"]);
      })();
    });

    it("applies a file filter (only .md)", (t) => {
      if (!guard(t)) return;
      return (async () => {
        const tree = await scraper.getTree("", { files: { include: "*.md" } });
        const folder = tree.children?.find((c) => c.name === "Folder");
        assert.deepEqual(
          (folder?.children ?? []).map((c) => c.name),
          ["Doc4.md"]
        );
        // No non-.md files should remain at the root.
        const rootFiles = (tree.children ?? []).filter((c) => c.type === "file");
        assert.equal(rootFiles.length, 0);
      })();
    });

    it("reads a file's bytes (Folder/Doc4.md = 306 bytes)", (t) => {
      if (!guard(t)) return;
      return (async () => {
        const bytes = await scraper.readFile("Folder/Doc4.md");
        assert.equal(bytes.length, 306);
      })();
    });
  });
}

// REST verifies exact sizes; DOM verifies names/types only (sizes not scraped).
strategySuite("rest", { checkSizes: true });
strategySuite("dom", { checkSizes: false });

describe("download (rest)", () => {
  let scraper: OneDriveScraper;
  let available = true;
  let dest = "";

  before(async () => {
    scraper = makeScraper("rest");
    try {
      await scraper.connect();
      dest = await mkdtemp(join(tmpdir(), "onedrive-test-"));
    } catch {
      available = false;
    }
  });

  after(async () => {
    await scraper?.close().catch(() => {});
    if (dest) await rm(dest, { recursive: true, force: true });
  });

  it("downloads all files reconstructing the folder structure", (t) => {
    if (!available) {
      t.skip("no valid session");
      return;
    }
    return (async () => {
      const results = await scraper.download("", { destDir: dest });
      assert.equal(results.filter((r) => r.status === "error").length, 0);
      for (const exp of EXPECTED_FILES) {
        const st = await stat(join(dest, exp.path));
        assert.equal(st.size, exp.size, `downloaded size of ${exp.path}`);
      }
    })();
  });

  it("filters downloads and flattens with reconstruct=false", (t) => {
    if (!available) {
      t.skip("no valid session");
      return;
    }
    return (async () => {
      const flatDest = await mkdtemp(join(tmpdir(), "onedrive-flat-"));
      try {
        const results = await scraper.download("", {
          destDir: flatDest,
          reconstruct: false,
          files: { include: "*.md" },
        });
        const downloaded = results.filter((r) => r.status === "downloaded");
        assert.equal(downloaded.length, 1);
        const st = await stat(join(flatDest, "Doc4.md")); // flat, no Folder/
        assert.equal(st.size, 306);
      } finally {
        await rm(flatDest, { recursive: true, force: true });
      }
    })();
  });
});

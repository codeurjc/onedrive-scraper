/**
 * Basic usage example.
 *
 * Run during development with:  npm run example
 * (Uses `tsx`, so no build step is needed.)
 *
 * The first run opens a browser window — sign in (including MFA). Subsequent
 * runs reuse the persisted session in ./.onedrive-session.
 */

import { OneDriveScraper, treeToText } from "../src/index.js";

const FOLDER_URL =
  "https://urjc-my.sharepoint.com/my?id=%2Fpersonal%2Fmicael%5Fgallego%5Furjc%5Fes%2FDocuments%2FPruebaOneDriveScrapper";

async function main() {
  const scraper = new OneDriveScraper({
    folderUrl: FOLDER_URL,
    sessionDir: ".onedrive-session",
  });

  await scraper.connect();
  try {
    // 1. List the root folder (like readdir).
    console.log("Root contents:");
    for (const entry of await scraper.list()) {
      console.log(` - [${entry.type}] ${entry.name}`);
    }

    // 2. Print the full tree.
    const tree = await scraper.getTree("", { maxDepth: Infinity });
    console.log("\nTree:\n" + treeToText(tree, { showSize: true }));

    // 3. Download only Word/Markdown files, recreating the structure.
    console.log("\nDownloading .docx and .md files…");
    const results = await scraper.download("", {
      destDir: "./downloads",
      reconstruct: true,
      files: { include: ["*.docx", "*.md"] },
    });
    console.log(
      `Downloaded ${results.filter((r) => r.status === "downloaded").length} files.`
    );
  } finally {
    await scraper.close();
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

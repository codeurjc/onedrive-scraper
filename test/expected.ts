/**
 * Known contents of the shared test folder (the "well-known" folder from the
 * project spec). Used by the live integration tests.
 *
 *   PruebaOneDriveScrapper/
 *   ├── Folder/
 *   │   ├── Doc3.pptx
 *   │   └── Doc4.md
 *   ├── Doc1.docx
 *   └── Doc2.xlsx
 */

export const TEST_FOLDER_URL =
  "https://urjc-my.sharepoint.com/my?id=%2Fpersonal%2Fmicael%5Fgallego%5Furjc%5Fes%2FDocuments%2FPruebaOneDriveScrapper";

export const SESSION_DIR = ".onedrive-session";

/** Expected entries at the root, keyed by name. Sizes are exact byte lengths. */
export const EXPECTED_ROOT = [
  { name: "Folder", type: "folder" as const },
  { name: "Doc1.docx", type: "file" as const, size: 9452 },
  { name: "Doc2.xlsx", type: "file" as const, size: 6624 },
];

/** Expected entries inside `Folder/`. */
export const EXPECTED_SUBFOLDER = [
  { name: "Doc3.pptx", type: "file" as const, size: 28041 },
  { name: "Doc4.md", type: "file" as const, size: 306 },
];

/** Full expected set of files, by relative path. */
export const EXPECTED_FILES = [
  { path: "Doc1.docx", size: 9452 },
  { path: "Doc2.xlsx", size: 6624 },
  { path: "Folder/Doc3.pptx", size: 28041 },
  { path: "Folder/Doc4.md", size: 306 },
];

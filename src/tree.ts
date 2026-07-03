/**
 * Serialization of a folder tree to text and JSON.
 */

import type { OneDriveEntry, OneDriveTreeNode } from "./types.js";

/** A plain, JSON-serializable representation of a tree node. */
export interface SerializedTreeNode {
  name: string;
  path: string;
  type: "file" | "folder";
  serverRelativeUrl: string;
  size?: number;
  modified?: string;
  childCount?: number;
  children?: SerializedTreeNode[];
}

/** Convert a tree into a plain object suitable for `JSON.stringify`. */
export function treeToJson(node: OneDriveTreeNode): SerializedTreeNode {
  const out: SerializedTreeNode = {
    name: node.name,
    path: node.path,
    type: node.type,
    serverRelativeUrl: node.serverRelativeUrl,
  };
  if (node.size !== undefined) out.size = node.size;
  if (node.modified) out.modified = node.modified.toISOString();
  if (node.childCount !== undefined) out.childCount = node.childCount;
  if (node.children) out.children = node.children.map(treeToJson);
  return out;
}

export interface TreeTextOptions {
  /** Append file sizes, e.g. `Doc1.docx (12.3 KB)`. Default `false`. */
  showSize?: boolean;
}

/** Render a tree as an indented ASCII listing. */
export function treeToText(
  node: OneDriveTreeNode,
  options: TreeTextOptions = {}
): string {
  const lines: string[] = [];
  const rootLabel = node.name || node.serverRelativeUrl || ".";
  lines.push(node.type === "folder" ? `${rootLabel}/` : rootLabel);
  renderChildren(node.children ?? [], "", lines, options);
  return lines.join("\n");
}

function renderChildren(
  children: OneDriveTreeNode[],
  prefix: string,
  lines: string[],
  options: TreeTextOptions
): void {
  const sorted = [...children].sort(sortEntries);
  sorted.forEach((child, index) => {
    const last = index === sorted.length - 1;
    const branch = last ? "└── " : "├── ";
    lines.push(prefix + branch + label(child, options));
    if (child.type === "folder" && child.children) {
      renderChildren(
        child.children,
        prefix + (last ? "    " : "│   "),
        lines,
        options
      );
    }
  });
}

function label(node: OneDriveTreeNode, options: TreeTextOptions): string {
  if (node.type === "folder") return `${node.name}/`;
  if (options.showSize && node.size !== undefined) {
    return `${node.name} (${formatBytes(node.size)})`;
  }
  return node.name;
}

/** Folders before files, then alphabetical (case-insensitive). */
export function sortEntries(a: OneDriveEntry, b: OneDriveEntry): number {
  if (a.type !== b.type) return a.type === "folder" ? -1 : 1;
  return a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
}

/** Flatten a tree into a list of entries (depth-first, folders before files). */
export function flattenTree(node: OneDriveTreeNode): OneDriveEntry[] {
  const out: OneDriveEntry[] = [];
  const walk = (n: OneDriveTreeNode) => {
    for (const child of [...(n.children ?? [])].sort(sortEntries)) {
      const { children, ...entry } = child;
      out.push(entry);
      if (child.type === "folder") walk(child);
    }
  };
  walk(node);
  return out;
}

/** Human-readable byte size. */
export function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes)) return String(bytes);
  const units = ["B", "KB", "MB", "GB", "TB"];
  let value = bytes;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit++;
  }
  const rounded = unit === 0 ? value : Math.round(value * 10) / 10;
  return `${rounded} ${units[unit]}`;
}

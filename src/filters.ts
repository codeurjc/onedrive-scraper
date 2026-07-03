/**
 * Filtering of files and folders by name or full relative path.
 *
 * A pattern can be a glob string, a `RegExp`, or a predicate function.
 * Rules can target the base `name` (default) or the full relative `path`.
 */

import type {
  FilterInput,
  FilterOptions,
  FilterPattern,
  FilterRule,
  OneDriveEntry,
} from "./types.js";

/** Convert a glob string into a `RegExp`. Supports `*`, `**`, `?`, and `{a,b}`. */
export function globToRegExp(glob: string): RegExp {
  let re = "";
  for (let i = 0; i < glob.length; i++) {
    const c = glob[i];
    switch (c) {
      case "*":
        if (glob[i + 1] === "*") {
          // `**` matches across path separators.
          re += ".*";
          i++;
          // Consume an immediately following slash so `**/x` also matches `x`.
          if (glob[i + 1] === "/") i++;
        } else {
          // `*` matches anything except a path separator.
          re += "[^/]*";
        }
        break;
      case "?":
        re += "[^/]";
        break;
      case "{": {
        const end = glob.indexOf("}", i);
        if (end === -1) {
          re += "\\{";
        } else {
          const options = glob
            .slice(i + 1, end)
            .split(",")
            .map((o) => o.replace(/[.+^${}()|[\]\\]/g, "\\$&"));
          re += `(?:${options.join("|")})`;
          i = end;
        }
        break;
      }
      case ".":
      case "+":
      case "^":
      case "$":
      case "(":
      case ")":
      case "|":
      case "[":
      case "]":
      case "\\":
        re += "\\" + c;
        break;
      default:
        re += c;
    }
  }
  return new RegExp(`^${re}$`, "i");
}

function normalizeRule(input: FilterPattern | FilterRule): FilterRule {
  if (
    typeof input === "string" ||
    input instanceof RegExp ||
    typeof input === "function"
  ) {
    return { pattern: input, target: "name" };
  }
  return { target: "name", ...input };
}

function toRuleArray(input: FilterInput | undefined): FilterRule[] {
  if (input === undefined) return [];
  const arr = Array.isArray(input) ? input : [input];
  return arr.map(normalizeRule);
}

function matchRule(rule: FilterRule, entry: OneDriveEntry): boolean {
  const value = rule.target === "path" ? entry.path : entry.name;
  const { pattern } = rule;
  if (typeof pattern === "function") return pattern(value, entry);
  if (pattern instanceof RegExp) return pattern.test(value);
  return globToRegExp(pattern).test(value);
}

/**
 * A compiled filter: call `test(entry)` to decide whether an entry passes.
 * An entry passes when it matches at least one `include` rule (or there are
 * none) and matches no `exclude` rule.
 */
export interface CompiledFilter {
  test(entry: OneDriveEntry): boolean;
}

/** Compile {@link FilterOptions} into a reusable {@link CompiledFilter}. */
export function compileFilter(options?: FilterOptions): CompiledFilter {
  const include = toRuleArray(options?.include);
  const exclude = toRuleArray(options?.exclude);
  return {
    test(entry: OneDriveEntry): boolean {
      if (exclude.some((r) => matchRule(r, entry))) return false;
      if (include.length === 0) return true;
      return include.some((r) => matchRule(r, entry));
    },
  };
}

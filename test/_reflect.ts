// ── Reading the type model at run time ──────────────────────────────
//
// Several tests derive their expectations from `src/_types.ts` rather
// than transcribing them, so a field added to the model breaks the test
// that should have covered it. Types are erased at run time, so the
// source is what there is to read.
//
// This used to be three copies of
// `new RegExp('export interface ' + name + ' \\{([\\s\\S]*?)\\n\\}')`,
// which has two failure modes worth naming (#474):
//
//   1. `[\s\S]*?\n\}` stops at the first `}` in column 0. A JSDoc block
//      holding a code example, or any nested type closed at column 0,
//      truncates the interface silently — and a *shorter* field list
//      makes a derive-from-source test pass, not fail.
//   2. It cannot match `export interface Foo extends Bar {` at all, and
//      it never followed `extends`, so inherited fields were invisible.
//      That is not hypothetical: during the audit behind #439 a finding
//      was filed against `JsonReadOptions` for a missing `typeInference`
//      that was there all along, on the interface it extends.
//
// So: strip comments, count braces, follow `extends`.

import { readFileSync } from "node:fs"

/**
 * Files searched for a declaration, in order.
 *
 * `_types.ts` holds the model, but `extends` shows up in the per-format
 * option types — and following it is the whole point, so the reader has
 * to be able to see them.
 */
const SOURCES = [
  "../src/_types.ts",
  "../src/json/reader.ts",
  "../src/json/flatten.ts",
  "../src/csv/reader.ts",
]

const cache = new Map<string, string>()

function sources(): string[] {
  return SOURCES.map((path) => {
    let text = cache.get(path)
    if (text === undefined) {
      text = stripComments(readFileSync(new URL(path, import.meta.url), "utf-8"))
      cache.set(path, text)
    }
    return text
  })
}

/**
 * Remove block and line comments, keeping every newline so line-anchored
 * matching still lines up with the original.
 *
 * Done before brace counting, so a `}` inside a JSDoc example cannot end
 * an interface early, and before field matching, so a commented-out field
 * is not collected as a real one.
 */
function stripComments(text: string): string {
  return text
    .replace(/\/\*[\s\S]*?\*\//g, (m) => m.replace(/[^\n]/g, " "))
    .replace(/\/\/[^\n]*/g, (m) => " ".repeat(m.length))
}

/** The declaration head and body of one interface, comments removed. */
function declarationOf(iface: string): { extends: string[]; body: string } {
  let text: string | undefined
  let head: RegExpExecArray | null = null
  for (const candidate of sources()) {
    head = new RegExp(`export interface ${iface}\\b([^{]*)\\{`).exec(candidate)
    if (head) {
      text = candidate
      break
    }
  }
  if (!head || text === undefined) {
    throw new Error(`interface ${iface} not found — did it get renamed?`)
  }

  // Brace-count from the opening brace to its match. `[\s\S]*?\n\}` was
  // the shortcut here and it is the one that truncates.
  const open = head.index + head[0].length
  let depth = 1
  let i = open
  for (; i < text.length && depth > 0; i++) {
    if (text[i] === "{") depth++
    else if (text[i] === "}") depth--
  }
  if (depth !== 0) throw new Error(`interface ${iface} has no closing brace`)

  const heritage = head[1]!.trim()
  const bases = heritage.startsWith("extends")
    ? heritage
        .slice("extends".length)
        .split(",")
        // `extends Foo<Bar>` — the base's own name is what we can look up.
        .map((b) => b.trim().replace(/<.*$/, ""))
        .filter(Boolean)
    : []

  return { extends: bases, body: text.slice(open, i - 1) }
}

/**
 * Field names declared directly on an interface — not on anything it
 * extends, and not on nested object types inside it.
 *
 * Depth tracking is what excludes the nested ones: a field of an inline
 * `{ ... }` type sits at depth 1 or deeper and is not part of this
 * interface's own surface.
 */
export function ownFieldsOf(iface: string): string[] {
  const { body } = declarationOf(iface)
  const fields: string[] = []
  let depth = 0
  let atLineStart = true

  for (let i = 0; i < body.length; i++) {
    const ch = body[i]!
    if (ch === "{" || ch === "(" || ch === "[") depth++
    else if (ch === "}" || ch === ")" || ch === "]") depth--
    else if (ch === "\n") {
      atLineStart = true
      continue
    }

    if (!atLineStart || depth !== 0) {
      if (ch !== " ") atLineStart = false
      continue
    }
    if (ch === " ") continue

    // First non-space on a line at depth 0 — the only place a field of
    // this interface can be declared.
    atLineStart = false
    const rest = body.slice(i)
    const match = /^(?:readonly\s+)?(\w+)\s*\??\s*:/.exec(rest)
    if (match) fields.push(match[1]!)
  }

  return fields
}

/**
 * Every field an interface has, including those it inherits.
 *
 * A base that is not itself declared in `_types.ts` — a TypeScript
 * built-in, or a type imported from elsewhere — contributes nothing
 * rather than throwing, because it is not part of hucre's model and no
 * test derives expectations from it.
 */
export function fieldsOf(iface: string): string[] {
  const seen = new Set<string>()
  const out: string[] = []

  const visit = (name: string, guard: Set<string>): void => {
    if (guard.has(name)) return
    guard.add(name)

    let decl: { extends: string[]; body: string }
    try {
      decl = declarationOf(name)
    } catch {
      return
    }
    for (const base of decl.extends) visit(base, guard)
    for (const field of ownFieldsOf(name)) {
      if (!seen.has(field)) {
        seen.add(field)
        out.push(field)
      }
    }
  }

  visit(iface, new Set())
  return out
}

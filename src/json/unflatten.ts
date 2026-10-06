// ── JSON Unflatten ───────────────────────────────────────────────────
// Rebuild dot-path keyed rows into nested objects — the inverse of
// json/flatten's object flattening.
//
// `src/xml/data-writer.ts` reconstructs a tree from dot-paths too, and this
// is deliberately not shared with it. That one builds a `TreeNode`
// (attributes / `#text` / an ordered child Map, every value stringified) and
// has to decide per segment whether the segment is an attribute, mixed
// content, or an element name it must validate. The two functions agree on
// one line — split the key on "." — and disagree on the node type, the value
// type, the collision rule, and the safety model. Merging them would mean a
// tree builder parameterised by all four, which is more configuration than
// shared body, and any change to the shared part would silently be a change
// to the XML output.

import { isCellError } from "../cell-error"
import type { CellValue } from "../_types"

/** A rebuilt row: leaf cells plus the nested objects reconstructed from paths. */
export type UnflattenedRow = Record<string, unknown>

/**
 * Rebuild one flat, dot-path keyed row into a nested object.
 *
 * ```ts
 * unflattenRow({ sku: "P1", "pricing.cost": 100 })
 * // → { sku: "P1", pricing: { cost: 100 } }
 * ```
 *
 * Three rules are worth knowing before you rely on this:
 *
 * **Every dot is a separator.** `flatten` joins with "." and does not escape
 * dots that were already in a key, so `{"a.b": 1}` is indistinguishable from
 * `{a: {b: 1}}` by the time it reaches here — flatten itself already merges
 * the two, last one winning. A literal-dot key comes back nested.
 *
 * **Numeric segments stay object keys.** `{"a.0": 1}` becomes
 * `{a: {"0": 1}}`, not `{a: [1]}` — `flatten` never emits an index, because
 * it joins primitive arrays into one cell and JSON-encodes arrays of
 * objects. Guessing arrays here would invent a shape the flat form never
 * meant, and joined arrays are not recoverable either way: `"1, 2"` and the
 * literal string `"1, 2"` are the same cell.
 *
 * **Conflicts stay flat rather than overwrite.** If a path would have to
 * replace a value already placed — `{a: 1, "a.b": 2}` — the conflicting key
 * is kept verbatim instead, so no cell is ever dropped.
 */
export function unflattenRow(row: Record<string, CellValue>): UnflattenedRow {
  // Null-prototype containers all the way down. This is the whole
  // prototype-pollution defence and it is not incidental: `"__proto__"` and
  // `"constructor"` are ordinary keys in a flattened row (flatten.ts keeps
  // them, deliberately), so a path like `__proto__.polluted` would otherwise
  // reach Object.prototype's setter on the way back up. With no prototype
  // there is nothing to walk into and nothing to poison — the segment
  // becomes an own property, which is exactly what round-tripping wants.
  const root: UnflattenedRow = Object.create(null)

  for (const key of Object.keys(row)) {
    const value = row[key] as CellValue
    const dot = key.indexOf(".")
    if (dot === -1) {
      root[key] = value
      continue
    }

    const path = key.split(".")
    let node = root
    let blocked = false

    for (let i = 0; i < path.length - 1; i++) {
      const segment = path[i]!
      const existing = node[segment]
      if (existing === undefined) {
        const child: UnflattenedRow = Object.create(null)
        node[segment] = child
        node = child
      } else if (isContainer(existing)) {
        node = existing
      } else {
        // A leaf already occupies this segment. Once we start creating
        // containers we can never hit this branch, so nothing half-built is
        // left behind when we bail.
        blocked = true
        break
      }
    }

    const last = path[path.length - 1]!
    if (blocked || isContainer(node[last])) {
      root[key] = value
      continue
    }
    node[last] = value
  }

  return root
}

/** Rebuild every row of a flat table. See {@link unflattenRow}. */
export function unflattenRows(rows: Record<string, CellValue>[]): UnflattenedRow[] {
  return rows.map(unflattenRow)
}

/** A node we built, as opposed to a cell value. Cells are never plain objects. */
function isContainer(value: unknown): value is UnflattenedRow {
  return (
    typeof value === "object" && value !== null && !(value instanceof Date) && !isCellError(value)
  )
}

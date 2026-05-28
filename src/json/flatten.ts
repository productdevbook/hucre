// ── JSON Flatten ──────────────────────────────────────────────────────
// Flatten nested objects/arrays into dot-path keyed CellValue records.

import type { CellValue } from "../_types"

export interface FlattenOptions {
  /** Flatten nested objects into dot-path keys. Default: true. */
  flatten?: boolean
  /** Separator for joined primitive arrays. Default: ", ". */
  arrayJoin?: string
  /** Maximum recursion depth for flattening. Default: 32. */
  maxDepth?: number
}

/**
 * Convert an arbitrary JS value into a `CellValue` flat object.
 *
 * - Primitives (`string` / `number` / `boolean` / `null`) → single-cell
 * - `Date` → preserved as `Date`
 * - Plain objects → flattened with dot-path keys (when `flatten: true`)
 * - Arrays of primitives → joined with `arrayJoin`
 * - Arrays of objects → JSON.stringify (cannot be flattened in a tabular row)
 * - When `flatten: false`, nested objects are JSON.stringify'd into a single cell
 */
export function flattenValue(
  value: unknown,
  options: FlattenOptions = {},
): Record<string, CellValue> {
  const flatten = options.flatten ?? true
  const arrayJoin = options.arrayJoin ?? ", "
  const maxDepth = options.maxDepth ?? 32

  const out: Record<string, CellValue> = {}
  walk(value, "", out, flatten, arrayJoin, maxDepth, 0)
  return out
}

function walk(
  value: unknown,
  prefix: string,
  out: Record<string, CellValue>,
  flatten: boolean,
  arrayJoin: string,
  maxDepth: number,
  depth: number,
): void {
  if (value === null || value === undefined) {
    if (prefix) out[prefix] = null
    return
  }

  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    out[prefix] = value
    return
  }

  if (value instanceof Date) {
    out[prefix] = value
    return
  }

  if (Array.isArray(value)) {
    if (value.length === 0) {
      out[prefix] = ""
      return
    }
    const allPrimitive = value.every(
      (v) => v === null || typeof v === "string" || typeof v === "number" || typeof v === "boolean",
    )
    if (allPrimitive) {
      out[prefix] = value.map((v) => (v === null ? "" : String(v))).join(arrayJoin)
    } else {
      out[prefix] = JSON.stringify(value)
    }
    return
  }

  if (typeof value === "object") {
    const obj = value as Record<string, unknown>
    const keys = Object.keys(obj)
    if (keys.length === 0) {
      if (prefix) out[prefix] = ""
      return
    }

    // At the row level (depth 0) we always descend into top-level keys; the
    // `flatten` toggle controls whether nested objects are recursed further.
    if (depth === 0) {
      for (const key of keys) {
        walk(obj[key], key, out, flatten, arrayJoin, maxDepth, depth + 1)
      }
      return
    }

    if (!flatten || depth >= maxDepth) {
      out[prefix] = JSON.stringify(value)
      return
    }

    for (const key of keys) {
      const nextKey = prefix ? `${prefix}.${key}` : key
      walk(obj[key], nextKey, out, flatten, arrayJoin, maxDepth, depth + 1)
    }
    return
  }

  // Functions / symbols / bigint — fall back to String
  out[prefix] = String(value)
}

/**
 * Compute the union of all keys appearing in a list of flattened rows,
 * preserving first-seen order.
 */
export function collectHeaders(rows: Record<string, CellValue>[]): string[] {
  const seen = new Set<string>()
  const headers: string[] = []
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      if (!seen.has(key)) {
        seen.add(key)
        headers.push(key)
      }
    }
  }
  return headers
}

// ── JSON Flatten ──────────────────────────────────────────────────────
// Flatten nested objects/arrays into dot-path keyed CellValue records.

import type { CellValue } from "../_types"
import { reviveIsoDate } from "../_date"

export interface FlattenOptions {
  /** Flatten nested objects into dot-path keys. Default: true. */
  flatten?: boolean
  /** Separator for joined primitive arrays. Default: ", ". */
  arrayJoin?: string
  /** Maximum recursion depth for flattening. Default: 32. */
  maxDepth?: number
  /**
   * Revive ISO 8601 strings as `Date`. Default: false.
   *
   * Named to match `CsvReadOptions.typeInference`, and it accepts exactly the
   * same instants — but it does *less*, on purpose. CSV has to guess numbers
   * and booleans out of text; JSON already carries them, so a JSON string is
   * a string by the author's choice and coercing `"007"` to `7` would destroy
   * information the document was explicit about. Dates are the only type JSON
   * genuinely cannot express, so they are the only thing inferred here.
   */
  typeInference?: boolean
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
  const typeInference = options.typeInference ?? false

  // Null-prototype so keys like "__proto__" / "constructor" (which
  // JSON.parse produces as ordinary own properties) are stored as plain
  // entries instead of hitting the prototype setter — which would silently
  // drop the value (primitives) or corrupt the object (objects). Returned
  // as-is: spread, JSON.stringify, Object.keys, and for-in all work on a
  // null-prototype object; only inherited methods like .hasOwnProperty are
  // absent (use Object.hasOwn / the `in` operator instead).
  const out: Record<string, CellValue> = Object.create(null)
  walk(value, "", out, { flatten, arrayJoin, maxDepth, typeInference }, 0)
  return out
}

interface WalkConfig {
  flatten: boolean
  arrayJoin: string
  maxDepth: number
  typeInference: boolean
}

function walk(
  value: unknown,
  prefix: string,
  out: Record<string, CellValue>,
  cfg: WalkConfig,
  depth: number,
): void {
  if (value === null || value === undefined) {
    if (prefix) out[prefix] = null
    return
  }

  if (typeof value === "string") {
    // Inference runs on the leaf, so a Date nested three levels down is
    // revived just like a top-level one.
    out[prefix] = cfg.typeInference ? (reviveIsoDate(value) ?? value) : value
    return
  }

  if (typeof value === "number" || typeof value === "boolean") {
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
      out[prefix] = value.map((v) => (v === null ? "" : String(v))).join(cfg.arrayJoin)
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
        walk(obj[key], key, out, cfg, depth + 1)
      }
      return
    }

    if (!cfg.flatten || depth >= cfg.maxDepth) {
      out[prefix] = JSON.stringify(value)
      return
    }

    for (const key of keys) {
      const nextKey = prefix ? `${prefix}.${key}` : key
      walk(obj[key], nextKey, out, cfg, depth + 1)
    }
    return
  }

  // Functions / symbols / bigint — fall back to String
  out[prefix] = String(value)
}

/**
 * Revive ISO 8601 strings as `Date` throughout an already-parsed JSON value,
 * in place, and return it.
 *
 * `flattenValue` does this on the way down for readers that flatten; the
 * NDJSON stream reader yields rows verbatim unless `flattenRows` is set, so
 * it needs the same rule applied to a tree it is not otherwise touching.
 *
 * Values are only ever *replaced*, never added, so an own `__proto__` key
 * carrying a date string is overwritten as the own property JSON.parse made
 * it — the prototype setter is never reached.
 */
export function reviveDates(value: unknown, maxDepth = 32, depth = 0): unknown {
  if (typeof value === "string") return reviveIsoDate(value) ?? value
  if (value === null || typeof value !== "object" || depth >= maxDepth) return value

  if (Array.isArray(value)) {
    for (let i = 0; i < value.length; i++) {
      value[i] = reviveDates(value[i], maxDepth, depth + 1)
    }
    return value
  }

  if (value instanceof Date) return value

  const obj = value as Record<string, unknown>
  for (const key of Object.keys(obj)) {
    obj[key] = reviveDates(obj[key], maxDepth, depth + 1)
  }
  return obj
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

// ── JSON Reader ──────────────────────────────────────────────────────
// Read JSON arrays / single objects / { rowsAt: [...] } shapes
// into tabular { data, headers }.

import type { CellValue } from "../_types"
import { ParseError } from "../errors"
import { collectHeaders, flattenValue, type FlattenOptions } from "./flatten"

export interface JsonReadOptions extends FlattenOptions {
  /**
   * Dot-path to the array of rows inside the parsed JSON.
   *
   * Default behaviour when `rowsAt` is omitted:
   *   1. Top-level array → use it
   *   2. Top-level object with a single array property → use that
   *      (the common `{ products: [...] }` pattern)
   *   3. Top-level object otherwise → treat as a single row
   */
  rowsAt?: string
  /** Transform header values. */
  transformHeader?: (header: string, index: number) => string
  /** Transform each cell value. */
  transformValue?: (value: CellValue, header: string, rowIndex: number) => CellValue
  /** Maximum number of rows to return. */
  maxRows?: number
}

export interface JsonReadResult<T extends Record<string, CellValue> = Record<string, CellValue>> {
  data: T[]
  headers: string[]
}

const TEXT_DECODER = new TextDecoder("utf-8")

function toString(input: string | Uint8Array): string {
  if (typeof input === "string") return input
  return TEXT_DECODER.decode(input)
}

function getByPath(value: unknown, path: string): unknown {
  if (!path) return value
  const parts = path.split(".")
  let current: unknown = value
  for (const part of parts) {
    if (current === null || current === undefined || typeof current !== "object") {
      return undefined
    }
    current = (current as Record<string, unknown>)[part]
  }
  return current
}

/**
 * Parse a JSON string (or UTF-8 encoded Uint8Array) into tabular rows.
 *
 * Top-level shapes accepted:
 *   - Array of objects → each element is a row
 *   - Object with a single array property → that array is the rows
 *   - Object without `rowsAt` → treated as a single-row table
 *
 * Use `rowsAt` to pick a specific path inside a deeper JSON document.
 */
export function parseJson<T extends Record<string, CellValue> = Record<string, CellValue>>(
  input: string | Uint8Array,
  options?: JsonReadOptions,
): JsonReadResult<T> {
  const text = toString(input)
  if (text.trim() === "") {
    return { data: [], headers: [] }
  }

  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch (err) {
    throw new ParseError(`Invalid JSON: ${(err as Error).message}`, undefined, { cause: err })
  }

  return parseValue<T>(parsed, options)
}

/**
 * Parse an already-parsed JSON value (e.g. from `await response.json()`).
 */
export function parseValue<T extends Record<string, CellValue> = Record<string, CellValue>>(
  value: unknown,
  options?: JsonReadOptions,
): JsonReadResult<T> {
  const rows = extractRows(value, options?.rowsAt)
  return rowsToResult<T>(rows, options)
}

function extractRows(value: unknown, rowsAt: string | undefined): unknown[] {
  if (rowsAt !== undefined) {
    const target = getByPath(value, rowsAt)
    if (Array.isArray(target)) return target
    if (target === null || target === undefined) {
      throw new ParseError(`No data found at rowsAt path "${rowsAt}"`)
    }
    if (typeof target === "object") return [target]
    throw new ParseError(`Value at rowsAt path "${rowsAt}" is not an object or array`)
  }

  if (value === null || value === undefined) {
    throw new ParseError("JSON input must be an object or an array of objects")
  }

  if (Array.isArray(value)) return value

  if (typeof value === "object") {
    // Look for a single-array property
    const obj = value as Record<string, unknown>
    const entries = Object.entries(obj)
    const arrayEntries = entries.filter(([, v]) => Array.isArray(v))
    if (arrayEntries.length === 1) {
      return arrayEntries[0]![1] as unknown[]
    }
    // Fallback: treat as single row
    return [obj]
  }

  throw new ParseError("JSON input must be an object or an array of objects")
}

function rowsToResult<T extends Record<string, CellValue>>(
  rows: unknown[],
  options?: JsonReadOptions,
): JsonReadResult<T> {
  const flatOpts: FlattenOptions = {
    flatten: options?.flatten,
    arrayJoin: options?.arrayJoin,
    maxDepth: options?.maxDepth,
  }

  const limit = options?.maxRows ?? Infinity
  const flat: Record<string, CellValue>[] = []
  for (const row of rows) {
    if (flat.length >= limit) break
    if (row === null || row === undefined) {
      flat.push({})
      continue
    }
    if (typeof row !== "object" || Array.isArray(row)) {
      // Wrap a primitive/array element as a single-column "value" row
      flat.push(flattenValue({ value: row }, flatOpts))
      continue
    }
    flat.push(flattenValue(row, flatOpts))
  }

  let headers = collectHeaders(flat)
  if (options?.transformHeader) {
    headers = headers.map((h, i) => options.transformHeader!(h, i))
  }

  // Re-key data rows to match transformed headers if header transform was applied
  const originalHeaders = collectHeaders(flat)
  const headerMap = options?.transformHeader
    ? new Map(originalHeaders.map((h, i) => [h, headers[i]!]))
    : null

  const data: T[] = []
  for (let r = 0; r < flat.length; r++) {
    const src = flat[r]!
    const obj: Record<string, CellValue> = {}
    for (const origKey of originalHeaders) {
      const outKey = headerMap ? headerMap.get(origKey)! : origKey
      let val = src[origKey] ?? null
      if (options?.transformValue) {
        val = options.transformValue(val, outKey, r)
      }
      obj[outKey] = val
    }
    data.push(obj as T)
  }

  return { data, headers }
}

/**
 * Parse NDJSON / JSON Lines input — one JSON value per line.
 *
 * Blank lines are skipped. By default a malformed line throws a `ParseError`;
 * pass `onError` to collect errors instead and continue.
 */
export interface NdjsonReadOptions extends Omit<JsonReadOptions, "rowsAt"> {
  /**
   * Called when a line fails to parse. If provided, the line is skipped
   * instead of throwing.
   */
  onError?: (line: string, lineNumber: number, error: Error) => void
}

export function parseNdjson<T extends Record<string, CellValue> = Record<string, CellValue>>(
  input: string | Uint8Array,
  options?: NdjsonReadOptions,
): JsonReadResult<T> {
  const text = toString(input)
  if (text.trim() === "") return { data: [], headers: [] }

  const lines = text.split(/\r?\n/)
  const values: unknown[] = []
  for (let idx = 0; idx < lines.length; idx++) {
    const line = lines[idx]!
    if (line.trim() === "") continue
    try {
      values.push(JSON.parse(line))
    } catch (err) {
      if (options?.onError) {
        options.onError(line, idx + 1, err as Error)
        continue
      }
      throw new ParseError(
        `Invalid NDJSON on line ${idx + 1}: ${(err as Error).message}`,
        { line: idx + 1 },
        { cause: err },
      )
    }
  }

  return rowsToResult<T>(values, options)
}

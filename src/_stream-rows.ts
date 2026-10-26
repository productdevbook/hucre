// ── Row filters shared by the streaming readers ──────────────────────
// `sheet`, `range` and `maxRows` mean the same in `streamXlsxRows` and
// `streamXlsRows`, so the two apply one implementation of each.

import type { CellValue, StreamRow } from "./_types"
import { ParseError } from "./errors"
import { parseCellRef } from "./xlsx/worksheet"

export interface RangeFilter {
  startRow: number
  endRow: number
  startCol: number
  endCol: number
}

/**
 * Parse a range reference like "A1:D10" into 0-based row/col bounds.
 * Single-cell refs like "B2" are also accepted (start == end).
 */
export function parseRangeFilter(ref: string): RangeFilter {
  const parts = ref.split(":")
  if (parts.length === 0 || parts.length > 2) {
    throw new ParseError(`Invalid range reference: "${ref}"`)
  }
  const start = parseCellRef(parts[0]!)
  const end = parts.length > 1 ? parseCellRef(parts[1]!) : start
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  }
}

/**
 * Apply the range filter to a freshly-yielded row. Returns the row to emit
 * (with cells outside the column range nulled out) or `null` if the row
 * itself falls outside the row range.
 *
 * Mirrors the batch reader (`parseWorksheet`): values stay aligned to
 * their original column index, and cells outside the column window are
 * masked to `null` rather than removed, so callers can still address
 * `row.values[colIndex]` for columns inside the range.
 */
export function applyRangeFilter(row: StreamRow, range: RangeFilter): StreamRow | null {
  if (row.index < range.startRow || row.index > range.endRow) return null
  const len = Math.max(row.values.length, range.endCol + 1)
  const out: CellValue[] = Array.from({ length: len }, () => null)
  const upper = Math.min(row.values.length - 1, range.endCol)
  for (let c = range.startCol; c <= upper; c++) {
    out[c] = row.values[c] ?? null
  }
  return { index: row.index, values: out }
}

/** The sheet a stream reads: the first, else the one `sheetSpec` names by 0-based index or by name. */
export function resolveTargetSheet<T extends { name: string }>(
  allSheets: T[],
  sheetSpec?: number | string,
): T | null {
  if (sheetSpec === undefined) {
    // Default: first sheet
    return allSheets[0] ?? null
  }

  if (typeof sheetSpec === "number") {
    return sheetSpec >= 0 && sheetSpec < allSheets.length ? allSheets[sheetSpec] : null
  }

  return allSheets.find((s) => s.name === sheetSpec) ?? null
}

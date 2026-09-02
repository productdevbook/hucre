// ── Sheet Conversion Utilities ──────────────────────────────────────
// Helper functions to convert Sheet data into objects or arrays.

import type { CellValue, Sheet } from "./_types"
import { rowsToObjects } from "./_objects"

/**
 * Options for {@link sheetToObjects}.
 */
export interface SheetToObjectsOptions {
  /** 0-based row index to use as headers. Default: 0. */
  headerRow?: number
  /**
   * Skip rows where every cell is null/empty. Default: true — the same
   * default as `readObjects`, `readXlsxObjects` and `readOdsObjects`,
   * which this used to disagree with by hard-coding `false`.
   */
  skipEmptyRows?: boolean
}

/**
 * Result shape for {@link sheetToObjects}, mirroring `XlsxObjectsResult`
 * and `OdsObjectsResult`.
 */
export interface SheetObjectsResult<
  T extends Record<string, CellValue> = Record<string, CellValue>,
> {
  data: T[]
  headers: string[]
}

/**
 * Convert sheet rows to objects keyed by a header row, plus the detected
 * headers.
 *
 * A pure in-memory projection with no transform hooks. For
 * `transformHeader` / `transformValue` / `maxRows`, read through
 * `readXlsxObjects`, `readOdsObjects`, or `readObjects` instead.
 *
 * @param sheet - The sheet to convert
 * @returns `{ data, headers }` — the same shape every `*Objects` reader returns
 */
export function sheetToObjects<T extends Record<string, CellValue> = Record<string, CellValue>>(
  sheet: Sheet,
  options?: SheetToObjectsOptions,
): SheetObjectsResult<T> {
  return rowsToObjects<T>(sheet.rows, {
    headerRow: options?.headerRow ?? 0,
    skipEmptyRows: options?.skipEmptyRows ?? true,
  })
}

/**
 * Convert sheet rows to a 2D array with headers extracted from the first row.
 *
 * @param sheet - The sheet to convert
 * @returns Object with `headers` (string[]) and `data` (remaining rows)
 */
export function sheetToArrays(sheet: Sheet): {
  headers: string[]
  data: CellValue[][]
} {
  if (sheet.rows.length === 0) {
    return { headers: [], data: [] }
  }

  const headerRow = sheet.rows[0]!
  const headers = headerRow.map((h) => {
    if (h === null || h === undefined) return ""
    return String(h).trim()
  })

  const data = sheet.rows.slice(1)
  return { headers, data }
}

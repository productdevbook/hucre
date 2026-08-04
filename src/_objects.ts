// ── Shared row → object projection ──────────────────────────────────
//
// Every `*Objects` reader (`readObjects`, `readXlsxObjects`,
// `readOdsObjects`, `parseCsvObjects`, `sheetToObjects`) turns a 2D row
// array into `{ data, headers }` the same way. Keeping one implementation
// here is what makes them behave identically — three near-copies of this
// loop is how they drifted apart in the first place (#365).
//
// Internal: not exported from any entry point.

import type { CellValue, Sheet, Workbook } from "./_types"
import { ParseError } from "./errors"

/** The projection knobs shared by the object readers. */
export interface RowsToObjectsOptions {
  /** Row index to use as headers. Callers resolve the default. */
  headerRow: number
  /** Skip rows where every cell is null/undefined/"". */
  skipEmptyRows: boolean
  /** Transform header values (after String/trim normalization). */
  transformHeader?: (header: string, index: number) => string
  /** Transform each cell value. */
  transformValue?: (
    value: CellValue,
    header: string,
    rowIndex: number,
    colIndex: number,
  ) => CellValue
  /** Maximum number of data rows to return (after the header row). */
  maxRows?: number
}

/** The result shape every `*Objects` reader returns. */
export interface ObjectsResult<T extends Record<string, CellValue> = Record<string, CellValue>> {
  data: T[]
  headers: string[]
}

/**
 * Project `rows` onto objects keyed by the header row.
 *
 * `rowIndex` handed to `transformValue` is the index into `rows`, not the
 * index into the returned data — the row's position in the source is what
 * callers reach for when reporting problems.
 */
export function rowsToObjects<T extends Record<string, CellValue> = Record<string, CellValue>>(
  rows: CellValue[][],
  options: RowsToObjectsOptions,
): ObjectsResult<T> {
  const {
    headerRow: headerRowIdx,
    skipEmptyRows,
    transformHeader,
    transformValue,
    maxRows,
  } = options

  if (headerRowIdx < 0 || rows.length <= headerRowIdx) {
    return { data: [], headers: [] }
  }

  let headers = rows[headerRowIdx]!.map((h) => {
    if (h === null || h === undefined) return ""
    return String(h).trim()
  })

  if (transformHeader) {
    headers = headers.map((h, i) => transformHeader(h, i))
  }

  const data: T[] = []
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    if (maxRows !== undefined && data.length >= maxRows) break
    const row = rows[i]!

    if (skipEmptyRows && row.every((v) => v === null || v === undefined || v === "")) {
      continue
    }

    const obj: Record<string, CellValue> = {}
    for (let j = 0; j < headers.length; j++) {
      let val: CellValue = j < row.length ? (row[j] ?? null) : null
      if (transformValue) {
        val = transformValue(val, headers[j]!, i, j)
      }
      obj[headers[j]!] = val
    }
    data.push(obj as T)
  }

  return { data, headers }
}

/**
 * Resolve a `sheet` selector (0-based index or name) against a workbook,
 * throwing the same typed errors from every reader that accepts one.
 */
export function selectSheet(workbook: Workbook, selector: number | string): Sheet {
  if (workbook.sheets.length === 0) {
    throw new ParseError("Workbook contains no sheets")
  }

  const sheet =
    typeof selector === "number"
      ? workbook.sheets[selector]
      : workbook.sheets.find((s) => s.name === selector)

  if (!sheet) {
    throw new ParseError(
      typeof selector === "number"
        ? `Sheet index ${selector} out of range (workbook has ${workbook.sheets.length} sheet(s))`
        : `Sheet "${selector}" not found`,
    )
  }

  return sheet
}

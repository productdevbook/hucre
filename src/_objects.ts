// ── Shared row → object projection ──────────────────────────────────
//
// Every `*Objects` reader (`readObjects`, `readXlsxObjects`,
// `readOdsObjects`, `parseCsvObjects`, `sheetToObjects`) turns a 2D row
// array into `{ data, headers }` the same way. Keeping one implementation
// here is what makes them behave identically — three near-copies of this
// loop is how they drifted apart in the first place (#365).
//
// Internal: not exported from any entry point.

import type { Cell, CellValue, Sheet, Workbook } from "./_types"
import type { ValueSheet } from "./_sheet-values"
import { sheetGrid, type GridProjectionOptions } from "./_sheet-grid"
import { assertGridSize } from "./_grid"
import { ParseError } from "./errors"

/** The projection knobs shared by the object readers. */
export interface RowsToObjectsOptions extends GridProjectionOptions {
  /** 0-based header row. Default: 0. */
  headerRow?: number
  /** Skip rows where every effective cell is null/"". Default: true. */
  skipEmptyRows?: boolean
  /** Transform header values (after String/trim normalization). */
  transformHeader?: (header: string, index: number) => string
  /** Transform each cell value. */
  transformValue?: (
    value: CellValue,
    header: string,
    rowIndex: number,
    colIndex: number,
  ) => CellValue
  /** Maximum returned data rows, after empty-row filtering; not a parse limit. */
  maxRows?: number
}

/** The result shape every `*Objects` reader returns. */
export interface ObjectsResult<T extends Record<string, CellValue> = Record<string, CellValue>> {
  data: T[]
  headers: string[]
}

/**
 * Make a header list usable as object keys, by renaming only what would
 * otherwise be lost.
 *
 * A header that is unique — **including a single blank one** — is left
 * exactly as it is. An empty header keying `""` is a settled contract
 * across every `*Objects` reader and is pinned by several tests; the loss
 * was never the blank key, it was the *second* column sharing it.
 *
 * So only repeats are renamed. A repeated name gets `_2`, `_3`; a
 * repeated blank gets `column<N>` for its 1-based position, since `_2`
 * alone would read as nothing at all.
 */
export function disambiguate(headers: string[]): string[] {
  const used = new Set<string>()
  return headers.map((header, index) => {
    if (!used.has(header)) {
      used.add(header)
      return header
    }
    let name = header === "" ? `column${index + 1}` : `${header}_2`
    let ordinal = 2
    while (used.has(name)) {
      ordinal++
      name = header === "" ? `column${index + 1}_${ordinal}` : `${header}_${ordinal}`
    }
    used.add(name)
    return name
  })
}

/** Lazy table projection; output bounds count selected rows, not sparse gaps. */
export function projectTable(
  sheet: ValueSheet<Partial<Cell>>,
  options: RowsToObjectsOptions = {},
): { headers: string[]; rows: Generator<[number, CellValue[]]> } {
  const grid = sheetGrid(sheet)
  const headerRow = options.headerRow ?? 0
  if (!Number.isInteger(headerRow) || headerRow < 0 || headerRow >= grid.height) {
    return { headers: [], rows: (function* (): Generator<[number, CellValue[]]> {})() }
  }
  const skipEmptyRows = options.skipEmptyRows ?? true
  const available = grid.height - headerRow - 1
  const dataRows =
    options.maxRows !== undefined && options.maxRows < available
      ? Math.max(0, Math.ceil(options.maxRows))
      : available
  // Keeping blanks makes the output rectangle known up front. A per-row
  // check alone would build up to the limit before rejecting a sparse
  // corner, and would already have invoked user transforms on that work.
  assertGridSize(skipEmptyRows ? 1 : dataRows + 1, grid.width, options.maxTotalCells)
  const headers = disambiguate(
    grid.row(headerRow).map((value, col) => {
      const header = value === null ? "" : String(value).trim()
      return options.transformHeader ? options.transformHeader(header, col) : header
    }),
  )
  function* rows(): Generator<[number, CellValue[]]> {
    let count = 0
    for (const index of grid.indexes(headerRow + 1, skipEmptyRows)) {
      if (options.maxRows !== undefined && count >= options.maxRows) break
      // Check before even one more row is padded. Sparse reads bypass the
      // reader's box limit, but an object export still allocates its output.
      assertGridSize(count + 2, headers.length, options.maxTotalCells)
      yield [index, grid.row(index, headers.length)]
      count++
    }
  }
  return { headers, rows: rows() }
}

/** Prototype-looking labels are ordinary own properties, never setters. */
export function rowObject(headers: string[], values: CellValue[]): Record<string, CellValue> {
  return Object.fromEntries(headers.map((header, col) => [header, values[col] ?? null]))
}

/** Physical row indexes survive sparse selection and empty-row filtering. */
export function rowsToObjects<T extends Record<string, CellValue> = Record<string, CellValue>>(
  source: CellValue[][] | ValueSheet<Partial<Cell>>,
  options: RowsToObjectsOptions = {},
): ObjectsResult<T> {
  const { headers, rows } = projectTable(Array.isArray(source) ? { rows: source } : source, options)
  const data: T[] = []
  for (const [index, values] of rows) {
    if (options.transformValue) {
      for (let col = 0; col < values.length; col++) {
        values[col] = options.transformValue(values[col]!, headers[col]!, index, col)
      }
    }
    data.push(rowObject(headers, values) as T)
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

/**
 * The union of every key appearing in a list of objects, in first-seen
 * order.
 *
 * The object writers used to take their column set from `data[0]` alone,
 * so any key absent from the first record — an optional field, a column
 * that appears halfway through an export — was dropped along with the
 * rows that only had it. See #439.
 */
export function collectHeaders(rows: Record<string, unknown>[]): string[] {
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

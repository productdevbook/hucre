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

  // Two columns sharing a header used to collapse into one key: the later
  // column overwrote the earlier and its values were gone. That is not
  // exotic input — it is what a real spreadsheet looks like when someone
  // repeated a label or left two spacer columns. See #439 §AG.
  //
  // Only repeats are renamed; the first column with a given name keeps it,
  // which is what a caller reading `data[0].name` expects.
  headers = disambiguate(headers)

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

/**
 * The union of every key appearing in a list of objects, in first-seen
 * order.
 *
 * The object writers used to take their column set from `data[0]` alone,
 * so any key absent from the first record — an optional field, a column
 * that appears halfway through an export — was dropped along with the
 * rows that only had it. See #439.
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

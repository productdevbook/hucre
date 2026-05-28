// ── XLSX Object Shorthand ──────────────────────────────────────────────
// Header-row-based read/write helpers that mirror parseCsvObjects ergonomics.

import type { CellValue, ReadInput, ReadOptions, WriteOutput } from "../_types"
import { ParseError } from "../errors"
import { readXlsx } from "./reader"
import { writeXlsx } from "./writer"

/**
 * Options for {@link readXlsxObjects}.
 */
export interface XlsxObjectsReadOptions extends Omit<ReadOptions, "sheets"> {
  /** Sheet to read from. Index (0-based) or sheet name. Default: 0. */
  sheet?: number | string
  /** 0-based row index to use as headers. Default: 0. */
  headerRow?: number
  /** Skip rows where every cell is null/empty. Default: true. */
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
  /** Maximum number of data rows to return (after the header row). */
  maxRows?: number
}

/**
 * Result shape for {@link readXlsxObjects}, mirroring `parseCsvObjects`.
 */
export interface XlsxObjectsResult<
  T extends Record<string, CellValue> = Record<string, CellValue>,
> {
  data: T[]
  headers: string[]
}

/**
 * Read an XLSX file and return its rows as an array of objects keyed by
 * header values, plus the detected headers.
 *
 * Mirror of `parseCsvObjects` for binary spreadsheets — the typical
 * `readXlsx → wb.sheets[0].rows[0] as headers, slice(1) as data` boilerplate
 * collapses to a single call.
 */
export async function readXlsxObjects<
  T extends Record<string, CellValue> = Record<string, CellValue>,
>(input: ReadInput, options?: XlsxObjectsReadOptions): Promise<XlsxObjectsResult<T>> {
  const headerRowIdx = options?.headerRow ?? 0
  const skipEmpty = options?.skipEmptyRows ?? true
  const sheetSelector = options?.sheet ?? 0

  const {
    sheet: _sheet,
    headerRow: _hr,
    skipEmptyRows: _se,
    transformHeader,
    transformValue,
    maxRows,
    ...readOpts
  } = options ?? {}

  const wb = await readXlsx(input, readOpts)
  if (wb.sheets.length === 0) {
    throw new ParseError("Workbook contains no sheets")
  }

  const sheet =
    typeof sheetSelector === "number"
      ? wb.sheets[sheetSelector]
      : wb.sheets.find((s) => s.name === sheetSelector)

  if (!sheet) {
    throw new ParseError(
      typeof sheetSelector === "number"
        ? `Sheet index ${sheetSelector} out of range (workbook has ${wb.sheets.length} sheet(s))`
        : `Sheet "${sheetSelector}" not found`,
    )
  }

  if (sheet.rows.length <= headerRowIdx) {
    return { data: [], headers: [] }
  }

  const headerRow = sheet.rows[headerRowIdx]!
  let headers = headerRow.map((h) => {
    if (h === null || h === undefined) return ""
    return String(h).trim()
  })

  if (transformHeader) {
    headers = headers.map((h, i) => transformHeader(h, i))
  }

  const data: T[] = []
  for (let i = headerRowIdx + 1; i < sheet.rows.length; i++) {
    if (maxRows !== undefined && data.length >= maxRows) break
    const row = sheet.rows[i]!

    if (skipEmpty && row.every((v) => v === null || v === undefined || v === "")) {
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
 * Options for {@link writeXlsxObjects}.
 */
export interface XlsxObjectsWriteOptions {
  /** Output sheet name. Default: "Sheet1". */
  sheetName?: string
  /**
   * Explicit column order. If omitted, headers are derived from the keys
   * of the first object (in insertion order).
   */
  headers?: string[]
  /** Write a header row as the first row. Default: true. */
  writeHeaders?: boolean
}

/**
 * Write an array of objects to an XLSX file. Header order is taken from
 * `options.headers`, falling back to the keys of the first object.
 *
 * Symmetric counterpart to {@link readXlsxObjects} and `writeCsvObjects`.
 */
export async function writeXlsxObjects(
  data: Record<string, CellValue>[],
  options?: XlsxObjectsWriteOptions,
): Promise<WriteOutput> {
  const sheetName = options?.sheetName ?? "Sheet1"
  const writeHeaders = options?.writeHeaders ?? true

  const headers = options?.headers ?? (data.length > 0 ? Object.keys(data[0]!) : [])

  const rows: CellValue[][] = []
  if (writeHeaders) {
    rows.push(headers.slice())
  }
  for (const obj of data) {
    rows.push(
      headers.map((key) => {
        const val = obj[key]
        return val === undefined ? null : val
      }),
    )
  }

  return await writeXlsx({ sheets: [{ name: sheetName, rows }] })
}

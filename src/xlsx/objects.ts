// ── XLSX Object Shorthand ──────────────────────────────────────────────
// Header-row-based read/write helpers that mirror parseCsvObjects ergonomics.

import type { CellValue, ReadInput, XlsxReadOptions, WriteOutput } from "../_types"
import {
  collectHeaders,
  rowsToObjects,
  selectSheet,
  type RowsToObjectsOptions,
  type ObjectsResult,
} from "../_objects"
import { readXlsx } from "./reader"
import { writeXlsx } from "./writer"

/**
 * Options for {@link readXlsxObjects}.
 */
export interface XlsxObjectsReadOptions
  extends Omit<XlsxReadOptions, "sheets">, RowsToObjectsOptions {
  /** Sheet to read from. Index (0-based) or sheet name. Default: 0. */
  sheet?: number | string
}

/**
 * Result shape for {@link readXlsxObjects}, mirroring `parseCsvObjects`.
 */
export type XlsxObjectsResult<T extends Record<string, CellValue> = Record<string, CellValue>> =
  ObjectsResult<T>

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
  const {
    sheet: sheetSelector = 0,
    headerRow = 0,
    skipEmptyRows = true,
    transformHeader,
    transformValue,
    maxRows,
    ...readOpts
  } = options ?? {}

  const wb = await readXlsx(input, readOpts)
  const sheet = selectSheet(wb, sheetSelector)

  return rowsToObjects<T>(sheet, {
    headerRow,
    skipEmptyRows,
    transformHeader,
    transformValue,
    maxRows,
    maxTotalCells: readOpts.maxTotalCells,
  })
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
  writeHeader?: boolean
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
  const writeHeader = options?.writeHeader ?? true

  const headers = options?.headers ?? collectHeaders(data)

  const rows: CellValue[][] = []
  if (writeHeader) {
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

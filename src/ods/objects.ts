// ── ODS Object Shorthand ──────────────────────────────────────────────
// Header-row-based read/write helpers that mirror parseCsvObjects ergonomics.

import type { CellValue, ReadInput, OdsReadOptions, WriteOutput } from "../_types"
import { collectHeaders, rowsToObjects, selectSheet } from "../_objects"
import { readOds } from "./reader"
import { writeOds } from "./writer"

/**
 * Options for {@link readOdsObjects}.
 */
export interface OdsObjectsReadOptions extends Omit<OdsReadOptions, "sheets"> {
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
 * Result shape for {@link readOdsObjects}, mirroring `parseCsvObjects`.
 */
export interface OdsObjectsResult<T extends Record<string, CellValue> = Record<string, CellValue>> {
  data: T[]
  headers: string[]
}

/**
 * Read an ODS file and return its rows as an array of objects keyed by
 * header values, plus the detected headers.
 */
export async function readOdsObjects<
  T extends Record<string, CellValue> = Record<string, CellValue>,
>(input: ReadInput, options?: OdsObjectsReadOptions): Promise<OdsObjectsResult<T>> {
  const {
    sheet: sheetSelector = 0,
    headerRow = 0,
    skipEmptyRows = true,
    transformHeader,
    transformValue,
    maxRows,
    ...readOpts
  } = options ?? {}

  const wb = await readOds(input, readOpts)
  const sheet = selectSheet(wb, sheetSelector)

  return rowsToObjects<T>(sheet.rows, {
    headerRow,
    skipEmptyRows,
    transformHeader,
    transformValue,
    maxRows,
  })
}

/**
 * Options for {@link writeOdsObjects}.
 */
export interface OdsObjectsWriteOptions {
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
 * Write an array of objects to an ODS file.
 */
export async function writeOdsObjects(
  data: Record<string, CellValue>[],
  options?: OdsObjectsWriteOptions,
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

  return await writeOds({ sheets: [{ name: sheetName, rows }] })
}

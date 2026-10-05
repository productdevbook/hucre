// ── ODS Object Shorthand ──────────────────────────────────────────────
// Header-row-based read/write helpers that mirror parseCsvObjects ergonomics.

import type { CellValue, ReadInput, OdsReadOptions, WriteOutput } from "../_types"
import {
  collectHeaders,
  rowsToObjects,
  selectSheet,
  type RowsToObjectsOptions,
  type ObjectsResult,
} from "../_objects"
import { readOds } from "./reader"
import { writeOds } from "./writer"

/**
 * Options for {@link readOdsObjects}.
 */
export interface OdsObjectsReadOptions
  extends Omit<OdsReadOptions, "sheets">, RowsToObjectsOptions {
  /** Sheet to read from. Index (0-based) or sheet name. Default: 0. */
  sheet?: number | string
}

/**
 * Result shape for {@link readOdsObjects}, mirroring `parseCsvObjects`.
 */
export type OdsObjectsResult<T extends Record<string, CellValue> = Record<string, CellValue>> =
  ObjectsResult<T>

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

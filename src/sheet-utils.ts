// Sheet conversion and object readers use one model projection.
import type { CellValue, Sheet } from "./_types"
import {
  rowsToObjects,
  projectTable,
  type RowsToObjectsOptions,
  type ObjectsResult,
} from "./_objects"

/** In-memory projection has the same knobs as the object readers. */
export interface SheetToObjectsOptions extends RowsToObjectsOptions {}
export type SheetObjectsResult<T extends Record<string, CellValue> = Record<string, CellValue>> =
  ObjectsResult<T>

/** Project effective values, skipping blank rows by default. */
export function sheetToObjects<T extends Record<string, CellValue> = Record<string, CellValue>>(
  sheet: Sheet,
  options?: SheetToObjectsOptions,
): SheetObjectsResult<T> {
  return rowsToObjects<T>(sheet, options)
}

export interface SheetToArraysOptions extends Pick<
  RowsToObjectsOptions,
  "headerRow" | "skipEmptyRows" | "maxRows" | "maxTotalCells"
> {}

/** Split a table into headers and rectangular values; keep blank rows by default. */
export function sheetToArrays(
  sheet: Sheet,
  options?: SheetToArraysOptions,
): {
  headers: string[]
  data: CellValue[][]
} {
  const { headers, rows } = projectTable(sheet, {
    ...options,
    skipEmptyRows: options?.skipEmptyRows ?? false,
  })
  return { headers, data: Array.from(rows, ([, values]) => values) }
}

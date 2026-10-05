import { isCellError } from "../cell-error"
import type { Sheet } from "../_types"

import { sheetToArrays, type SheetToArraysOptions } from "../sheet-utils"
import { rowObject } from "../_objects"

export interface JsonExportOptions extends SheetToArraysOptions {
  /**
   * Output format. Default: "objects"
   *
   * Only `"objects"` has a reader: `parseJson` reads it back, and
   * `jsonToWorkbook` reads a whole workbook of it. `"arrays"` and
   * `"columns"` are **write-only** presentation shapes — handoffs to a
   * charting library or a dataframe, which is the reason they exist — and
   * `parseJson` will not reconstruct a table from either. Export in
   * `"objects"` if the JSON has to come back into hucre.
   */
  format?: "objects" | "arrays" | "columns"
  /** Pretty print. Default: false */
  pretty?: boolean
}

/**
 * Custom replacer for JSON.stringify that converts Date objects to ISO strings.
 */
function dateReplacer(_key: string, value: unknown): unknown {
  if (value instanceof Date) {
    return value.toISOString()
  }
  if (isCellError(value)) return value.error
  return value
}

/**
 * Export a sheet as a JSON string.
 *
 * Formats:
 * - `"objects"` (default): `[{Name:"Widget", Price:9.99}, ...]` — readable
 *   back with `parseJson`
 * - `"arrays"`: `{headers:["Name","Price"], data:[["Widget",9.99], ...]}` —
 *   write-only
 * - `"columns"`: `{Name:["Widget","Gadget"], Price:[9.99,24.5]}` (columnar) —
 *   write-only
 */
export function toJson(sheet: Sheet, options?: JsonExportOptions): string {
  const { headers, data } = sheetToArrays(sheet, options)
  const format = options?.format ?? "objects"
  const output =
    format === "arrays"
      ? { headers, data }
      : format === "columns"
        ? Object.fromEntries(headers.map((header, col) => [header, data.map((row) => row[col]!)]))
        : data.map((row) => rowObject(headers, row))
  return JSON.stringify(output, dateReplacer, options?.pretty ? 2 : undefined)
}

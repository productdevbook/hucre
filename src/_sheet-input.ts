// Authoring has one row source before any format-specific consumer runs.
// XLSX used data first, ODS used rows first, and text writers ignored data.
// Resolve here so values, links, header offsets and explicit overrides
// cannot drift between worksheet, table, pivot and presentation output.
import type { CellInput, CellStyle, ColumnDef, SheetInput } from "./_types"
import { collectHeaders } from "./_objects"
import { splitInlineCells } from "./_inline-cells"
import { InvalidArgumentError } from "./errors"

/** A column numFmt is a default; the explicit style's format wins. */
export function columnCellStyle(col: ColumnDef): CellStyle | undefined {
  return col.numFmt && !col.style?.numFmt ? { ...col.style, numFmt: col.numFmt } : col.style
}

/** Generated headers keep physical positions and omit the data-only numFmt. */
export function columnHeaders(columns?: readonly ColumnDef[]): CellInput[] | null {
  if (!columns?.some((col) => col.header !== undefined)) return null
  return columns.map((col) => {
    const value = col.header ?? col.key ?? ""
    return col.numFmt && !col.style?.numFmt ? { value, style: col.style ?? {} } : value
  })
}

/** One object projection: empty keys are valid, inherited fields are not data. */
export function objectRow(
  item: Record<string, unknown>,
  columns: readonly ColumnDef[],
): CellInput[] {
  return columns.map((col) => {
    const key = col.key ?? col.header
    return key !== undefined && Object.hasOwn(item, key) ? ((item[key] as CellInput) ?? null) : null
  })
}

/** Resolve object rows and lift rich cells without mutating the input. */
export function normalizeSheetInput<T extends SheetInput>(
  sheet: T,
  checkSize?: (height: number, width: number) => void,
): T {
  if (sheet.data === undefined) return splitInlineCells(sheet)
  if (sheet.rows !== undefined) {
    throw new InvalidArgumentError(
      `Sheet "${sheet.name}" supplies both rows and data; choose one row source.`,
    )
  }
  const columns: ColumnDef[] =
    sheet.columns ?? collectHeaders(sheet.data).map((key) => ({ key, header: key }))
  const headers = columnHeaders(columns)
  // Sparse object records become a full column rectangle. Bound that
  // expansion before looking up values or constructing any data rows.
  checkSize?.(sheet.data.length + (headers ? 1 : 0), columns.length)
  const rows: CellInput[][] = []
  if (headers) rows.push(headers)
  for (const record of sheet.data) {
    rows.push(objectRow(record, columns))
  }
  return splitInlineCells({ ...sheet, data: undefined, columns, rows })
}

// Authoring has one row source before any format-specific consumer runs.
// XLSX used data first, ODS used rows first, and text writers ignored data.
// Resolve here so values, links, header offsets and explicit overrides
// cannot drift between worksheet, table, pivot and presentation output.
import type { CellInput, CellStyle, ColumnDef, SheetInput } from "./_types"
import { collectHeaders } from "./_objects"
import { splitInlineCells } from "./_inline-cells"
import { isHyperlinkValue } from "./xlsx/hyperlink"
import { InvalidArgumentError } from "./errors"

/** A column numFmt is a default; the explicit style's format wins. */
export function columnCellStyle(col: ColumnDef): CellStyle | undefined {
  return col.numFmt && !col.style?.numFmt ? { ...col.style, numFmt: col.numFmt } : col.style
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
  const hasHeaders = columns.some((col) => col.header !== undefined)
  // Sparse object records become a full column rectangle. Bound that
  // expansion before looking up values or constructing any data rows.
  checkSize?.(sheet.data.length + Number(hasHeaders), columns.length)
  const rows: CellInput[][] = []
  if (hasHeaders) {
    rows.push(
      columns.map((col) => {
        const value = col.header ?? col.key ?? ""
        // The generated header receives the column's explicit style, but
        // not its numeric data format. An empty style suppresses the default.
        return col.numFmt && !col.style?.numFmt ? { value, style: col.style ?? {} } : value
      }),
    )
  }
  for (const record of sheet.data) {
    rows.push(
      columns.map((col) => {
        const key = col.key ?? col.header
        const value = key !== undefined && Object.hasOwn(record, key) ? (record[key] ?? null) : null
        if (!isHyperlinkValue(value)) return value
        const hyperlink = value.hyperlink.startsWith("#")
          ? { target: "", location: value.hyperlink.slice(1), display: value.text }
          : { target: value.hyperlink, display: value.text }
        return {
          value: value.text,
          hyperlink:
            value.tooltip === undefined ? hyperlink : { ...hyperlink, tooltip: value.tooltip },
        }
      }),
    )
  }
  return splitInlineCells({ ...sheet, data: undefined, columns, rows })
}

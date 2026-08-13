// ── Ranges Derived From A Write Sheet ────────────────────────────────
// The two places a range is *computed* rather than given: the reserved
// print defined names, and a table's extent.
//
// Both write paths need them — `writeXlsx` (authoring) and `saveXlsx`
// (round-trip) — and both used to carry their own copy. Two copies of one
// question is how they drift: `applyPrintDefinedNames` in the reader
// depends on the writer deriving these names from `pageSetup` and nowhere
// else (#407), which is only true while every writer derives them the
// same way. One implementation, two callers.

import type { NamedRange, TableDefinition, WriteSheet } from "../_types"
import { colToLetter } from "./worksheet-writer"

/**
 * Build the full list of named ranges, merging user-defined ranges with
 * auto-generated _xlnm.Print_Area and _xlnm.Print_Titles from sheet pageSetup.
 */
export function buildNamedRanges(sheets: WriteSheet[], userRanges?: NamedRange[]): NamedRange[] {
  const result: NamedRange[] = userRanges ? [...userRanges] : []

  for (const sheet of sheets) {
    const ps = sheet.pageSetup
    if (!ps) continue

    // Print area → _xlnm.Print_Area
    if (ps.printArea) {
      result.push({
        name: "_xlnm.Print_Area",
        range: `${sheet.name}!${ps.printArea}`,
        scope: sheet.name,
      })
    }

    // Print titles (repeat rows and/or columns)
    const titleParts: string[] = []
    if (ps.printTitlesRow) {
      titleParts.push(`${sheet.name}!${ps.printTitlesRow}`)
    }
    if (ps.printTitlesColumn) {
      titleParts.push(`${sheet.name}!${ps.printTitlesColumn}`)
    }
    if (titleParts.length > 0) {
      result.push({
        name: "_xlnm.Print_Titles",
        range: titleParts.join(","),
        scope: sheet.name,
      })
    }
  }

  return result
}

/**
 * Auto-calculate table range from sheet data and table column count.
 * Assumes header row is row 1 and data fills remaining rows.
 */
export function computeTableRange(table: TableDefinition, sheet: WriteSheet): string {
  const colCount = table.columns.length
  let rowCount = 0

  if (sheet.rows) {
    rowCount = sheet.rows.length
  } else if (sheet.data) {
    // Object data: data rows + 1 header row (if columns have headers)
    const hasHeaders = sheet.columns?.some((c) => c.header)
    rowCount = sheet.data.length + (hasHeaders ? 1 : 0)
  }

  // Add total row if requested
  if (table.showTotalRow) {
    rowCount += 1
  }

  // Minimum: 1 header row + 0 data rows = 1 row
  if (rowCount < 1) rowCount = 1

  const startCol = colToLetter(0)
  const endCol = colToLetter(colCount - 1)
  return `${startCol}1:${endCol}${rowCount}`
}

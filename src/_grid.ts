import type { CellValue } from "./_types"
import { ParseError } from "./errors"
import { MAX_TOTAL_CELLS } from "./limits"

/** Check the rectangle before allocating any slots, including deferred rows. */
export function assertGridSize(
  height: number,
  width: number,
  limit: number = MAX_TOTAL_CELLS,
): void {
  const total = height * width
  if (total > limit) {
    throw new ParseError(
      `Sheet spans ${height} rows x ${width} columns (${total} cells), over the ${limit} limit. ` +
        "Raise `maxTotalCells` if the sheet really is this large.",
    )
  }
}

/**
 * Pad every row to the widest, in place, so `rows` is the dense rectangle
 * `Sheet.rows` promises: `rows[r][c]` is safe without a guard on either
 * index. `readXlsx` has always done this; the ODS, CSV and HTML paths
 * returned `[]` for an empty row and left a short line short, so the same
 * sheet read three ways had three shapes.
 */
export function padToRectangle(
  rows: CellValue[][],
  limit: number = MAX_TOTAL_CELLS,
): CellValue[][] {
  let width = 0
  for (const row of rows) if (row.length > width) width = row.length
  // A wide first row followed by short ones can be small until padding.
  // Checking only the slots already parsed misses this amplification.
  assertGridSize(rows.length, width, limit)
  for (const row of rows) while (row.length < width) row.push(null)
  return rows
}

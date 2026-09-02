import type { CellValue } from "./_types"

/**
 * Pad every row to the widest, in place, so `rows` is the dense rectangle
 * `Sheet.rows` promises: `rows[r][c]` is safe without a guard on either
 * index. `readXlsx` has always done this; the ODS, CSV and HTML paths
 * returned `[]` for an empty row and left a short line short, so the same
 * sheet read three ways had three shapes.
 */
export function padToRectangle(rows: CellValue[][]): CellValue[][] {
  let width = 0
  for (const row of rows) if (row.length > width) width = row.length
  for (const row of rows) while (row.length < width) row.push(null)
  return rows
}

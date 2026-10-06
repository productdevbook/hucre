// Projection reads the same value/cache overlays as editing. Keep sparse
// row selection separate from rectangle allocation: a far-away value is
// one record when blanks are skipped, not a million intermediate arrays.
import type { Cell, CellValue } from "./_types"
import { cellEntries, getCell } from "./cell-store"
import { effectiveValue, visitValues, type ValueSheet } from "./_sheet-values"
import { assertGridSize } from "./_grid"

export interface GridProjectionOptions {
  /** Maximum projected cells, including headers. Default: 20,000,000. */
  maxTotalCells?: number
}

interface ValueGrid {
  height: number
  width: number
  value(row: number, col: number): CellValue
  row(index: number, columns?: number): CellValue[]
  indexes(start: number, skipEmptyRows: boolean): Generator<number>
}

export function sheetGrid(sheet: ValueSheet<Partial<Cell>>): ValueGrid {
  let height = sheet.rows.length
  let width = 0
  for (const row of sheet.rows) width = Math.max(width, row?.length ?? 0)
  for (const [row, col] of cellEntries(sheet.cells)) {
    height = Math.max(height, row + 1)
    width = Math.max(width, col + 1)
  }
  const value = (row: number, col: number): CellValue =>
    effectiveValue(getCell(sheet.cells, row, col), sheet.rows[row]?.[col] ?? null)
  return {
    height,
    width,
    value,
    row(index: number, columns = width): CellValue[] {
      return Array.from({ length: columns }, (_, col) => value(index, col))
    },
    *indexes(start: number, skipEmptyRows: boolean): Generator<number> {
      if (skipEmptyRows) {
        const populated = new Set<number>()
        visitValues(sheet, (value, row) => {
          if (row >= start && value !== null && value !== "") populated.add(row)
        })
        yield* [...populated].sort((a, b) => a - b)
      } else {
        for (let row = start; row < height; row++) yield row
      }
    },
  }
}

/** Dense presentation formats pay for the box; check before creating it. */
export function denseValues(sheet: ValueSheet<Partial<Cell>>, limit?: number): CellValue[][] {
  const grid = sheetGrid(sheet)
  assertGridSize(grid.height, grid.width, limit)
  // Reader grids are already rectangular. Value-only exports can borrow
  // them instead of duplicating every slot merely to inspect the model.
  if (!sheet.cells?.size) {
    let rectangular = true
    for (const row of sheet.rows) {
      if (!row || row.length !== grid.width) {
        rectangular = false
        break
      }
    }
    if (rectangular) return sheet.rows
  }
  return Array.from({ length: grid.height }, (_, row) => grid.row(row))
}

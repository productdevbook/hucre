// ── Sheet Operations ────────────────────────────────────────────────
// In-memory row/column manipulation utilities for Sheet objects.

import type { Sheet, MergeRange, RowDef, Workbook, Cell, CellValue } from "./_types"
import { parseCellRef } from "./xlsx/worksheet"
import { toRange, type RangeLike } from "./cell-utils"
import { rangeRef } from "./xlsx/worksheet-writer"
import { cloneCellStyle } from "./_style"
import { InvalidArgumentError } from "./errors"
import { shiftFormula, shiftRangeRef, type RefShift } from "./_refs"

// ── Range Helpers ────────────────────────────────────────────────────

/**
 * Parse a range string like "A1:D10" into 0-based coordinates.
 */
function parseRange(range: string): MergeRange {
  const parts = range.split(":")
  const start = parseCellRef(parts[0])
  const end = parts.length > 1 ? parseCellRef(parts[1]) : start
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  }
}

/**
 * Build a range string from 0-based coordinates.
 */
function buildRange(r: MergeRange): string {
  return rangeRef(r.startRow, r.startCol, r.endRow, r.endCol)
}

/**
 * Shift row references in a range string by a given delta.
 * Only rows >= threshold are shifted.
 */
function shiftRangeRows(range: string, threshold: number, delta: number): string {
  const r = parseRange(range)
  if (r.startRow >= threshold) r.startRow += delta
  if (r.endRow >= threshold) r.endRow += delta
  return buildRange(r)
}

/**
 * Shift column references in a range string by a given delta.
 * Only columns >= threshold are shifted.
 */
function shiftRangeCols(range: string, threshold: number, delta: number): string {
  const r = parseRange(range)
  if (r.startCol >= threshold) r.startCol += delta
  if (r.endCol >= threshold) r.endCol += delta
  return buildRange(r)
}

// ── Reference maintenance ────────────────────────────────────────────

/**
 * Move everything on a sheet that *names* a position rather than
 * occupying one: formula text, the formulas inside data validations and
 * conditional rules, sparkline ranges, page breaks and text-box anchors.
 *
 * The cells themselves are moved by the caller; this is the other half,
 * and it was missing entirely. A formula below an insertion still pointed
 * where its arguments used to be, which is the one thing "insert a row"
 * is supposed to take care of. See #439 §D.
 *
 * Not covered, and not coverable from here: `Workbook.namedRanges`,
 * defined names, external-link caches and pivot caches all live on the
 * workbook, and these operations are handed a `Sheet`.
 */
function shiftReferences(sheet: Sheet, shift: RefShift): void {
  if (sheet.cells) {
    for (const cell of sheet.cells.values()) {
      if (cell.formula) cell.formula = shiftFormula(cell.formula, shift)
      if (cell.formulaRef) cell.formulaRef = shiftFormula(cell.formulaRef, shift)
    }
  }

  if (sheet.dataValidations) {
    for (const dv of sheet.dataValidations) {
      if (dv.formula1) dv.formula1 = shiftFormula(dv.formula1, shift)
      if (dv.formula2) dv.formula2 = shiftFormula(dv.formula2, shift)
    }
  }

  if (sheet.conditionalRules) {
    for (const rule of sheet.conditionalRules) {
      if (Array.isArray(rule.formula)) {
        rule.formula = rule.formula.map((f) => shiftFormula(f, shift))
      } else if (typeof rule.formula === "string") {
        rule.formula = shiftFormula(rule.formula, shift)
      }
    }
  }

  if (sheet.sparklines) {
    // A sparkline whose whole source was deleted has nothing left to draw.
    sheet.sparklines = sheet.sparklines.filter((sparkline) => {
      const dataRange = shiftRangeRef(sparkline.dataRange, shift)
      if (dataRange === undefined) return false
      sparkline.dataRange = dataRange
      const location = shiftRangeRef(sparkline.location, shift)
      if (location === undefined) return false
      sparkline.location = location
      return true
    })
  }

  const breaks = shift.axis === "row" ? sheet.rowBreaks : sheet.colBreaks
  if (breaks) {
    const moved: number[] = []
    for (const at of breaks) {
      const next = shiftIndex(at, shift)
      if (next !== null && !moved.includes(next)) moved.push(next)
    }
    moved.sort((a, b) => a - b)
    if (shift.axis === "row") sheet.rowBreaks = moved
    else sheet.colBreaks = moved
  }

  if (sheet.textBoxes) {
    for (const box of sheet.textBoxes) {
      shiftAnchor(box.anchor, shift)
    }
  }
}

/** One index, or `null` when the row or column it names was deleted. */
function shiftIndex(value: number, shift: RefShift): number | null {
  if (shift.delta > 0) return value >= shift.at ? value + shift.delta : value
  const removed = -shift.delta
  if (value >= shift.at + removed) return value - removed
  if (value >= shift.at) return null
  return value
}

/** Move a drawing anchor's corners, clamping one that lands in the gap. */
function shiftAnchor(
  anchor: { from: { row: number; col: number }; to?: { row: number; col: number } },
  shift: RefShift,
): void {
  const key = shift.axis === "row" ? "row" : "col"
  const from = shiftIndex(anchor.from[key], shift)
  anchor.from[key] = from ?? shift.at
  if (anchor.to) {
    const to = shiftIndex(anchor.to[key], shift)
    anchor.to[key] = to ?? shift.at
  }
}

// ── Row Width Helper ─────────────────────────────────────────────────

function getRowWidth(sheet: Sheet): number {
  let width = 0
  for (const row of sheet.rows) {
    if (row.length > width) width = row.length
  }
  if (sheet.columns && sheet.columns.length > width) {
    width = sheet.columns.length
  }
  return width
}

function makeEmptyRow(width: number): null[] {
  const row: null[] = []
  for (let i = 0; i < width; i++) row.push(null)
  return row
}

// ── Insert Rows ──────────────────────────────────────────────────────

/**
 * Insert rows at the given position (0-based), shifting existing rows down.
 * Updates merge ranges, data validations, conditional rules, auto filter,
 * images, and cells Map keys.
 */
export function insertRows(sheet: Sheet, rowIndex: number, count: number): void {
  if (count <= 0) return

  const width = getRowWidth(sheet)
  const newRows: null[][] = []
  for (let i = 0; i < count; i++) {
    newRows.push(makeEmptyRow(width))
  }

  // Insert into rows array
  sheet.rows.splice(rowIndex, 0, ...newRows)

  // Update cells Map
  if (sheet.cells && sheet.cells.size > 0) {
    const updated = new Map<string, import("./_types").Cell>()
    for (const [key, cell] of sheet.cells) {
      const [rowStr, colStr] = key.split(",")
      const row = Number(rowStr)
      const col = Number(colStr)
      if (row >= rowIndex) {
        updated.set(`${row + count},${col}`, cell)
      } else {
        updated.set(key, cell)
      }
    }
    sheet.cells = updated
  }

  // Update merge ranges
  if (sheet.merges) {
    for (const merge of sheet.merges) {
      if (merge.startRow >= rowIndex) {
        merge.startRow += count
        merge.endRow += count
      } else if (merge.endRow >= rowIndex) {
        // Merge starts before insertion but ends at or after — expand it
        merge.endRow += count
      }
    }
  }

  // Update data validations
  if (sheet.dataValidations) {
    for (const dv of sheet.dataValidations) {
      dv.range = shiftRangeRows(dv.range, rowIndex, count)
    }
  }

  // Update conditional rules
  if (sheet.conditionalRules) {
    for (const rule of sheet.conditionalRules) {
      rule.range = shiftRangeRows(rule.range, rowIndex, count)
    }
  }

  // Update auto filter
  if (sheet.autoFilter) {
    sheet.autoFilter.range = shiftRangeRows(sheet.autoFilter.range, rowIndex, count)
  }

  // Update image anchors
  if (sheet.images) {
    for (const img of sheet.images) {
      if (img.anchor.from.row >= rowIndex) {
        img.anchor.from.row += count
      }
      if (img.anchor.to && img.anchor.to.row >= rowIndex) {
        img.anchor.to.row += count
      }
    }
  }

  // Update row defs
  if (sheet.rowDefs && sheet.rowDefs.size > 0) {
    const updated = new Map<number, RowDef>()
    for (const [row, def] of sheet.rowDefs) {
      if (row >= rowIndex) {
        updated.set(row + count, def)
      } else {
        updated.set(row, def)
      }
    }
    sheet.rowDefs = updated
  }

  // Update table ranges
  if (sheet.tables) {
    for (const table of sheet.tables) {
      if (table.range) {
        table.range = shiftRangeRows(table.range, rowIndex, count)
      }
    }
  }

  shiftReferences(sheet, { axis: "row", at: rowIndex, delta: count })
}

// ── Delete Rows ──────────────────────────────────────────────────────

/**
 * Delete rows starting at the given position (0-based), shifting remaining rows up.
 * Removes merges fully within deleted range. Adjusts merges that partially overlap.
 */
export function deleteRows(sheet: Sheet, rowIndex: number, count: number): void {
  if (count <= 0) return

  const deleteEnd = rowIndex + count // exclusive

  // Remove rows from array
  sheet.rows.splice(rowIndex, count)

  // Update cells Map
  if (sheet.cells && sheet.cells.size > 0) {
    const updated = new Map<string, import("./_types").Cell>()
    for (const [key, cell] of sheet.cells) {
      const [rowStr, colStr] = key.split(",")
      const row = Number(rowStr)
      const col = Number(colStr)
      if (row >= rowIndex && row < deleteEnd) {
        // Cell is in deleted range — remove it
        continue
      } else if (row >= deleteEnd) {
        updated.set(`${row - count},${col}`, cell)
      } else {
        updated.set(key, cell)
      }
    }
    sheet.cells = updated
  }

  // Update merge ranges
  if (sheet.merges) {
    sheet.merges = sheet.merges.filter((merge) => {
      // Fully within deleted range — remove
      if (merge.startRow >= rowIndex && merge.endRow < deleteEnd) {
        return false
      }
      return true
    })

    for (const merge of sheet.merges) {
      if (merge.startRow >= deleteEnd) {
        // Entirely below deleted range — shift up
        merge.startRow -= count
        merge.endRow -= count
      } else if (merge.endRow >= deleteEnd) {
        // Partially overlapping: starts before or at deletion, ends after
        if (merge.startRow >= rowIndex) {
          // Starts within deleted range — clamp start to rowIndex
          merge.startRow = rowIndex
          merge.endRow -= count
        } else {
          // Starts before deleted range — shrink end
          merge.endRow -= count
        }
      } else if (merge.endRow >= rowIndex) {
        // Ends within deleted range but starts before — clamp end
        merge.endRow = rowIndex - 1
      }
    }

    // Drop merges that no longer merge anything. `start > end` is
    // incoherent; `start === end` on both axes is a one-cell merge, which
    // is not what any spreadsheet means by the word — Excel writes
    // `<mergeCell ref="B3:B3"/>` for nothing, and a shrunk range should
    // disappear the way a fully-deleted one already does.
    sheet.merges = sheet.merges.filter(
      (m) =>
        m.startRow <= m.endRow &&
        m.startCol <= m.endCol &&
        !(m.startRow === m.endRow && m.startCol === m.endCol),
    )
  }

  // Update data validations
  if (sheet.dataValidations) {
    sheet.dataValidations = sheet.dataValidations.filter((dv) => {
      const r = parseRange(dv.range)
      // Remove if fully within deleted range
      if (r.startRow >= rowIndex && r.endRow < deleteEnd) return false
      return true
    })
    for (const dv of sheet.dataValidations) {
      dv.range = shiftDeletedRangeRows(dv.range, rowIndex, count)
    }
  }

  // Update conditional rules
  if (sheet.conditionalRules) {
    sheet.conditionalRules = sheet.conditionalRules.filter((rule) => {
      const r = parseRange(rule.range)
      if (r.startRow >= rowIndex && r.endRow < deleteEnd) return false
      return true
    })
    for (const rule of sheet.conditionalRules) {
      rule.range = shiftDeletedRangeRows(rule.range, rowIndex, count)
    }
  }

  // Update auto filter
  if (sheet.autoFilter) {
    const r = parseRange(sheet.autoFilter.range)
    if (r.startRow >= rowIndex && r.endRow < deleteEnd) {
      sheet.autoFilter = undefined
    } else {
      sheet.autoFilter.range = shiftDeletedRangeRows(sheet.autoFilter.range, rowIndex, count)
    }
  }

  // Update image anchors
  if (sheet.images) {
    sheet.images = sheet.images.filter((img) => {
      // Remove images whose anchor starts in deleted range
      return !(img.anchor.from.row >= rowIndex && img.anchor.from.row < deleteEnd)
    })
    for (const img of sheet.images) {
      if (img.anchor.from.row >= deleteEnd) {
        img.anchor.from.row -= count
      }
      if (img.anchor.to && img.anchor.to.row >= deleteEnd) {
        img.anchor.to.row -= count
      }
    }
  }

  // Update row defs
  if (sheet.rowDefs && sheet.rowDefs.size > 0) {
    const updated = new Map<number, RowDef>()
    for (const [row, def] of sheet.rowDefs) {
      if (row >= rowIndex && row < deleteEnd) {
        continue // deleted
      } else if (row >= deleteEnd) {
        updated.set(row - count, def)
      } else {
        updated.set(row, def)
      }
    }
    sheet.rowDefs = updated
  }

  // Update table ranges
  if (sheet.tables) {
    sheet.tables = sheet.tables.filter((table) => {
      if (!table.range) return true
      const r = parseRange(table.range)
      return !(r.startRow >= rowIndex && r.endRow < deleteEnd)
    })
    for (const table of sheet.tables) {
      if (table.range) {
        table.range = shiftDeletedRangeRows(table.range, rowIndex, count)
      }
    }
  }

  shiftReferences(sheet, { axis: "row", at: rowIndex, delta: -count })
}

/**
 * Shift row references in a range string after deletion.
 * Rows >= deleteEnd shift up by count.
 * Rows within [rowIndex, deleteEnd) are clamped.
 */
function shiftDeletedRangeRows(range: string, rowIndex: number, count: number): string {
  const deleteEnd = rowIndex + count
  const r = parseRange(range)

  if (r.startRow >= deleteEnd) {
    r.startRow -= count
  } else if (r.startRow >= rowIndex) {
    r.startRow = rowIndex
  }

  if (r.endRow >= deleteEnd) {
    r.endRow -= count
  } else if (r.endRow >= rowIndex) {
    r.endRow = rowIndex > 0 ? rowIndex - 1 : 0
  }

  return buildRange(r)
}

// ── Insert Columns ───────────────────────────────────────────────────

/**
 * Insert columns at the given position (0-based), shifting existing columns right.
 * Updates merge ranges, data validations, conditional rules, auto filter,
 * images, column defs, and cells Map keys.
 */
export function insertColumns(sheet: Sheet, colIndex: number, count: number): void {
  if (count <= 0) return

  const nulls: null[] = makeEmptyRow(count)

  // Insert nulls into each row
  for (const row of sheet.rows) {
    // Extend row if it's shorter than colIndex
    while (row.length < colIndex) row.push(null)
    row.splice(colIndex, 0, ...nulls)
  }

  // Update column defs
  if (sheet.columns) {
    const newCols: import("./_types").ColumnDef[] = []
    for (let i = 0; i < count; i++) newCols.push({})
    // Ensure columns array is long enough
    while (sheet.columns.length < colIndex) sheet.columns.push({})
    sheet.columns.splice(colIndex, 0, ...newCols)
  }

  // Update cells Map
  if (sheet.cells && sheet.cells.size > 0) {
    const updated = new Map<string, import("./_types").Cell>()
    for (const [key, cell] of sheet.cells) {
      const [rowStr, colStr] = key.split(",")
      const row = Number(rowStr)
      const col = Number(colStr)
      if (col >= colIndex) {
        updated.set(`${row},${col + count}`, cell)
      } else {
        updated.set(key, cell)
      }
    }
    sheet.cells = updated
  }

  // Update merge ranges
  if (sheet.merges) {
    for (const merge of sheet.merges) {
      if (merge.startCol >= colIndex) {
        merge.startCol += count
        merge.endCol += count
      } else if (merge.endCol >= colIndex) {
        merge.endCol += count
      }
    }
  }

  // Update data validations
  if (sheet.dataValidations) {
    for (const dv of sheet.dataValidations) {
      dv.range = shiftRangeCols(dv.range, colIndex, count)
    }
  }

  // Update conditional rules
  if (sheet.conditionalRules) {
    for (const rule of sheet.conditionalRules) {
      rule.range = shiftRangeCols(rule.range, colIndex, count)
    }
  }

  // Update auto filter
  if (sheet.autoFilter) {
    sheet.autoFilter.range = shiftRangeCols(sheet.autoFilter.range, colIndex, count)
  }

  // Update image anchors
  if (sheet.images) {
    for (const img of sheet.images) {
      if (img.anchor.from.col >= colIndex) {
        img.anchor.from.col += count
      }
      if (img.anchor.to && img.anchor.to.col >= colIndex) {
        img.anchor.to.col += count
      }
    }
  }

  // Update table ranges
  if (sheet.tables) {
    for (const table of sheet.tables) {
      if (table.range) {
        table.range = shiftRangeCols(table.range, colIndex, count)
      }
    }
  }

  shiftReferences(sheet, { axis: "col", at: colIndex, delta: count })
}

// ── Delete Columns ───────────────────────────────────────────────────

/**
 * Delete columns starting at the given position (0-based), shifting remaining columns left.
 * Removes merges fully within deleted range. Adjusts merges that partially overlap.
 */
export function deleteColumns(sheet: Sheet, colIndex: number, count: number): void {
  if (count <= 0) return

  const deleteEnd = colIndex + count // exclusive

  // Remove columns from each row
  for (const row of sheet.rows) {
    if (colIndex < row.length) {
      row.splice(colIndex, Math.min(count, row.length - colIndex))
    }
  }

  // Update column defs
  if (sheet.columns) {
    if (colIndex < sheet.columns.length) {
      sheet.columns.splice(colIndex, Math.min(count, sheet.columns.length - colIndex))
    }
  }

  // Update cells Map
  if (sheet.cells && sheet.cells.size > 0) {
    const updated = new Map<string, import("./_types").Cell>()
    for (const [key, cell] of sheet.cells) {
      const [rowStr, colStr] = key.split(",")
      const row = Number(rowStr)
      const col = Number(colStr)
      if (col >= colIndex && col < deleteEnd) {
        continue // deleted
      } else if (col >= deleteEnd) {
        updated.set(`${row},${col - count}`, cell)
      } else {
        updated.set(key, cell)
      }
    }
    sheet.cells = updated
  }

  // Update merge ranges
  if (sheet.merges) {
    sheet.merges = sheet.merges.filter((merge) => {
      if (merge.startCol >= colIndex && merge.endCol < deleteEnd) {
        return false
      }
      return true
    })

    for (const merge of sheet.merges) {
      if (merge.startCol >= deleteEnd) {
        merge.startCol -= count
        merge.endCol -= count
      } else if (merge.endCol >= deleteEnd) {
        if (merge.startCol >= colIndex) {
          merge.startCol = colIndex
          merge.endCol -= count
        } else {
          merge.endCol -= count
        }
      } else if (merge.endCol >= colIndex) {
        merge.endCol = colIndex - 1
      }
    }

    // Same rule as deleteRows: a range shrunk to one cell is no longer a
    // merge.
    sheet.merges = sheet.merges.filter(
      (m) =>
        m.startRow <= m.endRow &&
        m.startCol <= m.endCol &&
        !(m.startRow === m.endRow && m.startCol === m.endCol),
    )
  }

  // Update data validations
  if (sheet.dataValidations) {
    sheet.dataValidations = sheet.dataValidations.filter((dv) => {
      const r = parseRange(dv.range)
      if (r.startCol >= colIndex && r.endCol < deleteEnd) return false
      return true
    })
    for (const dv of sheet.dataValidations) {
      dv.range = shiftDeletedRangeCols(dv.range, colIndex, count)
    }
  }

  // Update conditional rules
  if (sheet.conditionalRules) {
    sheet.conditionalRules = sheet.conditionalRules.filter((rule) => {
      const r = parseRange(rule.range)
      if (r.startCol >= colIndex && r.endCol < deleteEnd) return false
      return true
    })
    for (const rule of sheet.conditionalRules) {
      rule.range = shiftDeletedRangeCols(rule.range, colIndex, count)
    }
  }

  // Update auto filter
  if (sheet.autoFilter) {
    const r = parseRange(sheet.autoFilter.range)
    if (r.startCol >= colIndex && r.endCol < deleteEnd) {
      sheet.autoFilter = undefined
    } else {
      sheet.autoFilter.range = shiftDeletedRangeCols(sheet.autoFilter.range, colIndex, count)
    }
  }

  // Update image anchors
  if (sheet.images) {
    sheet.images = sheet.images.filter((img) => {
      return !(img.anchor.from.col >= colIndex && img.anchor.from.col < deleteEnd)
    })
    for (const img of sheet.images) {
      if (img.anchor.from.col >= deleteEnd) {
        img.anchor.from.col -= count
      }
      if (img.anchor.to && img.anchor.to.col >= deleteEnd) {
        img.anchor.to.col -= count
      }
    }
  }

  // Update table ranges
  if (sheet.tables) {
    sheet.tables = sheet.tables.filter((table) => {
      if (!table.range) return true
      const r = parseRange(table.range)
      return !(r.startCol >= colIndex && r.endCol < deleteEnd)
    })
    for (const table of sheet.tables) {
      if (table.range) {
        table.range = shiftDeletedRangeCols(table.range, colIndex, count)
      }
    }
  }

  shiftReferences(sheet, { axis: "col", at: colIndex, delta: -count })
}

/**
 * Shift column references in a range string after deletion.
 */
function shiftDeletedRangeCols(range: string, colIndex: number, count: number): string {
  const deleteEnd = colIndex + count
  const r = parseRange(range)

  if (r.startCol >= deleteEnd) {
    r.startCol -= count
  } else if (r.startCol >= colIndex) {
    r.startCol = colIndex
  }

  if (r.endCol >= deleteEnd) {
    r.endCol -= count
  } else if (r.endCol >= colIndex) {
    r.endCol = colIndex > 0 ? colIndex - 1 : 0
  }

  return buildRange(r)
}

// ── Move Rows ────────────────────────────────────────────────────────

/**
 * Move rows from one position to another.
 * Extracts `count` rows starting at `fromIndex` and inserts them at `toIndex`.
 * `toIndex` is the target position in the original (pre-move) coordinate space.
 */
export function moveRows(sheet: Sheet, fromIndex: number, count: number, toIndex: number): void {
  if (count <= 0 || fromIndex === toIndex) return

  // Extract rows
  const extractedRows = sheet.rows.splice(fromIndex, count)

  // Extract cells for moved rows
  const extractedCells = new Map<string, import("./_types").Cell>()
  if (sheet.cells) {
    for (const [key, cell] of sheet.cells) {
      const [rowStr] = key.split(",")
      const row = Number(rowStr)
      if (row >= fromIndex && row < fromIndex + count) {
        extractedCells.set(key, cell)
        sheet.cells.delete(key)
      }
    }
  }

  // Extract row defs for moved rows
  const extractedRowDefs = new Map<number, RowDef>()
  if (sheet.rowDefs) {
    for (const [row, def] of sheet.rowDefs) {
      if (row >= fromIndex && row < fromIndex + count) {
        extractedRowDefs.set(row, def)
        sheet.rowDefs.delete(row)
      }
    }
  }

  // After removing from source, adjust target index
  let adjustedTo = toIndex
  if (toIndex > fromIndex) {
    adjustedTo = toIndex - count
  }

  // Re-insert rows at adjusted position
  sheet.rows.splice(adjustedTo, 0, ...extractedRows)

  // Rebuild cells Map: shift all remaining cells, then re-add extracted
  if (sheet.cells || extractedCells.size > 0) {
    const newCells = new Map<string, import("./_types").Cell>()

    // Re-key all existing cells based on their new row positions
    if (sheet.cells) {
      // After splice-out and splice-in, we need to rebuild row indices
      // The simplest approach: re-scan all rows and assign cell positions
      // based on the final row layout.
      // But cells map may have entries that don't correspond to rows array.
      // Safer approach: rebuild by tracking position changes.

      // After removal: rows above fromIndex stay, rows at fromIndex+ shift up by count
      // After insertion: rows at adjustedTo+ shift down by count
      for (const [key, cell] of sheet.cells) {
        const [rowStr, colStr] = key.split(",")
        let row = Number(rowStr)
        const col = Number(colStr)

        // After removal of [fromIndex, fromIndex+count):
        if (row >= fromIndex) {
          row -= count
        }
        // After insertion at adjustedTo:
        if (row >= adjustedTo) {
          row += count
        }

        newCells.set(`${row},${col}`, cell)
      }
    }

    // Re-add extracted cells at their new positions
    for (const [key, cell] of extractedCells) {
      const [rowStr, colStr] = key.split(",")
      const originalRow = Number(rowStr)
      const col = Number(colStr)
      const offset = originalRow - fromIndex
      const newRow = adjustedTo + offset
      newCells.set(`${newRow},${col}`, cell)
    }

    sheet.cells = newCells.size > 0 ? newCells : undefined
  }

  // Rebuild row defs
  if (sheet.rowDefs || extractedRowDefs.size > 0) {
    const newRowDefs = new Map<number, RowDef>()

    if (sheet.rowDefs) {
      for (const [row, def] of sheet.rowDefs) {
        let newRow = row
        if (newRow >= fromIndex) {
          newRow -= count
        }
        if (newRow >= adjustedTo) {
          newRow += count
        }
        newRowDefs.set(newRow, def)
      }
    }

    for (const [row, def] of extractedRowDefs) {
      const offset = row - fromIndex
      newRowDefs.set(adjustedTo + offset, def)
    }

    sheet.rowDefs = newRowDefs.size > 0 ? newRowDefs : undefined
  }
}

// ── Hide Rows ────────────────────────────────────────────────────────

/**
 * Set row hidden state for `count` rows starting at `startRow`.
 * @param hidden - Default true. Pass false to unhide.
 */
export function hideRows(
  sheet: Sheet,
  startRow: number,
  count: number,
  hidden: boolean = true,
): void {
  if (!sheet.rowDefs) sheet.rowDefs = new Map()
  for (let i = startRow; i < startRow + count; i++) {
    const existing = sheet.rowDefs.get(i) || {}
    existing.hidden = hidden
    sheet.rowDefs.set(i, existing)
  }
}

// ── Hide Columns ─────────────────────────────────────────────────────

/**
 * Set column hidden state for `count` columns starting at `startCol`.
 * @param hidden - Default true. Pass false to unhide.
 */
export function hideColumns(
  sheet: Sheet,
  startCol: number,
  count: number,
  hidden: boolean = true,
): void {
  if (!sheet.columns) sheet.columns = []
  // Ensure columns array is large enough
  while (sheet.columns.length <= startCol + count - 1) {
    sheet.columns.push({})
  }
  for (let i = startCol; i < startCol + count; i++) {
    sheet.columns[i].hidden = hidden
  }
}

// ── Group Rows ───────────────────────────────────────────────────────

/**
 * Set outline level for rows in range [startRow, endRow] (inclusive, 0-based).
 * @param level - Outline level (default 1). Set to 0 to ungroup.
 */
export function groupRows(sheet: Sheet, startRow: number, endRow: number, level: number = 1): void {
  if (!sheet.rowDefs) sheet.rowDefs = new Map()
  for (let i = startRow; i <= endRow; i++) {
    const existing = sheet.rowDefs.get(i) || {}
    existing.outlineLevel = level
    sheet.rowDefs.set(i, existing)
  }
}

// ── Deep Clone Helpers ────────────────────────────────────────────────

function cloneCell(cell: Cell): Cell {
  const result: Cell = { value: cell.value, type: cell.type }
  if (cell.style) result.style = cloneCellStyle(cell.style)
  if (cell.checkbox !== undefined) result.checkbox = cell.checkbox
  if (cell.formula !== undefined) result.formula = cell.formula
  if (cell.formulaResult !== undefined) result.formulaResult = cell.formulaResult
  // The formula's *shape*, not just its text. Dropping these turned a
  // shared-formula slave cell — `{ formula: "", formulaType: "shared",
  // formulaSharedIndex: 3 }` — into a plain `{ formula: "" }`, which the
  // writer then emitted as an empty `<f/>`. An array formula lost its
  // spill range, and a dynamic array lost its metadata link (#423).
  if (cell.formulaType !== undefined) result.formulaType = cell.formulaType
  if (cell.formulaSharedIndex !== undefined) result.formulaSharedIndex = cell.formulaSharedIndex
  if (cell.formulaRef !== undefined) result.formulaRef = cell.formulaRef
  if (cell.formulaDynamic !== undefined) result.formulaDynamic = cell.formulaDynamic
  if (cell.richText)
    result.richText = cell.richText.map((r) => ({
      text: r.text,
      font: r.font
        ? { ...r.font, color: r.font.color ? { ...r.font.color } : undefined }
        : undefined,
    }))
  if (cell.hyperlink) result.hyperlink = { ...cell.hyperlink }
  if (cell.comment) {
    result.comment = { text: cell.comment.text, author: cell.comment.author }
    if (cell.comment.richText) {
      result.comment.richText = cell.comment.richText.map((r) => ({
        text: r.text,
        font: r.font
          ? { ...r.font, color: r.font.color ? { ...r.font.color } : undefined }
          : undefined,
      }))
    }
  }
  return result
}

// ── Clone Sheet ─────────────────────────────────────────────────────

/**
 * Deep clone a sheet (all data, styles, merges, validations, etc.).
 * The cloned sheet gets a new name.
 */
export function cloneSheet(sheet: Sheet, newName: string): Sheet {
  // Deep copy rows
  const rows = sheet.rows.map((row) => [...row])

  const cloned: Sheet = { name: newName, rows }
  if (sheet.kind !== undefined) cloned.kind = sheet.kind

  // Deep copy cells Map
  if (sheet.cells && sheet.cells.size > 0) {
    const cells = new Map<string, Cell>()
    for (const [key, cell] of sheet.cells) {
      cells.set(key, cloneCell(cell))
    }
    cloned.cells = cells
  }

  // Deep copy columns
  if (sheet.columns) {
    cloned.columns = sheet.columns.map((col) => ({
      ...col,
      style: col.style ? cloneCellStyle(col.style) : undefined,
    }))
  }

  // Deep copy rowDefs
  if (sheet.rowDefs && sheet.rowDefs.size > 0) {
    const rowDefs = new Map<number, RowDef>()
    for (const [key, def] of sheet.rowDefs) {
      rowDefs.set(key, { ...def })
    }
    cloned.rowDefs = rowDefs
  }

  // Deep copy merges
  if (sheet.merges) {
    cloned.merges = sheet.merges.map((m) => ({ ...m }))
  }

  // Deep copy data validations
  if (sheet.dataValidations) {
    cloned.dataValidations = sheet.dataValidations.map((dv) => ({
      ...dv,
      values: dv.values ? [...dv.values] : undefined,
    }))
  }

  // Deep copy conditional rules
  if (sheet.conditionalRules) {
    cloned.conditionalRules = sheet.conditionalRules.map((rule) => {
      const clonedRule = { ...rule }
      if (rule.style) clonedRule.style = cloneCellStyle(rule.style)
      if (rule.formula && Array.isArray(rule.formula)) clonedRule.formula = [...rule.formula]
      if (rule.colorScale) {
        clonedRule.colorScale = {
          cfvo: rule.colorScale.cfvo.map((c) => ({ ...c })),
          colors: [...rule.colorScale.colors],
        }
      }
      if (rule.dataBar) {
        clonedRule.dataBar = {
          cfvo: rule.dataBar.cfvo.map((c) => ({ ...c })),
          color: rule.dataBar.color,
        }
      }
      if (rule.iconSet) {
        clonedRule.iconSet = {
          ...rule.iconSet,
          cfvo: rule.iconSet.cfvo.map((c) => ({ ...c })),
        }
      }
      return clonedRule
    })
  }

  // Copy autoFilter
  if (sheet.autoFilter) {
    cloned.autoFilter = { ...sheet.autoFilter }
  }

  // Copy freezePane
  if (sheet.freezePane) {
    cloned.freezePane = { ...sheet.freezePane }
  }

  // Deep copy images
  if (sheet.images) {
    // Spread rather than enumerate: listing the fields by hand is how
    // `altText` and `title` came to be dropped. Only the two nested
    // members need their own copy.
    cloned.images = sheet.images.map((img) => {
      const copy = { ...img, data: new Uint8Array(img.data) }
      copy.anchor = { ...img.anchor, from: { ...img.anchor.from } }
      if (img.anchor.to) copy.anchor.to = { ...img.anchor.to }
      return copy
    })
  }

  // Copy protection
  if (sheet.protection) {
    cloned.protection = { ...sheet.protection }
  }

  // Copy pageSetup
  if (sheet.pageSetup) {
    cloned.pageSetup = {
      ...sheet.pageSetup,
      margins: sheet.pageSetup.margins ? { ...sheet.pageSetup.margins } : undefined,
    }
  }

  // Copy headerFooter
  if (sheet.headerFooter) {
    cloned.headerFooter = { ...sheet.headerFooter }
  }

  // Copy view
  if (sheet.view) {
    cloned.view = {
      ...sheet.view,
      tabColor: sheet.view.tabColor ? { ...sheet.view.tabColor } : undefined,
    }
  }

  // Copy hidden/veryHidden
  if (sheet.hidden !== undefined) cloned.hidden = sheet.hidden
  if (sheet.veryHidden !== undefined) cloned.veryHidden = sheet.veryHidden

  // Deep copy tables
  if (sheet.tables) {
    cloned.tables = sheet.tables.map((table) => ({
      ...table,
      columns: table.columns.map((col) => ({ ...col })),
    }))
  }

  // Deep copy charts. Charts are plain JSON-serializable records (no
  // Map / Uint8Array / function members), so a structuredClone gives a
  // faithful independent copy without hand-walking the deep axis / series
  // / dataLabels trees. Carrying them here is what lets
  // copySheetToWorkbook bring charts across workbooks (issue #136).
  if (sheet.charts && sheet.charts.length > 0) {
    cloned.charts = structuredClone(sheet.charts)
  }

  // ── The rest of the sheet ──
  //
  // These used to be dropped silently — a "deep clone" that returned a
  // sheet with no sparklines, no text boxes, no page breaks and no
  // background image. `copySheetToWorkbook` is built on this, so copying
  // a sheet between workbooks lost them too. See #439 §N.
  //
  // Everything here is plain JSON-serialisable data except the background
  // image, which is bytes, so `structuredClone` is the faithful copy for
  // the trees and a `slice()` for the buffer.
  if (sheet.splitPane) cloned.splitPane = { ...sheet.splitPane }
  if (sheet.rowBreaks) cloned.rowBreaks = [...sheet.rowBreaks]
  if (sheet.colBreaks) cloned.colBreaks = [...sheet.colBreaks]
  if (sheet.outlineProperties) cloned.outlineProperties = { ...sheet.outlineProperties }
  if (sheet.backgroundImage) cloned.backgroundImage = sheet.backgroundImage.slice()
  if (sheet.sparklines) cloned.sparklines = structuredClone(sheet.sparklines)
  if (sheet.textBoxes) cloned.textBoxes = structuredClone(sheet.textBoxes)
  if (sheet.threadedComments) cloned.threadedComments = structuredClone(sheet.threadedComments)
  if (sheet.pivotTables) cloned.pivotTables = structuredClone(sheet.pivotTables)
  if (sheet.slicers) cloned.slicers = structuredClone(sheet.slicers)
  if (sheet.timelines) cloned.timelines = structuredClone(sheet.timelines)
  if (sheet.a11y) cloned.a11y = { ...sheet.a11y }
  if (sheet.defaultRowHeight !== undefined) cloned.defaultRowHeight = sheet.defaultRowHeight
  if (sheet.defaultColWidth !== undefined) cloned.defaultColWidth = sheet.defaultColWidth

  return cloned
}

// ── Copy Sheet To Workbook ──────────────────────────────────────────

/**
 * Copy a sheet from one workbook to another.
 * Clones the sheet and appends it to the target workbook.
 */
export function copySheetToWorkbook(
  sourceSheet: Sheet,
  targetWorkbook: Workbook,
  newName?: string,
): void {
  const cloned = cloneSheet(sourceSheet, newName ?? sourceSheet.name)
  targetWorkbook.sheets.push(cloned)
}

// ── Copy Range ──────────────────────────────────────────────────────

/**
 * Copy a range of cells from one location to another within the same sheet.
 * Copies values, styles, and merges.
 */
export function copyRange(
  sheet: Sheet,
  sourceRange: RangeLike,
  targetStart: { startRow: number; startCol: number } | string,
): void {
  // Either form of either argument — `copyRange(s, "A1:C3", "E1")` and the
  // coordinate spelling describe the same move. See #474.
  const source = toRange(sourceRange)
  const target =
    typeof targetStart === "string"
      ? (({ row, col }) => ({ startRow: row, startCol: col }))(parseCellRef(targetStart))
      : targetStart

  const rowCount = source.endRow - source.startRow + 1
  const colCount = source.endCol - source.startCol + 1

  // Ensure rows array is large enough for target
  const targetEndRow = target.startRow + rowCount - 1
  while (sheet.rows.length <= targetEndRow) {
    sheet.rows.push([])
  }

  // Read all source values and cells first (to handle overlapping ranges)
  const sourceValues: import("./_types").CellValue[][] = []
  const sourceCells: (Cell | null)[][] = []

  for (let r = 0; r < rowCount; r++) {
    sourceValues.push([])
    sourceCells.push([])
    for (let c = 0; c < colCount; c++) {
      const srcRow = source.startRow + r
      const srcCol = source.startCol + c

      // Read value
      const row = sheet.rows[srcRow]
      sourceValues[r].push(row && srcCol < row.length ? row[srcCol] : null)

      // Read cell
      if (sheet.cells) {
        const key = `${srcRow},${srcCol}`
        const cell = sheet.cells.get(key)
        sourceCells[r].push(cell ? cloneCell(cell) : null)
      } else {
        sourceCells[r].push(null)
      }
    }
  }

  // Write values and cells to target
  for (let r = 0; r < rowCount; r++) {
    const tgtRow = target.startRow + r
    const row = sheet.rows[tgtRow]

    for (let c = 0; c < colCount; c++) {
      const tgtCol = target.startCol + c

      // Extend row if needed
      while (row.length <= tgtCol) row.push(null)
      row[tgtCol] = sourceValues[r][c]

      // Copy cell data
      const srcCell = sourceCells[r][c]
      if (srcCell) {
        if (!sheet.cells) sheet.cells = new Map()
        sheet.cells.set(`${tgtRow},${tgtCol}`, srcCell)
      } else if (sheet.cells) {
        sheet.cells.delete(`${tgtRow},${tgtCol}`)
      }
    }
  }

  // Copy merges that are fully within the source range
  if (sheet.merges) {
    const newMerges: MergeRange[] = []
    for (const merge of sheet.merges) {
      if (
        merge.startRow >= source.startRow &&
        merge.endRow <= source.endRow &&
        merge.startCol >= source.startCol &&
        merge.endCol <= source.endCol
      ) {
        const rowOffset = target.startRow - source.startRow
        const colOffset = target.startCol - source.startCol
        newMerges.push({
          startRow: merge.startRow + rowOffset,
          startCol: merge.startCol + colOffset,
          endRow: merge.endRow + rowOffset,
          endCol: merge.endCol + colOffset,
        })
      }
    }
    // Append new merges (avoid duplicates by checking if already exists)
    for (const nm of newMerges) {
      const exists = sheet.merges.some(
        (m) =>
          m.startRow === nm.startRow &&
          m.startCol === nm.startCol &&
          m.endRow === nm.endRow &&
          m.endCol === nm.endCol,
      )
      if (!exists) {
        sheet.merges.push(nm)
      }
    }
  }
}

// ── Move Sheet ──────────────────────────────────────────────────────

/**
 * Reorder sheets in a workbook.
 */
export function moveSheet(workbook: Workbook, fromIndex: number, toIndex: number): void {
  if (fromIndex === toIndex) return
  const [sheet] = workbook.sheets.splice(fromIndex, 1)
  workbook.sheets.splice(toIndex, 0, sheet)
}

// ── Remove Sheet ────────────────────────────────────────────────────

/**
 * Remove a sheet from a workbook.
 */
export function removeSheet(workbook: Workbook, index: number): void {
  workbook.sheets.splice(index, 1)
  // Adjust activeSheet if needed
  if (workbook.activeSheet !== undefined) {
    if (workbook.activeSheet === index) {
      // If we removed the active sheet, set to the previous sheet or 0
      workbook.activeSheet =
        workbook.sheets.length > 0 ? Math.min(index, workbook.sheets.length - 1) : 0
    } else if (workbook.activeSheet > index) {
      workbook.activeSheet--
    }
  }
}

// ── Cell Search ─────────────────────────────────────────────────────

/**
 * Find cells matching a value or predicate.
 *
 * @param sheet - The sheet to search
 * @param predicate - A value to match exactly, or a function `(value, row, col) => boolean`
 * @returns Array of matching cells with their positions and values
 */
export function findCells(
  sheet: Sheet,
  predicate: CellValue | RegExp | ((value: CellValue, row: number, col: number) => boolean),
): Array<{ row: number; col: number; value: CellValue }> {
  const results: Array<{ row: number; col: number; value: CellValue }> = []
  const isFn = typeof predicate === "function"
  // `replaceCells` has always taken a RegExp; this one took a predicate
  // instead, so "find the cells I am about to replace" could not be
  // written with the same argument. Both take all three forms now.
  const isRegExp = predicate instanceof RegExp

  for (let r = 0; r < sheet.rows.length; r++) {
    const row = sheet.rows[r]!
    for (let c = 0; c < row.length; c++) {
      const value = row[c] ?? null
      let match: boolean
      if (isFn) {
        match = (predicate as (value: CellValue, row: number, col: number) => boolean)(value, r, c)
      } else if (isRegExp) {
        // Same rule as replaceCells: a RegExp tests strings only. `lastIndex`
        // on a /g pattern would make the result depend on call order, so it
        // is reset before each test.
        predicate.lastIndex = 0
        match = typeof value === "string" && predicate.test(value)
      } else {
        match = value === predicate
      }
      if (match) {
        results.push({ row: r, col: c, value })
      }
    }
  }

  return results
}

/**
 * Find and replace cell values in a sheet.
 *
 * @param sheet - The sheet to modify (mutated in place)
 * @param find - The value or RegExp to search for
 * @param replace - The replacement value. For RegExp finds on string cells,
 *                  if replace is a string, `String.replace(regex, replace)` is used.
 * @returns The number of cells that were modified
 */
export function replaceCells(sheet: Sheet, find: CellValue | RegExp, replace: CellValue): number {
  let count = 0

  for (let r = 0; r < sheet.rows.length; r++) {
    const row = sheet.rows[r]!
    for (let c = 0; c < row.length; c++) {
      const value = row[c] ?? null

      if (find instanceof RegExp) {
        // RegExp matching: only applies to string cells
        if (typeof value === "string" && find.test(value)) {
          if (typeof replace === "string") {
            // Reset lastIndex for global regexes
            find.lastIndex = 0
            row[c] = value.replace(find, replace)
          } else {
            row[c] = replace
          }
          // Reset lastIndex after test() for global regexes
          find.lastIndex = 0
          syncCellOverride(sheet, r, c, row[c]!)
          count++
        }
      } else {
        // Exact value matching
        if (value === find) {
          row[c] = replace
          syncCellOverride(sheet, r, c, replace)
          count++
        }
      }
    }
  }

  return count
}

// ── Sort Rows ────────────────────────────────────────────────────────

/**
 * Sort sheet rows by the values in a given column.
 * Handles mixed types: nulls last, numbers < strings < booleans.
 *
 * @param sheet - The sheet to sort (mutated in place)
 * @param colIndex - 0-based column index to sort by
 * @param order - Sort order: "asc" (default) or "desc"
 */
export function sortRows(sheet: Sheet, colIndex: number, order?: "asc" | "desc"): void {
  const desc = order === "desc"

  // A merged range pins cells to positions, and a sort moves rows past
  // those positions — there is no arrangement that keeps both. Excel
  // refuses the operation outright for the same reason; sorting anyway
  // left the merge covering whatever happened to land there. A merge
  // wholly inside one row is unaffected, since that row moves as a unit.
  const spansRows = sheet.merges?.some((m) => m.endRow > m.startRow)
  if (spansRows) {
    throw new InvalidArgumentError(
      "Cannot sort rows: the sheet has a merged range spanning more than one row, " +
        "and no ordering can keep both the sort and the merge. Remove the merge first.",
    )
  }

  // Everything keyed by row index has to move with its row: the per-cell
  // override Map (styles, formulas, hyperlinks), the row definitions
  // (heights, hidden, outline levels), and single-row merges. Tag each row
  // with its original index, sort, then remap through old→new.
  const tagged = sheet.rows.map((row, i) => ({ row, i }))
  tagged.sort((a, b) => {
    const va = colIndex < a.row.length ? (a.row[colIndex] ?? null) : null
    const vb = colIndex < b.row.length ? (b.row[colIndex] ?? null) : null
    return compareCellValues(va, vb, desc)
  })

  const oldToNew = new Map<number, number>()
  for (let newIdx = 0; newIdx < tagged.length; newIdx++) {
    oldToNew.set(tagged[newIdx]!.i, newIdx)
  }

  sheet.rows = tagged.map((t) => t.row)

  if (sheet.cells && sheet.cells.size > 0) {
    const remapped = new Map<string, Cell>()
    for (const [key, cell] of sheet.cells) {
      const comma = key.indexOf(",")
      const oldRow = Number(key.slice(0, comma))
      const col = key.slice(comma + 1)
      const newRow = oldToNew.get(oldRow)
      // Keep non-positional keys untouched if any slipped in.
      remapped.set(newRow === undefined ? key : `${newRow},${col}`, cell)
    }
    sheet.cells = remapped
  }

  if (sheet.rowDefs && sheet.rowDefs.size > 0) {
    const remapped = new Map<number, RowDef>()
    for (const [row, def] of sheet.rowDefs) {
      remapped.set(oldToNew.get(row) ?? row, def)
    }
    sheet.rowDefs = remapped
  }

  if (sheet.merges) {
    for (const merge of sheet.merges) {
      const moved = oldToNew.get(merge.startRow)
      if (moved !== undefined) {
        merge.startRow = moved
        merge.endRow = moved
      }
    }
  }
}

/**
 * Keep a sheet's per-cell override Map in sync when a row value changes via
 * {@link replaceCells}: if an override exists at (row,col), update its
 * `value` so the writer (which prefers the override) doesn't emit the stale
 * pre-replace value.
 */
function syncCellOverride(sheet: Sheet, row: number, col: number, value: CellValue): void {
  const existing = sheet.cells?.get(`${row},${col}`)
  if (existing) existing.value = value
}

/** Compare two cell values for sorting: nulls last, numbers < strings < booleans. */
/**
 * Compare two cell values for {@link sortRows}.
 *
 * `desc` is applied to the *value* comparison only. Negating the whole
 * result flipped the null rule with it, so descending floated blanks to
 * the top — against this function's own contract and against Excel,
 * which sinks blanks in both directions. See #392.
 */
function compareCellValues(a: CellValue, b: CellValue, desc = false): number {
  // Nulls last, regardless of direction
  if (a === null && b === null) return 0
  if (a === null) return 1
  if (b === null) return -1

  return desc ? -compareNonNull(a, b) : compareNonNull(a, b)
}

function compareNonNull(a: CellValue, b: CellValue): number {
  const ta = typeRank(a)
  const tb = typeRank(b)
  if (ta !== tb) return ta - tb

  // Same type
  if (typeof a === "number" && typeof b === "number") return a - b
  if (typeof a === "string" && typeof b === "string") return a.localeCompare(b)
  if (typeof a === "boolean" && typeof b === "boolean") return (a ? 1 : 0) - (b ? 1 : 0)
  if (a instanceof Date && b instanceof Date) return a.getTime() - b.getTime()
  return 0
}

function typeRank(v: CellValue): number {
  if (v === null) return 4
  if (typeof v === "number") return 0
  if (v instanceof Date) return 1
  if (typeof v === "string") return 2
  if (typeof v === "boolean") return 3
  return 4
}

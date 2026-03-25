// ── Sheet Operations ────────────────────────────────────────────────
// In-memory row/column manipulation utilities for Sheet objects.

import type { Sheet, MergeRange, RowDef, Workbook, Cell, CellStyle, CellValue } from "./_types";
import { parseCellRef } from "./xlsx/worksheet";
import { rangeRef } from "./xlsx/worksheet-writer";

// ── Range Helpers ────────────────────────────────────────────────────

/**
 * Parse a range string like "A1:D10" into 0-based coordinates.
 */
function parseRange(range: string): MergeRange {
  const parts = range.split(":");
  const start = parseCellRef(parts[0]);
  const end = parts.length > 1 ? parseCellRef(parts[1]) : start;
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  };
}

/**
 * Build a range string from 0-based coordinates.
 */
function buildRange(r: MergeRange): string {
  return rangeRef(r.startRow, r.startCol, r.endRow, r.endCol);
}

/**
 * Shift row references in a range string by a given delta.
 * Only rows >= threshold are shifted.
 */
function shiftRangeRows(range: string, threshold: number, delta: number): string {
  const r = parseRange(range);
  if (r.startRow >= threshold) r.startRow += delta;
  if (r.endRow >= threshold) r.endRow += delta;
  return buildRange(r);
}

/**
 * Shift column references in a range string by a given delta.
 * Only columns >= threshold are shifted.
 */
function shiftRangeCols(range: string, threshold: number, delta: number): string {
  const r = parseRange(range);
  if (r.startCol >= threshold) r.startCol += delta;
  if (r.endCol >= threshold) r.endCol += delta;
  return buildRange(r);
}

// ── Row Width Helper ─────────────────────────────────────────────────

function getRowWidth(sheet: Sheet): number {
  let width = 0;
  for (const row of sheet.rows) {
    if (row.length > width) width = row.length;
  }
  if (sheet.columns && sheet.columns.length > width) {
    width = sheet.columns.length;
  }
  return width;
}

function makeEmptyRow(width: number): null[] {
  const row: null[] = [];
  for (let i = 0; i < width; i++) row.push(null);
  return row;
}

// ── Insert Rows ──────────────────────────────────────────────────────

/**
 * Insert rows at the given position (0-based), shifting existing rows down.
 * Updates merge ranges, data validations, conditional rules, auto filter,
 * images, and cells Map keys.
 */
export function insertRows(sheet: Sheet, rowIndex: number, count: number): void {
  if (count <= 0) return;

  const width = getRowWidth(sheet);
  const newRows: null[][] = [];
  for (let i = 0; i < count; i++) {
    newRows.push(makeEmptyRow(width));
  }

  // Insert into rows array
  sheet.rows.splice(rowIndex, 0, ...newRows);

  // Update cells Map
  if (sheet.cells && sheet.cells.size > 0) {
    const updated = new Map<string, import("./_types").Cell>();
    for (const [key, cell] of sheet.cells) {
      const [rowStr, colStr] = key.split(",");
      const row = Number(rowStr);
      const col = Number(colStr);
      if (row >= rowIndex) {
        updated.set(`${row + count},${col}`, cell);
      } else {
        updated.set(key, cell);
      }
    }
    sheet.cells = updated;
  }

  // Update merge ranges
  if (sheet.merges) {
    for (const merge of sheet.merges) {
      if (merge.startRow >= rowIndex) {
        merge.startRow += count;
        merge.endRow += count;
      } else if (merge.endRow >= rowIndex) {
        // Merge starts before insertion but ends at or after — expand it
        merge.endRow += count;
      }
    }
  }

  // Update data validations
  if (sheet.dataValidations) {
    for (const dv of sheet.dataValidations) {
      dv.range = shiftRangeRows(dv.range, rowIndex, count);
    }
  }

  // Update conditional rules
  if (sheet.conditionalRules) {
    for (const rule of sheet.conditionalRules) {
      rule.range = shiftRangeRows(rule.range, rowIndex, count);
    }
  }

  // Update auto filter
  if (sheet.autoFilter) {
    sheet.autoFilter.range = shiftRangeRows(sheet.autoFilter.range, rowIndex, count);
  }

  // Update image anchors
  if (sheet.images) {
    for (const img of sheet.images) {
      if (img.anchor.from.row >= rowIndex) {
        img.anchor.from.row += count;
      }
      if (img.anchor.to && img.anchor.to.row >= rowIndex) {
        img.anchor.to.row += count;
      }
    }
  }

  // Update row defs
  if (sheet.rowDefs && sheet.rowDefs.size > 0) {
    const updated = new Map<number, RowDef>();
    for (const [row, def] of sheet.rowDefs) {
      if (row >= rowIndex) {
        updated.set(row + count, def);
      } else {
        updated.set(row, def);
      }
    }
    sheet.rowDefs = updated;
  }

  // Update table ranges
  if (sheet.tables) {
    for (const table of sheet.tables) {
      if (table.range) {
        table.range = shiftRangeRows(table.range, rowIndex, count);
      }
    }
  }
}

// ── Delete Rows ──────────────────────────────────────────────────────

/**
 * Delete rows starting at the given position (0-based), shifting remaining rows up.
 * Removes merges fully within deleted range. Adjusts merges that partially overlap.
 */
export function deleteRows(sheet: Sheet, rowIndex: number, count: number): void {
  if (count <= 0) return;

  const deleteEnd = rowIndex + count; // exclusive

  // Remove rows from array
  sheet.rows.splice(rowIndex, count);

  // Update cells Map
  if (sheet.cells && sheet.cells.size > 0) {
    const updated = new Map<string, import("./_types").Cell>();
    for (const [key, cell] of sheet.cells) {
      const [rowStr, colStr] = key.split(",");
      const row = Number(rowStr);
      const col = Number(colStr);
      if (row >= rowIndex && row < deleteEnd) {
        // Cell is in deleted range — remove it
        continue;
      } else if (row >= deleteEnd) {
        updated.set(`${row - count},${col}`, cell);
      } else {
        updated.set(key, cell);
      }
    }
    sheet.cells = updated;
  }

  // Update merge ranges
  if (sheet.merges) {
    sheet.merges = sheet.merges.filter((merge) => {
      // Fully within deleted range — remove
      if (merge.startRow >= rowIndex && merge.endRow < deleteEnd) {
        return false;
      }
      return true;
    });

    for (const merge of sheet.merges) {
      if (merge.startRow >= deleteEnd) {
        // Entirely below deleted range — shift up
        merge.startRow -= count;
        merge.endRow -= count;
      } else if (merge.endRow >= deleteEnd) {
        // Partially overlapping: starts before or at deletion, ends after
        if (merge.startRow >= rowIndex) {
          // Starts within deleted range — clamp start to rowIndex
          merge.startRow = rowIndex;
          merge.endRow -= count;
        } else {
          // Starts before deleted range — shrink end
          merge.endRow -= count;
        }
      } else if (merge.endRow >= rowIndex) {
        // Ends within deleted range but starts before — clamp end
        merge.endRow = rowIndex - 1;
      }
    }

    // Remove degenerate merges (start > end)
    sheet.merges = sheet.merges.filter((m) => m.startRow <= m.endRow && m.startCol <= m.endCol);
  }

  // Update data validations
  if (sheet.dataValidations) {
    sheet.dataValidations = sheet.dataValidations.filter((dv) => {
      const r = parseRange(dv.range);
      // Remove if fully within deleted range
      if (r.startRow >= rowIndex && r.endRow < deleteEnd) return false;
      return true;
    });
    for (const dv of sheet.dataValidations) {
      dv.range = shiftDeletedRangeRows(dv.range, rowIndex, count);
    }
  }

  // Update conditional rules
  if (sheet.conditionalRules) {
    sheet.conditionalRules = sheet.conditionalRules.filter((rule) => {
      const r = parseRange(rule.range);
      if (r.startRow >= rowIndex && r.endRow < deleteEnd) return false;
      return true;
    });
    for (const rule of sheet.conditionalRules) {
      rule.range = shiftDeletedRangeRows(rule.range, rowIndex, count);
    }
  }

  // Update auto filter
  if (sheet.autoFilter) {
    const r = parseRange(sheet.autoFilter.range);
    if (r.startRow >= rowIndex && r.endRow < deleteEnd) {
      sheet.autoFilter = undefined;
    } else {
      sheet.autoFilter.range = shiftDeletedRangeRows(sheet.autoFilter.range, rowIndex, count);
    }
  }

  // Update image anchors
  if (sheet.images) {
    sheet.images = sheet.images.filter((img) => {
      // Remove images whose anchor starts in deleted range
      return !(img.anchor.from.row >= rowIndex && img.anchor.from.row < deleteEnd);
    });
    for (const img of sheet.images) {
      if (img.anchor.from.row >= deleteEnd) {
        img.anchor.from.row -= count;
      }
      if (img.anchor.to && img.anchor.to.row >= deleteEnd) {
        img.anchor.to.row -= count;
      }
    }
  }

  // Update row defs
  if (sheet.rowDefs && sheet.rowDefs.size > 0) {
    const updated = new Map<number, RowDef>();
    for (const [row, def] of sheet.rowDefs) {
      if (row >= rowIndex && row < deleteEnd) {
        continue; // deleted
      } else if (row >= deleteEnd) {
        updated.set(row - count, def);
      } else {
        updated.set(row, def);
      }
    }
    sheet.rowDefs = updated;
  }

  // Update table ranges
  if (sheet.tables) {
    sheet.tables = sheet.tables.filter((table) => {
      if (!table.range) return true;
      const r = parseRange(table.range);
      return !(r.startRow >= rowIndex && r.endRow < deleteEnd);
    });
    for (const table of sheet.tables) {
      if (table.range) {
        table.range = shiftDeletedRangeRows(table.range, rowIndex, count);
      }
    }
  }
}

/**
 * Shift row references in a range string after deletion.
 * Rows >= deleteEnd shift up by count.
 * Rows within [rowIndex, deleteEnd) are clamped.
 */
function shiftDeletedRangeRows(range: string, rowIndex: number, count: number): string {
  const deleteEnd = rowIndex + count;
  const r = parseRange(range);

  if (r.startRow >= deleteEnd) {
    r.startRow -= count;
  } else if (r.startRow >= rowIndex) {
    r.startRow = rowIndex;
  }

  if (r.endRow >= deleteEnd) {
    r.endRow -= count;
  } else if (r.endRow >= rowIndex) {
    r.endRow = rowIndex > 0 ? rowIndex - 1 : 0;
  }

  return buildRange(r);
}

// ── Insert Columns ───────────────────────────────────────────────────

/**
 * Insert columns at the given position (0-based), shifting existing columns right.
 * Updates merge ranges, data validations, conditional rules, auto filter,
 * images, column defs, and cells Map keys.
 */
export function insertColumns(sheet: Sheet, colIndex: number, count: number): void {
  if (count <= 0) return;

  const nulls: null[] = makeEmptyRow(count);

  // Insert nulls into each row
  for (const row of sheet.rows) {
    // Extend row if it's shorter than colIndex
    while (row.length < colIndex) row.push(null);
    row.splice(colIndex, 0, ...nulls);
  }

  // Update column defs
  if (sheet.columns) {
    const newCols: import("./_types").ColumnDef[] = [];
    for (let i = 0; i < count; i++) newCols.push({});
    // Ensure columns array is long enough
    while (sheet.columns.length < colIndex) sheet.columns.push({});
    sheet.columns.splice(colIndex, 0, ...newCols);
  }

  // Update cells Map
  if (sheet.cells && sheet.cells.size > 0) {
    const updated = new Map<string, import("./_types").Cell>();
    for (const [key, cell] of sheet.cells) {
      const [rowStr, colStr] = key.split(",");
      const row = Number(rowStr);
      const col = Number(colStr);
      if (col >= colIndex) {
        updated.set(`${row},${col + count}`, cell);
      } else {
        updated.set(key, cell);
      }
    }
    sheet.cells = updated;
  }

  // Update merge ranges
  if (sheet.merges) {
    for (const merge of sheet.merges) {
      if (merge.startCol >= colIndex) {
        merge.startCol += count;
        merge.endCol += count;
      } else if (merge.endCol >= colIndex) {
        merge.endCol += count;
      }
    }
  }

  // Update data validations
  if (sheet.dataValidations) {
    for (const dv of sheet.dataValidations) {
      dv.range = shiftRangeCols(dv.range, colIndex, count);
    }
  }

  // Update conditional rules
  if (sheet.conditionalRules) {
    for (const rule of sheet.conditionalRules) {
      rule.range = shiftRangeCols(rule.range, colIndex, count);
    }
  }

  // Update auto filter
  if (sheet.autoFilter) {
    sheet.autoFilter.range = shiftRangeCols(sheet.autoFilter.range, colIndex, count);
  }

  // Update image anchors
  if (sheet.images) {
    for (const img of sheet.images) {
      if (img.anchor.from.col >= colIndex) {
        img.anchor.from.col += count;
      }
      if (img.anchor.to && img.anchor.to.col >= colIndex) {
        img.anchor.to.col += count;
      }
    }
  }

  // Update table ranges
  if (sheet.tables) {
    for (const table of sheet.tables) {
      if (table.range) {
        table.range = shiftRangeCols(table.range, colIndex, count);
      }
    }
  }
}

// ── Delete Columns ───────────────────────────────────────────────────

/**
 * Delete columns starting at the given position (0-based), shifting remaining columns left.
 * Removes merges fully within deleted range. Adjusts merges that partially overlap.
 */
export function deleteColumns(sheet: Sheet, colIndex: number, count: number): void {
  if (count <= 0) return;

  const deleteEnd = colIndex + count; // exclusive

  // Remove columns from each row
  for (const row of sheet.rows) {
    if (colIndex < row.length) {
      row.splice(colIndex, Math.min(count, row.length - colIndex));
    }
  }

  // Update column defs
  if (sheet.columns) {
    if (colIndex < sheet.columns.length) {
      sheet.columns.splice(colIndex, Math.min(count, sheet.columns.length - colIndex));
    }
  }

  // Update cells Map
  if (sheet.cells && sheet.cells.size > 0) {
    const updated = new Map<string, import("./_types").Cell>();
    for (const [key, cell] of sheet.cells) {
      const [rowStr, colStr] = key.split(",");
      const row = Number(rowStr);
      const col = Number(colStr);
      if (col >= colIndex && col < deleteEnd) {
        continue; // deleted
      } else if (col >= deleteEnd) {
        updated.set(`${row},${col - count}`, cell);
      } else {
        updated.set(key, cell);
      }
    }
    sheet.cells = updated;
  }

  // Update merge ranges
  if (sheet.merges) {
    sheet.merges = sheet.merges.filter((merge) => {
      if (merge.startCol >= colIndex && merge.endCol < deleteEnd) {
        return false;
      }
      return true;
    });

    for (const merge of sheet.merges) {
      if (merge.startCol >= deleteEnd) {
        merge.startCol -= count;
        merge.endCol -= count;
      } else if (merge.endCol >= deleteEnd) {
        if (merge.startCol >= colIndex) {
          merge.startCol = colIndex;
          merge.endCol -= count;
        } else {
          merge.endCol -= count;
        }
      } else if (merge.endCol >= colIndex) {
        merge.endCol = colIndex - 1;
      }
    }

    sheet.merges = sheet.merges.filter((m) => m.startRow <= m.endRow && m.startCol <= m.endCol);
  }

  // Update data validations
  if (sheet.dataValidations) {
    sheet.dataValidations = sheet.dataValidations.filter((dv) => {
      const r = parseRange(dv.range);
      if (r.startCol >= colIndex && r.endCol < deleteEnd) return false;
      return true;
    });
    for (const dv of sheet.dataValidations) {
      dv.range = shiftDeletedRangeCols(dv.range, colIndex, count);
    }
  }

  // Update conditional rules
  if (sheet.conditionalRules) {
    sheet.conditionalRules = sheet.conditionalRules.filter((rule) => {
      const r = parseRange(rule.range);
      if (r.startCol >= colIndex && r.endCol < deleteEnd) return false;
      return true;
    });
    for (const rule of sheet.conditionalRules) {
      rule.range = shiftDeletedRangeCols(rule.range, colIndex, count);
    }
  }

  // Update auto filter
  if (sheet.autoFilter) {
    const r = parseRange(sheet.autoFilter.range);
    if (r.startCol >= colIndex && r.endCol < deleteEnd) {
      sheet.autoFilter = undefined;
    } else {
      sheet.autoFilter.range = shiftDeletedRangeCols(sheet.autoFilter.range, colIndex, count);
    }
  }

  // Update image anchors
  if (sheet.images) {
    sheet.images = sheet.images.filter((img) => {
      return !(img.anchor.from.col >= colIndex && img.anchor.from.col < deleteEnd);
    });
    for (const img of sheet.images) {
      if (img.anchor.from.col >= deleteEnd) {
        img.anchor.from.col -= count;
      }
      if (img.anchor.to && img.anchor.to.col >= deleteEnd) {
        img.anchor.to.col -= count;
      }
    }
  }

  // Update table ranges
  if (sheet.tables) {
    sheet.tables = sheet.tables.filter((table) => {
      if (!table.range) return true;
      const r = parseRange(table.range);
      return !(r.startCol >= colIndex && r.endCol < deleteEnd);
    });
    for (const table of sheet.tables) {
      if (table.range) {
        table.range = shiftDeletedRangeCols(table.range, colIndex, count);
      }
    }
  }
}

/**
 * Shift column references in a range string after deletion.
 */
function shiftDeletedRangeCols(range: string, colIndex: number, count: number): string {
  const deleteEnd = colIndex + count;
  const r = parseRange(range);

  if (r.startCol >= deleteEnd) {
    r.startCol -= count;
  } else if (r.startCol >= colIndex) {
    r.startCol = colIndex;
  }

  if (r.endCol >= deleteEnd) {
    r.endCol -= count;
  } else if (r.endCol >= colIndex) {
    r.endCol = colIndex > 0 ? colIndex - 1 : 0;
  }

  return buildRange(r);
}

// ── Move Rows ────────────────────────────────────────────────────────

/**
 * Move rows from one position to another.
 * Extracts `count` rows starting at `fromIndex` and inserts them at `toIndex`.
 * `toIndex` is the target position in the original (pre-move) coordinate space.
 */
export function moveRows(sheet: Sheet, fromIndex: number, count: number, toIndex: number): void {
  if (count <= 0 || fromIndex === toIndex) return;

  // Extract rows
  const extractedRows = sheet.rows.splice(fromIndex, count);

  // Extract cells for moved rows
  const extractedCells = new Map<string, import("./_types").Cell>();
  if (sheet.cells) {
    for (const [key, cell] of sheet.cells) {
      const [rowStr] = key.split(",");
      const row = Number(rowStr);
      if (row >= fromIndex && row < fromIndex + count) {
        extractedCells.set(key, cell);
        sheet.cells.delete(key);
      }
    }
  }

  // Extract row defs for moved rows
  const extractedRowDefs = new Map<number, RowDef>();
  if (sheet.rowDefs) {
    for (const [row, def] of sheet.rowDefs) {
      if (row >= fromIndex && row < fromIndex + count) {
        extractedRowDefs.set(row, def);
        sheet.rowDefs.delete(row);
      }
    }
  }

  // After removing from source, adjust target index
  let adjustedTo = toIndex;
  if (toIndex > fromIndex) {
    adjustedTo = toIndex - count;
  }

  // Re-insert rows at adjusted position
  sheet.rows.splice(adjustedTo, 0, ...extractedRows);

  // Rebuild cells Map: shift all remaining cells, then re-add extracted
  if (sheet.cells || extractedCells.size > 0) {
    const newCells = new Map<string, import("./_types").Cell>();

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
        const [rowStr, colStr] = key.split(",");
        let row = Number(rowStr);
        const col = Number(colStr);

        // After removal of [fromIndex, fromIndex+count):
        if (row >= fromIndex) {
          row -= count;
        }
        // After insertion at adjustedTo:
        if (row >= adjustedTo) {
          row += count;
        }

        newCells.set(`${row},${col}`, cell);
      }
    }

    // Re-add extracted cells at their new positions
    for (const [key, cell] of extractedCells) {
      const [rowStr, colStr] = key.split(",");
      const originalRow = Number(rowStr);
      const col = Number(colStr);
      const offset = originalRow - fromIndex;
      const newRow = adjustedTo + offset;
      newCells.set(`${newRow},${col}`, cell);
    }

    sheet.cells = newCells.size > 0 ? newCells : undefined;
  }

  // Rebuild row defs
  if (sheet.rowDefs || extractedRowDefs.size > 0) {
    const newRowDefs = new Map<number, RowDef>();

    if (sheet.rowDefs) {
      for (const [row, def] of sheet.rowDefs) {
        let newRow = row;
        if (newRow >= fromIndex) {
          newRow -= count;
        }
        if (newRow >= adjustedTo) {
          newRow += count;
        }
        newRowDefs.set(newRow, def);
      }
    }

    for (const [row, def] of extractedRowDefs) {
      const offset = row - fromIndex;
      newRowDefs.set(adjustedTo + offset, def);
    }

    sheet.rowDefs = newRowDefs.size > 0 ? newRowDefs : undefined;
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
  if (!sheet.rowDefs) sheet.rowDefs = new Map();
  for (let i = startRow; i < startRow + count; i++) {
    const existing = sheet.rowDefs.get(i) || {};
    existing.hidden = hidden;
    sheet.rowDefs.set(i, existing);
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
  if (!sheet.columns) sheet.columns = [];
  // Ensure columns array is large enough
  while (sheet.columns.length <= startCol + count - 1) {
    sheet.columns.push({});
  }
  for (let i = startCol; i < startCol + count; i++) {
    sheet.columns[i].hidden = hidden;
  }
}

// ── Group Rows ───────────────────────────────────────────────────────

/**
 * Set outline level for rows in range [startRow, endRow] (inclusive, 0-based).
 * @param level - Outline level (default 1). Set to 0 to ungroup.
 */
export function groupRows(sheet: Sheet, startRow: number, endRow: number, level: number = 1): void {
  if (!sheet.rowDefs) sheet.rowDefs = new Map();
  for (let i = startRow; i <= endRow; i++) {
    const existing = sheet.rowDefs.get(i) || {};
    existing.outlineLevel = level;
    sheet.rowDefs.set(i, existing);
  }
}

// ── Deep Clone Helpers ────────────────────────────────────────────────

function cloneStyle(style: CellStyle): CellStyle {
  const result: CellStyle = {};
  if (style.font)
    result.font = { ...style.font, color: style.font.color ? { ...style.font.color } : undefined };
  if (style.fill) {
    if (style.fill.type === "pattern") {
      result.fill = {
        type: "pattern",
        pattern: style.fill.pattern,
        fgColor: style.fill.fgColor ? { ...style.fill.fgColor } : undefined,
        bgColor: style.fill.bgColor ? { ...style.fill.bgColor } : undefined,
      };
    } else {
      result.fill = {
        type: "gradient",
        degree: style.fill.degree,
        stops: style.fill.stops.map((s) => ({ position: s.position, color: { ...s.color } })),
      };
    }
  }
  if (style.border) {
    result.border = {
      ...style.border,
      top: style.border.top
        ? {
            ...style.border.top,
            color: style.border.top.color ? { ...style.border.top.color } : undefined,
          }
        : undefined,
      right: style.border.right
        ? {
            ...style.border.right,
            color: style.border.right.color ? { ...style.border.right.color } : undefined,
          }
        : undefined,
      bottom: style.border.bottom
        ? {
            ...style.border.bottom,
            color: style.border.bottom.color ? { ...style.border.bottom.color } : undefined,
          }
        : undefined,
      left: style.border.left
        ? {
            ...style.border.left,
            color: style.border.left.color ? { ...style.border.left.color } : undefined,
          }
        : undefined,
      diagonal: style.border.diagonal
        ? {
            ...style.border.diagonal,
            color: style.border.diagonal.color ? { ...style.border.diagonal.color } : undefined,
          }
        : undefined,
    };
  }
  if (style.alignment) result.alignment = { ...style.alignment };
  if (style.numFmt !== undefined) result.numFmt = style.numFmt;
  if (style.protection) result.protection = { ...style.protection };
  return result;
}

function cloneCell(cell: Cell): Cell {
  const result: Cell = { value: cell.value, type: cell.type };
  if (cell.style) result.style = cloneStyle(cell.style);
  if (cell.formula !== undefined) result.formula = cell.formula;
  if (cell.formulaResult !== undefined) result.formulaResult = cell.formulaResult;
  if (cell.richText)
    result.richText = cell.richText.map((r) => ({
      text: r.text,
      font: r.font
        ? { ...r.font, color: r.font.color ? { ...r.font.color } : undefined }
        : undefined,
    }));
  if (cell.hyperlink) result.hyperlink = { ...cell.hyperlink };
  if (cell.comment) {
    result.comment = { text: cell.comment.text, author: cell.comment.author };
    if (cell.comment.richText) {
      result.comment.richText = cell.comment.richText.map((r) => ({
        text: r.text,
        font: r.font
          ? { ...r.font, color: r.font.color ? { ...r.font.color } : undefined }
          : undefined,
      }));
    }
  }
  return result;
}

// ── Clone Sheet ─────────────────────────────────────────────────────

/**
 * Deep clone a sheet (all data, styles, merges, validations, etc.).
 * The cloned sheet gets a new name.
 */
export function cloneSheet(sheet: Sheet, newName: string): Sheet {
  // Deep copy rows
  const rows = sheet.rows.map((row) => [...row]);

  const cloned: Sheet = { name: newName, rows };

  // Deep copy cells Map
  if (sheet.cells && sheet.cells.size > 0) {
    const cells = new Map<string, Cell>();
    for (const [key, cell] of sheet.cells) {
      cells.set(key, cloneCell(cell));
    }
    cloned.cells = cells;
  }

  // Deep copy columns
  if (sheet.columns) {
    cloned.columns = sheet.columns.map((col) => ({
      ...col,
      style: col.style ? cloneStyle(col.style) : undefined,
    }));
  }

  // Deep copy rowDefs
  if (sheet.rowDefs && sheet.rowDefs.size > 0) {
    const rowDefs = new Map<number, RowDef>();
    for (const [key, def] of sheet.rowDefs) {
      rowDefs.set(key, { ...def });
    }
    cloned.rowDefs = rowDefs;
  }

  // Deep copy merges
  if (sheet.merges) {
    cloned.merges = sheet.merges.map((m) => ({ ...m }));
  }

  // Deep copy data validations
  if (sheet.dataValidations) {
    cloned.dataValidations = sheet.dataValidations.map((dv) => ({
      ...dv,
      values: dv.values ? [...dv.values] : undefined,
    }));
  }

  // Deep copy conditional rules
  if (sheet.conditionalRules) {
    cloned.conditionalRules = sheet.conditionalRules.map((rule) => {
      const clonedRule = { ...rule };
      if (rule.style) clonedRule.style = cloneStyle(rule.style);
      if (rule.formula && Array.isArray(rule.formula)) clonedRule.formula = [...rule.formula];
      if (rule.colorScale) {
        clonedRule.colorScale = {
          cfvo: rule.colorScale.cfvo.map((c) => ({ ...c })),
          colors: [...rule.colorScale.colors],
        };
      }
      if (rule.dataBar) {
        clonedRule.dataBar = {
          cfvo: rule.dataBar.cfvo.map((c) => ({ ...c })),
          color: rule.dataBar.color,
        };
      }
      if (rule.iconSet) {
        clonedRule.iconSet = {
          ...rule.iconSet,
          cfvo: rule.iconSet.cfvo.map((c) => ({ ...c })),
        };
      }
      return clonedRule;
    });
  }

  // Copy autoFilter
  if (sheet.autoFilter) {
    cloned.autoFilter = { ...sheet.autoFilter };
  }

  // Copy freezePane
  if (sheet.freezePane) {
    cloned.freezePane = { ...sheet.freezePane };
  }

  // Deep copy images
  if (sheet.images) {
    cloned.images = sheet.images.map((img) => ({
      data: new Uint8Array(img.data),
      type: img.type,
      anchor: {
        from: { ...img.anchor.from },
        to: img.anchor.to ? { ...img.anchor.to } : undefined,
      },
      width: img.width,
      height: img.height,
    }));
  }

  // Copy protection
  if (sheet.protection) {
    cloned.protection = { ...sheet.protection };
  }

  // Copy pageSetup
  if (sheet.pageSetup) {
    cloned.pageSetup = {
      ...sheet.pageSetup,
      margins: sheet.pageSetup.margins ? { ...sheet.pageSetup.margins } : undefined,
    };
  }

  // Copy headerFooter
  if (sheet.headerFooter) {
    cloned.headerFooter = { ...sheet.headerFooter };
  }

  // Copy view
  if (sheet.view) {
    cloned.view = {
      ...sheet.view,
      tabColor: sheet.view.tabColor ? { ...sheet.view.tabColor } : undefined,
    };
  }

  // Copy hidden/veryHidden
  if (sheet.hidden !== undefined) cloned.hidden = sheet.hidden;
  if (sheet.veryHidden !== undefined) cloned.veryHidden = sheet.veryHidden;

  // Deep copy tables
  if (sheet.tables) {
    cloned.tables = sheet.tables.map((table) => ({
      ...table,
      columns: table.columns.map((col) => ({ ...col })),
    }));
  }

  return cloned;
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
  const cloned = cloneSheet(sourceSheet, newName ?? sourceSheet.name);
  targetWorkbook.sheets.push(cloned);
}

// ── Copy Range ──────────────────────────────────────────────────────

/**
 * Copy a range of cells from one location to another within the same sheet.
 * Copies values, styles, and merges.
 */
export function copyRange(
  sheet: Sheet,
  source: { startRow: number; startCol: number; endRow: number; endCol: number },
  target: { startRow: number; startCol: number },
): void {
  const rowCount = source.endRow - source.startRow + 1;
  const colCount = source.endCol - source.startCol + 1;

  // Ensure rows array is large enough for target
  const targetEndRow = target.startRow + rowCount - 1;
  while (sheet.rows.length <= targetEndRow) {
    sheet.rows.push([]);
  }

  // Read all source values and cells first (to handle overlapping ranges)
  const sourceValues: import("./_types").CellValue[][] = [];
  const sourceCells: (Cell | null)[][] = [];

  for (let r = 0; r < rowCount; r++) {
    sourceValues.push([]);
    sourceCells.push([]);
    for (let c = 0; c < colCount; c++) {
      const srcRow = source.startRow + r;
      const srcCol = source.startCol + c;

      // Read value
      const row = sheet.rows[srcRow];
      sourceValues[r].push(row && srcCol < row.length ? row[srcCol] : null);

      // Read cell
      if (sheet.cells) {
        const key = `${srcRow},${srcCol}`;
        const cell = sheet.cells.get(key);
        sourceCells[r].push(cell ? cloneCell(cell) : null);
      } else {
        sourceCells[r].push(null);
      }
    }
  }

  // Write values and cells to target
  for (let r = 0; r < rowCount; r++) {
    const tgtRow = target.startRow + r;
    const row = sheet.rows[tgtRow];

    for (let c = 0; c < colCount; c++) {
      const tgtCol = target.startCol + c;

      // Extend row if needed
      while (row.length <= tgtCol) row.push(null);
      row[tgtCol] = sourceValues[r][c];

      // Copy cell data
      const srcCell = sourceCells[r][c];
      if (srcCell) {
        if (!sheet.cells) sheet.cells = new Map();
        sheet.cells.set(`${tgtRow},${tgtCol}`, srcCell);
      } else if (sheet.cells) {
        sheet.cells.delete(`${tgtRow},${tgtCol}`);
      }
    }
  }

  // Copy merges that are fully within the source range
  if (sheet.merges) {
    const newMerges: MergeRange[] = [];
    for (const merge of sheet.merges) {
      if (
        merge.startRow >= source.startRow &&
        merge.endRow <= source.endRow &&
        merge.startCol >= source.startCol &&
        merge.endCol <= source.endCol
      ) {
        const rowOffset = target.startRow - source.startRow;
        const colOffset = target.startCol - source.startCol;
        newMerges.push({
          startRow: merge.startRow + rowOffset,
          startCol: merge.startCol + colOffset,
          endRow: merge.endRow + rowOffset,
          endCol: merge.endCol + colOffset,
        });
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
      );
      if (!exists) {
        sheet.merges.push(nm);
      }
    }
  }
}

// ── Move Sheet ──────────────────────────────────────────────────────

/**
 * Reorder sheets in a workbook.
 */
export function moveSheet(workbook: Workbook, fromIndex: number, toIndex: number): void {
  if (fromIndex === toIndex) return;
  const [sheet] = workbook.sheets.splice(fromIndex, 1);
  workbook.sheets.splice(toIndex, 0, sheet);
}

// ── Remove Sheet ────────────────────────────────────────────────────

/**
 * Remove a sheet from a workbook.
 */
export function removeSheet(workbook: Workbook, index: number): void {
  workbook.sheets.splice(index, 1);
  // Adjust activeSheet if needed
  if (workbook.activeSheet !== undefined) {
    if (workbook.activeSheet === index) {
      // If we removed the active sheet, set to the previous sheet or 0
      workbook.activeSheet =
        workbook.sheets.length > 0 ? Math.min(index, workbook.sheets.length - 1) : 0;
    } else if (workbook.activeSheet > index) {
      workbook.activeSheet--;
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
  predicate: CellValue | ((value: CellValue, row: number, col: number) => boolean),
): Array<{ row: number; col: number; value: CellValue }> {
  const results: Array<{ row: number; col: number; value: CellValue }> = [];
  const isFn = typeof predicate === "function";

  for (let r = 0; r < sheet.rows.length; r++) {
    const row = sheet.rows[r]!;
    for (let c = 0; c < row.length; c++) {
      const value = row[c] ?? null;
      const match = isFn
        ? (predicate as (value: CellValue, row: number, col: number) => boolean)(value, r, c)
        : value === predicate;
      if (match) {
        results.push({ row: r, col: c, value });
      }
    }
  }

  return results;
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
  let count = 0;

  for (let r = 0; r < sheet.rows.length; r++) {
    const row = sheet.rows[r]!;
    for (let c = 0; c < row.length; c++) {
      const value = row[c] ?? null;

      if (find instanceof RegExp) {
        // RegExp matching: only applies to string cells
        if (typeof value === "string" && find.test(value)) {
          if (typeof replace === "string") {
            // Reset lastIndex for global regexes
            find.lastIndex = 0;
            row[c] = value.replace(find, replace);
          } else {
            row[c] = replace;
          }
          // Reset lastIndex after test() for global regexes
          find.lastIndex = 0;
          count++;
        }
      } else {
        // Exact value matching
        if (value === find) {
          row[c] = replace;
          count++;
        }
      }
    }
  }

  return count;
}

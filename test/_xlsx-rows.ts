// ── The rows `readXlsx` has, as `streamXlsxRows` has to yield them ──
//
// Two test files compare the stream's row set with the buffered
// reader's, and each had its own way of reading the buffered side: one
// counted every key in sparse `Sheet.cells`, the other left out the
// entries that only carry a hyperlink or a comment. They agreed only
// because no hand-built case had a hyperlink. One definition, so the
// rule cannot drift between them.

import type { Cell } from "../src/_types"

/**
 * The 0-based rows of a sheet read with `sparse: true` that hold a cell
 * `readXlsx` counted toward the bounding box.
 *
 * In sparse mode every such cell lands in `Sheet.cells`, and nothing
 * else does — except the entries `readXlsx` adds afterwards to hang a
 * hyperlink or a comment on a position that had no `<c>`. Those come
 * from `<hyperlinks>` and the comments part, not from `sheetData`, so a
 * reader of `sheetData` rows — the stream — knows nothing about them,
 * and they are left out.
 */
export function bufferedRowIndexes(sheet: { cells?: Map<string, Cell> }): number[] {
  const rows = new Set<number>()
  for (const [key, cell] of sheet.cells ?? []) {
    const annotationOnly =
      cell.value === null && cell.formula === undefined && (cell.hyperlink || cell.comment)
    if (!annotationOnly) rows.add(Number(key.split(",")[0]))
  }
  return [...rows].sort((a, b) => a - b)
}

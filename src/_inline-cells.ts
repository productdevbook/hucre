// ── Cell objects written inline in `rows` ───────────────────────────
//
// `WriteSheet.rows` is the grid and `WriteSheet.cells` the per-cell
// detail, keyed `"row,col"`. Styling one cell therefore meant naming its
// position twice — once in the row, once in the map — and keeping the
// two in step by hand.
//
// The streaming writer never had that split: `addRow` has taken
// `{ value, style, formula }` inline since it existed. `writeOdsStream`
// takes the shape too, though only the `value` and `formula` of it —
// per-cell styles are the buffered ODS writer's alone.
//
// So the two XLSX writers disagreed about what a row entry may be, and
// the buffered one did not refuse the shape it did not accept —
// `resolveRows` read a cell object as a value, and the cell came out
// **empty**. Value, style and formula all gone, no error. See #433.
//
// Rather than teach every consumer of `rows` about a second shape — the
// two writers, the auto-width measurer, the pivot source collector, the
// table extent — an inline cell is lifted into `cells` once, before any
// of them runs. `rows` stays a grid of values, `cells` stays the one
// place per-cell detail lives, and an explicit `cells` entry still wins
// over an inline one at the same position.

import type { Cell, CellValue, WriteSheet } from "./_types"
import { isHyperlinkValue } from "./xlsx/hyperlink"

/**
 * A cell written where a value goes: `{ value, style }`, `{ formula }`,
 * or any other part of a {@link Cell}.
 */
export type InlineCell = Partial<Cell>

/**
 * Whether a row entry is a cell object rather than a value.
 *
 * `Date` is the only object a `CellValue` can be, and a
 * `HyperlinkValue` — `{ text, hyperlink }`, both strings — is the object
 * the `data[]` path already accepts in a value position. Everything else
 * is a cell object: not because the shape was inspected, but because
 * nothing else was ever a legal entry, so the alternative to reading it
 * as one is dropping it.
 */
export function isInlineCell(v: unknown): v is InlineCell {
  return (
    typeof v === "object" &&
    v !== null &&
    !(v instanceof Date) &&
    !Array.isArray(v) &&
    !isHyperlinkValue(v)
  )
}

/**
 * The value of a row entry, whichever shape it arrived in.
 *
 * A total function rather than a cast: a consumer that calls it is
 * correct on a sheet that went through {@link splitInlineCells} and on
 * one that did not, so the compiler is being told something true rather
 * than being overruled.
 */
export function toCellValue(v: CellValue | InlineCell): CellValue {
  return isInlineCell(v) ? (v.value ?? null) : v
}

/**
 * {@link toCellValue} over a grid, without copying one that is already
 * all values — which is every grid a caller wrote before #433, and most
 * of them since. The scan is one `typeof` per cell; a 100,000 × 12 sheet
 * that a CSV or JSON writer is about to walk anyway is not worth
 * duplicating to satisfy a type.
 */
export function toCellValues(rows: Array<Array<CellValue | InlineCell>>): CellValue[][] {
  for (const row of rows) {
    for (const v of row) {
      if (isInlineCell(v)) return rows.map((r) => r.map(toCellValue))
    }
  }
  return rows as CellValue[][]
}

/**
 * Lift any inline cell objects out of `sheet.rows` into `sheet.cells`.
 *
 * Returns the sheet **unchanged** when there are none, which is the
 * usual case — the scan is one `typeof` per cell and allocates nothing
 * until it finds something. A sheet that does carry them is copied
 * shallowly; the caller's arrays and map are never mutated.
 */
export function splitInlineCells(sheet: WriteSheet): WriteSheet {
  const rows = sheet.rows
  if (!rows) return sheet

  let found = false
  for (const row of rows) {
    for (const v of row) {
      if (isInlineCell(v)) {
        found = true
        break
      }
    }
    if (found) break
  }
  if (!found) return sheet

  const plainRows: CellValue[][] = []
  const lifted = new Map<string, Partial<Cell>>()

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!
    const plain: CellValue[] = new Array(row.length)
    for (let c = 0; c < row.length; c++) {
      const v = row[c]
      if (isInlineCell(v)) {
        lifted.set(`${r},${c}`, v)
        // The value stays in the grid too, so everything that reads only
        // `rows` — auto-width, a pivot's source range, a table's extent —
        // sees the cell rather than a hole.
        plain[c] = v.value ?? null
      } else {
        plain[c] = v as CellValue
      }
    }
    plainRows[r] = plain
  }

  // The caller's own `cells` is applied second, so where both describe a
  // position the explicit map wins — the same precedence `cells` already
  // has over `rows`.
  if (sheet.cells) {
    for (const [key, cell] of sheet.cells) lifted.set(key, cell)
  }

  return { ...sheet, rows: plainRows, cells: lifted }
}

/** {@link splitInlineCells} over a workbook's sheets. */
export function splitInlineCellsInSheets(sheets: WriteSheet[]): WriteSheet[] {
  let changed = false
  const out = sheets.map((s) => {
    const next = splitInlineCells(s)
    if (next !== s) changed = true
    return next
  })
  return changed ? out : sheets
}

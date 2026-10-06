import { createCellStore, setCell, cellEntries, getCell } from "./cell-store"
// ── Cell objects written inline in `rows` ───────────────────────────
//
// `SheetInput.rows` is the grid and `SheetInput.cells` the per-cell
// detail, indexed by numeric coordinates. Styling one cell therefore meant
// naming its position twice — once in the row, once in the store — and
// keeping the two in step by hand.
//
// Every spreadsheet writer resolves CellInput through the same boundary,
// including explicit formula caches and link() shorthand. Buffered input
// lifts metadata into the numeric store; streaming keeps it inline until
// serialization. See #433 for the original disagreement between them.
//
// Rather than teach every consumer of `rows` about a second shape — the
// two writers, the auto-width measurer, the pivot source collector, the
// table extent — an inline cell is lifted into `cells` once, before any
// of them runs. `rows` stays a grid of values, `cells` stays the one
// place per-cell detail lives, and an explicit `cells` entry still wins
// over an inline one at the same position.

import { isCellError } from "./cell-error"
import type { Cell, CellInput, CellValue, HyperlinkValue, SheetInput } from "./_types"
import { isHyperlinkValue } from "./xlsx/hyperlink"
import { effectiveValue } from "./_sheet-values"

/**
 * A cell written where a value goes: `{ value, style }`, `{ formula }`,
 * or any other part of a {@link Cell}.
 */
export type InlineCell = Partial<Cell>

/** Undefined inherits; null is a deliberate value, including an empty cache. */
export function mergeDefined<T extends object>(base: T, override: Partial<T>): T {
  const merged = { ...base }
  for (const key of Object.keys(override) as Array<keyof T>) {
    const value = override[key]
    if (value !== undefined) merged[key] = value
  }
  return merged
}

/**
 * Whether a row entry is a cell object rather than a value.
 *
 * Dates and structured errors are scalar values. Every other supported
 * object is rich input: either Partial<Cell> or the link() shorthand.
 */
function isRichCell(v: unknown): v is InlineCell | HyperlinkValue {
  return (
    typeof v === "object" &&
    v !== null &&
    !(v instanceof Date) &&
    !Array.isArray(v) &&
    !isCellError(v)
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
export function toCellValue(v: CellInput): CellValue {
  if (!isRichCell(v)) return v
  return isHyperlinkValue(v) ? v.text : (v.value ?? null)
}

/** Preserve the full inline model; an explicit cache, including null, wins. */
export function resolveCellInput(input: CellInput): InlineCell & { value: CellValue } {
  if (!isRichCell(input)) return { value: input ?? null }
  if (isHyperlinkValue(input)) {
    const internal = input.hyperlink.startsWith("#")
    return {
      value: input.text,
      hyperlink: {
        target: internal ? "" : input.hyperlink,
        location: internal ? input.hyperlink.slice(1) : undefined,
        display: input.text,
        tooltip: input.tooltip,
      },
    }
  }
  const cell = { ...input, value: effectiveValue(input, null) }
  // Streaming historically used value as a formula cache. Carry that
  // fallback through the common boundary instead of rebuilding a smaller
  // cell in each adapter and losing explicit caches or future fields.
  if (cell.formula !== undefined && cell.formulaResult === undefined)
    cell.formulaResult = cell.value
  return cell
}

/**
 * {@link toCellValue} over a grid, without copying one that is already
 * all values — which is every grid a caller wrote before #433, and most
 * of them since. The scan is one `typeof` per cell; a 100,000 × 12 sheet
 * that a CSV or JSON writer is about to walk anyway is not worth
 * duplicating to satisfy a type.
 */
export function toCellValues(rows: CellInput[][]): CellValue[][] {
  for (const row of rows) {
    for (const v of row) {
      if (isRichCell(v)) return rows.map((r) => r.map(toCellValue))
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
 * shallowly; the caller's arrays and store are never mutated.
 */
export function splitInlineCells<T extends SheetInput>(sheet: T): T {
  const rows = sheet.rows
  if (!rows) return sheet

  if (!rows.some((row) => row.some(isRichCell))) return sheet

  const plainRows: CellValue[][] = []
  const lifted = createCellStore<Partial<Cell>>()

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!
    const plain: CellValue[] = []
    for (let c = 0; c < row.length; c++) {
      const v = row[c]
      if (isRichCell(v)) {
        const cell = resolveCellInput(v)
        setCell(lifted, r, c, cell)
        // The value stays in the grid too, so everything that reads only
        // `rows` — auto-width, a pivot's source range, a table's extent —
        // sees the cell rather than a hole.
        plain.push(cell.value)
      } else {
        plain.push(v)
      }
    }
    plainRows[r] = plain
  }

  // The caller's own `cells` is applied second, so where both describe a
  // position the explicit store wins — the same precedence `cells` already
  // has over `rows`.
  if (sheet.cells) {
    for (const [row, col, cell] of cellEntries(sheet.cells)) {
      const inline = getCell(lifted, row, col)
      // A partial style override must not discard the row's formula/cache.
      setCell(lifted, row, col, inline ? mergeDefined(inline, cell) : cell)
    }
  }

  return { ...sheet, rows: plainRows, cells: lifted }
}

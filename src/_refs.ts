// ── Shifting references when rows or columns move ────────────────────
//
// Adjusting formula references is the *defining* behaviour of "insert a
// row" in a spreadsheet, and `sheet-ops` did none of it: a formula below
// an insertion still pointed where its arguments used to be. So did every
// range in a data validation, a conditional rule, an auto-filter, a
// sparkline. The cells moved and the references stayed. See #439 §D.
//
// The machinery was already here — `replaceA1Ranges` in cell-utils walks
// the references in a formula while leaving string literals, function
// names and embedded identifiers alone. It just was not wired to
// anything but the ODS writer.

import { colToLetter, letterToCol, replaceA1Ranges, type A1RangeMatch } from "./cell-utils"

/** What moved, from where, by how much. */
export interface RefShift {
  axis: "row" | "col"
  /** 0-based index at which the insertion or deletion starts. */
  at: number
  /**
   * Positive to insert, negative to delete. `delete` also means the
   * `|delta|` indexes from `at` are gone, and a reference into them
   * becomes `#REF!`.
   */
  delta: number
  /**
   * The sheet the operation happened on. A reference qualified with a
   * *different* sheet is left alone — `Sheet2!A5` does not move because
   * a row was inserted in Sheet1.
   */
  sheetName?: string
}

/** A parsed A1 reference, keeping the `$` of each coordinate. */
interface ParsedRef {
  colAbs: boolean
  col: number
  rowAbs: boolean
  row: number
}

const REF_PATTERN = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/i

function parseRef(ref: string): ParsedRef | null {
  const m = REF_PATTERN.exec(ref)
  if (!m) return null
  const row = Number(m[4]) - 1
  if (!Number.isInteger(row) || row < 0) return null
  return {
    colAbs: m[1] === "$",
    col: letterToCol(m[2]!.toUpperCase()),
    rowAbs: m[3] === "$",
    row,
  }
}

function formatRef(ref: ParsedRef): string {
  return `${ref.colAbs ? "$" : ""}${colToLetter(ref.col)}${ref.rowAbs ? "$" : ""}${ref.row + 1}`
}

const REF_ERROR = "#REF!"

/**
 * Where one coordinate lands after the shift.
 *
 * `null` means the row or column it names was deleted. Note that `$` does
 * not exempt a reference: absoluteness governs what happens when a
 * formula is *copied*, not when the grid moves underneath it, and Excel
 * shifts `$A$5` on an insert exactly as it shifts `A5`.
 */
function shiftCoordinate(value: number, shift: RefShift): number | null {
  if (shift.delta > 0) {
    return value >= shift.at ? value + shift.delta : value
  }
  const removed = -shift.delta
  if (value >= shift.at + removed) return value - removed
  if (value >= shift.at) return null
  return value
}

/** Which coordinate of a reference this shift moves. */
function coordinateOf(ref: ParsedRef, axis: "row" | "col"): number {
  return axis === "row" ? ref.row : ref.col
}

function withCoordinate(ref: ParsedRef, axis: "row" | "col", value: number): ParsedRef {
  return axis === "row" ? { ...ref, row: value } : { ...ref, col: value }
}

/**
 * Rewrite the references in a formula for a row or column that moved.
 *
 * Ranges are handled as ranges, not as two independent endpoints: a range
 * that a deletion swallows whole becomes `#REF!` once, rather than
 * `#REF!:#REF!`, and one the deletion clips simply shrinks.
 */
export function shiftFormula(formula: string, shift: RefShift): string {
  if (shift.delta === 0) return formula

  return replaceA1Ranges(formula, (match) => shiftMatch(match, shift))
}

/**
 * The qualifier, exactly as the formula spelled it.
 *
 * `replaceA1Ranges` strips the `!` and a leading `$` and hands back the
 * rest verbatim — so a name needing quotes still carries them, which is
 * what makes it a legal qualifier. This used to re-quote it, escaping the
 * quotes already there, and turned a reference to `My Sheet` into one
 * Excel rejects.
 */
function qualifierOf(match: A1RangeMatch): string {
  return match.sheet1 === undefined ? "" : `${match.sheet1}!`
}

/**
 * The sheet *name* a qualifier denotes, with Excel's quoting removed.
 *
 * Only for comparing against {@link RefShift.sheetName}, which is a name
 * rather than a qualifier. A quoted qualifier never equalled the bare
 * name, so a formula pointing at its own sheet that way was taken for a
 * foreign one and left unshifted while the rows moved under it.
 */
function unquoteSheet(name: string): string {
  return name.length >= 2 && name.startsWith("'") && name.endsWith("'")
    ? name.slice(1, -1).replace(/''/g, "'")
    : name
}

function shiftMatch(match: A1RangeMatch, shift: RefShift): string {
  // A reference into another sheet is not affected by this sheet's rows
  // moving. An unqualified one is local by definition.
  if (match.sheet1 !== undefined && unquoteSheet(match.sheet1) !== shift.sheetName) {
    return rebuild(match)
  }

  const start = parseRef(match.ref1)
  if (!start) return rebuild(match)

  if (match.ref2 === undefined) {
    const moved = shiftCoordinate(coordinateOf(start, shift.axis), shift)
    if (moved === null) return REF_ERROR
    return `${qualifierOf(match)}${formatRef(withCoordinate(start, shift.axis, moved))}`
  }

  const end = parseRef(match.ref2)
  if (!end) return rebuild(match)

  const startAt = coordinateOf(start, shift.axis)
  const endAt = coordinateOf(end, shift.axis)

  if (shift.delta < 0) {
    const removed = -shift.delta
    // The deletion covers the whole span, so there is nothing left to
    // point at.
    if (startAt >= shift.at && endAt < shift.at + removed) return REF_ERROR
  }

  // Clamp each endpoint into the surviving grid rather than letting it
  // become #REF!: a range clipped by a deletion shrinks, which is what
  // Excel does and what the caller means.
  const movedStart = shiftCoordinate(startAt, shift) ?? shift.at
  const movedEnd = shiftCoordinate(endAt, shift) ?? Math.max(shift.at - 1, 0)

  const lo = Math.min(movedStart, movedEnd)
  const hi = Math.max(movedStart, movedEnd)

  return (
    `${qualifierOf(match)}${formatRef(withCoordinate(start, shift.axis, lo))}` +
    `:${formatRef(withCoordinate(end, shift.axis, hi))}`
  )
}

function rebuild(match: A1RangeMatch): string {
  const head = `${qualifierOf(match)}${match.ref1}`
  if (match.ref2 === undefined) return head
  const tail = match.sheet2 === undefined ? match.ref2 : `${match.sheet2}!${match.ref2}`
  return `${head}:${tail}`
}

/**
 * Rewrite a bare A1 range — `"A1:D10"`, the shape `DataValidation.range`,
 * `ConditionalRule.range`, `AutoFilter.range` and `TableDefinition.range`
 * all use.
 *
 * Returns `undefined` when the whole range was deleted, so the caller can
 * drop the thing that owned it rather than keep a `#REF!` range.
 */
export function shiftRangeRef(range: string, shift: RefShift): string | undefined {
  const shifted = shiftFormula(range, shift)
  return shifted.includes(REF_ERROR) ? undefined : shifted
}

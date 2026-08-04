// ── Cell & Range Utility Functions ──────────────────────────────────
// Re-exports existing utilities and adds new ones.

import { parseCellRef } from "./xlsx/worksheet"
import { colToLetter } from "./xlsx/worksheet-writer"

// Re-export existing functions
export { parseCellRef } from "./xlsx/worksheet"
export { colToLetter, cellRef, rangeRef } from "./xlsx/worksheet-writer"

// ── New Utilities ──────────────────────────────────────────────────

/**
 * Matches an A1-style cell reference that is NOT part of a larger
 * identifier and NOT a function name. The negative lookbehind rejects a
 * preceding letter/digit/underscore (so the "G10" inside "LOG10" or the
 * tail of a defined name is left alone); the negative lookahead rejects a
 * following "(" (so "SUM(" / "ATAN2(" are not treated as the column "SUM"
 * etc.) and a following letter/digit (so "A1B" isn't split).
 */
const A1_REF_TOKEN = /(?<![A-Za-z0-9_$])(\$?[A-Z]{1,3}\$?\d+)(?![A-Za-z0-9_(])/g

/**
 * The sheet a reference is qualified with, written either Excel's way
 * (`Sheet2!A1`, `'My Sheet'!A1`) or with the dot separator OpenFormula and
 * LibreOffice's UI use (`Sheet2.A1`). A name needing quotes is quoted the
 * same way in both dialects — single quotes, an embedded quote doubled —
 * so the quoted form passes through untouched.
 */
const SHEET_QUALIFIER = "(?:\\$?(?:'(?:[^']|'')*'|[A-Za-z_][A-Za-z0-9_]*)[.!])?"

/** The reference itself, as matched inside {@link A1_RANGE_TOKEN}. */
const A1_REF = "\\$?[A-Z]{1,3}\\$?\\d+"

/**
 * Like {@link A1_REF_TOKEN} but also captures the sheet qualifier of each
 * reference and an optional `:ref2` range tail. The qualifier has to be part
 * of the match, not something the caller finds afterwards: a consumer that
 * only ever sees `A1` cannot tell `Sheet2!A1` from `A1`, and the ODS writer
 * used to prove it by emitting `Sheet2.[.A1]` — brackets around the cell but
 * not the sheet, which is not a reference ODF recognises. The lookbehind
 * also rejects a preceding `.`/`!` so that a qualifier this pattern failed
 * to parse can never be left outside the match.
 */
const A1_RANGE_TOKEN = new RegExp(
  `(?<![A-Za-z0-9_$.!])(${SHEET_QUALIFIER})(${A1_REF})(?::(${SHEET_QUALIFIER})(${A1_REF}))?(?![A-Za-z0-9_(])`,
  "g",
)

/** One A1 reference or range matched by {@link replaceA1Ranges}. */
export interface A1RangeMatch {
  /** Sheet qualifier of the first ref, separator stripped — `undefined` when the ref is local. */
  sheet1?: string
  ref1: string
  /** Sheet qualifier of the second ref; only a 3-D range such as `Sheet1!A1:Sheet3!B2` has one. */
  sheet2?: string
  /** Present only for a range. */
  ref2?: string
}

/** Strip the `.`/`!` separator and the `$` of an absolute sheet locator. */
function normalizeQualifier(raw: string): string | undefined {
  if (!raw) return undefined
  const name = raw.slice(0, -1)
  return name.startsWith("$") ? name.slice(1) : name
}

/**
 * Run a replacer over the A1 cell references / ranges in an Excel formula,
 * skipping quoted string literals, function names, and embedded
 * identifiers (same safety rules as {@link replaceA1Refs}). `replacer`
 * receives the reference's parts and returns the full replacement for the
 * matched span, sheet qualifier included.
 */
export function replaceA1Ranges(
  formula: string,
  replacer: (match: A1RangeMatch) => string,
): string {
  const parts = formula.split(/("(?:[^"]|"")*")/)
  for (let i = 0; i < parts.length; i++) {
    if (i % 2 === 1) continue
    parts[i] = parts[i].replace(
      A1_RANGE_TOKEN,
      (_m, q1: string, r1: string, q2?: string, r2?: string) =>
        replacer({
          sheet1: normalizeQualifier(q1),
          ref1: r1,
          sheet2: q2 ? normalizeQualifier(q2) : undefined,
          ref2: r2,
        }),
    )
  }
  return parts.join("")
}

/**
 * Run a replacer over the A1 cell references in an Excel formula while
 * leaving quoted string literals (`"..."`), function names, and embedded
 * identifiers untouched. The naive `/[A-Z]+\d+/g` approach corrupts
 * `LOG10(...)`, `ATAN2`, and `"AB1"` — this splits the formula on string
 * literals first and only rewrites refs in the code spans.
 *
 * `replacer` receives the matched reference text (e.g. `"$A$1"`) and
 * returns its replacement.
 */
export function replaceA1Refs(formula: string, replacer: (ref: string) => string): string {
  // Split on double-quoted string literals (Excel doubles "" to escape a
  // quote, which this regex keeps inside one literal token).
  const parts = formula.split(/("(?:[^"]|"")*")/)
  for (let i = 0; i < parts.length; i++) {
    // Odd indices are the captured string literals — leave them verbatim.
    if (i % 2 === 1) continue
    parts[i] = parts[i].replace(A1_REF_TOKEN, (m) => replacer(m))
  }
  return parts.join("")
}

/**
 * Convert a column letter (e.g. "A", "Z", "AA") to a 0-based column index.
 * This is the inverse of `colToLetter`.
 *
 *   "A" → 0, "Z" → 25, "AA" → 26, "ZZ" → 701
 */
export function letterToCol(letter: string): number {
  let col = 0
  for (let i = 0; i < letter.length; i++) {
    const code = letter.charCodeAt(i)
    // Support both uppercase and lowercase
    let value: number
    if (code >= 65 && code <= 90) {
      value = code - 64 // A=1, B=2, ...
    } else if (code >= 97 && code <= 122) {
      value = code - 96 // a=1, b=2, ...
    } else {
      break
    }
    col = col * 26 + value
  }
  return col - 1 // Convert to 0-based
}

/**
 * Parse a range string like "A1:D10" into 0-based coordinates.
 */
export function parseRange(range: string): {
  startRow: number
  startCol: number
  endRow: number
  endCol: number
} {
  const parts = range.split(":")
  const start = parseCellRef(parts[0]!)
  const end = parts.length > 1 ? parseCellRef(parts[1]!) : start
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  }
}

/**
 * Check if a cell (0-based row and col) falls within a range.
 */
export function isInRange(
  cellRow: number,
  cellCol: number,
  range: { startRow: number; startCol: number; endRow: number; endCol: number },
): boolean {
  return (
    cellRow >= range.startRow &&
    cellRow <= range.endRow &&
    cellCol >= range.startCol &&
    cellCol <= range.endCol
  )
}

// ── R1C1 Notation ────────────────────────────────────────────────

/**
 * Convert an R1C1-style formula reference to A1-style.
 *
 * - Absolute: `R2C3` → `$C$2`
 * - Relative: `R[1]C[-1]` (from row 5, col 5) → `D6`
 * - Mixed: `R2C[1]` → (absolute row 2, relative col)
 *
 * Replaces all R1C1 references in the formula string.
 */
export function r1c1ToA1(formula: string, currentRow: number, currentCol: number): string {
  // Match R1C1 patterns: R[n]C[n], RnCn, R[n]Cn, RnC[n]
  return formula.replace(
    /R(\[-?\d+\]|\d+)C(\[-?\d+\]|\d+)/g,
    (_match, rowPart: string, colPart: string) => {
      let row: number
      let col: number
      let rowAbs = true
      let colAbs = true

      if (rowPart.startsWith("[")) {
        // Relative row
        row = currentRow + parseInt(rowPart.slice(1, -1), 10)
        rowAbs = false
      } else {
        row = parseInt(rowPart, 10) - 1 // R1C1 is 1-based
      }

      if (colPart.startsWith("[")) {
        // Relative col
        col = currentCol + parseInt(colPart.slice(1, -1), 10)
        colAbs = false
      } else {
        col = parseInt(colPart, 10) - 1 // R1C1 is 1-based
      }

      const letter = colToLetter(col)
      const colStr = colAbs ? `$${letter}` : letter
      const rowStr = rowAbs ? `$${row + 1}` : `${row + 1}`
      return `${colStr}${rowStr}`
    },
  )
}

/**
 * Convert an A1-style cell reference to R1C1-style.
 *
 * - `$C$2` → `R2C3` (absolute)
 * - `D6` (from row 5, col 5) → `R[1]C[-1]` (relative)
 * - Mixed: `$C6` (from row 5) → `R[1]C3`
 *
 * Replaces all A1 references in the formula string.
 */
export function a1ToR1C1(formula: string, currentRow?: number, currentCol?: number): string {
  // Match A1 patterns: $A$1, A1, $A1, A$1, AA100 — but skip refs inside
  // function names ("LOG10") and string literals ("AB1").
  return replaceA1Refs(formula, (ref) => {
    const m = /(\$?)([A-Z]{1,3})(\$?)(\d+)/.exec(ref)!
    const colDollar = m[1]
    const colLetters = m[2]
    const rowDollar = m[3]
    const rowDigits = m[4]
    {
      const col = letterToCol(colLetters) // 0-based
      const row = parseInt(rowDigits, 10) - 1 // 0-based
      const colAbs = colDollar === "$"
      const rowAbs = rowDollar === "$"

      let rowPart: string
      if (rowAbs || currentRow === undefined) {
        rowPart = `${row + 1}` // 1-based absolute
      } else {
        const diff = row - currentRow
        rowPart = `[${diff}]`
      }

      let colPart: string
      if (colAbs || currentCol === undefined) {
        colPart = `${col + 1}` // 1-based absolute
      } else {
        const diff = col - currentCol
        colPart = `[${diff}]`
      }

      return `R${rowPart}C${colPart}`
    }
  })
}

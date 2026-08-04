// ── Write-path validation ────────────────────────────────────────────
// Every check hucre performs used to be on the read side. The write path
// trusted the caller completely, so an ordinary mistake — a sheet named
// after a date range, a name copied from a report title — produced a file
// Excel opens with "unreadable content" and no warning from hucre.
//
// These run before any bytes are produced, so a rejected workbook leaves
// no half-written output.

import { InvalidArgumentError } from "./errors"
import { MAX_COL_INDEX } from "./limits"

/**
 * Excel's hard limit on a sheet name. Enforced by Excel's UI, by the
 * XLSX format, and by LibreOffice for ODS.
 */
export const MAX_SHEET_NAME_LENGTH = 31

/**
 * Excel's limit on the text in one cell, and on a formula's length.
 *
 * Neither is enforced, and that is deliberate. Both are *application*
 * limits, not format limits: ECMA-376 imposes no such cap, the file
 * stays valid OOXML, and LibreOffice, pandas and hucre's own reader all
 * handle longer values. Excel truncates the display rather than refusing
 * the file. Throwing here would make hucre stricter than the format it
 * writes, and would break using it as a general interchange engine.
 *
 * Sheet names are different, and are enforced: an illegal one makes the
 * whole workbook unreadable rather than one cell lossy. See #364.
 *
 * Exported so a caller targeting Excel specifically can check.
 */
export const MAX_CELL_TEXT_LENGTH = 32_767
export const MAX_FORMULA_LENGTH = 8_192

/**
 * Characters Excel forbids in a sheet name. They collide with range
 * syntax (`Sheet1!A1`, `[Book1]Sheet1`) or with path separators.
 */
const FORBIDDEN_SHEET_NAME_CHARS = /[[\]:*?/\\]/

/**
 * Reserved by Excel for the change-tracking sheet. Case-insensitive —
 * Excel rejects `history` just as firmly as `History`.
 */
const RESERVED_SHEET_NAME = "history"

/**
 * Validate one sheet name against Excel's rules.
 *
 * Throws rather than sanitizing: silently truncating a 40-character name
 * or stripping its colons produces a workbook whose sheets are not the
 * ones the caller asked for, and range references built against the
 * original names would then dangle.
 */
export function validateSheetName(name: string, index: number): void {
  const where = `sheet ${index + 1}`

  if (typeof name !== "string" || name.length === 0) {
    throw new InvalidArgumentError(`Sheet name is empty (${where}); Excel requires a name`)
  }

  if (name.length > MAX_SHEET_NAME_LENGTH) {
    throw new InvalidArgumentError(
      `Sheet name "${name}" is ${name.length} characters (${where}); ` +
        `Excel allows at most ${MAX_SHEET_NAME_LENGTH}`,
    )
  }

  const forbidden = name.match(FORBIDDEN_SHEET_NAME_CHARS)
  if (forbidden) {
    throw new InvalidArgumentError(
      `Sheet name "${name}" contains ${JSON.stringify(forbidden[0])} (${where}); ` +
        `Excel forbids [ ] : * ? / \\ in sheet names`,
    )
  }

  // A leading or trailing apostrophe breaks quoted range references —
  // 'My Sheet'!A1 becomes unparseable.
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new InvalidArgumentError(
      `Sheet name "${name}" starts or ends with an apostrophe (${where}); ` +
        `Excel forbids it because it breaks quoted range references`,
    )
  }

  if (name.toLowerCase() === RESERVED_SHEET_NAME) {
    throw new InvalidArgumentError(
      `Sheet name "${name}" is reserved by Excel for change tracking (${where})`,
    )
  }
}

/**
 * Validate every sheet name in a workbook, including uniqueness.
 *
 * Excel compares sheet names case-insensitively, so `Data` and `data`
 * collide — a workbook carrying both opens as damaged.
 */
export function validateSheetNames(sheets: ReadonlyArray<{ name: string }>): void {
  const seen = new Map<string, number>()

  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i]!.name
    validateSheetName(name, i)

    const key = name.toLowerCase()
    const previous = seen.get(key)
    if (previous !== undefined) {
      throw new InvalidArgumentError(
        `Duplicate sheet name "${name}" (sheets ${previous + 1} and ${i + 1}); ` +
          `Excel compares sheet names case-insensitively`,
      )
    }
    seen.set(key, i)
  }
}

/**
 * Guard a 0-based column index before it becomes an `r=` attribute.
 *
 * `colToLetter` is pure arithmetic and produced nonsense for anything
 * outside the grid — `-1` gave `"@"`, `NaN` gave a NUL character, `1.5`
 * silently truncated, and `16384` gave `"XFE"`, one past Excel's last
 * column. Each produced a cell reference no reader can parse, from a
 * file that otherwise looked fine. See #364.
 */
export function validateColumnIndex(col: number): void {
  if (!Number.isInteger(col) || col < 0 || col > MAX_COL_INDEX) {
    throw new InvalidArgumentError(
      `Column index ${col} is not a valid 0-based column ` + `(Excel allows 0..${MAX_COL_INDEX})`,
    )
  }
}

// ── Value type inference ─────────────────────────────────────────────
// Text formats hand back strings, and every one of them has to answer the
// same question: does this look like a number, a boolean, a date? The CSV
// reader, the CSV streaming reader and the HTML table importer all ask it,
// and three private copies of the answer is three ways to disagree — the
// HTML importer used to run a bare `Number(text)` and turned "007" into 7
// while `parseCsv` on the same text returned "007". One implementation.

import type { CellValue } from "./_types"
import { reviveIsoDate } from "./_date"

/**
 * Infer a number, boolean or ISO date from cell text. Anything that is not
 * confidently one of those comes back unchanged, and non-strings pass
 * through untouched.
 *
 * `preserveLeadingZeros` keeps "0123", "007" and "00" as strings — the
 * product codes, ZIP codes and phone numbers that a bare `Number()` eats.
 */
export function inferType(value: CellValue, preserveLeadingZeros: boolean): CellValue {
  if (value === null) return null
  if (typeof value !== "string") return value

  const trimmed = value.trim()
  if (trimmed === "") return value

  // Boolean detection — only the literal true/false. "yes"/"no" are NOT
  // coerced: they collide with real data (the ISO country code "NO", a
  // yes/no/maybe survey column) and most CSV libraries don't coerce them.
  const lower = trimmed.toLowerCase()
  if (lower === "true") return true
  if (lower === "false") return false

  // ISO 8601 date detection, before number so a partial number cannot
  // swallow it. The rule itself lives in `_date.ts` — JSON revives dates
  // too, and one date rule is the whole point of this module.
  const asDate = reviveIsoDate(trimmed)
  if (asDate) return asDate

  // Leading-zero preservation: keep strings like "0123", "007", "00" as strings.
  // Exceptions: "0.xxx" decimals are still parsed.
  if (preserveLeadingZeros && trimmed.length > 1 && trimmed[0] === "0" && trimmed[1] !== ".") {
    return value
  }

  // Number detection
  const asNumber = parseNumber(trimmed)
  if (asNumber !== null) return asNumber

  return value
}

/**
 * Parse a decimal number, or return null when the text is not one.
 *
 * Deliberately stricter than `Number()`: the pattern below rejects hex
 * (`0x1A`), binary and octal literals, `Infinity` and `NaN`, all of which
 * `Number()` accepts and none of which a spreadsheet cell means.
 */
export function parseNumber(s: string): number | null {
  // Handle locale-aware numbers like "1,234.56" or "1,234"
  // Strip commas that are thousands separators (followed by 3 digits)
  const stripped = s.replace(/,(\d{3})/g, "$1")
  // Now try parsing
  if (stripped === "" || stripped === "-" || stripped === "+") return null
  // Must look like a number (avoid parsing random strings)
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(stripped)) return null
  const n = Number(stripped)
  if (Number.isNaN(n)) return null
  if (!Number.isFinite(n)) return null
  return n
}

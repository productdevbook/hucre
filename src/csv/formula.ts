// Formula-injection escaping, shared by every CSV writer and reader.
//
// It lives in its own module because the escape and its inverse have to
// agree on one list of trigger characters: the reader strips a leading
// apostrophe only where the writer would have added one (#408). Two copies
// of the list would drift, and a drifted list corrupts data in exactly the
// cases the escape exists to protect.

// Characters that trigger formula interpretation in Excel/Sheets/LibreOffice
// Covers: formulas (=), unary operators (+, -), at-sign (@), whitespace injection (\t, \r, \n), null byte (\0)
const FORMULA_PREFIXES = ["=", "+", "-", "@", "\t", "\r", "\n", "\0", "|"]

// DDE and dangerous function patterns (case-insensitive)
const DANGEROUS_PATTERNS = [
  /^=cmd\b/i,
  /^=HYPERLINK\s*\(/i,
  /^=IMPORTXML\s*\(/i,
  /^=IMPORTDATA\s*\(/i,
  /^=IMPORTFEED\s*\(/i,
  /^=IMPORTHTML\s*\(/i,
  /^=IMPORTRANGE\s*\(/i,
  /^=IMAGE\s*\(/i,
]

/**
 * Prefix a string value with a single quote if it starts with a formula-triggering character
 * or matches a dangerous function/DDE pattern.
 */
export function escapeFormula(value: string): string {
  if (value.length === 0) return value

  // Check prefix characters
  if (FORMULA_PREFIXES.includes(value[0]!)) {
    return "'" + value
  }

  // Check dangerous patterns (DDE, data exfiltration via HYPERLINK, etc.)
  for (const pattern of DANGEROUS_PATTERNS) {
    if (pattern.test(value)) {
      return "'" + value
    }
  }

  return value
}

/**
 * Undo {@link escapeFormula}: drop a leading apostrophe when the character
 * behind it is one the writer escapes for.
 *
 * The narrow test is the point. Stripping every leading apostrophe would
 * eat one from `'quoted'` and from any value a human typed that way; this
 * only touches shapes the writer can actually produce. Every dangerous
 * pattern starts with `=`, which is in {@link FORMULA_PREFIXES}, so the
 * prefix check covers those too.
 *
 * One ambiguity survives and cannot be removed: a source value that itself
 * began `'-5` is written unescaped (the apostrophe is not a trigger), and
 * un-escaping then reads it as `-5`. Documented on `unescapeFormulae`.
 */
export function unescapeFormula(value: string): string {
  if (value.length < 2 || value[0] !== "'") return value
  return FORMULA_PREFIXES.includes(value[1]!) ? value.slice(1) : value
}

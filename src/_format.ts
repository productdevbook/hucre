// ── Number Format Renderer ─────────────────────────────────────────────
//
// Apply an Excel number format string to a value and return formatted text.
// Handles: General, numbers, currency, percentage, scientific, date/time,
// fractions, accounting, multi-section formats, color codes, conditions,
// locale prefixes.
// ─────────────────────────────────────────────────────────────────────

import { isDateFormat, formatDate, serialToDate, dateToSerial } from "./_date"
import { InvalidArgumentError } from "./errors"

// ── Locale Definitions ──────────────────────────────────────────────

export interface LocaleFormat {
  /** Decimal separator character */
  decimal: string
  /** Thousands grouping separator character */
  thousands: string
  /** Currency symbol */
  currency: string
  /**
   * Digits per group, right to left: the first entry is the rightmost
   * group and the last repeats for everything further left.
   *
   * `[3]` for most locales. `[3, 2]` for the Indian system, where
   * 12,345,678 is written 1,23,45,678 — the separator was already right
   * there and the positions were not. Read from `Intl`, so any locale
   * that groups unusually is covered without a table. See #474.
   */
  groupSizes: number[]
}

/**
 * Currency symbols hucre has always carried for these four tags.
 *
 * The formatter does not read this field and never has — only `decimal`
 * and `thousands` are used — but `LocaleFormat` is public, so the values
 * it used to return are kept for the tags that had them.
 */
const KNOWN_CURRENCY: Record<string, string> = {
  "en-US": "$",
  "de-DE": "\u20AC",
  "fr-FR": "\u20AC",
  "tr-TR": "\u20BA",
}

const localeCache = new Map<string, LocaleFormat>()

/**
 * Resolve a BCP 47 tag to its separators.
 *
 * This used to be a four-entry table — `en-US`, `de-DE`, `fr-FR`,
 * `tr-TR` — and any other tag resolved to `undefined`, which the
 * formatter treated as "use the defaults". So
 * `formatValue(1234.5, "#,##0.00", { locale: "es-ES" })` returned
 * `1,234.50`: an en-US rendering, silently, for a locale the caller had
 * explicitly asked for. See #439 §R.
 *
 * `Intl.NumberFormat` knows every tag, and separators are all the
 * formatter needs, so the table is gone. A tag `Intl` rejects now throws
 * rather than being answered wrongly.
 */
function resolveLocale(locale?: string): LocaleFormat | undefined {
  if (!locale) return undefined

  const cached = localeCache.get(locale)
  if (cached) return cached

  let parts: Intl.NumberFormatPart[]
  try {
    parts = new Intl.NumberFormat(locale).formatToParts(1234567.8)
  } catch (error) {
    throw new InvalidArgumentError(
      `Unusable locale "${locale}". Pass a BCP 47 tag Intl.NumberFormat accepts, ` +
        "or omit `locale` for the default separators.",
      { cause: error },
    )
  }

  const resolved: LocaleFormat = {
    decimal: parts.find((p) => p.type === "decimal")?.value ?? ".",
    thousands: parts.find((p) => p.type === "group")?.value ?? ",",
    currency: KNOWN_CURRENCY[locale] ?? "",
    groupSizes: readGroupSizes(locale),
  }
  localeCache.set(locale, resolved)
  return resolved
}

/**
 * Ask `Intl` where a locale puts its group separators, by formatting a
 * number long enough to show three groups and measuring the digit runs.
 *
 * Deriving beats tabulating: `en-IN`, `hi-IN`, `bn-IN`, `ne-NP` and the
 * rest come out right without anyone listing them, and a locale whose
 * grouping `Intl` later revises follows along.
 */
function readGroupSizes(locale: string): number[] {
  let formatted: string
  try {
    formatted = new Intl.NumberFormat(locale, { useGrouping: true }).format(1234567890)
  } catch {
    return [3]
  }

  // Runs of digits, left to right; reverse so index 0 is the rightmost.
  const runs = (formatted.match(/\d+/g) ?? []).map((r) => r.length).reverse()
  if (runs.length < 2) return [3]

  // The leftmost run is whatever digits were left over, not a group size.
  const sizes = runs.slice(0, -1)
  // Collapse a uniform tail: [3, 3, 3] is [3], [3, 2, 2] is [3, 2].
  while (sizes.length > 1 && sizes[sizes.length - 1] === sizes[sizes.length - 2]) {
    sizes.pop()
  }
  return sizes.every((n) => n > 0) ? sizes : [3]
}

/**
 * Re-group an already-grouped integer string for a locale.
 *
 * The formatter groups in threes with `,` before this runs, which is
 * right for almost every locale and wrong for the Indian system. Anything
 * that is not plain digits and commas is left alone — a Special format
 * like `000-00-0000` interleaves literals, and re-grouping those digits
 * would be nonsense.
 */
function regroup(intStr: string, sizes: number[], separator: string): string {
  if (!/^[\d,]*$/.test(intStr)) return intStr.replace(/,/g, separator)

  const digits = intStr.replace(/,/g, "")
  if (digits.length === 0) return intStr

  const out: string[] = []
  let remaining = digits
  for (let i = 0; remaining.length > 0; i++) {
    const size = sizes[Math.min(i, sizes.length - 1)]!
    out.unshift(remaining.slice(-size))
    remaining = remaining.slice(0, -size)
  }
  return out.join(separator)
}

export interface FormatOptions {
  /** BCP 47 locale tag for number formatting (e.g. "de-DE", "tr-TR"). */
  locale?: string
  /**
   * Use the 1904 date system (Excel for Mac legacy). When true, date serials
   * are interpreted/produced relative to 1904-01-01 instead of 1900-01-01.
   */
  is1904?: boolean
}

/**
 * Apply an Excel number format string to a value and return formatted text.
 *
 * @param value - The raw cell value (number, string, boolean, Date)
 * @param numFmt - Excel number format string (e.g., "#,##0.00", "0%", "yyyy-mm-dd")
 * @param options - Optional formatting options (locale, etc.)
 * @returns Formatted string
 */
export function formatValue(value: unknown, numFmt: string, options?: FormatOptions): string {
  // Null/undefined → ""
  if (value === null || value === undefined) {
    return ""
  }

  // Boolean → "TRUE"/"FALSE"
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE"
  }

  // No format or "General"
  if (!numFmt || /^General$/i.test(numFmt.trim())) {
    if (value instanceof Date) {
      return value.toISOString()
    }
    return String(value)
  }

  // Parse sections: positive;negative;zero;text
  const sections = splitSections(numFmt)

  // If value is a string
  if (typeof value === "string") {
    // Use text section (4th) if available, otherwise return as-is
    const textSection = sections.length >= 4 ? sections[3] : sections[0]
    return applyTextSection(value, textSection)
  }

  // Convert Date to serial for numeric formatting
  let numValue: number
  if (value instanceof Date) {
    numValue = dateToSerial(value, options?.is1904)
  } else if (typeof value === "number") {
    numValue = value
  } else {
    return String(value)
  }

  // Select the right section based on value sign
  let section: string
  if (sections.length >= 3) {
    if (numValue > 0) {
      section = sections[0]
    } else if (numValue < 0) {
      section = sections[1]
      numValue = Math.abs(numValue) // negative section handles sign display
    } else {
      section = sections[2]
    }
  } else if (sections.length === 2) {
    if (numValue >= 0) {
      section = sections[0]
    } else {
      section = sections[1]
      numValue = Math.abs(numValue)
    }
  } else {
    section = sections[0]
    // For single-section, keep sign handling in the formatting
  }

  // Check for conditions in the section like [>100]
  const condResult = extractCondition(section)
  if (condResult.condition) {
    // With conditions, we use all sections but match against condition
    if (sections.length >= 2) {
      if (evaluateCondition(typeof value === "number" ? value : numValue, condResult.condition)) {
        section = condResult.rest
      } else {
        section = sections.length >= 2 ? stripCondition(sections[1]) : condResult.rest
      }
    } else {
      section = condResult.rest
    }
  }

  const localeInfo = resolveLocale(options?.locale)
  return applyNumberSection(numValue, section, localeInfo, options?.is1904)
}

// ── Section Parsing ─────────────────────────────────────────────────

/**
 * Split format string by unquoted semicolons.
 * Respects quoted strings and escaped characters.
 */
function splitSections(fmt: string): string[] {
  const sections: string[] = []
  let current = ""
  let inQuote = false
  let i = 0

  while (i < fmt.length) {
    const ch = fmt[i]

    if (ch === "\\") {
      current += ch
      i++
      if (i < fmt.length) {
        current += fmt[i]
        i++
      }
      continue
    }

    if (ch === '"') {
      inQuote = !inQuote
      current += ch
      i++
      continue
    }

    if (ch === ";" && !inQuote) {
      sections.push(current)
      current = ""
      i++
      continue
    }

    current += ch
    i++
  }

  sections.push(current)
  return sections
}

// ── Color & Locale Stripping ────────────────────────────────────────

/** Strip color codes like [Red], [Blue], [Color 3] etc. */
function stripColorCodes(fmt: string): string {
  return fmt.replace(/\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\s*\d+)\]/gi, "")
}

/** Strip locale prefixes like [$-409], [$€-407], [$-F800] */
function stripLocalePrefix(fmt: string): string {
  return fmt.replace(/\[\$[^\]]*\]/g, "")
}

/** Strip fill/padding characters like _( and *  */
function stripFillPadding(fmt: string): string {
  return fmt.replace(/[_*]./g, "")
}

/** Clean a format section: remove color, locale, padding */
function cleanSection(fmt: string): string {
  let cleaned = stripColorCodes(fmt)
  cleaned = stripLocalePrefix(cleaned)
  cleaned = stripFillPadding(cleaned)
  return cleaned
}

// ── Condition Handling ──────────────────────────────────────────────

interface Condition {
  operator: string
  value: number
}

function extractCondition(fmt: string): { condition: Condition | null; rest: string } {
  const match = fmt.match(/\[([<>=!]+)(-?\d+(?:\.\d+)?)\]/)
  if (!match) {
    return { condition: null, rest: fmt }
  }

  return {
    condition: { operator: match[1], value: Number(match[2]) },
    rest: fmt.replace(match[0], ""),
  }
}

function stripCondition(fmt: string): string {
  return fmt.replace(/\[([<>=!]+)(-?\d+(?:\.\d+)?)\]/, "")
}

function evaluateCondition(value: number, cond: Condition): boolean {
  switch (cond.operator) {
    case ">":
      return value > cond.value
    case "<":
      return value < cond.value
    case ">=":
      return value >= cond.value
    case "<=":
      return value <= cond.value
    case "=":
    case "==":
      return value === cond.value
    case "<>":
    case "!=":
      return value !== cond.value
    default:
      return true
  }
}

// ── Text Section ────────────────────────────────────────────────────

function applyTextSection(value: string, section: string): string {
  const cleaned = cleanSection(section)

  // Expand quoted strings and backslash-escaped chars, then replace @ with value
  const expanded = expandLiterals(cleaned)

  if (expanded.includes("@")) {
    return expanded.replace(/@/g, value)
  }

  // If no @ placeholder, return value as-is
  return value
}

/** Expand quoted strings ("text") and backslash escapes (\c) into literal text */
function expandLiterals(fmt: string): string {
  let result = ""
  let i = 0
  while (i < fmt.length) {
    if (fmt[i] === '"') {
      i++
      while (i < fmt.length && fmt[i] !== '"') {
        result += fmt[i]
        i++
      }
      i++ // skip closing quote
    } else if (fmt[i] === "\\") {
      i++
      if (i < fmt.length) {
        result += fmt[i]
        i++
      }
    } else {
      result += fmt[i]
      i++
    }
  }
  return result
}

// ── Number Section ──────────────────────────────────────────────────

function applyNumberSection(
  value: number,
  section: string,
  locale?: LocaleFormat,
  is1904?: boolean,
): string {
  const cleaned = cleanSection(section)

  // Text format: @ — return as string
  if (cleaned.trim() === "@") {
    return String(value)
  }

  // Check if it's a date format — delegate to formatDate
  if (isDateFormat(cleaned)) {
    const date = serialToDate(value, is1904)
    return formatDate(date, section, value) // Pass original section + serial for elapsed time
  }

  // Percentage: multiply by 100
  if (cleaned.includes("%")) {
    return formatPercentage(value, cleaned, locale)
  }

  // Scientific notation
  if (/[eE][+-]/.test(cleaned) || /[eE]\d/.test(cleaned)) {
    return formatScientific(value, cleaned, locale)
  }

  // Fractions
  if (isFractionFormat(cleaned)) {
    return formatFraction(value, cleaned)
  }

  // Regular number format
  return formatNumber(value, cleaned, locale)
}

// ── Percentage ──────────────────────────────────────────────────────

function formatPercentage(value: number, fmt: string, locale?: LocaleFormat): string {
  const percentValue = value * 100
  // Remove the % sign, format the number, then add % back
  const numFmt = fmt.replace(/%/g, "")
  const formatted = formatNumber(percentValue, numFmt, locale)
  return formatted + "%"
}

// ── Scientific Notation ─────────────────────────────────────────────

function formatScientific(value: number, fmt: string, locale?: LocaleFormat): string {
  // Parse the format: e.g., "0.00E+00"
  const match = fmt.match(/^([#0?.,]*?)([eE])([+-])(\d+)$/)
  if (!match) {
    // Fallback: determine decimal places from the mantissa part
    const decMatch = fmt.match(/\.([0#?]+)[eE]/)
    const decPlaces = decMatch ? decMatch[1].length : 2
    const expStr = value.toExponential(decPlaces)
    return formatExponentialString(expStr, fmt)
  }

  const mantissaFmt = match[1]
  const eChar = match[2]
  const signChar = match[3]
  const expDigits = match[4].length

  // Count decimal places in mantissa
  const dotIdx = mantissaFmt.indexOf(".")
  const decPlaces = dotIdx >= 0 ? mantissaFmt.length - dotIdx - 1 : 0

  const expStr = value.toExponential(decPlaces)
  const parts = expStr.split(/[eE]/)
  let mantissa = parts[0]
  let exp = Number.parseInt(parts[1], 10)

  // Apply locale decimal separator
  if (locale && locale.decimal !== ".") {
    mantissa = mantissa.replace(".", locale.decimal)
  }

  const expSign = exp >= 0 ? "+" : "-"
  const absExp = Math.abs(exp).toString().padStart(expDigits, "0")

  const displaySign = signChar === "+" ? expSign : exp < 0 ? "-" : ""

  return mantissa + eChar + displaySign + absExp
}

function formatExponentialString(expStr: string, fmt: string): string {
  const parts = expStr.split(/[eE]/)
  const mantissa = parts[0]
  let exp = Number.parseInt(parts[1], 10)

  // Determine E character case
  const eChar = fmt.includes("E") ? "E" : "e"
  const hasPlus = fmt.includes("E+") || fmt.includes("e+")

  const expSign = exp >= 0 ? "+" : "-"
  const absExp = Math.abs(exp).toString().padStart(2, "0")

  const displaySign = hasPlus ? expSign : exp < 0 ? "-" : ""

  return mantissa + eChar + displaySign + absExp
}

// ── Fraction Format ─────────────────────────────────────────────────

/** Numerator/denominator part of a fraction format; the denominator may be a literal number ("?/16"). */
const FRACTION_PARTS = /([?#0]+)\/(\d+|[?#0]+)/

function isFractionFormat(fmt: string): boolean {
  // A slash with digit placeholders against it: "# ?/?", "# ??/??", and the
  // fixed-denominator built-ins "As halves" (`# ?/2`), "As eighths" (`# ?/8`),
  // "As sixteenths" (`# ??/16`).
  //
  // One placeholder either side is enough — "?/?" and "0/2" are as much
  // fractions as "??/??" is. The old pattern wanted a placeholder *before*
  // the numerator run, so a one-character numerator never matched and the
  // format fell through to formatNumber, which renders the "/" as a literal
  // and loses the numerator: 2.5 under "?/?" came out " /3". See #402.
  //
  // Recognising and parsing off the same pattern keeps the two from
  // disagreeing about what a fraction is. Nothing else is dragged in: dates
  // are claimed before this is reached and put no placeholder against their
  // slashes, and an escaped slash ("0\/0") still carries its backslash.
  return FRACTION_PARTS.test(maskLiterals(fmt))
}

/**
 * Blank out quoted runs and escaped characters, keeping the string the
 * same length so an index into the result still points at the original.
 *
 * Fraction detection scans for placeholders around a slash, and a literal
 * can contain both: `0.00" 0/2"` was read as a fraction spec and rendered
 * `"3 1/2"` instead of `"3.50 0/2"`. See #429 — a regression from #402,
 * where unifying detection with the parse regex inherited the parse
 * regex's blind spot. The old pattern excluded this case by accident,
 * through the same quirk that made `"?/?"` unrecognisable.
 *
 * This shares its definition of "literal" with {@link extractLiterals} —
 * a quoted run, or a backslash and the character after it. A test pins
 * the two in agreement, because two implementations of one concept is
 * what caused #429 in the first place.
 */
function maskLiterals(fmt: string): string {
  let out = ""
  let i = 0

  while (i < fmt.length) {
    const ch = fmt[i]

    if (ch === '"') {
      const close = fmt.indexOf('"', i + 1)
      // An unterminated quote runs to the end, matching extractLiterals,
      // which consumes to the end rather than treating the quote as data.
      const end = close === -1 ? fmt.length : close + 1
      out += MASK.repeat(end - i)
      i = end
      continue
    }

    if (ch === "\\") {
      // The backslash and whatever it escapes, even at the very end.
      const span = i + 1 < fmt.length ? 2 : 1
      out += MASK.repeat(span)
      i += span
      continue
    }

    out += ch
    i++
  }

  return out
}

/** Stands in for a literal character: never a placeholder, never a slash. */
const MASK = "\u0001"

function formatFraction(value: number, fmt: string): string {
  // Determine denominator precision from format. Matched against the
  // masked form so a slash inside a literal cannot be mistaken for the
  // fraction bar; the mask preserves length, so the index and the
  // placeholder groups still describe the real format. See #429.
  const fracMatch = maskLiterals(fmt).match(FRACTION_PARTS)
  if (!fracMatch) {
    return String(value)
  }

  const fracAt = fracMatch.index ?? 0

  // Everything outside the numerator/denominator group is literal text, and
  // it is read with the machinery the plain number path already uses, so a
  // quoted run, a "\$" escape and a bare "$" mean here exactly what they mean
  // under "0.00". Building the output from the digits alone dropped all of it
  // — "$?/?" rendered as "5/2", and worse, the "-" of a negative section
  // ("# ?/?;-# ?/?") vanished with it, since that section is handed the
  // absolute value and the sign lives only in the format. See #426.
  //
  // Literals sit *outside* the "?" padding, which belongs to the fraction
  // rather than to the field: 2.5 under "$??/??" is "$ 5/ 2", not " $5/ 2".
  // The prefix leads the whole part and the suffix trails the denominator.
  // The prefix also leads the sign, matching formatNumber's "$-1,234.50".
  const head = extractLiterals(fmt.slice(0, fracAt))
  // The tail is all literal: nothing after the denominator has a number to
  // render, so expandLiterals — quotes and escapes to text — is the whole job.
  const tail = expandLiterals(fmt.slice(fracAt + fracMatch[0].length))

  // A whole part is only rendered when the format has a placeholder in front
  // of the fraction — "?" counts just as much as "#" and "0" ("? ?/?" is a
  // mixed number, "??/??" is improper). Reading it off `head.core` rather
  // than the raw text keeps a placeholder character that is only literal
  // ('"#"?/?') from being mistaken for an integer slot.
  const hasIntPart = /[#0?]/.test(head.core)

  const intPart = Math.trunc(value)
  // With no whole-part placeholder there is nowhere to put the integer, so
  // Excel folds it into the numerator: 2.5 under "??/??" is "5/2", not
  // "1/2". Formatting the remainder alone would drop the whole part
  // silently — a different number, not a different presentation. See #397.
  const target = hasIntPart ? Math.abs(value - intPart) : Math.abs(value)

  const denomLen = fracMatch[2].length

  // A denominator written in digits is fixed ("?/16"); one written in
  // placeholders is searched for. "0" and "00" are both at once, but a
  // literal denominator of zero means nothing, so they parse to 0 and the
  // guard below sends them to the search — which is where Excel puts them,
  // since "0" is a placeholder character.
  const fixedDenom = /^\d+$/.test(fracMatch[2]) ? Number.parseInt(fracMatch[2], 10) : 0

  let bestNum: number
  let bestDen: number

  if (target === 0) {
    bestNum = 0
    bestDen = fixedDenom > 0 ? fixedDenom : 1
  } else if (fixedDenom > 0) {
    bestDen = fixedDenom
    bestNum = Math.round(target * fixedDenom)
  } else {
    // Find best fraction with denominator up to 10^denomLen
    const maxDen = Math.pow(10, denomLen) - 1
    const result = findBestFraction(target, maxDen)
    bestNum = result.num
    bestDen = result.den
  }

  // Nothing left for the fraction area: either the value is whole, or the
  // remainder rounded away against a denominator the format fixed (0.1 over
  // halves). Excel prints the whole part rather than a zero numerator —
  // "3 0/2" is not something it ever renders. See #397.
  if (bestNum === 0) {
    // Whether there is an integer slot is the same question `hasIntPart`
    // already answered; asking it a second time off the raw format text is
    // how the two answers drift apart (a "$" prefix defeated the old test).
    // The blanked fraction area stays between the number and the suffix.
    if (hasIntPart && intPart !== 0) {
      return head.prefix + String(intPart) + tail
    }
    return head.prefix + String(intPart) + "      " + tail // padded like Excel
  }

  // Build the formatted string.
  // The sign has to come from the value: Math.trunc(-0.5) is -0, which is
  // neither `!== 0` nor `< 0`, so the sign would be lost for -1 < value < 0.
  // A format that writes its own "-" is left to it, as formatNumber does.
  const sign = value < 0 && !head.prefix.includes("-") ? "-" : ""
  // What separates the whole part from the numerator is whatever the format
  // put there — the space in "# ?/?" is literal text, not punctuation the
  // formatter owns.
  const whole = hasIntPart && intPart !== 0 ? String(Math.abs(intPart)) + head.suffix : ""

  const numStr = String(bestNum).padStart(fracMatch[1].length, " ")
  const denStr = String(bestDen).padStart(fracMatch[2].length, " ")

  return head.prefix + sign + whole + numStr + "/" + denStr + tail
}

function findBestFraction(value: number, maxDen: number): { num: number; den: number } {
  let bestNum = 0
  let bestDen = 1
  let bestError = Math.abs(value)

  for (let den = 1; den <= maxDen; den++) {
    const num = Math.round(value * den)
    const error = Math.abs(value - num / den)
    if (error < bestError) {
      bestError = error
      bestNum = num
      bestDen = den
      if (error === 0) break
    }
  }

  return { num: bestNum, den: bestDen }
}

// ── Number Formatting ───────────────────────────────────────────────

function formatNumber(value: number, fmt: string, locale?: LocaleFormat): string {
  // Extract currency symbol and literal text from the format
  const { prefix, suffix, core } = extractLiterals(fmt)

  if (!core.trim()) {
    // No number placeholders at all — return just the literal text
    return prefix + suffix
  }

  // A comma *between* digit placeholders turns on group separators; a comma
  // *after* the last placeholder scales the value down by 1000 each. The two
  // are independent — "#,##0,," both groups and divides by a million.
  const useThousandSep = /[#0?],[#0?]/.test(core)

  // Count trailing commas (each divides by 1000)
  const scaleMatch = core.match(/,+$/)
  const scaleDown = scaleMatch ? scaleMatch[0].length : 0

  let scaledValue = value
  for (let s = 0; s < scaleDown; s++) {
    scaledValue /= 1000
  }

  // Determine decimal places
  const dotIndex = core.indexOf(".")
  let decimalPlaces = 0
  if (dotIndex >= 0) {
    const afterDot = core.slice(dotIndex + 1).replace(/[^0#?]/g, "")
    decimalPlaces = afterDot.length
  }

  // Round the value
  const roundedValue = roundToDecimal(Math.abs(scaledValue), decimalPlaces)
  const isNegative = value < 0

  // Split into integer and decimal parts
  const [intStr, decStr] = splitNumber(roundedValue, decimalPlaces)

  // Format integer part
  const intFmt = dotIndex >= 0 ? core.slice(0, dotIndex) : core
  const formattedInt = formatIntegerPart(intStr, intFmt.replace(/,/g, ""), useThousandSep)

  // Format decimal part
  let formattedDec = ""
  if (dotIndex >= 0) {
    const decFmt = core.slice(dotIndex + 1)
    formattedDec = "." + formatDecimalPart(decStr, decFmt)
  }

  // Apply locale-specific separators if requested
  let localizedInt = formattedInt
  let localizedDec = formattedDec
  if (locale) {
    // Unconditional when the format asked for grouping: the separator and
    // the positions both come from the locale, and a locale that matches
    // the defaults regroups to the same string anyway.
    if (useThousandSep) {
      localizedInt = regroup(localizedInt, locale.groupSizes, locale.thousands)
    }
    if (locale.decimal !== "." && localizedDec.length > 0) {
      // Replace the leading "." with locale decimal
      localizedDec = locale.decimal + localizedDec.slice(1)
    }
  }

  // Combine
  let result = prefix
  if (isNegative && !prefix.includes("-")) {
    // Only add minus if the format doesn't already write one itself
    result += "-"
  }
  result += localizedInt + localizedDec + suffix

  return result
}

/**
 * Extract literal prefix/suffix text and the core number format.
 */
function extractLiterals(fmt: string): { prefix: string; suffix: string; core: string } {
  let prefix = ""
  let suffix = ""
  let core = ""
  let i = 0
  let foundDigitPlaceholder = false
  let afterDigits = false

  while (i < fmt.length) {
    const ch = fmt[i]

    // Quoted string
    if (ch === '"') {
      let literal = ""
      i++
      while (i < fmt.length && fmt[i] !== '"') {
        literal += fmt[i]
        i++
      }
      i++ // skip closing quote
      if (!foundDigitPlaceholder) {
        prefix += literal
      } else {
        afterDigits = true
        suffix += literal
      }
      continue
    }

    // Escaped character
    if (ch === "\\") {
      i++
      if (i < fmt.length) {
        if (!foundDigitPlaceholder) {
          prefix += fmt[i]
        } else {
          afterDigits = true
          suffix += fmt[i]
        }
        i++
      }
      continue
    }

    // Digit placeholders or format chars.
    // "+" and "-" are deliberately *not* here: Excel treats them as literal
    // text (the explicit sign of a negative section, for instance), so they
    // must reach the prefix/suffix rather than being swallowed by `core`,
    // where nothing would ever render them.
    if ("#0?.,%Ee".includes(ch)) {
      if (afterDigits && "#0?".includes(ch)) {
        // More digit placeholders after suffix text — unusual but handle it
        core += suffix + ch
        suffix = ""
        afterDigits = false
      } else {
        core += ch
      }
      if ("#0?".includes(ch)) {
        foundDigitPlaceholder = true
      }
      i++
      continue
    }

    // Comma within number section
    if (ch === ",") {
      if (foundDigitPlaceholder) {
        core += ch
      }
      i++
      continue
    }

    // Currency symbols and other characters
    if (!foundDigitPlaceholder) {
      prefix += ch
    } else {
      afterDigits = true
      suffix += ch
    }
    i++
  }

  return { prefix, suffix, core }
}

function roundToDecimal(value: number, decimals: number): number {
  const factor = Math.pow(10, decimals)
  return Math.round(value * factor) / factor
}

function splitNumber(value: number, decimalPlaces: number): [string, string] {
  const fixed = value.toFixed(decimalPlaces)
  const dotIdx = fixed.indexOf(".")
  if (dotIdx < 0) {
    return [fixed, ""]
  }
  return [fixed.slice(0, dotIdx), fixed.slice(dotIdx + 1)]
}

function formatIntegerPart(intStr: string, fmt: string, useThousandSep: boolean): string {
  // Count minimum digits from format (0s require digits, # are optional)
  const minDigits = (fmt.match(/0/g) || []).length
  const hasHash = fmt.includes("#")

  // If all # and value is 0, show nothing (e.g. "#.00" renders 0.5 as ".50")
  let digits = intStr
  if (digits === "0" && minDigits === 0 && hasHash) {
    digits = ""
  }

  // Plain run of digit placeholders — pad and group as a single block.
  if (!/[^0#?]/.test(fmt)) {
    let padded = digits
    if (padded.length < minDigits) {
      padded = padded.padStart(minDigits, "0")
    }
    if (useThousandSep && padded.length > 0) {
      padded = addThousandSeparators(padded)
    }
    return padded
  }

  // The format interleaves literal characters with digit placeholders —
  // Excel's Special formats ("000-00-0000", "(000) 000-0000"). Excel fills
  // placeholders right-to-left, keeps the literals in place, and lets the
  // leftmost placeholder absorb every surplus digit.
  const firstPlaceholder = fmt.search(/[0#?]/)
  const out: string[] = []
  let d = digits.length - 1

  for (let i = fmt.length - 1; i >= 0; i--) {
    const ch = fmt[i]
    if (ch !== "0" && ch !== "#" && ch !== "?") {
      out.push(ch)
      continue
    }
    if (d >= 0) {
      if (i === firstPlaceholder) {
        out.push(digits.slice(0, d + 1))
        d = -1
      } else {
        out.push(digits[d])
        d--
      }
    } else if (ch === "0") {
      out.push("0")
    } else if (ch === "?") {
      out.push(" ")
    }
  }

  return out.reverse().join("")
}

function formatDecimalPart(decStr: string, fmt: string): string {
  // The format contains 0, #, ? placeholders
  let result = ""
  const cleanFmt = fmt.replace(/[^0#?]/g, "")
  // `decStr` comes from toFixed(), so it is always padded out to the full
  // placeholder count. Trailing zeros there are insignificant digits, which
  // is what "?" renders as a space.
  const significant = decStr.replace(/0+$/, "").length

  for (let i = 0; i < cleanFmt.length; i++) {
    const placeholder = cleanFmt[i]
    const digit = i < decStr.length ? decStr[i] : "0"

    switch (placeholder) {
      case "0":
        // Always show digit
        result += digit
        break
      case "#":
        // Show digit only if significant (trailing zeros suppressed)
        // Check if there are any non-zero digits from this position onwards
        if (hasSignificantDigits(decStr, i)) {
          result += digit
        }
        break
      case "?":
        // Show digit or space
        if (i < significant) {
          result += digit
        } else {
          result += " "
        }
        break
    }
  }

  return result
}

function hasSignificantDigits(str: string, fromIndex: number): boolean {
  for (let i = fromIndex; i < str.length; i++) {
    if (str[i] !== "0") return true
  }
  return false
}

function addThousandSeparators(intStr: string): string {
  // Handle negative sign
  const negative = intStr.startsWith("-")
  const digits = negative ? intStr.slice(1) : intStr

  let result = ""
  const len = digits.length
  for (let i = 0; i < len; i++) {
    if (i > 0 && (len - i) % 3 === 0) {
      result += ","
    }
    result += digits[i]
  }

  return negative ? "-" + result : result
}

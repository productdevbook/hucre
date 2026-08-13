// ── ODS Writer ──────────────────────────────────────────────────────
// Generates valid OpenDocument Spreadsheet (.ods) files.

import type {
  WriteOptions,
  WriteOutput,
  CellValue,
  WorkbookProperties,
  WriteSheet,
  Cell,
  CellStyle,
  FontStyle,
  MergeRange,
  NamedRange,
  RichTextRun,
} from "../_types"
import { ZipWriter } from "../zip/writer"
import { validateSheetNames } from "../_validate"
import { unwrapCellValue } from "../xlsx/hyperlink"
import { xmlDocument, xmlElement, xmlSelfClose, xmlEscape as escapeXmlText } from "../xml/writer"
import { replaceA1Ranges, toRanges } from "../cell-utils"

const encoder = /* @__PURE__ */ new TextEncoder()

/**
 * Escape text content for an ODF document.
 *
 * The whole of this file — and the two streaming ODS writers, which
 * import this — goes through here rather than calling `xmlEscape`
 * directly, so that the carriage-return spelling is decided once.
 *
 * ODF has no `_xHHHH_` convention. Sharing the XLSX escaper meant a
 * string with a CR in it was written as `_x000D_` and read back with the
 * escape still in it — and LibreOffice showed the same seven characters,
 * because to anything that is not Excel that is all they are.
 */
export function odsEscape(text: string): string {
  return escapeXmlText(text, "charRef")
}

// ── ODS Namespaces ──────────────────────────────────────────────────

const NS_OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
const NS_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
const NS_TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
const NS_STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
const NS_FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
const NS_NUMBER = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
const NS_SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
const NS_META = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
const NS_DC = "http://purl.org/dc/elements/1.1/"
const NS_XLINK = "http://www.w3.org/1999/xlink"
const NS_OF = "urn:oasis:names:tc:opendocument:xmlns:of:1.2"

export const MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet"

// ── Helpers ─────────────────────────────────────────────────────────

/**
 * Format a number for display in <text:p>.
 * - Integers (including floats with no fractional part like 12.0) → "12"
 * - Floats → reasonable decimal places, no floating-point artifacts
 */
export function formatNumberDisplay(value: number): string {
  if (Number.isInteger(value)) return String(value)
  // Use toPrecision(15) to avoid floating-point artifacts (JS has ~17 significant digits),
  // then parseFloat to strip trailing zeros
  return String(parseFloat(value.toPrecision(15)))
}

function formatOdsDate(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, "Z")
}

export function formatOdsDateValue(date: Date): string {
  // ODS date values use ISO 8601 without time zone: YYYY-MM-DDTHH:MM:SS
  // Must use UTC methods to avoid local timezone offset corruption
  const y = date.getUTCFullYear()
  const m = String(date.getUTCMonth() + 1).padStart(2, "0")
  const d = String(date.getUTCDate()).padStart(2, "0")
  const hh = String(date.getUTCHours()).padStart(2, "0")
  const mm = String(date.getUTCMinutes()).padStart(2, "0")
  const ss = String(date.getUTCSeconds()).padStart(2, "0")
  return `${y}-${m}-${d}T${hh}:${mm}:${ss}`
}

// ── Number Format Translation ───────────────────────────────────────
//
// Translate an Excel-style format code (e.g. "0.00%", "yyyy-mm-dd",
// "[HH]:MM") into an ODS data-style element from the
// `urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0` namespace.
//
// ODS data styles are wrappers — the actual presentation lives in
// child `<number:number>`, `<number:year>`, `<number:hours>` etc
// elements that describe each token of the formatted output. A cell
// references a data style by name through `style:data-style-name` on
// its `<style:style>` element.

type OdsNumFmtKind = "number" | "percentage" | "currency" | "date" | "time" | "text"

interface OdsNumFmtDef {
  /** Element local name (no namespace prefix) */
  kind: OdsNumFmtKind
  /** Children of the data-style element, in document order */
  children: string[]
  /** Extra attributes on the data-style root (e.g. `number:truncate-on-overflow`) */
  attrs?: Record<string, string>
}

function isDateFormat(code: string): boolean {
  // Strip locale / colour tags / bracketed duration tokens so they don't
  // interfere — `[HH]`, `[red]`, `[$$-409]` are not date markers.
  const stripped = code.replace(/\[[^\]]*\]/g, "")
  if (/[yY]/.test(stripped) || /[dD]/.test(stripped)) return true

  // `mmm` and `mmmm` are the month-name formats and are always dates —
  // the same rule `isDateFormat` in `src/_date.ts` ends with, which every
  // reader uses to tell a date cell from a number. This copy did not have
  // it, so `mmm` reached the plain-number branch, found no `#` or `0` to
  // work with, and became literal text: a cell formatted `mmm` displayed
  // the letters "mmm" rather than "Mar".
  //
  // A lone `m` or `mm` stays out, in both. After an `h` it is minutes,
  // and there is nothing in the code itself to settle it.
  return /m{3,}/i.test(stripped)
}

function isTimeFormat(code: string): boolean {
  // Bracketed [h], [m], [s] always mean elapsed-time, otherwise the bare
  // letters h or s mark a clock-style time. Excel uses `m` for both month
  // and minute; we let the date check above win when only `m` is present.
  if (/\[[hHmMsS]+\]/.test(code)) return true
  if (/[hHsS]/.test(code)) return true
  return false
}

function isPercentageFormat(code: string): boolean {
  // Strip quoted literals so a literal "%" inside a string doesn't trigger
  return /(?<!\\)%/.test(code.replace(/"[^"]*"/g, "").replace(/\\./g, ""))
}

function detectCurrencySymbol(code: string): string | undefined {
  // [$<symbol>-<locale>] form, e.g. [$$-409], [$€-2], [$£-809]
  const bracketed = code.match(/\[\$([^-\]]*)(?:-[^\]]*)?\]/)
  if (bracketed && bracketed[1]) return bracketed[1]
  // "$" or "€" etc. as a quoted literal
  const quoted = code.match(/"([^"]+)"/)
  if (quoted && /[$€£¥₺₽₹]/.test(quoted[1])) return quoted[1]
  // Bare $ or another symbol, outside any bracketed section. Excel's
  // built-in date and time formats carry a locale tag like [$-F800], and
  // scanning the whole code found the `$` inside it — so the long-date
  // format became a currency style and every date token was lost. See
  // #390.
  const outsideBrackets = code.replace(/\[[^\]]*\]/g, "")
  const bare = outsideBrackets.match(/[$€£¥₺₽₹]/)
  if (bare) return bare[0]
  return undefined
}

function decimalsFromCode(code: string): number {
  // Find the first decimal section like "0.00" / "#.##" / "#,##0.000"
  const m = code.match(/[0#]\.([0#]+)/)
  return m ? m[1].length : 0
}

/**
 * How many digits a number format always shows, and how many at most.
 *
 * `0` is a digit that is always shown, `#` one shown only if there is
 * something to show — so `#.##` displays 1234.5 as "1234.5" and `0.00`
 * as "1234.50". ODF carries both counts: `number:decimal-places` is the
 * maximum and `number:min-decimal-places` the minimum, with
 * `number:min-integer-digits` doing the same job left of the point.
 *
 * #535 recorded the distinction as one ODS could not carry. It can; the
 * writer was simply emitting one count and a hardcoded `1`.
 */
function digitCounts(code: string): {
  minInteger: number
  minDecimals: number
  maxDecimals: number
} {
  // Only the first section — Excel's `positive;negative;zero` are
  // separate styles, and the caller has already split them.
  const section = code.split(";")[0] ?? ""
  const digits = section.match(/[0#?]+(?:,[0#?]+)*(?:\.[0#?]+)?/)?.[0] ?? ""
  const [integerPart = "", decimalPart = ""] = digits.split(".")

  return {
    // `?` is a digit-or-space placeholder; it holds a position, so it
    // counts as shown for this purpose.
    minInteger: (integerPart.match(/[0?]/g) ?? []).length,
    minDecimals: (decimalPart.match(/[0?]/g) ?? []).length,
    maxDecimals: decimalPart.length,
  }
}

function hasGrouping(code: string): boolean {
  // A comma between digit placeholders with three after it is the
  // thousands separator. Matching only `#,##0` and `0,000` missed
  // `#,###` — a wholly optional grouped integer — which then lost its
  // separator entirely and came back as `0`.
  //
  // The three placeholders matter: a trailing `0,,` is Excel's
  // scale-by-millions, not grouping, and has nothing after the commas.
  return /[0#?],[0#?]{3}/.test(code)
}

/**
 * A scientific format, and the widths ODF needs to describe it.
 *
 * Excel spells it `0.00E+00`: digits, then `E` with a sign, then the
 * exponent's own digits. ODF has `<number:scientific-number>` for this,
 * and without it the exponent is simply lost — the code fell through to
 * the plain-number branch, which reads the decimals and nothing else, so
 * `0.00E+00` became `0.00` and the file displayed a plain decimal in
 * LibreOffice as well.
 *
 * Quoted literals are stripped first: `"EUR "#,##0.00` has an `E` in it
 * and is not scientific.
 */
function parseScientificFormat(code: string):
  | {
      decimals: number
      integerDigits: number
      exponentDigits: number
      /** `E+` forces a sign on a positive exponent; `E-` does not. */
      forcedSign: boolean
      /** 3 for engineering notation (`##0.0E+0`), 1 otherwise. */
      interval: number
    }
  | undefined {
  const unquoted = code.replace(/"[^"]*"/g, "").replace(/\\./g, "")
  const match = unquoted.match(/^([#0]*)(?:\.([0#]+))?[Ee]([+-])([0#]+)$/)
  if (!match) return undefined

  // `min-integer-digits` counts the mandatory digits — the `0`s. `##0.0E+0`
  // asks for engineering notation, where the integer part steps in threes;
  // ODF expresses that with `number:exponent-interval`, which this does not
  // write, so such a code comes back as `0.0E+0`.
  const integerRun = match[1] ?? ""
  const integerDigits = (integerRun.match(/0/g) ?? []).length
  return {
    integerDigits: Math.max(integerDigits, 1),
    decimals: match[2]?.length ?? 0,
    exponentDigits: match[4]!.length,
    forcedSign: match[3] === "+",
    // `##0` asks for the integer part to stay between 1 and 999, which
    // is engineering notation: the exponent steps in threes. ODF says
    // that with `number:exponent-interval`, and the run's length is the
    // step. Anything else is a step of one.
    interval: integerRun.length > 1 ? integerRun.length : 1,
  }
}

/** Build a `<number:number>` child for numeric / percentage / currency styles */
function buildNumberChild(code: string, grouping: boolean): string {
  const { minInteger, minDecimals, maxDecimals } = digitCounts(code)
  const attrs: Record<string, string> = {
    "number:decimal-places": String(maxDecimals),
    "number:min-decimal-places": String(minDecimals),
    "number:min-integer-digits": String(minInteger),
  }
  if (grouping) attrs["number:grouping"] = "true"
  return xmlSelfClose("number:number", attrs)
}

/** Translate a date-format code into a sequence of `<number:*>` children */
function buildDateChildren(code: string): string[] {
  const out: string[] = []
  // Tokenise: longest matches first. Bracketed elapsed tokens like `[hh]`
  // collapse onto the same hour/minute/second elements but with style:long
  // when two letters are present.
  const tokenRegex =
    /\[[hH]{1,2}\]|\[[mM]{1,2}\]|\[[sS]{1,2}\]|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|am\/pm|"[^"]*"|\\.|./g

  // Heuristic: an `m` token is treated as minute when the previous non-literal
  // token is hours, otherwise as month.
  let lastWasHours = false
  let nextIsSeconds = false
  const tokens: string[] = []
  let match: RegExpExecArray | null
  while ((match = tokenRegex.exec(code)) !== null) {
    tokens.push(match[0])
  }

  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i]
    const lower = tok.toLowerCase()
    nextIsSeconds = false
    for (let j = i + 1; j < tokens.length; j++) {
      const t = tokens[j].toLowerCase()
      if (t === "s" || t === "ss") {
        nextIsSeconds = true
        break
      }
      if (
        t === "h" ||
        t === "hh" ||
        t === "m" ||
        t === "mm" ||
        t === "d" ||
        t === "dd" ||
        t === "yyyy" ||
        t === "yy"
      )
        break
    }

    // Bracketed elapsed tokens: `[h]`, `[hh]`, `[m]`, `[mm]`, `[s]`, `[ss]`.
    if (/^\[[hH]{1,2}\]$/.test(tok)) {
      const long = tok.length === 4 // `[hh]` -> 4 chars
      const attrs: Record<string, string> = {}
      if (long) attrs["number:style"] = "long"
      out.push(xmlSelfClose("number:hours", attrs))
      lastWasHours = true
      continue
    }
    if (/^\[[mM]{1,2}\]$/.test(tok)) {
      const long = tok.length === 4
      const attrs: Record<string, string> = {}
      if (long) attrs["number:style"] = "long"
      out.push(xmlSelfClose("number:minutes", attrs))
      lastWasHours = false
      continue
    }
    if (/^\[[sS]{1,2}\]$/.test(tok)) {
      const long = tok.length === 4
      const attrs: Record<string, string> = {}
      if (long) attrs["number:style"] = "long"
      out.push(xmlSelfClose("number:seconds", attrs))
      lastWasHours = false
      continue
    }

    if (lower === "yyyy") {
      out.push(xmlSelfClose("number:year", { "number:style": "long" }))
      lastWasHours = false
    } else if (lower === "yy") {
      out.push(xmlSelfClose("number:year"))
      lastWasHours = false
    } else if (lower === "mmmm") {
      out.push(xmlSelfClose("number:month", { "number:textual": "true", "number:style": "long" }))
      lastWasHours = false
    } else if (lower === "mmm") {
      out.push(xmlSelfClose("number:month", { "number:textual": "true" }))
      lastWasHours = false
    } else if (lower === "mm" || lower === "m") {
      // minute when between hours and seconds, else month
      const isMinute = lastWasHours || nextIsSeconds
      if (isMinute) {
        const attrs: Record<string, string> = {}
        if (lower === "mm") attrs["number:style"] = "long"
        out.push(xmlSelfClose("number:minutes", attrs))
      } else {
        const attrs: Record<string, string> = {}
        if (lower === "mm") attrs["number:style"] = "long"
        out.push(xmlSelfClose("number:month", attrs))
        lastWasHours = false
      }
    } else if (lower === "dddd") {
      out.push(xmlSelfClose("number:day-of-week", { "number:style": "long" }))
      lastWasHours = false
    } else if (lower === "ddd") {
      out.push(xmlSelfClose("number:day-of-week"))
      lastWasHours = false
    } else if (lower === "dd") {
      out.push(xmlSelfClose("number:day", { "number:style": "long" }))
      lastWasHours = false
    } else if (lower === "d") {
      out.push(xmlSelfClose("number:day"))
      lastWasHours = false
    } else if (lower === "hh") {
      out.push(xmlSelfClose("number:hours", { "number:style": "long" }))
      lastWasHours = true
    } else if (lower === "h") {
      out.push(xmlSelfClose("number:hours"))
      lastWasHours = true
    } else if (lower === "ss" || lower === "s") {
      // Excel spells fractional seconds by writing the places after the
      // token — `ss.0`, `ss.00`. ODF puts the count on the element. The
      // tokeniser hands `.` and each `0` over separately, so they are
      // consumed here; left alone they fell through to the literal-text
      // branch and `hh:mm:ss.0` came back as `hh:mm:ss."0"`, a format
      // that displays the digit rather than the fraction.
      const attrs: Record<string, string> = {}
      if (lower === "ss") attrs["number:style"] = "long"

      if (tokens[i + 1] === ".") {
        let j = i + 2
        while (tokens[j] === "0") j++
        const places = j - (i + 2)
        if (places > 0) {
          attrs["number:decimal-places"] = String(places)
          i = j - 1
        }
      }

      out.push(xmlSelfClose("number:seconds", attrs))
      lastWasHours = false
    } else if (lower === "am/pm") {
      out.push(xmlSelfClose("number:am-pm"))
      lastWasHours = false
    } else if (tok.startsWith('"') && tok.endsWith('"')) {
      out.push(xmlElement("number:text", undefined, odsEscape(tok.slice(1, -1))))
    } else if (tok.startsWith("\\") && tok.length === 2) {
      out.push(xmlElement("number:text", undefined, odsEscape(tok.slice(1))))
    } else {
      // Literal separator (`-`, `/`, `:`, `.`, ` `, etc.)
      out.push(xmlElement("number:text", undefined, odsEscape(tok)))
    }
  }

  return out
}

/**
 * Split an Excel format code into its `positive;negative;zero;text`
 * sections. A `;` inside a quoted literal or a bracketed tag is part of the
 * section, not a separator.
 */
function splitFormatSections(code: string): string[] {
  const sections: string[] = []
  let current = ""
  let quoted = false
  let bracketed = false

  for (let i = 0; i < code.length; i++) {
    const ch = code[i]!
    if (ch === "\\" && i + 1 < code.length) {
      current += ch + code[++i]
      continue
    }
    if (quoted) {
      current += ch
      if (ch === '"') quoted = false
      continue
    }
    if (ch === '"') quoted = true
    else if (ch === "[") bracketed = true
    else if (ch === "]") bracketed = false
    else if (ch === ";" && !bracketed) {
      sections.push(current)
      current = ""
      continue
    }
    current += ch
  }
  sections.push(current)
  return sections
}

/** Strip Excel's non-printing directives from a run of literal text. */
function unquoteLiteral(text: string): string {
  return (
    text
      // `_x` reserves the width of `x`; `*x` repeats `x` to fill the cell —
      // both are spacing hints with no ODF equivalent and no text of their own.
      .replace(/[_*]./g, "")
      .replace(/\\(.)/g, "$1")
      .replace(/"/g, "")
  )
}

/**
 * The literal text around the `#`/`0` core of a numeric section — the `-`
 * of `-#,##0.00`, the parentheses of `(#,##0.00)`, a trailing unit. Without
 * it a negative section renders unsigned, which matters because ODF reaches
 * that section through a `<style:map>` rather than by negating.
 */
function numberLiterals(code: string): { prefix: string; suffix: string } {
  const first = code.search(/[#0?]/)
  if (first < 0) return { prefix: unquoteLiteral(code), suffix: "" }
  let last = first
  for (let i = code.length - 1; i > first; i--) {
    if (/[#0?]/.test(code[i]!)) {
      last = i
      break
    }
  }
  return {
    prefix: unquoteLiteral(code.slice(0, first)),
    suffix: unquoteLiteral(code.slice(last + 1)),
  }
}

/**
 * Convert one section of an Excel-style `numFmt` code into an ODS
 * data-style definition. Returns `undefined` for sections the writer cannot
 * translate — those are silently dropped rather than emitting an invalid
 * style.
 */
function translateNumFmt(code: string): OdsNumFmtDef | undefined {
  if (!code) return undefined

  const trimmed = code.trim()
  // `General` is the *absence* of a format, and so is the empty string.
  if (trimmed === "General" || trimmed === "") return undefined

  // `@` is not. It says "show this as text", and ODF spells that
  // `<number:text-style>` with a `<number:text-content/>` placeholder —
  // LibreOffice writes one into every document it saves. #535 recorded
  // this as a loss ODS could not carry, which was wrong; the report in
  // #547 found the spelling by crossing the grammar with a LibreOffice
  // file.
  if (trimmed === "@") {
    return { kind: "text", children: [xmlSelfClose("number:text-content")] }
  }

  // Excel's built-in date and time codes carry a locale prefix —
  // `[$-409]mmm-yy`, `[$-F800]dddd, mmmm dd, yyyy`. It is a hint about
  // which locale's month and day names to use, not content, and ODF
  // expresses that through the style's language attributes rather than
  // inline. Left in place it was serialized character by character as
  // <number:text>, so the code round-tripped as `"["$"-"4""0""9""]"mmm-yy`.
  // A currency symbol written as `[$€-407]` keeps its symbol, because
  // detectCurrencySymbol reads that form before we get here. See #390.
  //
  // A colour tag (`[Red]`) is dropped for the same reason — and dropping it
  // is what keeps `[White]0.00` from being mistaken for a time format by
  // the bare `h` in "White".
  const section = trimmed
    .replace(/\[\$-[^\]]*\]/g, "")
    .replace(/\[(?:black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)\]/gi, "")

  // Before the currency and date checks: an exponent's `E` is neither,
  // but the plain-number branch at the bottom would swallow it silently.
  const scientific = parseScientificFormat(section)
  if (scientific) {
    return {
      kind: "number",
      children: [
        xmlSelfClose("number:scientific-number", {
          "number:decimal-places": String(scientific.decimals),
          "number:min-integer-digits": String(scientific.integerDigits),
          "number:min-exponent-digits": String(scientific.exponentDigits),
          ...(scientific.forcedSign ? { "number:forced-exponent-sign": "true" } : {}),
          ...(scientific.interval > 1
            ? { "number:exponent-interval": String(scientific.interval) }
            : {}),
        }),
      ],
    }
  }

  if (isPercentageFormat(section)) {
    const decimals = decimalsFromCode(section)
    const grouping = hasGrouping(section)
    const children = [
      buildNumberChild(section, grouping),
      xmlElement("number:text", undefined, "%"),
    ]
    return { kind: "percentage", children }
  }

  const currency = detectCurrencySymbol(section)
  if (currency) {
    const decimals = decimalsFromCode(section)
    const grouping = hasGrouping(section)
    const symbol = xmlElement("number:currency-symbol", undefined, odsEscape(currency))
    const number = buildNumberChild(section, grouping)
    // Detect symbol position: leading vs trailing
    const beforeNum = /^[^0#]*(\$|\[\$|"[$€£¥₺₽₹])/.test(section)
    const children = beforeNum ? [symbol, number] : [number, symbol]
    return { kind: "currency", children, attrs: { "number:automatic-order": "true" } }
  }

  // Quoted literals are text, not tokens. In `#,##0" days"` the `d` and
  // `y` are letters in a word; in `"Sales"#,##0` the `s` is. Read as
  // tokens they made the whole code a date or time format, and
  // `buildDateChildren` then emitted every `#` and `0` as
  // `<number:text>` — so `"Sales"#,##0` came back as
  // `"Sales""#","#""#""0"`, a format with no placeholder left in it.
  //
  // `isPercentageFormat` above already strips them, and the colour tags
  // stripped further up are the same problem found earlier: `[White]`
  // was being read as a time format for the `h` in "White".
  const tokens = section.replace(/"[^"]*"/g, "").replace(/\\./g, "")

  const time = isTimeFormat(tokens)
  const date = isDateFormat(tokens)
  // Bracketed-hour durations like `[HH]:MM` are pure time; check time first.
  const isElapsed = /\[[hHmMsS]+\]/.test(tokens)
  if (isElapsed || (time && !date)) {
    const def: OdsNumFmtDef = {
      kind: "time",
      children: buildDateChildren(section),
    }
    // `[H]`, `[M]`, `[S]` request a duration presentation that doesn't wrap
    // at 24h — ODS marks this with `number:truncate-on-overflow="false"`.
    if (isElapsed) def.attrs = { "number:truncate-on-overflow": "false" }
    return def
  }
  if (date) {
    return { kind: "date", children: buildDateChildren(section) }
  }

  // Plain number
  const { prefix, suffix } = numberLiterals(section)
  const children: string[] = []
  if (prefix) children.push(xmlElement("number:text", undefined, odsEscape(prefix)))
  // A section with no `#`/`0` at all is pure literal text — Excel's third
  // section is often `"-"` for zero — and gets no <number:number> child.
  if (/[#0?]/.test(section)) {
    children.push(buildNumberChild(section, hasGrouping(section)))
  } else if (children.length === 0) {
    return undefined
  }
  if (suffix) children.push(xmlElement("number:text", undefined, odsEscape(suffix)))
  return { kind: "number", children }
}

/**
 * `<table:named-expressions>`, or nothing when there is nothing to name.
 *
 * ODF spells a range `$Sheet1.$A$1:$Sheet1.$B$5` where Excel writes
 * `Sheet1!$A$1:$B$5`: the sheet is prefixed with `$` and joined by a dot,
 * and *both* halves carry it. A name whose range this cannot parse is
 * dropped rather than written malformed — `table:cell-range-address` is
 * required by the grammar, so a half-built one would invalidate the
 * document for the sake of a name nobody can resolve.
 *
 * Only workbook-level names. ODF scopes a name to a sheet by putting the
 * block inside that `<table:table>` instead, which `NamedRange.scope`
 * could drive; PARITY records that it does not yet.
 */
function buildNamedExpressions(ranges: NamedRange[] | undefined): string | undefined {
  if (!ranges || ranges.length === 0) return undefined

  const children: string[] = []
  for (const named of ranges) {
    const address = excelRangeToOdsAddress(named.range)
    if (!address) continue
    children.push(
      xmlSelfClose("table:named-range", {
        "table:name": named.name,
        "table:cell-range-address": address,
      }),
    )
  }

  if (children.length === 0) return undefined
  return xmlElement("table:named-expressions", undefined, children)
}

/** `Sheet1!$A$1:$B$2` -> `$Sheet1.$A$1:$Sheet1.$B$2`, or undefined. */
function excelRangeToOdsAddress(range: string): string | undefined {
  const match =
    /^(?:'([^']*(?:''[^']*)*)'|([A-Za-z_][\w.]*))!(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?$/.exec(
      range.trim(),
    )
  if (!match) return undefined

  const sheet = match[1] !== undefined ? match[1].replace(/''/g, "'") : match[2]!
  // The grammar's own regular expression for `cell-range-address` allows
  // a sheet name to be either a run with no dot, space or apostrophe in
  // it, or a quoted string. `$My Sheet.$A$1` is neither, and jing said so
  // — the round trip did not, because this reader is lenient about it.
  const quoted = /^[^.' ]+$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`
  const locator = odsSheetLocator(quoted)
  const start = `${locator}.${match[3]}`
  return match[4] ? `${start}:${locator}.${match[4]}` : start
}

// ── Style Generation ────────────────────────────────────────────────

/** Maps a CellStyle to a unique string key for deduplication */
function styleKey(style: CellStyle): string {
  const parts: string[] = []
  if (style.font?.bold) parts.push("b")
  if (style.font?.italic) parts.push("i")
  if (style.font?.size) parts.push(`sz${style.font.size}`)
  if (style.font?.color?.rgb) parts.push(`fc${style.font.color.rgb}`)
  if (style.fill?.type === "pattern" && style.fill.fgColor?.rgb) {
    parts.push(`bg${style.fill.fgColor.rgb}`)
  }
  if (style.numFmt) parts.push(`nf:${style.numFmt}`)
  return parts.join("|")
}

/**
 * The `<style:text-properties>` attributes a font maps onto. Shared by cell
 * styles and by the `text`-family styles rich-text runs reference, so a run
 * and a whole-cell font of the same description serialize identically.
 */
function fontTextProps(font: FontStyle | undefined): Record<string, string> {
  const props: Record<string, string> = {}
  if (!font) return props
  if (font.bold) props["fo:font-weight"] = "bold"
  if (font.italic) props["fo:font-style"] = "italic"
  if (font.size) props["fo:font-size"] = `${font.size}pt`
  if (font.color?.rgb) props["fo:color"] = `#${font.color.rgb}`
  return props
}

/** Generate a <style:style> element for a cell style */
function generateStyleElement(name: string, style: CellStyle, dataStyleName?: string): string {
  const textProps = fontTextProps(style.font)
  const cellProps: Record<string, string> = {}

  if (style.fill?.type === "pattern" && style.fill.fgColor?.rgb) {
    cellProps["fo:background-color"] = `#${style.fill.fgColor.rgb}`
  }

  // The order is the grammar's, not a preference. `<style:style>` with
  // `style:family="table-cell"` is a sequence — table-cell-properties,
  // then paragraph-properties, then text-properties — and writing the
  // text first produced a document the OASIS schema rejects for every
  // cell that had both a font and a fill. Nothing here noticed, because
  // every test reads the file back with hucre and LibreOffice opens it
  // regardless; `jing` against the published RELAX NG is what found it.
  const children: string[] = []
  if (Object.keys(cellProps).length > 0) {
    children.push(xmlSelfClose("style:table-cell-properties", cellProps))
  }
  if (Object.keys(textProps).length > 0) {
    children.push(xmlSelfClose("style:text-properties", textProps))
  }

  const attrs: Record<string, string> = { "style:name": name, "style:family": "table-cell" }
  if (dataStyleName) attrs["style:data-style-name"] = dataStyleName

  return xmlElement("style:style", attrs, children)
}

// ── Style Collector ────────────────────────────────────────────────

export interface StyleCollector {
  /** Map from style key → style name (e.g. "ce1") */
  styleMap: Map<string, string>
  /** Map from style name → XML element string */
  styleElements: Map<string, string>
  /** Counter for generating unique cell-style names */
  counter: number
  /** Map from numFmt code → data-style name (e.g. "N100") */
  dataStyleMap: Map<string, string>
  /** Map from data-style name → XML element string */
  dataStyleElements: Map<string, string>
  /** Counter for generating unique data-style names */
  dataStyleCounter: number
  /** Map from font key → text-style name (e.g. "T1"), for rich-text runs */
  textStyleMap: Map<string, string>
  /** Map from text-style name → XML element string */
  textStyleElements: Map<string, string>
  /** Counter for generating unique text-style names */
  textStyleCounter: number
}

export function createStyleCollector(): StyleCollector {
  return {
    styleMap: new Map(),
    styleElements: new Map(),
    counter: 1,
    dataStyleMap: new Map(),
    dataStyleElements: new Map(),
    dataStyleCounter: 100,
    textStyleMap: new Map(),
    textStyleElements: new Map(),
    textStyleCounter: 1,
  }
}

/**
 * A `text`-family style for one rich-text run. Runs reference it from
 * `<text:span text:style-name="T1">`; a run with nothing to say about its
 * font gets no style and is written as bare text.
 */
function getOrCreateTextStyleName(collector: StyleCollector, font: FontStyle): string {
  const props = fontTextProps(font)
  if (Object.keys(props).length === 0) return ""

  const key = styleKey({ font })
  const existing = collector.textStyleMap.get(key)
  if (existing) return existing

  const name = `T${collector.textStyleCounter++}`
  collector.textStyleMap.set(key, name)
  collector.textStyleElements.set(
    name,
    xmlElement(
      "style:style",
      { "style:name": name, "style:family": "text" },
      xmlSelfClose("style:text-properties", props),
    ),
  )
  return name
}

/** Emit one `<number:*-style>` element and return the name it was given. */
function emitDataStyle(
  collector: StyleCollector,
  def: OdsNumFmtDef,
  extraChildren: string[] = [],
  isVolatile = false,
): string {
  const name = `N${collector.dataStyleCounter++}`
  const attrs: Record<string, string> = { "style:name": name }
  // A style only reachable through another style's <style:map> is marked
  // volatile so consumers keep it even though no cell references it.
  if (isVolatile) attrs["style:volatile"] = "true"
  if (def.attrs) Object.assign(attrs, def.attrs)

  collector.dataStyleElements.set(
    name,
    xmlElement(`number:${def.kind}-style`, attrs, [...def.children, ...extraChildren]),
  )
  return name
}

function getOrCreateDataStyleName(collector: StyleCollector, numFmt: string): string | undefined {
  const existing = collector.dataStyleMap.get(numFmt)
  if (existing) return existing

  const sections = splitFormatSections(numFmt)
  const positive = translateNumFmt(sections[0]!)
  if (!positive) return undefined

  // Excel packs up to four sections into one code — positive;negative;zero;text
  // — while ODF gives each its own data style and links them from one
  // "main" style with <style:map>. Keeping only the first section threw the
  // rest away. See #405.
  //
  // The negative section is the main style and the others hang off it, which
  // is the shape LibreOffice writes: a style reached through a map renders
  // the value as-is, so the negative section has to carry its own minus.
  // The fourth (text) section has no ODF counterpart and is dropped.
  const negative = sections.length > 1 ? translateNumFmt(sections[1]!) : undefined
  const zero = sections.length > 2 ? translateNumFmt(sections[2]!) : undefined

  let name: string
  if (!negative) {
    name = emitDataStyle(collector, positive)
  } else {
    const maps = [
      xmlSelfClose("style:map", {
        "style:condition": "value()>0",
        "style:apply-style-name": emitDataStyle(collector, positive, [], true),
      }),
    ]
    if (zero) {
      maps.push(
        xmlSelfClose("style:map", {
          "style:condition": "value()=0",
          "style:apply-style-name": emitDataStyle(collector, zero, [], true),
        }),
      )
    }
    name = emitDataStyle(collector, negative, maps)
  }

  collector.dataStyleMap.set(numFmt, name)
  return name
}

function hasVisualProps(style: CellStyle): boolean {
  return Boolean(
    style.font?.bold ||
    style.font?.italic ||
    style.font?.size ||
    style.font?.color?.rgb ||
    (style.fill?.type === "pattern" && style.fill.fgColor?.rgb),
  )
}

export function getOrCreateStyleName(collector: StyleCollector, style: CellStyle): string {
  const key = styleKey(style)
  if (!key) return "" // No style properties

  const existing = collector.styleMap.get(key)
  if (existing) return existing

  const dataStyleName = style.numFmt ? getOrCreateDataStyleName(collector, style.numFmt) : undefined

  // If the only piece of styling is a numFmt that translated to nothing
  // (e.g. "General" / "@"), drop the cell-style entirely — emitting an
  // empty `<style:style>` would just bloat the document.
  if (!dataStyleName && !hasVisualProps(style)) return ""

  const name = `ce${collector.counter++}`
  collector.styleMap.set(key, name)
  collector.styleElements.set(name, generateStyleElement(name, style, dataStyleName))
  return name
}

// ── Formula Conversion ──────────────────────────────────────────────

/**
 * The sheet part of an OpenFormula reference. It lives *inside* the
 * brackets and carries `$` to mark the sheet absolute, the way LibreOffice
 * writes it: `[$Sheet2.A1]`. A local reference has no sheet part at all,
 * just the dot: `[.A1]`.
 */
function odsSheetLocator(sheet: string | undefined): string {
  return sheet ? `$${sheet}` : ""
}

/**
 * Convert an Excel-style formula to ODS formula syntax.
 * ODS formulas use `of:=` prefix and `[.A1]` cell references.
 */
export function excelFormulaToOds(formula: string): string {
  // Convert cell references like A1, $A$1, A1:B2 to ODS [.A1] notation,
  // handling ranges (A1:B2 → [.A1:.B2]) and cross-sheet references
  // (Sheet2!A1 → [$Sheet2.A1]) while leaving function names (LOG10),
  // string literals ("AB1"), and embedded identifiers untouched.
  //
  // The whole reference must be bracketed, sheet included — `Sheet2.[.A1]`
  // parses as nothing at all and LibreOffice refuses to evaluate the
  // formula. See #405.
  const converted = replaceA1Ranges(formula, ({ sheet1, ref1, sheet2, ref2 }) => {
    const start = `${odsSheetLocator(sheet1)}.${ref1}`
    // The second half of a range normally omits the sheet — it defaults to
    // the first half's — and only a 3-D range spells one out.
    return ref2 ? `[${start}:${odsSheetLocator(sheet2)}.${ref2}]` : `[${start}]`
  })
  return `of:=${converted}`
}

// ── Cell Serialization ──────────────────────────────────────────────

export interface CellContext {
  /** Cell override from sheet.cells */
  cellOverride?: Partial<Cell>
  /** Style name to apply (from style collector) */
  styleName?: string
  /** Merge span attributes */
  colSpan?: number
  rowSpan?: number
}

/** Serialize rich-text runs as `<text:span>`s, styled per run. */
function richTextSpans(runs: RichTextRun[], collector: StyleCollector): string {
  let out = ""
  for (const run of runs) {
    const text = odsEscape(run.text)
    const styleName = run.font ? getOrCreateTextStyleName(collector, run.font) : ""
    out += styleName ? xmlElement("text:span", { "text:style-name": styleName }, text) : text
  }
  return out
}

/**
 * The `<text:p>` holding a cell's visible content: the rich-text runs when
 * the cell has any, otherwise `display`.
 *
 * In ODF the anchor of a hyperlink *is* the cell's text, so a link on a
 * number, a date or a boolean is written exactly like one on a string —
 * the typed value stays in `office:value` / `office:date-value` beside it.
 * `Hyperlink.display` therefore replaces the text rather than sitting next
 * to it as it does in OOXML; a string cell whose display text differs from
 * its value reads back as the display text, because that is the only text
 * the format stores.
 */
function cellTextP(
  display: string,
  ctx: CellContext | undefined,
  collector: StyleCollector,
): string {
  const runs = ctx?.cellOverride?.richText
  const hyperlink = ctx?.cellOverride?.hyperlink

  let content = runs && runs.length > 0 ? richTextSpans(runs, collector) : odsEscape(display)
  if (hyperlink) {
    const anchor = hyperlink.display !== undefined ? odsEscape(hyperlink.display) : content
    content = xmlElement(
      "text:a",
      { "xlink:href": hyperlink.target, "xlink:type": "simple" },
      anchor,
    )
  }
  return xmlElement("text:p", undefined, content)
}

export function cellToOds(
  value: CellValue,
  ctx: CellContext | undefined,
  collector: StyleCollector,
): string {
  const attrs: Record<string, string> = {}
  const children: string[] = []

  if (ctx?.styleName) {
    attrs["table:style-name"] = ctx.styleName
  }
  if (ctx?.colSpan && ctx.colSpan > 1) {
    attrs["table:number-columns-spanned"] = String(ctx.colSpan)
  }
  if (ctx?.rowSpan && ctx.rowSpan > 1) {
    attrs["table:number-rows-spanned"] = String(ctx.rowSpan)
  }

  // Formula
  const formula = ctx?.cellOverride?.formula
  if (formula) {
    attrs["table:formula"] = excelFormulaToOds(formula)
  }

  // Hyperlink
  const hyperlink = ctx?.cellOverride?.hyperlink
  const richText = ctx?.cellOverride?.richText

  // A rich-text cell keeps its content in the runs and carries no scalar
  // value, so every branch below would have skipped it and written an empty
  // cell — the text simply vanished from content.xml. The runs win over
  // `value` the same way they do in the XLSX writer. See #405.
  if (richText && richText.length > 0) {
    attrs["office:value-type"] = "string"
    children.push(cellTextP("", ctx, collector))
    return xmlElement("table:table-cell", attrs, children)
  }

  if (value === null || value === undefined) {
    // A hyperlink with its own display text still has something to show in
    // an otherwise empty cell; a bare target has no anchor text to use.
    if (hyperlink?.display !== undefined) {
      attrs["office:value-type"] = "string"
      children.push(cellTextP("", ctx, collector))
      return xmlElement("table:table-cell", attrs, children)
    }
    if (Object.keys(attrs).length === 0) {
      return xmlSelfClose("table:table-cell")
    }
    return xmlElement("table:table-cell", attrs, children)
  }

  if (typeof value === "string") {
    attrs["office:value-type"] = "string"
    children.push(cellTextP(value, ctx, collector))
    return xmlElement("table:table-cell", attrs, children)
  }

  if (typeof value === "number") {
    // NaN / Infinity are not valid ODF floats — office:value="NaN" makes
    // LibreOffice read garbage. The XLSX writer already emits an empty
    // cell for these; ODS never got the same treatment. See #364.
    if (!Number.isFinite(value)) {
      return xmlElement("table:table-cell", attrs, [])
    }
    attrs["office:value-type"] = "float"
    attrs["office:value"] = String(value)
    children.push(cellTextP(formatNumberDisplay(value), ctx, collector))
    return xmlElement("table:table-cell", attrs, children)
  }

  if (typeof value === "boolean") {
    attrs["office:value-type"] = "boolean"
    attrs["office:boolean-value"] = value ? "true" : "false"
    children.push(cellTextP(value ? "TRUE" : "FALSE", ctx, collector))
    return xmlElement("table:table-cell", attrs, children)
  }

  if (value instanceof Date) {
    // An unparseable Date produced office:date-value="NaN-NaN-NaNT..."
    // — a corrupt file rather than an error. Emit an empty cell, the same
    // way non-finite numbers are handled. See #364.
    if (Number.isNaN(value.getTime())) {
      return xmlElement("table:table-cell", attrs, [])
    }
    const dateStr = formatOdsDateValue(value)
    attrs["office:value-type"] = "date"
    attrs["office:date-value"] = dateStr
    children.push(cellTextP(dateStr, ctx, collector))
    return xmlElement("table:table-cell", attrs, children)
  }

  if (Object.keys(attrs).length === 0) {
    return xmlSelfClose("table:table-cell")
  }
  return xmlElement("table:table-cell", attrs, children)
}

// ── Merge helpers ───────────────────────────────────────────────────

/** Build a set of covered cell positions from merge ranges */
function buildMergeMap(merges: MergeRange[] | undefined): {
  /** Cells that are the start of a merge: "row,col" → { colSpan, rowSpan } */
  starts: Map<string, { colSpan: number; rowSpan: number }>
  /** Cells covered by a merge (not the start cell) */
  covered: Set<string>
} {
  const starts = new Map<string, { colSpan: number; rowSpan: number }>()
  const covered = new Set<string>()

  if (!merges) return { starts, covered }

  for (const m of merges) {
    const colSpan = m.endCol - m.startCol + 1
    const rowSpan = m.endRow - m.startRow + 1
    starts.set(`${m.startRow},${m.startCol}`, { colSpan, rowSpan })

    for (let r = m.startRow; r <= m.endRow; r++) {
      for (let c = m.startCol; c <= m.endCol; c++) {
        if (r === m.startRow && c === m.startCol) continue
        covered.add(`${r},${c}`)
      }
    }
  }

  return { starts, covered }
}

// ── Row serialization with merge and cell override support ──────────

function rowToOds(
  row: CellValue[],
  rowIndex: number,
  sheet: WriteSheet,
  mergeMap: { starts: Map<string, { colSpan: number; rowSpan: number }>; covered: Set<string> },
  styleCollector: StyleCollector,
  maxCol: number,
): string {
  const cellElements: string[] = []

  // We need to emit cells for the full width including merge-covered columns
  const effectiveMax = Math.max(row.length - 1, maxCol)

  // Find the last column that has meaningful content (value, merge start, covered cell)
  let lastMeaningful = row.length - 1
  while (
    lastMeaningful >= 0 &&
    (row[lastMeaningful] === null || row[lastMeaningful] === undefined)
  ) {
    lastMeaningful--
  }
  // Also consider merge starts, covered cells, and per-cell overrides
  // beyond the last data value. An override carries a formula, a style or
  // a comment that the row array cannot express, so stopping at the last
  // non-null value dropped it silently — and the XLSX writer grows its
  // grid for exactly this case, so one WriteSheet produced different
  // documents per format. See #393.
  for (let c = lastMeaningful + 1; c <= effectiveMax; c++) {
    const key = `${rowIndex},${c}`
    if (mergeMap.starts.has(key) || mergeMap.covered.has(key) || sheet.cells?.has(key)) {
      lastMeaningful = c
    }
  }

  let i = 0
  while (i <= lastMeaningful) {
    const key = `${rowIndex},${i}`

    // Check if this cell is covered by a merge
    if (mergeMap.covered.has(key)) {
      // Count consecutive covered cells
      let count = 1
      while (i + count <= lastMeaningful && mergeMap.covered.has(`${rowIndex},${i + count}`)) {
        count++
      }
      if (count > 1) {
        cellElements.push(
          xmlSelfClose("table:covered-table-cell", {
            "table:number-columns-repeated": String(count),
          }),
        )
      } else {
        cellElements.push(xmlSelfClose("table:covered-table-cell"))
      }
      i += count
      continue
    }

    // Get cell override for values, formulas, hyperlinks, styles
    const cellOverride = sheet.cells?.get(key)

    // The override's value wins, matching resolveRows in the XLSX writer.
    // Reading only from `row` meant an override past the row's last value
    // serialized as an empty cell even once the grid had been grown to
    // reach it. See #393.
    const cell =
      cellOverride?.value !== undefined ? cellOverride.value : i < row.length ? row[i] : null

    // Build cell context
    const ctx: CellContext = {}
    if (cellOverride) ctx.cellOverride = cellOverride

    // Merge span
    const mergeInfo = mergeMap.starts.get(key)
    if (mergeInfo) {
      ctx.colSpan = mergeInfo.colSpan
      ctx.rowSpan = mergeInfo.rowSpan
    }

    // Style from cell override
    const style = cellOverride?.style
    if (style) {
      const name = getOrCreateStyleName(styleCollector, style)
      if (name) ctx.styleName = name
    }

    if (cell === null || cell === undefined) {
      if (Object.keys(ctx).length === 0 && !ctx.cellOverride && !mergeInfo) {
        // Plain empty cell — count consecutive empties
        let count = 1
        while (
          i + count <= lastMeaningful &&
          (i + count >= row.length || row[i + count] === null || row[i + count] === undefined) &&
          !mergeMap.covered.has(`${rowIndex},${i + count}`) &&
          !mergeMap.starts.has(`${rowIndex},${i + count}`) &&
          !sheet.cells?.has(`${rowIndex},${i + count}`)
        ) {
          count++
        }
        if (count > 1) {
          cellElements.push(
            xmlSelfClose("table:table-cell", {
              "table:number-columns-repeated": String(count),
            }),
          )
        } else {
          cellElements.push(xmlSelfClose("table:table-cell"))
        }
        i += count
        continue
      }
    }

    cellElements.push(cellToOds(cell, ctx, styleCollector))
    i++
  }

  return xmlElement("table:table-row", undefined, cellElements)
}

// ── content.xml ─────────────────────────────────────────────────────

function writeContentXml(options: WriteOptions): string {
  const { sheets } = options

  const styleCollector = createStyleCollector()
  const tableElements: string[] = []

  // First pass: collect styles and build table XML (deferred because styles go before body)
  const sheetXmlParts: string[][] = []

  for (const sheet of sheets) {
    const children: string[] = []

    // Resolve rows from rows or data
    let rows: CellValue[][] = []
    if (sheet.rows) {
      rows = sheet.rows
    } else if (sheet.data && sheet.columns) {
      // Generate header row + data rows from objects
      const keys = sheet.columns.map((c) => c.key ?? c.header ?? "")
      const hasHeaders = sheet.columns.some((c) => c.header)

      if (hasHeaders) {
        const headerRow = sheet.columns.map((c) => c.header ?? c.key ?? "")
        rows.push(headerRow)
      }

      for (const item of sheet.data) {
        const row = keys.map((k) => (k in item ? unwrapCellValue(item[k]) : null))
        rows.push(row)
      }
    }

    // Grow the grid to reach any per-cell override that sits past the
    // last row. The row loop below iterates `rows`, so an override at a
    // higher row index was never visited — the row-wise half of the same
    // defect as the trailing-column case in serializeRow. See #393.
    if (sheet.cells && sheet.cells.size > 0) {
      let maxOverrideRow = -1
      for (const key of sheet.cells.keys()) {
        const comma = key.indexOf(",")
        if (comma === -1) continue
        const r = Number(key.slice(0, comma))
        if (Number.isInteger(r) && r > maxOverrideRow) maxOverrideRow = r
      }
      if (maxOverrideRow >= rows.length) {
        if (rows === sheet.rows) rows = [...rows]
        while (rows.length <= maxOverrideRow) rows.push([])
      }
    }

    // A `WriteSheet` merge may be an A1 string; normalise once, here,
    // rather than at each of the three places below. See #474.
    const merges = toRanges(sheet.merges)

    // Build merge map
    const mergeMap = buildMergeMap(merges)

    // Determine column count (max width across all rows, considering merges)
    let colCount = 0
    for (const row of rows) {
      if (row.length > colCount) colCount = row.length
    }
    if (merges) {
      for (const m of merges) {
        if (m.endCol + 1 > colCount) colCount = m.endCol + 1
      }
    }

    // Determine max row needed (considering merges)
    let rowCount = rows.length
    if (merges) {
      for (const m of merges) {
        if (m.endRow + 1 > rowCount) rowCount = m.endRow + 1
      }
    }

    // Emit table:table-column element to declare column count
    if (colCount > 0) {
      if (colCount > 1) {
        children.push(
          xmlSelfClose("table:table-column", {
            "table:number-columns-repeated": String(colCount),
          }),
        )
      } else {
        children.push(xmlSelfClose("table:table-column"))
      }
    }

    // Emit rows (extend to cover merged rows beyond data)
    for (let r = 0; r < rowCount; r++) {
      const row = r < rows.length ? rows[r] : []
      children.push(rowToOds(row, r, sheet, mergeMap, styleCollector, colCount - 1))
    }

    sheetXmlParts.push(children)
  }

  // Now build the final XML with styles collected during serialization
  for (let i = 0; i < sheets.length; i++) {
    tableElements.push(
      xmlElement("table:table", { "table:name": sheets[i].name }, sheetXmlParts[i]),
    )
  }

  // `<table:named-expressions>` lives in the epilogue of
  // `<office:spreadsheet>` — after every `<table:table>`. The grammar is
  // a sequence, so anywhere else is invalid ODF.
  const namedExpressions = buildNamedExpressions(options.namedRanges)
  const spreadsheetBody = xmlElement("office:spreadsheet", undefined, [
    ...tableElements,
    ...(namedExpressions ? [namedExpressions] : []),
  ])
  const body = xmlElement("office:body", undefined, spreadsheetBody)

  // Build automatic styles from collected styles. Per ODS spec the data
  // styles (`<number:*-style>`) MUST appear before any `<style:style>`
  // that references them through `style:data-style-name`.
  const allStyleParts: string[] = [
    ...styleCollector.dataStyleElements.values(),
    ...styleCollector.styleElements.values(),
    ...styleCollector.textStyleElements.values(),
  ]
  const styleXml =
    allStyleParts.length > 0
      ? xmlElement("office:automatic-styles", undefined, allStyleParts)
      : xmlElement("office:automatic-styles", undefined, "")

  // Build content sections in order per ODS spec:
  // office:scripts, office:font-face-decls, office:automatic-styles, office:body
  const contentParts: string[] = []
  contentParts.push(xmlSelfClose("office:scripts"))
  contentParts.push(xmlElement("office:font-face-decls", undefined, ""))
  contentParts.push(styleXml)
  contentParts.push(body)

  return xmlDocument(
    "office:document-content",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:table": NS_TABLE,
      "xmlns:text": NS_TEXT,
      "xmlns:style": NS_STYLE,
      "xmlns:fo": NS_FO,
      "xmlns:number": NS_NUMBER,
      "xmlns:svg": NS_SVG,
      "xmlns:xlink": NS_XLINK,
      "xmlns:of": NS_OF,
      "office:version": "1.3",
    },
    contentParts,
  )
}

// ── meta.xml ────────────────────────────────────────────────────────

export function writeMetaXml(props?: WorkbookProperties): string {
  const children: string[] = []

  if (props?.title) {
    children.push(xmlElement("dc:title", undefined, odsEscape(props.title)))
  }
  if (props?.subject) {
    children.push(xmlElement("dc:subject", undefined, odsEscape(props.subject)))
  }
  if (props?.creator) {
    children.push(xmlElement("meta:initial-creator", undefined, odsEscape(props.creator)))
  }
  if (props?.description) {
    children.push(xmlElement("dc:description", undefined, odsEscape(props.description)))
  }
  if (props?.keywords) {
    children.push(xmlElement("meta:keyword", undefined, odsEscape(props.keywords)))
  }
  if (props?.created) {
    children.push(xmlElement("meta:creation-date", undefined, formatOdsDate(props.created)))
  }

  const modified = props?.modified ?? new Date()
  children.push(xmlElement("dc:date", undefined, formatOdsDate(modified)))

  children.push(xmlElement("meta:generator", undefined, "defter"))

  const metaContent = xmlElement("office:meta", undefined, children)

  return xmlDocument(
    "office:document-meta",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:meta": NS_META,
      "xmlns:dc": NS_DC,
      "office:version": "1.3",
    },
    metaContent,
  )
}

// ── styles.xml ──────────────────────────────────────────────────────

export function writeStylesXml(): string {
  // ODS spec requires these child elements even if empty:
  // office:font-face-decls, office:styles, office:automatic-styles, office:master-styles
  const children: string[] = []
  children.push(xmlElement("office:font-face-decls", undefined, ""))
  children.push(xmlElement("office:styles", undefined, ""))
  children.push(xmlElement("office:automatic-styles", undefined, ""))
  children.push(xmlElement("office:master-styles", undefined, ""))

  return xmlDocument(
    "office:document-styles",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:style": NS_STYLE,
      "xmlns:text": NS_TEXT,
      "xmlns:table": NS_TABLE,
      "xmlns:fo": NS_FO,
      "xmlns:number": NS_NUMBER,
      "xmlns:svg": NS_SVG,
      "office:version": "1.3",
    },
    children,
  )
}

// ── settings.xml ─────────────────────────────────────────────────────

export function writeSettingsXml(): string {
  const NS_CONFIG = "urn:oasis:names:tc:opendocument:xmlns:config:1.0"

  return xmlDocument(
    "office:document-settings",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:config": NS_CONFIG,
      "office:version": "1.3",
    },
    // No `<office:settings/>`. The element is optional, but the grammar
    // requires *one or more* `config:config-item-set` inside it, so an
    // empty one is invalid — and hucre has no view state, cursor
    // position or window geometry to record. The part itself stays,
    // because the manifest lists it and a reader may look for it.
    "",
  )
}

// ── manifest.xml ────────────────────────────────────────────────────

export function writeManifestXml(): string {
  const NS_MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"

  const entries: string[] = []
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "/",
      "manifest:version": "1.3",
      "manifest:media-type": MIMETYPE,
    }),
  )
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "content.xml",
      "manifest:media-type": "text/xml",
    }),
  )
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "meta.xml",
      "manifest:media-type": "text/xml",
    }),
  )
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "styles.xml",
      "manifest:media-type": "text/xml",
    }),
  )
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "settings.xml",
      "manifest:media-type": "text/xml",
    }),
  )

  return xmlDocument(
    "manifest:manifest",
    {
      "xmlns:manifest": NS_MANIFEST,
      "manifest:version": "1.3",
    },
    entries,
  )
}

// ── Main Writer ─────────────────────────────────────────────────────

/**
 * Write a workbook to ODS format.
 * Returns a Uint8Array containing the ZIP archive.
 */
export async function writeOds(options: WriteOptions): Promise<WriteOutput> {
  // Same rules as XLSX: LibreOffice enforces Excel's sheet-name limits
  // for interoperability. See #364.
  validateSheetNames(options.sheets)

  const zip = new ZipWriter()

  // mimetype MUST be the first entry and MUST be stored uncompressed
  zip.add("mimetype", encoder.encode(MIMETYPE), { compress: false })

  // META-INF/manifest.xml
  zip.add("META-INF/manifest.xml", encoder.encode(writeManifestXml()))

  // content.xml — main spreadsheet data
  zip.add("content.xml", encoder.encode(writeContentXml(options)))

  // meta.xml — document metadata
  zip.add("meta.xml", encoder.encode(writeMetaXml(options.properties)))

  // styles.xml — style definitions
  zip.add("styles.xml", encoder.encode(writeStylesXml()))

  // settings.xml — document settings
  zip.add("settings.xml", encoder.encode(writeSettingsXml()))

  return zip.build()
}

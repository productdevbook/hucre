// ── ODS Reader ──────────────────────────────────────────────────────
// Reads OpenDocument Spreadsheet (.ods) files.

import type {
  Workbook,
  ReadOptions,
  ReadInput,
  Sheet,
  CellValue,
  WorkbookProperties,
  Cell,
  CellStyle,
  MergeRange,
  Hyperlink,
} from "../_types"
import { ParseError, ZipError } from "../errors"
import { assertNotEncrypted, readInputToUint8Array } from "../_input"
import { ZipReader } from "../zip/reader"
import { decodePart } from "../_decode"
import { parseXml } from "../xml/parser"
import { parseRange } from "../cell-utils"
import { parseUtcDefaultDateTime } from "../_date"
import { MAX_COL_INDEX, MAX_REPEAT_COUNT, MAX_ROW_INDEX, MAX_TOTAL_CELLS } from "../limits"

/**
 * Bound a repeat count taken from the file so it can drive
 * `String.repeat` safely. Non-numeric and non-positive values collapse
 * to 1, matching the ODF default for an omitted `text:c`.
 */
function clampRepeat(raw: string | undefined): number {
  const value = Number(raw ?? "1")
  if (!Number.isFinite(value) || value < 1) return 1
  return Math.min(Math.trunc(value), MAX_REPEAT_COUNT)
}
import type { XmlElement } from "../xml/parser"

// ── Helpers ─────────────────────────────────────────────────────────

function decodeUtf8(data: Uint8Array, path = "(unknown)"): string {
  return decodePart(data, path)
}

function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const child of el.children) {
    if (typeof child === "string") continue
    const local = child.local || child.tag
    if (local === localName) return child
  }
  return undefined
}

function findChildren(el: XmlElement, localName: string): XmlElement[] {
  const result: XmlElement[] = []
  for (const child of el.children) {
    if (typeof child === "string") continue
    const local = child.local || child.tag
    if (local === localName) result.push(child)
  }
  return result
}

// ── Style Parsing ───────────────────────────────────────────────────

interface OdsStyleDef {
  bold?: boolean
  italic?: boolean
  fontSize?: number
  fontColor?: string // hex with '#' prefix
  backgroundColor?: string // hex with '#' prefix
  /** Excel-style number format reconstructed from a data style */
  numFmt?: string
}

function parseStyles(doc: XmlElement): Map<string, OdsStyleDef> {
  const styles = new Map<string, OdsStyleDef>()

  // Styles live in <office:automatic-styles>
  const autoStyles = findChild(doc, "automatic-styles")
  if (!autoStyles) return styles

  // First pass — collect data-style definitions (`<number:*-style>`) so we
  // can resolve `style:data-style-name` references in the second pass.
  const dataStyleMap = parseDataStyles(autoStyles)

  const styleElements = findChildren(autoStyles, "style")
  for (const styleEl of styleElements) {
    const family = styleEl.attrs["style:family"]
    if (family !== "table-cell") continue

    const name = styleEl.attrs["style:name"]
    if (!name) continue

    const def: OdsStyleDef = {}

    // Parse text properties
    const textProps = findChild(styleEl, "text-properties")
    if (textProps) {
      if (textProps.attrs["fo:font-weight"] === "bold") {
        def.bold = true
      }
      if (textProps.attrs["fo:font-style"] === "italic") {
        def.italic = true
      }
      const fontSize = textProps.attrs["fo:font-size"]
      if (fontSize) {
        // Parse "12pt" → 12
        const match = fontSize.match(/^(\d+(?:\.\d+)?)/)
        if (match) def.fontSize = parseFloat(match[1])
      }
      const color = textProps.attrs["fo:color"]
      if (color) {
        def.fontColor = color
      }
    }

    // Parse cell properties (background)
    const cellProps = findChild(styleEl, "table-cell-properties")
    if (cellProps) {
      const bgColor = cellProps.attrs["fo:background-color"]
      if (bgColor && bgColor !== "transparent") {
        def.backgroundColor = bgColor
      }
    }

    // Resolve `style:data-style-name` to an Excel-style format code
    const dataStyleRef = styleEl.attrs["style:data-style-name"]
    if (dataStyleRef) {
      const numFmt = dataStyleMap.get(dataStyleRef)
      if (numFmt) def.numFmt = numFmt
    }

    styles.set(name, def)
  }

  return styles
}

// ── Data-style (number format) parsing ──────────────────────────────

/**
 * Parse `<number:number-style>`, `<number:percentage-style>`,
 * `<number:date-style>`, `<number:time-style>`, and
 * `<number:currency-style>` elements into Excel-compatible format codes.
 */
function parseDataStyles(autoStyles: XmlElement): Map<string, string> {
  const out = new Map<string, string>()
  /** style name → the sections it maps to, by `<style:map>` condition */
  const mapped = new Map<string, Array<{ condition: string; target: string }>>()

  for (const child of autoStyles.children) {
    if (typeof child === "string") continue
    const local = child.local || child.tag
    const name = child.attrs["style:name"]
    if (!name) continue

    let code: string | undefined
    if (local === "number-style") {
      code = serializeDataStyleChildren(child, "number")
    } else if (local === "percentage-style") {
      code = serializeDataStyleChildren(child, "percentage")
    } else if (local === "currency-style") {
      code = serializeDataStyleChildren(child, "currency")
    } else if (local === "date-style") {
      code = serializeDataStyleChildren(child, "date")
    } else if (local === "text-style") {
      code = serializeDataStyleChildren(child, "text")
    } else if (local === "time-style") {
      const truncate = child.attrs["number:truncate-on-overflow"]
      code = serializeDataStyleChildren(child, "time", truncate === "false")
    }
    if (!code) continue
    out.set(name, code)

    const maps = findChildren(child, "map")
    if (maps.length > 0) {
      mapped.set(
        name,
        maps.map((m) => ({
          condition: m.attrs["style:condition"] ?? "",
          target: m.attrs["style:apply-style-name"] ?? "",
        })),
      )
    }
  }

  // Reassemble Excel's `positive;negative;zero` from the styles a
  // `<style:map>` chains together. The style a cell points at holds the
  // negative section, and the maps name the styles for the other two — the
  // inverse of what getOrCreateDataStyleName writes, and of what
  // LibreOffice writes.
  for (const [name, maps] of mapped) {
    let positive: string | undefined
    let zero: string | undefined
    for (const { condition, target } of maps) {
      const code = out.get(target)
      if (!code) continue
      if (condition.includes(">")) positive = code
      else if (condition.includes("=")) zero = code
    }
    // Without a section for the values the main style does not cover, the
    // maps describe something this reader cannot express — keep the main
    // style's own code rather than inventing sections around it.
    if (!positive) continue
    const negative = out.get(name)!
    out.set(name, zero ? `${positive};${negative};${zero}` : `${positive};${negative}`)
  }

  return out
}

/**
 * The integer half of a number format, from the digit counts ODF gives.
 *
 * `min-integer-digits="0"` means no digit is mandatory — `#` — and one
 * or more means that many `0`s. Grouping needs three positions before
 * the separator, so a single mandatory digit is `#,##0` and two is
 * `#,#00`.
 */
function integerPattern(minInteger: number, grouping: boolean): string {
  if (!grouping) return minInteger <= 0 ? "#" : "0".repeat(minInteger)
  if (minInteger <= 0) return "#,###"
  const mandatory = "0".repeat(minInteger)
  return minInteger >= 3 ? `#,${mandatory}` : `#,${"#".repeat(3 - minInteger)}${mandatory}`
}

function serializeDataStyleChildren(
  el: XmlElement,
  kind: "number" | "percentage" | "currency" | "date" | "time" | "text",
  bracketDuration = false,
): string {
  let out = ""
  for (const child of el.children) {
    if (typeof child === "string") continue
    const local = child.local || child.tag
    if (local === "number") {
      // Reached while parsing styles.xml, before any cell data.
      const clamp = (raw: string | undefined, fallback: number): number => {
        if (raw === undefined) return fallback
        const n = parseInt(raw, 10)
        return Number.isFinite(n) ? Math.min(Math.max(n, 0), MAX_REPEAT_COUNT) : fallback
      }

      const maxDecimals = clamp(child.attrs["number:decimal-places"], 0)
      // Absent means every decimal is shown, which is what this reader
      // assumed before it read the attribute at all — so a file written
      // by a tool that omits it reads exactly as it used to.
      const minDecimals = Math.min(
        clamp(child.attrs["number:min-decimal-places"], maxDecimals),
        maxDecimals,
      )
      const minInteger = clamp(child.attrs["number:min-integer-digits"], 1)
      const grouping = child.attrs["number:grouping"] === "true"

      // `0` is a digit always shown and `#` one shown only when there is
      // something to show, so the two counts are the difference between
      // `0.00` and `#.##`. See #535, which recorded this as a loss.
      out += integerPattern(minInteger, grouping)
      if (maxDecimals > 0) {
        out += `.${"0".repeat(minDecimals)}${"#".repeat(maxDecimals - minDecimals)}`
      }
    } else if (local === "scientific-number") {
      // `<number:scientific-number number:decimal-places="2"
      //  number:min-integer-digits="1" number:min-exponent-digits="2"/>`
      // is Excel's `0.00E+00`. The sign is always written: ODF has no
      // attribute for a bare `E00`, and Excel's own scientific formats
      // all carry one.
      const decimals = Math.min(
        parseInt(child.attrs["number:decimal-places"] ?? "0", 10) || 0,
        MAX_REPEAT_COUNT,
      )
      const integerDigits = Math.min(
        parseInt(child.attrs["number:min-integer-digits"] ?? "1", 10) || 1,
        MAX_REPEAT_COUNT,
      )
      const exponentDigits = Math.min(
        parseInt(child.attrs["number:min-exponent-digits"] ?? "2", 10) || 2,
        MAX_REPEAT_COUNT,
      )
      out +=
        "0".repeat(integerDigits) +
        (decimals > 0 ? `.${"0".repeat(decimals)}` : "") +
        `E+${"0".repeat(exponentDigits)}`
    } else if (local === "text-content") {
      // The placeholder for the cell's own text — Excel's `@`.
      out += "@"
    } else if (local === "currency-symbol") {
      const text = child.children.filter((c: unknown) => typeof c === "string").join("")
      out += `"${text}"`
    } else if (local === "text") {
      const text = child.children.filter((c: unknown) => typeof c === "string").join("")
      // Single-character separators stay bare; longer literals get quoted.
      if (text.length === 1 && /[\s\-/:.,()%]/.test(text)) {
        out += text
      } else {
        out += `"${text}"`
      }
    } else if (local === "year") {
      out += child.attrs["number:style"] === "long" ? "yyyy" : "yy"
    } else if (local === "month") {
      const long = child.attrs["number:style"] === "long"
      const textual = child.attrs["number:textual"] === "true"
      out += textual ? (long ? "mmmm" : "mmm") : long ? "mm" : "m"
    } else if (local === "day") {
      out += child.attrs["number:style"] === "long" ? "dd" : "d"
    } else if (local === "day-of-week") {
      out += child.attrs["number:style"] === "long" ? "dddd" : "ddd"
    } else if (local === "hours") {
      const tok = child.attrs["number:style"] === "long" ? "hh" : "h"
      out += bracketDuration && !out.includes("[") ? `[${tok}]` : tok
    } else if (local === "minutes") {
      out += child.attrs["number:style"] === "long" ? "mm" : "m"
    } else if (local === "seconds") {
      out += child.attrs["number:style"] === "long" ? "ss" : "s"
      // `number:decimal-places` on seconds is Excel's `ss.0` / `ss.00`.
      const places = Math.min(
        parseInt(child.attrs["number:decimal-places"] ?? "0", 10) || 0,
        MAX_REPEAT_COUNT,
      )
      if (places > 0) out += `.${"0".repeat(places)}`
    } else if (local === "am-pm") {
      out += "AM/PM"
    }
  }
  if (kind === "percentage" && !out.endsWith("%")) out += "%"
  // ODS time elements may emit elapsed-hour brackets via the writer; if the
  // reader detects truncate-on-overflow=false, surface the bracket form.
  return out
}

/** Convert a parsed ODS style def into a CellStyle */
function odsStyleToCellStyle(def: OdsStyleDef): CellStyle {
  const style: CellStyle = {}

  if (def.bold || def.italic || def.fontSize || def.fontColor) {
    style.font = {}
    if (def.bold) style.font.bold = true
    if (def.italic) style.font.italic = true
    if (def.fontSize) style.font.size = def.fontSize
    if (def.fontColor) {
      // Strip '#' prefix for the rgb field
      const hex = def.fontColor.startsWith("#") ? def.fontColor.slice(1) : def.fontColor
      style.font.color = { rgb: hex.toUpperCase() }
    }
  }

  if (def.backgroundColor) {
    const hex = def.backgroundColor.startsWith("#")
      ? def.backgroundColor.slice(1)
      : def.backgroundColor
    style.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { rgb: hex.toUpperCase() },
    }
  }

  if (def.numFmt) {
    style.numFmt = def.numFmt
  }

  return style
}

// ── Hyperlink Parsing ───────────────────────────────────────────────

/** Find the first <text:a> hyperlink anywhere under an element. */
function findHyperlink(el: XmlElement): { href: string; display: string } | undefined {
  for (const child of el.children) {
    if (typeof child === "string") continue
    const local = child.local || child.tag
    if (local === "a") {
      const href = child.attrs["xlink:href"]
      if (href) return { href, display: collectText(child) }
    }
    const nested = findHyperlink(child)
    if (nested) return nested
  }
  return undefined
}

/** Extract text and hyperlink from a cell's children */
function extractTextAndHyperlink(cell: XmlElement): { text: string; hyperlink?: Hyperlink } {
  // A cell may hold multiple <text:p> paragraphs (ODF joins them with a
  // newline). Collect the full text of every paragraph — including any text
  // surrounding a hyperlink — rather than only the first paragraph or only
  // the anchor text.
  const paragraphs = findChildren(cell, "p")
  if (paragraphs.length === 0) return { text: "" }

  const text = paragraphs.map((p) => collectText(p)).join("\n")

  // Surface the first hyperlink found in the cell, if any, without dropping
  // the surrounding text.
  for (const p of paragraphs) {
    const link = findHyperlink(p)
    if (link) {
      return { text, hyperlink: { target: link.href, display: link.display } }
    }
  }

  return { text }
}

/** Recursively collect text from an element and its children,
 *  handling ODS special elements: text:span, text:s, text:line-break, text:tab */
function collectText(el: XmlElement): string {
  let text = ""
  for (const child of el.children) {
    if (typeof child === "string") {
      text += child
    } else {
      const local = child.local || child.tag
      if (local === "s") {
        // <text:s/> or <text:s text:c="N"/> — space characters
        // Free-form integer from the file — uncapped it reaches a raw
        // RangeError, or allocates a gigabyte for one cell. See #363.
        const count = clampRepeat(child.attrs["text:c"])
        text += " ".repeat(count)
      } else if (local === "line-break") {
        // <text:line-break/> — newline
        text += "\n"
      } else if (local === "tab") {
        // <text:tab/> — tab character
        text += "\t"
      } else {
        // <text:span> and any other element — recurse into children
        text += collectText(child)
      }
    }
  }
  return text
}

// ── Formula Parsing ─────────────────────────────────────────────────

/** A cell, whole-column or whole-row address, as it appears after the dot. */
const ODS_ADDRESS = /^\$?(?:[A-Za-z]{1,3}(?:\$?\d+)?|\d+)$/

/**
 * Convert one side of an OpenFormula reference — everything between the
 * brackets, or between the brackets and the `:` — to its Excel spelling.
 * Returns `undefined` when the text is not an address, which is the signal
 * to leave the reference alone rather than mangle it.
 */
function odsAddressToExcel(part: string): string | undefined {
  // The separator is the *last* dot: a sheet name may contain one
  // (`['Q1.2024'.A1]`), an address never does.
  const dot = part.lastIndexOf(".")
  if (dot < 0) return undefined
  const address = part.slice(dot + 1)
  if (!ODS_ADDRESS.test(address)) return undefined

  let sheet = part.slice(0, dot)
  // An external reference (`['budget.ods'#$Sheet1.A1]`) has no Excel
  // spelling this reader can produce — leave the whole thing verbatim.
  if (sheet.includes("#")) return undefined
  // `$Sheet1` marks the sheet absolute; Excel has no notation for that, and
  // a cross-sheet reference never shifts on copy anyway.
  if (sheet.startsWith("$")) sheet = sheet.slice(1)

  return sheet ? `${sheet}!${address}` : address
}

/**
 * Convert an ODS formula to Excel-style formula.
 * ODS: "of:=SUM([.A1:.A10])" → "SUM(A1:A10)"
 */
function odsFormulaToExcel(formula: string): string {
  // Strip "of:=" or "oooc:=" prefix
  let f = formula
  if (f.startsWith("of:=")) f = f.slice(4)
  else if (f.startsWith("oooc:=")) f = f.slice(6)
  else if (f.startsWith("=")) f = f.slice(1)

  // Convert [.A1:.B2] → A1:B2, [.A1] → A1, and the cross-sheet forms
  // LibreOffice writes — [$Sheet2.A1] / [Sheet2.A1] → Sheet2!A1,
  // [$Sheet2.A1:.B2] → Sheet2!A1:B2. Matching on a literal `[.` (as this
  // did) decoded only the local forms and left every cross-sheet reference
  // in the file as raw ODF text. See #405.
  //
  // Split on string literals first, as the writer does, so a bracketed
  // token inside one is left as the text it is.
  const parts = f.split(/("(?:[^"]|"")*")/)
  for (let i = 0; i < parts.length; i++) {
    if (i % 2 === 1) continue
    parts[i] = parts[i]!.replace(/\[([^\]]*)\]/g, (match, body: string) => {
      // A sheet name cannot contain `:`, so this only ever splits a range.
      const halves = body.split(":")
      if (halves.length > 2) return match
      const first = odsAddressToExcel(halves[0]!)
      if (first === undefined) return match
      if (halves.length === 1) return first
      const second = odsAddressToExcel(halves[1]!)
      if (second === undefined) return match
      return `${first}:${second}`
    })
  }

  return parts.join("")
}

// ── Date Parsing ────────────────────────────────────────────────────

/**
 * Parse an ODF date (`xsd:date` / `xsd:dateTime`), whose zone designator is
 * optional.
 *
 * `new Date(text)` cannot be used directly: ECMAScript reads an unqualified
 * date-*time* as local but a date-*only* string as UTC. The writer builds
 * `office:date-value` out of UTC components, so the reader used to shift
 * every value by the machine's offset — and because the shifted value was
 * written back out the same way, the error accumulated with each save.
 * Silent on a UTC machine, one day off after four round trips in Tokyo.
 * See #415.
 *
 * An explicit offset (`...+02:00`, `...Z`) is what the file says and is
 * honoured; only an unqualified time is taken to mean UTC.
 */
export function parseOdsDateTime(text: string): Date | undefined {
  return parseUtcDefaultDateTime(text)
}

// ── Cell Value Parsing ──────────────────────────────────────────────

function parseCellValue(cell: XmlElement): CellValue {
  const valueType = cell.attrs["office:value-type"] ?? cell.attrs["calcext:value-type"] ?? ""

  switch (valueType) {
    case "string": {
      // Get text from <text:p> children, including from nested <text:a> elements
      const { text } = extractTextAndHyperlink(cell)
      if (text) return text
      // Check office:string-value attribute
      const strVal = cell.attrs["office:string-value"]
      if (strVal !== undefined) return strVal
      return ""
    }

    case "float":
    case "currency":
    case "percentage": {
      const val = cell.attrs["office:value"]
      if (val !== undefined) return Number(val)
      return null
    }

    case "boolean": {
      const boolVal = cell.attrs["office:boolean-value"]
      if (boolVal === "true") return true
      if (boolVal === "false") return false
      return null
    }

    case "date": {
      const dateVal = cell.attrs["office:date-value"]
      if (dateVal) return parseOdsDateTime(dateVal) ?? null
      return null
    }

    case "time": {
      // ODS time values are ISO 8601 durations like PT12H30M
      const timeVal = cell.attrs["office:time-value"]
      if (timeVal) {
        return timeVal
      }
      return null
    }

    default: {
      // No explicit type — try to extract text
      const { text } = extractTextAndHyperlink(cell)
      if (text) return text
      return null
    }
  }
}

// ── Content XML Parsing ─────────────────────────────────────────────

function parseContentXml(xml: string, options?: ReadOptions): Sheet[] {
  const doc = parseXml(xml)
  const sheets: Sheet[] = []

  const cellLimit = options?.maxTotalCells ?? MAX_TOTAL_CELLS

  // Parse styles for use when readStyles is enabled
  const readStyles = options?.readStyles ?? false
  const styleDefs = readStyles ? parseStyles(doc) : new Map<string, OdsStyleDef>()

  // Navigate: document-content > body > spreadsheet > table
  const body = findChild(doc, "body")
  if (!body) return sheets

  const spreadsheet = findChild(body, "spreadsheet")
  if (!spreadsheet) return sheets

  const tables = findChildren(spreadsheet, "table")

  for (const table of tables) {
    const name = table.attrs["table:name"] ?? `Sheet${sheets.length + 1}`

    // Filter sheets if specified
    if (options?.sheets !== undefined) {
      const filter = options.sheets
      const idx = sheets.length
      let shouldRead: boolean
      if (typeof filter === "function") {
        // ODS does not expose visibility state in the table directory.
        shouldRead = filter({ name, index: idx }, idx)
      } else if (filter.length === 0) {
        shouldRead = true
      } else {
        shouldRead = filter.some((spec) => {
          if (typeof spec === "string") return spec === name
          if (typeof spec === "number") return spec === idx
          return false
        })
      }
      if (!shouldRead) {
        sheets.push({ name, rows: [] }) // placeholder to maintain index
        continue
      }
    }

    const rows: CellValue[][] = []
    const merges: MergeRange[] = []
    const cells = new Map<string, Cell>()
    const tableRows = findChildren(table, "table-row")

    let currentRow = 0
    let pendingEmptyRows = 0

    // `maxRows` and `range` are on the shared `ReadOptions`, whose doc makes
    // no format-specific claim — but this reader used to read neither, so a
    // caller bounding a large ODS file got the whole thing and no warning.
    // See #439 §U. `maxRows` stops the walk; `range` masks afterwards,
    // matching what readXlsx returns for the same option.
    const maxRowsLimit = options?.maxRows ?? 0 // 0 = unlimited
    const rangeFilter = options?.range ? parseRange(options.range) : undefined

    for (const tableRow of tableRows) {
      if (maxRowsLimit > 0 && rows.length >= maxRowsLimit) break
      const rowRepeat = Number(tableRow.attrs["table:number-rows-repeated"] ?? "1")

      // Collect cell entries with their repeat counts first,
      // so we can trim trailing nulls before expanding
      const cellEntries: Array<{
        value: CellValue
        repeat: number
        colSpan: number
        rowSpan: number
        isCovered: boolean
        styleName?: string
        formula?: string
        hyperlink?: Hyperlink
      }> = []

      for (const child of tableRow.children) {
        if (typeof child === "string") continue
        const local = child.local || child.tag

        if (local === "table-cell") {
          const colRepeat = Number(child.attrs["table:number-columns-repeated"] ?? "1")
          const colSpan = Number(child.attrs["table:number-columns-spanned"] ?? "1")
          const rowSpan = Number(child.attrs["table:number-rows-spanned"] ?? "1")
          const value = parseCellValue(child)
          const styleName = child.attrs["table:style-name"]
          const formulaAttr = child.attrs["table:formula"]
          const formula = formulaAttr ? odsFormulaToExcel(formulaAttr) : undefined
          const { hyperlink } = extractTextAndHyperlink(child)
          cellEntries.push({
            value,
            repeat: colRepeat,
            colSpan,
            rowSpan,
            isCovered: false,
            styleName,
            formula,
            hyperlink,
          })
        } else if (local === "covered-table-cell") {
          const colRepeat = Number(child.attrs["table:number-columns-repeated"] ?? "1")
          cellEntries.push({
            value: null,
            repeat: colRepeat,
            colSpan: 1,
            rowSpan: 1,
            isCovered: true,
          })
        }
      }

      // A style name carries data only when the caller asked for styles
      // and this reader can resolve it. LibreOffice ends every row with
      // a default-styled cell repeated to column 16,384; keeping that
      // unknown style turns five values into 16,384 cells. See #464.
      while (cellEntries.length > 0) {
        const last = cellEntries[cellEntries.length - 1]!
        const hasStyle = readStyles && last.styleName && styleDefs.has(last.styleName)
        if (
          last.value !== null ||
          last.colSpan !== 1 ||
          last.rowSpan !== 1 ||
          hasStyle ||
          last.formula ||
          last.hyperlink
        ) {
          break
        }
        cellEntries.pop()
      }

      // Expand into row data and collect metadata
      const rowData: CellValue[] = []
      let col = 0

      for (const entry of cellEntries) {
        // Clamp column repeats too: a non-trailing cell with a huge
        // number-columns-repeated would otherwise allocate past Excel's
        // column limit. (Trailing empty repeats are already trimmed above.)
        const repeat = Math.min(entry.repeat, MAX_COL_INDEX + 1)
        for (let r = 0; r < repeat; r++) {
          rowData.push(entry.value)

          // Collect merge ranges
          if (entry.colSpan > 1 || entry.rowSpan > 1) {
            merges.push({
              startRow: currentRow,
              startCol: col,
              endRow: currentRow + entry.rowSpan - 1,
              endCol: col + entry.colSpan - 1,
            })
          }

          // Collect cell metadata (formulas, hyperlinks, styles)
          const hasMetadata =
            entry.formula ||
            entry.hyperlink ||
            (readStyles && entry.styleName && styleDefs.has(entry.styleName))

          if (hasMetadata) {
            const cellData: Cell = {
              value: entry.value,
              type:
                entry.value === null
                  ? "empty"
                  : typeof entry.value === "string"
                    ? "string"
                    : typeof entry.value === "number"
                      ? "number"
                      : typeof entry.value === "boolean"
                        ? "boolean"
                        : entry.value instanceof Date
                          ? "date"
                          : "empty",
            }

            if (entry.formula) {
              cellData.formula = entry.formula
              cellData.type = "formula"
            }
            if (entry.hyperlink) {
              cellData.hyperlink = entry.hyperlink
            }
            if (readStyles && entry.styleName) {
              const styleDef = styleDefs.get(entry.styleName)
              if (styleDef) {
                cellData.style = odsStyleToCellStyle(styleDef)
              }
            }

            cells.set(`${currentRow},${col}`, cellData)
          }

          col++
        }
      }

      if (rowData.length === 0) {
        // An empty row is held back rather than pushed. Whether it is data
        // depends on what comes after it: an interior one carries position
        // and has to survive, while the run LibreOffice pads the end of a
        // sheet with — one row repeated a million times — is not. Deciding
        // that here would need lookahead; deferring costs nothing and keeps
        // the trailing run from ever being allocated. See #394.
        // A malformed repeat parses to NaN, and the populated path drops
        // such a row outright (Math.min(NaN, …) is NaN, so its loop never
        // runs) — keep NaN out of the accumulator rather than letting it
        // poison every later flush.
        if (rowRepeat > 0) pendingEmptyRows += rowRepeat
        // The row counter still advances: merges and `cells` are keyed off
        // it, so it has to track the file's own row numbering either way.
        currentRow += rowRepeat
        continue
      }

      // A populated row makes every held-back empty row an interior one, so
      // flush them at the positions the file gave them. They carry no cells,
      // which puts them outside the MAX_TOTAL_CELLS guard below — bound them
      // by the sheet's row limit instead, or a file of nothing but huge
      // repeated empty rows would allocate without limit.
      if (pendingEmptyRows > 0) {
        const flush = Math.min(pendingEmptyRows, MAX_ROW_INDEX + 1 - rows.length)
        for (let r = 0; r < flush; r++) {
          rows.push([])
        }
        pendingEmptyRows = 0
      }

      // A hostile file can set a huge number-rows-repeated on a one-cell row
      // to force millions of allocations — clamp to Excel's row limit.
      const effectiveRowRepeat = Math.min(rowRepeat, MAX_ROW_INDEX + 1)

      // Each repeat attribute is capped on its own, but the aggregate is
      // not: one row of 16,384 cells repeated 1,048,576 times is 1.7e10
      // slots from a couple hundred bytes of content.xml. See #363.
      const projected = (rows.length + effectiveRowRepeat) * rowData.length
      if (projected > cellLimit) {
        throw new ParseError(
          `Sheet spans ${projected} cells, over the ${cellLimit} limit. ` +
            "Raise `maxTotalCells` if the sheet really is this large.",
        )
      }

      for (let r = 0; r < effectiveRowRepeat; r++) {
        rows.push(effectiveRowRepeat === 1 && r === 0 ? rowData : [...rowData])
        if (r > 0 && merges.length > 0) {
          // For repeated rows with merges, we'd need to duplicate merge info
          // but this is an edge case; repeated rows with merges are uncommon
        }
        currentRow++
      }
    }

    // Trim trailing empty rows. The walk above no longer pushes any (a run
    // of empty rows is only flushed once a populated row follows it), so
    // this is a backstop rather than the mechanism.
    while (rows.length > 0 && rows[rows.length - 1].length === 0) {
      rows.pop()
    }

    // `maxRows` can overshoot by the tail of a repeated row, since a single
    // <table-row table:number-rows-repeated="N"> expands after the check.
    if (maxRowsLimit > 0 && rows.length > maxRowsLimit) rows.length = maxRowsLimit

    // `range` masks rather than drops, so column indexes stay stable and a
    // row outside the span is present and empty — the same shape readXlsx
    // returns for the same option.
    if (rangeFilter) {
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r]!
        const inRowSpan = r >= rangeFilter.startRow && r <= rangeFilter.endRow
        for (let c = 0; c < row.length; c++) {
          if (!inRowSpan || c < rangeFilter.startCol || c > rangeFilter.endCol) {
            row[c] = null
            cells.delete(`${r},${c}`)
          }
        }
      }
    }

    const sheet: Sheet = { name, rows }

    if (merges.length > 0) {
      sheet.merges = merges
    }

    if (cells.size > 0) {
      sheet.cells = cells
    }

    sheets.push(sheet)
  }

  // If filter was applied, remove placeholder sheets with empty rows
  if (options?.sheets !== undefined) {
    const filter = options.sheets
    return sheets.filter((s, idx) => {
      if (s.rows.length > 0 || s.merges !== undefined || s.cells !== undefined) {
        return true
      }
      // Empty sheets that were genuinely selected by the filter must be kept.
      if (typeof filter === "function") {
        return filter({ name: s.name, index: idx }, idx)
      }
      if (filter.length === 0) return true
      return filter.some((spec) => {
        if (typeof spec === "string") return spec === s.name
        return false
      })
    })
  }

  return sheets
}

// ── Meta XML Parsing ────────────────────────────────────────────────

function parseMetaXml(xml: string): Partial<WorkbookProperties> {
  const doc = parseXml(xml)
  const props: Partial<WorkbookProperties> = {}

  // Navigate to office:meta element
  const meta = findChild(doc, "meta")
  if (!meta) return props

  for (const child of meta.children) {
    if (typeof child === "string") continue
    const local = child.local || child.tag
    const text = child.children.filter((c: unknown) => typeof c === "string").join("")

    switch (local) {
      case "title":
        if (text) props.title = text
        break
      case "subject":
        if (text) props.subject = text
        break
      case "initial-creator":
        if (text) props.creator = text
        break
      case "description":
        if (text) props.description = text
        break
      case "keyword":
        if (text) props.keywords = text
        break
      case "creation-date":
        if (text) {
          // LibreOffice writes these without a zone designator, so they
          // need the same UTC reading as office:date-value. See #415.
          const d = parseOdsDateTime(text)
          if (d) props.created = d
        }
        break
      case "date":
        if (text) {
          const d = parseOdsDateTime(text)
          if (d) props.modified = d
        }
        break
    }
  }

  return props
}

// ── Main Reader ─────────────────────────────────────────────────────

/**
 * Read an ODS file and return a Workbook.
 * Input can be Uint8Array, ArrayBuffer, or ReadableStream&lt;Uint8Array&gt;.
 *
 * For ReadableStream input, the stream is fully buffered before parsing
 * because the ZIP central directory lives at the end of the archive.
 */
export async function readOds(input: ReadInput, options?: ReadOptions): Promise<Workbook> {
  const data = await readInputToUint8Array(input, options?.maxInputBytes)

  // ODF supports password-encrypted documents via the same OLE2 / CFB
  // envelope Office uses for XLSX. Catch it before the ZIP reader does
  // so callers see a typed `EncryptedFileError` rather than a generic
  // ZIP ParseError. Decryption is tracked in #156.
  assertNotEncrypted(data, "ods")

  // 1. Open ZIP archive
  let zip: ZipReader
  try {
    zip = new ZipReader(data, options?.maxDecompressedBytes)
  } catch (err) {
    if (err instanceof ZipError) throw err
    throw new ParseError("Failed to open ODS file: not a valid ZIP archive", undefined, {
      cause: err,
    })
  }

  // 2. Verify mimetype — ODF spec requires it as the first ZIP entry
  if (!zip.has("mimetype")) {
    throw new ParseError(
      "Invalid ODS: missing 'mimetype' entry. The file may not be a valid OpenDocument Spreadsheet.",
    )
  }
  const mimeData = await zip.extract("mimetype")
  const mime = decodeUtf8(mimeData).trim()
  if (!mime.startsWith("application/vnd.oasis.opendocument")) {
    throw new ParseError(
      `Invalid ODS mimetype: "${mime}". Expected an OpenDocument type starting with "application/vnd.oasis.opendocument".`,
    )
  }

  // 3. Parse content.xml (required)
  if (!zip.has("content.xml")) {
    throw new ParseError("Invalid ODS: missing content.xml")
  }
  const contentXml = decodeUtf8(await zip.extract("content.xml"), "content.xml")
  const sheets = parseContentXml(contentXml, options)

  // 4. Parse meta.xml (optional)
  let properties: WorkbookProperties | undefined
  if (zip.has("meta.xml")) {
    const metaXml = decodeUtf8(await zip.extract("meta.xml"), "meta.xml")
    const metaProps = parseMetaXml(metaXml)
    if (Object.keys(metaProps).length > 0) {
      properties = { ...metaProps }
    }
  }

  // 5. Build workbook
  const workbook: Workbook = {
    sheets,
  }

  if (properties) {
    workbook.properties = properties
  }

  return workbook
}

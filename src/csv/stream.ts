// ── CSV Streaming ────────────────────────────────────────────────────
// Stream CSV rows as a synchronous generator (line by line).
// Stream CSV writer builds output incrementally.

import type { CellValue, CsvReadOptions, CsvWriteOptions } from "../_types"
import { stripBom, detectDelimiter } from "./reader"

// ── Type inference (duplicated from reader to avoid coupling) ────────

const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?)?$/

function inferType(value: string): CellValue {
  const trimmed = value.trim()
  if (trimmed === "") return value

  const lower = trimmed.toLowerCase()
  if (lower === "true" || lower === "yes") return true
  if (lower === "false" || lower === "no") return false

  if (ISO_DATE_RE.test(trimmed)) {
    const d = new Date(trimmed)
    if (!Number.isNaN(d.getTime())) return d
  }

  const asNumber = parseNumber(trimmed)
  if (asNumber !== null) return asNumber

  return value
}

function parseNumber(s: string): number | null {
  const stripped = s.replace(/,(\d{3})/g, "$1")
  if (stripped === "" || stripped === "-" || stripped === "+") return null
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(stripped)) return null
  const n = Number(stripped)
  if (Number.isNaN(n) || !Number.isFinite(n)) return null
  return n
}

// ── Helpers ──────────────────────────────────────────────────────────

function startsWith(str: string, prefix: string, offset: number): boolean {
  if (offset + prefix.length > str.length) return false
  for (let i = 0; i < prefix.length; i++) {
    if (str[offset + i] !== prefix[i]) return false
  }
  return true
}

// ── Streaming CSV Reader ─────────────────────────────────────────────

/**
 * Stream CSV rows as a synchronous generator.
 * Processes the string incrementally and yields one row at a time.
 */
export function* streamCsvRows(
  input: string,
  options?: CsvReadOptions,
): Generator<CellValue[], void, undefined> {
  const skipBom = options?.skipBom !== false
  const quote = options?.quote ?? '"'
  const escape = options?.escape ?? '"'
  const doTypeInference = options?.typeInference ?? false
  const skipEmptyRows = options?.skipEmptyRows ?? false
  const commentChar = options?.comment
  const isHeaderMode = options?.header ?? false

  if (skipBom) {
    input = stripBom(input)
  }

  if (input.length === 0) return

  const delimiter = options?.delimiter ?? detectDelimiter(input)
  const len = input.length

  let i = 0
  let isFirstRow = true
  let _headerRow: string[] | null = null

  while (i < len) {
    // Parse one row
    const row: string[] = []
    let currentField = ""
    let inQuoted = false
    let rowDone = false

    while (i < len && !rowDone) {
      const ch = input[i]!

      if (inQuoted) {
        // Check for escape sequence
        if (ch === escape && i + 1 < len && input[i + 1] === quote) {
          currentField += quote
          i += 2
          continue
        }
        // End of quoted field
        if (ch === quote) {
          inQuoted = false
          i++
          continue
        }
        // Any other character inside quotes
        currentField += ch
        i++
        continue
      }

      // Not in quoted field
      if (startsWith(input, delimiter, i)) {
        row.push(currentField)
        currentField = ""
        i += delimiter.length
        continue
      }

      // Check for line endings
      if (ch === "\r") {
        row.push(currentField)
        currentField = ""
        if (i + 1 < len && input[i + 1] === "\n") {
          i += 2
        } else {
          i++
        }
        rowDone = true
        continue
      }

      if (ch === "\n") {
        row.push(currentField)
        currentField = ""
        i++
        rowDone = true
        continue
      }

      // Start of quoted field
      if (ch === quote && currentField === "") {
        inQuoted = true
        i++
        continue
      }

      currentField += ch
      i++
    }

    // End of input without trailing newline
    if (!rowDone) {
      if (currentField !== "" || row.length > 0) {
        row.push(currentField)
      } else {
        // Nothing left
        break
      }
    }

    // Skip empty rows if configured
    if (row.length === 0) continue
    if (skipEmptyRows && row.every((cell) => cell === "")) continue

    // Skip comment rows
    if (commentChar && row.length > 0 && row[0].startsWith(commentChar)) {
      continue
    }

    // Handle header row
    if (isFirstRow && isHeaderMode) {
      _headerRow = row
      isFirstRow = false
      continue
    }
    isFirstRow = false

    // Apply type inference if requested
    if (doTypeInference) {
      const typedRow: CellValue[] = row.map((v) => inferType(v))
      yield typedRow
    } else {
      yield row
    }
  }
}

// ── Streaming CSV Writer ─────────────────────────────────────────────

const UTF8_BOM = "\uFEFF"

export class CsvStreamWriter {
  private delimiter: string
  private lineSeparator: string
  private quote: string
  private quoteStyle: "all" | "required" | "none"
  private bom: boolean
  private dateFormat: string | undefined
  private nullValue: string
  private lines: string[] = []
  private headerWritten = false
  private headers: string[] | boolean | undefined

  constructor(options?: CsvWriteOptions) {
    this.delimiter = options?.delimiter ?? ","
    this.lineSeparator = options?.lineSeparator ?? "\r\n"
    this.quote = options?.quote ?? '"'
    this.quoteStyle = options?.quoteStyle ?? "required"
    this.bom = options?.bom ?? false
    this.dateFormat = options?.dateFormat
    this.nullValue = options?.nullValue ?? ""
    this.headers = options?.headers

    // Write header row immediately if string array provided
    if (Array.isArray(this.headers) && !this.headerWritten) {
      const headerLine = this.headers.map((h) => this.quoteField(h)).join(this.delimiter)
      this.lines.push(headerLine)
      this.headerWritten = true
    }
  }

  /** Add a row of values */
  addRow(values: CellValue[]): void {
    const line = values.map((v) => this.formatAndQuote(v)).join(this.delimiter)
    this.lines.push(line)
  }

  /** Finalize and return the CSV string */
  finish(): string {
    const parts: string[] = []

    if (this.bom) {
      parts.push(UTF8_BOM)
    }

    parts.push(this.lines.join(this.lineSeparator))

    return parts.join("")
  }

  // ── Private helpers ───────────────────────────────────────────────

  private formatAndQuote(value: CellValue): string {
    if (value === null || value === undefined) {
      if (this.quoteStyle === "all") {
        return this.quote + this.nullValue + this.quote
      }
      return this.nullValue
    }

    if (typeof value === "boolean") {
      return this.quoteField(value ? "true" : "false")
    }

    if (typeof value === "number") {
      return this.quoteField(this.formatNumber(value))
    }

    if (value instanceof Date) {
      return this.quoteField(this.formatDate(value))
    }

    return this.quoteField(String(value))
  }

  private quoteField(value: string): string {
    if (this.quoteStyle === "none") {
      return value
    }

    const needsQuoting =
      this.quoteStyle === "all" ||
      value.includes(this.delimiter) ||
      value.includes(this.quote) ||
      value.includes("\n") ||
      value.includes("\r")

    if (!needsQuoting) {
      return value
    }

    const escaped = value.replaceAll(this.quote, this.quote + this.quote)
    return this.quote + escaped + this.quote
  }

  private formatNumber(n: number): string {
    if (Number.isInteger(n) && Math.abs(n) >= 1e15) {
      return n.toFixed(0)
    }
    if (Math.abs(n) > 0 && Math.abs(n) < 1e-6) {
      return n.toFixed(20).replace(/0+$/, "").replace(/\.$/, ".0")
    }
    return String(n)
  }

  private formatDate(d: Date): string {
    if (!this.dateFormat) {
      return d.toISOString()
    }

    const year = d.getFullYear()
    const month = d.getMonth() + 1
    const day = d.getDate()
    const hours = d.getHours()
    const minutes = d.getMinutes()
    const seconds = d.getSeconds()

    return this.dateFormat
      .replace("YYYY", String(year))
      .replace("MM", String(month).padStart(2, "0"))
      .replace("DD", String(day).padStart(2, "0"))
      .replace("HH", String(hours).padStart(2, "0"))
      .replace("mm", String(minutes).padStart(2, "0"))
      .replace("ss", String(seconds).padStart(2, "0"))
  }
}

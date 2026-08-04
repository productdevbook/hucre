import type { CellValue, CsvWriteOptions } from "../_types"
import { escapeFormula } from "./formula"

// ── BOM constant ─────────────────────────────────────────────────────

const UTF8_BOM = "\uFEFF"

// ── Public API ───────────────────────────────────────────────────────

/**
 * Format a single CellValue for CSV output.
 */
export function formatCsvValue(value: CellValue, options?: CsvWriteOptions): string {
  const opts = normalizeWriteOptions(options)

  // null / undefined
  if (value === null || value === undefined) {
    return opts.nullValue
  }

  // Boolean
  if (typeof value === "boolean") {
    return value ? "true" : "false"
  }

  // Number
  if (typeof value === "number") {
    return formatNumber(value)
  }

  // Date
  if (value instanceof Date) {
    return formatDate(value, opts.dateFormat)
  }

  // String — apply formula escaping if enabled, then quoting
  let str = String(value)
  if (opts.escapeFormulae) {
    str = escapeFormula(str)
  }
  return quoteField(str, opts)
}

/**
 * Write a 2D array of cell values to a CSV string.
 */
export function writeCsv(rows: CellValue[][], options?: CsvWriteOptions): string {
  const opts = normalizeWriteOptions(options)
  const parts: string[] = []

  // BOM
  if (opts.bom) {
    parts.push(UTF8_BOM)
  }

  // Headers row
  if (opts.headers) {
    if (Array.isArray(opts.headers)) {
      parts.push(opts.headers.map((h) => quoteField(h, opts)).join(opts.delimiter))
      if (rows.length > 0) {
        parts.push(opts.lineSeparator)
      }
    }
  }

  // Data rows
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]!
    const line = row.map((cell) => formatAndQuote(cell, opts)).join(opts.delimiter)
    parts.push(line)
    if (i < rows.length - 1) {
      parts.push(opts.lineSeparator)
    }
  }

  return parts.join("")
}

/**
 * Write an array of objects to a CSV string.
 */
export function writeCsvObjects(
  data: Array<Record<string, CellValue>>,
  options?: CsvWriteOptions,
): string {
  const opts = normalizeWriteOptions(options)

  // If columns option is provided, use it as the column order
  const explicitColumns = options?.columns

  // Determine headers
  let headers: string[]
  if (explicitColumns) {
    headers = explicitColumns
  } else if (Array.isArray(opts.headers)) {
    headers = opts.headers
  } else if (opts.headers === true || opts.headers === undefined) {
    // Auto-detect from first object's keys
    if (data.length === 0) {
      return opts.bom ? UTF8_BOM : ""
    }
    headers = Object.keys(data[0]!)
  } else {
    // headers === false — no header row, but we still need column order
    if (data.length === 0) {
      return opts.bom ? UTF8_BOM : ""
    }
    headers = Object.keys(data[0]!)
    // Convert to rows and write without headers
    const rows: CellValue[][] = data.map((obj) =>
      headers.map((key) => {
        const val = obj[key]
        return val === undefined ? null : val
      }),
    )
    return writeCsv(rows, { ...options, headers: undefined })
  }

  // Convert objects to rows
  const rows: CellValue[][] = data.map((obj) =>
    headers.map((key) => {
      const val = obj[key]
      return val === undefined ? null : val
    }),
  )

  return writeCsv(rows, { ...options, headers })
}

// ── Internal helpers ─────────────────────────────────────────────────

interface NormalizedWriteOptions {
  delimiter: string
  lineSeparator: string
  quote: string
  quoteStyle: "all" | "required" | "none"
  headers: string[] | boolean | undefined
  bom: boolean
  dateFormat: string | undefined
  nullValue: string
  escapeFormulae: boolean
  comment: string | undefined
}

function normalizeWriteOptions(options?: CsvWriteOptions): NormalizedWriteOptions {
  return {
    delimiter: options?.delimiter ?? ",",
    lineSeparator: options?.lineSeparator ?? "\r\n",
    quote: options?.quote ?? '"',
    quoteStyle: options?.quoteStyle ?? "required",
    headers: options?.headers,
    bom: options?.bom ?? false,
    dateFormat: options?.dateFormat,
    nullValue: options?.nullValue ?? "",
    escapeFormulae: options?.escapeFormulae ?? false,
    // "" would make startsWith() true for every value, so treat it as unset.
    comment: options?.comment || undefined,
  }
}

function formatAndQuote(value: CellValue, opts: NormalizedWriteOptions): string {
  if (value === null || value === undefined) {
    const raw = opts.nullValue
    if (opts.quoteStyle === "all") {
      return opts.quote + raw + opts.quote
    }
    return raw
  }

  if (typeof value === "boolean") {
    const raw = value ? "true" : "false"
    return quoteField(raw, opts)
  }

  if (typeof value === "number") {
    const raw = formatNumber(value)
    return quoteField(raw, opts)
  }

  if (value instanceof Date) {
    const raw = formatDate(value, opts.dateFormat)
    return quoteField(raw, opts)
  }

  let str = String(value)
  if (opts.escapeFormulae) {
    str = escapeFormula(str)
  }
  return quoteField(str, opts)
}

function quoteField(value: string, opts: NormalizedWriteOptions): string {
  if (opts.quoteStyle === "none") {
    return value
  }

  const needsQuoting =
    opts.quoteStyle === "all" ||
    value.includes(opts.delimiter) ||
    value.includes(opts.quote) ||
    value.includes("\n") ||
    value.includes("\r") ||
    // A leading comment character is quoted so a reader configured with
    // `comment` keeps the row instead of dropping the whole line (#408).
    // The reader only skips *unquoted* leading comment chars, so quoting
    // is a complete fix — and it is applied wherever the value sits, since
    // a caller may reorder or concatenate what we hand back.
    (opts.comment !== undefined && value.startsWith(opts.comment))

  if (!needsQuoting) {
    return value
  }

  // Escape quote characters by doubling them
  const escaped = value.replaceAll(opts.quote, opts.quote + opts.quote)
  return opts.quote + escaped + opts.quote
}

function formatNumber(n: number): string {
  // Avoid scientific notation for large integers
  if (Number.isInteger(n) && Math.abs(n) >= 1e15) {
    return n.toFixed(0)
  }
  // For very small numbers that would use scientific notation
  if (Math.abs(n) > 0 && Math.abs(n) < 1e-6) {
    return n.toFixed(20).replace(/0+$/, "").replace(/\.$/, ".0")
  }
  return String(n)
}

function formatDate(d: Date, format?: string): string {
  if (!format) {
    // See #364 — an unparseable Date threw a raw RangeError mid-write.
    if (Number.isNaN(d.getTime())) return ""
    return d.toISOString()
  }

  // Simple date format placeholders
  const year = d.getFullYear()
  const month = d.getMonth() + 1
  const day = d.getDate()
  const hours = d.getHours()
  const minutes = d.getMinutes()
  const seconds = d.getSeconds()

  return format
    .replace("YYYY", String(year))
    .replace("MM", String(month).padStart(2, "0"))
    .replace("DD", String(day).padStart(2, "0"))
    .replace("HH", String(hours).padStart(2, "0"))
    .replace("mm", String(minutes).padStart(2, "0"))
    .replace("ss", String(seconds).padStart(2, "0"))
}

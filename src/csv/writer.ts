import { isCellError } from "../cell-error"
import type { CellValue, CsvWriteOptions } from "../_types"
import { escapeFormula } from "./formula"
import { formatDate as formatExcelDate } from "../_date"
import { collectHeaders } from "../_objects"

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

  // String, or an error's token — apply formula escaping if enabled, then quoting
  let str = isCellError(value) ? value.error : value
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
  if (opts.headers && opts.writeHeader) {
    parts.push(opts.headers.map((h) => quoteField(h, opts)).join(opts.delimiter))
    if (rows.length > 0) {
      parts.push(opts.lineSeparator)
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
  } else if (opts.headers) {
    headers = opts.headers
  } else if (opts.writeHeader) {
    // Auto-detect from first object's keys
    if (data.length === 0) {
      return opts.bom ? UTF8_BOM : ""
    }
    headers = collectHeaders(data)
  } else {
    // writeHeader: false — no header row, but we still need column order
    if (data.length === 0) {
      return opts.bom ? UTF8_BOM : ""
    }
    headers = collectHeaders(data)
    // Convert to rows and write without headers
    const rows: CellValue[][] = data.map((obj) =>
      headers.map((key) => {
        const val = obj[key]
        return val === undefined ? null : val
      }),
    )
    return writeCsv(rows, { ...options, headers: undefined, writeHeader: false })
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
  headers: string[] | undefined
  writeHeader: boolean
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
    writeHeader: options?.writeHeader !== false,
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

  let str = isCellError(value) ? value.error : value
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

/**
 * Render a number for CSV.
 *
 * Excel shows a value written in exponent notation as `1E-07`, so the
 * plain decimal form is preferred where there is one — but only when it
 * is *the same number*. The expansion used to be unconditional, and
 * `toFixed(20)` caps at twenty decimal places:
 *
 *   Number.EPSILON  →  "0.00000000000000022204"   (five digits kept)
 *   Number.MIN_VALUE →  "0.0"                      (all of them lost)
 *
 * So the smallest values a caller could put in a cell came back as zero.
 * A prettier rendering is not worth a different number. See #474.
 */
function formatNumber(n: number): string {
  if (!Number.isFinite(n)) return String(n)

  // Large integers: `1e+21` reads as text in some importers.
  if (Number.isInteger(n) && Math.abs(n) >= 1e15) {
    const plain = n.toFixed(0)
    if (Number(plain) === n) return plain
  }

  // Small magnitudes, where JS switches to exponent notation at 1e-7.
  if (Math.abs(n) > 0 && Math.abs(n) < 1e-6) {
    const plain = n.toFixed(20).replace(/0+$/, "").replace(/\.$/, ".0")
    if (Number(plain) === n) return plain
  }

  return String(n)
}

/**
 * Render a `Date` for CSV output.
 *
 * Delegates to the library's own {@link formatExcelDate}, so a format
 * string means the same thing here as it does in a `numFmt`, in
 * `formatValue`, and in the public `formatDate` export. This file used to
 * carry a private substitute that accepted a different vocabulary
 * (`YYYY MM DD HH mm ss`), read *local* time components while the
 * no-format path read UTC, and substituted with a non-global `.replace()`
 * so a repeated token stayed literal. See #439.
 */
function formatDate(d: Date, format?: string): string {
  // See #364 — an unparseable Date threw a raw RangeError mid-write.
  if (Number.isNaN(d.getTime())) return ""
  if (!format) return d.toISOString()
  return formatExcelDate(d, format)
}

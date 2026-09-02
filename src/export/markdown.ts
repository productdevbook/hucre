import { isCellError } from "../cell-error"
import type { Sheet, CellValue } from "../_types"

export interface MarkdownExportOptions {
  /**
   * Treat the first row as the table header. Default: true.
   *
   * Renamed from `headerRow`, which elsewhere in the library is a 0-based
   * row *index*. See #365.
   */
  hasHeaderRow?: boolean
  /** Alignment per column. Default: left for strings, right for numbers */
  alignment?: Array<"left" | "center" | "right">
  /** Max column width (truncate with ...). Default: 50 */
  maxWidth?: number
  /**
   * Escape Markdown's inline syntax inside cells, so a cell reading
   * `*not emphasis*` renders as those words rather than as emphasis.
   * Default: true.
   *
   * Set false when the cells hold Markdown you want rendered — a column
   * of `**bold**` labels, say. The table's own structural characters
   * (pipes, newlines) are escaped either way, because losing those loses
   * the table. See #474.
   */
  escapeInline?: boolean
}

/**
 * Characters that change how a table cell renders. `\\` leads, so the
 * escapes this function adds are not themselves re-escaped.
 *
 * `_` is here despite GFM not treating intraword underscores as emphasis:
 * `_total_` is emphasis and `sub_total` is not, and telling them apart
 * needs the flanking rules. Escaping both is the honest trade — a cell of
 * `sub\_total` reads correctly even if it looks noisy in the source, and
 * `escapeInline: false` is there for anyone who disagrees.
 */
const MARKDOWN_INLINE = /[\\`*_[\]<]/g

/**
 * Escape characters that would break a Markdown table cell: pipes (column
 * separators) and newlines (row separators). A literal newline inside a
 * cell is rendered as `<br>` (GFM) so the table structure survives.
 *
 * With `escapeInline` (the default) the inline-formatting characters go
 * too. The `<br>` is inserted after that pass, so it is not escaped by it.
 */
function escapeCell(str: string, escapeInline: boolean): string {
  const inline = escapeInline ? str.replace(MARKDOWN_INLINE, "\\$&") : str
  return inline.replace(/\|/g, "\\|").replace(/\r\n|\r|\n/g, "<br>")
}

/** Format a cell value as a string for Markdown output */
function formatCellValue(value: CellValue, escapeInline: boolean): string {
  if (value === null || value === undefined) return ""
  if (value instanceof Date) {
    // See #364 — an unparseable Date threw a raw RangeError mid-write.
    if (Number.isNaN(value.getTime())) return ""
    return value.toISOString().slice(0, 10)
  }
  if (typeof value === "boolean") return String(value)
  if (typeof value === "number") return String(value)
  if (isCellError(value)) return value.error
  return escapeCell(String(value), escapeInline)
}

/** Truncate a string to maxWidth, adding "..." if truncated */
function truncate(str: string, maxWidth: number): string {
  if (str.length <= maxWidth) return str
  if (maxWidth <= 3) return str.slice(0, maxWidth)
  return str.slice(0, maxWidth - 3) + "..."
}

/** Detect the default alignment for a column based on the data types */
function detectAlignment(
  rows: CellValue[][],
  colIndex: number,
  startRow: number,
): "left" | "right" {
  for (let r = startRow; r < rows.length; r++) {
    const val = rows[r]?.[colIndex]
    if (val !== null && val !== undefined) {
      if (typeof val === "number") return "right"
      return "left"
    }
  }
  return "left"
}

/** Build the separator row (e.g., | --- | ---: |) */
function buildSeparator(alignments: Array<"left" | "center" | "right">, widths: number[]): string {
  const parts = alignments.map((align, i) => {
    const w = Math.max(widths[i], 3)
    // Each cell in padCell is: " " + content(w chars) + " " = w + 2 chars total
    // Separator must match that width exactly
    if (align === "right") return " " + "-".repeat(w) + ":"
    if (align === "center") return ":" + "-".repeat(w) + ":"
    return " " + "-".repeat(w) + " "
  })
  return "|" + parts.join("|") + "|"
}

/** Pad a cell value to a given width according to alignment */
function padCell(value: string, width: number, align: "left" | "center" | "right"): string {
  if (align === "right") {
    return " " + value.padStart(width) + " "
  }
  if (align === "center") {
    const totalPad = width - value.length
    const leftPad = Math.floor(totalPad / 2)
    const rightPad = totalPad - leftPad
    return " " + " ".repeat(leftPad) + value + " ".repeat(rightPad) + " "
  }
  return " " + value.padEnd(width) + " "
}

/**
 * Export a sheet as a Markdown table string.
 */
export function toMarkdown(sheet: Sheet, options?: MarkdownExportOptions): string {
  const opts: Required<MarkdownExportOptions> = {
    hasHeaderRow: options?.hasHeaderRow ?? true,
    alignment: options?.alignment ?? [],
    maxWidth: options?.maxWidth ?? 50,
    escapeInline: options?.escapeInline ?? true,
  }

  const rows = sheet.rows
  if (!rows || rows.length === 0) return ""

  // Determine number of columns
  let numCols = 0
  for (const row of rows) {
    if (row.length > numCols) numCols = row.length
  }
  if (numCols === 0) return ""

  // Format all cell values
  const formatted: string[][] = rows.map((row) => {
    const result: string[] = []
    for (let c = 0; c < numCols; c++) {
      const raw = formatCellValue(row[c], opts.escapeInline)
      result.push(truncate(raw, opts.maxWidth))
    }
    return result
  })

  // Determine the data start row (skip header if headerRow is true)
  const dataStartRow = opts.hasHeaderRow ? 1 : 0

  // Determine alignments
  const alignments: Array<"left" | "center" | "right"> = []
  for (let c = 0; c < numCols; c++) {
    if (opts.alignment && opts.alignment[c]) {
      alignments.push(opts.alignment[c])
    } else {
      alignments.push(detectAlignment(rows, c, dataStartRow))
    }
  }

  // Calculate column widths
  const widths: number[] = []
  for (let c = 0; c < numCols; c++) {
    let maxW = 3 // minimum width for separator
    for (const row of formatted) {
      if (row[c] && row[c].length > maxW) maxW = row[c].length
    }
    widths.push(maxW)
  }

  const lines: string[] = []

  if (opts.hasHeaderRow) {
    // Header row
    const headerCells = formatted[0].map((val, c) => padCell(val, widths[c], alignments[c]))
    lines.push("|" + headerCells.join("|") + "|")

    // Separator row
    lines.push(buildSeparator(alignments, widths))

    // Data rows
    for (let r = 1; r < formatted.length; r++) {
      const cells = formatted[r].map((val, c) => padCell(val, widths[c], alignments[c]))
      lines.push("|" + cells.join("|") + "|")
    }
  } else {
    // No header: generate a generic header from column indices
    const headerCells: string[] = []
    for (let c = 0; c < numCols; c++) {
      headerCells.push(padCell(String(c + 1), widths[c], alignments[c]))
    }
    lines.push("|" + headerCells.join("|") + "|")

    // Separator row
    lines.push(buildSeparator(alignments, widths))

    // All rows as data
    for (const row of formatted) {
      const cells = row.map((val, c) => padCell(val, widths[c], alignments[c]))
      lines.push("|" + cells.join("|") + "|")
    }
  }

  return lines.join("\n")
}

// ── Auto Size Calculator ────────────────────────────────────────────
// Re-exports auto column width and adds auto row height calculation.
// ─────────────────────────────────────────────────────────────────────

import type { CellValue } from "../_types"

export { calculateColumnWidth, measureValueWidth } from "./auto-width"

// ── Constants ────────────────────────────────────────────────────────

/** Default line height in points for Calibri 11pt (Excel default) */
const DEFAULT_LINE_HEIGHT = 15

/** Base font size for the default line height */
const BASE_FONT_SIZE = 11

/** Default column width in characters (used when no columnWidths provided) */
const DEFAULT_COLUMN_WIDTH = 8.43

// ── Row Height Calculation ──────────────────────────────────────────

/**
 * Count the number of visible characters in a string, treating CJK as 2 units.
 */
function measureTextWidth(text: string): number {
  let width = 0
  for (const char of text) {
    const cp = char.codePointAt(0)
    if (cp !== undefined && isCjk(cp)) {
      width += 2
    } else {
      width += 1
    }
  }
  return width
}

/**
 * Check if a code point is CJK (double-width).
 */
function isCjk(codePoint: number): boolean {
  return (
    (codePoint >= 0x4e00 && codePoint <= 0x9fff) ||
    (codePoint >= 0x3400 && codePoint <= 0x4dbf) ||
    (codePoint >= 0x20000 && codePoint <= 0x2a6df) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    (codePoint >= 0xac00 && codePoint <= 0xd7af) ||
    (codePoint >= 0x30a0 && codePoint <= 0x30ff) ||
    (codePoint >= 0x3040 && codePoint <= 0x309f) ||
    (codePoint >= 0x3000 && codePoint <= 0x303f) ||
    (codePoint >= 0xff00 && codePoint <= 0xff60) ||
    (codePoint >= 0xffe0 && codePoint <= 0xffe6)
  )
}

/**
 * Calculate the number of wrapped lines a text cell would occupy
 * given a column width in character units.
 */
function countWrappedLines(text: string, columnWidth: number): number {
  if (text.length === 0) return 1

  // Split on explicit newlines first
  const paragraphs = text.split("\n")
  let totalLines = 0

  for (const paragraph of paragraphs) {
    if (paragraph.length === 0) {
      totalLines += 1
      continue
    }
    const textWidth = measureTextWidth(paragraph)
    // Each paragraph wraps based on column width
    const wrappedLines = Math.ceil(textWidth / Math.max(columnWidth, 1))
    totalLines += Math.max(wrappedLines, 1)
  }

  return totalLines
}

/**
 * Calculate optimal row height based on cell content (wrap text, font size).
 *
 * Default row height is 15 points (Excel default for Calibri 11pt).
 * If text wraps, height = lineCount * lineHeight.
 * Line count is calculated as ceil(textLength / columnWidth) for each cell,
 * then the maximum across all cells in the row is used.
 *
 * @param values - All cell values in the row
 * @param options - Configuration options
 * @returns Optimal row height in points
 */
export function calculateRowHeight(
  values: CellValue[],
  options?: {
    /** Font size in points. Default: 11 (Calibri default) */
    fontSize?: number
    /** Whether text wrapping is enabled. Default: false */
    wrapText?: boolean
    /** Column widths in character units (one per cell). Default: 8.43 per column */
    columnWidths?: number[]
  },
): number {
  const fontSize = options?.fontSize ?? BASE_FONT_SIZE
  const wrapText = options?.wrapText ?? false

  // Line height scales proportionally with font size
  const lineHeight = (fontSize / BASE_FONT_SIZE) * DEFAULT_LINE_HEIGHT

  let maxLines = 1

  for (let i = 0; i < values.length; i++) {
    const value = values[i]
    if (value === null || value === undefined) continue

    const text = String(value)
    if (text.length === 0) continue

    if (wrapText) {
      const colWidth = options?.columnWidths?.[i] ?? DEFAULT_COLUMN_WIDTH
      const lines = countWrappedLines(text, colWidth)
      if (lines > maxLines) maxLines = lines
    } else {
      // Without wrap text, only explicit newlines increase height
      const newlineCount = text.split("\n").length
      if (newlineCount > maxLines) maxLines = newlineCount
    }
  }

  const height = maxLines * lineHeight

  // Round to nearest 0.25 (Excel row height granularity)
  return Math.ceil(height * 4) / 4
}

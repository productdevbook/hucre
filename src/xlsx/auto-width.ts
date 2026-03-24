// ── Auto Column Width Calculator ─────────────────────────────────────
// Calculates optimal column widths based on cell content.
//
// Excel column width unit = number of '0' characters of the "Normal" font
// that fit in a cell. Calibri 11pt is the Excel default.
//
// For proportional fonts, characters have varying widths. We use an
// average character width multiplier to approximate. Bold text is
// slightly wider (~1.05x). CJK characters count as 2 units.
// ─────────────────────────────────────────────────────────────────────

import type { CellValue, FontStyle } from "../_types";
import { formatDate } from "../_date";

// ── Constants ────────────────────────────────────────────────────────

/** Excel default minimum column width */
const DEFAULT_MIN_WIDTH = 8;

/** Excel maximum column width */
const MAX_COLUMN_WIDTH = 255;

/** Default cell padding (accounts for cell margins + filter dropdown arrow) */
const DEFAULT_PADDING = 2;

/**
 * Average character width multiplier for proportional fonts.
 * Calibri 11pt '0' character is the base unit. On average, characters in
 * Calibri are slightly wider than the '0' glyph due to variable widths.
 */
const PROPORTIONAL_FONT_MULTIPLIER = 1.1;

/** Bold text is ~5% wider than regular */
const BOLD_MULTIPLIER = 1.05;

// ── CJK Detection ───────────────────────────────────────────────────

/**
 * Check if a character code point is in a CJK (Chinese, Japanese, Korean) range.
 * CJK characters are typically displayed at double width.
 */
function isCjk(codePoint: number): boolean {
  return (
    // CJK Unified Ideographs
    (codePoint >= 0x4e00 && codePoint <= 0x9fff) ||
    // CJK Unified Ideographs Extension A
    (codePoint >= 0x3400 && codePoint <= 0x4dbf) ||
    // CJK Unified Ideographs Extension B
    (codePoint >= 0x20000 && codePoint <= 0x2a6df) ||
    // CJK Compatibility Ideographs
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    // Hangul Syllables
    (codePoint >= 0xac00 && codePoint <= 0xd7af) ||
    // Katakana
    (codePoint >= 0x30a0 && codePoint <= 0x30ff) ||
    // Hiragana
    (codePoint >= 0x3040 && codePoint <= 0x309f) ||
    // CJK Symbols and Punctuation
    (codePoint >= 0x3000 && codePoint <= 0x303f) ||
    // Fullwidth Forms
    (codePoint >= 0xff00 && codePoint <= 0xff60) ||
    // Halfwidth Forms (CJK portion)
    (codePoint >= 0xffe0 && codePoint <= 0xffe6)
  );
}

// ── Number Format Simulation ─────────────────────────────────────────

/**
 * Approximate the display string of a number given an Excel number format.
 * This is a simplified simulation -- it handles common patterns but does not
 * fully replicate Excel's format engine.
 */
function formatNumberForWidth(value: number, numFmt: string): string {
  const fmt = numFmt.trim();

  // Handle multiple sections (positive;negative;zero)
  const sections = fmt.split(";");
  let section: string;
  if (value > 0) {
    section = sections[0];
  } else if (value < 0 && sections.length > 1) {
    section = sections[1];
    value = Math.abs(value);
  } else if (value === 0 && sections.length > 2) {
    section = sections[2];
  } else {
    section = sections[0];
  }

  // Strip color/locale directives like [Red], [$-409], [$EUR]
  section = section.replace(/\[[^\]]*\]/g, "");
  // Strip escape characters
  section = section.replace(/\\./g, (m) => m[1]);
  // Strip quoted strings but keep their content for length calculation
  section = section.replace(/"([^"]*)"/g, "$1");

  // Percentage
  if (section.includes("%")) {
    const cleaned = section.replace(/%/g, "");
    const decimals = countDecimalPlaces(cleaned);
    const formatted = (value * 100).toFixed(decimals);
    return addThousandSeparator(formatted, section) + "%";
  }

  // Detect decimal places from format
  const decimals = countDecimalPlaces(section);

  // Check for thousand separator
  const formatted = value.toFixed(decimals);

  return addThousandSeparator(formatted, section);
}

/** Count the number of decimal places specified in a format pattern */
function countDecimalPlaces(fmt: string): number {
  const dotIdx = fmt.indexOf(".");
  if (dotIdx === -1) return 0;

  let count = 0;
  for (let i = dotIdx + 1; i < fmt.length; i++) {
    if (fmt[i] === "0" || fmt[i] === "#") {
      count++;
    } else {
      break;
    }
  }
  return count;
}

/** Add thousand separator commas if the format calls for it */
function addThousandSeparator(numStr: string, fmt: string): string {
  // Check if format has comma thousand separator (e.g., #,##0)
  if (!fmt.includes(",")) return numStr;

  const parts = numStr.split(".");
  let intPart = parts[0];
  const isNeg = intPart.startsWith("-");
  if (isNeg) intPart = intPart.slice(1);

  // Insert commas from right to left
  let result = "";
  for (let i = 0; i < intPart.length; i++) {
    if (i > 0 && (intPart.length - i) % 3 === 0) {
      result += ",";
    }
    result += intPart[i];
  }

  if (isNeg) result = "-" + result;
  if (parts[1] !== undefined) result += "." + parts[1];

  return result;
}

/**
 * Count prefix/suffix literal characters in a number format.
 * E.g., "$#,##0.00" has "$" prefix (1 char), "#,##0.00 TL" has " TL" suffix (3 chars).
 */
function countFormatLiterals(numFmt: string): number {
  let count = 0;
  // Strip sections, take first
  const section = numFmt.split(";")[0];
  // Strip color/locale
  const cleaned = section.replace(/\[[^\]]*\]/g, "");

  // Check for currency symbols and other literal prefixes/suffixes
  for (let i = 0; i < cleaned.length; i++) {
    const ch = cleaned[i];
    if (ch === '"') {
      // Quoted literal
      i++;
      while (i < cleaned.length && cleaned[i] !== '"') {
        count++;
        i++;
      }
    } else if (ch === "\\") {
      count++;
      i++; // skip escaped char
    } else if (ch === "$" || ch === "\u20AC" || ch === "\u00A3" || ch === "\u00A5") {
      // Common currency symbols
      count++;
    }
  }

  return count;
}

// ── Core Functions ───────────────────────────────────────────────────

/**
 * Measure the display width (in character units) of a single line of text.
 * CJK characters count as 2 units, all others as 1.
 */
function measureLineWidth(text: string): number {
  let width = 0;
  for (const char of text) {
    const codePoint = char.codePointAt(0);
    if (codePoint !== undefined && isCjk(codePoint)) {
      width += 2;
    } else {
      width += 1;
    }
  }
  return width;
}

/**
 * Calculate the display width of a formatted cell value.
 *
 * @param value - The cell value
 * @param numFmt - Optional Excel number format string
 * @returns Width in character units (before font multiplier and padding)
 */
export function measureValueWidth(value: CellValue, numFmt?: string): number {
  if (value === null || value === undefined) {
    return 0;
  }

  if (typeof value === "boolean") {
    // Excel displays TRUE (5 chars) or FALSE (6 chars)
    return value ? 5 : 6;
  }

  if (typeof value === "string") {
    if (value.length === 0) return 0;

    // Handle multiline: take the longest line
    if (value.includes("\n")) {
      const lines = value.split("\n");
      let maxWidth = 0;
      for (const line of lines) {
        const w = measureLineWidth(line);
        if (w > maxWidth) maxWidth = w;
      }
      return maxWidth;
    }

    return measureLineWidth(value);
  }

  if (typeof value === "number") {
    if (numFmt) {
      // Format the number and measure the result
      const formatted = formatNumberForWidth(value, numFmt);
      const extraLiterals = countFormatLiterals(numFmt);
      return measureLineWidth(formatted) + extraLiterals;
    }

    // Default number formatting: use toString
    const str = String(value);
    return measureLineWidth(str);
  }

  if (value instanceof Date) {
    if (numFmt) {
      const formatted = formatDate(value, numFmt);
      return measureLineWidth(formatted);
    }

    // Default date format: yyyy-mm-dd (10 chars)
    const y = value.getUTCFullYear();
    const m = String(value.getUTCMonth() + 1).padStart(2, "0");
    const d = String(value.getUTCDate()).padStart(2, "0");
    return measureLineWidth(`${y}-${m}-${d}`);
  }

  return 0;
}

/**
 * Calculate the optimal column width for a set of cell values.
 * Width is measured in Excel character units (number of '0' characters
 * of the Normal font that fit in a cell).
 *
 * @param values - All cell values in the column (including header)
 * @param options - Configuration options
 * @returns Optimal column width in Excel character units
 */
export function calculateColumnWidth(
  values: CellValue[],
  options?: {
    font?: FontStyle;
    numFmt?: string;
    minWidth?: number;
    maxWidth?: number;
    padding?: number;
  },
): number {
  const minWidth = options?.minWidth ?? DEFAULT_MIN_WIDTH;
  const maxWidth = options?.maxWidth ?? MAX_COLUMN_WIDTH;
  const padding = options?.padding ?? DEFAULT_PADDING;
  const isBold = options?.font?.bold === true;

  let maxContentWidth = 0;

  for (const value of values) {
    const w = measureValueWidth(value, options?.numFmt);
    if (w > maxContentWidth) {
      maxContentWidth = w;
    }
  }

  // Apply font multiplier for proportional fonts
  let width = maxContentWidth * PROPORTIONAL_FONT_MULTIPLIER;

  // Apply bold multiplier if needed
  if (isBold) {
    width *= BOLD_MULTIPLIER;
  }

  // Add padding
  width += padding;

  // Round up to nearest 0.5 (Excel snaps to these increments)
  width = Math.ceil(width * 2) / 2;

  // Clamp to min/max
  if (width < minWidth) width = minWidth;
  if (width > maxWidth) width = maxWidth;

  return width;
}

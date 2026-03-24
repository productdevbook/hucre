// ── Styles Parser ────────────────────────────────────────────────────
// Parses xl/styles.xml — number formats, fonts, fills, borders, cell XFs.

import type { XmlElement } from "../xml/parser";
import type {
  CellStyle,
  FontStyle,
  FillStyle,
  BorderStyle,
  AlignmentStyle,
  Color,
  BorderSide,
  BorderLineStyle,
  FillPattern,
} from "../_types";
import { parseXml } from "../xml/parser";
import { isDateFormat } from "../_date";

// ── Types ────────────────────────────────────────────────────────────

export interface ParsedStyles {
  numFmts: Map<number, string>;
  fonts: FontStyle[];
  fills: FillStyle[];
  borders: BorderStyle[];
  cellXfs: CellXf[];
}

export interface CellXf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  alignment?: AlignmentStyle;
  applyNumberFormat?: boolean;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyAlignment?: boolean;
}

// ── Built-in Number Formats ──────────────────────────────────────────

const BUILTIN_NUM_FMTS: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yyyy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

// ── Date Format Detection ────────────────────────────────────────────

/** Built-in format IDs that represent date/time formats */
const DATE_FMT_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51,
  52, 53, 54, 55, 56, 57, 58,
]);

// ── Parser ───────────────────────────────────────────────────────────

/**
 * Parse xl/styles.xml and extract number formats, fonts, fills, borders,
 * and cell format records (cellXfs).
 */
export function parseStyles(xml: string): ParsedStyles {
  const doc = parseXml(xml);

  const numFmts = new Map<number, string>();
  const fonts: FontStyle[] = [];
  const fills: FillStyle[] = [];
  const borders: BorderStyle[] = [];
  const cellXfs: CellXf[] = [];

  for (const child of doc.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    switch (local) {
      case "numFmts":
        parseNumFmts(child, numFmts);
        break;
      case "fonts":
        parseFonts(child, fonts);
        break;
      case "fills":
        parseFills(child, fills);
        break;
      case "borders":
        parseBorders(child, borders);
        break;
      case "cellXfs":
        parseCellXfs(child, cellXfs);
        break;
    }
  }

  return { numFmts, fonts, fills, borders, cellXfs };
}

// ── Number Formats ───────────────────────────────────────────────────

function parseNumFmts(el: XmlElement, numFmts: Map<number, string>): void {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === "numFmt") {
      const id = Number(child.attrs["numFmtId"]);
      const code = child.attrs["formatCode"] ?? "";
      if (!Number.isNaN(id)) {
        numFmts.set(id, code);
      }
    }
  }
}

// ── Fonts ────────────────────────────────────────────────────────────

function parseFonts(el: XmlElement, fonts: FontStyle[]): void {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === "font") {
      fonts.push(parseFont(child));
    }
  }
}

function parseFont(el: XmlElement): FontStyle {
  const font: FontStyle = {};

  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    switch (local) {
      case "b":
        font.bold = child.attrs["val"] !== "0" && child.attrs["val"] !== "false";
        break;
      case "i":
        font.italic = child.attrs["val"] !== "0" && child.attrs["val"] !== "false";
        break;
      case "u": {
        const val = child.attrs["val"];
        if (val === "double") font.underline = "double";
        else if (val === "singleAccounting") font.underline = "singleAccounting";
        else if (val === "doubleAccounting") font.underline = "doubleAccounting";
        else font.underline = true;
        break;
      }
      case "strike":
        font.strikethrough = child.attrs["val"] !== "0" && child.attrs["val"] !== "false";
        break;
      case "sz":
        if (child.attrs["val"]) font.size = Number(child.attrs["val"]);
        break;
      case "name":
        if (child.attrs["val"]) font.name = child.attrs["val"];
        break;
      case "color":
        font.color = parseColor(child);
        break;
      case "vertAlign":
        if (child.attrs["val"] === "superscript" || child.attrs["val"] === "subscript") {
          font.vertAlign = child.attrs["val"];
        }
        break;
      case "family":
        if (child.attrs["val"]) font.family = Number(child.attrs["val"]);
        break;
      case "charset":
        if (child.attrs["val"]) font.charset = Number(child.attrs["val"]);
        break;
      case "scheme":
        if (
          child.attrs["val"] === "major" ||
          child.attrs["val"] === "minor" ||
          child.attrs["val"] === "none"
        ) {
          font.scheme = child.attrs["val"];
        }
        break;
    }
  }

  return font;
}

// ── Colors ───────────────────────────────────────────────────────────

function parseColor(el: XmlElement): Color {
  const color: Color = {};
  if (el.attrs["rgb"]) {
    // ARGB format — strip the alpha channel prefix (first 2 hex chars)
    const rgb = el.attrs["rgb"];
    color.rgb = rgb.length === 8 ? rgb.slice(2) : rgb;
  }
  if (el.attrs["theme"]) color.theme = Number(el.attrs["theme"]);
  if (el.attrs["tint"]) color.tint = Number(el.attrs["tint"]);
  if (el.attrs["indexed"]) color.indexed = Number(el.attrs["indexed"]);
  return color;
}

// ── Fills ────────────────────────────────────────────────────────────

function parseFills(el: XmlElement, fills: FillStyle[]): void {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === "fill") {
      fills.push(parseFill(child));
    }
  }
}

function parseFill(el: XmlElement): FillStyle {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    if (local === "patternFill") {
      return parsePatternFill(child);
    }
    if (local === "gradientFill") {
      return parseGradientFill(child);
    }
  }

  // Default: none pattern fill
  return { type: "pattern", pattern: "none" };
}

function parsePatternFill(el: XmlElement): FillStyle {
  const pattern = (el.attrs["patternType"] ?? "none") as FillPattern;
  const result: FillStyle = { type: "pattern", pattern };

  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === "fgColor") {
      (result as { fgColor?: Color }).fgColor = parseColor(child);
    } else if (local === "bgColor") {
      (result as { bgColor?: Color }).bgColor = parseColor(child);
    }
  }

  return result;
}

function parseGradientFill(el: XmlElement): FillStyle {
  const degree = el.attrs["degree"] ? Number(el.attrs["degree"]) : undefined;
  const stops: Array<{ position: number; color: Color }> = [];

  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === "stop") {
      const position = Number(child.attrs["position"] ?? "0");
      for (const stopChild of child.children) {
        if (typeof stopChild === "string") continue;
        const stopLocal = stopChild.local || stopChild.tag;
        if (stopLocal === "color") {
          stops.push({ position, color: parseColor(stopChild) });
        }
      }
    }
  }

  return { type: "gradient", degree, stops };
}

// ── Borders ──────────────────────────────────────────────────────────

function parseBorders(el: XmlElement, borders: BorderStyle[]): void {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === "border") {
      borders.push(parseBorder(child));
    }
  }
}

function parseBorder(el: XmlElement): BorderStyle {
  const border: BorderStyle = {};
  const diagonalUp = el.attrs["diagonalUp"];
  const diagonalDown = el.attrs["diagonalDown"];
  if (diagonalUp === "1" || diagonalUp === "true") border.diagonalUp = true;
  if (diagonalDown === "1" || diagonalDown === "true") border.diagonalDown = true;

  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    const side = parseBorderSide(child);
    if (!side) continue;

    switch (local) {
      case "left":
        border.left = side;
        break;
      case "right":
        border.right = side;
        break;
      case "top":
        border.top = side;
        break;
      case "bottom":
        border.bottom = side;
        break;
      case "diagonal":
        border.diagonal = side;
        break;
    }
  }

  return border;
}

function parseBorderSide(el: XmlElement): BorderSide | undefined {
  const style = el.attrs["style"] as BorderLineStyle | undefined;
  if (!style) return undefined;

  const side: BorderSide = { style };

  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === "color") {
      side.color = parseColor(child);
    }
  }

  return side;
}

// ── Cell XFs ─────────────────────────────────────────────────────────

function parseCellXfs(el: XmlElement, cellXfs: CellXf[]): void {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === "xf") {
      cellXfs.push(parseCellXf(child));
    }
  }
}

function parseCellXf(el: XmlElement): CellXf {
  const xf: CellXf = {
    numFmtId: Number(el.attrs["numFmtId"] ?? "0"),
    fontId: Number(el.attrs["fontId"] ?? "0"),
    fillId: Number(el.attrs["fillId"] ?? "0"),
    borderId: Number(el.attrs["borderId"] ?? "0"),
  };

  if (el.attrs["applyNumberFormat"] === "1" || el.attrs["applyNumberFormat"] === "true") {
    xf.applyNumberFormat = true;
  }
  if (el.attrs["applyFont"] === "1" || el.attrs["applyFont"] === "true") {
    xf.applyFont = true;
  }
  if (el.attrs["applyFill"] === "1" || el.attrs["applyFill"] === "true") {
    xf.applyFill = true;
  }
  if (el.attrs["applyBorder"] === "1" || el.attrs["applyBorder"] === "true") {
    xf.applyBorder = true;
  }
  if (el.attrs["applyAlignment"] === "1" || el.attrs["applyAlignment"] === "true") {
    xf.applyAlignment = true;
  }

  // Parse alignment child element
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === "alignment") {
      xf.alignment = parseAlignment(child);
    }
  }

  return xf;
}

function parseAlignment(el: XmlElement): AlignmentStyle {
  const align: AlignmentStyle = {};

  if (el.attrs["horizontal"]) {
    align.horizontal = el.attrs["horizontal"] as AlignmentStyle["horizontal"];
  }
  if (el.attrs["vertical"]) {
    align.vertical = el.attrs["vertical"] as AlignmentStyle["vertical"];
  }
  if (el.attrs["wrapText"] === "1" || el.attrs["wrapText"] === "true") {
    align.wrapText = true;
  }
  if (el.attrs["shrinkToFit"] === "1" || el.attrs["shrinkToFit"] === "true") {
    align.shrinkToFit = true;
  }
  if (el.attrs["textRotation"]) {
    align.textRotation = Number(el.attrs["textRotation"]);
  }
  if (el.attrs["indent"]) {
    align.indent = Number(el.attrs["indent"]);
  }
  if (el.attrs["readingOrder"]) {
    const ro = Number(el.attrs["readingOrder"]);
    if (ro === 1) align.readingOrder = "ltr";
    else if (ro === 2) align.readingOrder = "rtl";
    else align.readingOrder = "context";
  }

  return align;
}

// ── Style Resolution ─────────────────────────────────────────────────

/**
 * Resolve a cell style index (the `s` attribute on a cell) to a full CellStyle.
 * Returns an object only containing applied style properties.
 */
export function resolveStyle(styles: ParsedStyles, styleIndex: number): CellStyle {
  const xf = styles.cellXfs[styleIndex];
  if (!xf) return {};

  const result: CellStyle = {};

  // Number format
  if (xf.numFmtId !== 0) {
    const fmt = styles.numFmts.get(xf.numFmtId) ?? BUILTIN_NUM_FMTS[xf.numFmtId];
    if (fmt) {
      result.numFmt = fmt;
    }
  }

  // Font
  if (xf.fontId < styles.fonts.length && xf.fontId !== 0) {
    result.font = styles.fonts[xf.fontId];
  }

  // Fill — skip index 0 (none) and index 1 (gray125 default)
  if (xf.fillId < styles.fills.length && xf.fillId > 1) {
    result.fill = styles.fills[xf.fillId];
  }

  // Border
  if (xf.borderId < styles.borders.length && xf.borderId !== 0) {
    result.border = styles.borders[xf.borderId];
  }

  // Alignment
  if (xf.alignment) {
    result.alignment = xf.alignment;
  }

  return result;
}

/**
 * Check if a style index represents a date format.
 * Uses both built-in format ID checks and custom numFmt string analysis.
 */
export function isDateStyle(styles: ParsedStyles, styleIndex: number): boolean {
  const xf = styles.cellXfs[styleIndex];
  if (!xf) return false;

  const numFmtId = xf.numFmtId;

  // Check built-in date format IDs
  if (DATE_FMT_IDS.has(numFmtId)) {
    return true;
  }

  // Check custom number formats
  const customFmt = styles.numFmts.get(numFmtId);
  if (customFmt) {
    return isDateFormat(customFmt);
  }

  // Check built-in format string
  const builtinFmt = BUILTIN_NUM_FMTS[numFmtId];
  if (builtinFmt) {
    return isDateFormat(builtinFmt);
  }

  return false;
}

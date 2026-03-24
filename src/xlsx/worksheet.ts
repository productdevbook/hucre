// ── Worksheet Parser ─────────────────────────────────────────────────
// Parses xl/worksheets/sheetN.xml into a Sheet object.

import type { Sheet, Cell, CellValue, MergeRange, RichTextRun, FontStyle } from "../_types";
import type { SharedString } from "./shared-strings";
import type { ParsedStyles } from "./styles";
import { resolveStyle, isDateStyle } from "./styles";
import { serialToDate } from "../_date";
import { parseSax, decodeOoxmlEscapes } from "../xml/parser";

// ── Types ────────────────────────────────────────────────────────────

export interface WorksheetContext {
  sharedStrings: SharedString[];
  styles: ParsedStyles | null;
  readStyles: boolean;
  dateSystem: "1900" | "1904";
}

// ── Cell Reference Parsing ───────────────────────────────────────────

/**
 * Parse a cell reference like "A1", "Z1", "AA1", "AZ1", "AAA1"
 * into 0-based { row, col }.
 */
export function parseCellRef(ref: string): { row: number; col: number } {
  let i = 0;
  let col = 0;

  // Parse column letters
  while (i < ref.length) {
    const code = ref.charCodeAt(i);
    if (code >= 65 && code <= 90) {
      // A-Z
      col = col * 26 + (code - 64);
      i++;
    } else if (code >= 97 && code <= 122) {
      // a-z
      col = col * 26 + (code - 96);
      i++;
    } else {
      break;
    }
  }

  col--; // Convert to 0-based

  // Parse row number
  const row = Number(ref.slice(i)) - 1; // Convert to 0-based

  return { row, col };
}

/**
 * Parse a range reference like "A1:B2" into start and end positions.
 */
function parseRangeRef(ref: string): MergeRange {
  const parts = ref.split(":");
  const start = parseCellRef(parts[0]);
  const end = parts.length > 1 ? parseCellRef(parts[1]) : start;

  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  };
}

// ── SAX-based Worksheet Parser ───────────────────────────────────────

/**
 * Parse a worksheet XML into a Sheet using SAX parsing for performance.
 * This avoids building a full DOM tree for large worksheets.
 */
export function parseWorksheet(xml: string, name: string, ctx: WorksheetContext): Sheet {
  const rows: CellValue[][] = [];
  const cells = new Map<string, Cell>();
  const merges: MergeRange[] = [];
  let maxCol = -1;
  let maxRow = -1;
  let hasCells = false;

  // SAX parsing state
  let inSheetData = false;
  let inRow = false;
  let inCell = false;
  let inValue = false;
  let inFormula = false;
  let inInlineStr = false;
  let inInlineT = false;
  let inMergeCells = false;

  // Rich text in inline strings
  let inInlineR = false;
  let inInlineRPr = false;
  let inInlineRT = false;

  // Current cell state
  let cellRef = "";
  let cellType = "";
  let cellStyleIndex = -1;
  let cellValueText = "";
  let cellFormulaText = "";
  let inlineText = "";

  // Inline rich text state
  let inlineRichText: RichTextRun[] = [];
  let currentRunText = "";
  let currentRunFont: FontStyle | undefined;
  let _fontPropTag = "";

  parseSax(xml, {
    onOpenTag(tag, attrs) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag;

      switch (local) {
        case "sheetData":
          inSheetData = true;
          break;
        case "row":
          if (inSheetData) inRow = true;
          break;
        case "c":
          if (inRow) {
            inCell = true;
            cellRef = attrs["r"] ?? "";
            cellType = attrs["t"] ?? "";
            cellStyleIndex = attrs["s"] ? Number(attrs["s"]) : -1;
            cellValueText = "";
            cellFormulaText = "";
            inlineText = "";
            inlineRichText = [];
          }
          break;
        case "v":
          if (inCell) inValue = true;
          break;
        case "f":
          if (inCell) inFormula = true;
          break;
        case "is":
          if (inCell) inInlineStr = true;
          break;
        case "t":
          if (inInlineStr && !inInlineR) {
            inInlineT = true;
          } else if (inInlineR) {
            inInlineRT = true;
          }
          break;
        case "r":
          if (inInlineStr) {
            inInlineR = true;
            currentRunText = "";
            currentRunFont = undefined;
          }
          break;
        case "rPr":
          if (inInlineR) {
            inInlineRPr = true;
            currentRunFont = {};
          }
          break;
        case "mergeCells":
          inMergeCells = true;
          break;
        case "mergeCell":
          if (inMergeCells && attrs["ref"]) {
            merges.push(parseRangeRef(attrs["ref"]));
          }
          break;
        default:
          // Handle font property tags inside rPr
          if (inInlineRPr && currentRunFont) {
            _fontPropTag = local;
            applyFontProp(currentRunFont, local, attrs);
          }
          break;
      }
    },

    onText(text) {
      if (inValue) {
        cellValueText += text;
      } else if (inFormula) {
        cellFormulaText += text;
      } else if (inInlineT) {
        inlineText += text;
      } else if (inInlineRT) {
        currentRunText += text;
      }
    },

    onCloseTag(tag) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag;

      switch (local) {
        case "sheetData":
          inSheetData = false;
          break;
        case "row":
          inRow = false;
          break;
        case "c":
          if (inCell) {
            processCell(
              cellRef,
              cellType,
              cellStyleIndex,
              cellValueText,
              cellFormulaText,
              inlineText,
              inlineRichText.length > 0 ? inlineRichText : undefined,
              ctx,
              rows,
              cells,
            );
            // Track max dimensions
            if (cellRef) {
              const pos = parseCellRef(cellRef);
              if (pos.col > maxCol) maxCol = pos.col;
              if (pos.row > maxRow) maxRow = pos.row;
              hasCells = true;
            }
            inCell = false;
          }
          break;
        case "v":
          inValue = false;
          break;
        case "f":
          inFormula = false;
          break;
        case "is":
          inInlineStr = false;
          break;
        case "t":
          if (inInlineRT) {
            inInlineRT = false;
          } else if (inInlineT) {
            inInlineT = false;
          }
          break;
        case "r":
          if (inInlineR) {
            const decodedRunText = decodeOoxmlEscapes(currentRunText);
            inlineRichText.push(
              currentRunFont
                ? { text: decodedRunText, font: currentRunFont }
                : { text: decodedRunText },
            );
            inInlineR = false;
          }
          break;
        case "rPr":
          inInlineRPr = false;
          break;
        case "mergeCells":
          inMergeCells = false;
          break;
        default:
          if (inInlineRPr) {
            _fontPropTag = "";
          }
          break;
      }
    },
  });

  // Ensure all rows have consistent length
  if (hasCells) {
    const colCount = maxCol + 1;
    for (let r = 0; r <= maxRow; r++) {
      if (!rows[r]) {
        rows[r] = Array.from({ length: colCount }, () => null) as CellValue[];
      } else {
        while (rows[r].length < colCount) {
          rows[r].push(null);
        }
      }
    }
  }

  const sheet: Sheet = {
    name,
    rows,
  };

  if (cells.size > 0) {
    sheet.cells = cells;
  }
  if (merges.length > 0) {
    sheet.merges = merges;
  }

  return sheet;
}

// ── Cell Processing ──────────────────────────────────────────────────

function processCell(
  ref: string,
  type: string,
  styleIndex: number,
  valueText: string,
  formulaText: string,
  inlineText: string,
  inlineRichText: RichTextRun[] | undefined,
  ctx: WorksheetContext,
  rows: CellValue[][],
  cells: Map<string, Cell>,
): void {
  if (!ref) return;

  const pos = parseCellRef(ref);
  const { row, col } = pos;

  // Ensure row array exists
  while (rows.length <= row) {
    rows.push([]);
  }
  while (rows[row].length <= col) {
    rows[row].push(null);
  }

  let value: CellValue = null;
  let cellType: Cell["type"] = "empty";
  let formula: string | undefined;
  let formulaResult: CellValue | undefined;
  let richText: RichTextRun[] | undefined;

  // Handle formula
  if (formulaText) {
    formula = formulaText;
  }

  // Determine cell value based on type
  switch (type) {
    case "s": {
      // Shared string
      const idx = Number(valueText);
      if (!Number.isNaN(idx) && idx >= 0 && idx < ctx.sharedStrings.length) {
        const ss = ctx.sharedStrings[idx];
        value = ss.text;
        if (ss.richText && ss.richText.length > 0) {
          richText = ss.richText;
          cellType = "richText";
        } else {
          cellType = "string";
        }
      } else {
        value = valueText;
        cellType = "string";
      }
      break;
    }
    case "str": {
      // Inline formula string result
      value = decodeOoxmlEscapes(valueText);
      cellType = formula ? "formula" : "string";
      break;
    }
    case "inlineStr": {
      // Inline string with <is> element
      if (inlineRichText && inlineRichText.length > 0) {
        value = inlineRichText.map((r) => r.text).join("");
        richText = inlineRichText;
        cellType = "richText";
      } else {
        value = decodeOoxmlEscapes(inlineText);
        cellType = "string";
      }
      break;
    }
    case "b": {
      // Boolean
      value = valueText === "1" || valueText.toLowerCase() === "true";
      cellType = "boolean";
      break;
    }
    case "e": {
      // Error
      value = valueText;
      cellType = "error";
      break;
    }
    case "n":
    default: {
      // Number (explicit or implied)
      if (valueText === "" && !formula) {
        // Empty cell
        value = null;
        cellType = "empty";
        break;
      }

      const num = Number(valueText);
      if (!Number.isNaN(num) && valueText !== "") {
        // Check if this is a date via style
        if (ctx.styles && styleIndex >= 0 && isDateStyle(ctx.styles, styleIndex)) {
          value = serialToDate(num, ctx.dateSystem === "1904");
          cellType = "date";
        } else {
          value = num;
          cellType = "number";
        }
      } else if (valueText !== "") {
        // Non-numeric value text (shouldn't happen, but be safe)
        value = valueText;
        cellType = "string";
      }

      if (formula) {
        formulaResult = value;
        cellType = "formula";
      }
      break;
    }
  }

  // Set the value in the rows array
  rows[row][col] = value;

  // Build Cell object if there's detail beyond the raw value
  const hasDetails =
    formula !== undefined ||
    richText !== undefined ||
    (ctx.readStyles && ctx.styles && styleIndex >= 0) ||
    cellType === "error" ||
    cellType === "formula" ||
    cellType === "richText";

  if (hasDetails) {
    const cell: Cell = {
      value,
      type: cellType,
    };
    if (formula) {
      cell.formula = formula;
      if (formulaResult !== undefined) {
        cell.formulaResult = formulaResult;
      }
    }
    if (richText) {
      cell.richText = richText;
    }
    if (ctx.readStyles && ctx.styles && styleIndex >= 0) {
      const style = resolveStyle(ctx.styles, styleIndex);
      if (Object.keys(style).length > 0) {
        cell.style = style;
      }
    }
    cells.set(`${row},${col}`, cell);
  }
}

// ── Inline Rich Text Font Properties ─────────────────────────────────

function applyFontProp(font: FontStyle, tag: string, attrs: Record<string, string>): void {
  switch (tag) {
    case "b":
      font.bold = attrs["val"] !== "0" && attrs["val"] !== "false";
      break;
    case "i":
      font.italic = attrs["val"] !== "0" && attrs["val"] !== "false";
      break;
    case "u": {
      const val = attrs["val"];
      if (val === "double") font.underline = "double";
      else font.underline = true;
      break;
    }
    case "strike":
      font.strikethrough = attrs["val"] !== "0" && attrs["val"] !== "false";
      break;
    case "sz":
      if (attrs["val"]) font.size = Number(attrs["val"]);
      break;
    case "rFont":
      if (attrs["val"]) font.name = attrs["val"];
      break;
    case "color":
      font.color = {};
      if (attrs["rgb"]) {
        const rgb = attrs["rgb"];
        font.color.rgb = rgb.length === 8 ? rgb.slice(2) : rgb;
      }
      if (attrs["theme"]) font.color.theme = Number(attrs["theme"]);
      if (attrs["tint"]) font.color.tint = Number(attrs["tint"]);
      if (attrs["indexed"]) font.color.indexed = Number(attrs["indexed"]);
      break;
    case "vertAlign":
      if (attrs["val"] === "superscript" || attrs["val"] === "subscript") {
        font.vertAlign = attrs["val"];
      }
      break;
    case "family":
      if (attrs["val"]) font.family = Number(attrs["val"]);
      break;
    case "charset":
      if (attrs["val"]) font.charset = Number(attrs["val"]);
      break;
    case "scheme":
      if (attrs["val"] === "major" || attrs["val"] === "minor" || attrs["val"] === "none") {
        font.scheme = attrs["val"];
      }
      break;
  }
}

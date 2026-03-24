// ── Worksheet Parser ─────────────────────────────────────────────────
// Parses xl/worksheets/sheetN.xml into a Sheet object.

import type {
  Sheet,
  Cell,
  CellValue,
  MergeRange,
  RichTextRun,
  FontStyle,
  Hyperlink,
  DataValidation,
  ValidationType,
  ValidationOperator,
} from "../_types";
import type { SharedString } from "./shared-strings";
import type { ParsedStyles } from "./styles";
import type { Relationship } from "./relationships";
import { resolveStyle, isDateStyle } from "./styles";
import { serialToDate } from "../_date";
import { parseSax, decodeOoxmlEscapes } from "../xml/parser";

// ── Types ────────────────────────────────────────────────────────────

export interface WorksheetContext {
  sharedStrings: SharedString[];
  styles: ParsedStyles | null;
  readStyles: boolean;
  dateSystem: "1900" | "1904";
  /** Worksheet-level relationships (from xl/worksheets/_rels/sheetN.xml.rels) */
  worksheetRels?: Relationship[];
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

  // Hyperlinks parsed from <hyperlinks> section
  interface RawHyperlink {
    ref: string;
    rId?: string;
    location?: string;
    tooltip?: string;
    display?: string;
  }
  const rawHyperlinks: RawHyperlink[] = [];

  // Data validations parsed from <dataValidations> section
  const dataValidations: DataValidation[] = [];

  // SAX parsing state
  let inSheetData = false;
  let inRow = false;
  let inCell = false;
  let inValue = false;
  let inFormula = false;
  let inInlineStr = false;
  let inInlineT = false;
  let inMergeCells = false;
  let inHyperlinks = false;
  let inDataValidations = false;
  let inDataValidation = false;
  let inDvFormula1 = false;
  let inDvFormula2 = false;

  // Current data validation state
  let dvFormula1Text = "";
  let dvFormula2Text = "";
  let dvAttrs: Record<string, string> = {};

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
        case "hyperlinks":
          inHyperlinks = true;
          break;
        case "hyperlink":
          if (inHyperlinks && attrs["ref"]) {
            const hl: RawHyperlink = { ref: attrs["ref"] };
            // r:id for external hyperlinks
            const rId = attrs["r:id"] ?? attrs["R:id"];
            if (rId) hl.rId = rId;
            if (attrs["location"]) hl.location = attrs["location"];
            if (attrs["tooltip"]) hl.tooltip = attrs["tooltip"];
            if (attrs["display"]) hl.display = attrs["display"];
            rawHyperlinks.push(hl);
          }
          break;
        case "dataValidations":
          inDataValidations = true;
          break;
        case "dataValidation":
          if (inDataValidations) {
            inDataValidation = true;
            dvAttrs = { ...attrs };
            dvFormula1Text = "";
            dvFormula2Text = "";
          }
          break;
        case "formula1":
          if (inDataValidation) inDvFormula1 = true;
          break;
        case "formula2":
          if (inDataValidation) inDvFormula2 = true;
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
      } else if (inDvFormula1) {
        dvFormula1Text += text;
      } else if (inDvFormula2) {
        dvFormula2Text += text;
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
        case "hyperlinks":
          inHyperlinks = false;
          break;
        case "dataValidations":
          inDataValidations = false;
          break;
        case "dataValidation":
          if (inDataValidation) {
            const dv = buildDataValidation(dvAttrs, dvFormula1Text, dvFormula2Text);
            if (dv) {
              dataValidations.push(dv);
            }
            inDataValidation = false;
          }
          break;
        case "formula1":
          inDvFormula1 = false;
          break;
        case "formula2":
          inDvFormula2 = false;
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

  // ── Resolve hyperlinks ──
  // Build a map of rId → target URL from worksheet relationships
  const relMap = new Map<string, string>();
  if (ctx.worksheetRels) {
    for (const rel of ctx.worksheetRels) {
      relMap.set(rel.id, rel.target);
    }
  }

  for (const hl of rawHyperlinks) {
    const pos = parseCellRef(hl.ref);
    const key = `${pos.row},${pos.col}`;

    // Get or create cell in the cells map
    let cell = cells.get(key);
    if (!cell) {
      cell = {
        value: (rows[pos.row] && rows[pos.row][pos.col]) ?? null,
        type: "string",
      };
      cells.set(key, cell);
    }

    const hyperlink: Hyperlink = { target: "" };

    if (hl.location) {
      // Internal hyperlink
      hyperlink.location = hl.location;
      hyperlink.target = hl.location;
    } else if (hl.rId) {
      // External hyperlink — resolve from relationships
      const target = relMap.get(hl.rId);
      if (target) {
        hyperlink.target = target;
      }
    }

    if (hl.tooltip) hyperlink.tooltip = hl.tooltip;
    if (hl.display) hyperlink.display = hl.display;

    cell.hyperlink = hyperlink;
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
  if (dataValidations.length > 0) {
    sheet.dataValidations = dataValidations;
  }

  return sheet;
}

// ── Data Validation Builder ─────────────────────────────────────────

const VALID_TYPES = new Set<string>([
  "list",
  "whole",
  "decimal",
  "date",
  "time",
  "textLength",
  "custom",
]);
const VALID_OPERATORS = new Set<string>([
  "between",
  "notBetween",
  "equal",
  "notEqual",
  "greaterThan",
  "lessThan",
  "greaterThanOrEqual",
  "lessThanOrEqual",
]);

function buildDataValidation(
  attrs: Record<string, string>,
  formula1Text: string,
  formula2Text: string,
): DataValidation | null {
  const typeStr = attrs["type"];
  if (!typeStr || !VALID_TYPES.has(typeStr)) return null;

  const sqref = attrs["sqref"];
  if (!sqref) return null;

  const dv: DataValidation = {
    type: typeStr as ValidationType,
    range: sqref,
  };

  // Operator
  const operatorStr = attrs["operator"];
  if (operatorStr && VALID_OPERATORS.has(operatorStr)) {
    dv.operator = operatorStr as ValidationOperator;
  }

  // Boolean flags (XLSX uses "1" for true)
  if (attrs["allowBlank"] === "1" || attrs["allowBlank"] === "true") {
    dv.allowBlank = true;
  }
  if (attrs["showInputMessage"] === "1" || attrs["showInputMessage"] === "true") {
    dv.showInputMessage = true;
  }
  if (attrs["showErrorMessage"] === "1" || attrs["showErrorMessage"] === "true") {
    dv.showErrorMessage = true;
  }

  // Error style
  const errorStyle = attrs["errorStyle"];
  if (errorStyle === "stop" || errorStyle === "warning" || errorStyle === "information") {
    dv.errorStyle = errorStyle;
  }

  // Input/error messages (XLSX uses promptTitle/prompt for input messages)
  if (attrs["promptTitle"]) dv.inputTitle = attrs["promptTitle"];
  if (attrs["prompt"]) dv.inputMessage = attrs["prompt"];
  if (attrs["errorTitle"]) dv.errorTitle = attrs["errorTitle"];
  if (attrs["error"]) dv.errorMessage = attrs["error"];

  // Formulas
  if (formula1Text) {
    if (typeStr === "list") {
      // Check if formula1 is a quoted comma-separated list: "val1,val2,val3"
      const trimmed = formula1Text.trim();
      if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
        // Quoted list — parse into values array
        const inner = trimmed.slice(1, -1);
        dv.values = inner.split(",");
      } else {
        // Formula reference (e.g. Sheet2!$A$1:$A$10)
        dv.formula1 = formula1Text;
      }
    } else {
      dv.formula1 = formula1Text;
    }
  }

  if (formula2Text) {
    dv.formula2 = formula2Text;
  }

  return dv;
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

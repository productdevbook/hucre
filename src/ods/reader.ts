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
} from "../_types";
import { ParseError, ZipError } from "../errors";
import { ZipReader } from "../zip/reader";
import { parseXml } from "../xml/parser";
import type { XmlElement } from "../xml/parser";

// ── Helpers ─────────────────────────────────────────────────────────

function toUint8Array(input: ReadInput): Uint8Array {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  throw new ParseError("Unsupported input type. Expected Uint8Array or ArrayBuffer.");
}

function decodeUtf8(data: Uint8Array): string {
  return new TextDecoder("utf-8").decode(data);
}

function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === localName) return child;
  }
  return undefined;
}

function findChildren(el: XmlElement, localName: string): XmlElement[] {
  const result: XmlElement[] = [];
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === localName) result.push(child);
  }
  return result;
}

// ── Style Parsing ───────────────────────────────────────────────────

interface OdsStyleDef {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontColor?: string; // hex with '#' prefix
  backgroundColor?: string; // hex with '#' prefix
}

function parseStyles(doc: XmlElement): Map<string, OdsStyleDef> {
  const styles = new Map<string, OdsStyleDef>();

  // Styles live in <office:automatic-styles>
  const autoStyles = findChild(doc, "automatic-styles");
  if (!autoStyles) return styles;

  const styleElements = findChildren(autoStyles, "style");
  for (const styleEl of styleElements) {
    const family = styleEl.attrs["style:family"];
    if (family !== "table-cell") continue;

    const name = styleEl.attrs["style:name"];
    if (!name) continue;

    const def: OdsStyleDef = {};

    // Parse text properties
    const textProps = findChild(styleEl, "text-properties");
    if (textProps) {
      if (textProps.attrs["fo:font-weight"] === "bold") {
        def.bold = true;
      }
      if (textProps.attrs["fo:font-style"] === "italic") {
        def.italic = true;
      }
      const fontSize = textProps.attrs["fo:font-size"];
      if (fontSize) {
        // Parse "12pt" → 12
        const match = fontSize.match(/^(\d+(?:\.\d+)?)/);
        if (match) def.fontSize = parseFloat(match[1]);
      }
      const color = textProps.attrs["fo:color"];
      if (color) {
        def.fontColor = color;
      }
    }

    // Parse cell properties (background)
    const cellProps = findChild(styleEl, "table-cell-properties");
    if (cellProps) {
      const bgColor = cellProps.attrs["fo:background-color"];
      if (bgColor && bgColor !== "transparent") {
        def.backgroundColor = bgColor;
      }
    }

    styles.set(name, def);
  }

  return styles;
}

/** Convert a parsed ODS style def into a CellStyle */
function odsStyleToCellStyle(def: OdsStyleDef): CellStyle {
  const style: CellStyle = {};

  if (def.bold || def.italic || def.fontSize || def.fontColor) {
    style.font = {};
    if (def.bold) style.font.bold = true;
    if (def.italic) style.font.italic = true;
    if (def.fontSize) style.font.size = def.fontSize;
    if (def.fontColor) {
      // Strip '#' prefix for the rgb field
      const hex = def.fontColor.startsWith("#") ? def.fontColor.slice(1) : def.fontColor;
      style.font.color = { rgb: hex.toUpperCase() };
    }
  }

  if (def.backgroundColor) {
    const hex = def.backgroundColor.startsWith("#")
      ? def.backgroundColor.slice(1)
      : def.backgroundColor;
    style.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { rgb: hex.toUpperCase() },
    };
  }

  return style;
}

// ── Hyperlink Parsing ───────────────────────────────────────────────

/** Extract text and hyperlink from a cell's children */
function extractTextAndHyperlink(cell: XmlElement): { text: string; hyperlink?: Hyperlink } {
  const textP = findChild(cell, "p");
  if (!textP) return { text: "" };

  // Look for <text:a> elements inside <text:p>
  for (const child of textP.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === "a") {
      const href = child.attrs["xlink:href"];
      const text = child.children.filter((c: unknown) => typeof c === "string").join("");
      if (href) {
        return {
          text,
          hyperlink: { target: href, display: text },
        };
      }
    }
  }

  // No hyperlink — collect all text content (including from nested elements)
  const text = collectText(textP);
  return { text };
}

/** Recursively collect text from an element and its children,
 *  handling ODS special elements: text:span, text:s, text:line-break, text:tab */
function collectText(el: XmlElement): string {
  let text = "";
  for (const child of el.children) {
    if (typeof child === "string") {
      text += child;
    } else {
      const local = child.local || child.tag;
      if (local === "s") {
        // <text:s/> or <text:s text:c="N"/> — space characters
        const count = Number(child.attrs["text:c"] ?? "1");
        text += " ".repeat(count > 0 ? count : 1);
      } else if (local === "line-break") {
        // <text:line-break/> — newline
        text += "\n";
      } else if (local === "tab") {
        // <text:tab/> — tab character
        text += "\t";
      } else {
        // <text:span> and any other element — recurse into children
        text += collectText(child);
      }
    }
  }
  return text;
}

// ── Formula Parsing ─────────────────────────────────────────────────

/**
 * Convert an ODS formula to Excel-style formula.
 * ODS: "of:=SUM([.A1:.A10])" → "SUM(A1:A10)"
 */
function odsFormulaToExcel(formula: string): string {
  // Strip "of:=" or "oooc:=" prefix
  let f = formula;
  if (f.startsWith("of:=")) f = f.slice(4);
  else if (f.startsWith("oooc:=")) f = f.slice(6);
  else if (f.startsWith("=")) f = f.slice(1);

  // Convert [.A1:.B2] → A1:B2 and [.A1] → A1
  f = f.replace(/\[\.([^\]:.]+)(?::\.([^\]]+))?\]/g, (_match, ref1: string, ref2?: string) => {
    if (ref2) return `${ref1}:${ref2}`;
    return ref1;
  });

  return f;
}

// ── Cell Value Parsing ──────────────────────────────────────────────

function parseCellValue(cell: XmlElement): CellValue {
  const valueType = cell.attrs["office:value-type"] ?? cell.attrs["calcext:value-type"] ?? "";

  switch (valueType) {
    case "string": {
      // Get text from <text:p> children, including from nested <text:a> elements
      const { text } = extractTextAndHyperlink(cell);
      if (text) return text;
      // Check office:string-value attribute
      const strVal = cell.attrs["office:string-value"];
      if (strVal !== undefined) return strVal;
      return "";
    }

    case "float":
    case "currency":
    case "percentage": {
      const val = cell.attrs["office:value"];
      if (val !== undefined) return Number(val);
      return null;
    }

    case "boolean": {
      const boolVal = cell.attrs["office:boolean-value"];
      if (boolVal === "true") return true;
      if (boolVal === "false") return false;
      return null;
    }

    case "date": {
      const dateVal = cell.attrs["office:date-value"];
      if (dateVal) {
        const d = new Date(dateVal);
        if (!Number.isNaN(d.getTime())) return d;
      }
      return null;
    }

    case "time": {
      // ODS time values are ISO 8601 durations like PT12H30M
      const timeVal = cell.attrs["office:time-value"];
      if (timeVal) {
        return timeVal;
      }
      return null;
    }

    default: {
      // No explicit type — try to extract text
      const { text } = extractTextAndHyperlink(cell);
      if (text) return text;
      return null;
    }
  }
}

// ── Content XML Parsing ─────────────────────────────────────────────

function parseContentXml(xml: string, options?: ReadOptions): Sheet[] {
  const doc = parseXml(xml);
  const sheets: Sheet[] = [];

  // Parse styles for use when readStyles is enabled
  const readStyles = options?.readStyles ?? false;
  const styleDefs = readStyles ? parseStyles(doc) : new Map<string, OdsStyleDef>();

  // Navigate: document-content > body > spreadsheet > table
  const body = findChild(doc, "body");
  if (!body) return sheets;

  const spreadsheet = findChild(body, "spreadsheet");
  if (!spreadsheet) return sheets;

  const tables = findChildren(spreadsheet, "table");

  for (const table of tables) {
    const name = table.attrs["table:name"] ?? `Sheet${sheets.length + 1}`;

    // Filter sheets if specified
    if (options?.sheets && options.sheets.length > 0) {
      const shouldRead = options.sheets.some((spec) => {
        if (typeof spec === "string") return spec === name;
        if (typeof spec === "number") return spec === sheets.length;
        return false;
      });
      if (!shouldRead) {
        sheets.push({ name, rows: [] }); // placeholder to maintain index
        continue;
      }
    }

    const rows: CellValue[][] = [];
    const merges: MergeRange[] = [];
    const cells = new Map<string, Cell>();
    const tableRows = findChildren(table, "table-row");

    let currentRow = 0;

    for (const tableRow of tableRows) {
      const rowRepeat = Number(tableRow.attrs["table:number-rows-repeated"] ?? "1");

      // Collect cell entries with their repeat counts first,
      // so we can trim trailing nulls before expanding
      const cellEntries: Array<{
        value: CellValue;
        repeat: number;
        colSpan: number;
        rowSpan: number;
        isCovered: boolean;
        styleName?: string;
        formula?: string;
        hyperlink?: Hyperlink;
      }> = [];

      for (const child of tableRow.children) {
        if (typeof child === "string") continue;
        const local = child.local || child.tag;

        if (local === "table-cell") {
          const colRepeat = Number(child.attrs["table:number-columns-repeated"] ?? "1");
          const colSpan = Number(child.attrs["table:number-columns-spanned"] ?? "1");
          const rowSpan = Number(child.attrs["table:number-rows-spanned"] ?? "1");
          const value = parseCellValue(child);
          const styleName = child.attrs["table:style-name"];
          const formulaAttr = child.attrs["table:formula"];
          const formula = formulaAttr ? odsFormulaToExcel(formulaAttr) : undefined;
          const { hyperlink } = extractTextAndHyperlink(child);
          cellEntries.push({
            value,
            repeat: colRepeat,
            colSpan,
            rowSpan,
            isCovered: false,
            styleName,
            formula,
            hyperlink,
          });
        } else if (local === "covered-table-cell") {
          const colRepeat = Number(child.attrs["table:number-columns-repeated"] ?? "1");
          cellEntries.push({
            value: null,
            repeat: colRepeat,
            colSpan: 1,
            rowSpan: 1,
            isCovered: true,
          });
        }
      }

      // Trim trailing null/empty entries (avoids expanding huge repeat counts like 16384)
      while (
        cellEntries.length > 0 &&
        cellEntries[cellEntries.length - 1].value === null &&
        cellEntries[cellEntries.length - 1].colSpan === 1 &&
        cellEntries[cellEntries.length - 1].rowSpan === 1 &&
        !cellEntries[cellEntries.length - 1].styleName &&
        !cellEntries[cellEntries.length - 1].formula &&
        !cellEntries[cellEntries.length - 1].hyperlink
      ) {
        cellEntries.pop();
      }

      // Expand into row data and collect metadata
      const rowData: CellValue[] = [];
      let col = 0;

      for (const entry of cellEntries) {
        for (let r = 0; r < entry.repeat; r++) {
          rowData.push(entry.value);

          // Collect merge ranges
          if (entry.colSpan > 1 || entry.rowSpan > 1) {
            merges.push({
              startRow: currentRow,
              startCol: col,
              endRow: currentRow + entry.rowSpan - 1,
              endCol: col + entry.colSpan - 1,
            });
          }

          // Collect cell metadata (formulas, hyperlinks, styles)
          const hasMetadata =
            entry.formula ||
            entry.hyperlink ||
            (readStyles && entry.styleName && styleDefs.has(entry.styleName));

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
            };

            if (entry.formula) {
              cellData.formula = entry.formula;
              cellData.type = "formula";
            }
            if (entry.hyperlink) {
              cellData.hyperlink = entry.hyperlink;
            }
            if (readStyles && entry.styleName) {
              const styleDef = styleDefs.get(entry.styleName);
              if (styleDef) {
                cellData.style = odsStyleToCellStyle(styleDef);
              }
            }

            cells.set(`${currentRow},${col}`, cellData);
          }

          col++;
        }
      }

      // Cap row repeats for empty rows to avoid memory issues
      // (LibreOffice may emit large row repeats for trailing empty rows)
      const effectiveRowRepeat = rowData.length > 0 ? rowRepeat : 0;

      for (let r = 0; r < effectiveRowRepeat; r++) {
        rows.push(effectiveRowRepeat === 1 && r === 0 ? rowData : [...rowData]);
        if (r > 0 && merges.length > 0) {
          // For repeated rows with merges, we'd need to duplicate merge info
          // but this is an edge case; repeated rows with merges are uncommon
        }
        currentRow++;
      }

      if (effectiveRowRepeat === 0) {
        // Still advance row counter for empty repeated rows
        currentRow += rowRepeat;
      }
    }

    // Trim trailing empty rows
    while (rows.length > 0 && rows[rows.length - 1].length === 0) {
      rows.pop();
    }

    const sheet: Sheet = { name, rows };

    if (merges.length > 0) {
      sheet.merges = merges;
    }

    if (cells.size > 0) {
      sheet.cells = cells;
    }

    sheets.push(sheet);
  }

  // If filter was applied, remove placeholder sheets with empty rows
  if (options?.sheets && options.sheets.length > 0) {
    return sheets.filter(
      (s) =>
        s.rows.length > 0 ||
        s.merges !== undefined ||
        s.cells !== undefined ||
        options.sheets!.some((spec) => {
          if (typeof spec === "string") return spec === s.name;
          return false;
        }),
    );
  }

  return sheets;
}

// ── Meta XML Parsing ────────────────────────────────────────────────

function parseMetaXml(xml: string): Partial<WorkbookProperties> {
  const doc = parseXml(xml);
  const props: Partial<WorkbookProperties> = {};

  // Navigate to office:meta element
  const meta = findChild(doc, "meta");
  if (!meta) return props;

  for (const child of meta.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    const text = child.children.filter((c: unknown) => typeof c === "string").join("");

    switch (local) {
      case "title":
        if (text) props.title = text;
        break;
      case "subject":
        if (text) props.subject = text;
        break;
      case "initial-creator":
        if (text) props.creator = text;
        break;
      case "description":
        if (text) props.description = text;
        break;
      case "keyword":
        if (text) props.keywords = text;
        break;
      case "creation-date":
        if (text) {
          const d = new Date(text);
          if (!Number.isNaN(d.getTime())) props.created = d;
        }
        break;
      case "date":
        if (text) {
          const d = new Date(text);
          if (!Number.isNaN(d.getTime())) props.modified = d;
        }
        break;
    }
  }

  return props;
}

// ── Main Reader ─────────────────────────────────────────────────────

/**
 * Read an ODS file and return a Workbook.
 * Input can be Uint8Array or ArrayBuffer.
 */
export async function readOds(input: ReadInput, options?: ReadOptions): Promise<Workbook> {
  const data = toUint8Array(input);

  // 1. Open ZIP archive
  let zip: ZipReader;
  try {
    zip = new ZipReader(data);
  } catch (err) {
    if (err instanceof ZipError) throw err;
    throw new ParseError("Failed to open ODS file: not a valid ZIP archive", undefined, {
      cause: err,
    });
  }

  // 2. Verify mimetype
  if (zip.has("mimetype")) {
    const mimeData = await zip.extract("mimetype");
    const mime = decodeUtf8(mimeData).trim();
    if (mime !== "application/vnd.oasis.opendocument.spreadsheet") {
      throw new ParseError(`Invalid ODS mimetype: ${mime}`);
    }
  }

  // 3. Parse content.xml (required)
  if (!zip.has("content.xml")) {
    throw new ParseError("Invalid ODS: missing content.xml");
  }
  const contentXml = decodeUtf8(await zip.extract("content.xml"));
  const sheets = parseContentXml(contentXml, options);

  // 4. Parse meta.xml (optional)
  let properties: WorkbookProperties | undefined;
  if (zip.has("meta.xml")) {
    const metaXml = decodeUtf8(await zip.extract("meta.xml"));
    const metaProps = parseMetaXml(metaXml);
    if (Object.keys(metaProps).length > 0) {
      properties = { ...metaProps };
    }
  }

  // 5. Build workbook
  const workbook: Workbook = {
    sheets,
  };

  if (properties) {
    workbook.properties = properties;
  }

  return workbook;
}

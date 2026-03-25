// ── Worksheet XML Writer ─────────────────────────────────────────────
// Generates xl/worksheets/sheetN.xml for an XLSX package.

import type {
  WriteSheet,
  CellValue,
  CellStyle,
  ConditionalRule,
  DataValidation,
  SheetProtection,
  PageSetup,
  PageMargins,
  HeaderFooter,
  PaperSize,
  RichTextRun,
  FontStyle,
  Color,
} from "../_types";
import type { StylesCollector } from "./styles-writer";
import { dateToSerial } from "../_date";
import { xmlDocument, xmlElement, xmlSelfClose, xmlEscape } from "../xml/writer";
import { calculateColumnWidth } from "./auto-width";
import { hashSheetPassword } from "./password";

// ── Hyperlink Relationship ────────────────────────────────────────

export interface HyperlinkRelationship {
  id: string;
  target: string;
}

export interface WorksheetResult {
  xml: string;
  hyperlinkRelationships: HyperlinkRelationship[];
  /** The rId used for the drawing reference (if sheet has images) */
  drawingRId: string | null;
  /** The rId used for legacy drawing (VML) reference (if sheet has comments) */
  legacyDrawingRId: string | null;
  /** The rId used for the comments file reference (if sheet has comments) */
  commentsRId: string | null;
  /** Whether this sheet has comments */
  hasComments: boolean;
  /** Table parts info: rId and global table index for each table */
  tables: Array<{ rId: string; globalTableIndex: number }>;
}

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// ── Column Letter Conversion ───────────────────────────────────────

/** Convert a 0-based column index to an Excel column letter (A, B, ... Z, AA, AB, ...) */
export function colToLetter(col: number): string {
  let result = "";
  let c = col;
  do {
    result = String.fromCharCode(65 + (c % 26)) + result;
    c = Math.floor(c / 26) - 1;
  } while (c >= 0);
  return result;
}

/** Build a cell reference like "A1" from 0-based row and col */
export function cellRef(row: number, col: number): string {
  return colToLetter(col) + (row + 1);
}

/** Build a range reference like "A1:D10" from 0-based coordinates */
export function rangeRef(
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): string {
  return `${cellRef(startRow, startCol)}:${cellRef(endRow, endCol)}`;
}

// ── Shared Strings Collector ───────────────────────────────────────

export interface SharedStringsCollector {
  /** Get or add a string, return its index */
  add(value: string): number;
  /** Get all strings in order */
  getAll(): string[];
  /** Get count */
  count(): number;
}

export function createSharedStrings(): SharedStringsCollector {
  const strings: string[] = [];
  const indexMap = new Map<string, number>();

  function add(value: string): number {
    const existing = indexMap.get(value);
    if (existing !== undefined) return existing;

    const idx = strings.length;
    strings.push(value);
    indexMap.set(value, idx);
    return idx;
  }

  function getAll(): string[] {
    return strings;
  }

  function count(): number {
    return strings.length;
  }

  return { add, getAll, count };
}

/** Generate xl/sharedStrings.xml */
export function writeSharedStringsXml(sharedStrings: SharedStringsCollector): string {
  const strings = sharedStrings.getAll();
  const count = strings.length;

  if (count === 0) {
    return xmlDocument("sst", { xmlns: NS_SPREADSHEET, count: 0, uniqueCount: 0 }, "");
  }

  const children: string[] = [];
  for (const str of strings) {
    const escaped = xmlEscape(str);
    const needsPreserve =
      str.length > 0 &&
      (str[0] === " " || str[str.length - 1] === " " || str.includes("\n") || str.includes("\t"));
    const tElement = needsPreserve
      ? `<t xml:space="preserve">${escaped}</t>`
      : xmlElement("t", undefined, escaped);
    children.push(xmlElement("si", undefined, [tElement]));
  }

  return xmlDocument("sst", { xmlns: NS_SPREADSHEET, count, uniqueCount: count }, children);
}

// ── Resolved Cell Data ─────────────────────────────────────────────

interface ResolvedCell {
  value: CellValue;
  style?: CellStyle;
  formula?: string;
  formulaResult?: CellValue;
  formulaType?: "shared" | "array";
  formulaSharedIndex?: number;
  formulaRef?: string;
  formulaDynamic?: boolean;
  richText?: RichTextRun[];
}

// ── Default date format ────────────────────────────────────────────

const DEFAULT_DATE_FORMAT = "yyyy-mm-dd";

/** Known Excel error value strings */
const EXCEL_ERRORS = new Set([
  "#VALUE!",
  "#REF!",
  "#N/A",
  "#NAME?",
  "#NULL!",
  "#DIV/0!",
  "#NUM!",
  "#GETTING_DATA",
]);

// ── Worksheet Writer ───────────────────────────────────────────────

/** Generate xl/worksheets/sheetN.xml along with any hyperlink relationships */
export function writeWorksheetXml(
  sheet: WriteSheet,
  styles: StylesCollector,
  sharedStrings: SharedStringsCollector,
  dateSystem?: "1900" | "1904",
  globalTableStartIndex?: number,
): WorksheetResult {
  const is1904 = dateSystem === "1904";

  // Resolve rows from data or rows
  const resolvedRows = resolveRows(sheet);
  const rowCount = resolvedRows.length;

  // Calculate max column count
  let maxCols = 0;
  for (const row of resolvedRows) {
    if (row.length > maxCols) maxCols = row.length;
  }
  if (sheet.columns && sheet.columns.length > maxCols) {
    maxCols = sheet.columns.length;
  }

  const parts: string[] = [];

  // ── SheetPr (tab color, outlinePr, etc.) — must come first per OOXML spec ──
  {
    const sheetPrChildren: string[] = [];
    if (sheet.view?.tabColor) {
      sheetPrChildren.push(xmlSelfClose("tabColor", serializeColorAttrs(sheet.view.tabColor)));
    }
    if (sheet.outlineProperties) {
      const outlineAttrs: Record<string, string | number | boolean> = {};
      if (sheet.outlineProperties.summaryBelow !== undefined) {
        outlineAttrs["summaryBelow"] = sheet.outlineProperties.summaryBelow ? 1 : 0;
      }
      if (sheet.outlineProperties.summaryRight !== undefined) {
        outlineAttrs["summaryRight"] = sheet.outlineProperties.summaryRight ? 1 : 0;
      }
      sheetPrChildren.push(xmlSelfClose("outlinePr", outlineAttrs));
    }
    if (sheetPrChildren.length > 0) {
      parts.push(xmlElement("sheetPr", undefined, sheetPrChildren));
    }
  }

  // ── Dimension (OOXML spec: after sheetPr, before sheetViews) ──
  if (rowCount > 0 || maxCols > 0) {
    const endRow = rowCount > 0 ? rowCount - 1 : 0;
    const endCol = maxCols > 0 ? maxCols - 1 : 0;
    parts.push(xmlSelfClose("dimension", { ref: rangeRef(0, 0, endRow, endCol) }));
  }

  // ── SheetViews (freeze panes, view settings) ──
  const sheetViewParts: string[] = [];

  if (sheet.freezePane) {
    const fp = sheet.freezePane;
    const topLeftCell = cellRef(fp.rows ?? 0, fp.columns ?? 0);
    const paneAttrs: Record<string, string | number> = {};

    if (fp.columns && fp.columns > 0) {
      paneAttrs["xSplit"] = fp.columns;
    }
    if (fp.rows && fp.rows > 0) {
      paneAttrs["ySplit"] = fp.rows;
    }
    paneAttrs["topLeftCell"] = topLeftCell;
    paneAttrs["state"] = "frozen";

    // Determine active pane
    const hasXSplit = fp.columns && fp.columns > 0;
    const hasYSplit = fp.rows && fp.rows > 0;

    if (hasXSplit && hasYSplit) {
      paneAttrs["activePane"] = "bottomRight";
    } else if (hasXSplit) {
      paneAttrs["activePane"] = "topRight";
    } else {
      paneAttrs["activePane"] = "bottomLeft";
    }

    sheetViewParts.push(xmlSelfClose("pane", paneAttrs));
  } else if (sheet.splitPane) {
    const sp = sheet.splitPane;
    const paneAttrs: Record<string, string | number> = {};

    const hasXSplit = sp.xSplit !== undefined && sp.xSplit > 0;
    const hasYSplit = sp.ySplit !== undefined && sp.ySplit > 0;

    if (hasXSplit) {
      paneAttrs["xSplit"] = sp.xSplit!;
    }
    if (hasYSplit) {
      paneAttrs["ySplit"] = sp.ySplit!;
    }
    paneAttrs["topLeftCell"] = "A1";
    paneAttrs["state"] = "split";

    if (hasXSplit && hasYSplit) {
      paneAttrs["activePane"] = "bottomRight";
    } else if (hasXSplit) {
      paneAttrs["activePane"] = "topRight";
    } else {
      paneAttrs["activePane"] = "bottomLeft";
    }

    sheetViewParts.push(xmlSelfClose("pane", paneAttrs));
  }

  const viewAttrs: Record<string, string | number> = {
    workbookViewId: 0,
  };

  if (sheet.view) {
    if (sheet.view.showGridLines === false) viewAttrs["showGridLines"] = 0;
    if (sheet.view.showRowColHeaders === false) viewAttrs["showRowColHeaders"] = 0;
    if (sheet.view.zoomScale !== undefined) viewAttrs["zoomScale"] = sheet.view.zoomScale;
    if (sheet.view.rightToLeft) viewAttrs["rightToLeft"] = 1;
  }

  parts.push(
    xmlElement("sheetViews", undefined, [
      sheetViewParts.length > 0
        ? xmlElement("sheetView", viewAttrs, sheetViewParts)
        : xmlSelfClose("sheetView", viewAttrs),
    ]),
  );

  // ── SheetFormatPr ──
  parts.push(xmlSelfClose("sheetFormatPr", { defaultRowHeight: 15 }));

  // ── Columns ──
  if (sheet.columns && sheet.columns.length > 0) {
    const colElements: string[] = [];
    for (let i = 0; i < sheet.columns.length; i++) {
      const col = sheet.columns[i];

      // Calculate auto-width if requested and no explicit width is set
      let effectiveWidth = col.width;
      if (col.autoWidth && effectiveWidth === undefined) {
        const columnValues: CellValue[] = [];
        for (const row of resolvedRows) {
          if (row && i < row.length && row[i]) {
            columnValues.push(row[i]!.value);
          }
        }
        effectiveWidth = calculateColumnWidth(columnValues, {
          font: col.style?.font,
          numFmt: col.numFmt ?? col.style?.numFmt,
        });
      }

      if (effectiveWidth !== undefined || col.hidden || col.outlineLevel || col.collapsed) {
        const colAttrs: Record<string, string | number | boolean> = {
          min: i + 1,
          max: i + 1,
        };
        if (effectiveWidth !== undefined) {
          colAttrs["width"] = effectiveWidth;
          colAttrs["customWidth"] = true;
        }
        if (col.hidden) {
          colAttrs["hidden"] = true;
        }
        if (col.outlineLevel) {
          colAttrs["outlineLevel"] = col.outlineLevel;
        }
        if (col.collapsed) {
          colAttrs["collapsed"] = true;
        }
        colElements.push(xmlSelfClose("col", colAttrs));
      }
    }
    if (colElements.length > 0) {
      parts.push(xmlElement("cols", undefined, colElements));
    }
  }

  // ── Sheet Data ──
  const rowElements: string[] = [];

  for (let r = 0; r < rowCount; r++) {
    const row = resolvedRows[r];
    const rowDef = sheet.rowDefs?.get(r);
    const hasRowDef =
      rowDef &&
      (rowDef.height !== undefined || rowDef.hidden || rowDef.outlineLevel || rowDef.collapsed);

    if ((!row || row.length === 0) && !hasRowDef) continue;

    const cellElements: string[] = [];
    let hasAnyCells = false;

    if (row) {
      for (let c = 0; c < row.length; c++) {
        const resolved = row[c];
        if (!resolved) continue;

        const cellXml = serializeCell(r, c, resolved, styles, sharedStrings, is1904);
        if (cellXml) {
          cellElements.push(cellXml);
          hasAnyCells = true;
        }
      }
    }

    if (hasAnyCells || hasRowDef) {
      const rowAttrs: Record<string, string | number | boolean> = { r: r + 1 };
      if (rowDef?.height !== undefined) {
        rowAttrs["ht"] = rowDef.height;
        rowAttrs["customHeight"] = 1;
      }
      if (rowDef?.hidden) {
        rowAttrs["hidden"] = 1;
      }
      if (rowDef?.outlineLevel) {
        rowAttrs["outlineLevel"] = rowDef.outlineLevel;
      }
      if (rowDef?.collapsed) {
        rowAttrs["collapsed"] = 1;
      }
      if (hasAnyCells) {
        rowElements.push(xmlElement("row", rowAttrs, cellElements));
      } else {
        rowElements.push(xmlSelfClose("row", rowAttrs));
      }
    }
  }

  parts.push(xmlElement("sheetData", undefined, rowElements.length > 0 ? rowElements : ""));

  // ── SheetProtection (OOXML: after sheetData, before autoFilter) ──
  if (sheet.protection) {
    parts.push(serializeSheetProtection(sheet.protection));
  }

  // ── Auto Filter (OOXML: after sheetProtection, before mergeCells) ──
  if (sheet.autoFilter) {
    if (sheet.autoFilter.columns && sheet.autoFilter.columns.length > 0) {
      const filterChildren: string[] = [];
      for (const col of sheet.autoFilter.columns) {
        if (col.filters && col.filters.length > 0) {
          const filterElements = col.filters.map((v) => xmlSelfClose("filter", { val: v }));
          filterChildren.push(
            xmlElement("filterColumn", { colId: col.colIndex }, [
              xmlElement("filters", undefined, filterElements),
            ]),
          );
        }
      }
      parts.push(xmlElement("autoFilter", { ref: sheet.autoFilter.range }, filterChildren));
    } else {
      parts.push(xmlSelfClose("autoFilter", { ref: sheet.autoFilter.range }));
    }
  }

  // ── Merge Cells ──
  if (sheet.merges && sheet.merges.length > 0) {
    const mergeElements = sheet.merges.map((m) =>
      xmlSelfClose("mergeCell", {
        ref: rangeRef(m.startRow, m.startCol, m.endRow, m.endCol),
      }),
    );
    parts.push(xmlElement("mergeCells", { count: sheet.merges.length }, mergeElements));
  }

  // ── Conditional Formatting ──
  if (sheet.conditionalRules && sheet.conditionalRules.length > 0) {
    parts.push(...serializeConditionalFormatting(sheet.conditionalRules, styles));
  }

  // ── Data Validations ──
  if (sheet.dataValidations && sheet.dataValidations.length > 0) {
    parts.push(serializeDataValidations(sheet.dataValidations));
  }

  // ── Hyperlinks ──
  const { xml: hyperlinksXml, relationships: hyperlinkRelationships } = collectHyperlinks(sheet);
  if (hyperlinksXml) {
    parts.push(hyperlinksXml);
  }

  // ── Print Options (only when pageSetup exists) ──
  if (sheet.pageSetup) {
    parts.push(xmlSelfClose("printOptions", { headings: 0, gridLines: 0 }));
  }

  // ── Page Margins ──
  parts.push(serializePageMargins(sheet.pageSetup?.margins));

  // ── Page Setup ──
  if (sheet.pageSetup) {
    parts.push(serializePageSetup(sheet.pageSetup));
  }

  // ── Header/Footer ──
  if (sheet.headerFooter) {
    parts.push(serializeHeaderFooter(sheet.headerFooter));
  }

  // ── Row Breaks ──
  if (sheet.rowBreaks && sheet.rowBreaks.length > 0) {
    const sorted = [...sheet.rowBreaks].sort((a, b) => a - b);
    const brkElements = sorted.map((row) =>
      xmlSelfClose("brk", { id: row + 1, max: 16383, man: 1 }),
    );
    parts.push(
      xmlElement(
        "rowBreaks",
        { count: sorted.length, manualBreakCount: sorted.length },
        brkElements,
      ),
    );
  }

  // ── Column Breaks ──
  if (sheet.colBreaks && sheet.colBreaks.length > 0) {
    const sorted = [...sheet.colBreaks].sort((a, b) => a - b);
    const brkElements = sorted.map((col) =>
      xmlSelfClose("brk", { id: col + 1, max: 1048575, man: 1 }),
    );
    parts.push(
      xmlElement(
        "colBreaks",
        { count: sorted.length, manualBreakCount: sorted.length },
        brkElements,
      ),
    );
  }

  // ── Drawing (images) ──
  let drawingRId: string | null = null;
  let nextRId = hyperlinkRelationships.length + 1;
  if (sheet.images && sheet.images.length > 0) {
    // Drawing rId comes after all hyperlink rIds
    drawingRId = `rId${nextRId}`;
    nextRId++;
    parts.push(xmlSelfClose("drawing", { "r:id": drawingRId }));
  }

  // ── Legacy Drawing (VML — for comments) ──
  let legacyDrawingRId: string | null = null;
  let commentsRId: string | null = null;
  let hasComments = false;
  if (sheet.cells) {
    for (const [, cell] of sheet.cells) {
      if (cell.comment) {
        hasComments = true;
        break;
      }
    }
  }
  if (hasComments) {
    legacyDrawingRId = `rId${nextRId}`;
    nextRId++;
    commentsRId = `rId${nextRId}`;
    nextRId++;
    parts.push(xmlSelfClose("legacyDrawing", { "r:id": legacyDrawingRId }));
  }

  // ── Table Parts ──
  const tableEntries: Array<{ rId: string; globalTableIndex: number }> = [];
  if (sheet.tables && sheet.tables.length > 0 && globalTableStartIndex !== undefined) {
    const tablePartElements: string[] = [];
    for (let t = 0; t < sheet.tables.length; t++) {
      const tableRId = `rId${nextRId}`;
      nextRId++;
      const globalIdx = globalTableStartIndex + t;
      tableEntries.push({ rId: tableRId, globalTableIndex: globalIdx });
      tablePartElements.push(xmlSelfClose("tablePart", { "r:id": tableRId }));
    }
    parts.push(xmlElement("tableParts", { count: sheet.tables.length }, tablePartElements));
  }

  return {
    xml: xmlDocument("worksheet", { xmlns: NS_SPREADSHEET, "xmlns:r": NS_R }, parts),
    hyperlinkRelationships,
    drawingRId,
    legacyDrawingRId,
    commentsRId,
    hasComments,
    tables: tableEntries,
  };
}

// ── Row Resolution ─────────────────────────────────────────────────

function resolveRows(sheet: WriteSheet): Array<Array<ResolvedCell | null>> {
  const resolved: Array<Array<ResolvedCell | null>> = [];

  if (sheet.data && sheet.columns) {
    // Object-based data with column keys
    const keys = sheet.columns.map((col) => col.key);

    // Add header row if columns have headers
    const hasHeaders = sheet.columns.some((col) => col.header);
    if (hasHeaders) {
      const headerRow: Array<ResolvedCell | null> = [];
      for (let c = 0; c < sheet.columns.length; c++) {
        const col = sheet.columns[c];
        headerRow.push({
          value: col.header ?? col.key ?? null,
          style: col.style,
        });
      }
      resolved.push(headerRow);
    }

    for (const obj of sheet.data) {
      const row: Array<ResolvedCell | null> = [];
      for (let c = 0; c < keys.length; c++) {
        const key = keys[c];
        const value = key !== undefined ? (obj[key] ?? null) : null;
        const col = sheet.columns[c];
        row.push({
          value,
          style: col.style,
          ...(col.numFmt && !col.style?.numFmt
            ? { style: { ...col.style, numFmt: col.numFmt } }
            : {}),
        });
      }
      resolved.push(row);
    }
  } else if (sheet.rows) {
    // Array-based rows
    for (const row of sheet.rows) {
      const resolvedRow: Array<ResolvedCell | null> = [];
      for (let c = 0; c < row.length; c++) {
        const value = row[c];
        resolvedRow.push({ value });
      }
      resolved.push(resolvedRow);
    }
  }

  // Apply cell overrides
  if (sheet.cells) {
    for (const [key, cellOverride] of sheet.cells) {
      const [rowStr, colStr] = key.split(",");
      const r = parseInt(rowStr, 10);
      const c = parseInt(colStr, 10);

      // Ensure row exists
      while (resolved.length <= r) {
        resolved.push([]);
      }
      const row = resolved[r];
      while (row.length <= c) {
        row.push(null);
      }

      const existing = row[c];
      row[c] = {
        value: cellOverride.value ?? existing?.value ?? null,
        style: cellOverride.style ?? existing?.style,
        formula: cellOverride.formula ?? existing?.formula,
        formulaResult: cellOverride.formulaResult ?? existing?.formulaResult,
        formulaType: cellOverride.formulaType ?? existing?.formulaType,
        formulaSharedIndex: cellOverride.formulaSharedIndex ?? existing?.formulaSharedIndex,
        formulaRef: cellOverride.formulaRef ?? existing?.formulaRef,
        formulaDynamic: cellOverride.formulaDynamic ?? existing?.formulaDynamic,
        richText: cellOverride.richText ?? existing?.richText,
      };
    }
  }

  return resolved;
}

// ── Cell Serialization ─────────────────────────────────────────────

function serializeCell(
  row: number,
  col: number,
  resolved: ResolvedCell,
  styles: StylesCollector,
  sharedStrings: SharedStringsCollector,
  is1904: boolean,
): string | null {
  const {
    value,
    style,
    formula,
    formulaResult,
    formulaType,
    formulaSharedIndex,
    formulaRef,
    formulaDynamic,
    richText,
  } = resolved;

  // Determine style index
  let styleIdx = 0;
  let effectiveStyle = style;

  // If value is Date and no numFmt specified, add default date format
  if (value instanceof Date && (!effectiveStyle || !effectiveStyle.numFmt)) {
    effectiveStyle = {
      ...effectiveStyle,
      numFmt: DEFAULT_DATE_FORMAT,
    };
  }

  if (effectiveStyle) {
    styleIdx = styles.addStyle(effectiveStyle);
  }

  const ref = cellRef(row, col);

  // Rich text cell (inline string)
  if (richText && richText.length > 0) {
    const cellAttrs: Record<string, string | number> = { r: ref, t: "inlineStr" };
    if (styleIdx !== 0) cellAttrs["s"] = styleIdx;
    const isContent = serializeRichTextRuns(richText);
    return xmlElement("c", cellAttrs, [xmlElement("is", undefined, isContent)]);
  }

  // Formula cell (including shared formula slave cells with empty formula text)
  if (formula !== undefined && formula !== null) {
    const cellAttrs: Record<string, string | number> = { r: ref };
    if (styleIdx !== 0) cellAttrs["s"] = styleIdx;

    // Build <f> element with appropriate attributes
    let fElement: string;
    if (formulaType === "shared") {
      const fAttrs: Record<string, string | number> = { t: "shared" };
      if (formulaSharedIndex !== undefined) fAttrs["si"] = formulaSharedIndex;
      if (formulaRef) fAttrs["ref"] = formulaRef;
      // Shared slave cell: no formula text → self-closing <f/>
      if (formula === "") {
        fElement = xmlSelfClose("f", fAttrs);
      } else {
        fElement = xmlElement("f", fAttrs, xmlEscape(formula));
      }
    } else if (formulaType === "array") {
      const fAttrs: Record<string, string | number> = { t: "array" };
      if (formulaRef) fAttrs["ref"] = formulaRef;
      if (formulaDynamic) fAttrs["cm"] = 1;
      fElement = xmlElement("f", fAttrs, xmlEscape(formula));
    } else {
      // Normal formula
      fElement = xmlElement("f", undefined, xmlEscape(formula));
    }

    const children: string[] = [fElement];

    // Cached formula result
    if (formulaResult !== undefined && formulaResult !== null) {
      if (typeof formulaResult === "string") {
        cellAttrs["t"] = "str";
        children.push(xmlElement("v", undefined, xmlEscape(formulaResult)));
      } else if (typeof formulaResult === "boolean") {
        cellAttrs["t"] = "b";
        children.push(xmlElement("v", undefined, formulaResult ? "1" : "0"));
      } else if (typeof formulaResult === "number") {
        children.push(xmlElement("v", undefined, String(formulaResult)));
      }
    }

    return xmlElement("c", cellAttrs, children);
  }

  // Null/undefined — skip if no style, otherwise emit empty cell with style
  if (value === null || value === undefined) {
    if (styleIdx !== 0) {
      return xmlSelfClose("c", { r: ref, s: styleIdx });
    }
    return null;
  }

  // Error value (e.g. #VALUE!, #REF!, #N/A, #NAME?, #NULL!, #DIV/0!, #NUM!)
  if (typeof value === "string" && EXCEL_ERRORS.has(value)) {
    const attrs: Record<string, string | number> = { r: ref, t: "e" };
    if (styleIdx !== 0) attrs["s"] = styleIdx;
    return xmlElement("c", attrs, [xmlElement("v", undefined, value)]);
  }

  // String value
  if (typeof value === "string") {
    const ssIdx = sharedStrings.add(value);
    const attrs: Record<string, string | number> = { r: ref, t: "s" };
    if (styleIdx !== 0) attrs["s"] = styleIdx;
    return xmlElement("c", attrs, [xmlElement("v", undefined, String(ssIdx))]);
  }

  // Number value
  if (typeof value === "number") {
    // Infinity, -Infinity, and NaN cannot be represented in OOXML — emit as empty cell
    if (!Number.isFinite(value)) {
      if (styleIdx !== 0) {
        return xmlSelfClose("c", { r: ref, s: styleIdx });
      }
      return null;
    }
    const attrs: Record<string, string | number> = { r: ref };
    if (styleIdx !== 0) attrs["s"] = styleIdx;
    return xmlElement("c", attrs, [xmlElement("v", undefined, String(value))]);
  }

  // Boolean value
  if (typeof value === "boolean") {
    const attrs: Record<string, string | number> = { r: ref, t: "b" };
    if (styleIdx !== 0) attrs["s"] = styleIdx;
    return xmlElement("c", attrs, [xmlElement("v", undefined, value ? "1" : "0")]);
  }

  // Date value
  if (value instanceof Date) {
    const serial = dateToSerial(value, is1904);
    const attrs: Record<string, string | number> = { r: ref };
    if (styleIdx !== 0) attrs["s"] = styleIdx;
    return xmlElement("c", attrs, [xmlElement("v", undefined, String(serial))]);
  }

  return null;
}

// ── Sheet Protection Serialization ────────────────────────────────

/**
 * Serialize a SheetProtection object into a `<sheetProtection>` XML element.
 *
 * XLSX attribute semantics:
 * - `sheet="1"` means the sheet IS protected
 * - `objects="1"` means objects ARE protected
 * - `scenarios="1"` means scenarios ARE protected
 * - All other boolean attrs (selectLockedCells, formatCells, etc.) use "1" to mean
 *   the action is PROHIBITED (not allowed).
 *
 * Our API uses intuitive "allow" booleans: `true` means the user CAN do it.
 * So we invert them when writing to XML: allow=true → attr="0" (not prohibited).
 */
function serializeSheetProtection(protection: SheetProtection): string {
  const attrs: Record<string, string | number> = {};

  // Password hash
  if (protection.password) {
    attrs["password"] = hashSheetPassword(protection.password);
  }

  // sheet, objects, scenarios: true means protected → "1"
  if (protection.sheet !== false) {
    // Default to protected when protection object exists
    attrs["sheet"] = 1;
  }
  if (protection.objects) {
    attrs["objects"] = 1;
  }
  if (protection.scenarios) {
    attrs["scenarios"] = 1;
  }

  // All other options: our API = "allow" booleans.
  // In XLSX, "1" = prohibited. So allow=true → "0", allow=false → "1".
  // We only emit the attribute if the user explicitly set it.
  const allowOptions: Array<[keyof SheetProtection, string]> = [
    ["selectLockedCells", "selectLockedCells"],
    ["selectUnlockedCells", "selectUnlockedCells"],
    ["formatCells", "formatCells"],
    ["formatColumns", "formatColumns"],
    ["formatRows", "formatRows"],
    ["insertColumns", "insertColumns"],
    ["insertRows", "insertRows"],
    ["insertHyperlinks", "insertHyperlinks"],
    ["deleteColumns", "deleteColumns"],
    ["deleteRows", "deleteRows"],
    ["sort", "sort"],
    ["autoFilter", "autoFilter"],
    ["pivotTables", "pivotTables"],
  ];

  for (const [prop, attr] of allowOptions) {
    const val = protection[prop];
    if (val !== undefined && typeof val === "boolean") {
      // true (allowed) → "0" (not prohibited), false (disallowed) → "1" (prohibited)
      attrs[attr] = val ? 0 : 1;
    }
  }

  return xmlSelfClose("sheetProtection", attrs);
}

// ── Data Validation Serialization ─────────────────────────────────

/** Serialize data validations into a `<dataValidations>` XML block */
function serializeDataValidations(validations: DataValidation[]): string {
  const dvElements: string[] = [];

  for (const dv of validations) {
    const attrs: Record<string, string | number> = {
      type: dv.type,
      sqref: dv.range,
    };

    if (dv.operator) {
      attrs["operator"] = dv.operator;
    }
    if (dv.allowBlank) {
      attrs["allowBlank"] = 1;
    }
    if (dv.showInputMessage) {
      attrs["showInputMessage"] = 1;
    }
    if (dv.showErrorMessage) {
      attrs["showErrorMessage"] = 1;
    }
    if (dv.errorStyle) {
      attrs["errorStyle"] = dv.errorStyle;
    }
    if (dv.inputTitle) {
      attrs["promptTitle"] = dv.inputTitle;
    }
    if (dv.inputMessage) {
      attrs["prompt"] = dv.inputMessage;
    }
    if (dv.errorTitle) {
      attrs["errorTitle"] = dv.errorTitle;
    }
    if (dv.errorMessage) {
      attrs["error"] = dv.errorMessage;
    }

    // Build formula children
    const children: string[] = [];

    if (dv.type === "list" && dv.values && dv.values.length > 0) {
      // List with explicit values: formula1 is quoted, comma-separated
      const quotedList = `"${dv.values.join(",")}"`;
      children.push(xmlElement("formula1", undefined, xmlEscape(quotedList)));
    } else if (dv.formula1 !== undefined) {
      children.push(xmlElement("formula1", undefined, xmlEscape(dv.formula1)));
    }

    if (dv.formula2 !== undefined) {
      children.push(xmlElement("formula2", undefined, xmlEscape(dv.formula2)));
    }

    if (children.length > 0) {
      dvElements.push(xmlElement("dataValidation", attrs, children));
    } else {
      dvElements.push(xmlSelfClose("dataValidation", attrs));
    }
  }

  return xmlElement("dataValidations", { count: validations.length }, dvElements);
}

// ── Hyperlink Collection ──────────────────────────────────────────

/**
 * Collect hyperlinks from the sheet's cell overrides and generate
 * the `<hyperlinks>` XML section plus external relationship entries.
 */
export function collectHyperlinks(sheet: WriteSheet): {
  xml: string;
  relationships: HyperlinkRelationship[];
} {
  if (!sheet.cells) {
    return { xml: "", relationships: [] };
  }

  const hyperlinkElements: string[] = [];
  const relationships: HyperlinkRelationship[] = [];
  let rIdCounter = 1;

  for (const [key, cellOverride] of sheet.cells) {
    if (!cellOverride.hyperlink) continue;

    const [rowStr, colStr] = key.split(",");
    const r = parseInt(rowStr, 10);
    const c = parseInt(colStr, 10);
    const ref = cellRef(r, c);
    const hl = cellOverride.hyperlink;

    const attrs: Record<string, string> = { ref };

    if (hl.location) {
      // Internal hyperlink — uses location attribute directly, no relationship needed
      attrs["location"] = hl.location;
    } else if (hl.target) {
      // External hyperlink — needs a relationship entry
      const rId = `rId${rIdCounter++}`;
      attrs["r:id"] = rId;
      relationships.push({ id: rId, target: hl.target });
    }

    if (hl.tooltip) {
      attrs["tooltip"] = hl.tooltip;
    }
    if (hl.display) {
      attrs["display"] = hl.display;
    }

    hyperlinkElements.push(xmlSelfClose("hyperlink", attrs));
  }

  if (hyperlinkElements.length === 0) {
    return { xml: "", relationships: [] };
  }

  return {
    xml: xmlElement("hyperlinks", undefined, hyperlinkElements),
    relationships,
  };
}

// ── Paper Size Map ──────────────────────────────────────────────────

const PAPER_SIZE_MAP: Record<PaperSize, number> = {
  letter: 1,
  legal: 5,
  a3: 8,
  a4: 9,
  a5: 11,
  b4: 12,
  b5: 13,
  executive: 7,
  tabloid: 3,
};

/** Reverse map: XLSX paper size number → PaperSize string */
export const PAPER_SIZE_REVERSE: Record<number, PaperSize> = {};
for (const [name, num] of Object.entries(PAPER_SIZE_MAP)) {
  PAPER_SIZE_REVERSE[num] = name as PaperSize;
}

// ── Page Margins Serialization ──────────────────────────────────────

/** Default Excel margins (in inches) */
const DEFAULT_MARGINS: Required<PageMargins> = {
  left: 0.7,
  right: 0.7,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3,
};

/** Serialize page margins. Always emits (Excel expects it). */
function serializePageMargins(margins?: PageMargins): string {
  const m = margins ?? {};
  return xmlSelfClose("pageMargins", {
    left: m.left ?? DEFAULT_MARGINS.left,
    right: m.right ?? DEFAULT_MARGINS.right,
    top: m.top ?? DEFAULT_MARGINS.top,
    bottom: m.bottom ?? DEFAULT_MARGINS.bottom,
    header: m.header ?? DEFAULT_MARGINS.header,
    footer: m.footer ?? DEFAULT_MARGINS.footer,
  });
}

// ── Page Setup Serialization ─────────────────────────────────────────

function serializePageSetup(ps: PageSetup): string {
  const attrs: Record<string, string | number> = {};

  if (ps.paperSize) {
    const num = PAPER_SIZE_MAP[ps.paperSize];
    if (num !== undefined) {
      attrs["paperSize"] = num;
    }
  }

  if (ps.orientation) {
    attrs["orientation"] = ps.orientation;
  }

  if (ps.scale !== undefined) {
    attrs["scale"] = ps.scale;
  }

  if (ps.fitToPage) {
    if (ps.fitToWidth !== undefined) {
      attrs["fitToWidth"] = ps.fitToWidth;
    }
    if (ps.fitToHeight !== undefined) {
      attrs["fitToHeight"] = ps.fitToHeight;
    }
  }

  if (ps.horizontalCentered) {
    attrs["horizontalCentered"] = 1;
  }

  if (ps.verticalCentered) {
    attrs["verticalCentered"] = 1;
  }

  // Only emit if there are attributes beyond default
  if (Object.keys(attrs).length === 0) {
    return "";
  }

  return xmlSelfClose("pageSetup", attrs);
}

// ── Header/Footer Serialization ──────────────────────────────────────

function serializeHeaderFooter(hf: HeaderFooter): string {
  const attrs: Record<string, string | number> = {};

  if (hf.differentOddEven) {
    attrs["differentOddEven"] = 1;
  }
  if (hf.differentFirst) {
    attrs["differentFirst"] = 1;
  }

  const children: string[] = [];

  if (hf.oddHeader) {
    children.push(xmlElement("oddHeader", undefined, xmlEscape(hf.oddHeader)));
  }
  if (hf.oddFooter) {
    children.push(xmlElement("oddFooter", undefined, xmlEscape(hf.oddFooter)));
  }
  if (hf.evenHeader) {
    children.push(xmlElement("evenHeader", undefined, xmlEscape(hf.evenHeader)));
  }
  if (hf.evenFooter) {
    children.push(xmlElement("evenFooter", undefined, xmlEscape(hf.evenFooter)));
  }
  if (hf.firstHeader) {
    children.push(xmlElement("firstHeader", undefined, xmlEscape(hf.firstHeader)));
  }
  if (hf.firstFooter) {
    children.push(xmlElement("firstFooter", undefined, xmlEscape(hf.firstFooter)));
  }

  if (children.length === 0) {
    return "";
  }

  return xmlElement("headerFooter", Object.keys(attrs).length > 0 ? attrs : undefined, children);
}

// ── Color Attribute Serialization ──────────────────────────────────────

/** Serialize a Color object into XML attributes for a color element */
function serializeColorAttrs(color: Color): Record<string, string | number> {
  const attrs: Record<string, string | number> = {};
  if (color.rgb !== undefined) {
    // XLSX expects ARGB format (8 chars), add "FF" alpha prefix if only 6 chars
    const rgb = color.rgb;
    attrs["rgb"] = rgb.length === 6 ? `FF${rgb}` : rgb;
  }
  if (color.theme !== undefined) {
    attrs["theme"] = color.theme;
  }
  if (color.tint !== undefined) {
    attrs["tint"] = color.tint;
  }
  if (color.indexed !== undefined) {
    attrs["indexed"] = color.indexed;
  }
  return attrs;
}

// ── Rich Text Serialization ──────────────────────────────────────────

/** Serialize an array of RichTextRun into XML elements for an <is> (inline string) block */
function serializeRichTextRuns(runs: RichTextRun[]): string[] {
  const elements: string[] = [];

  for (const run of runs) {
    const runChildren: string[] = [];

    // Run properties (<rPr>)
    if (run.font) {
      const rPrParts = serializeFontProps(run.font);
      if (rPrParts.length > 0) {
        runChildren.push(xmlElement("rPr", undefined, rPrParts));
      }
    }

    // Run text (<t>)
    // Use xml:space="preserve" to preserve whitespace
    const text = xmlEscape(run.text);
    const needsPreserve =
      run.text.length > 0 &&
      (run.text[0] === " " ||
        run.text[run.text.length - 1] === " " ||
        run.text.includes("\n") ||
        run.text.includes("\t"));
    if (needsPreserve) {
      runChildren.push(`<t xml:space="preserve">${text}</t>`);
    } else {
      runChildren.push(xmlElement("t", undefined, text));
    }

    elements.push(xmlElement("r", undefined, runChildren));
  }

  return elements;
}

/** Serialize FontStyle into individual XML elements for <rPr> */
function serializeFontProps(font: FontStyle): string[] {
  const parts: string[] = [];

  if (font.bold) {
    parts.push(xmlSelfClose("b"));
  }
  if (font.italic) {
    parts.push(xmlSelfClose("i"));
  }
  if (font.underline) {
    if (font.underline === true || font.underline === "single") {
      parts.push(xmlSelfClose("u"));
    } else {
      parts.push(xmlSelfClose("u", { val: font.underline }));
    }
  }
  if (font.strikethrough) {
    parts.push(xmlSelfClose("strike"));
  }
  if (font.vertAlign) {
    parts.push(xmlSelfClose("vertAlign", { val: font.vertAlign }));
  }
  if (font.size !== undefined) {
    parts.push(xmlSelfClose("sz", { val: font.size }));
  }
  if (font.color) {
    parts.push(xmlSelfClose("color", serializeColorAttrs(font.color)));
  }
  if (font.name) {
    parts.push(xmlSelfClose("rFont", { val: font.name }));
  }
  if (font.family !== undefined) {
    parts.push(xmlSelfClose("family", { val: font.family }));
  }
  if (font.charset !== undefined) {
    parts.push(xmlSelfClose("charset", { val: font.charset }));
  }
  if (font.scheme) {
    parts.push(xmlSelfClose("scheme", { val: font.scheme }));
  }

  return parts;
}

// ── Conditional Formatting Serialization ─────────────────────────

/**
 * Serialize conditional formatting rules into `<conditionalFormatting>` XML blocks.
 * Rules are grouped by range (sqref) — multiple rules on the same range go into one element.
 */
function serializeConditionalFormatting(
  rules: ConditionalRule[],
  styles: StylesCollector,
): string[] {
  // Group rules by range
  const byRange = new Map<string, ConditionalRule[]>();
  for (const rule of rules) {
    const existing = byRange.get(rule.range);
    if (existing) {
      existing.push(rule);
    } else {
      byRange.set(rule.range, [rule]);
    }
  }

  const elements: string[] = [];

  for (const [range, rangeRules] of byRange) {
    const cfRuleElements: string[] = [];

    for (const rule of rangeRules) {
      cfRuleElements.push(serializeCfRule(rule, styles));
    }

    elements.push(xmlElement("conditionalFormatting", { sqref: range }, cfRuleElements));
  }

  return elements;
}

/** Serialize a single `<cfRule>` element */
function serializeCfRule(rule: ConditionalRule, styles: StylesCollector): string {
  const attrs: Record<string, string | number | boolean> = {
    type: rule.type,
    priority: rule.priority,
  };

  // Register dxf style and set dxfId
  if (rule.style) {
    attrs["dxfId"] = styles.addDxf(rule.style);
  }

  if (rule.operator) {
    attrs["operator"] = rule.operator;
  }

  if (rule.stopIfTrue) {
    attrs["stopIfTrue"] = true;
  }

  // Text-based rule attributes
  if (rule.text !== undefined) {
    attrs["text"] = rule.text;
  }

  const children: string[] = [];

  // Formulas
  if (rule.formula !== undefined) {
    const formulas = Array.isArray(rule.formula) ? rule.formula : [rule.formula];
    for (const f of formulas) {
      children.push(xmlElement("formula", undefined, xmlEscape(f)));
    }
  }

  // Color scale
  if (rule.type === "colorScale" && rule.colorScale) {
    const csChildren: string[] = [];
    for (const cfvo of rule.colorScale.cfvo) {
      const cfvoAttrs: Record<string, string> = { type: cfvo.type };
      if (cfvo.value !== undefined) cfvoAttrs["val"] = cfvo.value;
      csChildren.push(xmlSelfClose("cfvo", cfvoAttrs));
    }
    for (const color of rule.colorScale.colors) {
      csChildren.push(xmlSelfClose("color", { rgb: color }));
    }
    children.push(xmlElement("colorScale", undefined, csChildren));
  }

  // Data bar
  if (rule.type === "dataBar" && rule.dataBar) {
    const dbChildren: string[] = [];
    for (const cfvo of rule.dataBar.cfvo) {
      const cfvoAttrs: Record<string, string> = { type: cfvo.type };
      if (cfvo.value !== undefined) cfvoAttrs["val"] = cfvo.value;
      dbChildren.push(xmlSelfClose("cfvo", cfvoAttrs));
    }
    dbChildren.push(xmlSelfClose("color", { rgb: rule.dataBar.color }));
    children.push(xmlElement("dataBar", undefined, dbChildren));
  }

  // Icon set
  if (rule.type === "iconSet" && rule.iconSet) {
    const isAttrs: Record<string, string | boolean> = {
      iconSet: rule.iconSet.iconSet,
    };
    if (rule.iconSet.reverse) isAttrs["reverse"] = true;
    if (rule.iconSet.showValue === false) isAttrs["showValue"] = false;

    const isChildren: string[] = [];
    for (const cfvo of rule.iconSet.cfvo) {
      const cfvoAttrs: Record<string, string> = { type: cfvo.type };
      if (cfvo.value !== undefined) cfvoAttrs["val"] = cfvo.value;
      isChildren.push(xmlSelfClose("cfvo", cfvoAttrs));
    }
    children.push(xmlElement("iconSet", isAttrs, isChildren));
  }

  if (children.length > 0) {
    return xmlElement("cfRule", attrs, children);
  }
  return xmlSelfClose("cfRule", attrs);
}

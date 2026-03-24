// ── Worksheet XML Writer ─────────────────────────────────────────────
// Generates xl/worksheets/sheetN.xml for an XLSX package.

import type { WriteSheet, CellValue, CellStyle, DataValidation } from "../_types";
import type { StylesCollector } from "./styles-writer";
import { dateToSerial } from "../_date";
import { xmlDocument, xmlElement, xmlSelfClose, xmlEscape } from "../xml/writer";
import { calculateColumnWidth } from "./auto-width";

// ── Hyperlink Relationship ────────────────────────────────────────

export interface HyperlinkRelationship {
  id: string;
  target: string;
}

export interface WorksheetResult {
  xml: string;
  hyperlinkRelationships: HyperlinkRelationship[];
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
    children.push(xmlElement("si", undefined, [xmlElement("t", undefined, xmlEscape(str))]));
  }

  return xmlDocument("sst", { xmlns: NS_SPREADSHEET, count, uniqueCount: count }, children);
}

// ── Resolved Cell Data ─────────────────────────────────────────────

interface ResolvedCell {
  value: CellValue;
  style?: CellStyle;
  formula?: string;
  formulaResult?: CellValue;
}

// ── Default date format ────────────────────────────────────────────

const DEFAULT_DATE_FORMAT = "yyyy-mm-dd";

// ── Worksheet Writer ───────────────────────────────────────────────

/** Generate xl/worksheets/sheetN.xml along with any hyperlink relationships */
export function writeWorksheetXml(
  sheet: WriteSheet,
  styles: StylesCollector,
  sharedStrings: SharedStringsCollector,
  dateSystem?: "1900" | "1904",
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
  }

  const viewAttrs: Record<string, string | number | boolean> = {
    workbookViewId: 0,
  };

  if (sheet.view) {
    if (sheet.view.showGridLines === false) viewAttrs["showGridLines"] = false;
    if (sheet.view.showRowColHeaders === false) viewAttrs["showRowColHeaders"] = false;
    if (sheet.view.zoomScale !== undefined) viewAttrs["zoomScale"] = sheet.view.zoomScale;
    if (sheet.view.rightToLeft) viewAttrs["rightToLeft"] = true;
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

      if (effectiveWidth !== undefined || col.hidden || col.outlineLevel) {
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
    if (!row || row.length === 0) continue;

    const cellElements: string[] = [];
    let hasAnyCells = false;

    for (let c = 0; c < row.length; c++) {
      const resolved = row[c];
      if (!resolved) continue;

      const cellXml = serializeCell(r, c, resolved, styles, sharedStrings, is1904);
      if (cellXml) {
        cellElements.push(cellXml);
        hasAnyCells = true;
      }
    }

    if (hasAnyCells) {
      rowElements.push(xmlElement("row", { r: r + 1 }, cellElements));
    }
  }

  parts.push(xmlElement("sheetData", undefined, rowElements.length > 0 ? rowElements : ""));

  // ── Merge Cells ──
  if (sheet.merges && sheet.merges.length > 0) {
    const mergeElements = sheet.merges.map((m) =>
      xmlSelfClose("mergeCell", {
        ref: rangeRef(m.startRow, m.startCol, m.endRow, m.endCol),
      }),
    );
    parts.push(xmlElement("mergeCells", { count: sheet.merges.length }, mergeElements));
  }

  // ── Auto Filter ──
  if (sheet.autoFilter) {
    parts.push(xmlSelfClose("autoFilter", { ref: sheet.autoFilter.range }));
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

  return {
    xml: xmlDocument("worksheet", { xmlns: NS_SPREADSHEET, "xmlns:r": NS_R }, parts),
    hyperlinkRelationships,
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
  const { value, style, formula, formulaResult } = resolved;

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

  // Formula cell
  if (formula) {
    const cellAttrs: Record<string, string | number> = { r: ref };
    if (styleIdx !== 0) cellAttrs["s"] = styleIdx;

    const children: string[] = [xmlElement("f", undefined, xmlEscape(formula))];

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

  // String value
  if (typeof value === "string") {
    const ssIdx = sharedStrings.add(value);
    const attrs: Record<string, string | number> = { r: ref, t: "s" };
    if (styleIdx !== 0) attrs["s"] = styleIdx;
    return xmlElement("c", attrs, [xmlElement("v", undefined, String(ssIdx))]);
  }

  // Number value
  if (typeof value === "number") {
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

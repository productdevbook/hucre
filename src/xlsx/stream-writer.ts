// ── Streaming XLSX Writer ────────────────────────────────────────────
// Incrementally builds an XLSX file row by row.
// Each addRow() serializes the row to XML immediately.
// finish() assembles all parts into a valid XLSX ZIP archive.

import type { CellValue, CellStyle, ColumnDef, FreezePane } from "../_types";
import { ZipWriter } from "../zip/writer";
import { writeContentTypes } from "./content-types-writer";
import { writeRootRels, writeWorkbookRels } from "./workbook-writer";
import { createStylesCollector } from "./styles-writer";
import { createSharedStrings, writeSharedStringsXml } from "./worksheet-writer";
import { cellRef } from "./worksheet-writer";
import { dateToSerial } from "../_date";
import { xmlDocument, xmlElement, xmlSelfClose } from "../xml/writer";

const encoder = /* @__PURE__ */ new TextEncoder();

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// ── Types ────────────────────────────────────────────────────────────

export interface StreamWriterOptions {
  /** Sheet name */
  name: string;
  /** Column definitions */
  columns?: ColumnDef[];
  /** Freeze pane */
  freezePane?: FreezePane;
  /** Date system. Default: "1900" */
  dateSystem?: "1900" | "1904";
}

// ── Default date format ─────────────────────────────────────────────

const DEFAULT_DATE_FORMAT = "yyyy-mm-dd";

// ── Stream Writer Class ─────────────────────────────────────────────

export class XlsxStreamWriter {
  private sheetName: string;
  private columns: ColumnDef[] | undefined;
  private freezePane: FreezePane | undefined;
  private dateSystem: "1900" | "1904";
  private styles = createStylesCollector();
  private sharedStrings = createSharedStrings();
  private rowXmlFragments: string[] = [];
  private rowCount = 0;
  private maxCols = 0;

  constructor(options: StreamWriterOptions) {
    this.sheetName = options.name;
    this.columns = options.columns;
    this.freezePane = options.freezePane;
    this.dateSystem = options.dateSystem ?? "1900";

    // If columns have headers, write the header row immediately
    if (this.columns && this.columns.some((col) => col.header)) {
      const headerValues: CellValue[] = this.columns.map((col) => col.header ?? col.key ?? null);
      this.addRow(headerValues);
    }
  }

  /** Add a row of values */
  addRow(values: CellValue[]): void {
    const rowIndex = this.rowCount;
    this.rowCount++;

    if (values.length > this.maxCols) {
      this.maxCols = values.length;
    }

    const is1904 = this.dateSystem === "1904";
    const cellElements: string[] = [];

    for (let c = 0; c < values.length; c++) {
      const value = values[c];
      const colDef = this.columns?.[c];
      let style: CellStyle | undefined = colDef?.style;

      // If numFmt on column but not in style, merge
      if (colDef?.numFmt && (!style || !style.numFmt)) {
        style = { ...style, numFmt: colDef.numFmt };
      }

      const cellXml = this.serializeCell(rowIndex, c, value, style, is1904);
      if (cellXml) {
        cellElements.push(cellXml);
      }
    }

    if (cellElements.length > 0) {
      this.rowXmlFragments.push(xmlElement("row", { r: rowIndex + 1 }, cellElements));
    }
  }

  /** Add a row from an object, using column definitions for value extraction.
   *  Requires columns with key accessors. */
  addObject(item: Record<string, unknown>): void {
    if (!this.columns) throw new Error("addObject requires columns with key accessors");
    const values: CellValue[] = this.columns.map((col) => {
      if (col.key !== undefined) return (item[col.key] ?? null) as CellValue;
      return null;
    });
    this.addRow(values);
  }

  /** Finalize and return the XLSX buffer */
  async finish(): Promise<Uint8Array> {
    const hasSharedStrings = this.sharedStrings.count() > 0;

    // Build worksheet XML
    const worksheetParts: string[] = [];

    // SheetViews (freeze panes)
    const sheetViewParts: string[] = [];
    if (this.freezePane) {
      const fp = this.freezePane;
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

    worksheetParts.push(
      xmlElement("sheetViews", undefined, [
        sheetViewParts.length > 0
          ? xmlElement("sheetView", { workbookViewId: 0 }, sheetViewParts)
          : xmlSelfClose("sheetView", { workbookViewId: 0 }),
      ]),
    );

    // SheetFormatPr
    worksheetParts.push(xmlSelfClose("sheetFormatPr", { defaultRowHeight: 15 }));

    // Columns
    if (this.columns && this.columns.length > 0) {
      const colElements: string[] = [];
      for (let i = 0; i < this.columns.length; i++) {
        const col = this.columns[i];
        if (col.width !== undefined || col.hidden || col.outlineLevel) {
          const colAttrs: Record<string, string | number | boolean> = {
            min: i + 1,
            max: i + 1,
          };
          if (col.width !== undefined) {
            colAttrs["width"] = col.width;
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
        worksheetParts.push(xmlElement("cols", undefined, colElements));
      }
    }

    // Sheet data — all accumulated row XML fragments
    worksheetParts.push(
      xmlElement(
        "sheetData",
        undefined,
        this.rowXmlFragments.length > 0 ? this.rowXmlFragments : "",
      ),
    );

    const worksheetXml = xmlDocument(
      "worksheet",
      { xmlns: NS_SPREADSHEET, "xmlns:r": NS_R },
      worksheetParts,
    );

    // Build workbook XML
    const sheetElements = [
      xmlSelfClose("sheet", {
        name: this.sheetName,
        sheetId: 1,
        "r:id": "rId1",
      }),
    ];
    const workbookParts: string[] = [];
    if (this.dateSystem === "1904") {
      workbookParts.push(xmlSelfClose("workbookPr", { date1904: 1 }));
    }
    workbookParts.push(xmlElement("sheets", undefined, sheetElements));
    const workbookXml = xmlDocument(
      "workbook",
      { xmlns: NS_SPREADSHEET, "xmlns:r": NS_R },
      workbookParts,
    );

    // Build ZIP archive
    const zip = new ZipWriter();

    // [Content_Types].xml
    zip.add(
      "[Content_Types].xml",
      encoder.encode(writeContentTypes({ sheetCount: 1, hasSharedStrings })),
    );

    // _rels/.rels
    zip.add("_rels/.rels", encoder.encode(writeRootRels()));

    // xl/workbook.xml
    zip.add("xl/workbook.xml", encoder.encode(workbookXml));

    // xl/_rels/workbook.xml.rels
    zip.add("xl/_rels/workbook.xml.rels", encoder.encode(writeWorkbookRels(1, hasSharedStrings)));

    // xl/styles.xml
    zip.add("xl/styles.xml", encoder.encode(this.styles.toXml()));

    // xl/sharedStrings.xml (if any strings)
    if (hasSharedStrings) {
      zip.add("xl/sharedStrings.xml", encoder.encode(writeSharedStringsXml(this.sharedStrings)));
    }

    // xl/worksheets/sheet1.xml
    zip.add("xl/worksheets/sheet1.xml", encoder.encode(worksheetXml));

    return zip.build();
  }

  // ── Private helpers ───────────────────────────────────────────────

  private serializeCell(
    row: number,
    col: number,
    value: CellValue,
    style: CellStyle | undefined,
    is1904: boolean,
  ): string | null {
    let effectiveStyle = style;

    // Add default date format for Date values without explicit format
    if (value instanceof Date && (!effectiveStyle || !effectiveStyle.numFmt)) {
      effectiveStyle = { ...effectiveStyle, numFmt: DEFAULT_DATE_FORMAT };
    }

    let styleIdx = 0;
    if (effectiveStyle) {
      styleIdx = this.styles.addStyle(effectiveStyle);
    }

    const ref = cellRef(row, col);

    // Null — skip if no style
    if (value === null || value === undefined) {
      if (styleIdx !== 0) {
        return xmlSelfClose("c", { r: ref, s: styleIdx });
      }
      return null;
    }

    // String
    if (typeof value === "string") {
      const ssIdx = this.sharedStrings.add(value);
      const attrs: Record<string, string | number> = { r: ref, t: "s" };
      if (styleIdx !== 0) attrs["s"] = styleIdx;
      return xmlElement("c", attrs, [xmlElement("v", undefined, String(ssIdx))]);
    }

    // Number
    if (typeof value === "number") {
      const attrs: Record<string, string | number> = { r: ref };
      if (styleIdx !== 0) attrs["s"] = styleIdx;
      return xmlElement("c", attrs, [xmlElement("v", undefined, String(value))]);
    }

    // Boolean
    if (typeof value === "boolean") {
      const attrs: Record<string, string | number> = { r: ref, t: "b" };
      if (styleIdx !== 0) attrs["s"] = styleIdx;
      return xmlElement("c", attrs, [xmlElement("v", undefined, value ? "1" : "0")]);
    }

    // Date — convert to serial number
    if (value instanceof Date) {
      const serial = dateToSerial(value, is1904);
      const attrs: Record<string, string | number> = { r: ref };
      if (styleIdx !== 0) attrs["s"] = styleIdx;
      return xmlElement("c", attrs, [xmlElement("v", undefined, String(serial))]);
    }

    return null;
  }
}

// ── Helpers ─────────────────────────────────────────────────────────

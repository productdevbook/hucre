// ── Streaming XLSX Writer ────────────────────────────────────────────
// Incrementally builds an XLSX file row by row.
// Each addRow() serializes the row to XML immediately.
// finish() assembles all parts into a valid XLSX ZIP archive.

import type { CellValue, CellStyle, ColumnDef, FreezePane } from "../_types"
import { ZipWriter } from "../zip/writer"
import { writeContentTypes } from "./content-types-writer"
import { writeRootRels, writeWorkbookRels } from "./workbook-writer"
import { createStylesCollector } from "./styles-writer"
import { createSharedStrings, writeSharedStringsXml } from "./worksheet-writer"
import { cellRef } from "./worksheet-writer"
import { dateToSerial } from "../_date"
import { xmlDocument, xmlElement, xmlSelfClose } from "../xml/writer"
import { writeThemeXml } from "./theme-writer"

const encoder = /* @__PURE__ */ new TextEncoder()

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

// ── Types ────────────────────────────────────────────────────────────

export interface StreamWriterOptions {
  /** Sheet name */
  name: string
  /** Column definitions */
  columns?: ColumnDef[]
  /** Freeze pane */
  freezePane?: FreezePane
  /** Date system. Default: "1900" */
  dateSystem?: "1900" | "1904"
  /**
   * When set, rows past this count are written into a new sheet named
   * `{name}_2`, `{name}_3`, ... (truncated to fit Excel's 31-char limit).
   *
   * Defaults to {@link XLSX_MAX_ROWS_PER_SHEET} (1,048,576) — Excel's hard
   * row limit. Pass an explicit number to roll over earlier (handy for
   * tests). Pass `Infinity` to disable the rollover.
   */
  maxRowsPerSheet?: number
  /**
   * When `true` (default) and the writer is set up with a column header
   * (either via `columns[].header` or the first call to `addRow`), the
   * header row is repeated as the first row of every rolled-over sheet.
   *
   * Set to `false` to leave new sheets without a header row.
   */
  repeatHeaders?: boolean
}

/** Excel's hard row limit since Excel 2007 (2^20). */
export const XLSX_MAX_ROWS_PER_SHEET = 1_048_576

// ── Default date format ─────────────────────────────────────────────

const DEFAULT_DATE_FORMAT = "yyyy-mm-dd"

// ── Stream Writer Class ─────────────────────────────────────────────

export class XlsxStreamWriter {
  private sheetName: string
  private columns: ColumnDef[] | undefined
  private freezePane: FreezePane | undefined
  private dateSystem: "1900" | "1904"
  private maxRowsPerSheet: number
  private repeatHeaders: boolean
  private styles = createStylesCollector()
  private sharedStrings = createSharedStrings()
  /**
   * One fragment array per sheet. New sheets are appended when the row
   * limit is reached.
   */
  private sheetFragments: string[][] = [[]]
  /** Row index within the *current* sheet, NOT the global count. */
  private currentSheetRowCount = 0
  /** Global row count across every sheet — preserves the original semantics. */
  private rowCount = 0
  private maxCols = 0
  /** Captured for `repeatHeaders`. Set when the first row is written. */
  private headerRowValues: CellValue[] | null = null
  /** Was the captured header injected by the constructor (vs. a user call)? */
  private headerWasFromColumns = false

  constructor(options: StreamWriterOptions) {
    this.sheetName = options.name
    this.columns = options.columns
    this.freezePane = options.freezePane
    this.dateSystem = options.dateSystem ?? "1900"
    this.maxRowsPerSheet = options.maxRowsPerSheet ?? XLSX_MAX_ROWS_PER_SHEET
    this.repeatHeaders = options.repeatHeaders ?? true

    if (this.maxRowsPerSheet < 2) {
      throw new Error("maxRowsPerSheet must be at least 2 (one header + one data row)")
    }

    // If columns have headers, write the header row immediately
    if (this.columns && this.columns.some((col) => col.header)) {
      const headerValues: CellValue[] = this.columns.map((col) => col.header ?? col.key ?? null)
      this.headerRowValues = headerValues.slice()
      this.headerWasFromColumns = true
      this.addRow(headerValues)
    }
  }

  /** Add a row of values */
  addRow(values: CellValue[]): void {
    // Capture the very first row as a fallback header for repeatHeaders, in
    // case the caller didn't supply column definitions but does want their
    // first row repeated when sheets roll over.
    if (this.rowCount === 0 && !this.headerRowValues) {
      this.headerRowValues = values.slice()
    }

    // Roll over before writing this row when the current sheet is full.
    if (this.currentSheetRowCount >= this.maxRowsPerSheet) {
      this.rolloverSheet()
    }

    const rowIndex = this.currentSheetRowCount
    this.currentSheetRowCount++
    this.rowCount++

    if (values.length > this.maxCols) {
      this.maxCols = values.length
    }

    const is1904 = this.dateSystem === "1904"
    const cellElements: string[] = []

    for (let c = 0; c < values.length; c++) {
      const value = values[c]
      const colDef = this.columns?.[c]
      let style: CellStyle | undefined = colDef?.style

      // If numFmt on column but not in style, merge
      if (colDef?.numFmt && (!style || !style.numFmt)) {
        style = { ...style, numFmt: colDef.numFmt }
      }

      const cellXml = this.serializeCell(rowIndex, c, value, style, is1904)
      if (cellXml) {
        cellElements.push(cellXml)
      }
    }

    if (cellElements.length > 0) {
      this.sheetFragments[this.sheetFragments.length - 1]!.push(
        xmlElement("row", { r: rowIndex + 1 }, cellElements),
      )
    }
  }

  /**
   * Open a new sheet for the next row. Optionally re-emits the captured
   * header row at the top of the new sheet.
   */
  private rolloverSheet(): void {
    this.sheetFragments.push([])
    this.currentSheetRowCount = 0

    if (this.repeatHeaders && this.headerRowValues) {
      // Re-emit the header row at row 0 of the new sheet. We bypass the
      // public `addRow` to avoid double-counting in `rowCount` and to dodge
      // the rollover guard at the top.
      const headerValues = this.headerRowValues
      const rowIndex = this.currentSheetRowCount
      this.currentSheetRowCount++
      const is1904 = this.dateSystem === "1904"
      const cellElements: string[] = []

      for (let c = 0; c < headerValues.length; c++) {
        const value = headerValues[c]
        const colDef = this.columns?.[c]
        let style: CellStyle | undefined = colDef?.style
        if (colDef?.numFmt && (!style || !style.numFmt)) {
          style = { ...style, numFmt: colDef.numFmt }
        }
        const cellXml = this.serializeCell(rowIndex, c, value, style, is1904)
        if (cellXml) cellElements.push(cellXml)
      }
      if (cellElements.length > 0) {
        this.sheetFragments[this.sheetFragments.length - 1]!.push(
          xmlElement("row", { r: rowIndex + 1 }, cellElements),
        )
      }
    }
  }

  /** Add a row from an object, using column definitions for value extraction.
   *  Requires columns with key accessors. */
  addObject(item: Record<string, unknown>): void {
    if (!this.columns) throw new Error("addObject requires columns with key accessors")
    const values: CellValue[] = this.columns.map((col) => {
      if (col.key !== undefined) return (item[col.key] ?? null) as CellValue
      return null
    })
    this.addRow(values)
  }

  /** Finalize and return the XLSX buffer */
  async finish(): Promise<Uint8Array> {
    const hasSharedStrings = this.sharedStrings.count() > 0
    const sheetCount = this.sheetFragments.length
    const sheetNames = this.generateSheetNames(sheetCount)

    // Build the same view/columns prelude for every emitted sheet.
    const sheetPrelude = this.buildSheetPrelude()

    // Build ZIP archive
    const zip = new ZipWriter()

    // [Content_Types].xml
    zip.add(
      "[Content_Types].xml",
      encoder.encode(writeContentTypes({ sheetCount, hasSharedStrings })),
    )

    // _rels/.rels
    zip.add("_rels/.rels", encoder.encode(writeRootRels()))

    // xl/_rels/workbook.xml.rels
    zip.add(
      "xl/_rels/workbook.xml.rels",
      encoder.encode(writeWorkbookRels(sheetCount, hasSharedStrings)),
    )

    // xl/styles.xml
    zip.add("xl/styles.xml", encoder.encode(this.styles.toXml()))

    // xl/theme/theme1.xml — declared in [Content_Types].xml and referenced
    // from workbook.xml.rels, so the part must actually be written or Excel
    // rejects the workbook as corrupt (matches the batch writer).
    zip.add("xl/theme/theme1.xml", encoder.encode(writeThemeXml()))

    // xl/sharedStrings.xml (if any strings)
    if (hasSharedStrings) {
      zip.add("xl/sharedStrings.xml", encoder.encode(writeSharedStringsXml(this.sharedStrings)))
    }

    // xl/worksheets/sheet{N}.xml — one entry per fragment array
    for (let s = 0; s < sheetCount; s++) {
      const fragments = this.sheetFragments[s]!
      const worksheetParts: string[] = []
      worksheetParts.push(...sheetPrelude)
      worksheetParts.push(xmlElement("sheetData", undefined, fragments.length > 0 ? fragments : ""))
      const worksheetXml = xmlDocument(
        "worksheet",
        { xmlns: NS_SPREADSHEET, "xmlns:r": NS_R },
        worksheetParts,
      )
      zip.add(`xl/worksheets/sheet${s + 1}.xml`, encoder.encode(worksheetXml))
    }

    // Build workbook XML
    const sheetElements: string[] = []
    for (let s = 0; s < sheetCount; s++) {
      sheetElements.push(
        xmlSelfClose("sheet", {
          name: sheetNames[s]!,
          sheetId: s + 1,
          "r:id": `rId${s + 1}`,
        }),
      )
    }
    const workbookParts: string[] = []
    if (this.dateSystem === "1904") {
      workbookParts.push(xmlSelfClose("workbookPr", { date1904: 1 }))
    }
    workbookParts.push(xmlElement("sheets", undefined, sheetElements))
    const workbookXml = xmlDocument(
      "workbook",
      { xmlns: NS_SPREADSHEET, "xmlns:r": NS_R },
      workbookParts,
    )

    // xl/workbook.xml
    zip.add("xl/workbook.xml", encoder.encode(workbookXml))

    return zip.build()
  }

  /** Build sheetView + sheetFormatPr + cols (same on every emitted sheet). */
  private buildSheetPrelude(): string[] {
    const parts: string[] = []

    // SheetViews (freeze panes)
    const sheetViewParts: string[] = []
    if (this.freezePane) {
      const fp = this.freezePane
      const topLeftCell = cellRef(fp.rows ?? 0, fp.columns ?? 0)
      const paneAttrs: Record<string, string | number> = {}
      if (fp.columns && fp.columns > 0) paneAttrs["xSplit"] = fp.columns
      if (fp.rows && fp.rows > 0) paneAttrs["ySplit"] = fp.rows
      paneAttrs["topLeftCell"] = topLeftCell
      paneAttrs["state"] = "frozen"
      const hasXSplit = fp.columns && fp.columns > 0
      const hasYSplit = fp.rows && fp.rows > 0
      if (hasXSplit && hasYSplit) paneAttrs["activePane"] = "bottomRight"
      else if (hasXSplit) paneAttrs["activePane"] = "topRight"
      else paneAttrs["activePane"] = "bottomLeft"
      sheetViewParts.push(xmlSelfClose("pane", paneAttrs))
    }
    parts.push(
      xmlElement("sheetViews", undefined, [
        sheetViewParts.length > 0
          ? xmlElement("sheetView", { workbookViewId: 0 }, sheetViewParts)
          : xmlSelfClose("sheetView", { workbookViewId: 0 }),
      ]),
    )

    // SheetFormatPr
    parts.push(xmlSelfClose("sheetFormatPr", { defaultRowHeight: 15 }))

    // Columns
    if (this.columns && this.columns.length > 0) {
      const colElements: string[] = []
      for (let i = 0; i < this.columns.length; i++) {
        const col = this.columns[i]
        if (col.width !== undefined || col.hidden || col.outlineLevel) {
          const colAttrs: Record<string, string | number | boolean> = {
            min: i + 1,
            max: i + 1,
          }
          if (col.width !== undefined) {
            colAttrs["width"] = col.width
            colAttrs["customWidth"] = true
          }
          if (col.hidden) colAttrs["hidden"] = true
          if (col.outlineLevel) colAttrs["outlineLevel"] = col.outlineLevel
          colElements.push(xmlSelfClose("col", colAttrs))
        }
      }
      if (colElements.length > 0) {
        parts.push(xmlElement("cols", undefined, colElements))
      }
    }

    return parts
  }

  /**
   * Generate `count` unique sheet names: the configured base name first,
   * then `{name}_2`, `{name}_3`, …. Each name is truncated to fit Excel's
   * 31-character limit by trimming the base, not the suffix.
   */
  private generateSheetNames(count: number): string[] {
    const names: string[] = []
    for (let i = 0; i < count; i++) {
      if (i === 0) {
        names.push(truncateSheetName(this.sheetName))
      } else {
        const suffix = `_${i + 1}`
        const room = 31 - suffix.length
        const base = this.sheetName.length > room ? this.sheetName.slice(0, room) : this.sheetName
        names.push(base + suffix)
      }
    }
    return names
  }

  // ── Private helpers ───────────────────────────────────────────────

  private serializeCell(
    row: number,
    col: number,
    value: CellValue,
    style: CellStyle | undefined,
    is1904: boolean,
  ): string | null {
    let effectiveStyle = style

    // Add default date format for Date values without explicit format
    if (value instanceof Date && (!effectiveStyle || !effectiveStyle.numFmt)) {
      effectiveStyle = { ...effectiveStyle, numFmt: DEFAULT_DATE_FORMAT }
    }

    let styleIdx = 0
    if (effectiveStyle) {
      styleIdx = this.styles.addStyle(effectiveStyle)
    }

    const ref = cellRef(row, col)

    // Null — skip if no style
    if (value === null || value === undefined) {
      if (styleIdx !== 0) {
        return xmlSelfClose("c", { r: ref, s: styleIdx })
      }
      return null
    }

    // String
    if (typeof value === "string") {
      const ssIdx = this.sharedStrings.add(value)
      const attrs: Record<string, string | number> = { r: ref, t: "s" }
      if (styleIdx !== 0) attrs["s"] = styleIdx
      return xmlElement("c", attrs, [xmlElement("v", undefined, String(ssIdx))])
    }

    // Number
    if (typeof value === "number") {
      // Infinity, -Infinity, and NaN cannot be represented in OOXML — emit as empty cell
      if (!Number.isFinite(value)) {
        if (styleIdx !== 0) {
          return xmlSelfClose("c", { r: ref, s: styleIdx })
        }
        return null
      }
      const attrs: Record<string, string | number> = { r: ref }
      if (styleIdx !== 0) attrs["s"] = styleIdx
      return xmlElement("c", attrs, [xmlElement("v", undefined, String(value))])
    }

    // Boolean
    if (typeof value === "boolean") {
      const attrs: Record<string, string | number> = { r: ref, t: "b" }
      if (styleIdx !== 0) attrs["s"] = styleIdx
      return xmlElement("c", attrs, [xmlElement("v", undefined, value ? "1" : "0")])
    }

    // Date — convert to serial number
    if (value instanceof Date) {
      const serial = dateToSerial(value, is1904)
      const attrs: Record<string, string | number> = { r: ref }
      if (styleIdx !== 0) attrs["s"] = styleIdx
      return xmlElement("c", attrs, [xmlElement("v", undefined, String(serial))])
    }

    return null
  }
}

// ── Helpers ─────────────────────────────────────────────────────────

/** Excel sheet names cap at 31 characters. */
function truncateSheetName(name: string): string {
  return name.length > 31 ? name.slice(0, 31) : name
}

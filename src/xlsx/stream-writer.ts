// ── Incremental & Streaming XLSX Writers ─────────────────────────────
//
// Two writers share the row-serialization core below:
//
// • `XlsxStreamWriter` — incremental. Each `addRow()` serializes to XML
//   right away, so no workbook object model is ever built, but the
//   serialized parts are held until `finish()` returns the whole file as
//   one buffer. Peak memory is O(data).
//
// • `writeXlsxStream()` — genuinely streaming. Rows are pulled from a
//   source on demand and the ZIP is emitted chunk by chunk, so peak
//   memory is O(distinct styles), independent of row count.

import type { CellValue, CellStyle, ColumnDef, FreezePane } from "../_types"
import { ZipWriter } from "../zip/writer"
import { zipStream, type ZipStreamEntry } from "../zip/stream-writer"
import { writeContentTypes } from "./content-types-writer"
import { writeRootRels, writeWorkbookRels } from "./workbook-writer"
import { createStylesCollector, type StylesCollector } from "./styles-writer"
import { createSharedStrings, writeSharedStringsXml } from "./worksheet-writer"
import { cellRef } from "./worksheet-writer"
import type { SharedStringsCollector } from "./worksheet-writer"
import { writeThemeXml } from "./theme-writer"
import { dateToSerial } from "../_date"
import { xmlDocument, xmlDeclaration, xmlElement, xmlSelfClose, xmlEscape } from "../xml/writer"

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

export interface XlsxWriteStreamOptions extends StreamWriterOptions {
  /**
   * Write strings as inline `<is><t>` cells instead of routing them
   * through `xl/sharedStrings.xml`. Default `true`.
   *
   * A shared string table has to be held in memory until the workbook is
   * finished, so it grows with the number of *distinct* strings and
   * undoes the constant-memory guarantee. Set to `false` when the data is
   * highly repetitive and a smaller file matters more than peak memory.
   */
  inlineStrings?: boolean
  /** DEFLATE the parts. Default `true`. */
  compress?: boolean
  /**
   * Emit ZIP64 records, lifting the 4 GiB ceiling on any single part and
   * on the archive. Default `false`.
   *
   * Turn this on when one worksheet's XML may exceed 4 GiB — very wide
   * sheets near the per-sheet row cap can. With it off, an overflow
   * throws rather than writing a corrupt file. ZIP64 archives need a
   * ZIP64-aware consumer; Excel 2007+ and hucre's own reader qualify.
   */
  zip64?: boolean
}

/** A streamed row: positional values, or an object read through `columns[].key`. */
export type XlsxStreamRow = CellValue[] | Record<string, unknown>

/** Excel's hard row limit since Excel 2007 (2^20). */
export const XLSX_MAX_ROWS_PER_SHEET = 1_048_576

// ── Default date format ─────────────────────────────────────────────

const DEFAULT_DATE_FORMAT = "yyyy-mm-dd"

/** Flush the XML accumulator once it crosses this many characters. */
const CHUNK_THRESHOLD = 64 * 1024

// ── Row Serialization Core ──────────────────────────────────────────

interface RowSerializerOptions {
  columns?: ColumnDef[]
  dateSystem?: "1900" | "1904"
  inlineStrings?: boolean
}

/**
 * Turns row values into `<row>` XML, collecting styles (and, unless
 * `inlineStrings`, shared strings) along the way. Both writers share
 * one of these so their output stays byte-identical.
 */
class RowSerializer {
  readonly styles: StylesCollector = createStylesCollector()
  readonly sharedStrings: SharedStringsCollector = createSharedStrings()
  private columns: ColumnDef[] | undefined
  private is1904: boolean
  private inlineStrings: boolean

  constructor(options: RowSerializerOptions) {
    this.columns = options.columns
    this.is1904 = options.dateSystem === "1904"
    this.inlineStrings = options.inlineStrings ?? false
  }

  /** Serialize one row. Returns `null` when every cell was empty. */
  serializeRow(rowIndex: number, values: CellValue[]): string | null {
    const cellElements: string[] = []

    for (let c = 0; c < values.length; c++) {
      const colDef = this.columns?.[c]
      let style: CellStyle | undefined = colDef?.style

      // If numFmt on column but not in style, merge
      if (colDef?.numFmt && (!style || !style.numFmt)) {
        style = { ...style, numFmt: colDef.numFmt }
      }

      const cellXml = this.serializeCell(rowIndex, c, values[c], style)
      if (cellXml) cellElements.push(cellXml)
    }

    if (cellElements.length === 0) return null
    return xmlElement("row", { r: rowIndex + 1 }, cellElements)
  }

  private serializeCell(
    row: number,
    col: number,
    value: CellValue,
    style: CellStyle | undefined,
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
      if (this.inlineStrings) {
        const attrs: Record<string, string | number> = { r: ref, t: "inlineStr" }
        if (styleIdx !== 0) attrs["s"] = styleIdx
        return xmlElement("c", attrs, [xmlElement("is", undefined, [textElement(value)])])
      }
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
      const serial = dateToSerial(value, this.is1904)
      const attrs: Record<string, string | number> = { r: ref }
      if (styleIdx !== 0) attrs["s"] = styleIdx
      return xmlElement("c", attrs, [xmlElement("v", undefined, String(serial))])
    }

    return null
  }
}

/** `<t>` element, preserving significant whitespace like the shared-string writer. */
function textElement(value: string): string {
  const escaped = xmlEscape(value)
  const needsPreserve =
    value.length > 0 &&
    (value[0] === " " ||
      value[value.length - 1] === " " ||
      value.includes("\n") ||
      value.includes("\t"))
  return needsPreserve
    ? `<t xml:space="preserve">${escaped}</t>`
    : xmlElement("t", undefined, escaped)
}

// ── Shared Sheet Scaffolding ────────────────────────────────────────

/** Build sheetView + sheetFormatPr + cols (same on every emitted sheet). */
function buildSheetPrelude(columns: ColumnDef[] | undefined, freezePane: FreezePane | undefined) {
  const parts: string[] = []

  // SheetViews (freeze panes)
  const sheetViewParts: string[] = []
  if (freezePane) {
    const fp = freezePane
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
  if (columns && columns.length > 0) {
    const colElements: string[] = []
    for (let i = 0; i < columns.length; i++) {
      const col = columns[i]
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
function generateSheetNames(baseName: string, count: number): string[] {
  const names: string[] = []
  for (let i = 0; i < count; i++) {
    if (i === 0) {
      names.push(truncateSheetName(baseName))
    } else {
      const suffix = `_${i + 1}`
      const room = 31 - suffix.length
      const base = baseName.length > room ? baseName.slice(0, room) : baseName
      names.push(base + suffix)
    }
  }
  return names
}

/** Build xl/workbook.xml for `count` sheets carrying the base name. */
function buildWorkbookXml(baseName: string, count: number, dateSystem: "1900" | "1904"): string {
  const sheetNames = generateSheetNames(baseName, count)
  const sheetElements: string[] = []
  for (let s = 0; s < count; s++) {
    sheetElements.push(
      xmlSelfClose("sheet", {
        name: sheetNames[s]!,
        sheetId: s + 1,
        "r:id": `rId${s + 1}`,
      }),
    )
  }

  const workbookParts: string[] = []
  if (dateSystem === "1904") {
    workbookParts.push(xmlSelfClose("workbookPr", { date1904: 1 }))
  }
  workbookParts.push(xmlElement("sheets", undefined, sheetElements))

  return xmlDocument("workbook", { xmlns: NS_SPREADSHEET, "xmlns:r": NS_R }, workbookParts)
}

/** Header values implied by `columns[].header`, or `null` when there are none. */
function headerFromColumns(columns: ColumnDef[] | undefined): CellValue[] | null {
  if (!columns || !columns.some((col) => col.header)) return null
  return columns.map((col) => col.header ?? col.key ?? null)
}

function validateMaxRowsPerSheet(value: number): void {
  if (value < 2) {
    throw new Error("maxRowsPerSheet must be at least 2 (one header + one data row)")
  }
}

// ── Stream Writer Class (incremental, buffered) ─────────────────────

/**
 * Incremental XLSX writer.
 *
 * Rows are serialized to XML as they arrive — no workbook object model
 * is built — but every serialized part is retained until {@link finish}
 * assembles the archive, so peak memory still scales with the data.
 * For constant-memory output use {@link writeXlsxStream} instead.
 */
export class XlsxStreamWriter {
  private sheetName: string
  private columns: ColumnDef[] | undefined
  private freezePane: FreezePane | undefined
  private dateSystem: "1900" | "1904"
  private maxRowsPerSheet: number
  private repeatHeaders: boolean
  private serializer: RowSerializer
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

  constructor(options: StreamWriterOptions) {
    this.sheetName = options.name
    this.columns = options.columns
    this.freezePane = options.freezePane
    this.dateSystem = options.dateSystem ?? "1900"
    this.maxRowsPerSheet = options.maxRowsPerSheet ?? XLSX_MAX_ROWS_PER_SHEET
    this.repeatHeaders = options.repeatHeaders ?? true
    this.serializer = new RowSerializer({
      columns: options.columns,
      dateSystem: this.dateSystem,
    })

    validateMaxRowsPerSheet(this.maxRowsPerSheet)

    // If columns have headers, write the header row immediately
    const headerValues = headerFromColumns(this.columns)
    if (headerValues) {
      this.headerRowValues = headerValues.slice()
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

    this.emit(rowIndex, values)
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
      const rowIndex = this.currentSheetRowCount
      this.currentSheetRowCount++
      this.emit(rowIndex, this.headerRowValues)
    }
  }

  private emit(rowIndex: number, values: CellValue[]): void {
    const xml = this.serializer.serializeRow(rowIndex, values)
    if (xml) {
      this.sheetFragments[this.sheetFragments.length - 1]!.push(xml)
    }
  }

  /** Add a row from an object, using column definitions for value extraction.
   *  Requires columns with key accessors. */
  addObject(item: Record<string, unknown>): void {
    if (!this.columns) throw new Error("addObject requires columns with key accessors")
    this.addRow(objectToValues(item, this.columns))
  }

  /** Finalize and return the XLSX buffer */
  async finish(): Promise<Uint8Array> {
    const hasSharedStrings = this.serializer.sharedStrings.count() > 0
    const sheetCount = this.sheetFragments.length

    // Build the same view/columns prelude for every emitted sheet.
    const sheetPrelude = buildSheetPrelude(this.columns, this.freezePane)

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
    zip.add("xl/styles.xml", encoder.encode(this.serializer.styles.toXml()))

    // xl/theme/theme1.xml — declared in [Content_Types].xml and referenced
    // from workbook.xml.rels, so the part must actually be written or Excel
    // rejects the workbook as corrupt (matches the batch writer).
    zip.add("xl/theme/theme1.xml", encoder.encode(writeThemeXml()))

    // xl/sharedStrings.xml (if any strings)
    if (hasSharedStrings) {
      zip.add(
        "xl/sharedStrings.xml",
        encoder.encode(writeSharedStringsXml(this.serializer.sharedStrings)),
      )
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

    // xl/workbook.xml
    zip.add(
      "xl/workbook.xml",
      encoder.encode(buildWorkbookXml(this.sheetName, sheetCount, this.dateSystem)),
    )

    return zip.build()
  }
}

// ── True Streaming Writer ───────────────────────────────────────────

/**
 * Write an XLSX workbook as a byte stream, pulling rows from `rows` only
 * as the consumer reads.
 *
 * Peak memory is independent of the row count: rows are serialized,
 * compressed, and flushed as they arrive, and nothing but the style
 * table (bounded by the number of *distinct* styles) plus one small ZIP
 * record per part is retained.
 *
 * ```ts
 * return new Response(writeXlsxStream({ name: "Export", columns }, rowSource), {
 *   headers: {
 *     "content-type":
 *       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
 *   },
 * })
 * ```
 *
 * Notes:
 * - Strings are written inline by default; see {@link XlsxWriteStreamOptions.inlineStrings}.
 * - Part sizes are unknown up front, so entries carry ZIP data
 *   descriptors. By default no ZIP64 records are emitted, which caps any
 *   single part — and the archive — at 4 GiB; see
 *   {@link XlsxWriteStreamOptions.zip64}.
 * - Compression needs `CompressionStream`; without it parts are stored
 *   uncompressed rather than buffered.
 */
export function writeXlsxStream(
  options: XlsxWriteStreamOptions,
  rows: AsyncIterable<XlsxStreamRow> | Iterable<XlsxStreamRow>,
): ReadableStream<Uint8Array> {
  return zipStream(xlsxStreamEntries(options, rows), { zip64: options.zip64 })
}

/** Lazily produce the ZIP entries for a streamed workbook. */
async function* xlsxStreamEntries(
  options: XlsxWriteStreamOptions,
  rows: AsyncIterable<XlsxStreamRow> | Iterable<XlsxStreamRow>,
): AsyncGenerator<ZipStreamEntry> {
  const dateSystem = options.dateSystem ?? "1900"
  const maxRowsPerSheet = options.maxRowsPerSheet ?? XLSX_MAX_ROWS_PER_SHEET
  const repeatHeaders = options.repeatHeaders ?? true
  const inlineStrings = options.inlineStrings ?? true
  const compress = options.compress ?? true
  const columns = options.columns

  validateMaxRowsPerSheet(maxRowsPerSheet)

  const serializer = new RowSerializer({ columns, dateSystem, inlineStrings })
  const prelude = buildSheetPrelude(columns, options.freezePane).join("")
  const cursor = createRowCursor(rows)

  let headerRow = headerFromColumns(columns)
  let globalRowCount = 0

  /** Stream one worksheet, stopping at the row cap or when rows run out. */
  async function* sheetChunks(sheetIndex: number): AsyncGenerator<Uint8Array> {
    const chunker = new XmlChunker()

    const open =
      xmlDeclaration() +
      `<worksheet xmlns="${NS_SPREADSHEET}" xmlns:r="${NS_R}">` +
      prelude +
      "<sheetData>"
    const openChunk = chunker.push(open)
    if (openChunk) yield openChunk

    let rowIndex = 0

    // Sheet 0 emits the column header as a real row (matching the
    // incremental writer); later sheets repeat it when asked to.
    if (headerRow && (sheetIndex === 0 || repeatHeaders)) {
      const xml = serializer.serializeRow(rowIndex, headerRow)
      rowIndex++
      if (sheetIndex === 0) globalRowCount++
      if (xml) {
        const chunk = chunker.push(xml)
        if (chunk) yield chunk
      }
    }

    while (rowIndex < maxRowsPerSheet) {
      const row = await cursor.next()
      if (row === undefined) break

      const values = Array.isArray(row) ? row : objectToValues(row, requireColumns(columns))

      // Without column headers the first row doubles as the repeated header.
      if (globalRowCount === 0 && !headerRow) {
        headerRow = values.slice()
      }

      const xml = serializer.serializeRow(rowIndex, values)
      rowIndex++
      globalRowCount++
      if (xml) {
        const chunk = chunker.push(xml)
        if (chunk) yield chunk
      }
    }

    const closeChunk = chunker.push("</sheetData></worksheet>")
    if (closeChunk) yield closeChunk
    const tail = chunker.flush()
    if (tail) yield tail
  }

  let sheetCount = 0
  try {
    for (;;) {
      const sheetIndex = sheetCount++
      yield {
        path: `xl/worksheets/sheet${sheetIndex + 1}.xml`,
        data: sheetChunks(sheetIndex),
        compress,
      }
      // The ZIP writer drains each entry before pulling the next, so by
      // now the sheet above is closed and the cursor is positioned on the
      // first row that didn't fit.
      if ((await cursor.peek()) === undefined) break
    }
  } finally {
    await cursor.close()
  }

  const hasSharedStrings = serializer.sharedStrings.count() > 0

  yield { path: "xl/styles.xml", data: encoder.encode(serializer.styles.toXml()), compress }

  if (hasSharedStrings) {
    yield {
      path: "xl/sharedStrings.xml",
      data: encoder.encode(writeSharedStringsXml(serializer.sharedStrings)),
      compress,
    }
  }

  yield { path: "xl/theme/theme1.xml", data: encoder.encode(writeThemeXml()), compress }

  yield {
    path: "xl/workbook.xml",
    data: encoder.encode(buildWorkbookXml(options.name, sheetCount, dateSystem)),
    compress,
  }

  yield {
    path: "xl/_rels/workbook.xml.rels",
    data: encoder.encode(writeWorkbookRels(sheetCount, hasSharedStrings)),
    compress,
  }

  yield { path: "_rels/.rels", data: encoder.encode(writeRootRels()), compress }

  // [Content_Types].xml lands last: the sheet count isn't known until the
  // rows run out. ZIP consumers resolve parts through the central
  // directory, so package order doesn't matter.
  yield {
    path: "[Content_Types].xml",
    data: encoder.encode(writeContentTypes({ sheetCount, hasSharedStrings })),
    compress,
  }
}

// ── Streaming helpers ───────────────────────────────────────────────

/** Accumulate XML text and hand it back as ~64 KB encoded chunks. */
class XmlChunker {
  private pending: string[] = []
  private size = 0

  /** Buffer `xml`, returning an encoded chunk once the threshold is crossed. */
  push(xml: string): Uint8Array | undefined {
    if (xml.length === 0) return undefined
    this.pending.push(xml)
    this.size += xml.length
    if (this.size < CHUNK_THRESHOLD) return undefined
    return this.flush()
  }

  /** Encode and release whatever is still buffered. */
  flush(): Uint8Array | undefined {
    if (this.size === 0) return undefined
    const text = this.pending.join("")
    this.pending = []
    this.size = 0
    return encoder.encode(text)
  }
}

interface RowCursor {
  /** Consume the next row, or `undefined` when the source is exhausted. */
  next(): Promise<XlsxStreamRow | undefined>
  /** Look at the next row without consuming it. */
  peek(): Promise<XlsxStreamRow | undefined>
  /** Release the underlying iterator. */
  close(): Promise<void>
}

function createRowCursor(
  source: AsyncIterable<XlsxStreamRow> | Iterable<XlsxStreamRow>,
): RowCursor {
  const iterator: AsyncIterator<XlsxStreamRow> | Iterator<XlsxStreamRow> =
    Symbol.asyncIterator in Object(source)
      ? (source as AsyncIterable<XlsxStreamRow>)[Symbol.asyncIterator]()
      : (source as Iterable<XlsxStreamRow>)[Symbol.iterator]()

  let lookahead: XlsxStreamRow | undefined
  let hasLookahead = false
  let done = false

  async function pull(): Promise<XlsxStreamRow | undefined> {
    if (done) return undefined
    const result = await iterator.next()
    if (result.done) {
      done = true
      return undefined
    }
    return result.value
  }

  return {
    async peek() {
      if (!hasLookahead) {
        lookahead = await pull()
        hasLookahead = true
      }
      return lookahead
    },
    async next() {
      if (hasLookahead) {
        hasLookahead = false
        const value = lookahead
        lookahead = undefined
        return value
      }
      return pull()
    },
    async close() {
      if (done) return
      done = true
      await iterator.return?.()
    },
  }
}

// ── Helpers ─────────────────────────────────────────────────────────

function requireColumns(columns: ColumnDef[] | undefined): ColumnDef[] {
  if (!columns) throw new Error("Object rows require columns with key accessors")
  return columns
}

function objectToValues(item: Record<string, unknown>, columns: ColumnDef[]): CellValue[] {
  return columns.map((col) => {
    if (col.key !== undefined) return (item[col.key] ?? null) as CellValue
    return null
  })
}

/** Excel sheet names cap at 31 characters. */
function truncateSheetName(name: string): string {
  return name.length > 31 ? name.slice(0, 31) : name
}

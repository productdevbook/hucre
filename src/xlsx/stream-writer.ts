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
//   memory retains distinct styles, optional shared strings and the current
//   physical sheet's links/comments, plus a header and ZIP part records.

import { resolveCellInput } from "../_inline-cells"
import { columnCellStyle, columnHeaders, objectRow } from "../_sheet-input"
import { createColumnWidthCollector } from "./auto-width"
import { SheetCellParts } from "./cell-parts"
import { METADATA_PART_PATH, writeMetadataXml } from "./metadata"
import { FPB_PART_PATH, writeFeaturePropertyBagXml } from "./feature-property-bag"
import type {
  AutoFilter,
  CellStyle,
  ColumnDef,
  ConditionalRule,
  FreezePane,
  MergeRange,
  RowDef,
  SpreadsheetStreamWriter,
  CellInput,
} from "../_types"
import { toRanges } from "../cell-utils"
import { ZipWriter } from "../zip/writer"
import { zipStream, type ZipStreamEntry } from "../zip/stream-writer"
import { writeContentTypes } from "./content-types-writer"
import { writeRootRels, writeWorkbookRels } from "./workbook-writer"
import { createStylesCollector, type StylesCollector } from "./styles-writer"
import { createSharedStrings, writeSharedStringsXml } from "./worksheet-writer"
import {
  cellRef,
  hasRowAttributes,
  rowAttributes,
  serializeAutoFilter,
  serializeCell,
  serializeColumns,
  serializeFrozenPane,
  serializeConditionalFormatting,
} from "./worksheet-writer"
import type { SharedStringsCollector } from "./worksheet-writer"
import { writeThemeXml } from "./theme-writer"
import {
  MAX_SHEET_NAME_LENGTH,
  validateSheetName,
  validateSheetNames,
  validateRowSize,
} from "../_validate"
import { MAX_ROW_INDEX } from "../limits"
import { InvalidArgumentError } from "../errors"
import { xmlDocument, xmlDeclaration, xmlElement, xmlSelfClose } from "../xml/writer"

const encoder = /* @__PURE__ */ new TextEncoder()

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

// ── Types ────────────────────────────────────────────────────────────

export interface XlsxStreamWriterOptions {
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
  /**
   * Row-level properties — `height`, `hidden`, `outlineLevel`, `collapsed` —
   * keyed by 0-based row index, counted the way the caller sees it: the key
   * keeps rising past a rollover, where the per-sheet index restarts at 0.
   * Consulted as each row is emitted, so it costs nothing per row that has
   * no entry.
   */
  rowDefs?: Map<number, RowDef>
  /**
   * Merged ranges. Written after the sheet data of the first sheet, so they
   * do not have to be known before the rows are streamed.
   */
  merges?: Array<MergeRange | string>
  /**
   * Auto-filter range. Written after the sheet data of the first sheet
   * (and before merges), matching the buffered writer.
   */
  autoFilter?: AutoFilter
  /**
   * Conditional formatting rules. Written after the sheet data of the first
   * sheet, matching the buffered writer.
   */
  conditionalRules?: ConditionalRule[]
}

/** Auto-width needs rows before columns; true streaming accepts fixed layout. */
export type XlsxStreamColumn = Omit<ColumnDef, "autoWidth">

export interface XlsxWriteStreamOptions extends Omit<XlsxStreamWriterOptions, "columns"> {
  columns?: XlsxStreamColumn[]
  /**
   * Write strings as inline `<is><t>` cells instead of routing them
   * through `xl/sharedStrings.xml`. Default `true`.
   *
   * A shared string table has to be held in memory until the workbook is
   * finished, so it grows with the number of *distinct* strings and
   * undoes the constant-memory guarantee. Set to `false` when the data is
   * highly repetitive and a smaller file matters more than peak memory.
   */
  stringMode?: "shared" | "inline"
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
export type XlsxStreamRow = CellInput[] | Record<string, unknown>

function serializeMergeCells(input: Array<MergeRange | string>): string {
  // A merge may be given as an A1 string; the file only knows coordinates.
  const merges = toRanges(input)!
  return xmlElement(
    "mergeCells",
    { count: merges.length },
    merges.map((merge) =>
      xmlSelfClose("mergeCell", {
        ref: `${cellRef(merge.startRow, merge.startCol)}:${cellRef(merge.endRow, merge.endCol)}`,
      }),
    ),
  )
}

/**
 * Copy a captured header row. `slice()` alone keeps a reference to each styled
 * cell, so a caller reusing one wrapper for later rows would see the header
 * repeated on rolled-over sheets carrying the mutated value.
 */
function snapshotRow(values: CellInput[]): CellInput[] {
  return structuredClone(values)
}

/**
 * One sheet of a streamed workbook, with its own name and layout.
 *
 * `rows` is pulled only as the consumer reads, exactly as in
 * {@link writeXlsxStream} — it is the sheet *list* that is eager, not the
 * data. See {@link writeXlsxStreamSheets}.
 */
export interface XlsxStreamSheet {
  /**
   * Sheet name. Must be unique across the workbook, compared
   * case-insensitively, because Excel compares them that way.
   */
  name: string
  /** Rows for this sheet, pulled on demand. */
  rows: AsyncIterable<XlsxStreamRow> | Iterable<XlsxStreamRow>
  /** Column definitions for this sheet. */
  columns?: XlsxStreamColumn[]
  /** Freeze pane for this sheet. */
  freezePane?: FreezePane
  /** Row-level properties for this sheet; see {@link XlsxStreamWriterOptions.rowDefs}. */
  rowDefs?: Map<number, RowDef>
  /** Merged ranges for this sheet; see {@link XlsxStreamWriterOptions.merges}. */
  merges?: Array<MergeRange | string>
  /** Auto-filter for this sheet; see {@link StreamWriterOptions.autoFilter}. */
  autoFilter?: AutoFilter
  /** Conditional formatting for this sheet; see {@link StreamWriterOptions.conditionalRules}. */
  conditionalRules?: ConditionalRule[]
  /** Overrides the workbook-level rollover cap for this sheet alone. */
  maxRowsPerSheet?: number
  /** Overrides the workbook-level header repetition for this sheet alone. */
  repeatHeaders?: boolean
}

/**
 * Workbook-wide options for {@link writeXlsxStreamSheets}.
 *
 * The per-sheet half of {@link XlsxWriteStreamOptions} — `name`,
 * `columns`, `freezePane`, `rowDefs`, `merges`, `autoFilter`,
 * `conditionalRules` — moves to
 * {@link XlsxStreamSheet}, since a multi-sheet workbook has one of each
 * *per sheet*. What is left is what
 * a workbook has exactly one of: the date system, the string strategy,
 * and the ZIP settings. `maxRowsPerSheet` and `repeatHeaders` stay here
 * as defaults every sheet inherits unless it sets its own.
 */
export interface XlsxWriteStreamWorkbookOptions {
  /** Date system. Default: "1900" */
  dateSystem?: "1900" | "1904"
  /** Default rollover cap; see {@link XlsxWriteStreamOptions.maxRowsPerSheet}. */
  maxRowsPerSheet?: number
  /** Default header repetition; see {@link XlsxWriteStreamOptions.repeatHeaders}. */
  repeatHeaders?: boolean
  /** See {@link XlsxWriteStreamOptions.stringMode}. */
  stringMode?: "shared" | "inline"
  /** DEFLATE the parts. Default `true`. */
  compress?: boolean
  /** See {@link XlsxWriteStreamOptions.zip64}. */
  zip64?: boolean
}

/** Excel's hard row limit since Excel 2007 (2^20). */
export const XLSX_MAX_ROWS_PER_SHEET: number = MAX_ROW_INDEX + 1

/** Flush the XML accumulator once it crosses this many characters. */
const CHUNK_THRESHOLD = 64 * 1024

// ── Row Serialization Core ──────────────────────────────────────────

interface RowSerializerOptions {
  columns?: ColumnDef[]
  dateSystem?: "1900" | "1904"
  inlineStrings?: boolean
  /**
   * Collectors to write into. A multi-sheet workbook hands the same two
   * to every sheet's serializer, because `xl/styles.xml` and
   * `xl/sharedStrings.xml` are workbook-wide parts: a style registered
   * while serializing sheet 2 has to end up in the same table sheet 1
   * indexes into. Omitted, each serializer gets its own pair.
   */
  styles?: StylesCollector
  sharedStrings?: SharedStringsCollector
}

/**
 * Turns row values into `<row>` XML, collecting styles (and, unless
 * `inlineStrings`, shared strings) along the way. Both writers share
 * one of these so their output stays byte-identical.
 */
class RowSerializer {
  readonly styles: StylesCollector
  readonly sharedStrings: SharedStringsCollector
  private columnStyles: Array<CellStyle | undefined>
  readonly columnWidths: ReturnType<typeof createColumnWidthCollector>
  hasDynamicArray = false
  private is1904: boolean
  private inlineStrings: boolean

  constructor(options: RowSerializerOptions) {
    this.styles = options.styles ?? createStylesCollector()
    this.sharedStrings = options.sharedStrings ?? createSharedStrings()
    this.columnStyles = options.columns?.map(columnCellStyle) ?? []
    this.columnWidths = createColumnWidthCollector(options.columns)
    this.is1904 = options.dateSystem === "1904"
    this.inlineStrings = options.inlineStrings ?? false
  }

  /** Serialize one row. Returns `null` when every cell was empty. */
  serializeRow(
    rowIndex: number,
    values: CellInput[],
    parts: SheetCellParts,
    rowDef?: RowDef,
  ): string | null {
    validateRowSize(rowIndex, values.length)
    const cellElements: string[] = []

    for (let c = 0; c < values.length; c++) {
      const cell = resolveCellInput(values[c])
      if (this.columnWidths.widths.size) this.columnWidths.add(c, cell)
      cell.style ??= this.columnStyles[c]
      if (cell.formulaDynamic) this.hasDynamicArray = true
      parts.add(rowIndex, c, cell)
      const cellXml = serializeCell(
        rowIndex,
        c,
        cell,
        this.styles,
        this.sharedStrings,
        this.is1904,
        this.inlineStrings,
      )
      if (cellXml) cellElements.push(cellXml)
    }

    const attrs = rowAttributes(rowIndex, rowDef)

    if (cellElements.length === 0) {
      // A row with no cells still has to be emitted when its definition asks
      // for something — a height, or a hidden/grouped row the caller wants.
      return hasRowAttributes(rowDef) ? xmlSelfClose("row", attrs) : null
    }
    return xmlElement("row", attrs, cellElements)
  }
}

// ── Shared Sheet Scaffolding ────────────────────────────────────────

/** Build sheetView + sheetFormatPr + cols (same on every emitted sheet). */
function buildSheetPrelude(
  columns: ColumnDef[] | undefined,
  freezePane: FreezePane | undefined,
  styles: StylesCollector,
  widths?: ReadonlyMap<number, number>,
): string[] {
  const parts: string[] = []

  // SheetViews (freeze panes)
  const sheetViewParts: string[] = []
  if (freezePane) {
    sheetViewParts.push(serializeFrozenPane(freezePane))
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

  const columnsXml = serializeColumns(columns, styles, widths)
  if (columnsXml) parts.push(columnsXml)

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
    names.push(i === 0 ? truncateSheetName(baseName) : suffixSheetName(baseName, i + 1))
  }
  return names
}

/** `{baseName}_{ordinal}`, trimmed so the whole name still fits Excel's cap. */
function suffixSheetName(baseName: string, ordinal: number): string {
  const suffix = `_${ordinal}`
  const room = MAX_SHEET_NAME_LENGTH - suffix.length
  const base = baseName.length > room ? baseName.slice(0, room) : baseName
  return base + suffix
}

/**
 * Name the `part`-th worksheet a single sheet rolled over into, skipping
 * any name already taken.
 *
 * One sheet's rollover can land on a name another sheet declared: a
 * workbook holding `Data` and `Data_2` rolls `Data` straight into the
 * name its neighbour owns, and Excel refuses to open a file with two
 * sheets of the same name. Bumping the ordinal until the name is free
 * keeps the rollover invisible to callers who never hit it, and keeps
 * the file valid for the ones who do.
 */
function claimSheetName(baseName: string, part: number, taken: Set<string>): string {
  let ordinal = part + 1
  let candidate = part === 0 ? truncateSheetName(baseName) : suffixSheetName(baseName, ordinal)

  while (taken.has(candidate.toLowerCase())) {
    ordinal++
    candidate = suffixSheetName(baseName, ordinal)
  }

  taken.add(candidate.toLowerCase())
  return candidate
}

/** Build xl/workbook.xml for `count` sheets carrying the base name. */
function buildWorkbookXml(baseName: string, count: number, dateSystem: "1900" | "1904"): string {
  return buildWorkbookXmlFor(generateSheetNames(baseName, count), dateSystem)
}

/** Build xl/workbook.xml for sheets whose names are already resolved. */
function buildWorkbookXmlFor(sheetNames: string[], dateSystem: "1900" | "1904"): string {
  const count = sheetNames.length
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

function validateMaxRowsPerSheet(value: number): void {
  if (value < 2) {
    throw new InvalidArgumentError("maxRowsPerSheet must be at least 2 (one header + one data row)")
  }
  // NaN previously made the lazy sheet loop emit empty sheets forever.
  if (value !== Infinity && (!Number.isInteger(value) || value > XLSX_MAX_ROWS_PER_SHEET)) {
    throw new InvalidArgumentError(
      `maxRowsPerSheet must be an integer in 2..${XLSX_MAX_ROWS_PER_SHEET}, or Infinity`,
    )
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
export class XlsxStreamWriter implements SpreadsheetStreamWriter {
  private sheetName: string
  private columns: ColumnDef[] | undefined
  private freezePane: FreezePane | undefined
  private dateSystem: "1900" | "1904"
  private rowDefs: Map<number, RowDef> | undefined
  private merges: Array<MergeRange | string> | undefined
  private autoFilter: AutoFilter | undefined
  private conditionalRules: ConditionalRule[] | undefined
  private maxRowsPerSheet: number
  private repeatHeaders: boolean
  private serializer: RowSerializer
  /**
   * One fragment array per sheet. New sheets are appended when the row
   * limit is reached.
   */
  private sheetFragments: string[][] = [[]]
  private sheetParts = [new SheetCellParts()]
  private done = false
  /** Row index within the *current* sheet, NOT the global count. */
  private currentSheetRowCount = 0
  /** Global row count across every sheet — preserves the original semantics. */
  private rowCount = 0
  private maxCols = 0
  /** Captured for `repeatHeaders`. Set when the first row is written. */
  private headerRowValues: CellInput[] | null = null

  constructor(options: XlsxStreamWriterOptions) {
    this.sheetName = options.name
    validateRowSize(0, options.columns?.length ?? 0)
    this.columns = options.columns
    this.freezePane = options.freezePane
    this.dateSystem = options.dateSystem ?? "1900"
    this.maxRowsPerSheet = options.maxRowsPerSheet ?? XLSX_MAX_ROWS_PER_SHEET
    this.repeatHeaders = options.repeatHeaders ?? true
    this.rowDefs = options.rowDefs
    this.merges = options.merges
    this.autoFilter = options.autoFilter
    this.conditionalRules = options.conditionalRules
    this.serializer = new RowSerializer({
      columns: options.columns,
      dateSystem: this.dateSystem,
    })

    // The base name is the caller's; rollover names are generated, and
    // those are still truncated below because the `_2` suffix can push a
    // legal 31-character base over the limit. See #364.
    validateSheetName(this.sheetName, 0)
    validateMaxRowsPerSheet(this.maxRowsPerSheet)

    // If columns have headers, write the header row immediately
    const headerValues = columnHeaders(this.columns)
    if (headerValues) {
      this.headerRowValues = headerValues.slice()
      this.addRow(headerValues)
    }
  }

  /** Add a row of values, each optionally carrying its own style or formula. */
  addRow(values: CellInput[]): void {
    if (this.done) throw new InvalidArgumentError("Cannot write to XlsxStreamWriter after finish()")
    // Reject before capturing a header, rolling over or advancing counters.
    const nextRow =
      this.currentSheetRowCount >= this.maxRowsPerSheet
        ? Number(this.repeatHeaders && this.headerRowValues !== null)
        : this.currentSheetRowCount
    validateRowSize(nextRow, values.length)
    // Capture the very first row as a fallback header for repeatHeaders, in
    // case the caller didn't supply column definitions but does want their
    // first row repeated when sheets roll over.
    if (this.rowCount === 0 && !this.headerRowValues) {
      this.headerRowValues = snapshotRow(values)
    }

    // Roll over before writing this row when the current sheet is full.
    if (this.currentSheetRowCount >= this.maxRowsPerSheet) {
      this.rolloverSheet()
    }

    const rowIndex = this.currentSheetRowCount
    const globalRow = this.rowCount
    this.currentSheetRowCount++
    this.rowCount++

    if (values.length > this.maxCols) {
      this.maxCols = values.length
    }

    // rowDefs are keyed by the caller's row number, which keeps counting past
    // a rollover — `rowIndex` restarts at 0 on every generated sheet.
    this.emit(rowIndex, values, globalRow)
  }

  /**
   * Open a new sheet for the next row. Optionally re-emits the captured
   * header row at the top of the new sheet.
   */
  private rolloverSheet(): void {
    this.sheetFragments.push([])
    this.sheetParts.push(new SheetCellParts())
    this.currentSheetRowCount = 0

    if (this.repeatHeaders && this.headerRowValues) {
      // Re-emit the header row at row 0 of the new sheet. We bypass the
      // public `addRow` to avoid double-counting in `rowCount` and to dodge
      // the rollover guard at the top.
      const rowIndex = this.currentSheetRowCount
      this.currentSheetRowCount++
      // The repeated header is the same logical row as the original, so it
      // carries the same row definition.
      this.emit(rowIndex, this.headerRowValues, 0)
    }
  }

  private emit(rowIndex: number, values: CellInput[], globalRow: number): void {
    const xml = this.serializer.serializeRow(
      rowIndex,
      values,
      this.sheetParts[this.sheetParts.length - 1],
      this.rowDefs?.get(globalRow),
    )
    if (xml) {
      this.sheetFragments[this.sheetFragments.length - 1]!.push(xml)
    }
  }

  /** Add a row from an object, using column definitions for value extraction.
   *  Requires columns with key accessors. */
  addObject(item: Record<string, unknown>): void {
    if (!this.columns)
      throw new InvalidArgumentError("addObject requires columns with key accessors")
    this.addRow(objectRow(item, this.columns))
  }

  /** Finalize and return the XLSX buffer */
  async finish(): Promise<Uint8Array> {
    this.done = true
    const hasSharedStrings = this.serializer.sharedStrings.count() > 0
    const sheetCount = this.sheetFragments.length
    // Differential formats must be registered before styles.xml is built.
    const firstTail = [
      this.autoFilter ? serializeAutoFilter(this.autoFilter) : "",
      this.merges?.length ? serializeMergeCells(this.merges) : "",
      this.conditionalRules?.length
        ? serializeConditionalFormatting(this.conditionalRules, this.serializer.styles).join("")
        : "",
    ].join("")
    const hasFeaturePropertyBag = this.serializer.styles.hasCheckboxFeature()
    const hasMetadata = this.serializer.hasDynamicArray
    const commentIndices = this.sheetParts.flatMap((parts, index) =>
      parts.hasComments ? [index + 1] : [],
    )

    // Build the same view/columns prelude for every emitted sheet.
    const sheetPrelude = buildSheetPrelude(
      this.columns,
      this.freezePane,
      this.serializer.styles,
      this.serializer.columnWidths.widths,
    )

    // Build ZIP archive
    const zip = new ZipWriter()

    // [Content_Types].xml
    zip.add(
      "[Content_Types].xml",
      encoder.encode(
        writeContentTypes({
          sheetCount,
          hasSharedStrings,
          commentIndices,
          hasFeaturePropertyBag,
          hasMetadata,
        }),
      ),
    )

    // _rels/.rels
    zip.add("_rels/.rels", encoder.encode(writeRootRels()))

    // xl/_rels/workbook.xml.rels
    zip.add(
      "xl/_rels/workbook.xml.rels",
      encoder.encode(
        writeWorkbookRels(
          sheetCount,
          hasSharedStrings,
          undefined,
          hasFeaturePropertyBag,
          undefined,
          undefined,
          undefined,
          undefined,
          undefined,
          undefined,
          hasMetadata,
        ),
      ),
    )

    // xl/styles.xml
    zip.add("xl/styles.xml", encoder.encode(this.serializer.styles.toXml()))

    // xl/theme/theme1.xml — declared in [Content_Types].xml and referenced
    // from workbook.xml.rels, so the part must actually be written or Excel
    // rejects the workbook as corrupt (matches the batch writer).
    zip.add("xl/theme/theme1.xml", encoder.encode(writeThemeXml()))
    if (hasMetadata) zip.add(METADATA_PART_PATH, encoder.encode(writeMetadataXml()))
    if (hasFeaturePropertyBag) zip.add(FPB_PART_PATH, encoder.encode(writeFeaturePropertyBagXml()))

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
      // Merges belong to the first sheet: a rollover splits one logical sheet
      // into several, and a range copied onto the continuation would cover
      // rows it was never meant to.
      if (s === 0) worksheetParts.push(firstTail)
      worksheetParts.push(this.sheetParts[s].toXml())
      const worksheetXml = xmlDocument(
        "worksheet",
        { xmlns: NS_SPREADSHEET, "xmlns:r": NS_R },
        worksheetParts,
      )
      zip.add(`xl/worksheets/sheet${s + 1}.xml`, encoder.encode(worksheetXml))
      for (const entry of this.sheetParts[s].entries(s + 1))
        zip.add(entry.path, encoder.encode(entry.xml))
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
 * Rows are serialized, compressed and flushed as they arrive. Retained
 * state is distinct styles, optional shared strings, the captured header,
 * current physical sheet's links/comments and one small ZIP record per
 * part. Link/comment metadata is released after that sheet's parts are
 * emitted; a sheet with a comment on every row still grows with row count.
 *
 * ```ts
 * return new Response(writeXlsxStream(rowSource, { name: "Export", columns }), {
 *   headers: {
 *     "content-type":
 *       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
 *   },
 * })
 * ```
 *
 * Notes:
 * - Strings are written inline by default; see {@link XlsxWriteStreamOptions.stringMode}.
 * - Part sizes are unknown up front, so entries carry ZIP data
 *   descriptors. By default no ZIP64 records are emitted, which caps any
 *   single part — and the archive — at 4 GiB; see
 *   {@link XlsxWriteStreamOptions.zip64}.
 * - Compression needs `CompressionStream`; without it parts are stored
 *   uncompressed rather than buffered.
 */
export function writeXlsxStream(
  rows: AsyncIterable<XlsxStreamRow> | Iterable<XlsxStreamRow>,
  options: XlsxWriteStreamOptions,
): ReadableStream<Uint8Array> {
  const { name, columns, freezePane, rowDefs, merges, autoFilter, conditionalRules, ...workbook } =
    options

  return writeXlsxStreamSheets(
    [{ name, rows, columns, freezePane, rowDefs, merges, autoFilter, conditionalRules }],
    workbook,
  )
}

/**
 * Write a multi-sheet XLSX workbook as a byte stream, pulling each
 * sheet's rows only as the consumer reads.
 *
 * {@link writeXlsxStream} streams a single sheet; a workbook that needs
 * two of them — a report and its rejects, say — had no streaming path at
 * all and had to fall back to {@link writeXlsx}, which builds the whole
 * object model first. Rows are not retained; styles and optional shared
 * strings span the workbook, while links/comments are retained only until
 * the current physical sheet's related parts have been emitted.
 *
 * ```ts
 * return new Response(
 *   writeXlsxStreamSheets([
 *     { name: "Accepted", rows: accepted, columns },
 *     { name: "Rejected", rows: rejected, columns },
 *   ]),
 *   { headers: { "content-type": XLSX_MIME } },
 * )
 * ```
 *
 * The sheet *list* is eager while the rows stay lazy. Sheets are few and
 * their names have to be validated before this returns, for the reason
 * given in {@link writeXlsxStream}: a generator that throws on first pull
 * throws after the caller already handed the stream to a `Response`.
 *
 * Notes:
 * - Sheets are written in the order given, and each one is drained before
 *   the next is pulled — the source for sheet 2 is not touched until
 *   sheet 1 runs out.
 * - `xl/styles.xml` and `xl/sharedStrings.xml` are workbook-wide, so a
 *   style or string is registered once no matter how many sheets use it.
 * - A sheet past its row cap rolls over exactly as in
 *   {@link writeXlsxStream}, and the rollover skips names already taken
 *   by another sheet.
 */
export function writeXlsxStreamSheets(
  sheets: readonly XlsxStreamSheet[],
  options: XlsxWriteStreamWorkbookOptions = {},
): ReadableStream<Uint8Array> {
  if (sheets.length === 0) {
    throw new InvalidArgumentError("A workbook needs at least one sheet")
  }

  validateSheetNames(sheets)
  for (const sheet of sheets) {
    validateRowSize(0, sheet.columns?.length ?? 0)
    // Columns precede sheetData; guessing a width would silently ignore
    // an accepted option, while reading all rows would destroy streaming.
    if (sheet.columns?.some((col: ColumnDef) => col.autoWidth && col.width === undefined)) {
      throw new InvalidArgumentError(
        "autoWidth requires buffered or incremental XLSX output; true streaming needs explicit widths",
      )
    }
    validateMaxRowsPerSheet(
      sheet.maxRowsPerSheet ?? options.maxRowsPerSheet ?? XLSX_MAX_ROWS_PER_SHEET,
    )
  }

  return zipStream(xlsxStreamEntries(sheets, options), { zip64: options.zip64 })
}

/** Lazily produce the ZIP entries for a streamed workbook. */
async function* xlsxStreamEntries(
  sheets: readonly XlsxStreamSheet[],
  options: XlsxWriteStreamWorkbookOptions,
): AsyncGenerator<ZipStreamEntry> {
  const dateSystem = options.dateSystem ?? "1900"
  const inlineStrings = options.stringMode !== "shared"
  const compress = options.compress ?? true

  // Workbook-wide parts: every sheet writes into the same two tables.
  const styles = createStylesCollector()
  const sharedStrings = createSharedStrings()
  let hasMetadata = false
  const commentIndices: number[] = []

  const sheetNames: string[] = []
  const takenNames = new Set(sheets.map((sheet) => sheet.name.toLowerCase()))

  for (const sheet of sheets) {
    const maxRowsPerSheet =
      sheet.maxRowsPerSheet ?? options.maxRowsPerSheet ?? XLSX_MAX_ROWS_PER_SHEET
    const repeatHeaders = sheet.repeatHeaders ?? options.repeatHeaders ?? true

    const serializer = new RowSerializer({
      columns: sheet.columns,
      dateSystem,
      inlineStrings,
      styles,
      sharedStrings,
    })
    const prelude = buildSheetPrelude(sheet.columns, sheet.freezePane, styles).join("")
    const cursor = createRowCursor(sheet.rows)

    let headerRow: CellInput[] | null = columnHeaders(sheet.columns)
    let sheetRowCount = 0
    let cellParts: SheetCellParts

    /** Stream one worksheet, stopping at the row cap or when rows run out. */
    async function* sheetChunks(part: number): AsyncGenerator<Uint8Array> {
      const chunker = new XmlChunker()

      const open =
        xmlDeclaration() +
        `<worksheet xmlns="${NS_SPREADSHEET}" xmlns:r="${NS_R}">` +
        prelude +
        "<sheetData>"
      const openChunk = chunker.push(open)
      if (openChunk) yield openChunk

      let rowIndex = 0

      // The first part emits the column header as a real row (matching the
      // incremental writer); later parts repeat it when asked to.
      if (headerRow && (part === 0 || repeatHeaders)) {
        // The header is the caller's row 0, whichever part it is repeated on.
        const xml = serializer.serializeRow(rowIndex, headerRow, cellParts, sheet.rowDefs?.get(0))
        rowIndex++
        if (part === 0) sheetRowCount++
        if (xml) {
          const chunk = chunker.push(xml)
          if (chunk) yield chunk
        }
      }

      while (rowIndex < maxRowsPerSheet) {
        const row = await cursor.next()
        if (row === undefined) break

        const values = Array.isArray(row) ? row : objectRow(row, requireColumns(sheet.columns))

        // Without column headers the first row doubles as the repeated header.
        if (sheetRowCount === 0 && !headerRow) {
          headerRow = snapshotRow(values)
        }

        // rowDefs are keyed by the caller's row number within this sheet,
        // which keeps counting past a rollover — `rowIndex` restarts at 0 on
        // every part.
        const xml = serializer.serializeRow(
          rowIndex,
          values,
          cellParts,
          sheet.rowDefs?.get(sheetRowCount),
        )
        rowIndex++
        sheetRowCount++
        if (xml) {
          const chunk = chunker.push(xml)
          if (chunk) yield chunk
        }
      }

      // Tail elements belong to the sheet's first part: a rollover splits one
      // logical sheet into several, and a range copied onto the continuation
      // would cover rows it was never meant to. Order matches the buffered
      // writer (and ECMA-376): autoFilter → mergeCells → conditionalFormatting.
      let sheetTail = "</sheetData>"
      if (part === 0 && sheet.autoFilter) {
        sheetTail += serializeAutoFilter(sheet.autoFilter)
      }
      if (part === 0 && sheet.merges?.length) {
        sheetTail += serializeMergeCells(sheet.merges)
      }
      if (part === 0 && sheet.conditionalRules?.length) {
        sheetTail += serializeConditionalFormatting(sheet.conditionalRules, styles).join("")
      }
      sheetTail += cellParts.toXml()
      sheetTail += "</worksheet>"
      const closeChunk = chunker.push(sheetTail)
      if (closeChunk) yield closeChunk
      const tail = chunker.flush()
      if (tail) yield tail
    }

    try {
      for (let part = 0; ; part++) {
        cellParts = new SheetCellParts()
        // The declared name is already reserved; a rollover has to claim
        // a free one.
        sheetNames.push(
          part === 0 ? truncateSheetName(sheet.name) : claimSheetName(sheet.name, part, takenNames),
        )

        yield {
          path: `xl/worksheets/sheet${sheetNames.length}.xml`,
          data: sheetChunks(part),
          compress,
        }
        const index = sheetNames.length
        if (cellParts.hasComments) commentIndices.push(index)
        for (const entry of cellParts.entries(index))
          yield { path: entry.path, data: encoder.encode(entry.xml), compress }
        hasMetadata ||= serializer.hasDynamicArray
        // The ZIP writer drains each entry before pulling the next, so by
        // now the sheet above is closed and the cursor is positioned on the
        // first row that didn't fit.
        if ((await cursor.peek()) === undefined) break
      }
    } finally {
      await cursor.close()
    }
  }

  const sheetCount = sheetNames.length
  const hasSharedStrings = sharedStrings.count() > 0
  const hasFeaturePropertyBag = styles.hasCheckboxFeature()

  yield { path: "xl/styles.xml", data: encoder.encode(styles.toXml()), compress }

  if (hasSharedStrings) {
    yield {
      path: "xl/sharedStrings.xml",
      data: encoder.encode(writeSharedStringsXml(sharedStrings)),
      compress,
    }
  }

  yield { path: "xl/theme/theme1.xml", data: encoder.encode(writeThemeXml()), compress }
  if (hasMetadata)
    yield { path: METADATA_PART_PATH, data: encoder.encode(writeMetadataXml()), compress }
  if (hasFeaturePropertyBag)
    yield { path: FPB_PART_PATH, data: encoder.encode(writeFeaturePropertyBagXml()), compress }

  yield {
    path: "xl/workbook.xml",
    data: encoder.encode(buildWorkbookXmlFor(sheetNames, dateSystem)),
    compress,
  }

  yield {
    path: "xl/_rels/workbook.xml.rels",
    data: encoder.encode(
      writeWorkbookRels(
        sheetCount,
        hasSharedStrings,
        undefined,
        hasFeaturePropertyBag,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        hasMetadata,
      ),
    ),
    compress,
  }

  yield { path: "_rels/.rels", data: encoder.encode(writeRootRels()), compress }

  // [Content_Types].xml lands last: the sheet count isn't known until the
  // rows run out. ZIP consumers resolve parts through the central
  // directory, so package order doesn't matter.
  yield {
    path: "[Content_Types].xml",
    data: encoder.encode(
      writeContentTypes({
        sheetCount,
        hasSharedStrings,
        commentIndices,
        hasFeaturePropertyBag,
        hasMetadata,
      }),
    ),
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
  if (!columns) throw new InvalidArgumentError("Object rows require columns with key accessors")
  return columns
}

/** Excel sheet names cap at 31 characters. */
function truncateSheetName(name: string): string {
  return name.length > MAX_SHEET_NAME_LENGTH ? name.slice(0, MAX_SHEET_NAME_LENGTH) : name
}

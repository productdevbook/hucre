// ── Worksheet Parser ─────────────────────────────────────────────────
// Parses xl/worksheets/sheetN.xml into a Sheet object.

import { cellError } from "../cell-error"
import type {
  ReadWarning,
  Sheet,
  Cell,
  CellValue,
  CellStyle,
  MergeRange,
  RichTextRun,
  FontStyle,
  Hyperlink,
  AutoFilter,
  ConditionalRule,
  ConditionalRuleType,
  DataValidation,
  ValidationType,
  ValidationOperator,
  SheetProtection,
  SheetView,
  Color,
  PageSetup,
  PageMargins,
  HeaderFooter,
  FreezePane,
  SplitPane,
  Sparkline,
} from "../_types"
import type { SharedString } from "./shared-strings"
import type { ParsedStyles } from "./styles"
import type { Relationship } from "./relationships"
import { resolveStyle, isDateStyle } from "./styles"
import { cloneCellStyle } from "../_style"
import { PAPER_SIZE_REVERSE } from "./worksheet-writer"
import { serialToDate } from "../_date"
import { parseSax, parseSaxStream, decodeOoxmlEscapes, type SaxHandlers } from "../xml/parser"
import { MAX_CELL_MAP_ENTRIES, MAX_COL_INDEX, MAX_ROW_INDEX, MAX_TOTAL_CELLS } from "../limits"
import { ParseError } from "../errors"

/**
 * The message for a sheet whose bounding box is over the limit.
 *
 * Pure, and exported, because the branch that matters cannot be reached
 * from a test workbook: telling a caller *not* to try `sparse` needs a
 * sheet with more than 16.7 million filled cells, which is a couple of
 * gigabytes to build for one string.
 *
 * The advice is the point. A sparse sheet — 82k values over a 305M-slot
 * box, 0.03% filled — wants `sparse: true`, and a dense one cannot use
 * it: the cell count that blew the box limit is the same count that
 * blows the `Map` behind `cells`. The message used to offer it either
 * way. See #501, #527.
 */
export function oversizeSheetMessage(
  name: string,
  rowCount: number,
  colCount: number,
  totalCells: number,
  cellCount: number,
  cellLimit: number,
): string {
  const density = totalCells > 0 ? (100 * cellCount) / totalCells : 100
  const sparseWouldFit = cellCount <= MAX_CELL_MAP_ENTRIES

  return (
    `Sheet "${name}" spans ${rowCount} rows x ${colCount} columns ` +
    `(${totalCells} cells, ${density.toFixed(2)}% of them filled), ` +
    `over the ${cellLimit} limit.\n` +
    `  - streamXlsxRows(input) reads it a row at a time, whatever the box.\n` +
    (sparseWouldFit
      ? `  - readXlsx(input, { sparse: true }) returns the cells and no grid.\n`
      : `  - \`sparse: true\` cannot help here: ${cellCount} filled cells is past ` +
        `the ${MAX_CELL_MAP_ENTRIES} a Map can hold.\n`) +
    `  - \`range\` or \`maxRows\` bound the area, if you know where the data is.\n` +
    `  - \`maxTotalCells\` raises the bound, if the sheet really is this large.`
  )
}

/**
 * Refuse the cell that would overflow `Sheet.cells`.
 *
 * V8 caps a `Map` at 2^24 entries and answers the next `set` with a raw
 * `RangeError: Map maximum size exceeded` — not a `HucreError`, naming
 * no sheet, saying nothing about spreadsheets. Checking the size first
 * costs one comparison per cell and turns that into a `ParseError` that
 * names where to go instead.
 *
 * `has` is only consulted at the boundary, so the common path is the
 * comparison alone.
 */
export function assertCellMapCapacity(
  cells: Map<string, Cell>,
  key: string,
  sheetName: string | undefined,
): void {
  if (cells.size < MAX_CELL_MAP_ENTRIES || cells.has(key)) return

  throw new ParseError(
    `Sheet "${sheetName ?? "?"}" has more than ${MAX_CELL_MAP_ENTRIES} filled cells, ` +
      `which is the most \`Sheet.cells\` can hold — a Map caps at 2^24 entries.\n` +
      `  - streamXlsxRows(input) reads it a row at a time and has no such bound.\n` +
      `The file is not damaged; it is larger than this model.`,
  )
}

// ── Types ────────────────────────────────────────────────────────────

export interface WorksheetContext {
  sharedStrings: SharedString[]
  styles: ParsedStyles | null
  readStyles: boolean
  dateSystem: "1900" | "1904"
  /** Worksheet-level relationships (from xl/worksheets/_rels/sheetN.xml.rels) */
  worksheetRels?: Relationship[]
  /** Maximum number of data rows to parse. Default: unlimited */
  maxRows?: number
  /** Cell range filter (e.g. "A1:D10"). Only cells within this range are returned. */
  range?: string
  /** Bounding-box ceiling; see ReadOptions.maxTotalCells. Default {@link MAX_TOTAL_CELLS}. */
  maxTotalCells?: number
  /** Skip the dense grid and return cells only; see ReadOptions.sparse. */
  sparse?: boolean
  /** Name of the sheet being parsed, so a warning can say where it was. */
  sheetName?: string
  /** Where a dropped reference is reported; see ReadOptions.onWarning. */
  onWarning?: (warning: ReadWarning) => void
  /**
   * One-based `cm` indexes that xl/metadata.xml resolves to a
   * dynamic-array (XLDAPR) record. Undefined when the package ships no
   * metadata part — see {@link isDynamicArrayCm}.
   */
  dynamicArrayCm?: Set<number>
}

/**
 * Decide whether a cell's `cm` index marks a dynamic array.
 *
 * With a metadata part present the index is resolved properly. Without
 * one there is nothing to resolve against, and the only producer known
 * to emit a bare `cm` is hucre itself before #423 — which always meant
 * "dynamic array" — so any non-zero index is taken at its word.
 */
function isDynamicArrayCm(raw: string | undefined, ctx: WorksheetContext): boolean {
  if (raw === undefined) return false
  const index = Number(raw)
  if (!Number.isFinite(index) || index <= 0) return false
  return ctx.dynamicArrayCm ? ctx.dynamicArrayCm.has(index) : true
}

// ── Cell Reference Parsing ───────────────────────────────────────────

/**
 * Parse a cell reference like "A1", "Z1", "AA1", "AZ1", "AAA1"
 * into 0-based { row, col }.
 */
export function parseCellRef(ref: string): { row: number; col: number } {
  let i = 0
  let col = 0

  // Parse column letters
  while (i < ref.length) {
    const code = ref.charCodeAt(i)
    if (code >= 65 && code <= 90) {
      // A-Z
      col = col * 26 + (code - 64)
      i++
    } else if (code >= 97 && code <= 122) {
      // a-z
      col = col * 26 + (code - 96)
      i++
    } else {
      break
    }
  }

  col-- // Convert to 0-based

  // Parse row number
  const row = Number(ref.slice(i)) - 1 // Convert to 0-based

  return { row, col }
}

/**
 * Bound a 1-based `<col min>` / `<col max>` attribute so it can safely
 * drive a loop.
 *
 * Without this a single `max="1e999"` makes the expansion loop run
 * forever (#355). An overflowing magnitude and outright garbage are
 * treated differently on purpose: `1e999` parses to `Infinity`, which
 * says "wider than representable", so it clamps to the last column the
 * same way an absurd-but-finite `99999999` does. A value that is not a
 * number at all carries no such intent and collapses to `fallback`,
 * dropping the malformed element.
 */
function clampColumnBound(raw: string | undefined, fallback: number): number {
  const value = Number(raw ?? "")
  if (Number.isNaN(value)) return fallback
  // Math.trunc leaves ±Infinity intact, so -Infinity fails the `< 1`
  // guard below and +Infinity is capped by the Math.min.
  const truncated = Math.trunc(value)
  if (truncated < 1) return fallback
  return Math.min(truncated, MAX_COL_INDEX + 1)
}

/**
 * Parse a range reference like "A1:B2" into start and end positions.
 */
function parseRangeRef(ref: string): MergeRange {
  const parts = ref.split(":")
  const start = parseCellRef(parts[0])
  const end = parts.length > 1 ? parseCellRef(parts[1]) : start

  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  }
}

// ── SAX-based Worksheet Parser ───────────────────────────────────────

/**
 * The SAX handlers for one worksheet, and the finalisation that turns
 * what they collected into a {@link Sheet}.
 *
 * Split out so the buffered and the streaming reader drive the *same*
 * handler set rather than two copies of it. A worksheet part over the
 * string ceiling has to be parsed a chunk at a time (#503), and a second
 * implementation of these 850 lines would be a second set of answers for
 * every field of the model — the two would drift on the first field
 * added to one and forgotten in the other, which is the failure mode
 * `CONTRIBUTING.md` calls the registers. There is exactly one
 * implementation; only the driver differs.
 */
function worksheetParser(
  name: string,
  ctx: WorksheetContext,
): { handlers: SaxHandlers; finish: () => Sheet } {
  const rows: CellValue[][] = []
  const cells = new Map<string, Cell>()
  const merges: MergeRange[] = []
  let maxCol = -1
  let maxRow = -1
  /** Cells that carried something — the numerator of the fill factor. */
  let cellCount = 0
  let hasCells = false

  // Range filter — parse once, use in cell processing
  let rangeFilter: MergeRange | undefined
  if (ctx.range) {
    rangeFilter = parseRangeRef(ctx.range)
  }

  // Hyperlinks parsed from <hyperlinks> section
  interface RawHyperlink {
    ref: string
    rId?: string
    location?: string
    tooltip?: string
    display?: string
  }
  const rawHyperlinks: RawHyperlink[] = []

  // Data validations parsed from <dataValidations> section
  const dataValidations: DataValidation[] = []

  // Conditional formatting rules parsed from <conditionalFormatting> sections
  const conditionalRules: ConditionalRule[] = []

  // Auto filter parsed from <autoFilter> element
  let autoFilter: AutoFilter | undefined
  let inAutoFilter = false
  let currentFilterColIndex = -1
  let currentFilterValues: string[] = []

  // Sheet protection parsed from <sheetProtection> element
  let sheetProtection: SheetProtection | undefined

  // Sheet view settings (gridlines, zoom, RTL, tab color)
  let sheetView: SheetView | undefined
  let inSheetPr = false
  let fitToPageFlag = false
  let outlineProperties: import("../_types").OutlineProperties | undefined

  // Freeze/Split pane parsed from <pane> element
  let freezePane: FreezePane | undefined
  let splitPane: SplitPane | undefined

  // Page setup / print settings
  let pageSetup: PageSetup | undefined
  let pageMargins: PageMargins | undefined
  let headerFooter: HeaderFooter | undefined

  // Page breaks
  const rowBreaks: number[] = []
  const colBreaks: number[] = []
  let inRowBreaks = false
  let inColBreaks = false

  // Sparkline SAX state
  const sparklines: Sparkline[] = []
  let inSparklineGroups = false
  let inSparklineGroup = false
  let inSparkline = false
  let inSparklineF = false
  let inSparklineSqref = false
  let sparklineGroupType = ""
  let sparklineGroupColor: Color | undefined
  let sparklineGroupMarkers = false
  let sparklineF = ""
  let sparklineSqref = ""

  // Header/footer SAX state
  let inHeaderFooter = false
  let inOddHeader = false
  let inOddFooter = false
  let inEvenHeader = false
  let inEvenFooter = false
  let inFirstHeader = false
  let inFirstFooter = false
  let hfText = ""

  // Row limit tracking (maxRows option)
  const maxRowsLimit = ctx.maxRows ?? 0 // 0 = unlimited
  let dataRowCount = 0

  // Row definitions (height, hidden, outlineLevel, collapsed)
  const rowDefs = new Map<number, import("../_types").RowDef>()

  // Column definitions (width, hidden, outlineLevel, collapsed) parsed from <col> elements
  const columnDefs: import("../_types").ColumnDef[] = []
  let defaultRowHeight: number | undefined
  let defaultColWidth: number | undefined
  let inCols = false

  // SAX parsing state
  let inSheetData = false
  let inRow = false
  let inCell = false
  let inValue = false
  let inFormula = false
  let inInlineStr = false
  let inInlineT = false
  let inMergeCells = false
  let inHyperlinks = false
  let inDataValidations = false
  let inDataValidation = false
  let inDvFormula1 = false
  let inDvFormula2 = false

  // Current data validation state
  let dvFormula1Text = ""
  let dvFormula2Text = ""
  let dvAttrs: Record<string, string> = {}

  // Conditional formatting SAX state
  let inConditionalFormatting = false
  let cfSqref = ""
  let inCfRule = false
  let cfRuleAttrs: Record<string, string> = {}
  let inCfFormula = false
  let cfFormulaText = ""
  let cfFormulas: string[] = []
  // colorScale state
  let inColorScale = false
  let csCfvos: Array<{ type: string; value?: string }> = []
  let csColors: Color[] = []
  // dataBar state
  let inDataBar = false
  let dbCfvos: Array<{ type: string; value?: string }> = []
  let dbColor: Color | undefined
  // iconSet state
  let inIconSet = false
  let isAttrs: Record<string, string> = {}
  let isCfvos: Array<{ type: string; value?: string }> = []

  // Rich text in inline strings
  let inInlineR = false
  let inInlineRPr = false
  let inInlineRT = false

  // Current cell state
  let cellRef = ""
  // Implicit-column tracking for cells lacking an `r` attribute (parity with
  // the streaming reader). Reset at the start of each row.
  let currentRowNum = 0 // 1-based row number from the row's `r` attr
  let implicitCol = 0 // 0-based next implicit column index
  let cellType = ""
  let cellStyleIndex = -1
  let cellValueText = ""
  let cellFormulaText = ""
  let cellFormulaType = "" // "shared", "array", or ""
  let cellFormulaSi = -1 // shared formula index
  let cellFormulaRef = "" // formula ref range
  let cellFormulaCm = false // dynamic array flag
  let inlineText = ""

  // Inline rich text state
  let inlineRichText: RichTextRun[] = []
  let currentRunText = ""
  let currentRunFont: FontStyle | undefined
  let _fontPropTag = ""

  const handlers: SaxHandlers = {
    onOpenTag(tag, attrs) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag

      switch (local) {
        case "cols":
          inCols = true
          break
        case "col":
          if (inCols) {
            // `min`/`max` are 1-based and drive the loop below, so they are
            // bounded before use: `max="1e999"` parses to Infinity and would
            // spin forever, and any finite value past Excel's column count
            // would allocate until the heap gives out. A range wider than the
            // sheet can hold is malformed, so it is clamped rather than
            // honoured. See #355.
            const minCol = clampColumnBound(attrs["min"], 1)
            const maxCol2 = clampColumnBound(attrs["max"], 0)
            const width = attrs["width"] ? Number(attrs["width"]) : undefined
            const hidden = attrs["hidden"] === "1" || attrs["hidden"] === "true"
            const outlineLevel = attrs["outlineLevel"] ? Number(attrs["outlineLevel"]) : undefined
            const collapsed = attrs["collapsed"] === "1" || attrs["collapsed"] === "true"
            const bestFit = attrs["bestFit"] === "1" || attrs["bestFit"] === "true"
            // `style` is the column's default cell format — the thing that
            // makes a whole column currency, including the cells nobody has
            // typed in. It was read by neither side; see #439 §W.
            const styleIndex = attrs["style"] !== undefined ? Number(attrs["style"]) : undefined
            const columnStyle =
              ctx.readStyles && ctx.styles && styleIndex !== undefined && styleIndex >= 0
                ? resolveStyle(ctx.styles, styleIndex)
                : undefined

            // Expand column range (min and max are 1-based in OOXML)
            for (let c = minCol; c <= maxCol2; c++) {
              const idx = c - 1 // Convert to 0-based
              // Ensure the array is long enough
              while (columnDefs.length <= idx) {
                columnDefs.push({})
              }
              const def: import("../_types").ColumnDef = {}
              if (width !== undefined && !Number.isNaN(width)) def.width = width
              if (hidden) def.hidden = true
              if (outlineLevel !== undefined && !Number.isNaN(outlineLevel) && outlineLevel > 0) {
                def.outlineLevel = outlineLevel
              }
              if (collapsed) def.collapsed = true
              if (bestFit) def.autoWidth = true
              if (columnStyle && Object.keys(columnStyle).length > 0) {
                // Each column gets its own copy: `columnDefs` is the caller's
                // to edit, and `resolveStyle` hands out shared records.
                def.style = cloneCellStyle(columnStyle)
              }
              if (Object.keys(def).length > 0) {
                columnDefs[idx] = def
              }
            }
          }
          break
        case "sheetFormatPr": {
          const dh = attrs["defaultRowHeight"]
          const dw = attrs["defaultColWidth"]
          // Excel writes 15 whether or not the sheet means anything by it,
          // so only a value that differs is a statement worth surfacing —
          // otherwise every sheet would come back carrying a default it
          // never set.
          if (dh !== undefined) {
            const height = Number(dh)
            if (Number.isFinite(height) && height > 0 && height !== 15) {
              defaultRowHeight = height
            }
          }
          if (dw !== undefined) {
            const widthValue = Number(dw)
            if (Number.isFinite(widthValue) && widthValue > 0) defaultColWidth = widthValue
          }
          break
        }
        case "sheetData":
          inSheetData = true
          break
        case "row":
          if (inSheetData) {
            // Check maxRows limit
            if (maxRowsLimit > 0 && dataRowCount >= maxRowsLimit) {
              break
            }
            inRow = true
            // Track the row number and reset implicit-column counter so cells
            // lacking an `r` attribute get sequential columns within this row.
            currentRowNum = Number(attrs["r"]) || currentRowNum + 1
            implicitCol = 0
            // Parse row-level attributes: ht, customHeight, hidden
            if (
              attrs["ht"] &&
              (attrs["customHeight"] === "1" || attrs["customHeight"] === "true")
            ) {
              const rowNum = Number(attrs["r"]) - 1 // 0-based
              const height = Number(attrs["ht"])
              if (!Number.isNaN(rowNum) && !Number.isNaN(height)) {
                const existing = rowDefs.get(rowNum) ?? {}
                existing.height = height
                rowDefs.set(rowNum, existing)
              }
            }
            if (attrs["hidden"] === "1" || attrs["hidden"] === "true") {
              const rowNum = Number(attrs["r"]) - 1
              if (!Number.isNaN(rowNum)) {
                const existing = rowDefs.get(rowNum) ?? {}
                existing.hidden = true
                rowDefs.set(rowNum, existing)
              }
            }
            if (attrs["outlineLevel"]) {
              const rowNum = Number(attrs["r"]) - 1
              const level = Number(attrs["outlineLevel"])
              if (!Number.isNaN(rowNum) && !Number.isNaN(level) && level > 0) {
                const existing = rowDefs.get(rowNum) ?? {}
                existing.outlineLevel = level
                rowDefs.set(rowNum, existing)
              }
            }
            if (attrs["collapsed"] === "1" || attrs["collapsed"] === "true") {
              const rowNum = Number(attrs["r"]) - 1
              if (!Number.isNaN(rowNum)) {
                const existing = rowDefs.get(rowNum) ?? {}
                existing.collapsed = true
                rowDefs.set(rowNum, existing)
              }
            }
          }
          break
        case "c":
          if (inRow) {
            inCell = true
            cellRef = attrs["r"] ?? ""
            cellType = attrs["t"] ?? ""
            cellStyleIndex = attrs["s"] ? Number(attrs["s"]) : -1
            cellValueText = ""
            cellFormulaText = ""
            cellFormulaType = ""
            cellFormulaSi = -1
            cellFormulaRef = ""
            // `cm` lives on `<c>` (§18.3.1.4) and is a one-based index
            // into xl/metadata.xml's cellMetadata collection, not a
            // boolean. hucre used to both write and read it on `<f>`,
            // which round-tripped with itself and with nothing else
            // (#423).
            cellFormulaCm = isDynamicArrayCm(attrs["cm"], ctx)
            inlineText = ""
            inlineRichText = []
          }
          break
        case "v":
          if (inCell) inValue = true
          break
        case "f":
          if (inSparkline) {
            inSparklineF = true
            sparklineF = ""
          } else if (inCell) {
            inFormula = true
            cellFormulaType = attrs["t"] ?? ""
            if (attrs["si"] !== undefined) {
              cellFormulaSi = Number(attrs["si"])
            }
            if (attrs["ref"]) {
              cellFormulaRef = attrs["ref"]
            }
            // Every hucre release up to 0.6 wrote the marker here, so
            // keep honouring it — those files are in the wild and the
            // attribute is meaningless on `<f>` for any other reason.
            if (attrs["cm"] !== undefined && isDynamicArrayCm(attrs["cm"], ctx)) {
              cellFormulaCm = true
            }
          }
          break
        case "is":
          if (inCell) inInlineStr = true
          break
        case "t":
          if (inInlineStr && !inInlineR) {
            inInlineT = true
          } else if (inInlineR) {
            inInlineRT = true
          }
          break
        case "r":
          if (inInlineStr) {
            inInlineR = true
            currentRunText = ""
            currentRunFont = undefined
          }
          break
        case "rPr":
          if (inInlineR) {
            inInlineRPr = true
            currentRunFont = {}
          }
          break
        case "sheetPr":
          inSheetPr = true
          break
        case "outlinePr":
          // Write-only until now: the type and the writer existed, but
          // nothing parsed it, so Sheet.outlineProperties was always
          // undefined and open -> save could not preserve it. See #359.
          if (inSheetPr) {
            const outline: import("../_types").OutlineProperties = {}
            if (attrs["summaryBelow"] !== undefined) {
              outline.summaryBelow =
                attrs["summaryBelow"] === "1" || attrs["summaryBelow"] === "true"
            }
            if (attrs["summaryRight"] !== undefined) {
              outline.summaryRight =
                attrs["summaryRight"] === "1" || attrs["summaryRight"] === "true"
            }
            if (Object.keys(outline).length > 0) outlineProperties = outline
          }
          break
        case "pageSetUpPr":
          // The real home of the fit-to-page toggle: <pageSetup> only
          // carries the page counts. Recorded separately from the
          // <pageSetup> attributes because the two elements are far apart
          // in the document and either may be absent. See #407.
          if (inSheetPr) {
            fitToPageFlag = attrs["fitToPage"] === "1" || attrs["fitToPage"] === "true"
          }
          break
        case "tabColor":
          if (inSheetPr) {
            if (!sheetView) sheetView = {}
            sheetView.tabColor = parseColorAttrs(attrs)
          }
          break
        case "sheetView":
          if (!inSheetData) {
            if (!sheetView) sheetView = {}
            if (attrs["showGridLines"] === "0" || attrs["showGridLines"] === "false") {
              sheetView.showGridLines = false
            }
            if (attrs["showRowColHeaders"] === "0" || attrs["showRowColHeaders"] === "false") {
              sheetView.showRowColHeaders = false
            }
            if (attrs["zoomScale"]) {
              const zoom = Number(attrs["zoomScale"])
              if (!Number.isNaN(zoom)) {
                sheetView.zoomScale = zoom
              }
            }
            if (attrs["rightToLeft"] === "1" || attrs["rightToLeft"] === "true") {
              sheetView.rightToLeft = true
            }
          }
          break
        case "pane":
          if (!inSheetData) {
            const state = attrs["state"]
            if (state === "frozen" || state === "frozenSplit") {
              // Freeze pane
              const xSplit = Number(attrs["xSplit"] || "0")
              const ySplit = Number(attrs["ySplit"] || "0")
              if (xSplit > 0 || ySplit > 0) {
                freezePane = {}
                if (xSplit > 0) freezePane.columns = xSplit
                if (ySplit > 0) freezePane.rows = ySplit
              }
            } else if (state === "split") {
              // Split pane
              const xSplit = Number(attrs["xSplit"] || "0")
              const ySplit = Number(attrs["ySplit"] || "0")
              if (xSplit > 0 || ySplit > 0) {
                splitPane = {}
                if (xSplit > 0) splitPane.xSplit = xSplit
                if (ySplit > 0) splitPane.ySplit = ySplit
              }
            }
          }
          break
        case "sheetProtection":
          sheetProtection = parseSheetProtectionAttrs(attrs)
          break
        case "autoFilter":
          if (attrs["ref"]) {
            autoFilter = { range: attrs["ref"] }
            inAutoFilter = true
          }
          break
        case "filterColumn":
          if (inAutoFilter && attrs["colId"] !== undefined) {
            currentFilterColIndex = Number(attrs["colId"])
            currentFilterValues = []
          }
          break
        case "filter":
          if (inAutoFilter && currentFilterColIndex >= 0 && attrs["val"] !== undefined) {
            currentFilterValues.push(attrs["val"])
          }
          break
        case "mergeCells":
          inMergeCells = true
          break
        case "mergeCell":
          if (inMergeCells && attrs["ref"]) {
            merges.push(parseRangeRef(attrs["ref"]))
          }
          break
        case "hyperlinks":
          inHyperlinks = true
          break
        case "hyperlink":
          if (inHyperlinks && attrs["ref"]) {
            const hl: RawHyperlink = { ref: attrs["ref"] }
            // r:id for external hyperlinks
            const rId = attrs["r:id"] ?? attrs["R:id"]
            if (rId) hl.rId = rId
            if (attrs["location"]) hl.location = attrs["location"]
            if (attrs["tooltip"]) hl.tooltip = attrs["tooltip"]
            if (attrs["display"]) hl.display = attrs["display"]
            rawHyperlinks.push(hl)
          }
          break
        case "conditionalFormatting":
          inConditionalFormatting = true
          cfSqref = attrs["sqref"] ?? ""
          break
        case "cfRule":
          if (inConditionalFormatting) {
            inCfRule = true
            cfRuleAttrs = { ...attrs }
            cfFormulas = []
            csCfvos = []
            csColors = []
            dbCfvos = []
            dbColor = undefined
            isCfvos = []
            isAttrs = {}
          }
          break
        case "colorScale":
          if (inCfRule) {
            inColorScale = true
            csCfvos = []
            csColors = []
          }
          break
        case "cfvo":
          if (inColorScale) {
            csCfvos.push({ type: attrs["type"] ?? "min", value: attrs["val"] })
          } else if (inDataBar) {
            dbCfvos.push({ type: attrs["type"] ?? "min", value: attrs["val"] })
          } else if (inIconSet) {
            isCfvos.push({ type: attrs["type"] ?? "min", value: attrs["val"] })
          }
          break
        case "dataBar":
          if (inCfRule) {
            inDataBar = true
            dbCfvos = []
            dbColor = undefined
          }
          break
        case "iconSet":
          if (inCfRule) {
            inIconSet = true
            isAttrs = { ...attrs }
            isCfvos = []
          }
          break
        case "dataValidations":
          inDataValidations = true
          break
        case "dataValidation":
          if (inDataValidations) {
            inDataValidation = true
            dvAttrs = { ...attrs }
            dvFormula1Text = ""
            dvFormula2Text = ""
          }
          break
        case "formula1":
          if (inDataValidation) inDvFormula1 = true
          break
        case "formula2":
          if (inDataValidation) inDvFormula2 = true
          break
        case "pageMargins":
          pageMargins = parsePageMarginsAttrs(attrs)
          break
        case "pageSetup":
          // Merge, don't replace: <printOptions> is written before
          // <pageSetup> in a worksheet, so assigning here dropped
          // whatever it had already contributed. See #360.
          pageSetup = { ...pageSetup, ...parsePageSetupAttrs(attrs, ctx) }
          break
        case "printOptions":
          // <printOptions> was never parsed, so showGridLines and
          // showRowColHeaders were write-only fields that always read back
          // as undefined. It can appear before or after <pageSetup>, so
          // merge rather than assign. See #360.
          pageSetup = applyPrintOptionsAttrs(pageSetup, attrs)
          break
        case "headerFooter":
          inHeaderFooter = true
          headerFooter = {}
          if (attrs["differentOddEven"] === "1" || attrs["differentOddEven"] === "true") {
            headerFooter.differentOddEven = true
          }
          if (attrs["differentFirst"] === "1" || attrs["differentFirst"] === "true") {
            headerFooter.differentFirst = true
          }
          break
        case "oddHeader":
          if (inHeaderFooter) {
            inOddHeader = true
            hfText = ""
          }
          break
        case "oddFooter":
          if (inHeaderFooter) {
            inOddFooter = true
            hfText = ""
          }
          break
        case "evenHeader":
          if (inHeaderFooter) {
            inEvenHeader = true
            hfText = ""
          }
          break
        case "evenFooter":
          if (inHeaderFooter) {
            inEvenFooter = true
            hfText = ""
          }
          break
        case "firstHeader":
          if (inHeaderFooter) {
            inFirstHeader = true
            hfText = ""
          }
          break
        case "firstFooter":
          if (inHeaderFooter) {
            inFirstFooter = true
            hfText = ""
          }
          break
        case "rowBreaks":
          inRowBreaks = true
          break
        case "colBreaks":
          inColBreaks = true
          break
        case "brk":
          if (inRowBreaks || inColBreaks) {
            const brkId = attrs["id"]
            if (brkId) {
              const index = Number(brkId) - 1 // Convert to 0-based
              if (inRowBreaks) {
                rowBreaks.push(index)
              } else {
                colBreaks.push(index)
              }
            }
          }
          break
        case "color":
          if (inColorScale) {
            csColors.push(parseColorAttrs(attrs))
          } else if (inDataBar) {
            dbColor = parseColorAttrs(attrs)
          } else if (inInlineRPr && currentRunFont) {
            applyFontProp(currentRunFont, local, attrs)
          }
          break
        case "formula":
          if (inCfRule && !inDataValidation) {
            inCfFormula = true
            cfFormulaText = ""
          }
          break
        case "sparklineGroups":
          inSparklineGroups = true
          break
        case "sparklineGroup":
          if (inSparklineGroups) {
            inSparklineGroup = true
            sparklineGroupType = attrs["type"] ?? "line"
            sparklineGroupColor = undefined
            sparklineGroupMarkers = attrs["markers"] === "1" || attrs["markers"] === "true"
          }
          break
        case "colorSeries":
          if (inSparklineGroup) sparklineGroupColor = parseColorAttrs(attrs)
          break
        case "sparkline":
          if (inSparklineGroup) {
            inSparkline = true
            sparklineF = ""
            sparklineSqref = ""
          }
          break
        default:
          // Handle xm:sqref inside sparkline
          if (inSparkline && local === "sqref") {
            inSparklineSqref = true
            sparklineSqref = ""
            break
          }
          // Handle font property tags inside rPr
          if (inInlineRPr && currentRunFont) {
            _fontPropTag = local
            applyFontProp(currentRunFont, local, attrs)
          }
          break
      }
    },

    onText(text) {
      if (inValue) {
        cellValueText += text
      } else if (inFormula) {
        cellFormulaText += text
      } else if (inCfFormula) {
        cfFormulaText += text
      } else if (inInlineT) {
        inlineText += text
      } else if (inInlineRT) {
        currentRunText += text
      } else if (inDvFormula1) {
        dvFormula1Text += text
      } else if (inDvFormula2) {
        dvFormula2Text += text
      } else if (
        inOddHeader ||
        inOddFooter ||
        inEvenHeader ||
        inEvenFooter ||
        inFirstHeader ||
        inFirstFooter
      ) {
        hfText += text
      } else if (inSparklineF) {
        sparklineF += text
      } else if (inSparklineSqref) {
        sparklineSqref += text
      }
    },

    onCloseTag(tag) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag

      switch (local) {
        case "cols":
          inCols = false
          break
        case "sheetData":
          inSheetData = false
          break
        case "row":
          if (inRow) {
            dataRowCount++
          }
          inRow = false
          break
        case "c":
          if (inCell) {
            // Resolve effective row/col. Cells with an `r` attribute use it;
            // cells without one fall back to implicit position within the row.
            const effRow = cellRef ? parseCellRef(cellRef).row : currentRowNum - 1
            const effCol = cellRef ? parseCellRef(cellRef).col : implicitCol
            // Advance implicit column for the next cell in this row.
            implicitCol = effCol + 1

            // Skip cells outside the range filter
            let skipCell = false
            if (rangeFilter) {
              if (
                effRow < rangeFilter.startRow ||
                effRow > rangeFilter.endRow ||
                effCol < rangeFilter.startCol ||
                effCol > rangeFilter.endCol
              ) {
                skipCell = true
              }
            }

            // A cell that will contribute nothing to the model must not
            // extend the sheet — nor allocate the row it sits in.
            //
            // Excel writes a self-closing `<c r="WVF45" s="3"/>` for every
            // position formatting was ever applied to, and a real
            // packing-list workbook had 145,315 of them against 197
            // values: `rows` came back 45 x 16,126 and `writeCsv` of it
            // was 727 KB, 99.75% bare commas, from 1.8 KB of data.
            //
            // Under `readStyles: true` those cells do carry information
            // and still count. See #492.
            const carriesData =
              cellValueText !== "" ||
              inlineText !== "" ||
              inlineRichText.length > 0 ||
              cellFormulaText !== "" ||
              cellType === "e" ||
              // An empty *inline* string is still a string. The producer
              // wrote `t="inlineStr"` and an `<is>` to say so, which is
              // not the contentless `<c r="WVF45" s="3"/>` this guard is
              // for. Deciding from the collected text alone made the two
              // spellings of one value disagree: a shared string carries
              // its index here, non-empty even when the string is empty,
              // so `""` survived that way and vanished the other.
              //
              // It reached hucre's own writers, because
              // `writeXlsxStream` defaults to inline strings.
              cellType === "inlineStr" ||
              (ctx.readStyles && cellStyleIndex >= 0) ||
              (ctx.styles && cellStyleIndex >= 0
                ? (ctx.styles.cellXfs[cellStyleIndex]?.hasCheckboxFeature ?? false)
                : false)

            if (!skipCell && carriesData) {
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
                cellFormulaType,
                cellFormulaSi,
                cellFormulaRef,
                cellFormulaCm,
                effRow,
                effCol,
              )
              // Track max dimensions
              if (effRow >= 0 && effCol >= 0) {
                if (effCol > maxCol) maxCol = effCol
                if (effRow > maxRow) maxRow = effRow
                hasCells = true
                cellCount++
              }
            }
            inCell = false
          }
          break
        case "v":
          inValue = false
          break
        case "f":
          if (inSparklineF) {
            inSparklineF = false
          } else {
            inFormula = false
          }
          break
        case "is":
          inInlineStr = false
          break
        case "t":
          if (inInlineRT) {
            inInlineRT = false
          } else if (inInlineT) {
            inInlineT = false
          }
          break
        case "r":
          if (inInlineR) {
            const decodedRunText = decodeOoxmlEscapes(currentRunText)
            inlineRichText.push(
              currentRunFont
                ? { text: decodedRunText, font: currentRunFont }
                : { text: decodedRunText },
            )
            inInlineR = false
          }
          break
        case "rPr":
          inInlineRPr = false
          break
        case "sheetPr":
          inSheetPr = false
          break
        case "mergeCells":
          inMergeCells = false
          break
        case "autoFilter":
          inAutoFilter = false
          break
        case "filterColumn":
          if (
            inAutoFilter &&
            autoFilter &&
            currentFilterColIndex >= 0 &&
            currentFilterValues.length > 0
          ) {
            if (!autoFilter.columns) autoFilter.columns = []
            autoFilter.columns.push({
              colIndex: currentFilterColIndex,
              filters: currentFilterValues,
            })
          }
          currentFilterColIndex = -1
          currentFilterValues = []
          break
        case "hyperlinks":
          inHyperlinks = false
          break
        case "conditionalFormatting":
          inConditionalFormatting = false
          break
        case "cfRule":
          if (inCfRule) {
            const cfRule = buildConditionalRule(
              cfRuleAttrs,
              cfSqref,
              cfFormulas,
              csCfvos,
              csColors,
              dbCfvos,
              dbColor,
              isCfvos,
              isAttrs,
              ctx.styles?.dxfs,
              ctx,
            )
            if (cfRule) {
              conditionalRules.push(cfRule)
            }
            inCfRule = false
          }
          break
        case "colorScale":
          inColorScale = false
          break
        case "dataBar":
          if (inCfRule) {
            inDataBar = false
          }
          break
        case "iconSet":
          inIconSet = false
          break
        case "formula":
          if (inCfFormula) {
            cfFormulas.push(cfFormulaText)
            inCfFormula = false
          }
          break
        case "dataValidations":
          inDataValidations = false
          break
        case "dataValidation":
          if (inDataValidation) {
            const dv = buildDataValidation(dvAttrs, dvFormula1Text, dvFormula2Text)
            if (dv) {
              dataValidations.push(dv)
            }
            inDataValidation = false
          }
          break
        case "formula1":
          inDvFormula1 = false
          break
        case "formula2":
          inDvFormula2 = false
          break
        case "headerFooter":
          inHeaderFooter = false
          break
        case "oddHeader":
          if (inOddHeader && headerFooter) {
            headerFooter.oddHeader = hfText
            inOddHeader = false
          }
          break
        case "oddFooter":
          if (inOddFooter && headerFooter) {
            headerFooter.oddFooter = hfText
            inOddFooter = false
          }
          break
        case "evenHeader":
          if (inEvenHeader && headerFooter) {
            headerFooter.evenHeader = hfText
            inEvenHeader = false
          }
          break
        case "evenFooter":
          if (inEvenFooter && headerFooter) {
            headerFooter.evenFooter = hfText
            inEvenFooter = false
          }
          break
        case "firstHeader":
          if (inFirstHeader && headerFooter) {
            headerFooter.firstHeader = hfText
            inFirstHeader = false
          }
          break
        case "firstFooter":
          if (inFirstFooter && headerFooter) {
            headerFooter.firstFooter = hfText
            inFirstFooter = false
          }
          break
        case "rowBreaks":
          inRowBreaks = false
          break
        case "colBreaks":
          inColBreaks = false
          break
        case "sparklineGroups":
          inSparklineGroups = false
          break
        case "sparklineGroup":
          inSparklineGroup = false
          break
        case "sparkline":
          if (inSparkline && sparklineSqref) {
            const sp: Sparkline = {
              location: sparklineSqref,
              dataRange: sparklineF,
            }
            if (sparklineGroupType && sparklineGroupType !== "line") {
              sp.type = sparklineGroupType as Sparkline["type"]
            }
            if (sparklineGroupColor) {
              sp.color = sparklineGroupColor
            }
            if (sparklineGroupMarkers) {
              sp.markers = true
            }
            sparklines.push(sp)
          }
          inSparkline = false
          break
        default:
          // Handle xm:sqref close inside sparkline
          if (inSparkline && local === "sqref") {
            inSparklineSqref = false
            break
          }
          if (inInlineRPr) {
            _fontPropTag = ""
          }
          break
      }
    },
  }

  function finish(): Sheet {
    // Ensure all rows have consistent length. Not in sparse mode: there is
    // no grid, which is the whole point, and the bounding-box limit below
    // has nothing to guard.
    if (hasCells && !ctx.sparse) {
      const colCount = maxCol + 1
      // The cost of a sheet is its bounding box, not its cell count — the
      // loop below fills every slot in it. Two in-bounds cells at opposite
      // corners describe 1.7e10 slots, which V8 answers with an OOM the
      // caller cannot catch, so the product is checked before allocating.
      const totalCells = (maxRow + 1) * colCount
      const cellLimit = ctx.maxTotalCells ?? MAX_TOTAL_CELLS
      if (totalCells > cellLimit) {
        // The options this used to name were the wrong three for the case
        // that actually hits it. A *sparse* sheet — 82k values scattered
        // over a 305M-slot box, 0.03% fill — is not large, so raising
        // `maxTotalCells` trades a clean error for a multi-gigabyte
        // allocation; `range` needs the caller to already know where the
        // data is; and `maxRows` bounds rows when the problem is columns.
        //
        // `streamXlsxRows` reads exactly this file today, one row at a
        // time, and the message never mentioned it. See #501.
        throw new ParseError(
          oversizeSheetMessage(name, maxRow + 1, colCount, totalCells, cellCount, cellLimit),
        )
      }
      for (let r = 0; r <= maxRow; r++) {
        if (!rows[r]) {
          rows[r] = Array.from({ length: colCount }, () => null) as CellValue[]
        } else {
          while (rows[r].length < colCount) {
            rows[r].push(null)
          }
        }
      }
    }

    // ── Resolve hyperlinks ──
    // Build a map of rId → target URL from worksheet relationships
    const relMap = new Map<string, string>()
    if (ctx.worksheetRels) {
      for (const rel of ctx.worksheetRels) {
        relMap.set(rel.id, rel.target)
      }
    }

    for (const hl of rawHyperlinks) {
      const pos = parseCellRef(hl.ref)
      const key = `${pos.row},${pos.col}`

      // Get or create cell in the cells map
      let cell = cells.get(key)
      if (!cell) {
        cell = {
          value: (rows[pos.row] && rows[pos.row][pos.col]) ?? null,
          type: "string",
        }
        // Far fewer hyperlinks than cells in any real file, but this is
        // the other place `cells` grows and the check is one comparison.
        assertCellMapCapacity(cells, key, name)
        cells.set(key, cell)
      }

      const hyperlink: Hyperlink = { target: "" }

      if (hl.location) {
        // Internal hyperlink
        hyperlink.location = hl.location
        hyperlink.target = hl.location
      } else if (hl.rId) {
        // External hyperlink — resolve from relationships
        const target = relMap.get(hl.rId)
        if (target) {
          hyperlink.target = target
        } else {
          // The cell keeps a hyperlink with an empty target, which reads as
          // a link that goes nowhere rather than as a missing relationship.
          // See #474.
          ctx.onWarning?.({
            code: "unresolved-hyperlink",
            message:
              `Cell ${hl.ref} links through ${hl.rId}, which the sheet's ` +
              "relationships do not define. Read with an empty target.",
            sheet: ctx.sheetName,
            row: pos.row,
            col: pos.col,
          })
        }
      }

      if (hl.tooltip) hyperlink.tooltip = hl.tooltip
      if (hl.display) hyperlink.display = hl.display

      cell.hyperlink = hyperlink
    }

    const sheet: Sheet = {
      name,
      rows,
    }

    if (cells.size > 0) {
      sheet.cells = cells
    }
    // Attach column definitions (width, hidden, outlineLevel, collapsed)
    if (defaultRowHeight !== undefined) sheet.defaultRowHeight = defaultRowHeight
    if (defaultColWidth !== undefined) sheet.defaultColWidth = defaultColWidth

    if (columnDefs.some((c) => Object.keys(c).length > 0)) {
      sheet.columns = columnDefs
    }
    if (merges.length > 0) {
      sheet.merges = merges
    }
    if (dataValidations.length > 0) {
      sheet.dataValidations = dataValidations
    }
    if (conditionalRules.length > 0) {
      sheet.conditionalRules = conditionalRules
    }
    if (autoFilter) {
      sheet.autoFilter = autoFilter
    }
    if (freezePane) {
      sheet.freezePane = freezePane
    }
    if (splitPane) {
      sheet.splitPane = splitPane
    }
    if (sheetProtection) {
      sheet.protection = sheetProtection
    }

    // Attach sheet view settings
    if (sheetView && Object.keys(sheetView).length > 0) {
      sheet.view = sheetView
    }

    // Attach page setup (merge margins into pageSetup if present)
    if (pageSetup || pageMargins || fitToPageFlag) {
      const ps: PageSetup = pageSetup ?? {}
      if (pageMargins) {
        ps.margins = pageMargins
      }
      if (fitToPageFlag) {
        ps.fitToPage = true
      }
      sheet.pageSetup = ps
    }

    // Attach header/footer
    if (headerFooter && Object.keys(headerFooter).length > 0) {
      sheet.headerFooter = headerFooter
    }

    // Attach page breaks
    if (rowBreaks.length > 0) {
      sheet.rowBreaks = rowBreaks.sort((a, b) => a - b)
    }
    if (colBreaks.length > 0) {
      sheet.colBreaks = colBreaks.sort((a, b) => a - b)
    }

    // Attach row definitions (height, hidden, outlineLevel)
    if (rowDefs.size > 0) {
      sheet.rowDefs = rowDefs
    }

    // Attach sparklines
    if (sparklines.length > 0) {
      sheet.sparklines = sparklines
    }

    // Attach outline properties
    if (outlineProperties) {
      sheet.outlineProperties = outlineProperties
    }

    return sheet
  }

  return { handlers, finish }
}

/**
 * Parse a worksheet XML into a Sheet using SAX parsing for performance.
 * This avoids building a full DOM tree for large worksheets.
 */
export function parseWorksheet(xml: string, name: string, ctx: WorksheetContext): Sheet {
  const { handlers, finish } = worksheetParser(name, ctx)
  parseSax(xml, handlers)
  return finish()
}

/**
 * Parse a worksheet from a stream of its bytes, for a part that cannot
 * become a string at all.
 *
 * V8 stops at 0x1fffffe8 characters, so `xl/worksheets/sheet2.xml` at
 * 589 MB is unrepresentable however much memory the machine has — the
 * buffered reader could decompress it and still not parse it. See #503.
 *
 * The result is the same `Sheet` {@link parseWorksheet} builds, because
 * it is the same handlers and the same finalisation; only the driver
 * differs. `parseSaxStream` holds back a tag or an entity split across a
 * chunk boundary, and every text handler accumulates with `+=`, so a run
 * arriving in pieces is assembled exactly as one arriving whole.
 */
export async function parseWorksheetStream(
  stream: ReadableStream<Uint8Array>,
  name: string,
  ctx: WorksheetContext,
): Promise<Sheet> {
  const { handlers, finish } = worksheetParser(name, ctx)
  // `strict` so a truncated part is an error here, as it is for the
  // buffered driver. See `endOfInput` in the parser for why it is not
  // the default.
  await parseSaxStream(stream, handlers, { strict: true })
  return finish()
}

// ── Sheet Protection Parser ─────────────────────────────────────────

/**
 * Parse `<sheetProtection>` attributes into a SheetProtection object.
 *
 * XLSX attribute semantics:
 * - `sheet="1"` → sheet IS protected
 * - `objects="1"` → objects ARE protected
 * - `scenarios="1"` → scenarios ARE protected
 * - All other attrs: "1" = action is PROHIBITED → we convert to allow=false
 *   "0" = action is ALLOWED → we convert to allow=true
 */
function parseSheetProtectionAttrs(attrs: Record<string, string>): SheetProtection {
  const prot: SheetProtection = {}

  // password is stored as hex hash — we store it as-is (hashed form)
  // We do NOT store it as the `password` field since that's the raw plaintext in our API.
  // Instead we skip it — the hash is one-way and can't be reversed.
  // The presence of a password attr just means the sheet was password-protected.

  if (attrs["sheet"] === "1" || attrs["sheet"] === "true") {
    prot.sheet = true
  }
  if (attrs["objects"] === "1" || attrs["objects"] === "true") {
    prot.objects = true
  }
  if (attrs["scenarios"] === "1" || attrs["scenarios"] === "true") {
    prot.scenarios = true
  }

  // All other options: XLSX "1" = prohibited → our API allow = false
  const allowOptions: Array<[string, keyof SheetProtection]> = [
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
  ]

  for (const [attr, prop] of allowOptions) {
    const val = attrs[attr]
    if (val !== undefined) {
      // "1" or "true" = prohibited → allow = false
      // "0" or "false" = allowed → allow = true
      ;(prot as Record<string, boolean>)[prop] = !(val === "1" || val === "true")
    }
  }

  return prot
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
])
const VALID_OPERATORS = new Set<string>([
  "between",
  "notBetween",
  "equal",
  "notEqual",
  "greaterThan",
  "lessThan",
  "greaterThanOrEqual",
  "lessThanOrEqual",
])

function buildDataValidation(
  attrs: Record<string, string>,
  formula1Text: string,
  formula2Text: string,
): DataValidation | null {
  const typeStr = attrs["type"]
  if (!typeStr || !VALID_TYPES.has(typeStr)) return null

  const sqref = attrs["sqref"]
  if (!sqref) return null

  const dv: DataValidation = {
    type: typeStr as ValidationType,
    range: sqref,
  }

  // Operator
  const operatorStr = attrs["operator"]
  if (operatorStr && VALID_OPERATORS.has(operatorStr)) {
    dv.operator = operatorStr as ValidationOperator
  }

  // Boolean flags (XLSX uses "1" for true)
  if (attrs["allowBlank"] === "1" || attrs["allowBlank"] === "true") {
    dv.allowBlank = true
  }
  if (attrs["showInputMessage"] === "1" || attrs["showInputMessage"] === "true") {
    dv.showInputMessage = true
  }
  if (attrs["showErrorMessage"] === "1" || attrs["showErrorMessage"] === "true") {
    dv.showErrorMessage = true
  }

  // Error style
  const errorStyle = attrs["errorStyle"]
  if (errorStyle === "stop" || errorStyle === "warning" || errorStyle === "information") {
    dv.errorStyle = errorStyle
  }

  // Input/error messages (XLSX uses promptTitle/prompt for input messages)
  if (attrs["promptTitle"]) dv.inputTitle = attrs["promptTitle"]
  if (attrs["prompt"]) dv.inputMessage = attrs["prompt"]
  if (attrs["errorTitle"]) dv.errorTitle = attrs["errorTitle"]
  if (attrs["error"]) dv.errorMessage = attrs["error"]

  // Formulas
  if (formula1Text) {
    if (typeStr === "list") {
      // Check if formula1 is a quoted comma-separated list: "val1,val2,val3"
      const trimmed = formula1Text.trim()
      if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
        // Quoted list — parse into values array
        const inner = trimmed.slice(1, -1)
        dv.values = inner.split(",")
      } else {
        // Formula reference (e.g. Sheet2!$A$1:$A$10)
        dv.formula1 = formula1Text
      }
    } else {
      dv.formula1 = formula1Text
    }
  }

  if (formula2Text) {
    dv.formula2 = formula2Text
  }

  return dv
}

// ── Conditional Rule Builder ─────────────────────────────────────────

const VALID_CF_TYPES = new Set<string>([
  "cellIs",
  "expression",
  "colorScale",
  "dataBar",
  "iconSet",
  "top10",
  "aboveAverage",
  "duplicateValues",
  "uniqueValues",
  "containsText",
  "notContainsText",
  "beginsWith",
  "endsWith",
  "containsBlanks",
  "notContainsBlanks",
])

function buildConditionalRule(
  attrs: Record<string, string>,
  sqref: string,
  formulas: string[],
  csCfvos: Array<{ type: string; value?: string }>,
  csColors: Color[],
  dbCfvos: Array<{ type: string; value?: string }>,
  dbColor: Color | undefined,
  isCfvos: Array<{ type: string; value?: string }>,
  isAttrsObj: Record<string, string>,
  dxfs: CellStyle[] | undefined,
  ctx?: WorksheetContext,
): ConditionalRule | null {
  const typeStr = attrs["type"]
  if (!typeStr || !VALID_CF_TYPES.has(typeStr)) return null
  if (!sqref) return null

  const rule: ConditionalRule = {
    type: typeStr as ConditionalRuleType,
    priority: Number(attrs["priority"] ?? "1"),
    range: sqref,
  }

  // Operator
  const operatorStr = attrs["operator"]
  if (operatorStr && VALID_OPERATORS.has(operatorStr)) {
    rule.operator = operatorStr as ValidationOperator
  }

  // dxfId indexes the workbook's <dxfs> block, so the rule's formatting
  // only exists once styles.xml has been parsed. A file can legitimately
  // reference a dxfId we have no entry for (styles.xml missing, or the
  // index out of range); leave `style` absent rather than invent one.
  const dxfId = Number(attrs["dxfId"])
  if (dxfs && !Number.isNaN(dxfId)) {
    const dxf = dxfs[dxfId]
    // An empty <dxf/> carries no formatting — surfacing `{}` would claim
    // a style the rule does not have. Shared with any other rule pointing
    // at the same dxfId, following the same contract as a resolved cell
    // style; see resolveStyle in ./styles.ts.
    if (dxf && Object.keys(dxf).length > 0) rule.style = dxf
    // A rule whose formatting silently vanished still applies — it just
    // paints nothing, which looks like the rule not working rather than
    // like a damaged file. See #474.
    else if (!dxf) {
      ctx?.onWarning?.({
        code: "unresolved-dxf",
        message:
          `Conditional rule on ${sqref} asks for differential format ${dxfId}, ` +
          `which the file does not have (${dxfs.length} present). The rule keeps ` +
          "its condition and loses its formatting.",
        sheet: ctx.sheetName,
      })
    }
  }

  // stopIfTrue
  if (attrs["stopIfTrue"] === "1" || attrs["stopIfTrue"] === "true") {
    rule.stopIfTrue = true
  }

  // text attribute (for containsText, beginsWith, endsWith, etc.)
  if (attrs["text"] !== undefined) {
    rule.text = attrs["text"]
  }

  // Formulas
  if (formulas.length === 1) {
    rule.formula = formulas[0]
  } else if (formulas.length > 1) {
    rule.formula = formulas
  }

  // colorScale
  if (typeStr === "colorScale" && csCfvos.length > 0) {
    rule.colorScale = {
      cfvo: csCfvos.map((c) => ({
        type: c.type as "min" | "max" | "num" | "percent" | "percentile",
        value: c.value,
      })),
      colors: csColors,
    }
  }

  // dataBar
  if (typeStr === "dataBar" && dbCfvos.length > 0) {
    rule.dataBar = {
      cfvo: dbCfvos.map((c) => ({
        type: c.type as "min" | "max" | "num" | "percent" | "percentile",
        value: c.value,
      })),
      color: dbColor ?? {},
    }
  }

  // iconSet
  if (typeStr === "iconSet" && isCfvos.length > 0) {
    rule.iconSet = {
      iconSet: isAttrsObj["iconSet"] ?? "3TrafficLights1",
      cfvo: isCfvos.map((c) => ({
        type: c.type as "min" | "num" | "percent" | "percentile",
        value: c.value,
      })),
    }
    if (isAttrsObj["reverse"] === "1" || isAttrsObj["reverse"] === "true") {
      rule.iconSet.reverse = true
    }
    if (isAttrsObj["showValue"] === "0" || isAttrsObj["showValue"] === "false") {
      rule.iconSet.showValue = false
    }
  }

  return rule
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
  formulaType?: string,
  formulaSi?: number,
  formulaRef?: string,
  formulaCm?: boolean,
  fallbackRow?: number,
  fallbackCol?: number,
): void {
  // When the `r` attribute is missing, fall back to implicit row/col position
  // (parity with the streaming reader).
  const pos =
    ref !== ""
      ? parseCellRef(ref)
      : fallbackRow !== undefined && fallbackCol !== undefined
        ? { row: fallbackRow, col: fallbackCol }
        : null
  if (!pos) return
  const { row, col } = pos

  // Two different failures, treated differently on purpose — the same
  // distinction `clampColumnBound` draws a few lines down.
  //
  // A reference *past the grid* (`AAAAAA1`, row 2,000,000) is a resource
  // claim: honouring it would allocate billions of null slots and OOM
  // the process, and there is no partial answer that is not a fabricated
  // one. That throws, and always has.
  //
  // A reference that is *not a reference* (`B` with no row, `A0`) claims
  // nothing. It used to throw too, so one malformed `r` attribute cost
  // the whole sheet where every other content damage costs one cell.
  // It now drops the cell and says so. See #473.
  if (row > MAX_ROW_INDEX || col > MAX_COL_INDEX) {
    throw new ParseError(
      `Cell reference "${ref}" is outside the supported sheet bounds (max row ${
        MAX_ROW_INDEX + 1
      }, max col ${MAX_COL_INDEX + 1})`,
    )
  }
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0) {
    ctx.onWarning?.({
      code: "malformed-cell-ref",
      message:
        `Cell reference "${ref}" is not a cell reference — a column needs a ` +
        "row number and rows are 1-based. The cell is dropped; the rest of " +
        "the sheet is read.",
      sheet: ctx.sheetName,
    })
    return
  }

  // Ensure row array exists. Skipped in sparse mode — allocating the row
  // out to the cell's column is the cost being avoided, and it is paid
  // here rather than in the densify pass at the end. See #501.
  if (!ctx.sparse) {
    while (rows.length <= row) {
      rows.push([])
    }
    while (rows[row].length <= col) {
      rows[row].push(null)
    }
  }

  let value: CellValue = null
  let cellType: Cell["type"] = "empty"
  let formula: string | undefined
  let formulaResult: CellValue | undefined
  let richText: RichTextRun[] | undefined

  // Handle formula (including shared formula slave cells with no text)
  if (formulaText) {
    formula = formulaText
  } else if (formulaType === "shared" && formulaSi !== undefined && formulaSi >= 0) {
    // Shared formula slave cell: no formula text, but has si attribute
    formula = ""
  }

  // Determine cell value based on type
  switch (type) {
    case "s": {
      // Shared string
      const idx = Number(valueText)
      if (!Number.isNaN(idx) && idx >= 0 && idx < ctx.sharedStrings.length) {
        const ss = ctx.sharedStrings[idx]
        value = ss.text
        if (ss.richText && ss.richText.length > 0) {
          richText = ss.richText
          cellType = "richText"
        } else {
          cellType = "string"
        }
      } else {
        // Out-of-bounds SST index — return null (consistent with the
        // streaming reader), not the raw index string. Reported, because
        // `null` here is otherwise indistinguishable from an empty cell.
        ctx.onWarning?.({
          code: "unresolved-shared-string",
          message:
            `Cell ${ref || `${row},${col}`} points at shared string ${valueText}, ` +
            `which the file does not have (${ctx.sharedStrings.length} present). Read as empty.`,
          sheet: ctx.sheetName,
          row,
          col,
        })
        value = null
        cellType = "empty"
      }
      break
    }
    case "str": {
      // Inline formula string result. `formulaResult` used to be set in
      // the numeric arm alone, so a cached result survived only when it
      // happened to be a number — and `readXlsx` → `writeXlsx` dropped
      // the rest, emitting `<f>` with no `<v>`. The writer has always
      // been able to write them back. See #497.
      value = decodeOoxmlEscapes(valueText)
      cellType = formula ? "formula" : "string"
      if (formula) formulaResult = value
      break
    }
    case "inlineStr": {
      // Inline string with <is> element
      if (inlineRichText && inlineRichText.length > 0) {
        value = inlineRichText.map((r) => r.text).join("")
        richText = inlineRichText
        cellType = "richText"
      } else {
        value = decodeOoxmlEscapes(inlineText)
        cellType = "string"
      }
      break
    }
    case "b": {
      // Boolean
      value = valueText === "1" || valueText.toLowerCase() === "true"
      cellType = formula ? "formula" : "boolean"
      if (formula) formulaResult = value
      break
    }
    case "e": {
      // Error.
      //
      // A cell carrying a formula reports `type: "formula"` here, as the
      // numeric and string arms do. It used to report `"error"` on the
      // way in and `"formula"` on the way back out, which cannot both be
      // right — and the round trip is the side with a second opinion.
      // `value` still holds the error token either way, so spotting an
      // error by its value is unaffected; a *hard-coded* error cell,
      // which carries no formula, still reports `"error"`. See #497.
      value = cellError(valueText)
      cellType = formula ? "formula" : "error"
      if (formula) formulaResult = value
      break
    }
    case "d": {
      // ISO 8601 date (ECMA-376 Part 1, §18.18.11 ST_CellType). Every
      // other member of that enumeration had a case; this one fell
      // through to `n`, where `Number("2024-03-17")` is NaN and the
      // value landed in the "shouldn't happen, but be safe" arm as a
      // *string*. It does happen: openpyxl writes it whenever
      // `iso_dates=True`. See #496.
      //
      // The value is an instant, not an offset from an epoch, so
      // `date1904` must NOT be applied — unlike the serial path below,
      // where the same day is 1,462 days apart between the two systems.
      const parsed = parseIsoCellDate(valueText)
      if (parsed) {
        value = parsed
        cellType = "date"
      } else if (valueText !== "") {
        // A bare time (`13:45:30`, which openpyxl emits for a
        // `datetime.time`) is not an ISO 8601 date-time and has no day
        // to anchor it. Left as text rather than guessed onto an epoch.
        value = valueText
        cellType = "string"
      } else {
        value = null
        cellType = "empty"
      }
      if (formula) {
        formulaResult = value
        cellType = "formula"
      }
      break
    }
    case "n":
    default: {
      // Number (explicit or implied)
      if (valueText === "" && !formula) {
        // Empty cell
        value = null
        cellType = "empty"
        break
      }

      const num = Number(valueText)
      if (!Number.isNaN(num) && valueText !== "") {
        // Check if this is a date via style
        if (ctx.styles && styleIndex >= 0 && isDateStyle(ctx.styles, styleIndex)) {
          value = serialToDate(num, ctx.dateSystem)
          cellType = "date"
        } else {
          value = num
          cellType = "number"
        }
      } else if (valueText !== "") {
        // Non-numeric value text (shouldn't happen, but be safe)
        value = valueText
        cellType = "string"
      }

      if (formula) {
        formulaResult = value
        cellType = "formula"
      }
      break
    }
  }

  // Set the value in the rows array. In sparse mode there is no grid to
  // set it in: the whole point is that the bounding box is not paid for.
  // See #501.
  if (!ctx.sparse) rows[row][col] = value

  // Detect Excel 2024 checkbox feature on this cell's xf — independent of
  // readStyles so the flag round-trips even without full style hydration.
  const isCheckbox =
    ctx.styles && styleIndex >= 0
      ? (ctx.styles.cellXfs[styleIndex]?.hasCheckboxFeature ?? false)
      : false

  // Build Cell object if there's detail beyond the raw value — or always,
  // in sparse mode, where `cells` is the only place a value can live.
  const hasDetails =
    ctx.sparse ||
    formula !== undefined ||
    richText !== undefined ||
    (ctx.readStyles && ctx.styles && styleIndex >= 0) ||
    isCheckbox ||
    cellType === "error" ||
    cellType === "formula" ||
    cellType === "richText"

  if (hasDetails) {
    const cell: Cell = {
      value,
      type: cellType,
    }
    if (isCheckbox) {
      cell.checkbox = true
    }
    if (formula !== undefined) {
      cell.formula = formula
      if (formulaResult !== undefined) {
        cell.formulaResult = formulaResult
      }
      // Store formula type metadata
      if (formulaType === "shared") {
        cell.formulaType = "shared"
        if (formulaSi !== undefined && formulaSi >= 0) {
          cell.formulaSharedIndex = formulaSi
        }
        if (formulaRef) {
          cell.formulaRef = formulaRef
        }
      } else if (formulaType === "array") {
        cell.formulaType = "array"
        if (formulaRef) {
          cell.formulaRef = formulaRef
        }
      }
      // The dynamic-array flag is independent of the formula type — the
      // reader used to surface it only for `t="array"`, mirroring the
      // writer's matching restriction (#407).
      if (formulaCm) {
        cell.formulaDynamic = true
      }
    }
    if (richText) {
      cell.richText = richText
    }
    if (ctx.readStyles && ctx.styles && styleIndex >= 0) {
      const style = resolveStyle(ctx.styles, styleIndex)
      if (Object.keys(style).length > 0) {
        cell.style = style
      } else if (styleIndex >= ctx.styles.cellXfs.length) {
        // The xf the cell names is not in the file, so the cell comes back
        // unstyled — indistinguishable from one that never had a format.
        ctx.onWarning?.({
          code: "unresolved-style",
          message:
            `Cell ${ref || `${row},${col}`} points at cell format ${styleIndex}, ` +
            `which the file does not have (${ctx.styles.cellXfs.length} present). Read unstyled.`,
          sheet: ctx.sheetName,
          row,
          col,
        })
      }
    }
    const cellKey = `${row},${col}`
    assertCellMapCapacity(cells, cellKey, ctx.sheetName)
    cells.set(cellKey, cell)
  }
}

// ── Inline Rich Text Font Properties ─────────────────────────────────

function applyFontProp(font: FontStyle, tag: string, attrs: Record<string, string>): void {
  switch (tag) {
    case "b":
      font.bold = attrs["val"] !== "0" && attrs["val"] !== "false"
      break
    case "i":
      font.italic = attrs["val"] !== "0" && attrs["val"] !== "false"
      break
    case "u": {
      const val = attrs["val"]
      if (val === "double") font.underline = "double"
      else font.underline = true
      break
    }
    case "strike":
      font.strikethrough = attrs["val"] !== "0" && attrs["val"] !== "false"
      break
    case "sz":
      if (attrs["val"]) font.size = Number(attrs["val"])
      break
    case "rFont":
      if (attrs["val"]) font.name = attrs["val"]
      break
    case "color":
      font.color = parseColorAttrs(attrs)
      break
    case "vertAlign":
      if (attrs["val"] === "superscript" || attrs["val"] === "subscript") {
        font.vertAlign = attrs["val"]
      }
      break
    case "family":
      if (attrs["val"]) font.family = Number(attrs["val"])
      break
    case "charset":
      if (attrs["val"]) font.charset = Number(attrs["val"])
      break
    case "scheme":
      if (attrs["val"] === "major" || attrs["val"] === "minor" || attrs["val"] === "none") {
        font.scheme = attrs["val"]
      }
      break
  }
}

// ── Page Margins Parser ────────────────────────────────────────────────

function parsePageMarginsAttrs(attrs: Record<string, string>): PageMargins {
  const m: PageMargins = {}
  if (attrs["left"]) m.left = Number(attrs["left"])
  if (attrs["right"]) m.right = Number(attrs["right"])
  if (attrs["top"]) m.top = Number(attrs["top"])
  if (attrs["bottom"]) m.bottom = Number(attrs["bottom"])
  if (attrs["header"]) m.header = Number(attrs["header"])
  if (attrs["footer"]) m.footer = Number(attrs["footer"])
  return m
}

// ── Page Setup Parser ──────────────────────────────────────────────────

/** Reverse map: XLSX paper size number → PaperSize string */
// The name↔code table lives with the writer, so the two cannot disagree —
// there used to be a second copy here, and keeping two tables in step is
// exactly the kind of thing nobody checks. See #439 §Q.

/**
 * Merge `<printOptions>` attributes into the sheet's page setup.
 *
 * The element can appear either side of `<pageSetup>` in a worksheet, so
 * this merges into whatever exists rather than replacing it — and creates
 * the object when `<printOptions>` comes first.
 */
function applyPrintOptionsAttrs(
  existing: PageSetup | undefined,
  attrs: Record<string, string>,
): PageSetup {
  const ps: PageSetup = existing ?? {}
  const isTrue = (value: string | undefined): boolean => value === "1" || value === "true"

  if (isTrue(attrs["gridLines"])) ps.showGridLines = true
  if (isTrue(attrs["headings"])) ps.showRowColHeaders = true
  if (isTrue(attrs["horizontalCentered"])) ps.horizontalCentered = true
  if (isTrue(attrs["verticalCentered"])) ps.verticalCentered = true

  return ps
}

function parsePageSetupAttrs(attrs: Record<string, string>, ctx?: WorksheetContext): PageSetup {
  const ps: PageSetup = {}

  if (attrs["paperSize"]) {
    const num = Number(attrs["paperSize"])
    // A code with no name round-trips as the number rather than vanishing.
    if (Number.isInteger(num) && num > 0) ps.paperSize = PAPER_SIZE_REVERSE[num] ?? num
    else {
      // Anything else is not a paper size, so the sheet comes back
      // claiming the default one. Reported, because the printed output
      // then differs from the file and nothing else says why. See #474.
      ctx?.onWarning?.({
        code: "unusable-paper-size",
        message:
          `Page setup names paper size "${attrs["paperSize"]}", which is not a ` +
          "positive integer code. Dropped; the sheet reads with no paper size set.",
        sheet: ctx.sheetName,
      })
    }
  }

  if (attrs["orientation"] === "landscape" || attrs["orientation"] === "portrait") {
    ps.orientation = attrs["orientation"]
  }

  if (attrs["scale"]) {
    ps.scale = Number(attrs["scale"])
  }

  if (attrs["fitToWidth"] !== undefined || attrs["fitToHeight"] !== undefined) {
    ps.fitToPage = true
    if (attrs["fitToWidth"]) ps.fitToWidth = Number(attrs["fitToWidth"])
    if (attrs["fitToHeight"]) ps.fitToHeight = Number(attrs["fitToHeight"])
  }

  // Excel writes the centering flags on <printOptions>; hucre used to
  // write them here. Keep accepting them from <pageSetup> so files from
  // older versions still round-trip. See #360.
  if (attrs["horizontalCentered"] === "1" || attrs["horizontalCentered"] === "true") {
    ps.horizontalCentered = true
  }

  if (attrs["verticalCentered"] === "1" || attrs["verticalCentered"] === "true") {
    ps.verticalCentered = true
  }

  // ── The rest of CT_PageSetup (#470) ──────────────────────────────
  // Read unconditionally, defaults included: this is the roundtrip path
  // as well as the read path, and dropping an attribute because it
  // happened to equal its default would rewrite a file the caller only
  // opened. The writer is the side that elides defaults.
  if (attrs["paperWidth"]) ps.paperWidth = attrs["paperWidth"]
  if (attrs["paperHeight"]) ps.paperHeight = attrs["paperHeight"]

  const firstPageNumber = intAttr(attrs["firstPageNumber"])
  if (firstPageNumber !== undefined) ps.firstPageNumber = firstPageNumber
  if (attrs["useFirstPageNumber"] !== undefined) {
    ps.useFirstPageNumber = isTruthyAttr(attrs["useFirstPageNumber"])
  }

  if (attrs["pageOrder"] === "overThenDown" || attrs["pageOrder"] === "downThenOver") {
    ps.pageOrder = attrs["pageOrder"]
  }
  if (isTruthyAttr(attrs["blackAndWhite"])) ps.blackAndWhite = true
  if (isTruthyAttr(attrs["draft"])) ps.draft = true

  const comments = attrs["cellComments"]
  if (comments === "none" || comments === "asDisplayed" || comments === "atEnd") {
    ps.cellComments = comments
  }

  const errors = attrs["errors"]
  if (errors === "displayed" || errors === "blank" || errors === "dash" || errors === "NA") {
    ps.errors = errors
  }

  const copies = intAttr(attrs["copies"])
  if (copies !== undefined) ps.copies = copies
  const hDpi = intAttr(attrs["horizontalDpi"])
  if (hDpi !== undefined) ps.horizontalDpi = hDpi
  const vDpi = intAttr(attrs["verticalDpi"])
  if (vDpi !== undefined) ps.verticalDpi = vDpi
  if (attrs["usePrinterDefaults"] !== undefined) {
    ps.usePrinterDefaults = isTruthyAttr(attrs["usePrinterDefaults"])
  }

  return ps
}

/** `"1"` / `"true"` — the two spellings ECMA-376 allows for xsd:boolean. */
function isTruthyAttr(value: string | undefined): boolean {
  return value === "1" || value === "true"
}

/**
 * A non-negative integer attribute, or `undefined` when it is absent or
 * not one. A hostile `copies="NaN"` becoming `NaN` on the model would
 * serialize back out as the literal string `NaN`.
 */
function intAttr(value: string | undefined): number | undefined {
  if (value === undefined) return undefined
  const n = Number(value)
  return Number.isInteger(n) && n >= 0 ? n : undefined
}

// ── Color Attribute Parser ──────────────────────────────────────────────

/**
 * Parse a colour element's attributes — `<tabColor>`, a font or fill
 * `<color>`, a conditional-format scale stop, a sparkline series. One
 * reader for all of them: the CF and sparkline sites used to read `rgb`
 * alone and lose theme colours.
 */
function parseColorAttrs(attrs: Record<string, string>): Color {
  const color: Color = {}
  if (attrs["rgb"]) {
    const rgb = attrs["rgb"]
    // Strip ARGB alpha prefix if present (8 chars → 6 chars)
    color.rgb = rgb.length === 8 ? rgb.slice(2) : rgb
  }
  if (attrs["theme"]) {
    color.theme = Number(attrs["theme"])
  }
  if (attrs["tint"]) {
    color.tint = Number(attrs["tint"])
  }
  if (attrs["indexed"]) {
    color.indexed = Number(attrs["indexed"])
  }
  return color
}

/**
 * Parse the value of a `t="d"` cell — an ISO 8601 date or date-time.
 *
 * Deliberately strict. `new Date(text)` accepts a great deal that is not
 * ISO 8601 and answers `Invalid Date` for the rest, so a loose parse here
 * would turn arbitrary cell text into dates. The shapes accepted are the
 * ones ECMA-376 §18.18.11 describes and that producers actually write:
 * `YYYY-MM-DD`, optionally with a time, optionally with a zone.
 *
 * An unqualified time is read as UTC, for the same reason the ODS and
 * docProps readers do it: every format hucre reads records an absolute
 * moment, and local time would make the same file mean different things
 * on different machines. See #415, #474.
 *
 * Exported for `stream-reader.ts`, which needs the same answer. #496
 * added the `t="d"` case here and not there, so the streaming reader
 * returned the raw text where this returned a `Date` — one fix, two
 * implementations, and only one of them got it. Sharing the parser is
 * what stops the two drifting on the next shape someone accepts.
 */
export function parseIsoCellDate(text: string): Date | undefined {
  const trimmed = text.trim()
  if (
    !/^\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?$/.test(trimmed)
  ) {
    return undefined
  }
  const zoned = /(?:Z|[+-]\d{2}:?\d{2})$/.test(trimmed)
  const hasTime = /[T ]\d{2}:/.test(trimmed)
  const normalized = trimmed.replace(" ", "T")
  const date = new Date(hasTime && !zoned ? `${normalized}Z` : normalized)
  return Number.isNaN(date.getTime()) ? undefined : date
}

// ── Cell Value Types ────────────────────────────────────────────────

import type { CellError } from "./cell-error"

export type CellValue = string | number | boolean | Date | CellError | null

export type CellType =
  | "string"
  | "number"
  | "boolean"
  | "date"
  | "error"
  | "formula"
  | "richText"
  | "empty"

// ── Color ──────────────────────────────────────────────────────────

export interface Color {
  /** Hex RGB string without '#', e.g. "FF0000" */
  rgb?: string
  /** Theme color index */
  theme?: number
  /** Tint applied to theme color (-1.0 to 1.0) */
  tint?: number
  /** Indexed color (legacy) */
  indexed?: number
}

// ── Font ───────────────────────────────────────────────────────────

export interface FontStyle {
  name?: string
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean | "single" | "double" | "singleAccounting" | "doubleAccounting"
  strikethrough?: boolean
  color?: Color
  vertAlign?: "superscript" | "subscript"
  family?: number
  charset?: number
  scheme?: "major" | "minor" | "none"
}

// ── Fill ───────────────────────────────────────────────────────────

export type FillPattern =
  | "none"
  | "solid"
  | "darkDown"
  | "darkGray"
  | "darkGrid"
  | "darkHorizontal"
  | "darkTrellis"
  | "darkUp"
  | "darkVertical"
  | "gray0625"
  | "gray125"
  | "lightDown"
  | "lightGray"
  | "lightGrid"
  | "lightHorizontal"
  | "lightTrellis"
  | "lightUp"
  | "lightVertical"
  | "mediumGray"

export interface PatternFill {
  type: "pattern"
  pattern: FillPattern
  fgColor?: Color
  bgColor?: Color
}

export interface GradientFill {
  type: "gradient"
  degree?: number
  stops: Array<{ position: number; color: Color }>
}

export type FillStyle = PatternFill | GradientFill

// ── Border ─────────────────────────────────────────────────────────

export type BorderLineStyle =
  | "thin"
  | "medium"
  | "thick"
  | "dotted"
  | "dashed"
  | "double"
  | "hair"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot"

export interface BorderSide {
  style: BorderLineStyle
  color?: Color
}

export interface BorderStyle {
  top?: BorderSide
  right?: BorderSide
  bottom?: BorderSide
  left?: BorderSide
  diagonal?: BorderSide
  diagonalUp?: boolean
  diagonalDown?: boolean
}

// ── Alignment ──────────────────────────────────────────────────────

export interface AlignmentStyle {
  horizontal?:
    | "left"
    | "center"
    | "right"
    | "fill"
    | "justify"
    | "centerContinuous"
    | "distributed"
    | "general"
  vertical?: "top" | "center" | "bottom" | "justify" | "distributed"
  wrapText?: boolean
  shrinkToFit?: boolean
  textRotation?: number
  indent?: number
  readingOrder?: "ltr" | "rtl" | "context"
}

// ── Cell Style ─────────────────────────────────────────────────────

export interface CellStyle {
  font?: FontStyle
  fill?: FillStyle
  border?: BorderStyle
  alignment?: AlignmentStyle
  numFmt?: string
  protection?: CellProtection
}

export interface CellProtection {
  locked?: boolean
  hidden?: boolean
}

// ── Rich Text ──────────────────────────────────────────────────────

export interface RichTextRun {
  text: string
  font?: FontStyle
}

// ── Hyperlink ──────────────────────────────────────────────────────

export interface Hyperlink {
  target: string
  tooltip?: string
  display?: string
  /** Internal reference (e.g. "Sheet2!A1") */
  location?: string
}

/**
 * A rich hyperlink value that can be placed inline in a {@link WriteSheet.data}
 * row object, keyed by a column's `key`. The display text and link target live
 * together, so a "Link" column needs no parallel `cells` coordinate map.
 *
 * @example
 * ```ts
 * data: [{ id: "abc", link: { text: "Open", hyperlink: "https://example.com/abc" } }]
 * ```
 */
export interface HyperlinkValue {
  /** Display text shown in the cell. */
  text: string
  /** Link destination — an external URL, or an internal ref prefixed with `#` (e.g. `"#Sheet2!A1"`). */
  hyperlink: string
  /** Optional hover tooltip. */
  tooltip?: string
}

// ── Comment ────────────────────────────────────────────────────────

export interface CellComment {
  author?: string
  text: string
  richText?: RichTextRun[]
}

// ── Cell ───────────────────────────────────────────────────────────

export interface Cell {
  value: CellValue
  type: CellType
  /**
   * The cell's format.
   *
   * On a cell that came from a reader, the nested `font` / `fill` /
   * `border` objects are **shared** with every other cell of the same
   * format — see {@link ReadOptions.readStyles}. Copy with
   * `cloneCellStyle` before mutating one cell's format in place.
   */
  style?: CellStyle
  /**
   * Render this cell as an Excel 2024 native checkbox. Only meaningful for
   * boolean cells; the value drives the checked state.
   *
   * Implemented via Microsoft's FeaturePropertyBag extension to OOXML
   * (the `{C7286773-470A-42A8-94C5-96B5CB345126}` cell-XF complement).
   * Requires Microsoft 365; older Excel and LibreOffice fall back to the
   * raw `TRUE`/`FALSE` value.
   */
  checkbox?: boolean
  formula?: string
  formulaResult?: CellValue
  /** Formula type: "shared" | "array". Undefined means normal formula. */
  formulaType?: "shared" | "array"
  /** Shared formula index (si attribute) */
  formulaSharedIndex?: number
  /** Range this formula applies to (ref attribute on master cell) */
  formulaRef?: string
  /**
   * Dynamic array flag (`cm="1"`). Independent of {@link formulaType} —
   * a spilling function set as a plain formula carries it just as an
   * explicit `"array"` formula does.
   */
  formulaDynamic?: boolean
  richText?: RichTextRun[]
  hyperlink?: Hyperlink
  comment?: CellComment
}

/**
 * What a writer accepts where a cell goes: a bare value, or a cell
 * object — `{ value, style }`, `{ formula }`, anything a {@link Cell}
 * carries. One type for every writer; v1 had four spellings of it
 * (`Partial<Cell>`, `StreamStyledCell`, `OdsStyledCell`, `OdsWriteCell`).
 */
export type CellInput = CellValue | Partial<Cell>

// ── Column Definition ──────────────────────────────────────────────

export interface ColumnDef {
  /** Column header text */
  header?: string
  /** Key for object-based data */
  key?: string
  /** Column width in characters */
  width?: number
  /** Auto-calculate optimal width from cell content */
  autoWidth?: boolean
  /**
   * Default style for every cell in the column. Applies whether the rows
   * come from {@link WriteSheet.data} or {@link WriteSheet.rows} — on the
   * `data[]` path the generated header row gets it too.
   */
  style?: CellStyle
  /** Number format. Folded into {@link style}; an explicit `style.numFmt` wins. */
  numFmt?: string
  /** Hide column */
  hidden?: boolean
  /** Outline level (grouping) */
  outlineLevel?: number
  /** Whether this outline group is collapsed */
  collapsed?: boolean
}

// ── Merge Range ────────────────────────────────────────────────────

export interface MergeRange {
  /** Start row (0-based) */
  startRow: number
  /** Start column (0-based) */
  startCol: number
  /** End row (0-based, inclusive) */
  endRow: number
  /** End column (0-based, inclusive) */
  endCol: number
}

// ── Data Validation ────────────────────────────────────────────────

export type ValidationType =
  | "list"
  | "whole"
  | "decimal"
  | "date"
  | "time"
  | "textLength"
  | "custom"

export type ValidationOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "lessThan"
  | "greaterThanOrEqual"
  | "lessThanOrEqual"

export interface DataValidation {
  type: ValidationType
  operator?: ValidationOperator
  formula1?: string
  formula2?: string
  /** List values (for type: "list") */
  values?: string[]
  allowBlank?: boolean
  showInputMessage?: boolean
  showErrorMessage?: boolean
  inputTitle?: string
  inputMessage?: string
  errorTitle?: string
  errorMessage?: string
  errorStyle?: "stop" | "warning" | "information"
  /** Cell range (e.g. "A1:A100") */
  range: string
}

// ── Conditional Formatting ─────────────────────────────────────────

export type ConditionalRuleType =
  | "cellIs"
  | "expression"
  | "colorScale"
  | "dataBar"
  | "iconSet"
  | "top10"
  | "aboveAverage"
  | "duplicateValues"
  | "uniqueValues"
  | "containsText"
  | "notContainsText"
  | "beginsWith"
  | "endsWith"
  | "containsBlanks"
  | "notContainsBlanks"

export interface ConditionalRule {
  type: ConditionalRuleType
  priority: number
  operator?: ValidationOperator
  formula?: string | string[]
  style?: CellStyle
  stopIfTrue?: boolean
  range: string
  /**
   * Color scale configuration. Colours are {@link Color}, the same shape
   * fonts and fills use — a scale built from theme colours used to read
   * back as `""` and be written back as `rgb=""`, because the field could
   * only hold an RGB string.
   */
  colorScale?: {
    cfvo: Array<{
      type: "min" | "max" | "num" | "percent" | "percentile"
      value?: string
    }>
    colors: Color[]
  }
  /** Data bar configuration */
  dataBar?: {
    cfvo: Array<{
      type: "min" | "max" | "num" | "percent" | "percentile"
      value?: string
    }>
    color: Color
  }
  /** Icon set configuration */
  iconSet?: {
    iconSet: string // "3Arrows", "3TrafficLights1", etc.
    cfvo: Array<{
      type: "min" | "num" | "percent" | "percentile"
      value?: string
    }>
    reverse?: boolean
    showValue?: boolean
  }
  /** Text value for containsText, notContainsText, beginsWith, endsWith */
  text?: string
}

// ── Auto Filter ────────────────────────────────────────────────────

export interface AutoFilter {
  /** Range (e.g. "A1:D100") */
  range: string
  /** Column filter criteria */
  columns?: Array<{
    /** 0-based column index within the autoFilter range */
    colIndex: number
    /** List of values to filter by */
    filters?: string[]
  }>
}

// ── Freeze Pane ────────────────────────────────────────────────────

export interface FreezePane {
  /** Number of rows to freeze from top */
  rows?: number
  /** Number of columns to freeze from left */
  columns?: number
}

// ── Split Pane ─────────────────────────────────────────────────────

export interface SplitPane {
  /** Horizontal split position in twips (1/20 of a point) */
  xSplit?: number
  /** Vertical split position in twips (1/20 of a point) */
  ySplit?: number
}

// ── Named Range ────────────────────────────────────────────────────

export interface NamedRange {
  name: string
  /** Cell range reference (e.g. "Sheet1!$A$1:$D$10") */
  range: string
  /** Scope: undefined = workbook level, string = sheet name */
  scope?: string
  comment?: string
}

// ── Page Setup / Print ─────────────────────────────────────────────

/**
 * A paper size, either by name or by its raw OOXML `paperSize` code.
 *
 * The names cover what people ask for; the number is the escape hatch.
 * Excel defines about 120 codes and hucre used to model nine, dropping
 * anything else **silently** on read — so a workbook set to A6 lost its
 * page size with no error and nothing in the parity statement. A code
 * with no name here now round-trips as the number it is. See #439 §Q.
 */
export type PaperSize = PaperSizeName | (number & {})

export type PaperSizeName =
  | "letter"
  | "letterSmall"
  | "tabloid"
  | "ledger"
  | "legal"
  | "statement"
  | "executive"
  | "a3"
  | "a4"
  | "a4Small"
  | "a5"
  | "b4"
  | "b5"
  | "folio"
  | "quarto"
  | "note"
  | "envelope9"
  | "envelope10"
  | "envelope11"
  | "envelope12"
  | "envelope14"
  | "cSheet"
  | "dSheet"
  | "eSheet"
  | "envelopeDL"
  | "envelopeC5"
  | "envelopeC3"
  | "envelopeC4"
  | "envelopeC6"
  | "envelopeC65"
  | "envelopeB4"
  | "envelopeB5"
  | "envelopeB6"
  | "envelopeItaly"
  | "envelopeMonarch"
  | "envelopePersonal"
  | "fanfoldUS"
  | "fanfoldGermanStd"
  | "fanfoldGermanLegal"
  | "a6"
  | "japanesePostcard"
  | "japaneseDoublePostcard"

export interface PageSetup {
  paperSize?: PaperSize
  orientation?: "portrait" | "landscape"
  fitToPage?: boolean
  fitToWidth?: number
  fitToHeight?: number
  scale?: number
  margins?: PageMargins
  /**
   * Print area as a bare A1 range (e.g. `"$A$1:$D$50"`), without a sheet
   * qualifier. Stored in the file as the reserved `_xlnm.Print_Area`
   * defined name; the reader folds that name back into this field rather
   * than surfacing it in {@link Workbook.namedRanges}, so the setting has
   * one representation in both directions.
   */
  printArea?: string
  /** Rows repeated at the top of every page (e.g. `"$1:$1"`). Stored in `_xlnm.Print_Titles`. */
  printTitlesRow?: string
  /** Columns repeated at the left of every page (e.g. `"$A:$A"`). Stored in `_xlnm.Print_Titles`. */
  printTitlesColumn?: string
  showGridLines?: boolean
  showRowColHeaders?: boolean
  horizontalCentered?: boolean
  verticalCentered?: boolean
  /**
   * Page number to print on the first page. Excel only honours it when
   * `useFirstPageNumber` is also set, and the writer sets that flag for
   * you whenever this is present — a `firstPageNumber` that silently did
   * nothing would be worse than not having the field.
   */
  firstPageNumber?: number
  /**
   * Whether {@link firstPageNumber} is used at all. Written implicitly
   * with `firstPageNumber`; carried here so a file that sets the flag
   * without the number, or vice versa, round-trips as it was.
   */
  useFirstPageNumber?: boolean
  /**
   * Order pages are laid out in when the sheet is wider and taller than
   * one page. Default: `"downThenOver"`.
   */
  pageOrder?: "downThenOver" | "overThenDown"
  /** Print without colour. */
  blackAndWhite?: boolean
  /** Print without graphics. */
  draft?: boolean
  /** Where cell comments are printed. Default: `"none"`. */
  cellComments?: "none" | "asDisplayed" | "atEnd"
  /** How cells holding errors are printed. Default: `"displayed"`. */
  errors?: "displayed" | "blank" | "dash" | "NA"
  /** Number of copies. Default: 1. */
  copies?: number
  /** Horizontal print resolution in DPI. Default: 600. */
  horizontalDpi?: number
  /** Vertical print resolution in DPI. Default: 600. */
  verticalDpi?: number
  /**
   * Custom page width as an ST_PositiveUniversalMeasure — a number with a
   * unit, e.g. `"210mm"`, `"8.5in"`, `"21cm"`. The only way to express a
   * page size that has no {@link PaperSize} code.
   *
   * Excel reads this in preference to `paperSize` when both are present.
   * Set it together with {@link paperHeight}; one alone describes nothing.
   */
  paperWidth?: string
  /** Custom page height; see {@link paperWidth}. */
  paperHeight?: string
  /**
   * Whether the printer's own defaults are used for the settings this
   * sheet does not name. Default: true.
   */
  usePrinterDefaults?: boolean
}

export interface PageMargins {
  top?: number
  right?: number
  bottom?: number
  left?: number
  header?: number
  footer?: number
}

export interface HeaderFooter {
  oddHeader?: string
  oddFooter?: string
  evenHeader?: string
  evenFooter?: string
  firstHeader?: string
  firstFooter?: string
  differentOddEven?: boolean
  differentFirst?: boolean
}

// ── Sparkline ─────────────────────────────────────────────────────

export interface Sparkline {
  /** Cell where the sparkline is displayed */
  location: string
  /** Data range (e.g. "Sheet1!B2:F2") */
  dataRange: string
  /** Type: line, column, or win/loss (stacked) */
  type?: "line" | "column" | "stacked"
  /** Series colour. Default: Excel's `376092`. */
  color?: Color
  /** Show markers */
  markers?: boolean
}

// ── TextBox ───────────────────────────────────────────────────────

export interface SheetTextBox {
  text: string
  anchor: {
    from: { row: number; col: number }
    to?: { row: number; col: number }
  }
  width?: number
  height?: number
  style?: {
    fontSize?: number
    bold?: boolean
    color?: string
    fillColor?: string
    borderColor?: string
  }
  /** Alternative text for screen readers (lands in xdr:cNvPr/@descr). */
  altText?: string
  /** Title/caption for the shape (lands in xdr:cNvPr/@title). */
  title?: string
}

// ── Threaded Comments (Excel 365+) ─────────────────────────────────

/**
 * A person who can author or be mentioned in threaded comments.
 * Stored in the workbook-wide `xl/persons/person.xml` part.
 */
export interface ThreadedCommentPerson {
  /** Stable GUID identifying this person within the workbook. */
  id: string
  /** Display name shown in Excel's comment pane (required by the schema). */
  displayName: string
  /** Identity-system user id, e.g. the Azure AD object id. */
  userId?: string
  /** Identity provider name, e.g. "AD" or "PeoplePicker". */
  providerId?: string
}

/**
 * An `@person` mention inside a threaded comment's text. Indices are
 * UTF-16 code-unit offsets into the comment text.
 */
export interface ThreadedCommentMention {
  mentionPersonId: string
  mentionId: string
  startIndex: number
  length: number
}

/**
 * A single message in a thread on `xl/threadedComments/threadedCommentN.xml`.
 * Top-level messages declare a `ref`; replies omit it and link to their
 * parent through `parentId`.
 */
export interface ThreadedComment {
  id: string
  /** A1-style cell ref. Required for thread roots, omitted for replies. */
  ref?: string
  /** GUID matching a {@link ThreadedCommentPerson.id}. */
  personId: string
  /** GUID of the parent comment when this is a reply. */
  parentId?: string
  /** ISO-8601 timestamp from the `dT` attribute. */
  date?: string
  /** Comment body. */
  text: string
  /** Whether the thread is marked resolved. */
  done?: boolean
  /** `@person` mentions inside the text. */
  mentions?: ThreadedCommentMention[]
}

// ── Image ──────────────────────────────────────────────────────────

export interface SheetImage {
  data: Uint8Array
  type: "png" | "jpeg" | "gif" | "svg" | "webp"
  /** Anchor to cell */
  anchor: {
    from: { row: number; col: number }
    to?: { row: number; col: number }
  }
  /**
   * Rendered size in pixels at 96 DPI, stored as EMU in the drawing's
   * `<a:ext>`. Absent on write means the writer's own default size, which
   * is then what the reader reports — a file records a size either way.
   */
  width?: number
  height?: number
  /** Alternative text for screen readers (lands in xdr:cNvPr/@descr). */
  altText?: string
  /** Title/caption for the image (lands in xdr:cNvPr/@title). */
  title?: string
}

// ── Charts (write/clone surface) ────────────────────────────────────
//
// The chart write/clone interfaces have been moved to
// `./xlsx/chart/types.ts` so the chart submodules share one authoritative
// home. We re-export them here so existing consumers
// (`import { SheetChart, ... } from "hucre"`) keep working unchanged.

import type { SheetChart } from "./xlsx/chart/types"

export type {
  ChartBorderDash,
  ChartColor,
  ChartDataLabelPosition,
  ChartDataLabels,
  ChartDataPoint,
  ChartDataTable,
  ChartDisplayBlanksAs,
  ChartErrorBarDirection,
  ChartErrorBarType,
  ChartErrorBarValType,
  ChartErrorBars,
  ChartLegendEntry,
  ChartLineCap,
  ChartLineCompound,
  ChartLineDashStyle,
  ChartLineStroke,
  ChartManualLayout,
  ChartMarker,
  ChartMarkerSymbol,
  ChartProtection,
  ChartScatterStyle,
  ChartSeries,
  ChartShape3D,
  ChartThemeColor,
  ChartThemeColorName,
  ChartTrendline,
  ChartTrendlineType,
  ChartView3D,
  SheetChart,
  WriteChartKind,
} from "./xlsx/chart/types"

// ── Accessibility ──────────────────────────────────────────────────

/**
 * Per-sheet accessibility metadata. Hints to screen readers and
 * input to `audit` from the `hucre/a11y` entry point.
 */
export interface SheetA11y {
  /**
   * Short, human-readable summary of the sheet's purpose. If the
   * workbook does not already declare a `properties.description`,
   * the first non-empty summary across the workbook is copied there
   * so screen readers announce it when the file is opened.
   */
  summary?: string
  /**
   * 0-based row index that should be treated as the column-header
   * row. Used by the audit to verify a header is present and to
   * cross-check tables that span the same range.
   */
  headerRow?: number
}

/** Severity of an accessibility finding. */
export type A11ySeverity = "error" | "warning" | "info"

/** Stable code identifying an accessibility issue. */
export type A11yCode =
  | "no-doc-title"
  | "no-doc-description"
  | "no-header-row"
  | "missing-alt-text"
  | "merged-header-row"
  | "low-contrast"
  | "empty-sheet"
  | "blank-row-in-data"

/** Pinpoint where an issue applies. */
export interface A11yLocation {
  sheet?: string
  /** Cell reference like "B5" or range like "A1:D1". */
  ref?: string
  /** Image index inside `sheet.images`. */
  image?: number
  /** Text-box index inside `sheet.textBoxes`. */
  textBox?: number
}

export interface A11yIssue {
  type: A11ySeverity
  code: A11yCode
  message: string
  location?: A11yLocation
}

// ── Sheet Protection ───────────────────────────────────────────────

export interface SheetProtection {
  password?: string
  sheet?: boolean
  objects?: boolean
  scenarios?: boolean
  selectLockedCells?: boolean
  selectUnlockedCells?: boolean
  formatCells?: boolean
  formatColumns?: boolean
  formatRows?: boolean
  insertColumns?: boolean
  insertRows?: boolean
  insertHyperlinks?: boolean
  deleteColumns?: boolean
  deleteRows?: boolean
  sort?: boolean
  autoFilter?: boolean
  pivotTables?: boolean
}

// ── Sheet View ─────────────────────────────────────────────────────

export interface SheetView {
  showGridLines?: boolean
  showRowColHeaders?: boolean
  zoomScale?: number
  rightToLeft?: boolean
  tabColor?: Color
}

// ── Table (ListObject) ────────────────────────────────────────────

export interface TableDefinition {
  /** Table name (must be unique in workbook, used in structured references) */
  name: string
  /** Display name */
  displayName?: string
  /** Cell range (e.g. "A1:D10") — if not provided, auto-calculated from data */
  range?: string
  /** Column definitions */
  columns: TableColumn[]
  /** Table style name (e.g. "TableStyleMedium2") */
  style?: string
  /** Show banded rows. Default: true */
  showRowStripes?: boolean
  /** Show banded columns. Default: false */
  showColumnStripes?: boolean
  /**
   * Show auto-filter. Default when writing: true. On read this reports
   * whether the table part actually carries an `<autoFilter>` — a table
   * without one has no filter dropdowns, so it reads back `false`.
   */
  showAutoFilter?: boolean
  /** Show total row. Default: false */
  showTotalRow?: boolean
}

export interface TableColumn {
  /** Column header name */
  name: string
  /** Total row function (sum, count, average, min, max, countNums, stdDev, var, custom) */
  totalFunction?: string
  /** Total row formula (for custom) */
  totalFormula?: string
  /** Total row label (text in total cell) */
  totalLabel?: string
}

// ── Row Definition ────────────────────────────────────────────────

export interface RowDef {
  /** Row height in points */
  height?: number
  /** Hide row */
  hidden?: boolean
  /** Outline level (grouping) */
  outlineLevel?: number
  /** Whether this outline group is collapsed */
  collapsed?: boolean
}

// ── Sheet ──────────────────────────────────────────────────────────

/**
 * What kind of sheet a tab holds.
 *
 * `xl/workbook.xml`'s `<sheets>` lists every tab whatever its kind —
 * ECMA-376 `CT_Sheet` covers worksheets, chart sheets, dialog sheets and
 * macro sheets alike, and the *relationship type* is what tells them
 * apart. Only a worksheet has cells.
 */
export type SheetKind = "worksheet" | "chartsheet" | "dialogsheet" | "macrosheet"

export interface Sheet {
  name: string
  /**
   * Cell values as a **dense rectangle**: every row is an array, every
   * row is the same length, and no element is `undefined`.
   *
   * That is what makes `rows[r][c]` safe without a guard on either
   * index, and it is what the readers' bounding-box limits are sized
   * against — the cost of a sheet is its box, not its cell count, which
   * is why {@link ReadOptions.maxTotalCells} bounds the product.
   *
   * It went unwritten and two readers did not hold it: `readXls` and
   * `readXlsb` padded a row only to its own last cell and never
   * allocated a row Excel left empty, so one authored sheet saved three
   * ways came back three shapes, and a gap row came back as `undefined`
   * — which `CellValue` cannot express. See #494.
   */
  rows: CellValue[][]
  /**
   * The kind of tab this is. Absent means `"worksheet"`, which is what
   * all but a handful of sheets are.
   *
   * A workbook containing a chart sheet used to fail to read *entirely*
   * — the chart sheet's relationship is not a `worksheet` one, so the
   * lookup missed and the reader threw, taking every ordinary worksheet
   * beside it down too. Non-worksheet tabs are now read as empty sheets
   * so the indices still line up with Excel's tab bar, and this field is
   * what tells you one apart from a worksheet that happens to be empty.
   * Read-only: hucre cannot author a chart sheet. See #499.
   */
  kind?: SheetKind
  /** Detailed cell data (keyed by "row,col" e.g. "0,2") */
  cells?: Map<string, Cell>
  columns?: ColumnDef[]
  /** Row-level properties (keyed by 0-based row index) */
  rowDefs?: Map<number, RowDef>
  /**
   * Default row height in points, for rows with no `rowDefs` entry.
   * Excel's own default is 15. Written to `<sheetFormatPr defaultRowHeight>`.
   *
   * Before this existed the writer emitted a hard-coded 15 and the reader
   * looked at `<sheetFormatPr>` not at all, so a workbook whose default was
   * 24 came back through readXlsx → writeXlsx with every unstyled row
   * shortened. See #439 §X.
   */
  defaultRowHeight?: number
  /**
   * Default column width in characters, for columns with no `columns[]`
   * entry. Written to `<sheetFormatPr defaultColWidth>`; absent means
   * Excel picks its own from the default font.
   */
  defaultColWidth?: number
  merges?: MergeRange[]
  dataValidations?: DataValidation[]
  conditionalRules?: ConditionalRule[]
  autoFilter?: AutoFilter
  freezePane?: FreezePane
  splitPane?: SplitPane
  images?: SheetImage[]
  protection?: SheetProtection
  pageSetup?: PageSetup
  headerFooter?: HeaderFooter
  view?: SheetView
  hidden?: boolean
  /** Very hidden (only unhideable via VBA) */
  veryHidden?: boolean
  /** Excel Tables (ListObject) defined on this sheet */
  tables?: TableDefinition[]
  /** Row page breaks (0-based row indices) */
  rowBreaks?: number[]
  /** Column page breaks (0-based column indices) */
  colBreaks?: number[]
  /** Outline properties (controls summary row/column position) */
  outlineProperties?: OutlineProperties
  /** Background image data (extracted from worksheet picture relationship) */
  backgroundImage?: Uint8Array
  /** Sparklines (mini-charts in cells) */
  sparklines?: Sparkline[]
  /** Text boxes (shapes with text) */
  textBoxes?: SheetTextBox[]
  /**
   * Excel 365 threaded comments for this sheet. Stored physically in
   * `xl/threadedComments/threadedCommentN.xml` and resolved against
   * the workbook-wide person list (`Workbook.persons`).
   */
  threadedComments?: ThreadedComment[]
  /** Accessibility metadata for screen readers and the `audit` helper. */
  a11y?: SheetA11y
  /**
   * Pivot table instances hosted on this sheet. The body lives in
   * `xl/pivotTables/pivotTableN.xml`; each instance points at a
   * workbook-level cache via `cacheId`.
   */
  pivotTables?: PivotTable[]
  /**
   * Slicers attached to this sheet (Excel 2010+). Resolved from
   * `xl/slicers/slicerN.xml` parts referenced via this sheet's rels.
   */
  slicers?: Slicer[]
  /**
   * Timeline slicers attached to this sheet (Excel 2013+). Resolved from
   * `xl/timelines/timelineN.xml` parts referenced via this sheet's rels.
   */
  timelines?: Timeline[]
  /**
   * Charts anchored on this sheet, resolved from `xl/charts/chartN.xml`
   * parts referenced via the sheet's drawing. Hucre does not yet author
   * charts; the entries surface for inspection on read and survive
   * roundtrip when the sheet has no hucre-managed images.
   */
  charts?: Chart[]
}

// ── Workbook Properties ────────────────────────────────────────────

export interface WorkbookProperties {
  title?: string
  subject?: string
  creator?: string
  keywords?: string
  description?: string
  lastModifiedBy?: string
  created?: Date
  modified?: Date
  company?: string
  manager?: string
  category?: string
  /** Custom properties */
  custom?: Record<string, string | number | boolean | Date>
}

// ── External Workbook Links ────────────────────────────────────────

/** Cached cell type as encoded in `cell/@t`. Mirrors OOXML cell type codes. */
export type ExternalCellType = "n" | "s" | "b" | "e" | "str"

export interface ExternalCachedCell {
  /** A1-style reference within the external sheet. */
  ref: string
  type: ExternalCellType
  /** Cached value. Strings include error text for `t="e"`. */
  value: string | number | boolean
}

export interface ExternalSheetData {
  /** 0-based index into the external workbook's sheet list. */
  sheetId: number
  cells: ExternalCachedCell[]
}

export interface ExternalDefinedName {
  name: string
  refersTo?: string
  /** Sheet-local index when present; omitted for workbook-level names. */
  sheetId?: number
}

/**
 * A reference to another workbook resolved via
 * `xl/externalLinks/externalLinkN.xml`. Cached values follow Excel's
 * formula syntax `[N]Sheet!Ref`, where `N` is this entry's 1-based
 * position in `Workbook.externalLinks`.
 */
export interface ExternalLink {
  /** Target path of the linked workbook (URL, file path, or local entry). */
  target: string
  /** Almost always `"External"`. Mirrors the `TargetMode` attribute. */
  targetMode?: "External" | "Internal"
  /** External workbook's sheets in declaration order. */
  sheetNames: string[]
  /** Cached cell values, keyed by external sheet id. */
  sheetData: ExternalSheetData[]
  /** Defined names declared in the external workbook. */
  definedNames?: ExternalDefinedName[]
}

// ── Cell-Embedded Images (WPS DISPIMG / cellimages) ───────────────

/**
 * An image embedded inside a cell via the WPS Office cellimages mechanism
 * (also recognized by recent Excel versions). The image is referenced from
 * a cell formula `=_xlfn.DISPIMG("<id>", 1)` and the binary lives in the
 * package as a regular media part. Unlike `SheetImage` (which is anchored
 * to a drawing rectangle on a sheet), a `CellImage` is workbook-wide and
 * can be referenced from any number of cells.
 */
export interface CellImage {
  /**
   * Stable image identifier as it appears inside the DISPIMG formula
   * (`name` attribute on `xdr:cNvPr`). For example `"ID_2A8C..."`.
   */
  id: string
  /** Image binary, extracted from the package media folder. */
  data: Uint8Array
  /** Image format inferred from the media file extension. */
  type: SheetImage["type"]
  /** Optional human-readable description (`descr` attribute). */
  description?: string
}

// ── Pivot Tables ───────────────────────────────────────────────────

/**
 * Aggregation function for a pivot table data field. Mirrors the
 * `subtotal` attribute on `<c:dataField>` in OOXML.
 */
export type PivotDataFieldFunction =
  | "sum"
  | "count"
  | "average"
  | "max"
  | "min"
  | "product"
  | "countNums"
  | "stdDev"
  | "stdDevp"
  | "var"
  | "varp"

/**
 * Field role in a pivot table layout. `row`, `col`, `page`, and `data`
 * mirror the four standard axes; `hidden` means the field exists in the
 * cache but is not currently placed on any axis.
 */
export type PivotFieldAxis = "row" | "col" | "page" | "data" | "hidden"

export interface PivotField {
  /**
   * Display name. Reads from the `<cacheField name="...">` attribute on
   * the matching field index in the pivot cache definition.
   */
  name: string
  /**
   * Where the field appears in the pivot table. `hidden` covers cache
   * fields that are present but not placed on any axis.
   */
  axis: PivotFieldAxis
  /** When `axis === "data"`, the aggregation applied to the values. */
  function?: PivotDataFieldFunction
  /**
   * Display name overlay for data fields (the `name` attribute on
   * `<dataField>`). Falls back to `name` when absent.
   */
  displayName?: string
}

/**
 * A pivot table instance, attached to the sheet that hosts its layout.
 * The `cacheId` references one of the workbook-level pivot caches that
 * back this table.
 */
export interface PivotTable {
  /** Pivot table name (`<pivotTableDefinition name="...">`). */
  name: string
  /**
   * Index into `Workbook.pivotCaches`. Mirrors the workbook-level
   * `cacheId` attribute on `<pivotCache>` rather than the per-table
   * relationship — that way a model author who reorders the cache
   * array keeps the link sound.
   */
  cacheId: number
  /**
   * Output range on the host sheet, e.g. `"A3:D20"`. Empty string when
   * the source omits a `<location>` element.
   */
  location: string
  /** Number of header rows above the data rows. */
  firstHeaderRow?: number
  /** Number of body rows reserved for column-axis labels. */
  firstDataRow?: number
  /** Column index of the first data row (0-based). */
  firstDataCol?: number
  /** Number of pages declared in `<pageFields>`. */
  rowPageCount?: number
  /** Number of column-axis page-break positions. */
  colPageCount?: number
  /**
   * Pivot fields in declaration order. The position in this array is
   * the field index used by `<rowItems>`, `<colItems>`, etc.
   */
  fields: PivotField[]
  /** Pivot-table style name (`<pivotTableStyleInfo name="...">`). */
  styleName?: string
  /** Whether the data field caption is shown. */
  dataCaption?: string
}

/**
 * Workbook-level pivot cache: source range plus cached field metadata.
 * Multiple pivot tables can share a cache so the same source data only
 * gets indexed once.
 */
export interface PivotCache {
  /**
   * Cache id Excel uses to wire pivot tables to caches. Mirrors the
   * `cacheId` attribute on `<workbook><pivotCaches><pivotCache>`.
   */
  cacheId: number
  /**
   * Source range, e.g. `"Sheet1!$A$1:$C$100"` or a defined-name
   * reference. Empty string for non-worksheet sources.
   */
  sourceRef?: string
  /** Source sheet name when the source is a worksheet range. */
  sourceSheet?: string
  /**
   * Source type: `worksheet` (range or table on a sheet), `external`
   * (linked workbook / database), `consolidation`, or `scenario`. Most
   * real workbooks use `worksheet`.
   */
  sourceType?: "worksheet" | "external" | "consolidation" | "scenario"
  /** Cached field names in declaration order. */
  fieldNames: string[]
  /** Whether a `pivotCacheRecords{N}.xml` part is present. */
  hasRecords?: boolean
}

/**
 * A data field placement on a {@link WritePivotTable}.
 *
 * `field` names a column in the source data; `function` selects the
 * aggregation Excel applies (`sum` is the default). `displayName` becomes
 * the column header on the rendered pivot — it defaults to
 * `"<Function> of <field>"`, mirroring Excel's auto-label.
 */
export interface WritePivotDataField {
  /** Source column name (must match an entry in the source header row). */
  field: string
  /** Aggregation function. Default: `"sum"`. */
  function?: PivotDataFieldFunction
  /** Optional display name override. Default: e.g. `"Sum of Revenue"`. */
  displayName?: string
  /** Optional number format for aggregated values. Default: General. */
  numberFormat?: string
}

/**
 * Author a pivot table on a sheet.
 *
 * Phase 1 covers the most common dashboard use case: a tabular source on
 * one sheet, summarised onto another sheet with row / column / value
 * fields. Hucre emits the pivot cache (definition + cached records), the
 * pivot table layout, and all required relationships and content types.
 *
 * The actual numeric layout (row totals, grand totals, value cells) is
 * left for Excel to compute on first open via `<calcPr fullCalcOnLoad="1"/>`
 * — Phase 1 ships the structural skeleton, not pre-computed cells.
 */
export interface WritePivotTable {
  /** Pivot table name shown in Excel's `Field List`. */
  name: string
  /**
   * Source sheet name. Defaults to the sheet the pivot is declared on
   * when omitted — handy for pivots that summarise their own sheet's
   * data.
   */
  sourceSheet?: string
  /**
   * Source range covering the header row plus all data rows
   * (e.g. `"A1:C100"`). Auto-detected from the source sheet's `rows`
   * length when omitted.
   */
  sourceRange?: string
  /**
   * Top-left anchor for the rendered pivot table on the host sheet
   * (e.g. `"A3"`). Default: `"A1"`.
   */
  targetCell?: string
  /** Source columns laid out on the row axis, in order. */
  rows?: string[]
  /** Source columns laid out on the column axis, in order. */
  columns?: string[]
  /** Source columns laid out as page (filter) fields, in order. */
  pages?: string[]
  /** Aggregated value fields. Each entry becomes one data column. */
  values: WritePivotDataField[]
  /**
   * Pivot table style name (e.g. `"PivotStyleLight16"`). Default:
   * `"PivotStyleLight16"` — the modern Excel default.
   */
  styleName?: string
  /**
   * Caption shown above the data fields when there is more than one.
   * Default: `"Values"` (Excel's built-in caption).
   */
  dataCaption?: string
}

// ── Slicers & Timelines ────────────────────────────────────────────

/**
 * A slicer (Excel 2010+ visual filter). Slicers live on a worksheet and
 * are backed by a {@link SlicerCache} that holds the actual filter state.
 *
 * Slicers come from `xl/slicers/slicerN.xml`. Each slicer entry inside
 * a slicer file is exposed as one record in {@link Sheet.slicers}.
 */
export interface Slicer {
  /** Programmatic name. Mirrors `slicer/@name`. */
  name: string
  /** Slicer cache identifier this slicer references. Mirrors `slicer/@cache`. */
  cache: string
  /** Display caption shown in the header. Mirrors `slicer/@caption`. */
  caption?: string
  /** Number of columns in the slicer button grid. Mirrors `slicer/@columnCount`. */
  columnCount?: number
  /** Built-in style id, e.g. `SlicerStyleLight1`. Mirrors `slicer/@style`. */
  style?: string
  /** Sort order for items. Mirrors `slicer/@sortOrder` (e.g. `ascending`, `descending`). */
  sortOrder?: string
  /** Row height in EMUs. Mirrors `slicer/@rowHeight`. */
  rowHeight?: number
}

/**
 * Workbook-level slicer cache. Stores the filter source and selection
 * state shared by one or more {@link Slicer} instances.
 *
 * Slicer caches come from `xl/slicerCaches/slicerCacheN.xml`.
 */
export interface SlicerCache {
  /** Programmatic name. Mirrors `slicerCacheDefinition/@name`. */
  name: string
  /** Source identifier — typically the cache definition's source ref. */
  sourceName?: string
  /**
   * Pivot tables this cache filters, when sourced from a pivot table.
   * Each entry is the `tabId` (sheet index) + `name` of a pivot table.
   */
  pivotTables?: SlicerCachePivotTable[]
  /** Excel Table this cache filters, when sourced from a table. */
  tableSource?: SlicerCacheTableSource
}

export interface SlicerCachePivotTable {
  /** 0-based sheet tab id of the sheet hosting the pivot table. */
  tabId: number
  /** Pivot table name. */
  name: string
}

export interface SlicerCacheTableSource {
  /** Excel Table name. */
  name: string
  /** Column referenced in the table. */
  column?: string
}

/**
 * Timeline slicer (Excel 2013+ date-range filter). Like {@link Slicer}
 * but constrained to date columns and rendered as a draggable date band.
 *
 * Timelines come from `xl/timelines/timelineN.xml`.
 */
export interface Timeline {
  /** Programmatic name. */
  name: string
  /** Cache identifier this timeline references. */
  cache: string
  /** Display caption. */
  caption?: string
  /** Built-in style id, e.g. `TimeSlicerStyleLight1`. */
  style?: string
  /** Granularity: `years`, `quarters`, `months`, or `days`. */
  level?: string
  /** Whether the time-level selector is shown. */
  showHeader?: boolean
  /** Whether the selection-label band is shown. */
  showSelectionLabel?: boolean
  /** Whether the time-level row is shown. */
  showTimeLevel?: boolean
  /** Whether the horizontal scrollbar is shown. */
  showHorizontalScrollbar?: boolean
}

/**
 * Workbook-level timeline cache. Stores the date column and selected
 * range shared by one or more {@link Timeline} instances.
 *
 * Timeline caches come from `xl/timelineCaches/timelineCacheN.xml`.
 */
export interface TimelineCache {
  /** Programmatic name. */
  name: string
  /** Source identifier. */
  sourceName?: string
  /** Pivot tables this cache filters. */
  pivotTables?: SlicerCachePivotTable[]
}

// ── Charts (read surface) ───────────────────────────────────────────
//
// The chart read interfaces have been moved to `./xlsx/chart/types.ts`.
// Re-exported here so existing consumers
// (`import { Chart, ChartAxisInfo, ... } from "hucre"`) keep working
// unchanged.

import type { Chart } from "./xlsx/chart/types"

export type {
  Chart,
  ChartAnchor,
  ChartAxisCrossBetween,
  ChartAxisCrosses,
  ChartAxisDispUnit,
  ChartAxisDispUnits,
  ChartAxisGridlines,
  ChartAxisInfo,
  ChartAxisLabelAlign,
  ChartAxisNumberFormat,
  ChartAxisScale,
  ChartAxisTickLabelPosition,
  ChartAxisTickMark,
  ChartBarGrouping,
  ChartDataLabelsInfo,
  ChartKind,
  ChartLegendPosition,
  ChartLineAreaGrouping,
  ChartSeriesInfo,
} from "./xlsx/chart/types"

// ── Workbook ───────────────────────────────────────────────────────

export interface Workbook {
  sheets: Sheet[]
  properties?: WorkbookProperties
  namedRanges?: NamedRange[]
  /** Date system: 1900 (default/Windows) or 1904 (Mac) */
  dateSystem?: "1900" | "1904"
  /**
   * Default font for the workbook — `fonts[0]` in `xl/styles.xml`, the
   * entry every cell format inherits from unless it names another.
   */
  defaultFont?: FontStyle
  /**
   * Active sheet index — the tab the file opens on.
   *
   * Undefined when the file opens on the first tab: `activeTab="0"` is the
   * OOXML default and is indistinguishable from a file that says nothing,
   * so both collapse to `undefined` and round-trip identically.
   */
  activeSheet?: number
  /** Theme color palette (resolved from xl/theme/theme1.xml) */
  themeColors?: string[]
  /** Workbook-level protection */
  workbookProtection?: {
    lockStructure?: boolean
    lockWindows?: boolean
  }
  /**
   * Workbook-wide person directory referenced from threaded comments.
   * Each `ThreadedComment.personId` resolves against this list.
   */
  persons?: ThreadedCommentPerson[]
  /**
   * External workbook references, resolved from
   * `xl/externalLinks/externalLinkN.xml`. The 1-based position in this
   * array matches the `[N]` prefix used in formulas like `[1]Sheet1!A1`.
   */
  externalLinks?: ExternalLink[]
  /**
   * Cell-embedded images (WPS DISPIMG mechanism).
   *
   * Resolved from `xl/cellimages.xml`. Cells reference these images via
   * `=_xlfn.DISPIMG("<id>", 1)` formulas — match `CellImage.id` against
   * the first argument to look up the binary.
   */
  cellImages?: CellImage[]
  /**
   * Workbook-level pivot caches resolved from
   * `xl/pivotCache/pivotCacheDefinitionN.xml`. Sheet-level
   * `PivotTable.cacheId` references entries here.
   */
  pivotCaches?: PivotCache[]
  /**
   * Slicer caches resolved from `xl/slicerCaches/slicerCacheN.xml`.
   * The 1-based position in this array matches the `N` in the source path.
   */
  slicerCaches?: SlicerCache[]
  /**
   * Timeline caches resolved from `xl/timelineCaches/timelineCacheN.xml`.
   * The 1-based position in this array matches the `N` in the source path.
   */
  timelineCaches?: TimelineCache[]
}

// ── Read diagnostics ───────────────────────────────────────────────

/** What a reader had to drop, and where. */
export interface ReadWarning {
  /** What kind of problem this is, for programmatic handling. */
  code:
    | "unresolved-shared-string"
    | "unresolved-style"
    | "unresolved-dxf"
    | "unresolved-hyperlink"
    | "unusable-paper-size"
    | "malformed-cell-ref"
  /** A sentence a person can act on. */
  message: string
  /** The sheet it happened in, when the reader knows. */
  sheet?: string
  /** 0-based cell position, when the problem is a cell's. */
  row?: number
  col?: number
}

// ── Read Options ───────────────────────────────────────────────────

/**
 * Lightweight metadata exposed to a {@link ReadOptions.sheets} predicate
 * before the worksheet body is parsed. Includes the cheaply-known fields
 * read from the workbook directory — name, index, and visibility state.
 *
 * `hidden` and `veryHidden` are XLSX-only; ODS does not expose visibility
 * in the table directory and they will be `undefined`.
 */
export interface SheetFilterInfo {
  /** Sheet name as declared in the workbook directory. */
  name: string
  /** 0-based position in the workbook's sheet list. */
  index: number
  /** XLSX `<sheet state="hidden">`. Undefined for ODS. */
  hidden?: boolean
  /** XLSX `<sheet state="veryHidden">`. Undefined for ODS. */
  veryHidden?: boolean
}

/**
 * Predicate form of {@link ReadOptions.sheets}. Receives one
 * {@link SheetFilterInfo} per sheet in workbook order; returning `true`
 * includes the sheet, `false` skips it.
 */
export type SheetFilter = (info: SheetFilterInfo, index: number) => boolean

/**
 * The options every reader honours.
 *
 * Each reader takes its own extension of this — {@link XlsxReadOptions},
 * {@link OdsReadOptions}, {@link XlsbReadOptions}, {@link XlsReadOptions}
 * — carrying exactly the fields it reads. v1 had one `ReadOptions` for
 * all four and a table in its doc comment saying which reader ignored
 * what; `readXls(bytes, { password })` compiled and did nothing. The type
 * is the table now.
 */
export interface ReadOptionsBase {
  /**
   * Maximum number of bytes buffered from a `ReadableStream` input.
   * Default: 1 GiB ({@link MAX_INPUT_BYTES}). A stream that exceeds it
   * fails with a `ParseError` instead of growing until the process runs
   * out of memory. Ignored for `Uint8Array` / `ArrayBuffer` input, which
   * the caller has already allocated.
   */
  maxInputBytes?: number
  /**
   * Maximum number of cells a single sheet may be normalized into —
   * `rows` is a dense rectangle, so this bounds the bounding box rather
   * than the cell count. Default: 20,000,000 ({@link MAX_TOTAL_CELLS}).
   *
   * The default refuses two legal cells at `A1` and `XFD1048576`, which
   * describe 1.7e10 slots from a few hundred bytes of XML. It also
   * refuses a legitimate 25-million-cell sheet, which is why this is a
   * number rather than a ceiling: raise it when you know the file, and
   * budget roughly 8 bytes per slot for the array alone.
   *
   */
  maxTotalCells?: number
}

/** Options of readers whose container is a ZIP archive. */
export interface ZipReadOptions {
  /**
   * Maximum number of bytes any single ZIP entry may decompress to.
   * Default: 2 GiB ({@link MAX_DECOMPRESSED_BYTES}).
   *
   * This is the zip-bomb bound — an entry that claims a small compressed
   * size and expands past it fails with a `ZipError` rather than being
   * allowed to allocate. Raising it is the one on this list where a
   * caller should be sure the input is trusted.
   *
   * Honoured wherever the container is a ZIP: `readXlsx`, `readOds`.
   */
  maxDecompressedBytes?: number
}

/** Options of readers that can open an ECMA-376 Agile-encrypted package. */
export interface EncryptedReadOptions {
  /** Password for encrypted files */
  password?: string
  /**
   * Maximum password-derivation spin count accepted from an encrypted
   * workbook. Default: 10,000,000 ({@link MAX_SPIN_COUNT}).
   *
   * Office writes 100,000. The bound exists so a hostile file cannot
   * name a count that pins a CPU for minutes; raising it means agreeing
   * to spend that time.
   */
  maxSpinCount?: number
}

/** Options `readXlsx` (and `openXlsx`) honour. */
export interface XlsxReadOptions extends ReadOptionsBase, ZipReadOptions, EncryptedReadOptions {
  /**
   * Which sheets to read.
   * - `Array<number | string>` — explicit indexes and/or names.
   * - `(info, index) => boolean` — predicate evaluated against
   *   {@link SheetFilterInfo} before each worksheet body is parsed.
   *   Useful for selecting by visibility, e.g.
   *   `sheets: (info) => !info.hidden && !info.veryHidden`.
   *
   * Default: all sheets.
   */
  sheets?: Array<number | string> | SheetFilter
  /**
   * Date system override. Default: `"auto"`, which takes the file's own
   * `date1904` flag.
   */
  dateSystem?: "1900" | "1904" | "auto"
  /**
   * Whether to read styles. Default: false (faster without).
   *
   * **A resolved style's parts are shared, not copied.** `xl/styles.xml`
   * holds one font, fill and border record per distinct format, and every
   * cell that indexes it gets that same object — copying per cell nearly
   * doubles peak memory on a styled read for a guarantee most callers
   * never need. So `cells.get(a).style.font === cells.get(b).style.font`
   * whenever `a` and `b` share a format, and writing through one changes
   * both. Use `cloneCellStyle` before editing a single cell's format.
   */
  readStyles?: boolean
  /** Maximum number of data rows to read per sheet. Default: unlimited */
  maxRows?: number
  /** Cell range to read (e.g. "A1:D10"). Only cells within this range are returned. */
  range?: string
  /**
   * Return cells without materializing the grid. Default: false.
   *
   * `Sheet.rows` is a dense rectangle, so the cost of a read is the
   * bounding box rather than the cell count — which is right for almost
   * every sheet and wrong for a sparse one. A real workbook with 82,000
   * values scattered over a 305,612,208-slot box (0.03% filled) could
   * not be read at all: raising {@link maxTotalCells} trades a clean
   * error for a multi-gigabyte allocation, `range` needs you to already
   * know where the data is, and `maxRows` bounds rows when the problem
   * is columns. See #501.
   *
   * With this set, `rows` comes back empty and every cell that carries
   * something is in {@link Sheet.cells}, keyed `"row,col"`. Memory then
   * tracks the values rather than the box, and the bounding-box limit
   * does not apply because nothing dense is built.
   *
   * `streamXlsxRows` is the other answer and the better one when you
   * only need to walk the rows once; this is for random access, or for
   * when you want a `Workbook`.
   *
   * XLSX only.
   */
  sparse?: boolean
  /**
   * Called for each thing a reader had to drop.
   *
   * The readers are lenient on purpose — a corrupt reference yields
   * `null` rather than an exception, because a spreadsheet is a format
   * you receive rather than one you control. But leniency used to be the
   * *only* mode: a cell pointing at a shared string that is not there
   * came back as `null`, indistinguishable from a cell that was
   * genuinely empty, and nothing said which. See #439 §S.
   *
   * ```ts
   * const warnings: ReadWarning[] = []
   * const wb = await readXlsx(bytes, { onWarning: (w) => warnings.push(w) })
   * if (warnings.length) console.warn(`${warnings.length} problem(s) in this file`)
   * ```
   *
   * Nothing changes when it is omitted. This is a side channel, not part
   * of the document, which is why it is a callback rather than a field on
   * `Workbook`.
   */
  onWarning?: (warning: ReadWarning) => void
}

/**
 * Options `readOds` honours. ODS stores ISO date strings, so there is no
 * 1900/1904 system to pick; ODS encryption is not implemented (#156).
 */
export interface OdsReadOptions extends ReadOptionsBase, ZipReadOptions {
  /**
   * Which sheets to read.
   * - `Array<number | string>` — explicit indexes and/or names.
   * - `(info, index) => boolean` — predicate evaluated against
   *   {@link SheetFilterInfo} before each worksheet body is parsed.
   *   Useful for selecting by visibility, e.g.
   *   `sheets: (info) => !info.hidden && !info.veryHidden`.
   *
   * Default: all sheets.
   */
  sheets?: Array<number | string> | SheetFilter
  /**
   * Whether to read styles. Default: false (faster without).
   *
   * **A resolved style's parts are shared, not copied.** `xl/styles.xml`
   * holds one font, fill and border record per distinct format, and every
   * cell that indexes it gets that same object — copying per cell nearly
   * doubles peak memory on a styled read for a guarantee most callers
   * never need. So `cells.get(a).style.font === cells.get(b).style.font`
   * whenever `a` and `b` share a format, and writing through one changes
   * both. Use `cloneCellStyle` before editing a single cell's format.
   */
  readStyles?: boolean
  /** Maximum number of data rows to read per sheet. Default: unlimited */
  maxRows?: number
  /** Cell range to read (e.g. "A1:D10"). Only cells within this range are returned. */
  range?: string
}

/**
 * Options `readXlsb` honours. The binary reader surfaces values, sheet
 * names and merges only, so there are no styles to ask for and no
 * per-sheet selection.
 */
export interface XlsbReadOptions extends ReadOptionsBase, ZipReadOptions, EncryptedReadOptions {
  /**
   * Date system override. Default: `"auto"`, which takes the file's own
   * `date1904` flag.
   */
  dateSystem?: "1900" | "1904" | "auto"
}

/** Options `readXls` honours. A `.xls` is a CFB container, not a ZIP. */
export interface XlsReadOptions extends ReadOptionsBase {
  /**
   * Date system override. Default: `"auto"`, which takes the file's own
   * `date1904` flag.
   */
  dateSystem?: "1900" | "1904" | "auto"
}

/**
 * What `read()` accepts: it does not know the format before it looks at
 * the bytes, so it takes the widest reader's options and hands the
 * detected reader the fields it understands.
 */
export type ReadOptions = XlsxReadOptions

// ── Write Options ──────────────────────────────────────────────────

export interface WriteOptions {
  sheets: WriteSheet[]
  properties?: WorkbookProperties
  namedRanges?: NamedRange[]
  defaultFont?: FontStyle
  dateSystem?: "1900" | "1904"
  /** Active sheet index (0-based). Default: 0 */
  activeSheet?: number
  /** Workbook-level protection (lock structure/windows) */
  workbookProtection?: {
    lockStructure?: boolean
    lockWindows?: boolean
    password?: string
  }
  /** String storage mode. Default: "shared"
   *  - "shared": shared string table (smaller files with repeated strings)
   *  - "inline": inline strings per cell (faster write, larger files)
   */
  stringMode?: "shared" | "inline"
  /** VBA project binary (vbaProject.bin) to embed. Output becomes macro-enabled (.xlsm). */
  vbaProject?: Uint8Array
  /**
   * Encrypt the output as a password-protected workbook (ECMA-376 Agile,
   * the Excel 2010+ scheme). The result is an OLE2/CFB container that Excel
   * opens after prompting for the password.
   *
   * `spinCount` is the password key-derivation iteration count (default
   * 100000, matching Excel). Lower it only when the speed/security trade-off
   * genuinely calls for it — the value is stored in the file, so any reader
   * (including Excel) honors it.
   */
  encryption?: { password: string; spinCount?: number }
}

export interface WriteSheet {
  name: string
  columns?: ColumnDef[]
  /**
   * Raw row data (array of arrays).
   *
   * An entry is a {@link CellValue}, or a cell object — `{ value, style }`,
   * `{ formula }`, anything a {@link Cell} carries — written where the
   * value goes. The streaming writers have taken that shape since they
   * existed; the buffered ones now do too, so styling one cell no longer
   * means naming its position again in {@link cells} (#433). Where both
   * describe a position, {@link cells} wins.
   */
  rows?: Array<Array<CellValue | Partial<Cell>>>
  /**
   * Object data (array of objects — uses column keys). A value may be a scalar
   * {@link CellValue} or a rich {@link HyperlinkValue} for inline clickable links.
   */
  data?: Array<Record<string, CellValue | HyperlinkValue>>
  /** Detailed cell overrides (keyed by "row,col") */
  cells?: Map<string, Partial<Cell>>
  /**
   * Default row height in points, for rows with no `rowDefs` entry.
   * Excel's own default is 15. Written to `<sheetFormatPr defaultRowHeight>`.
   *
   * Before this existed the writer emitted a hard-coded 15 and the reader
   * looked at `<sheetFormatPr>` not at all, so a workbook whose default was
   * 24 came back through readXlsx → writeXlsx with every unstyled row
   * shortened. See #439 §X.
   */
  defaultRowHeight?: number
  /**
   * Default column width in characters, for columns with no `columns[]`
   * entry. Written to `<sheetFormatPr defaultColWidth>`; absent means
   * Excel picks its own from the default font.
   */
  defaultColWidth?: number
  /**
   * Merged ranges, as coordinates or as A1 strings — `"A1:C1"` and
   * `{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }` mean the same
   * thing. See #474; the read model stays coordinates, because that is
   * what the reader produces.
   */
  merges?: Array<MergeRange | string>
  dataValidations?: DataValidation[]
  conditionalRules?: ConditionalRule[]
  autoFilter?: AutoFilter
  freezePane?: FreezePane
  splitPane?: SplitPane
  images?: SheetImage[]
  protection?: SheetProtection
  pageSetup?: PageSetup
  headerFooter?: HeaderFooter
  view?: SheetView
  hidden?: boolean
  veryHidden?: boolean
  /** Excel Tables (ListObject) to define on this sheet */
  tables?: TableDefinition[]
  /** Row page breaks (0-based row indices) */
  rowBreaks?: number[]
  /** Column page breaks (0-based column indices) */
  colBreaks?: number[]
  /** Row-level properties (keyed by 0-based row index) */
  rowDefs?: Map<number, RowDef>
  /** Outline properties (controls summary row/column position) */
  outlineProperties?: OutlineProperties
  /** Background image for the worksheet (watermark) */
  backgroundImage?: Uint8Array
  /** Sparklines (mini-charts in cells) */
  sparklines?: Sparkline[]
  /** Text boxes (shapes with text) */
  textBoxes?: SheetTextBox[]
  /**
   * Native Excel charts (bar, column, line, pie, scatter, area). Charts
   * share the worksheet's drawing part with images and text boxes.
   */
  charts?: SheetChart[]
  // No `threadedComments` here, deliberately. Authoring Excel 365 threaded
  // comments is a roadmap item, not a shipped feature: it needs a
  // `xl/threadedComments/` part, a workbook-wide `xl/persons/person.xml`,
  // and the legacy `<comment>` fallback Excel expects alongside them. The
  // field used to sit here typed and accepted, and was silently discarded
  // — see #404. `Sheet.threadedComments` is real: they are read, and
  // preserved through `openXlsx` → `saveXlsx`.
  /**
   * Pivot tables anchored on this sheet. The source data is read from
   * either the same sheet or a sibling sheet identified by
   * {@link WritePivotTable.sourceSheet}.
   */
  pivotTables?: WritePivotTable[]
  /** Accessibility metadata for screen readers and the `audit` helper. */
  a11y?: SheetA11y
}

// ── Outline Properties ────────────────────────────────────────────

export interface OutlineProperties {
  /** Summary rows appear below detail rows. Default: true */
  summaryBelow?: boolean
  /** Summary columns appear to the right of detail columns. Default: true */
  summaryRight?: boolean
}

// ── CSV Options ────────────────────────────────────────────────────

/**
 * Options shared by `parseCsv`, `parseCsvObjects` and `streamCsvRows` —
 * every one of them means the same thing in all three.
 *
 * `schema` used to live here and was honoured by no CSV reader at all; it
 * was removed before v1 rather than frozen. Validate with
 * `validateWithSchema` on the parsed rows, which does implement it.
 */
export interface CsvReadOptions {
  /**
   * How to decode byte input. Ignored when a string is passed — the
   * caller has already decided.
   *
   * Any label from the WHATWG Encoding Standard, which is what
   * `TextDecoder` accepts: `"utf-8"`, `"utf-16le"`, `"windows-1254"`,
   * `"iso-8859-9"`, and the rest. Default: the encoding the file's
   * byte-order mark declares, or UTF-8 when it carries none.
   *
   * There is no detection beyond the mark. A mark is a statement the file
   * makes about itself; telling windows-1254 from windows-1252 by byte
   * frequency is a guess, and a wrong one often enough to be worse than
   * asking. If your source is a legacy-encoded export — Excel on a
   * Turkish or Central European Windows, say — name it. See #475.
   */
  encoding?: string
  /** Field delimiter. Default: auto-detect */
  delimiter?: string
  /** Quote character. Default: " */
  quote?: string
  /**
   * Escape character. Default: " (RFC 4180 doubled quotes)
   *
   * Read-only on purpose: set it to read a foreign dialect (a backslash
   * escape, say), but the writers always emit RFC 4180 doubled quotes.
   * Writing a backslash dialect losslessly would need the escape character
   * itself escaped, which this parser does not decode — a half-implemented
   * `escape` on the write side would corrupt a value ending in one.
   */
  escape?: string
  /**
   * Whether the first row is a header. Default: false. The row is still
   * returned; it names columns for `transformValue`, and `skipHeaderRow`
   * is what consumes it. The same name `toHtml` and `toMarkdown` use for
   * the same question.
   */
  hasHeaderRow?: boolean
  /** Skip BOM if present. Default: true */
  skipBom?: boolean
  /** Type inference for numbers, booleans, dates. Default: false */
  typeInference?: boolean
  /** Keep strings with leading zeros (e.g. "0123") as strings instead of converting to numbers. Default: true */
  preserveLeadingZeros?: boolean
  /** Skip empty rows. Default: false */
  skipEmptyRows?: boolean
  /** Comment character (lines starting with this are skipped) */
  comment?: string
  /** Maximum number of data rows to parse. When set, parsing stops after this many rows. */
  maxRows?: number
  /** Skip the first N lines before parsing (useful for files with metadata headers above the CSV data). */
  skipLines?: number
  /** Called for each row during parsing, enabling progressive processing without buffering all rows. */
  onRow?: (row: CellValue[], index: number) => void
  /** Transform each header string when `header: true`. Called on each header value. */
  transformHeader?: (header: string, index: number) => string
  /** Transform each cell value after type inference. Called on every cell. */
  transformValue?: (value: CellValue, header: string, row: number, col: number) => CellValue
  /** Fast mode: skip quote handling and just split by delimiter/newlines. Faster for files known to have no quoted fields. Default: false */
  fastMode?: boolean
  /**
   * Drop the header row from the output instead of yielding it.
   *
   * `header: true` only marks the first row as a header — it is still
   * returned, and only used to name columns for {@link transformValue}.
   * Set this when you want the header consumed rather than emitted.
   * Default: false
   */
  skipHeaderRow?: boolean
  /**
   * Undo {@link CsvWriteOptions.escapeFormulae}: drop the leading `'` from
   * values that start with one of the characters the writer escapes for
   * (`= + - @ | \t \r \n \0`). Runs before type inference, so `'-5` reads
   * back as the number -5. Default: false
   *
   * Set it only for input produced with `escapeFormulae: true` — a source
   * value that genuinely began `'-5` is written unescaped, and this cannot
   * tell the two apart. Values whose apostrophe is followed by anything
   * else (`'quoted'`, `'tis`) are never touched.
   */
  unescapeFormulae?: boolean
}

export interface CsvWriteOptions {
  /** Field delimiter. Default: "," */
  delimiter?: string
  /** Line separator. Default: "\r\n" (CRLF per RFC 4180) */
  lineSeparator?: string
  /** Quote character. Default: " */
  quote?: string
  /** Quote style. Default: "required" */
  quoteStyle?: "all" | "required" | "none"
  /**
   * Header names to write, when the rows carry none of their own. For
   * `writeCsvObjects` and the streaming writers this is also the column
   * order; `columns` wins where both are given.
   */
  headers?: string[]
  /**
   * Whether to write a header line at all. Default: true wherever one is
   * known — explicit `headers`, or object rows whose keys name the columns.
   */
  writeHeader?: boolean
  /** Prepend UTF-8 BOM (for Excel compatibility). Default: false */
  bom?: boolean
  /**
   * Date format string. Default: ISO 8601 (`toISOString()`).
   *
   * Takes the same tokens as the exported `formatDate` and as a `numFmt`
   * anywhere else in the library — `yyyy`/`yy`, `mmmm`/`mmm`/`mm`/`m`,
   * `dddd`/`ddd`/`dd`/`d`, `hh`/`h`, `mm`/`m` (minutes, after an hour
   * token), `ss`, `AM/PM` — and is case-insensitive, so `YYYY-MM-DD` and
   * `yyyy-mm-dd` both work. Components are read in **UTC**, matching the
   * ISO default and every other date path in the library.
   *
   * **One-way.** The readers recognize ISO 8601 and nothing else, so a
   * `Date` written with the default round-trips as a `Date` under
   * `typeInference`, while any custom format comes back a string — a
   * reader cannot tell `03/04/2024` in one convention from the other.
   * Use a custom format for output people read, not for output hucre
   * reads back.
   */
  dateFormat?: string
  /**
   * Null/undefined representation. Default: ""
   *
   * **One-way.** CSV has no null, so nothing on the read side turns the
   * token back into `null` — `nullValue: "NULL"` reads as the string
   * `"NULL"`, and the default reads as `""`. That is true of the default
   * too, which is why there is no inverse option: restoring `null` for
   * `""` would have to guess for every empty field in the file.
   */
  nullValue?: string
  /**
   * Escape formula injection by prefixing cells starting with =, +, -, @, \t, \r with a single quote. Default: false
   *
   * Reverse it on the way back in with
   * {@link CsvReadOptions.unescapeFormulae}; without that, the `'` is
   * part of the value and every round trip keeps it (#408).
   */
  escapeFormulae?: boolean
  /**
   * Comment character used by the reader this output is written for.
   * Values starting with it are quoted, so the reader keeps them as data
   * instead of discarding the line. Default: unset — a value starting with
   * `#` is written bare, and a reader with `comment: "#"` drops the row.
   *
   * No effect under `quoteStyle: "none"`, which cannot quote anything.
   */
  comment?: string
  /** Column keys to include (for writeCsvObjects). When provided, only these columns are output in this order. */
  columns?: string[]
}

// ── Schema Validation ──────────────────────────────────────────────

export type SchemaFieldType = "string" | "number" | "integer" | "boolean" | "date"

export interface SchemaField {
  /** Expected column header name (for matching) */
  column?: string
  /** Column index (0-based, alternative to column name) */
  columnIndex?: number
  type?: SchemaFieldType
  required?: boolean
  /** Custom validation function */
  validate?: (value: unknown) => boolean | string
  /** Transform value after parsing */
  transform?: (value: unknown) => unknown
  /** Regular expression pattern (for strings) */
  pattern?: RegExp
  /** Minimum value (for numbers) or length (for strings) */
  min?: number
  /** Maximum value (for numbers) or length (for strings) */
  max?: number
  /** Allowed values */
  enum?: unknown[]
  /** Default value for empty cells */
  default?: unknown
}

export type SchemaDefinition = Record<string, SchemaField>

/**
 * One row/column schema failure produced by `validateWithSchema`.
 *
 * Named `SchemaValidationIssue` rather than `ValidationError` because the
 * `ValidationError` *class* (see `./errors`) is what strict mode throws;
 * this is the plain record collected in non-strict mode.
 */
export interface SchemaValidationIssue {
  /** 1-based row number */
  row: number
  /** Column name or index */
  column: string | number
  /** Error message */
  message: string
  /** The raw value that failed validation */
  value: unknown
  /** Field name in the schema */
  field: string
}

// ── Streaming ──────────────────────────────────────────────────────

/**
 * One row yielded by a streaming reader.
 *
 * Shared by `streamXlsxRows` and `streamOdsRows`, which previously had
 * two near-identical shapes under two names. `index` is carried because
 * sheet rows are sparse — a file may jump from row 1 to row 500, and
 * position in the iteration cannot recover that.
 *
 * `streamCsvRows` deliberately yields a bare `CellValue[]` instead: CSV
 * rows are dense and positional, so an index would be pure ceremony, and
 * the bare array is what keeps it the streaming mirror of `parseCsv`.
 */
/**
 * One row from a streaming reader — the same shape from every one of
 * them. v1 had four: `StreamRow` from XLSX and ODS, a bare array from
 * CSV, a bare object from NDJSON, and `XmlStreamRow` from XML.
 *
 * `T` is what a row holds: positional `CellValue[]` from the grid
 * formats, a record from NDJSON and XML.
 */
export interface StreamRow<T = CellValue[]> {
  /** 0-based row index within its sheet — the source position, so a gap means a skipped empty row. */
  index: number
  /** 0-based index of the sheet this row came from. `0` for single-sheet formats. */
  sheet: number
  /** The row. */
  values: T
}

// ── Input/Output Types ─────────────────────────────────────────────

export type ReadInput = Uint8Array | ArrayBuffer | ReadableStream<Uint8Array>
export type WriteOutput = Uint8Array

// ── Incremental writers ────────────────────────────────────────────

/**
 * The vocabulary the four incremental writers share — `XlsxStreamWriter`,
 * `CsvStreamWriter`, `NdjsonStreamWriter`, `OdsStreamWriter` — so a
 * format-agnostic export helper can be written once.
 *
 * v1's version said `finish(): string | Promise<Uint8Array>` and carried a
 * `toStream()` that, on three of the four, buffered everything and then
 * handed over one chunk. Both are gone: `finish()` is bytes everywhere,
 * and a writer that streams says so by having `toStream()` itself
 * (`NdjsonStreamWriter` does) rather than the interface promising it.
 */
export interface SpreadsheetStreamWriter {
  /** Append a row of positional values, or cells. */
  addRow(values: CellInput[]): void
  /** Append a row from an object, projected through the writer's columns. */
  addObject(item: Record<string, CellInput>): void
  /**
   * Close the writer and return its output as bytes. The text writers
   * also have `finishText()`, which returns the same output as a string.
   */
  finish(): Promise<Uint8Array>
}

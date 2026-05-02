// ── Cell Value Types ────────────────────────────────────────────────

export type CellValue = string | number | boolean | Date | null;

export type CellType =
  | "string"
  | "number"
  | "boolean"
  | "date"
  | "error"
  | "formula"
  | "richText"
  | "empty";

// ── Color ──────────────────────────────────────────────────────────

export interface Color {
  /** Hex RGB string without '#', e.g. "FF0000" */
  rgb?: string;
  /** Theme color index */
  theme?: number;
  /** Tint applied to theme color (-1.0 to 1.0) */
  tint?: number;
  /** Indexed color (legacy) */
  indexed?: number;
}

// ── Font ───────────────────────────────────────────────────────────

export interface FontStyle {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | "single" | "double" | "singleAccounting" | "doubleAccounting";
  strikethrough?: boolean;
  color?: Color;
  vertAlign?: "superscript" | "subscript";
  family?: number;
  charset?: number;
  scheme?: "major" | "minor" | "none";
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
  | "mediumGray";

export interface PatternFill {
  type: "pattern";
  pattern: FillPattern;
  fgColor?: Color;
  bgColor?: Color;
}

export interface GradientFill {
  type: "gradient";
  degree?: number;
  stops: Array<{ position: number; color: Color }>;
}

export type FillStyle = PatternFill | GradientFill;

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
  | "slantDashDot";

export interface BorderSide {
  style: BorderLineStyle;
  color?: Color;
}

export interface BorderStyle {
  top?: BorderSide;
  right?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  diagonal?: BorderSide;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
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
    | "general";
  vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
  wrapText?: boolean;
  shrinkToFit?: boolean;
  textRotation?: number;
  indent?: number;
  readingOrder?: "ltr" | "rtl" | "context";
}

// ── Cell Style ─────────────────────────────────────────────────────

export interface CellStyle {
  font?: FontStyle;
  fill?: FillStyle;
  border?: BorderStyle;
  alignment?: AlignmentStyle;
  numFmt?: string;
  protection?: CellProtection;
}

export interface CellProtection {
  locked?: boolean;
  hidden?: boolean;
}

// ── Rich Text ──────────────────────────────────────────────────────

export interface RichTextRun {
  text: string;
  font?: FontStyle;
}

// ── Hyperlink ──────────────────────────────────────────────────────

export interface Hyperlink {
  target: string;
  tooltip?: string;
  display?: string;
  /** Internal reference (e.g. "Sheet2!A1") */
  location?: string;
}

// ── Comment ────────────────────────────────────────────────────────

export interface CellComment {
  author?: string;
  text: string;
  richText?: RichTextRun[];
}

// ── Cell ───────────────────────────────────────────────────────────

export interface Cell {
  value: CellValue;
  type: CellType;
  style?: CellStyle;
  /**
   * Render this cell as an Excel 2024 native checkbox. Only meaningful for
   * boolean cells; the value drives the checked state.
   *
   * Implemented via Microsoft's FeaturePropertyBag extension to OOXML
   * (the `{C7286773-470A-42A8-94C5-96B5CB345126}` cell-XF complement).
   * Requires Microsoft 365; older Excel and LibreOffice fall back to the
   * raw `TRUE`/`FALSE` value.
   */
  checkbox?: boolean;
  formula?: string;
  formulaResult?: CellValue;
  /** Formula type: "shared" | "array". Undefined means normal formula. */
  formulaType?: "shared" | "array";
  /** Shared formula index (si attribute) */
  formulaSharedIndex?: number;
  /** Range this formula applies to (ref attribute on master cell) */
  formulaRef?: string;
  /** Dynamic array flag (cm="1") */
  formulaDynamic?: boolean;
  richText?: RichTextRun[];
  hyperlink?: Hyperlink;
  comment?: CellComment;
}

// ── Column Definition ──────────────────────────────────────────────

export interface ColumnDef {
  /** Column header text */
  header?: string;
  /** Key for object-based data */
  key?: string;
  /** Column width in characters */
  width?: number;
  /** Auto-calculate optimal width from cell content */
  autoWidth?: boolean;
  /** Default style for the column */
  style?: CellStyle;
  /** Number format */
  numFmt?: string;
  /** Hide column */
  hidden?: boolean;
  /** Outline level (grouping) */
  outlineLevel?: number;
  /** Whether this outline group is collapsed */
  collapsed?: boolean;
}

// ── Merge Range ────────────────────────────────────────────────────

export interface MergeRange {
  /** Start row (0-based) */
  startRow: number;
  /** Start column (0-based) */
  startCol: number;
  /** End row (0-based, inclusive) */
  endRow: number;
  /** End column (0-based, inclusive) */
  endCol: number;
}

// ── Data Validation ────────────────────────────────────────────────

export type ValidationType =
  | "list"
  | "whole"
  | "decimal"
  | "date"
  | "time"
  | "textLength"
  | "custom";

export type ValidationOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "lessThan"
  | "greaterThanOrEqual"
  | "lessThanOrEqual";

export interface DataValidation {
  type: ValidationType;
  operator?: ValidationOperator;
  formula1?: string;
  formula2?: string;
  /** List values (for type: "list") */
  values?: string[];
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  inputTitle?: string;
  inputMessage?: string;
  errorTitle?: string;
  errorMessage?: string;
  errorStyle?: "stop" | "warning" | "information";
  /** Cell range (e.g. "A1:A100") */
  range: string;
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
  | "notContainsBlanks";

export interface ConditionalRule {
  type: ConditionalRuleType;
  priority: number;
  operator?: ValidationOperator;
  formula?: string | string[];
  style?: CellStyle;
  stopIfTrue?: boolean;
  range: string;
  /** Color scale configuration */
  colorScale?: {
    cfvo: Array<{
      type: "min" | "max" | "num" | "percent" | "percentile";
      value?: string;
    }>;
    colors: string[]; // hex ARGB colors like "FF63BE7B"
  };
  /** Data bar configuration */
  dataBar?: {
    cfvo: Array<{
      type: "min" | "max" | "num" | "percent" | "percentile";
      value?: string;
    }>;
    color: string;
  };
  /** Icon set configuration */
  iconSet?: {
    iconSet: string; // "3Arrows", "3TrafficLights1", etc.
    cfvo: Array<{
      type: "min" | "num" | "percent" | "percentile";
      value?: string;
    }>;
    reverse?: boolean;
    showValue?: boolean;
  };
  /** Text value for containsText, notContainsText, beginsWith, endsWith */
  text?: string;
}

// ── Auto Filter ────────────────────────────────────────────────────

export interface AutoFilter {
  /** Range (e.g. "A1:D100") */
  range: string;
  /** Column filter criteria */
  columns?: Array<{
    /** 0-based column index within the autoFilter range */
    colIndex: number;
    /** List of values to filter by */
    filters?: string[];
  }>;
}

// ── Freeze Pane ────────────────────────────────────────────────────

export interface FreezePane {
  /** Number of rows to freeze from top */
  rows?: number;
  /** Number of columns to freeze from left */
  columns?: number;
}

// ── Split Pane ─────────────────────────────────────────────────────

export interface SplitPane {
  /** Horizontal split position in twips (1/20 of a point) */
  xSplit?: number;
  /** Vertical split position in twips (1/20 of a point) */
  ySplit?: number;
}

// ── Named Range ────────────────────────────────────────────────────

export interface NamedRange {
  name: string;
  /** Cell range reference (e.g. "Sheet1!$A$1:$D$10") */
  range: string;
  /** Scope: undefined = workbook level, string = sheet name */
  scope?: string;
  comment?: string;
}

// ── Page Setup / Print ─────────────────────────────────────────────

export type PaperSize =
  | "letter"
  | "legal"
  | "a3"
  | "a4"
  | "a5"
  | "b4"
  | "b5"
  | "executive"
  | "tabloid";

export interface PageSetup {
  paperSize?: PaperSize;
  orientation?: "portrait" | "landscape";
  fitToPage?: boolean;
  fitToWidth?: number;
  fitToHeight?: number;
  scale?: number;
  margins?: PageMargins;
  printArea?: string;
  printTitlesRow?: string;
  printTitlesColumn?: string;
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  horizontalCentered?: boolean;
  verticalCentered?: boolean;
}

export interface PageMargins {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
  header?: number;
  footer?: number;
}

export interface HeaderFooter {
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
  differentOddEven?: boolean;
  differentFirst?: boolean;
}

// ── Sparkline ─────────────────────────────────────────────────────

export interface Sparkline {
  /** Cell where the sparkline is displayed */
  location: string;
  /** Data range (e.g. "Sheet1!B2:F2") */
  dataRange: string;
  /** Type: line, column, or win/loss (stacked) */
  type?: "line" | "column" | "stacked";
  /** Color (hex RGB without '#', e.g. "376092") */
  color?: string;
  /** Show markers */
  markers?: boolean;
}

// ── TextBox ───────────────────────────────────────────────────────

export interface SheetTextBox {
  text: string;
  anchor: {
    from: { row: number; col: number };
    to?: { row: number; col: number };
  };
  width?: number;
  height?: number;
  style?: {
    fontSize?: number;
    bold?: boolean;
    color?: string;
    fillColor?: string;
    borderColor?: string;
  };
  /** Alternative text for screen readers (lands in xdr:cNvPr/@descr). */
  altText?: string;
  /** Title/caption for the shape (lands in xdr:cNvPr/@title). */
  title?: string;
}

// ── Threaded Comments (Excel 365+) ─────────────────────────────────

/**
 * A person who can author or be mentioned in threaded comments.
 * Stored in the workbook-wide `xl/persons/person.xml` part.
 */
export interface ThreadedCommentPerson {
  /** Stable GUID identifying this person within the workbook. */
  id: string;
  /** Display name shown in Excel's comment pane (required by the schema). */
  displayName: string;
  /** Identity-system user id, e.g. the Azure AD object id. */
  userId?: string;
  /** Identity provider name, e.g. "AD" or "PeoplePicker". */
  providerId?: string;
}

/**
 * An `@person` mention inside a threaded comment's text. Indices are
 * UTF-16 code-unit offsets into the comment text.
 */
export interface ThreadedCommentMention {
  mentionPersonId: string;
  mentionId: string;
  startIndex: number;
  length: number;
}

/**
 * A single message in a thread on `xl/threadedComments/threadedCommentN.xml`.
 * Top-level messages declare a `ref`; replies omit it and link to their
 * parent through `parentId`.
 */
export interface ThreadedComment {
  id: string;
  /** A1-style cell ref. Required for thread roots, omitted for replies. */
  ref?: string;
  /** GUID matching a {@link ThreadedCommentPerson.id}. */
  personId: string;
  /** GUID of the parent comment when this is a reply. */
  parentId?: string;
  /** ISO-8601 timestamp from the `dT` attribute. */
  date?: string;
  /** Comment body. */
  text: string;
  /** Whether the thread is marked resolved. */
  done?: boolean;
  /** `@person` mentions inside the text. */
  mentions?: ThreadedCommentMention[];
}

// ── Image ──────────────────────────────────────────────────────────

export interface SheetImage {
  data: Uint8Array;
  type: "png" | "jpeg" | "gif" | "svg" | "webp";
  /** Anchor to cell */
  anchor: {
    from: { row: number; col: number };
    to?: { row: number; col: number };
  };
  width?: number;
  height?: number;
  /** Alternative text for screen readers (lands in xdr:cNvPr/@descr). */
  altText?: string;
  /** Title/caption for the image (lands in xdr:cNvPr/@title). */
  title?: string;
}

// ── Charts ─────────────────────────────────────────────────────────

/**
 * Chart kinds supported by {@link writeXlsx} when authoring charts via
 * {@link WriteSheet.charts}. Phase 1 covers the four most common chart
 * families — bar/column, line, pie, and scatter — plus area as a
 * thin variant of `line`.
 *
 * Distinct from the read-side {@link ChartKind} (which mirrors the
 * full set of OOXML chart-type element local names) — the write side
 * exposes only the kinds the chart author can emit today.
 */
export type WriteChartKind = "bar" | "column" | "line" | "pie" | "scatter" | "area";

/**
 * A single data series inside a chart.
 *
 * `values` and `categories` are A1-style cell range references.
 * Provide either a sheet-qualified reference (e.g. `"Sheet1!$B$2:$B$4"`)
 * or a bare range (`"B2:B4"`). Bare ranges are auto-qualified with the
 * sheet that owns the chart.
 */
export interface ChartSeries {
  /** Series name shown in the legend (e.g. "Revenue"). */
  name?: string;
  /** A1-style range with the series numeric values (e.g. "B2:B10"). */
  values: string;
  /** A1-style range with the category labels (e.g. "A2:A10"). */
  categories?: string;
  /** Optional fill color as a 6-digit RGB hex string (e.g. "1F77B4"). */
  color?: string;
}

/**
 * A chart embedded into a worksheet via the drawing layer.
 *
 * Excel anchors charts to cells using the same `xdr:twoCellAnchor`
 * mechanism it uses for images. The chart is stored in
 * `xl/charts/chartN.xml` and wired into the worksheet through a
 * drawing part.
 */
export interface SheetChart {
  /**
   * Chart family. `"bar"` is horizontal, `"column"` is vertical (the
   * Excel default). Both map to `<c:barChart>` with different
   * `<c:barDir>` values.
   */
  type: WriteChartKind;
  /** Optional chart title rendered above the plot area. */
  title?: string;
  /** One or more data series. */
  series: ChartSeries[];
  /** Cell anchor — `to` defaults to a 6×15 area below `from`. */
  anchor: {
    from: { row: number; col: number };
    to?: { row: number; col: number };
  };
  /**
   * Bar/column subtype. Default: `"clustered"`. `"stacked"` and
   * `"percentStacked"` group series end-to-end. Ignored for non-bar
   * chart kinds.
   */
  barGrouping?: "clustered" | "stacked" | "percentStacked";
  /**
   * Whether the legend is shown and where. Default: `"right"` for
   * pie/bar/line/area, `"bottom"` for scatter. Pass `false` to hide
   * the legend.
   */
  legend?: false | "top" | "bottom" | "left" | "right" | "topRight";
  /** Show the chart-level title element. Default: `true` when `title` is set. */
  showTitle?: boolean;
  /** Alternative text for screen readers (lands in xdr:cNvPr/@descr). */
  altText?: string;
  /** Caption for the chart frame (lands in xdr:cNvPr/@title). */
  frameTitle?: string;
}

// ── Accessibility ──────────────────────────────────────────────────

/**
 * Per-sheet accessibility metadata. Hints to screen readers and
 * input to {@link audit} from the `hucre/a11y` entry point.
 */
export interface SheetA11y {
  /**
   * Short, human-readable summary of the sheet's purpose. If the
   * workbook does not already declare a `properties.description`,
   * the first non-empty summary across the workbook is copied there
   * so screen readers announce it when the file is opened.
   */
  summary?: string;
  /**
   * 0-based row index that should be treated as the column-header
   * row. Used by the audit to verify a header is present and to
   * cross-check tables that span the same range.
   */
  headerRow?: number;
}

/** Severity of an accessibility finding. */
export type A11ySeverity = "error" | "warning" | "info";

/** Stable code identifying an accessibility issue. */
export type A11yCode =
  | "no-doc-title"
  | "no-doc-description"
  | "no-header-row"
  | "missing-alt-text"
  | "merged-header-row"
  | "low-contrast"
  | "empty-sheet"
  | "blank-row-in-data";

/** Pinpoint where an issue applies. */
export interface A11yLocation {
  sheet?: string;
  /** Cell reference like "B5" or range like "A1:D1". */
  ref?: string;
  /** Image index inside `sheet.images`. */
  image?: number;
  /** Text-box index inside `sheet.textBoxes`. */
  textBox?: number;
}

export interface A11yIssue {
  type: A11ySeverity;
  code: A11yCode;
  message: string;
  location?: A11yLocation;
}

// ── Sheet Protection ───────────────────────────────────────────────

export interface SheetProtection {
  password?: string;
  sheet?: boolean;
  objects?: boolean;
  scenarios?: boolean;
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  insertHyperlinks?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  pivotTables?: boolean;
}

// ── Sheet View ─────────────────────────────────────────────────────

export interface SheetView {
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  zoomScale?: number;
  rightToLeft?: boolean;
  tabColor?: Color;
}

// ── Table (ListObject) ────────────────────────────────────────────

export interface TableDefinition {
  /** Table name (must be unique in workbook, used in structured references) */
  name: string;
  /** Display name */
  displayName?: string;
  /** Cell range (e.g. "A1:D10") — if not provided, auto-calculated from data */
  range?: string;
  /** Column definitions */
  columns: TableColumn[];
  /** Table style name (e.g. "TableStyleMedium2") */
  style?: string;
  /** Show banded rows. Default: true */
  showRowStripes?: boolean;
  /** Show banded columns. Default: false */
  showColumnStripes?: boolean;
  /** Show auto-filter. Default: true */
  showAutoFilter?: boolean;
  /** Show total row. Default: false */
  showTotalRow?: boolean;
}

export interface TableColumn {
  /** Column header name */
  name: string;
  /** Total row function (sum, count, average, min, max, countNums, stdDev, var, custom) */
  totalFunction?: string;
  /** Total row formula (for custom) */
  totalFormula?: string;
  /** Total row label (text in total cell) */
  totalLabel?: string;
}

// ── Row Definition ────────────────────────────────────────────────

export interface RowDef {
  /** Row height in points */
  height?: number;
  /** Hide row */
  hidden?: boolean;
  /** Outline level (grouping) */
  outlineLevel?: number;
  /** Whether this outline group is collapsed */
  collapsed?: boolean;
}

// ── Sheet ──────────────────────────────────────────────────────────

export interface Sheet {
  name: string;
  rows: CellValue[][];
  /** Detailed cell data (keyed by "row,col" e.g. "0,2") */
  cells?: Map<string, Cell>;
  columns?: ColumnDef[];
  /** Row-level properties (keyed by 0-based row index) */
  rowDefs?: Map<number, RowDef>;
  merges?: MergeRange[];
  dataValidations?: DataValidation[];
  conditionalRules?: ConditionalRule[];
  autoFilter?: AutoFilter;
  freezePane?: FreezePane;
  splitPane?: SplitPane;
  images?: SheetImage[];
  protection?: SheetProtection;
  pageSetup?: PageSetup;
  headerFooter?: HeaderFooter;
  view?: SheetView;
  hidden?: boolean;
  /** Very hidden (only unhideable via VBA) */
  veryHidden?: boolean;
  /** Excel Tables (ListObject) defined on this sheet */
  tables?: TableDefinition[];
  /** Row page breaks (0-based row indices) */
  rowBreaks?: number[];
  /** Column page breaks (0-based column indices) */
  colBreaks?: number[];
  /** Outline properties (controls summary row/column position) */
  outlineProperties?: OutlineProperties;
  /** Background image data (extracted from worksheet picture relationship) */
  backgroundImage?: Uint8Array;
  /** Sparklines (mini-charts in cells) */
  sparklines?: Sparkline[];
  /** Text boxes (shapes with text) */
  textBoxes?: SheetTextBox[];
  /**
   * Excel 365 threaded comments for this sheet. Stored physically in
   * `xl/threadedComments/threadedCommentN.xml` and resolved against
   * the workbook-wide person list (`Workbook.persons`).
   */
  threadedComments?: ThreadedComment[];
  /** Accessibility metadata for screen readers and the `audit` helper. */
  a11y?: SheetA11y;
  /**
   * Pivot table instances hosted on this sheet. The body lives in
   * `xl/pivotTables/pivotTableN.xml`; each instance points at a
   * workbook-level cache via `cacheId`.
   */
  pivotTables?: PivotTable[];
  /**
   * Slicers attached to this sheet (Excel 2010+). Resolved from
   * `xl/slicers/slicerN.xml` parts referenced via this sheet's rels.
   */
  slicers?: Slicer[];
  /**
   * Timeline slicers attached to this sheet (Excel 2013+). Resolved from
   * `xl/timelines/timelineN.xml` parts referenced via this sheet's rels.
   */
  timelines?: Timeline[];
  /**
   * Charts anchored on this sheet, resolved from `xl/charts/chartN.xml`
   * parts referenced via the sheet's drawing. Hucre does not yet author
   * charts; the entries surface for inspection on read and survive
   * roundtrip when the sheet has no hucre-managed images.
   */
  charts?: Chart[];
}

// ── Workbook Properties ────────────────────────────────────────────

export interface WorkbookProperties {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  created?: Date;
  modified?: Date;
  company?: string;
  manager?: string;
  category?: string;
  /** Custom properties */
  custom?: Record<string, string | number | boolean | Date>;
}

// ── External Workbook Links ────────────────────────────────────────

/** Cached cell type as encoded in `cell/@t`. Mirrors OOXML cell type codes. */
export type ExternalCellType = "n" | "s" | "b" | "e" | "str";

export interface ExternalCachedCell {
  /** A1-style reference within the external sheet. */
  ref: string;
  type: ExternalCellType;
  /** Cached value. Strings include error text for `t="e"`. */
  value: string | number | boolean;
}

export interface ExternalSheetData {
  /** 0-based index into the external workbook's sheet list. */
  sheetId: number;
  cells: ExternalCachedCell[];
}

export interface ExternalDefinedName {
  name: string;
  refersTo?: string;
  /** Sheet-local index when present; omitted for workbook-level names. */
  sheetId?: number;
}

/**
 * A reference to another workbook resolved via
 * `xl/externalLinks/externalLinkN.xml`. Cached values follow Excel's
 * formula syntax `[N]Sheet!Ref`, where `N` is this entry's 1-based
 * position in `Workbook.externalLinks`.
 */
export interface ExternalLink {
  /** Target path of the linked workbook (URL, file path, or local entry). */
  target: string;
  /** Almost always `"External"`. Mirrors the `TargetMode` attribute. */
  targetMode?: "External" | "Internal";
  /** External workbook's sheets in declaration order. */
  sheetNames: string[];
  /** Cached cell values, keyed by external sheet id. */
  sheetData: ExternalSheetData[];
  /** Defined names declared in the external workbook. */
  definedNames?: ExternalDefinedName[];
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
  id: string;
  /** Image binary, extracted from the package media folder. */
  data: Uint8Array;
  /** Image format inferred from the media file extension. */
  type: SheetImage["type"];
  /** Optional human-readable description (`descr` attribute). */
  description?: string;
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
  | "varp";

/**
 * Field role in a pivot table layout. `row`, `col`, `page`, and `data`
 * mirror the four standard axes; `hidden` means the field exists in the
 * cache but is not currently placed on any axis.
 */
export type PivotFieldAxis = "row" | "col" | "page" | "data" | "hidden";

export interface PivotField {
  /**
   * Display name. Reads from the `<cacheField name="...">` attribute on
   * the matching field index in the pivot cache definition.
   */
  name: string;
  /**
   * Where the field appears in the pivot table. `hidden` covers cache
   * fields that are present but not placed on any axis.
   */
  axis: PivotFieldAxis;
  /** When `axis === "data"`, the aggregation applied to the values. */
  function?: PivotDataFieldFunction;
  /**
   * Display name overlay for data fields (the `name` attribute on
   * `<dataField>`). Falls back to `name` when absent.
   */
  displayName?: string;
}

/**
 * A pivot table instance, attached to the sheet that hosts its layout.
 * The `cacheId` references one of the workbook-level pivot caches that
 * back this table.
 */
export interface PivotTable {
  /** Pivot table name (`<pivotTableDefinition name="...">`). */
  name: string;
  /**
   * Index into `Workbook.pivotCaches`. Mirrors the workbook-level
   * `cacheId` attribute on `<pivotCache>` rather than the per-table
   * relationship — that way a model author who reorders the cache
   * array keeps the link sound.
   */
  cacheId: number;
  /**
   * Output range on the host sheet, e.g. `"A3:D20"`. Empty string when
   * the source omits a `<location>` element.
   */
  location: string;
  /** Number of header rows above the data rows. */
  firstHeaderRow?: number;
  /** Number of body rows reserved for column-axis labels. */
  firstDataRow?: number;
  /** Column index of the first data row (0-based). */
  firstDataCol?: number;
  /** Number of pages declared in `<pageFields>`. */
  rowPageCount?: number;
  /** Number of column-axis page-break positions. */
  colPageCount?: number;
  /**
   * Pivot fields in declaration order. The position in this array is
   * the field index used by `<rowItems>`, `<colItems>`, etc.
   */
  fields: PivotField[];
  /** Pivot-table style name (`<pivotTableStyleInfo name="...">`). */
  styleName?: string;
  /** Whether the data field caption is shown. */
  dataCaption?: string;
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
  cacheId: number;
  /**
   * Source range, e.g. `"Sheet1!$A$1:$C$100"` or a defined-name
   * reference. Empty string for non-worksheet sources.
   */
  sourceRef?: string;
  /** Source sheet name when the source is a worksheet range. */
  sourceSheet?: string;
  /**
   * Source type: `worksheet` (range or table on a sheet), `external`
   * (linked workbook / database), `consolidation`, or `scenario`. Most
   * real workbooks use `worksheet`.
   */
  sourceType?: "worksheet" | "external" | "consolidation" | "scenario";
  /** Cached field names in declaration order. */
  fieldNames: string[];
  /** Whether a `pivotCacheRecords{N}.xml` part is present. */
  hasRecords?: boolean;
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
  field: string;
  /** Aggregation function. Default: `"sum"`. */
  function?: PivotDataFieldFunction;
  /** Optional display name override. Default: e.g. `"Sum of Revenue"`. */
  displayName?: string;
  /** Optional number format for aggregated values. Default: General. */
  numberFormat?: string;
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
  name: string;
  /**
   * Source sheet name. Defaults to the sheet the pivot is declared on
   * when omitted — handy for pivots that summarise their own sheet's
   * data.
   */
  sourceSheet?: string;
  /**
   * Source range covering the header row plus all data rows
   * (e.g. `"A1:C100"`). Auto-detected from the source sheet's `rows`
   * length when omitted.
   */
  sourceRange?: string;
  /**
   * Top-left anchor for the rendered pivot table on the host sheet
   * (e.g. `"A3"`). Default: `"A1"`.
   */
  targetCell?: string;
  /** Source columns laid out on the row axis, in order. */
  rows?: string[];
  /** Source columns laid out on the column axis, in order. */
  columns?: string[];
  /** Source columns laid out as page (filter) fields, in order. */
  pages?: string[];
  /** Aggregated value fields. Each entry becomes one data column. */
  values: WritePivotDataField[];
  /**
   * Pivot table style name (e.g. `"PivotStyleLight16"`). Default:
   * `"PivotStyleLight16"` — the modern Excel default.
   */
  styleName?: string;
  /**
   * Caption shown above the data fields when there is more than one.
   * Default: `"Values"` (Excel's built-in caption).
   */
  dataCaption?: string;
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
  name: string;
  /** Slicer cache identifier this slicer references. Mirrors `slicer/@cache`. */
  cache: string;
  /** Display caption shown in the header. Mirrors `slicer/@caption`. */
  caption?: string;
  /** Number of columns in the slicer button grid. Mirrors `slicer/@columnCount`. */
  columnCount?: number;
  /** Built-in style id, e.g. `SlicerStyleLight1`. Mirrors `slicer/@style`. */
  style?: string;
  /** Sort order for items. Mirrors `slicer/@sortOrder` (e.g. `ascending`, `descending`). */
  sortOrder?: string;
  /** Row height in EMUs. Mirrors `slicer/@rowHeight`. */
  rowHeight?: number;
}

/**
 * Workbook-level slicer cache. Stores the filter source and selection
 * state shared by one or more {@link Slicer} instances.
 *
 * Slicer caches come from `xl/slicerCaches/slicerCacheN.xml`.
 */
export interface SlicerCache {
  /** Programmatic name. Mirrors `slicerCacheDefinition/@name`. */
  name: string;
  /** Source identifier — typically the cache definition's source ref. */
  sourceName?: string;
  /**
   * Pivot tables this cache filters, when sourced from a pivot table.
   * Each entry is the `tabId` (sheet index) + `name` of a pivot table.
   */
  pivotTables?: SlicerCachePivotTable[];
  /** Excel Table this cache filters, when sourced from a table. */
  tableSource?: SlicerCacheTableSource;
}

export interface SlicerCachePivotTable {
  /** 0-based sheet tab id of the sheet hosting the pivot table. */
  tabId: number;
  /** Pivot table name. */
  name: string;
}

export interface SlicerCacheTableSource {
  /** Excel Table name. */
  name: string;
  /** Column referenced in the table. */
  column?: string;
}

/**
 * Timeline slicer (Excel 2013+ date-range filter). Like {@link Slicer}
 * but constrained to date columns and rendered as a draggable date band.
 *
 * Timelines come from `xl/timelines/timelineN.xml`.
 */
export interface Timeline {
  /** Programmatic name. */
  name: string;
  /** Cache identifier this timeline references. */
  cache: string;
  /** Display caption. */
  caption?: string;
  /** Built-in style id, e.g. `TimeSlicerStyleLight1`. */
  style?: string;
  /** Granularity: `years`, `quarters`, `months`, or `days`. */
  level?: string;
  /** Whether the time-level selector is shown. */
  showHeader?: boolean;
  /** Whether the selection-label band is shown. */
  showSelectionLabel?: boolean;
  /** Whether the time-level row is shown. */
  showTimeLevel?: boolean;
  /** Whether the horizontal scrollbar is shown. */
  showHorizontalScrollbar?: boolean;
}

/**
 * Workbook-level timeline cache. Stores the date column and selected
 * range shared by one or more {@link Timeline} instances.
 *
 * Timeline caches come from `xl/timelineCaches/timelineCacheN.xml`.
 */
export interface TimelineCache {
  /** Programmatic name. */
  name: string;
  /** Source identifier. */
  sourceName?: string;
  /** Pivot tables this cache filters. */
  pivotTables?: SlicerCachePivotTable[];
}

// ── Charts ────────────────────────────────────────────────────────

/**
 * Chart kind reported by {@link Chart.kinds}. Mirrors the OOXML
 * chart-type element local names (`c:barChart`, `c:lineChart`, ...).
 * A single chart can mix multiple kinds (combo chart), in which case
 * every kind appears in `kinds` in the order it's declared.
 */
export type ChartKind =
  | "bar"
  | "bar3D"
  | "line"
  | "line3D"
  | "pie"
  | "pie3D"
  | "doughnut"
  | "area"
  | "area3D"
  | "scatter"
  | "bubble"
  | "radar"
  | "surface"
  | "surface3D"
  | "stock"
  | "ofPie";

/**
 * A single series surfaced from a parsed chart.
 *
 * Field semantics mirror what {@link ChartSeries} accepts on the write
 * side, so a `ChartSeriesInfo` returned by {@link Chart.series} can be
 * used as the basis for cloning a chart with new bindings.
 *
 * `valuesRef` and `categoriesRef` are the raw `<c:f>` formula strings
 * extracted from the chart XML — typically sheet-qualified A1 ranges
 * like `"Sheet1!$B$2:$B$10"`. They may be `undefined` when the series
 * embeds literal numbers (`<c:numLit>`) instead of referencing a range.
 */
export interface ChartSeriesInfo {
  /** Chart kind that owns this series (matches {@link Chart.kinds}). */
  kind: ChartKind;
  /** 0-based position inside the chart-type element. */
  index: number;
  /** Series name pulled from `<c:tx>` (literal `<c:v>` or strRef cache). */
  name?: string;
  /** Raw `<c:f>` for `<c:val>` / `<c:yVal>`. */
  valuesRef?: string;
  /** Raw `<c:f>` for `<c:cat>` / `<c:xVal>`. */
  categoriesRef?: string;
  /** 6-digit RGB hex from `<c:spPr><a:solidFill><a:srgbClr val>`. */
  color?: string;
}

/**
 * A chart anchored on a sheet via the sheet's drawing part.
 *
 * Charts come from `xl/charts/chartN.xml`. Hucre exposes the
 * structural metadata needed to recognize, introspect, and clone the
 * chart; the chart body is preserved verbatim through roundtrip.
 */
/**
 * Legend placement reported by {@link Chart.legend}.
 *
 * Values mirror the {@link SheetChart.legend} options on the writer
 * side, so a parsed legend position slots straight back into a clone
 * target. `false` is reported when the chart explicitly omits the
 * legend element (Excel's "no legend" state); `undefined` means the
 * chart did not declare a legend at all.
 */
export type ChartLegendPosition = "top" | "bottom" | "left" | "right" | "topRight";

/**
 * Bar/column grouping reported by {@link Chart.barGrouping}.
 *
 * Pulled from `<c:barChart><c:grouping val="..."/></c:barChart>`.
 * `"standard"` is the OOXML value for non-stacked, non-percent layouts
 * — it is excluded here because the writer's
 * {@link SheetChart.barGrouping} models the same default as the
 * absence of the field. Only the stacked variants surface, which is
 * what callers need to detect when cloning a stacked template.
 */
export type ChartBarGrouping = "clustered" | "stacked" | "percentStacked";

export interface Chart {
  /** Chart-type elements present in `<c:plotArea>`, in declaration order. */
  kinds: ChartKind[];
  /** Number of `<c:ser>` series across every chart-type element. */
  seriesCount: number;
  /** Plain-text title pulled from `<c:title>`, when present. */
  title?: string;
  /**
   * Per-series metadata across every chart-type element, in
   * declaration order. Empty when the chart has no `<c:ser>` children.
   */
  series?: ChartSeriesInfo[];
  /**
   * Legend placement pulled from `<c:legend><c:legendPos val=".."/>`.
   * Reported as `false` when the chart explicitly omits the legend
   * element (Excel's "no legend" state). `undefined` means the chart
   * did not declare a legend at all — Excel falls back to its default
   * placement in that case.
   */
  legend?: false | ChartLegendPosition;
  /**
   * Grouping pulled from the first `<c:barChart>` element, when the
   * chart has one. Surfaces only the stacked variants — the OOXML
   * `"standard"` / `"clustered"` values both round-trip cleanly to
   * the writer's `"clustered"` default, but only the explicit
   * `clustered` value is reported here for symmetry with the writer's
   * {@link SheetChart.barGrouping} field.
   */
  barGrouping?: ChartBarGrouping;
}

// ── Workbook ───────────────────────────────────────────────────────

export interface Workbook {
  sheets: Sheet[];
  properties?: WorkbookProperties;
  namedRanges?: NamedRange[];
  /** Date system: 1900 (default/Windows) or 1904 (Mac) */
  dateSystem?: "1900" | "1904";
  /** Default font for the workbook */
  defaultFont?: FontStyle;
  /** Active sheet index */
  activeSheet?: number;
  /** Theme color palette (resolved from xl/theme/theme1.xml) */
  themeColors?: string[];
  /** Workbook-level protection */
  workbookProtection?: {
    lockStructure?: boolean;
    lockWindows?: boolean;
  };
  /**
   * Workbook-wide person directory referenced from threaded comments.
   * Each `ThreadedComment.personId` resolves against this list.
   */
  persons?: ThreadedCommentPerson[];
  /**
   * External workbook references, resolved from
   * `xl/externalLinks/externalLinkN.xml`. The 1-based position in this
   * array matches the `[N]` prefix used in formulas like `[1]Sheet1!A1`.
   */
  externalLinks?: ExternalLink[];
  /**
   * Cell-embedded images (WPS DISPIMG mechanism).
   *
   * Resolved from `xl/cellimages.xml`. Cells reference these images via
   * `=_xlfn.DISPIMG("<id>", 1)` formulas — match `CellImage.id` against
   * the first argument to look up the binary.
   */
  cellImages?: CellImage[];
  /**
   * Workbook-level pivot caches resolved from
   * `xl/pivotCache/pivotCacheDefinitionN.xml`. Sheet-level
   * `PivotTable.cacheId` references entries here.
   */
  pivotCaches?: PivotCache[];
  /**
   * Slicer caches resolved from `xl/slicerCaches/slicerCacheN.xml`.
   * The 1-based position in this array matches the `N` in the source path.
   */
  slicerCaches?: SlicerCache[];
  /**
   * Timeline caches resolved from `xl/timelineCaches/timelineCacheN.xml`.
   * The 1-based position in this array matches the `N` in the source path.
   */
  timelineCaches?: TimelineCache[];
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
  name: string;
  /** 0-based position in the workbook's sheet list. */
  index: number;
  /** XLSX `<sheet state="hidden">`. Undefined for ODS. */
  hidden?: boolean;
  /** XLSX `<sheet state="veryHidden">`. Undefined for ODS. */
  veryHidden?: boolean;
}

/**
 * Predicate form of {@link ReadOptions.sheets}. Receives one
 * {@link SheetFilterInfo} per sheet in workbook order; returning `true`
 * includes the sheet, `false` skips it.
 */
export type SheetFilter = (info: SheetFilterInfo, index: number) => boolean;

export interface ReadOptions {
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
  sheets?: Array<number | string> | SheetFilter;
  /** Which row is the header row (1-based). Default: none */
  headerRow?: number;
  /** Schema for validation and type coercion */
  schema?: SchemaDefinition;
  /** Date system override. Default: auto-detect from file */
  dateSystem?: "1900" | "1904" | "auto";
  /** Whether to read styles. Default: false (faster without) */
  readStyles?: boolean;
  /** Password for encrypted files */
  password?: string;
  /** Maximum number of data rows to read per sheet. Default: unlimited */
  maxRows?: number;
  /** Cell range to read (e.g. "A1:D10"). Only cells within this range are returned. */
  range?: string;
}

// ── Write Options ──────────────────────────────────────────────────

export interface WriteOptions {
  sheets: WriteSheet[];
  properties?: WorkbookProperties;
  namedRanges?: NamedRange[];
  defaultFont?: FontStyle;
  dateSystem?: "1900" | "1904";
  /** Active sheet index (0-based). Default: 0 */
  activeSheet?: number;
  /** Workbook-level protection (lock structure/windows) */
  workbookProtection?: {
    lockStructure?: boolean;
    lockWindows?: boolean;
    password?: string;
  };
  /** String storage mode. Default: "shared"
   *  - "shared": shared string table (smaller files with repeated strings)
   *  - "inline": inline strings per cell (faster write, larger files)
   */
  stringMode?: "shared" | "inline";
  /** VBA project binary (vbaProject.bin) to embed. Output becomes macro-enabled (.xlsm). */
  vbaProject?: Uint8Array;
}

export interface WriteSheet {
  name: string;
  columns?: ColumnDef[];
  /** Raw row data (array of arrays) */
  rows?: CellValue[][];
  /** Object data (array of objects — uses column keys) */
  data?: Array<Record<string, CellValue>>;
  /** Detailed cell overrides (keyed by "row,col") */
  cells?: Map<string, Partial<Cell>>;
  merges?: MergeRange[];
  dataValidations?: DataValidation[];
  conditionalRules?: ConditionalRule[];
  autoFilter?: AutoFilter;
  freezePane?: FreezePane;
  splitPane?: SplitPane;
  images?: SheetImage[];
  protection?: SheetProtection;
  pageSetup?: PageSetup;
  headerFooter?: HeaderFooter;
  view?: SheetView;
  hidden?: boolean;
  veryHidden?: boolean;
  /** Excel Tables (ListObject) to define on this sheet */
  tables?: TableDefinition[];
  /** Row page breaks (0-based row indices) */
  rowBreaks?: number[];
  /** Column page breaks (0-based column indices) */
  colBreaks?: number[];
  /** Row-level properties (keyed by 0-based row index) */
  rowDefs?: Map<number, RowDef>;
  /** Outline properties (controls summary row/column position) */
  outlineProperties?: OutlineProperties;
  /** Background image for the worksheet (watermark) */
  backgroundImage?: Uint8Array;
  /** Sparklines (mini-charts in cells) */
  sparklines?: Sparkline[];
  /** Text boxes (shapes with text) */
  textBoxes?: SheetTextBox[];
  /**
   * Native Excel charts (bar, column, line, pie, scatter, area). Charts
   * share the worksheet's drawing part with images and text boxes.
   */
  charts?: SheetChart[];
  /** Excel 365 threaded comments for this sheet. */
  threadedComments?: ThreadedComment[];
  /**
   * Pivot tables anchored on this sheet. The source data is read from
   * either the same sheet or a sibling sheet identified by
   * {@link WritePivotTable.sourceSheet}.
   */
  pivotTables?: WritePivotTable[];
  /** Accessibility metadata for screen readers and the `audit` helper. */
  a11y?: SheetA11y;
}

// ── Outline Properties ────────────────────────────────────────────

export interface OutlineProperties {
  /** Summary rows appear below detail rows. Default: true */
  summaryBelow?: boolean;
  /** Summary columns appear to the right of detail columns. Default: true */
  summaryRight?: boolean;
}

// ── CSV Options ────────────────────────────────────────────────────

export interface CsvReadOptions {
  /** Field delimiter. Default: auto-detect */
  delimiter?: string;
  /** Line separator. Default: auto-detect */
  lineSeparator?: string;
  /** Quote character. Default: " */
  quote?: string;
  /** Escape character. Default: " (RFC 4180 doubled quotes) */
  escape?: string;
  /** Whether first row is header. Default: false */
  header?: boolean;
  /** Skip BOM if present. Default: true */
  skipBom?: boolean;
  /** Type inference for numbers, booleans, dates. Default: false */
  typeInference?: boolean;
  /** Keep strings with leading zeros (e.g. "0123") as strings instead of converting to numbers. Default: true */
  preserveLeadingZeros?: boolean;
  /** Schema for validation */
  schema?: SchemaDefinition;
  /** Encoding. Default: "utf-8" */
  encoding?: string;
  /** Skip empty rows. Default: false */
  skipEmptyRows?: boolean;
  /** Comment character (lines starting with this are skipped) */
  comment?: string;
  /** Maximum number of data rows to parse. When set, parsing stops after this many rows. */
  maxRows?: number;
  /** Skip the first N lines before parsing (useful for files with metadata headers above the CSV data). */
  skipLines?: number;
  /** Called for each row during parsing, enabling progressive processing without buffering all rows. */
  onRow?: (row: CellValue[], index: number) => void;
  /** Transform each header string when `header: true`. Called on each header value. */
  transformHeader?: (header: string, index: number) => string;
  /** Transform each cell value after type inference. Called on every cell. */
  transformValue?: (value: CellValue, header: string, row: number, col: number) => CellValue;
  /** Fast mode: skip quote handling and just split by delimiter/newlines. Faster for files known to have no quoted fields. Default: false */
  fastMode?: boolean;
}

export interface CsvWriteOptions {
  /** Field delimiter. Default: "," */
  delimiter?: string;
  /** Line separator. Default: "\r\n" (CRLF per RFC 4180) */
  lineSeparator?: string;
  /** Quote character. Default: " */
  quote?: string;
  /** Quote style. Default: "required" */
  quoteStyle?: "all" | "required" | "none";
  /** Headers row from column names */
  headers?: string[] | boolean;
  /** Prepend UTF-8 BOM (for Excel compatibility). Default: false */
  bom?: boolean;
  /** Date format string. Default: ISO 8601 */
  dateFormat?: string;
  /** Null/undefined representation. Default: "" */
  nullValue?: string;
  /** Escape formula injection by prefixing cells starting with =, +, -, @, \t, \r with a single quote. Default: false */
  escapeFormulae?: boolean;
  /** Column keys to include (for writeCsvObjects). When provided, only these columns are output in this order. */
  columns?: string[];
}

// ── Schema Validation ──────────────────────────────────────────────

export type SchemaFieldType = "string" | "number" | "integer" | "boolean" | "date";

export interface SchemaField {
  /** Expected column header name (for matching) */
  column?: string;
  /** Column index (0-based, alternative to column name) */
  columnIndex?: number;
  type?: SchemaFieldType;
  required?: boolean;
  /** Custom validation function */
  validate?: (value: unknown) => boolean | string;
  /** Transform value after parsing */
  transform?: (value: unknown) => unknown;
  /** Regular expression pattern (for strings) */
  pattern?: RegExp;
  /** Minimum value (for numbers) or length (for strings) */
  min?: number;
  /** Maximum value (for numbers) or length (for strings) */
  max?: number;
  /** Allowed values */
  enum?: unknown[];
  /** Default value for empty cells */
  default?: unknown;
}

export type SchemaDefinition = Record<string, SchemaField>;

export interface ValidationError {
  /** 1-based row number */
  row: number;
  /** Column name or index */
  column: string | number;
  /** Error message */
  message: string;
  /** The raw value that failed validation */
  value: unknown;
  /** Field name in the schema */
  field: string;
}

export interface ReadResult<T = Record<string, unknown>> {
  /** Parsed and validated rows */
  data: T[];
  /** Validation errors (if schema provided) */
  errors: ValidationError[];
  /** Raw sheet data */
  sheets: Sheet[];
}

// ── Streaming Types ────────────────────────────────────────────────

export interface StreamReadOptions extends ReadOptions {
  /** Batch size for row events. Default: 1 */
  batchSize?: number;
}

export interface StreamWriteOptions extends WriteOptions {
  /** Sheet being written */
  sheet: WriteSheet;
}

// ── Input/Output Types ─────────────────────────────────────────────

export type ReadInput = Uint8Array | ArrayBuffer | ReadableStream<Uint8Array>;
export type WriteOutput = Uint8Array;

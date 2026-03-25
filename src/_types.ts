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
}

// ── Read Options ───────────────────────────────────────────────────

export interface ReadOptions {
  /** Which sheets to read (by index or name). Default: all */
  sheets?: Array<number | string>;
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

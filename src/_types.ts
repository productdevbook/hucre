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
 * {@link WriteSheet.charts}. Covers the most common chart families —
 * bar/column, line, pie, doughnut, scatter, and area.
 *
 * Distinct from the read-side {@link ChartKind} (which mirrors the
 * full set of OOXML chart-type element local names) — the write side
 * exposes only the kinds the chart author can emit today.
 */
export type WriteChartKind = "bar" | "column" | "line" | "pie" | "doughnut" | "scatter" | "area";

/**
 * Where a data label is placed relative to its data point.
 *
 * Mirrors the OOXML `c:dLblPos` value space. Not every chart kind
 * accepts every position — Excel will silently fall back to a sensible
 * default when an invalid combination is requested:
 *
 * - **Bar / column**: `"ctr"`, `"inEnd"`, `"inBase"`, `"outEnd"` (clustered) or `"ctr"`, `"inEnd"`, `"inBase"` (stacked).
 * - **Line / area / scatter**: `"t"`, `"b"`, `"l"`, `"r"`, `"ctr"`.
 * - **Pie / doughnut**: `"ctr"`, `"inEnd"`, `"outEnd"`, `"bestFit"`.
 */
export type ChartDataLabelPosition =
  | "t"
  | "b"
  | "l"
  | "r"
  | "ctr"
  | "inEnd"
  | "inBase"
  | "outEnd"
  | "bestFit";

/**
 * Configuration for the small text annotations Excel paints next to
 * each data point. Maps to the OOXML `<c:dLbls>` element.
 *
 * Apply at the chart level via {@link SheetChart.dataLabels} to label
 * every series, or at the series level via
 * {@link ChartSeries.dataLabels} to override a single series. A
 * series-level `dataLabels` always wins over the chart-level default,
 * including when the value is `false` (which suppresses the labels for
 * that series alone).
 *
 * At least one of `showValue`, `showCategoryName`, `showSeriesName`,
 * or `showPercent` should be `true` for the labels to render anything
 * meaningful — Excel hides the label box when no toggle is on.
 */
export interface ChartDataLabels {
  /** Show the numeric value of each data point. */
  showValue?: boolean;
  /** Show the category (X-axis) label. */
  showCategoryName?: boolean;
  /** Show the series name. Useful with multi-series legends collapsed. */
  showSeriesName?: boolean;
  /** Show the value as a percent of total. Pie / doughnut only. */
  showPercent?: boolean;
  /**
   * Where the label sits relative to its point. See
   * {@link ChartDataLabelPosition} for the valid set per chart kind.
   * Omit to let Excel pick a default (`outEnd` for bar/column,
   * `r` for line/scatter, `bestFit` for pie).
   */
  position?: ChartDataLabelPosition;
  /**
   * Separator between concatenated label parts when more than one
   * `show*` toggle is on. Defaults to `", "`. Common alternatives:
   * `" "`, `"; "`, `"\n"` (newline).
   */
  separator?: string;
}

/**
 * Preset dash pattern for a chart series line stroke.
 *
 * Mirrors the OOXML `ST_PresetLineDashVal` enum exactly. Each value
 * names a stock pattern Excel paints without needing a custom dash
 * array. The Excel "Format Data Series → Line → Dash type" UI exposes
 * these stock patterns; Excel ignores any unrecognized value.
 */
export type ChartLineDashStyle =
  | "solid"
  | "dot"
  | "dash"
  | "lgDash"
  | "dashDot"
  | "lgDashDot"
  | "lgDashDotDot"
  | "sysDash"
  | "sysDot"
  | "sysDashDot"
  | "sysDashDotDot";

/**
 * Per-series line stroke styling for line / scatter charts.
 *
 * Maps to the `<a:ln>` element nested inside `<c:ser><c:spPr>` — the
 * same wrapper that already carries the series fill color. Only
 * meaningful on `line` and `scatter` series; the field is silently
 * dropped on every other chart family at all three layers (read,
 * write, clone), since dashing and stroke width have no visible effect
 * on bar / pie / doughnut / area renderings.
 *
 * Every field is optional — a bare `{}` collapses to no stroke
 * configuration and leaves Excel's per-series default in place. Set
 * `dash: "solid"` to explicitly reset a template's dashed stroke back
 * to a continuous line.
 */
export interface ChartLineStroke {
  /**
   * Preset dash pattern. See {@link ChartLineDashStyle} for the
   * accepted set.
   */
  dash?: ChartLineDashStyle;
  /**
   * Stroke width in points. Excel's UI exposes the 0.25 – 13.5 pt band;
   * the writer clamps anything outside that range and rounds to the
   * nearest quarter-point so a round-trip cannot drift. The OOXML
   * attribute is in EMU (1 pt = 12 700 EMU); the writer performs the
   * conversion and the reader inverts it. Non-finite values are
   * dropped so the writer can elide the attribute entirely.
   */
  width?: number;
}

/**
 * Marker symbol shape rendered at each data point on a line / scatter
 * series.
 *
 * Mirrors the OOXML `ST_MarkerStyle` enum exactly. `"none"` suppresses
 * the marker (the Excel default for line charts beyond the first
 * series); `"auto"` defers to Excel's series-rotation default; every
 * other value pins a specific shape. `"picture"` is intentionally
 * omitted — it requires a separately-embedded picture part that Phase 1
 * native chart authoring does not support.
 */
export type ChartMarkerSymbol =
  | "none"
  | "auto"
  | "circle"
  | "square"
  | "diamond"
  | "triangle"
  | "x"
  | "star"
  | "dot"
  | "dash"
  | "plus";

/**
 * Per-series marker styling for line / scatter charts.
 *
 * Maps to `<c:marker>` inside `<c:ser>`. Only meaningful on `line` and
 * `scatter` series — the OOXML schema places `<c:marker>` exclusively
 * on `CT_LineSer` and `CT_ScatterSer`, so the field is silently
 * dropped on every other chart family at all three layers (read,
 * write, clone).
 *
 * Every field is optional — a bare `{}` collapses to no marker
 * configuration and leaves Excel's per-series default in place. Set
 * `symbol: "none"` to explicitly hide the marker (useful for a
 * scatter clone whose template uses markers but the dashboard wants
 * a clean line).
 */
export interface ChartMarker {
  /** Shape of the marker glyph. See {@link ChartMarkerSymbol}. */
  symbol?: ChartMarkerSymbol;
  /**
   * Marker glyph size in points, in the OOXML range `2..72`. Excel's
   * UI clamps values outside this band. Default (when omitted): Excel
   * picks a series-rotation default (typically `5`).
   */
  size?: number;
  /**
   * Marker fill color as a 6-digit RGB hex string (e.g. `"1F77B4"`).
   * Maps to `<c:marker><c:spPr><a:solidFill><a:srgbClr val="..">`.
   */
  fill?: string;
  /**
   * Marker outline color as a 6-digit RGB hex string. Maps to
   * `<c:marker><c:spPr><a:ln><a:solidFill><a:srgbClr val="..">`.
   */
  line?: string;
}

/**
 * How Excel paints a series across cells whose value is missing or
 * blank. Mirrors the OOXML `ST_DispBlanksAs` enum exactly and matches
 * the three options Excel exposes under "Select Data Source → Hidden
 * and Empty Cells":
 *
 * - `"gap"` — leave a gap at the missing point (the OOXML default and
 *   what Excel selects in fresh chart UI). A line chart shows a break,
 *   a bar chart simply skips the bar.
 * - `"zero"` — substitute `0` for the missing value, so a line chart
 *   drops to the X axis and bar charts render a flush-zero bar.
 * - `"span"` — connect adjacent points across the gap (line / scatter
 *   only; Excel falls back to `"gap"` for bar / pie / area).
 */
export type ChartDisplayBlanksAs = "gap" | "zero" | "span";

/**
 * Scatter sub-style applied at the chart level. Maps to the OOXML
 * `ST_ScatterStyle` enum which sits inside `<c:scatterChart>` as
 * `<c:scatterStyle val=".."/>`. Excel exposes the same six presets
 * under "Change Chart Type → XY (Scatter)":
 *
 * - `"none"`         — markers only, no connecting line and no curves.
 *                      Equivalent to `"marker"` in modern Excel UI.
 * - `"line"`         — straight-line segments between points, no markers.
 * - `"lineMarker"`   — straight-line segments with markers (Excel's
 *                      reference default and the writer's fallback).
 * - `"marker"`       — markers only, no line. Same render as `"none"`;
 *                      OOXML lists both for legacy compatibility.
 * - `"smooth"`       — smoothed (Catmull-Rom-style) curves between
 *                      points, no markers.
 * - `"smoothMarker"` — smoothed curves with markers.
 *
 * Distinct from the per-series {@link ChartSeries.smooth} flag — the
 * series-level toggle paints individual points, while `scatterStyle`
 * is the chart-wide preset Excel selects in the chart-type picker.
 * When both are set, the OOXML schema lets Excel render the union
 * (smooth chart with the series-level smooth still emitted), but
 * Excel's UI normally pairs them: `scatterStyle: "smooth"` implies
 * smoothed series, `scatterStyle: "lineMarker"` implies straight ones.
 */
export type ChartScatterStyle =
  | "none"
  | "line"
  | "lineMarker"
  | "marker"
  | "smooth"
  | "smoothMarker";

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
  /**
   * Per-series data label override. Pass `false` to suppress labels
   * for this series even when the chart-level
   * {@link SheetChart.dataLabels} enables them.
   */
  dataLabels?: ChartDataLabels | false;
  /**
   * Smooth the line connecting data points using a Catmull-Rom-style
   * spline. Maps to `<c:smooth val="..">` inside the `<c:ser>` element.
   * Only meaningful for `line` and `scatter` charts — ignored for every
   * other chart kind (the OOXML schema does not allow `<c:smooth>` on
   * bar / column / pie / doughnut / area series).
   *
   * Default: `false` (straight-line segments). Set `true` to render the
   * curved variant Excel offers under "Format Data Series → Line →
   * Smoothed line".
   */
  smooth?: boolean;
  /**
   * Per-series line stroke (dash pattern + width) for line / scatter
   * charts. Maps to `<a:ln>` inside `<c:ser><c:spPr>`. Ignored on every
   * other chart family — bar / column / pie / doughnut / area never
   * render a connecting line, so dashing and stroke width have no
   * visible effect there. See {@link ChartLineStroke}.
   */
  stroke?: ChartLineStroke;
  /**
   * Per-series marker styling. Only meaningful for `line` and
   * `scatter` charts — the OOXML schema places `<c:marker>` on
   * `CT_LineSer` / `CT_ScatterSer` only. Ignored on every other
   * chart family at write time.
   */
  marker?: ChartMarker;
  /**
   * Invert the fill color when the value is negative. Maps to
   * `<c:invertIfNegative val=".."/>` inside the `<c:ser>` element.
   * Only meaningful for `bar` and `column` charts — the OOXML schema
   * places `<c:invertIfNegative>` exclusively on `CT_BarSer` and
   * `CT_Bar3DSer`, so the field is silently dropped on every other
   * chart family at write time.
   *
   * Default: `false` (negative bars share the series fill color).
   * Set `true` to mirror Excel's "Format Data Series → Fill → Invert
   * if negative" toggle, which paints negative bars with white (or
   * the inverted color when the spreadsheet supplies one).
   */
  invertIfNegative?: boolean;
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
   * Bar/column gap width as a percentage of the bar width — the empty
   * space between adjacent category groups. Accepted range: `0` – `500`
   * (the OOXML `ST_GapAmount` schema). Excel's default is `150` (each
   * group's gap equals 1.5× the bar width). Smaller values pack groups
   * tighter; `0` removes the gap entirely. Maps to
   * `<c:barChart><c:gapWidth val=".."/></c:barChart>`. Ignored for
   * non-bar / non-column chart kinds.
   */
  gapWidth?: number;
  /**
   * Bar/column series overlap as a percentage of the bar width.
   * Accepted range: `-100` – `100` (the OOXML `ST_Overlap` schema).
   * Negative values open a gap between series within a group, positive
   * values stack them on top of each other. Excel's default is `0` for
   * `clustered` (side-by-side) and `100` for `stacked` /
   * `percentStacked` (fully overlapped). Maps to
   * `<c:barChart><c:overlap val=".."/></c:barChart>`. Ignored for
   * non-bar / non-column chart kinds.
   */
  overlap?: number;
  /**
   * Line subtype. Default: `"standard"`. `"stacked"` accumulates
   * series end-to-end, `"percentStacked"` normalizes each category to
   * 100%. Ignored for non-line chart kinds. Maps to
   * `<c:lineChart><c:grouping val="..."/></c:lineChart>`.
   */
  lineGrouping?: "standard" | "stacked" | "percentStacked";
  /**
   * Area subtype. Default: `"standard"`. `"stacked"` paints series on
   * top of each other, `"percentStacked"` normalizes each category to
   * 100%. Ignored for non-area chart kinds. Maps to
   * `<c:areaChart><c:grouping val="..."/></c:areaChart>`.
   */
  areaGrouping?: "standard" | "stacked" | "percentStacked";
  /**
   * Doughnut hole size as a percentage of the outer radius. Accepted
   * range: 10 – 90 (Excel's UI clamps values outside this band).
   * Default: `50` — the Excel default. Ignored for non-doughnut chart
   * kinds.
   */
  holeSize?: number;
  /**
   * Pie / doughnut starting angle in degrees, measured clockwise from
   * the 12 o'clock position. Accepted range: 0 – 360 (the OOXML schema
   * range). Default: `0` — the Excel default (first slice begins at
   * 12 o'clock). Maps to `<c:firstSliceAng val=".."/>`. Ignored for
   * non-pie / non-doughnut chart kinds.
   *
   * Useful for rotating the first wedge into a specific quadrant when
   * composing a dashboard whose pie / doughnut charts should align
   * visually (e.g. `90` to start at 3 o'clock).
   */
  firstSliceAng?: number;
  /**
   * Whether the legend is shown and where. Default: `"right"` for
   * pie/doughnut/bar/line/area, `"bottom"` for scatter. Pass `false`
   * to hide the legend.
   */
  legend?: false | "top" | "bottom" | "left" | "right" | "topRight";
  /** Show the chart-level title element. Default: `true` when `title` is set. */
  showTitle?: boolean;
  /** Alternative text for screen readers (lands in xdr:cNvPr/@descr). */
  altText?: string;
  /** Caption for the chart frame (lands in xdr:cNvPr/@title). */
  frameTitle?: string;
  /**
   * Chart-level data labels applied to every series that does not set
   * its own {@link ChartSeries.dataLabels}. Pass a single
   * {@link ChartDataLabels} object to enable Excel's small in-chart
   * value/category annotations.
   */
  dataLabels?: ChartDataLabels;
  /**
   * How Excel renders missing / blank cells in the source data. Maps
   * to `<c:dispBlanksAs val=".."/>` on `<c:chart>`. Default: `"gap"`
   * (the OOXML default Excel itself emits). Set `"zero"` to anchor the
   * line / bar to the X axis at missing points, or `"span"` to
   * connect across the gap on line and scatter charts. See
   * {@link ChartDisplayBlanksAs} for the accepted set.
   */
  dispBlanksAs?: ChartDisplayBlanksAs;
  /**
   * Vary the color of each data point within the same series. Maps to
   * `<c:varyColors val=".."/>` on the chart-type element
   * (`<c:barChart>`, `<c:lineChart>`, `<c:pieChart>`, ...). Excel
   * exposes the same toggle under "Format Data Series → Fill →
   * Vary colors by point".
   *
   * Excel's per-family defaults differ:
   *   - `pie`, `doughnut`         → `true`  (each slice gets a unique color)
   *   - `bar`, `column`, `line`,
   *     `area`, `scatter`         → `false` (every point on a series
   *                                  shares one color)
   *
   * The writer falls back to those per-family defaults when the field
   * is omitted, so a fresh chart matches Excel's reference
   * serialization. Pin `true` on a single-series bar / column chart to
   * paint each bar a different color, or pin `false` on a doughnut to
   * collapse every wedge to the same color (Excel's "single color"
   * preset).
   *
   * The OOXML schema places `<c:varyColors>` on every chart-type
   * element except `surfaceChart`, `surface3DChart`, and `stockChart`.
   * Hucre's writer emits the element on every authored family, so
   * `varyColors` round-trips on bar / column / line / pie / doughnut /
   * area / scatter charts; surface / stock are not authored by hucre's
   * writer.
   */
  varyColors?: boolean;
  /**
   * Scatter sub-style for `scatter` charts. Maps to
   * `<c:scatterChart><c:scatterStyle val=".."/></c:scatterChart>`.
   * Default: `"lineMarker"` (Excel's chart-picker default — straight
   * lines with markers). Pass `"smooth"` for Excel's "Scatter with
   * Smooth Lines", `"marker"` / `"none"` for "Scatter with Only
   * Markers", `"line"` for "Scatter with Straight Lines", and
   * `"smoothMarker"` for "Scatter with Smooth Lines and Markers". See
   * {@link ChartScatterStyle} for the full preset list.
   *
   * Ignored for every other chart kind — the OOXML schema places
   * `<c:scatterStyle>` exclusively on `<c:scatterChart>`. Use the
   * per-series {@link ChartSeries.smooth} flag to pick a curve on a
   * line chart or pin smoothing on individual scatter series.
   */
  scatterStyle?: ChartScatterStyle;
  /**
   * Whether the chart only plots data from visible cells. Maps to
   * `<c:plotVisOnly val=".."/>` on `<c:chart>`. Mirrors Excel's
   * "Hidden and Empty Cells → Show data in hidden rows and columns"
   * checkbox: when the box is checked, hidden cells stay in the chart
   * and `plotVisOnly` is `false`; when unchecked (the Excel UI
   * default), hidden cells drop out and `plotVisOnly` is `true`.
   *
   * Default: `true` — the OOXML schema default and what every fresh
   * Excel chart emits. Set `false` to keep hidden rows / columns in
   * the rendered chart, useful when the source data range hides helper
   * cells or the dashboard's filter view should not affect the chart.
   *
   * The writer always emits the element so the rendered intent is
   * explicit on roundtrip — Excel itself includes it in every reference
   * serialization.
   */
  plotVisOnly?: boolean;
  /**
   * Whether the chart frame is drawn with rounded corners. Maps to
   * `<c:roundedCorners val=".."/>` on `<c:chartSpace>` (a sibling of
   * `<c:chart>`, not a child). Mirrors Excel's "Format Chart Area →
   * Border → Rounded corners" toggle.
   *
   * Default: `false` — the OOXML schema default and what every fresh
   * Excel chart emits. Set `true` to soften the chart frame's outer
   * edge, useful when matching a dashboard whose other charts already
   * carry the rounded look from a template.
   *
   * The writer always emits the element so the rendered intent is
   * explicit on roundtrip — Excel itself includes it in every reference
   * serialization.
   */
  roundedCorners?: boolean;
  /**
   * Per-axis configuration rendered alongside the plot area. The `x`
   * axis is the category axis for bar/column/line/area (or the bottom
   * value axis for scatter); the `y` axis is the value axis. Ignored
   * for `pie` and `doughnut` charts because they have no axes in
   * OOXML.
   *
   * `title` maps to a `<c:title>` element nested inside the matching
   * `<c:catAx>` / `<c:valAx>`. Pass an empty string or omit the entry
   * to skip the title — Excel renders no axis label by default.
   *
   * `gridlines` toggles `<c:majorGridlines>` / `<c:minorGridlines>`.
   * Omitting the field skips both — useful when porting a clean look
   * across cloned charts. Set `major: true` to draw the heavier
   * reference lines that Excel shows by default on the value axis;
   * `minor: true` adds the lighter half-step lines.
   *
   * `scale` pins the value axis to explicit `<c:min>` / `<c:max>` /
   * `<c:majorUnit>` / `<c:minorUnit>` / `<c:logBase>` bounds. Excel
   * auto-computes any field omitted from the object. Bar/column/line/
   * area charts apply scaling to the Y axis (`<c:valAx>`); scatter
   * charts apply it to whichever axis the field is set on.
   *
   * `numberFormat` pins the tick-label format via `<c:numFmt>` —
   * useful when the cloned chart needs a different format from the
   * source data range (e.g. forcing `"0.00%"` on a percentage chart
   * whose underlying cells are stored as decimals).
   *
   * `tickLblSkip` and `tickMarkSkip` thin out a crowded category axis.
   * Both map to category-axis-only OOXML elements (`<c:tickLblSkip>` /
   * `<c:tickMarkSkip>` on `CT_CatAx` / `CT_DateAx`); they have no slot
   * on `<c:valAx>` and are silently ignored on the value axis or on
   * scatter charts (whose two axes are both value axes).
   *
   * `hidden` collapses the axis line, tick marks, and tick labels off
   * the rendered chart by emitting `<c:delete val="1"/>`. Maps to
   * Excel's "Format Axis -> Axis Options -> Labels -> Show axis" toggle
   * (and the matching context-menu "Delete" action). Useful for
   * minimal "sparkline-style" dashboard tiles where only the data
   * series should remain visible.
   */
  axes?: {
    /** Category axis (bar/column/line/area) or X value axis (scatter). */
    x?: {
      title?: string;
      gridlines?: ChartAxisGridlines;
      scale?: ChartAxisScale;
      numberFormat?: ChartAxisNumberFormat;
      /**
       * Major tick-mark style. Maps to
       * `<c:catAx><c:majorTickMark val=".."/></c:catAx>` (or
       * `<c:valAx>` for scatter). Default: `"out"` — Excel's reference
       * serialization. See {@link ChartAxisTickMark}.
       */
      majorTickMark?: ChartAxisTickMark;
      /**
       * Minor tick-mark style. Maps to
       * `<c:catAx><c:minorTickMark val=".."/></c:catAx>` (or
       * `<c:valAx>` for scatter). Default: `"none"` — Excel's
       * reference serialization. See {@link ChartAxisTickMark}.
       */
      minorTickMark?: ChartAxisTickMark;
      /**
       * Tick-label position. Maps to
       * `<c:catAx><c:tickLblPos val=".."/></c:catAx>` (or
       * `<c:valAx>` for scatter). Default: `"nextTo"` — Excel's
       * reference serialization. See {@link ChartAxisTickLabelPosition}.
       */
      tickLblPos?: ChartAxisTickLabelPosition;
      /**
       * Reverse the axis plotting order. Maps to
       * `<c:scaling><c:orientation val="maxMin"/></c:scaling>` —
       * Excel's "Categories in reverse order" / "Values in reverse
       * order" toggle. Default: `false` (the OOXML `"minMax"` default).
       *
       * On a category axis, reversing flips the order in which
       * categories are drawn (right-to-left on a column chart, top-to-
       * bottom on a bar chart). On a value axis, reversing flips the
       * numeric direction so the maximum sits at the origin and the
       * minimum at the far end. Useful when porting templates that
       * pin a specific reading direction (e.g. dates on a horizontal
       * bar chart with the most recent at the top).
       */
      reverse?: boolean;
      /**
       * Show every Nth tick label on a category axis. `1` (the OOXML
       * default) shows every label; `2` shows every other one; `3`
       * shows every third, and so on. Maps to
       * `<c:catAx><c:tickLblSkip val="N"/></c:catAx>`. Only meaningful
       * for bar / column / line / area charts (whose X axis is
       * `<c:catAx>`); silently ignored for scatter (both axes are
       * value axes) and pie / doughnut (no axes at all). Accepted
       * range: positive integers 1..32767 (the OOXML
       * `ST_SkipIntervals` schema). Values outside the range or
       * non-positive are dropped at write time.
       */
      tickLblSkip?: number;
      /**
       * Show every Nth tick mark on a category axis. Same `1`-default
       * semantics as {@link tickLblSkip} but for the short tick lines
       * Excel paints alongside each label. Maps to
       * `<c:catAx><c:tickMarkSkip val="N"/></c:catAx>`. Same
       * scope-restriction as `tickLblSkip` — category axes only.
       */
      tickMarkSkip?: number;
      /**
       * Distance between the tick labels and the axis line on a
       * category axis, expressed as a percentage of the default
       * spacing. `100` (the OOXML default) renders Excel's reference
       * spacing; lower values pull the labels in towards the axis,
       * higher values push them out. Maps to
       * `<c:catAx><c:lblOffset val="N"/></c:catAx>`. Only meaningful
       * for bar / column / line / area charts (whose X axis is
       * `<c:catAx>`); silently ignored for scatter (both axes are
       * value axes) and pie / doughnut (no axes at all). Accepted
       * range: `0..1000` (the OOXML `ST_LblOffsetPercent` schema).
       * Values outside the range are dropped at write time.
       */
      lblOffset?: number;
      /**
       * Horizontal alignment of the tick labels on a category axis —
       * `"ctr"` (center, the OOXML default), `"l"` (left), or `"r"`
       * (right). Maps to `<c:catAx><c:lblAlgn val=".."/></c:catAx>`.
       * Useful when category labels are wrapped onto multiple lines
       * and the default centered alignment looks ragged against a
       * column chart's left-aligned bars. Excel's UI exposes the
       * three presets under "Format Axis -> Alignment" on a category
       * axis only.
       *
       * Only meaningful for bar / column / line / area charts (whose X
       * axis is `<c:catAx>`); silently ignored for scatter (both axes
       * are value axes) and pie / doughnut (no axes at all). The OOXML
       * schema (`ST_LblAlgn`) restricts the value to the three tokens
       * above; unknown tokens are dropped at write time. See
       * {@link ChartAxisLabelAlign}.
       */
      lblAlgn?: ChartAxisLabelAlign;
      /**
       * Hide the entire axis (line, tick marks, tick labels). Maps to
       * `<c:catAx><c:delete val="1"/></c:catAx>` (or the matching
       * `<c:valAx>` element on scatter). Default: `false` — Excel
       * paints the axis. Set `true` to collapse a noisy axis off a
       * sparkline-style dashboard tile.
       *
       * Excel still reserves the layout slot the axis would have
       * occupied, so a hidden category axis on a column chart leaves a
       * thin gap at the bottom of the plot area where the labels would
       * have rendered — pair with `<c:layout>` overrides on the parent
       * `<c:plotArea>` if you need to reclaim that space (hucre does
       * not surface a layout knob today; the writer falls back to
       * Excel's auto-layout in either case).
       *
       * The flag is silently ignored on `pie` / `doughnut` charts
       * because the OOXML schema places no axes on those families.
       */
      hidden?: boolean;
    };
    /** Value axis. */
    y?: {
      title?: string;
      gridlines?: ChartAxisGridlines;
      scale?: ChartAxisScale;
      numberFormat?: ChartAxisNumberFormat;
      /**
       * Major tick-mark style for the value axis. Maps to
       * `<c:valAx><c:majorTickMark val=".."/></c:valAx>`. Default:
       * `"out"`. See {@link ChartAxisTickMark}.
       */
      majorTickMark?: ChartAxisTickMark;
      /**
       * Minor tick-mark style for the value axis. Maps to
       * `<c:valAx><c:minorTickMark val=".."/></c:valAx>`. Default:
       * `"none"`. See {@link ChartAxisTickMark}.
       */
      minorTickMark?: ChartAxisTickMark;
      /**
       * Tick-label position for the value axis. Maps to
       * `<c:valAx><c:tickLblPos val=".."/></c:valAx>`. Default:
       * `"nextTo"`. See {@link ChartAxisTickLabelPosition}.
       */
      tickLblPos?: ChartAxisTickLabelPosition;
      /**
       * Hide the entire value axis (line, tick marks, tick labels).
       * Maps to `<c:valAx><c:delete val="1"/></c:valAx>`. Default:
       * `false`. See {@link SheetChart.axes.x.hidden} for the full
       * semantics — the value-axis flag mirrors the X-axis flag.
       */
      hidden?: boolean;
      /**
       * Reverse the value axis plotting order. Maps to
       * `<c:valAx><c:scaling><c:orientation val="maxMin"/></c:scaling></c:valAx>`.
       * Default: `false` (the OOXML `"minMax"` default).
       *
       * Mirrors {@link SheetChart.axes.x.reverse} for the value axis —
       * setting `true` flips the numeric direction so the maximum sits
       * at the origin and the minimum at the far end.
       */
      reverse?: boolean;
    };
  };
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
  /**
   * Series-level data labels parsed from the `<c:ser><c:dLbls>` block.
   * Falls back to the chart-level {@link Chart.dataLabels} when this
   * series carries no override of its own.
   */
  dataLabels?: ChartDataLabelsInfo;
  /**
   * Smoothed-line flag pulled from `<c:ser><c:smooth val=".."/>`.
   * Surfaces only on `line` / `scatter` series — the OOXML schema places
   * `<c:smooth>` exclusively on `CT_LineSer` and `CT_ScatterSer`. `false`
   * collapses to `undefined` because it matches the OOXML default and
   * round-trips identically with absence of the field.
   */
  smooth?: boolean;
  /**
   * Line stroke pulled from `<c:ser><c:spPr><a:ln>` — preset dash
   * pattern and width in points. Surfaces only on `line` / `scatter`
   * series so a dashed-stroke template round-trips through
   * `parseChart` → {@link cloneChart} → `writeXlsx`. Field semantics
   * mirror the write-side {@link ChartLineStroke}, so the value can be
   * fed straight into a clone without transformation.
   */
  stroke?: ChartLineStroke;
  /**
   * Marker styling parsed from `<c:ser><c:marker>`. Surfaces only on
   * `line` / `scatter` series — the OOXML schema places `<c:marker>`
   * exclusively on `CT_LineSer` and `CT_ScatterSer`. Empty marker
   * blocks (no symbol, size, or color) collapse to `undefined` so a
   * round-trip keeps the read-side shape minimal. Field semantics
   * mirror the write-side {@link ChartMarker}, so the value can be
   * fed straight into {@link cloneChart} without transformation.
   */
  marker?: ChartMarker;
  /**
   * Invert-if-negative flag pulled from
   * `<c:ser><c:invertIfNegative val=".."/>`. Surfaces only on `bar`
   * (and `bar3D`) series — the OOXML schema places
   * `<c:invertIfNegative>` exclusively on `CT_BarSer` / `CT_Bar3DSer`.
   * `false` collapses to `undefined` because it matches the OOXML
   * default and round-trips identically with absence of the field.
   */
  invertIfNegative?: boolean;
}

/**
 * Read-side mirror of {@link ChartDataLabels}. Exposes the same toggle
 * fields parsed from `<c:dLbls>` so a `ChartSeriesInfo` returned by
 * `parseChart` can be fed straight into {@link cloneChart} without
 * transformation.
 */
export interface ChartDataLabelsInfo {
  showValue?: boolean;
  showCategoryName?: boolean;
  showSeriesName?: boolean;
  showPercent?: boolean;
  position?: ChartDataLabelPosition;
  separator?: string;
}

/**
 * Cell-anchored placement for a chart on its host sheet.
 *
 * Mirrors the `<xdr:from>` / `<xdr:to>` pair on the drawing-layer
 * `xdr:twoCellAnchor` (or the `<xdr:from>` alone for a
 * `xdr:oneCellAnchor`). Coordinates are 0-based row/col indices into
 * the worksheet — identical to the convention used by
 * {@link SheetImage.anchor} and {@link SheetChart.anchor}, so a parsed
 * `ChartAnchor` slots straight back into the writer's shape.
 *
 * `to` is optional because Excel also supports `xdr:oneCellAnchor`
 * (chart pinned to a single cell with intrinsic size).
 * `xdr:absoluteAnchor` (EMU-positioned) does not surface here — those
 * charts are reported with `anchor` undefined.
 */
export interface ChartAnchor {
  /** Top-left cell (`<xdr:from>`). */
  from: { row: number; col: number };
  /** Bottom-right cell (`<xdr:to>`). Omitted for one-cell anchors. */
  to?: { row: number; col: number };
}

/**
 * Major / minor gridline visibility for a chart axis.
 *
 * Excel paints horizontal or vertical reference lines across the plot
 * area, anchored to the major or minor tick marks of an axis. The
 * presence of `<c:majorGridlines>` / `<c:minorGridlines>` inside an
 * `<c:catAx>` or `<c:valAx>` toggles them on; absence of the element
 * means the gridline is off (Excel's default for the value axis is
 * major-on/minor-off, but the OOXML serialization is explicit either
 * way — the writer mirrors what the model says).
 */
export interface ChartAxisGridlines {
  /** Whether the axis declares `<c:majorGridlines>`. */
  major?: boolean;
  /** Whether the axis declares `<c:minorGridlines>`. */
  minor?: boolean;
}

/**
 * Per-axis metadata pulled from the chart's `<c:catAx>` / `<c:valAx>`
 * elements.
 *
 * Surfaces the structural pieces that dashboard cloning needs to
 * preserve through a `parseChart` → {@link cloneChart} → `writeXlsx`
 * round-trip — currently the axis title and the gridline visibility.
 */
/**
 * Value-axis scaling pulled from `<c:scaling>` — bounds plus tick
 * spacing. Excel reports a numeric scale for any value-axis chart;
 * absent on category axes (`<c:catAx>` tolerates `<c:scaling>` but
 * populates only `<c:orientation>` there).
 *
 * All four numeric fields are optional — a chart may declare any
 * subset and Excel auto-computes the rest. Hucre surfaces only the
 * explicitly declared values, so a round-trip cannot accidentally pin
 * an axis to numbers Excel would otherwise have inferred.
 */
export interface ChartAxisScale {
  /** `<c:min>` — value where the axis starts. */
  min?: number;
  /** `<c:max>` — value where the axis ends. */
  max?: number;
  /** `<c:majorUnit>` — spacing between major tick marks. Must be > 0. */
  majorUnit?: number;
  /** `<c:minorUnit>` — spacing between minor tick marks. Must be > 0. */
  minorUnit?: number;
  /**
   * `<c:logBase>` — log base for a logarithmic scale. Excel restricts
   * this to 2–1000; the parser does not enforce that range, but the
   * writer rejects values outside it.
   */
  logBase?: number;
}

/**
 * Axis number-format spec pulled from `<c:numFmt>`. Mirrors what Excel
 * emits for tick labels — an OOXML number-format code (e.g.
 * `"#,##0"`, `"0.00%"`, `"$#,##0.00"`) and a `sourceLinked` flag that
 * tells Excel whether to inherit the cell number format from the
 * underlying data range.
 */
export interface ChartAxisNumberFormat {
  /** OOXML format code (e.g. `"#,##0"`, `"0.00%"`). */
  formatCode: string;
  /**
   * When `true`, Excel ignores `formatCode` and pulls the format
   * straight from the source data range. Defaults to `false` when
   * omitted — the pinned `formatCode` wins.
   */
  sourceLinked?: boolean;
}

/**
 * Axis tick-mark style — where Excel paints the short tick lines that
 * mark major or minor unit boundaries on a category or value axis.
 *
 * Maps to the OOXML `ST_TickMark` enumeration which sits inside
 * `<c:catAx>` / `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` as
 * `<c:majorTickMark val=".."/>` and `<c:minorTickMark val=".."/>`:
 *
 * - `"none"`  — no tick marks rendered at all.
 * - `"in"`    — tick marks point inward (toward the plot area).
 * - `"out"`   — tick marks point outward (away from the plot area).
 *               OOXML default for `<c:majorTickMark>`.
 * - `"cross"` — tick marks straddle the axis line.
 *
 * Excel's UI exposes the same four presets under "Format Axis →
 * Tick Marks → Major type / Minor type". The OOXML default for
 * `<c:minorTickMark>` is `"none"` (Excel's UI also defaults to "None"
 * for the minor type on a freshly-drawn axis).
 */
export type ChartAxisTickMark = "none" | "in" | "out" | "cross";

/**
 * Axis tick-label position — where Excel paints the numeric / category
 * labels relative to the axis line.
 *
 * Maps to the OOXML `ST_TickLblPos` enumeration which sits inside
 * `<c:catAx>` / `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` as
 * `<c:tickLblPos val=".."/>`:
 *
 * - `"nextTo"` — labels sit alongside the axis line at the closest
 *                edge of the plot area. OOXML default.
 * - `"low"`    — labels pinned to the low end of the perpendicular
 *                axis (left for value axes, bottom for category axes).
 *                Useful when the axis crosses elsewhere but labels
 *                should stay anchored to the chart edge.
 * - `"high"`   — mirror of `"low"`; labels pinned to the high end.
 * - `"none"`   — no labels rendered. Excel's UI exposes this as
 *                "Format Axis → Labels → Label Position → None".
 */
export type ChartAxisTickLabelPosition = "nextTo" | "low" | "high" | "none";

/**
 * Horizontal alignment for category-axis tick labels — where Excel
 * anchors each label inside its allocated cell along the axis.
 *
 * Maps to the OOXML `ST_LblAlgn` enumeration which sits inside
 * `<c:catAx>` / `<c:dateAx>` as `<c:lblAlgn val=".."/>`. The element
 * does not exist on `<c:valAx>` / `<c:serAx>`:
 *
 * - `"ctr"` — labels centered along the axis. OOXML default and what
 *             Excel paints on a freshly-drawn category axis.
 * - `"l"`   — labels pinned to the left edge of their slot. Useful for
 *             multi-line wrapped labels on a column chart that should
 *             align flush with the leftmost gridline.
 * - `"r"`   — labels pinned to the right edge of their slot.
 *
 * Excel's UI exposes the three presets under "Format Axis ->
 * Alignment -> Text alignment" on a category axis. Pie / doughnut and
 * scatter charts have no category axis, so the field is dropped on
 * those families.
 */
export type ChartAxisLabelAlign = "ctr" | "l" | "r";

export interface ChartAxisInfo {
  /** Plain-text title from the axis's `<c:title>`. Omitted when absent. */
  title?: string;
  /**
   * Major / minor gridline visibility. Omitted when neither
   * `<c:majorGridlines>` nor `<c:minorGridlines>` is declared on the
   * axis (i.e. Excel's "no gridlines" state for both).
   */
  gridlines?: ChartAxisGridlines;
  /**
   * Numeric scaling (`<c:min>` / `<c:max>` / `<c:majorUnit>` /
   * `<c:minorUnit>` / `<c:logBase>`). Omitted when the axis declared
   * none of those children — Excel auto-computes the bounds in that
   * case and the reader leaves the inference up to the consumer.
   */
  scale?: ChartAxisScale;
  /**
   * Tick-label number format (`<c:numFmt>`). Omitted when the axis
   * does not declare one. Mirrors `formatCode` / `sourceLinked` on
   * the writer side.
   */
  numberFormat?: ChartAxisNumberFormat;
  /**
   * Major tick-mark style pulled from `<c:majorTickMark>`. Omitted
   * when absent or when the axis declared the OOXML default `"out"` —
   * absence and the default round-trip identically through
   * {@link cloneChart}, so collapsing the default keeps the parsed
   * shape minimal. See {@link ChartAxisTickMark}.
   */
  majorTickMark?: ChartAxisTickMark;
  /**
   * Minor tick-mark style pulled from `<c:minorTickMark>`. Omitted
   * when absent or when the axis declared the OOXML default `"none"`.
   * See {@link ChartAxisTickMark}.
   */
  minorTickMark?: ChartAxisTickMark;
  /**
   * Tick-label position pulled from `<c:tickLblPos>`. Omitted when
   * absent or when the axis declared the OOXML default `"nextTo"` —
   * absence and the default round-trip identically through
   * {@link cloneChart}, so collapsing the default keeps the parsed
   * shape minimal. See {@link ChartAxisTickLabelPosition}.
   */
  tickLblPos?: ChartAxisTickLabelPosition;
  /**
   * Reverse-axis flag pulled from
   * `<c:scaling><c:orientation val=".."/></c:scaling>`. Surfaces `true`
   * only when the axis pinned `"maxMin"` (Excel's "Categories /
   * Values in reverse order" toggle); the OOXML default `"minMax"`
   * collapses to `undefined` so absence and the default round-trip
   * identically through {@link cloneChart}. Mirrors the writer-side
   * {@link SheetChart.axes.x.reverse} field, so a parsed value slots
   * straight back into a clone target without transformation.
   */
  reverse?: boolean;
  /**
   * Tick-label skip interval pulled from `<c:tickLblSkip val=".."/>`.
   * Surfaces only on category axes (`<c:catAx>` / `<c:dateAx>`) — the
   * OOXML schema does not place this element on `<c:valAx>`. The
   * default `1` (show every label) collapses to `undefined` so absence
   * and the default round-trip identically through {@link cloneChart}.
   * Out-of-range values (non-positive or > 32767) are dropped rather
   * than fabricated.
   */
  tickLblSkip?: number;
  /**
   * Tick-mark skip interval pulled from `<c:tickMarkSkip val=".."/>`.
   * Same scope (category axes only) and default-collapse semantics as
   * {@link tickLblSkip}.
   */
  tickMarkSkip?: number;
  /**
   * Label offset pulled from `<c:lblOffset val=".."/>`, expressed as a
   * percentage of the default axis-label spacing. Surfaces only on
   * category axes (`<c:catAx>` / `<c:dateAx>`) — the OOXML schema
   * (`ST_LblOffsetPercent`) does not place this element on `<c:valAx>`
   * or `<c:serAx>`. The default `100` (Excel's reference spacing)
   * collapses to `undefined` so absence and the default round-trip
   * identically through {@link cloneChart}. Accepted range is `0..1000`;
   * out-of-range values are dropped rather than fabricated.
   */
  lblOffset?: number;
  /**
   * Tick-label horizontal alignment pulled from `<c:lblAlgn val=".."/>`.
   * Surfaces only on category axes (`<c:catAx>` / `<c:dateAx>`) — the
   * OOXML schema (`ST_LblAlgn`) does not place this element on
   * `<c:valAx>` or `<c:serAx>`. The default `"ctr"` (Excel's reference
   * centered alignment) collapses to `undefined` so absence and the
   * default round-trip identically through {@link cloneChart}. Unknown
   * tokens drop to `undefined` rather than fabricate a value the
   * writer would never emit. See {@link ChartAxisLabelAlign}.
   */
  lblAlgn?: ChartAxisLabelAlign;
  /**
   * Axis hidden flag pulled from `<c:delete val=".."/>`. Surfaces
   * `true` when the axis pinned `val="1"` (Excel's "Format Axis ->
   * Show axis = off" toggle). The OOXML default `val="0"` (and absence
   * of the element) collapse to `undefined` so absence and the default
   * round-trip identically through {@link cloneChart}. The reader
   * accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` /
   * `"0"` / `"false"`); unknown values and missing `val` attributes
   * drop to `undefined`.
   */
  hidden?: boolean;
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

/**
 * Line/area grouping reported by {@link Chart.lineGrouping} and
 * {@link Chart.areaGrouping}.
 *
 * Pulled from `<c:lineChart><c:grouping val="..."/></c:lineChart>` or
 * `<c:areaChart><c:grouping val="..."/></c:areaChart>`. Only the
 * stacked variants surface — `"standard"` is the OOXML default and
 * is collapsed to `undefined` for symmetry with the writer's
 * {@link SheetChart.lineGrouping} / {@link SheetChart.areaGrouping}
 * defaults.
 */
export type ChartLineAreaGrouping = "stacked" | "percentStacked";

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
   * Cell anchor pulled from the host drawing's `<xdr:twoCellAnchor>` /
   * `<xdr:oneCellAnchor>`. Undefined when the drawing positions the
   * chart with `<xdr:absoluteAnchor>` (EMU-positioned, no cell anchor)
   * or when the drawing's anchor element is missing the `from` block.
   */
  anchor?: ChartAnchor;
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
  /**
   * Grouping pulled from the first `<c:lineChart>` element, when the
   * chart has one. Surfaces only `"stacked"` / `"percentStacked"` —
   * the OOXML `"standard"` value is the writer default and collapses
   * to `undefined` here.
   */
  lineGrouping?: ChartLineAreaGrouping;
  /**
   * Grouping pulled from the first `<c:areaChart>` element, when the
   * chart has one. Surfaces only `"stacked"` / `"percentStacked"` —
   * the OOXML `"standard"` value is the writer default and collapses
   * to `undefined` here.
   */
  areaGrouping?: ChartLineAreaGrouping;
  /**
   * Chart-level data label defaults parsed from the first chart-type
   * element's `<c:dLbls>` block. Series-level overrides on
   * {@link ChartSeriesInfo.dataLabels} take precedence.
   */
  dataLabels?: ChartDataLabelsInfo;
  /**
   * Per-axis metadata. `x` corresponds to the chart's `<c:catAx>`
   * (category axis on bar/column/line/area) or the first `<c:valAx>`
   * on scatter. `y` corresponds to the value axis. Both fields are
   * omitted on charts that have no axes (e.g. pie/doughnut) or when
   * neither axis carries a title.
   */
  axes?: {
    x?: ChartAxisInfo;
    y?: ChartAxisInfo;
  };
  /**
   * Doughnut hole size pulled from the chart's `<c:doughnutChart>
   * <c:holeSize val=".."/>`, expressed as a percentage of the outer
   * radius (1–99). Omitted on non-doughnut charts and on doughnut
   * charts that do not declare the element.
   */
  holeSize?: number;
  /**
   * Bar/column gap width pulled from the first `<c:barChart>` /
   * `<c:bar3DChart>` element's `<c:gapWidth val=".."/>`, expressed as a
   * percentage of the bar width. Range: 0–500. The OOXML default of
   * `150` collapses to `undefined` so absence and the default
   * round-trip identically — symmetric with how the writer's
   * {@link SheetChart.gapWidth} treats the absence of the field.
   * Omitted on non-bar / non-column charts.
   */
  gapWidth?: number;
  /**
   * Bar/column series overlap pulled from the first `<c:barChart>` /
   * `<c:bar3DChart>` element's `<c:overlap val=".."/>`, expressed as a
   * percentage of the bar width. Range: -100..100. The OOXML default of
   * `0` collapses to `undefined` so absence and the default round-trip
   * identically — symmetric with how the writer's
   * {@link SheetChart.overlap} treats the absence of the field.
   * Omitted on non-bar / non-column charts.
   */
  overlap?: number;
  /**
   * Pie / doughnut starting angle in degrees pulled from the first
   * `<c:pieChart>` / `<c:doughnutChart>` element's
   * `<c:firstSliceAng val=".."/>`. Range: 0–360. `0` collapses to
   * `undefined` because it is the OOXML default (first slice at the
   * 12 o'clock position) — the writer's
   * {@link SheetChart.firstSliceAng} treats the absence of the field
   * the same way. Omitted on non-pie / non-doughnut charts.
   */
  firstSliceAng?: number;
  /**
   * How the chart renders missing / blank cells, pulled from
   * `<c:chart><c:dispBlanksAs val=".."/>`. The OOXML default of
   * `"gap"` collapses to `undefined` so absence and the default
   * round-trip identically through {@link cloneChart} — symmetric with
   * the writer's {@link SheetChart.dispBlanksAs} field. Surfaces
   * `"zero"` and `"span"` literally; unknown values are dropped rather
   * than fabricated.
   */
  dispBlanksAs?: ChartDisplayBlanksAs;
  /**
   * Vary-colors-by-point flag pulled from the first chart-type
   * element's `<c:varyColors val=".."/>`. Reflects Excel's
   * per-family default by collapsing matching values to `undefined`:
   *
   *   - On `pie`, `pie3D`, `doughnut`, `ofPie` charts, the OOXML
   *     default is `true` — `<c:varyColors val="1"/>` and absence both
   *     collapse to `undefined`; only an explicit `<c:varyColors val="0"/>`
   *     surfaces `false`.
   *   - On every other chart family the OOXML default is `false` —
   *     `<c:varyColors val="0"/>` and absence both collapse to
   *     `undefined`; only an explicit `<c:varyColors val="1"/>`
   *     surfaces `true`.
   *
   * The asymmetric collapse keeps the parsed shape minimal — a pure
   * round-trip of a stock chart returns no `varyColors` field, while
   * a template that overrides the per-family default surfaces the
   * non-default value so {@link cloneChart} can carry it through.
   * Omitted on chart families that have no `<c:varyColors>` slot
   * (`surface`, `surface3D`, `stock`).
   */
  varyColors?: boolean;
  /**
   * Scatter sub-style pulled from `<c:scatterChart><c:scatterStyle
   * val=".."/></c:scatterChart>`. Reflects which of Excel's six XY
   * scatter presets the chart was authored with — `"none"`, `"line"`,
   * `"lineMarker"`, `"marker"`, `"smooth"`, or `"smoothMarker"`. The
   * OOXML default `"marker"` collapses to `undefined` (Excel's reference
   * serialization actually emits `"lineMarker"` even at the UI default,
   * so the reader does not pin a default of its own — both `"marker"`
   * and `"lineMarker"` surface literally so a clone preserves what the
   * template said).
   *
   * Omitted on every chart family except `scatter`; the OOXML schema
   * places `<c:scatterStyle>` exclusively on `<c:scatterChart>`.
   */
  scatterStyle?: ChartScatterStyle;
  /**
   * Plot-visible-only flag pulled from
   * `<c:chart><c:plotVisOnly val=".."/>`. Reflects Excel's "Hidden and
   * Empty Cells → Show data in hidden rows and columns" toggle (the
   * checkbox is the inverse of this flag — checked means `false`,
   * unchecked means `true`).
   *
   * The OOXML default `true` collapses to `undefined` so absence and
   * the default round-trip identically through {@link cloneChart} —
   * only an explicit `<c:plotVisOnly val="0"/>` surfaces `false`. The
   * reader accepts the OOXML truthy / falsy spellings (`"1"` / `"true"`
   * / `"0"` / `"false"`); unknown values and missing `val` attributes
   * drop to `undefined`.
   */
  plotVisOnly?: boolean;
  /**
   * Rounded-corners flag pulled from
   * `<c:chartSpace><c:roundedCorners val=".."/>`. Reflects Excel's
   * "Format Chart Area → Border → Rounded corners" toggle, which paints
   * the chart frame with rounded edges instead of the default square
   * border.
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * the default round-trip identically through {@link cloneChart} —
   * only an explicit `<c:roundedCorners val="1"/>` surfaces `true`.
   * The reader accepts the OOXML truthy / falsy spellings (`"1"` /
   * `"true"` / `"0"` / `"false"`); unknown values and missing `val`
   * attributes drop to `undefined`.
   *
   * Note: `<c:roundedCorners>` lives on `<c:chartSpace>`, not inside
   * `<c:chart>` — the toggle styles the outer frame, not the plot area.
   */
  roundedCorners?: boolean;
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

// ── High-Level API ──────────────────────────────────────────────────
export { read, write, readObjects, writeObjects } from "./defter"
export type { ReadObjectsOptions, ReadObjectsResult, WriteObjectsTableOption } from "./defter"

// ── XLSX ────────────────────────────────────────────────────────────
export { readXlsx } from "./xlsx/reader"
export { readXlsb } from "./xlsx/xlsb/reader"
export { readXls } from "./xls/reader"
export { writeXlsx } from "./xlsx/writer"
export { link } from "./xlsx/hyperlink"
export { openXlsx, saveXlsx } from "./xlsx/roundtrip"
export type { RoundtripWorkbook } from "./xlsx/roundtrip"
export { hashSheetPassword } from "./xlsx/password"
export { calculateColumnWidth, measureValueWidth, calculateRowHeight } from "./xlsx/auto-size"
// ── Low-level OOXML part parsers ───────────────────────────────────
//
// These take a raw XML string from inside an .xlsx package and return
// hucre's internal model of it. They now live at `hucre/ooxml`, which is
// explicitly outside the v1 stability commitment — their shapes mirror
// the parse pipeline and move when it does.
//
// Kept here for backward compatibility. Prefer `hucre/ooxml`.
//
// @deprecated Import from `hucre/ooxml` instead.
export { parseThemeColors, resolveThemeColor } from "./xlsx/theme"
export { streamXlsxRows } from "./xlsx/stream-reader"
export type { StreamRow } from "./xlsx/stream-reader"
export {
  XlsxStreamWriter,
  writeXlsxStream,
  writeXlsxStreamSheets,
  XLSX_MAX_ROWS_PER_SHEET,
} from "./xlsx/stream-writer"
export type {
  StreamWriterOptions,
  XlsxStreamWriterOptions,
  XlsxWriteStreamOptions,
  XlsxWriteStreamWorkbookOptions,
  XlsxStreamRow,
  XlsxStreamSheet,
  StreamStyledCell,
} from "./xlsx/stream-writer"
export { readXlsxObjects, writeXlsxObjects } from "./xlsx/objects"
export type {
  XlsxObjectsReadOptions,
  XlsxObjectsResult,
  XlsxObjectsWriteOptions,
} from "./xlsx/objects"

// ── ODS ────────────────────────────────────────────────────────────
export { readOds } from "./ods/reader"
export { writeOds } from "./ods/writer"
export { writeOdsStream } from "./ods/stream-writer"
export type { OdsWriteRow, OdsWriteCell, OdsStreamWriteOptions } from "./ods/stream-writer"
export { streamOdsRows } from "./ods/stream"
export { readOdsObjects, writeOdsObjects } from "./ods/objects"
export type { OdsObjectsReadOptions, OdsObjectsResult, OdsObjectsWriteOptions } from "./ods/objects"

// ── CSV ────────────────────────────────────────────────────────────
export {
  parseCsv,
  parseCsvObjects,
  detectDelimiter,
  stripBom,
  writeCsv,
  writeCsvObjects,
  formatCsvValue,
  fetchCsv,
} from "./csv/index"
export type { CsvObjectsResult } from "./csv/index"
export { streamCsvRows, CsvStreamWriter, writeCsvStream } from "./csv/stream"
export { decodeCsvInput, detectBom } from "./csv/encoding"
export type { CsvInput, BomEncoding } from "./csv/encoding"
export type { CsvStreamRow, CsvStreamWriterOptions } from "./csv/stream"

// ── JSON ───────────────────────────────────────────────────────────
export {
  parseJson,
  parseValue,
  parseNdjson,
  writeJson,
  writeNdjson,
  workbookToJson,
  jsonToWorkbook,
  NdjsonStreamWriter,
  streamNdjsonRows,
  readNdjsonStream,
  writeNdjsonStream,
  // Exported from hucre/json but not from the root, so the two surfaces
  // disagreed about what the JSON API is.
  flattenValue,
  collectHeaders,
  unflattenRow,
  unflattenRows,
} from "./json"
export type {
  JsonReadOptions,
  JsonReadResult,
  JsonToWorkbookOptions,
  NdjsonReadOptions,
  JsonWriteOptions,
  WorkbookToJsonOptions,
  NdjsonStreamReadOptions,
  NdjsonStreamRow,
  NdjsonStreamWriterOptions,
  FlattenOptions,
  UnflattenedRow,
} from "./json"

// ── XML ────────────────────────────────────────────────────────────
export { readXml, writeXml, writeXmlStream } from "./xml"
export type { XmlReadOptions, XmlReadResult, XmlWriteOptions } from "./xml"

// ── Schema Validation ──────────────────────────────────────────────
export { validateWithSchema } from "./_schema"

// ── Threaded Comments (Excel 365+) ─────────────────────────────────
export { parsePersons, parseThreadedComments } from "./xlsx/threaded-comments-reader"
export type { ThreadedComment, ThreadedCommentMention, ThreadedCommentPerson } from "./_types"

// ── Accessibility ──────────────────────────────────────────────────
export * as a11y from "./a11y"
export type { A11yIssue, A11ySeverity, A11yCode, A11yLocation, SheetA11y } from "./_types"

// ── External Workbook Links ────────────────────────────────────────
export { parseExternalLink } from "./xlsx/external-link-reader"
export type {
  ExternalLink,
  ExternalCellType,
  ExternalCachedCell,
  ExternalSheetData,
  ExternalDefinedName,
} from "./_types"

// ── Cell-Embedded Images (WPS DISPIMG) ────────────────────────────
export { parseCellImages, assembleCellImages, REL_CELL_IMAGES } from "./xlsx/cell-images-reader"
export type { ParsedCellImageRef } from "./xlsx/cell-images-reader"
export type { CellImage } from "./_types"

// ── Pivot Tables ───────────────────────────────────────────────────
export {
  parsePivotTable,
  parsePivotCacheDefinition,
  attachPivotCacheFields,
} from "./xlsx/pivot-reader"
export type {
  PivotTable,
  PivotCache,
  PivotField,
  PivotFieldAxis,
  PivotDataFieldFunction,
  WritePivotTable,
  WritePivotDataField,
} from "./_types"

// ── Slicers & Timelines ────────────────────────────────────────────
export {
  parseSlicers,
  parseSlicerCache,
  parseTimelines,
  parseTimelineCache,
} from "./xlsx/slicer-reader"
export type {
  Slicer,
  SlicerCache,
  SlicerCachePivotTable,
  SlicerCacheTableSource,
  Timeline,
  TimelineCache,
} from "./_types"

// ── Charts ─────────────────────────────────────────────────────────
export { parseChart } from "./xlsx/chart-reader"
export { cloneChart, chartKindToWriteKind } from "./xlsx/chart-clone"
export type { CloneChartOptions, CloneChartSeriesOverride } from "./xlsx/chart-clone"
export { addChart, getCharts } from "./xlsx/chart-helpers"
export type { ChartLocation } from "./xlsx/chart-helpers"
export type {
  Chart,
  ChartAnchor,
  ChartAxisGridlines,
  ChartAxisInfo,
  ChartAxisTickLabelPosition,
  ChartAxisTickMark,
  ChartBarGrouping,
  ChartDataLabelPosition,
  ChartDataLabels,
  ChartDataLabelsInfo,
  ChartDataTable,
  ChartDisplayBlanksAs,
  ChartKind,
  ChartLegendEntry,
  ChartLegendPosition,
  ChartLineAreaGrouping,
  ChartLineDashStyle,
  ChartLineStroke,
  ChartManualLayout,
  ChartMarker,
  ChartMarkerSymbol,
  ChartProtection,
  ChartSeriesInfo,
  ChartView3D,
} from "./_types"

// ── Date Utilities ─────────────────────────────────────────────────
export {
  serialToDate,
  dateToSerial,
  isDateFormat,
  formatDate,
  parseDate,
  serialToTime,
  timeToSerial,
} from "./_date"

// ── Number Format ─────────────────────────────────────────────────
export { formatValue } from "./_format"

// ── Style Utilities ───────────────────────────────────────────────
export { cloneCellStyle } from "./_style"
export type { FormatOptions, LocaleFormat } from "./_format"

// ── Builder Pattern ──────────────────────────────────────────────
export { WorkbookBuilder, SheetBuilder } from "./builder"

// ── Template Engine ──────────────────────────────────────────────
export { fillTemplate } from "./template"

// ── Read model → write model ─────────────────────────────────────
export { toWriteOptions, toWriteSheet } from "./write-model"
export type { WriteModelDrop, ToWriteOptionsOptions } from "./write-model"

// ── Sheet Operations ──────────────────────────────────────────────
export {
  insertRows,
  deleteRows,
  insertColumns,
  deleteColumns,
  moveRows,
  hideRows,
  hideColumns,
  groupRows,
  cloneSheet,
  copySheetToWorkbook,
  copyRange,
  moveSheet,
  removeSheet,
  findCells,
  replaceCells,
  sortRows,
} from "./sheet-ops"

// ── Web Worker Helpers ───────────────────────────────────────────
export { serializeWorkbook, deserializeWorkbook } from "./worker"
export type {
  SerializedWorkbook,
  SerializedSheet,
  SerializedCell,
  SerializedCellValue,
  SerializedSheetImage,
  SerializedWorkbookProperties,
} from "./worker"

// ── Cell Utilities ─────────────────────────────────────────────────
export {
  parseCellRef,
  colToLetter,
  cellRef,
  rangeRef,
  letterToCol,
  parseRange,
  isInRange,
  r1c1ToA1,
  a1ToR1C1,
  // Normalise either spelling of a range to coordinates (#474).
  toRange,
  toRanges,
} from "./cell-utils"
export type { RangeLike } from "./cell-utils"

// ── Sheet Utilities ──────────────────────────────────────────────
export { sheetToObjects, sheetToArrays } from "./sheet-utils"
export type { SheetObjectsResult, SheetToObjectsOptions } from "./sheet-utils"

// ── Export (HTML / Markdown / JSON / TSV) ────────────────────────────
export { toHtml, toMarkdown, toJson, fromHtml } from "./export/index"
export type {
  HtmlExportOptions,
  HtmlImportOptions,
  MarkdownExportOptions,
  JsonExportOptions,
} from "./export/index"
export { writeTsv, writeTsvObjects } from "./export/tsv"

// ── Image Utilities ──────────────────────────────────────────────
export { imageFromBase64 } from "./image"

// ── Errors ─────────────────────────────────────────────────────────
export {
  HucreError,
  /** @deprecated Use {@link HucreError}. Same class object — `instanceof` is unaffected. */
  DefterError,
  ParseError,
  ZipError,
  XmlError,
  ValidationError,
  InvalidArgumentError,
  UnsupportedFormatError,
  EncryptedFileError,
  DecryptionError,
} from "./errors"

// ── Resource limits ────────────────────────────────────────────────
//
// The bounds the readers defend themselves with. They are exported so a
// caller can quote the number in their own message, compare against it
// before handing a file over, or pass a raised one back in through the
// matching `ReadOptions` field — rather than hard-coding 20,000,000 and
// hoping it does not move. See #471.
export {
  MAX_COL_INDEX,
  MAX_DECOMPRESSED_BYTES,
  MAX_INPUT_BYTES,
  MAX_REPEAT_COUNT,
  MAX_ROW_INDEX,
  MAX_SPAN_CELLS,
  MAX_SPIN_COUNT,
  MAX_TOTAL_CELLS,
} from "./limits"

// ── Types ──────────────────────────────────────────────────────────
export type {
  /**
   * The vocabulary the three incremental writers share. Implemented by
   * `XlsxStreamWriter`, `CsvStreamWriter` and `NdjsonStreamWriter`. See #468.
   */
  SpreadsheetStreamWriter,
  // Cell
  CellValue,
  CellType,
  Cell,
  RichTextRun,
  Hyperlink,
  HyperlinkValue,
  CellComment,
  // Style
  CellStyle,
  CellProtection,
  FontStyle,
  FillStyle,
  PatternFill,
  GradientFill,
  FillPattern,
  BorderStyle,
  BorderSide,
  BorderLineStyle,
  AlignmentStyle,
  Color,
  // Sheet
  Sheet,
  ColumnDef,
  RowDef,
  MergeRange,
  DataValidation,
  ConditionalRule,
  AutoFilter,
  FreezePane,
  SplitPane,
  SheetImage,
  SheetChart,
  WriteChartKind,
  ChartSeries,
  SheetProtection,
  SheetView,
  PageSetup,
  PageMargins,
  HeaderFooter,
  NamedRange,
  // Workbook
  Workbook,
  WorkbookProperties,
  // Read
  ReadOptions,
  ReadInput,
  SheetFilter,
  SheetFilterInfo,
  ReadWarning,
  // Write
  WriteOptions,
  WriteSheet,
  WriteOutput,
  // Outline
  OutlineProperties,
  // CSV
  CsvReadOptions,
  CsvWriteOptions,
  // Schema
  SchemaDefinition,
  SchemaField,
  SchemaFieldType,
  SchemaValidationIssue,
  // Sheet features reachable through WriteSheet / Sheet but previously
  // unnameable — you could build one inline and never annotate it.
  TableDefinition,
  TableColumn,
  Sparkline,
  SheetTextBox,
  PaperSize,
  PaperSizeName,
  ValidationType,
  ValidationOperator,
  ConditionalRuleType,
  // Chart sub-types reachable through SheetChart / ChartSeries / Chart.
  ChartColor,
  ChartThemeColor,
  ChartThemeColorName,
  ChartDataPoint,
  ChartTrendline,
  ChartTrendlineType,
  ChartErrorBars,
  ChartErrorBarDirection,
  ChartErrorBarType,
  ChartErrorBarValType,
  ChartBorderDash,
  ChartLineCap,
  ChartLineCompound,
  ChartScatterStyle,
  ChartShape3D,
  ChartAxisCrosses,
  ChartAxisCrossBetween,
  ChartAxisDispUnit,
  ChartAxisDispUnits,
  ChartAxisLabelAlign,
  ChartAxisNumberFormat,
  ChartAxisScale,
} from "./_types"

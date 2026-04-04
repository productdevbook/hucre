// ── High-Level API ──────────────────────────────────────────────────
export { read, write, readObjects, writeObjects } from "./defter";

// ── XLSX ────────────────────────────────────────────────────────────
export { readXlsx } from "./xlsx/reader";
export { writeXlsx } from "./xlsx/writer";
export { openXlsx, saveXlsx } from "./xlsx/roundtrip";
export type { RoundtripWorkbook } from "./xlsx/roundtrip";
export { hashSheetPassword } from "./xlsx/password";
export { calculateColumnWidth, measureValueWidth, calculateRowHeight } from "./xlsx/auto-size";
export { parseThemeColors, resolveThemeColor } from "./xlsx/theme";
export { streamXlsxRows } from "./xlsx/stream-reader";
export type { StreamRow } from "./xlsx/stream-reader";
export { XlsxStreamWriter } from "./xlsx/stream-writer";
export type { StreamWriterOptions } from "./xlsx/stream-writer";

// ── ODS ────────────────────────────────────────────────────────────
export { readOds } from "./ods/reader";
export { writeOds } from "./ods/writer";
export { streamOdsRows } from "./ods/stream";

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
} from "./csv/index";
export { streamCsvRows, CsvStreamWriter } from "./csv/stream";

// ── Schema Validation ──────────────────────────────────────────────
export { validateWithSchema } from "./_schema";

// ── Date Utilities ─────────────────────────────────────────────────
export {
  serialToDate,
  dateToSerial,
  isDateFormat,
  formatDate,
  parseDate,
  serialToTime,
  timeToSerial,
} from "./_date";

// ── Number Format ─────────────────────────────────────────────────
export { formatValue } from "./_format";
export type { FormatOptions, LocaleFormat } from "./_format";

// ── Builder Pattern ──────────────────────────────────────────────
export { WorkbookBuilder, SheetBuilder } from "./builder";

// ── Formula Helpers ─────────────────────────────────────────────
export * as fx from "./fx";

// ── Style Presets ───────────────────────────────────────────────
export { slate, ocean, forest, rose, minimal, applyPreset } from "./presets";

// ── Column Utilities ────────────────────────────────────────────
export { pickColumns, omitColumns } from "./column-utils";

// ── Template Engine ──────────────────────────────────────────────
export { fillTemplate } from "./template";

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
} from "./sheet-ops";

// ── Web Worker Helpers ───────────────────────────────────────────
export { serializeWorkbook, deserializeWorkbook, WORKER_SAFE_FUNCTIONS } from "./worker";
export type {
  SerializedWorkbook,
  SerializedSheet,
  SerializedCell,
  SerializedCellValue,
  SerializedSheetImage,
  SerializedWorkbookProperties,
} from "./worker";

// ── Cell Utilities ─────────────────────────────────────────────────
export {
  parseCellRef,
  colToLetter,
  cellRef,
  rangeRef,
  letterToCol,
  parseRange,
  isInRange,
} from "./cell-utils";

// ── Sheet Utilities ──────────────────────────────────────────────
export { sheetToObjects, sheetToArrays } from "./sheet-utils";

// ── Export (HTML / Markdown / JSON / TSV) ────────────────────────────
export { toHtml, toMarkdown, toJson, fromHtml } from "./export/index";
export type { HtmlExportOptions, MarkdownExportOptions, JsonExportOptions } from "./export/index";
export { writeTsv, writeTsvObjects } from "./export/tsv";

// ── Image Utilities ──────────────────────────────────────────────
export { imageFromBase64 } from "./image";

// ── Errors ─────────────────────────────────────────────────────────
export {
  DefterError,
  ParseError,
  ZipError,
  XmlError,
  ValidationError,
  UnsupportedFormatError,
  EncryptedFileError,
} from "./errors";

// ── Types ──────────────────────────────────────────────────────────
export type {
  // Cell
  CellValue,
  CellType,
  Cell,
  RichTextRun,
  Hyperlink,
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
  ReadResult,
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
  ValidationError as ValidationErrorType,
  // Column Builder
  ColumnSummary,
  ColumnCondition,
  StylePreset,
} from "./_types";

// ── High-Level API ──────────────────────────────────────────────────
export { read, write, readObjects, writeObjects } from "./defter";

// ── XLSX ────────────────────────────────────────────────────────────
export { readXlsx } from "./xlsx/reader";
export { writeXlsx } from "./xlsx/writer";
export { openXlsx, saveXlsx } from "./xlsx/roundtrip";
export type { RoundtripWorkbook } from "./xlsx/roundtrip";
export { hashSheetPassword } from "./xlsx/password";
export { streamXlsxRows } from "./xlsx/stream-reader";
export type { StreamRow } from "./xlsx/stream-reader";
export { XlsxStreamWriter } from "./xlsx/stream-writer";
export type { StreamWriterOptions } from "./xlsx/stream-writer";

// ── ODS ────────────────────────────────────────────────────────────
export { readOds } from "./ods/reader";
export { writeOds } from "./ods/writer";

// ── CSV ────────────────────────────────────────────────────────────
export {
  parseCsv,
  parseCsvObjects,
  detectDelimiter,
  stripBom,
  writeCsv,
  writeCsvObjects,
  formatCsvValue,
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
export { parseCellRef } from "./xlsx/worksheet";
export { colToLetter, cellRef, rangeRef } from "./xlsx/worksheet-writer";

// ── Export (HTML / Markdown) ───────────────────────────────────────
export { toHtml, toMarkdown } from "./export/index";
export type { HtmlExportOptions, MarkdownExportOptions } from "./export/index";

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
  // CSV
  CsvReadOptions,
  CsvWriteOptions,
  // Schema
  SchemaDefinition,
  SchemaField,
  SchemaFieldType,
  ValidationError as ValidationErrorType,
} from "./_types";

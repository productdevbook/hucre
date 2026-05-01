// ── High-Level API ──────────────────────────────────────────────────
export { read, write, readObjects, writeObjects } from "./defter";
export type { WriteObjectsTableOption } from "./defter";

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
export { readXlsxObjects, writeXlsxObjects } from "./xlsx/objects";
export type {
  XlsxObjectsReadOptions,
  XlsxObjectsResult,
  XlsxObjectsWriteOptions,
} from "./xlsx/objects";

// ── ODS ────────────────────────────────────────────────────────────
export { readOds } from "./ods/reader";
export { writeOds } from "./ods/writer";
export { streamOdsRows } from "./ods/stream";
export { readOdsObjects, writeOdsObjects } from "./ods/objects";
export type {
  OdsObjectsReadOptions,
  OdsObjectsResult,
  OdsObjectsWriteOptions,
} from "./ods/objects";

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

// ── JSON ───────────────────────────────────────────────────────────
export {
  parseJson,
  parseValue,
  parseNdjson,
  writeJson,
  writeNdjson,
  workbookToJson,
  NdjsonStreamWriter,
  readNdjsonStream,
} from "./json";
export type {
  JsonReadOptions,
  JsonReadResult,
  NdjsonReadOptions,
  JsonWriteOptions,
  WorkbookToJsonOptions,
  NdjsonStreamReadOptions,
  FlattenOptions,
} from "./json";

// ── XML ────────────────────────────────────────────────────────────
export { readXml, writeXml } from "./xml";
export type { XmlReadOptions, XmlReadResult, XmlWriteOptions } from "./xml";

// ── Schema Validation ──────────────────────────────────────────────
export { validateWithSchema } from "./_schema";

// ── Threaded Comments (Excel 365+) ─────────────────────────────────
export { parsePersons, parseThreadedComments } from "./xlsx/threaded-comments-reader";
export type { ThreadedComment, ThreadedCommentMention, ThreadedCommentPerson } from "./_types";

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
  r1c1ToA1,
  a1ToR1C1,
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
  SheetFilter,
  SheetFilterInfo,
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
} from "./_types";

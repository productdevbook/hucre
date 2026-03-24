// ── XLSX ────────────────────────────────────────────────────────────
export { readXlsx } from "./xlsx/reader";
export { writeXlsx } from "./xlsx/writer";

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

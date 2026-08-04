// ── hucre/xlsx entry point ────────────────────────────────────────────
// Read & write XLSX, plus read-only XLSB and legacy XLS.

export { readXlsx } from "./xlsx/reader"
export { writeXlsx } from "./xlsx/writer"
export { readXlsb } from "./xlsx/xlsb/reader"
export { readXls } from "./xls/reader"
export { readXlsxObjects, writeXlsxObjects } from "./xlsx/objects"
export type {
  XlsxObjectsReadOptions,
  XlsxObjectsResult,
  XlsxObjectsWriteOptions,
} from "./xlsx/objects"
export { link } from "./xlsx/hyperlink"
export type { HyperlinkValue } from "./_types"
export { openXlsx, saveXlsx } from "./xlsx/roundtrip"
export type { RoundtripWorkbook } from "./xlsx/roundtrip"
export { hashSheetPassword } from "./xlsx/password"
export { streamXlsxRows } from "./xlsx/stream-reader"
export type { StreamRow } from "./xlsx/stream-reader"
export { XlsxStreamWriter, writeXlsxStream, XLSX_MAX_ROWS_PER_SHEET } from "./xlsx/stream-writer"
export type {
  StreamWriterOptions,
  XlsxStreamWriterOptions,
  XlsxWriteStreamOptions,
  XlsxStreamRow,
} from "./xlsx/stream-writer"

// ── Sizing & theme helpers ─────────────────────────────────────────
export { calculateColumnWidth, measureValueWidth } from "./xlsx/auto-width"
export { calculateRowHeight } from "./xlsx/auto-size"
export { parseThemeColors, resolveThemeColor } from "./xlsx/theme"

// ── Cell Utilities ─────────────────────────────────────────────────
export { parseCellRef } from "./xlsx/worksheet"
export { colToLetter, cellRef, rangeRef } from "./xlsx/worksheet-writer"

// ── Shared types used by this entry point's signatures ──────────────
// Re-exported so `import type { WriteSheet } from "hucre/xlsx"` works
// without a second import from the root, which would pull the whole
// type graph back in and defeat the point of a format subpath.
export type {
  Cell,
  CellStyle,
  CellValue,
  ColumnDef,
  ConditionalRule,
  DataValidation,
  MergeRange,
  ReadInput,
  ReadOptions,
  Sheet,
  SheetChart,
  Sparkline,
  SheetTextBox,
  TableDefinition,
  Workbook,
  WorkbookProperties,
  WriteOptions,
  WriteSheet,
} from "./_types"

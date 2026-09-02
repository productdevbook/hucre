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
export {
  XlsxStreamWriter,
  writeXlsxStream,
  writeXlsxStreamSheets,
  XLSX_MAX_ROWS_PER_SHEET,
} from "./xlsx/stream-writer"
export type {
  XlsxStreamWriterOptions,
  XlsxWriteStreamOptions,
  XlsxWriteStreamWorkbookOptions,
  XlsxStreamRow,
  XlsxStreamSheet,
} from "./xlsx/stream-writer"

// ── Sizing & theme helpers ─────────────────────────────────────────
export { cloneCellStyle } from "./_style"
export { toWriteOptions, toWriteSheet } from "./write-model"
export type { WriteModelDrop, ToWriteOptionsOptions } from "./write-model"
export { calculateColumnWidth, measureValueWidth } from "./xlsx/auto-width"
export { calculateRowHeight } from "./xlsx/auto-size"

// ── Cell Utilities ─────────────────────────────────────────────────
//
// All nine, from one module. `hucre/xlsx` used to carry four of them and
// the root the other five, so anyone here who wanted `letterToCol` — a
// pure string helper with nothing XLSX-specific about it — had to pull a
// second entry point for it. The JSON surface had exactly this
// disagreement and it was settled before v1; this one was missed. See
// #474.
export {
  parseCellRef,
  colToLetter,
  letterToCol,
  cellRef,
  rangeRef,
  parseRange,
  isInRange,
  r1c1ToA1,
  a1ToR1C1,
  toRange,
  toRanges,
} from "./cell-utils"
export type { RangeLike } from "./cell-utils"

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
  XlsxReadOptions,
  XlsbReadOptions,
  ReadWarning,
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

// A cell may hold an error value; every writer takes one, and the spreadsheet readers produce them.
export { cellError, isCellError } from "./cell-error"
export type { CellError, CellErrorCode } from "./cell-error"

// What a writer takes where a cell goes: a value, or a cell object.
export type { CellInput } from "./_types"

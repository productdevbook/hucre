export { readXlsx } from "./xlsx/reader"
export { writeXlsx } from "./xlsx/writer"
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
  XlsxWriteStreamOptions,
  XlsxStreamRow,
} from "./xlsx/stream-writer"

// ── Cell Utilities ─────────────────────────────────────────────────
export { parseCellRef } from "./xlsx/worksheet"
export { colToLetter, cellRef, rangeRef } from "./xlsx/worksheet-writer"

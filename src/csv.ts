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

// ── Shared types used by this entry point's signatures ──────────────
export type { CellValue, CsvReadOptions, CsvWriteOptions } from "./_types"

// A cell may hold an error value; every writer takes one, and the spreadsheet readers produce them.
export { cellError, isCellError } from "./cell-error"
export type { CellError, CellErrorCode } from "./cell-error"

// Every stream*Rows reader yields this one shape.
export type { StreamRow } from "./_types"

// ── hucre/json entry point ────────────────────────────────────────────
// Read & write JSON arrays / NDJSON as tabular Workbook/Sheet rows.

export { parseJson, parseValue, parseNdjson, jsonToWorkbook } from "./json/reader"
export type {
  JsonReadOptions,
  JsonReadResult,
  JsonToWorkbookOptions,
  NdjsonReadOptions,
} from "./json/reader"

export { writeJson, writeNdjson, workbookToJson } from "./json/writer"
export type { JsonWriteOptions, WorkbookToJsonOptions } from "./json/writer"

export { NdjsonStreamWriter, streamNdjsonRows, writeNdjsonStream } from "./json/stream"
export type {
  NdjsonStreamReadOptions,
  NdjsonStreamWriterOptions,
  NdjsonStreamRow,
} from "./json/stream"

export { flattenValue, collectHeaders } from "./json/flatten"
export type { FlattenOptions } from "./json/flatten"

export { unflattenRow, unflattenRows } from "./json/unflatten"
export type { UnflattenedRow } from "./json/unflatten"

// ── Shared types used by this entry point's signatures ──────────────
export type { CellValue, Sheet, Workbook } from "./_types"

// A cell may hold an error value; every writer takes one, and the spreadsheet readers produce them.
export { cellError, isCellError } from "./cell-error"
export type { CellError, CellErrorCode } from "./cell-error"

// Every stream*Rows reader yields this one shape.
export type { StreamRow } from "./_types"

// What a writer takes where a cell goes: a value, or a cell object.
export type { CellInput } from "./_types"

// ── hucre/ods entry point ─────────────────────────────────────────────
// Read & write OpenDocument Spreadsheet (.ods) files.

export { readOds } from "./ods/reader"
export { writeOds } from "./ods/writer"
export { writeOdsStream } from "./ods/stream-writer"
export { OdsStreamWriter } from "./ods/incremental-writer"
export type {
  OdsStreamWriterOptions,
  OdsStyledCell,
  OdsIncrementalCell,
} from "./ods/incremental-writer"
export type { OdsWriteRow, OdsWriteCell, OdsStreamWriteOptions } from "./ods/stream-writer"
export { streamOdsRows } from "./ods/stream"
export { readOdsObjects, writeOdsObjects } from "./ods/objects"
export { toWriteOptions, toWriteSheet } from "./write-model"
export type { WriteModelDrop, ToWriteOptionsOptions } from "./write-model"
export type { OdsObjectsReadOptions, OdsObjectsResult, OdsObjectsWriteOptions } from "./ods/objects"

// ── Shared types used by this entry point's signatures ──────────────
// Re-exported so `import type { WriteSheet } from "hucre/ods"` works
// without a second import from the root, which would pull the whole
// type graph back in and defeat the point of a format subpath.
export type {
  Cell,
  CellStyle,
  CellValue,
  ColumnDef,
  MergeRange,
  ReadInput,
  OdsReadOptions,
  Sheet,
  Workbook,
  WorkbookProperties,
  WriteOptions,
  WriteSheet,
} from "./_types"

// A cell may hold an error value; every writer takes one, and the spreadsheet readers produce them.
export { cellError, isCellError } from "./cell-error"
export type { CellError, CellErrorCode } from "./cell-error"

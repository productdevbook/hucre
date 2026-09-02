// ── hucre/xml entry point ─────────────────────────────────────────────
// Read & write tabular XML (product feeds, ERP exports, GS1, etc.).

export { readXml } from "./xml/data-reader"
export type { XmlReadOptions, XmlReadResult } from "./xml/data-reader"

export { writeXml, writeXmlStream } from "./xml/data-writer"
export { streamXmlRows } from "./xml/stream-reader"
export type { XmlStreamRow, XmlStreamReadOptions } from "./xml/stream-reader"
export type { XmlWriteOptions } from "./xml/data-writer"

// ── Shared types used by this entry point's signatures ──────────────
export type { CellValue } from "./_types"

// A cell may hold an error value; every writer takes one, and the spreadsheet readers produce them.
export { cellError, isCellError } from "./cell-error"
export type { CellError, CellErrorCode } from "./cell-error"

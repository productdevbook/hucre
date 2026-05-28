// ── hucre/json entry point ────────────────────────────────────────────
// Read & write JSON arrays / NDJSON as tabular Workbook/Sheet rows.

export { parseJson, parseValue, parseNdjson } from "./json/reader"
export type { JsonReadOptions, JsonReadResult, NdjsonReadOptions } from "./json/reader"

export { writeJson, writeNdjson, workbookToJson } from "./json/writer"
export type { JsonWriteOptions, WorkbookToJsonOptions } from "./json/writer"

export { NdjsonStreamWriter, readNdjsonStream } from "./json/stream"
export type { NdjsonStreamReadOptions } from "./json/stream"

export { flattenValue, collectHeaders } from "./json/flatten"
export type { FlattenOptions } from "./json/flatten"

// ── True Streaming ODS Writer ───────────────────────────────────────
//
// ODS was the one that stood out. It has a streaming *reader* and had no
// streaming writer at all, so the format with the second-best support in
// the library could not produce a large file without holding the whole
// thing in memory. The ZIP layer already streams and `zipStream` is
// format-agnostic; what was missing was a row serializer that can be
// driven incrementally, the shape `RowSerializer` has for XLSX. See #467.
//
// What this deliberately does not do is styles. ODF puts
// `<office:automatic-styles>` *before* the body, so a style discovered
// while serializing row 900,000 has nowhere to go — the same shape as the
// shared-string table the XLSX streaming writer answers with inline
// strings, and ODF has no inline equivalent. Column widths are the
// exception and are carried, because `columns` is known before the first
// row. Everything else is values, which is what a million-row export is.

import { isCellError } from "../cell-error"
import type { CellValue, WorkbookProperties, CellInput } from "../_types"
import { zipStream, type ZipStreamEntry } from "../zip/stream-writer"
import { xmlEscape, xmlEscapeAttr } from "../xml/writer"
import { validateSheetNames } from "../_validate"

import {
  MIMETYPE,
  writeManifestXml,
  writeMetaXml,
  writeSettingsXml,
  writeStylesXml,
  formatOdsDateValue,
  formatNumberDisplay,
  excelFormulaToOds,
  odsEscape,
} from "./writer"

const encoder = /* @__PURE__ */ new TextEncoder()

/** A streamed row: positional values, each optionally carrying a formula. */

export interface OdsStreamWriteOptions {
  /** Sheet name. Excel's limits apply — LibreOffice enforces them too. */
  name?: string
  /**
   * Column widths in characters, and the header row to emit before the
   * data. Known before the first row, which is why these can be carried
   * when per-cell styles cannot.
   */
  columns?: Array<{ header?: string; width?: number }>
  /** Document properties written to `meta.xml`. */
  properties?: WorkbookProperties
  /**
   * Emit ZIP64 records, lifting the 4 GiB ceiling. Handed straight to
   * `zipStream`; see its note on why this is an up-front choice.
   */
  zip64?: boolean
}

/**
 * Write an ODS document as a byte stream, pulling rows from `rows` only
 * as the consumer reads.
 *
 * ```ts
 * return new Response(writeOdsStream(rowCursor, { name: "Export" }), {
 *   headers: {
 *     "content-type": "application/vnd.oasis.opendocument.spreadsheet",
 *   },
 * })
 * ```
 *
 * Peak memory is independent of the row count: each row is serialized,
 * encoded and enqueued on its own, and nothing is retained.
 *
 * **Values, not formatting.** `<office:automatic-styles>` precedes the
 * body in ODF, so a style first seen on row 900,000 has nowhere to be
 * declared. Column widths and a header row are carried because they are
 * known up front; per-cell styles are not, and `writeOds` remains the
 * path for a document that needs them. See #467.
 */
export function writeOdsStream(
  rows: AsyncIterable<CellInput[]> | Iterable<CellInput[]>,
  options?: OdsStreamWriteOptions,
): ReadableStream<Uint8Array> {
  const name = options?.name ?? "Sheet1"
  validateSheetNames([{ name }])

  const entries: ZipStreamEntry[] = [
    // mimetype MUST be first and MUST be stored uncompressed — the one
    // rule an ODF consumer checks before anything else.
    { path: "mimetype", data: encoder.encode(MIMETYPE), compress: false },
    { path: "META-INF/manifest.xml", data: encoder.encode(writeManifestXml()) },
    { path: "styles.xml", data: encoder.encode(writeStylesXml()) },
    { path: "meta.xml", data: encoder.encode(writeMetaXml(options?.properties)) },
    { path: "settings.xml", data: encoder.encode(writeSettingsXml()) },
    { path: "content.xml", data: contentChunks(rows, name, options?.columns) },
  ]

  return zipStream(entries, { zip64: options?.zip64 })
}

const CONTENT_HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
  ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' +
  ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
  ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
  ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"' +
  ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
  ' xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"' +
  ' xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"' +
  ' office:version="1.3">'

/** Serialize content.xml into ~64 KB encoded chunks, pulling lazily. */
async function* contentChunks(
  rows: AsyncIterable<CellInput[]> | Iterable<CellInput[]>,
  name: string,
  columns?: Array<{ header?: string; width?: number }>,
): AsyncGenerator<Uint8Array> {
  const CHUNK_BYTES = 64 * 1024
  let pending: string[] = []
  let pendingBytes = 0

  const push = function* (text: string): Generator<Uint8Array> {
    pending.push(text)
    pendingBytes += text.length
    if (pendingBytes >= CHUNK_BYTES) {
      yield encoder.encode(pending.join(""))
      pending = []
      pendingBytes = 0
    }
  }

  yield* push(CONTENT_HEAD)
  yield* push(automaticStyles(columns))
  yield* push(`<office:body><office:spreadsheet><table:table table:name="${xmlEscapeAttr(name)}">`)

  // Column declarations, which have to precede every row.
  const colCount = columns?.length ?? 0
  for (let i = 0; i < colCount; i++) {
    const width = columns![i]!.width
    yield* push(
      width === undefined
        ? "<table:table-column/>"
        : `<table:table-column table:style-name="co${i + 1}"/>`,
    )
  }

  if (columns?.some((c) => c.header !== undefined)) {
    yield* push(serializeRow(columns.map((c) => c.header ?? null)))
  }

  for await (const row of rows) {
    yield* push(serializeRow(row))
  }

  yield* push("</table:table></office:spreadsheet></office:body></office:document-content>")

  if (pending.length > 0) yield encoder.encode(pending.join(""))
}

/**
 * The style block, emitted before the body because ODF requires it there.
 *
 * Only column widths land here — they are the one thing known before the
 * first row arrives.
 */
function automaticStyles(columns?: Array<{ header?: string; width?: number }>): string {
  const parts: string[] = []
  columns?.forEach((col, i) => {
    if (col.width === undefined) return
    // ODF column widths are a physical measure; the same 7px-per-character
    // approximation the buffered writer uses, at 96 DPI.
    const inches = (col.width * 7 + 5) / 96
    parts.push(
      `<style:style style:name="co${i + 1}" style:family="table-column">` +
        `<style:table-column-properties style:column-width="${inches.toFixed(4)}in"/>` +
        "</style:style>",
    )
  })
  return `<office:automatic-styles>${parts.join("")}</office:automatic-styles>`
}

/** One `<table:table-row>`, values only. */
function serializeRow(row: CellInput[]): string {
  const cells: string[] = []
  for (const cell of row) {
    cells.push(
      cell !== null &&
        typeof cell === "object" &&
        !(cell instanceof Date) &&
        !isCellError(cell) &&
        !Array.isArray(cell)
        ? serializeCell(
            (cell as { value?: CellValue }).value ?? null,
            (cell as { formula?: string }).formula,
          )
        : serializeCell(cell as CellValue),
    )
  }
  return `<table:table-row>${cells.join("")}</table:table-row>`
}

/**
 * One `<table:table-cell>`.
 *
 * The value encoding matches `cellToOds` exactly, including its two
 * refusals: a non-finite number and an unparseable Date both produce an
 * empty cell rather than `office:value="NaN"`, which LibreOffice reads as
 * garbage and which used to make a corrupt file (#364).
 */
function serializeCell(value: CellValue, formula?: string): string {
  const attrs = formula ? ` table:formula="${xmlEscapeAttr(excelFormulaToOds(formula))}"` : ""

  if (value === null || value === undefined) {
    return attrs ? `<table:table-cell${attrs}/>` : "<table:table-cell/>"
  }

  if (typeof value === "string") {
    return (
      `<table:table-cell${attrs} office:value-type="string">` +
      `<text:p>${odsEscape(value)}</text:p></table:table-cell>`
    )
  }

  if (typeof value === "number") {
    if (!Number.isFinite(value)) return `<table:table-cell${attrs}></table:table-cell>`
    return (
      `<table:table-cell${attrs} office:value-type="float" office:value="${value}">` +
      `<text:p>${odsEscape(formatNumberDisplay(value))}</text:p></table:table-cell>`
    )
  }

  if (typeof value === "boolean") {
    return (
      `<table:table-cell${attrs} office:value-type="boolean" ` +
      `office:boolean-value="${value ? "true" : "false"}">` +
      `<text:p>${value ? "TRUE" : "FALSE"}</text:p></table:table-cell>`
    )
  }

  if (isCellError(value)) {
    return (
      `<table:table-cell${attrs} office:value-type="string" calcext:value-type="error">` +
      `<text:p>${xmlEscape(value.error)}</text:p></table:table-cell>`
    )
  }

  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return `<table:table-cell${attrs}></table:table-cell>`
    const iso = formatOdsDateValue(value)
    return (
      `<table:table-cell${attrs} office:value-type="date" office:date-value="${iso}">` +
      `<text:p>${iso}</text:p></table:table-cell>`
    )
  }

  return `<table:table-cell${attrs}/>`
}

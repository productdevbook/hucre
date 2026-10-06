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
// row. Values, formula caches, rich-text content and links share the buffered
// cell serializer; unsupported styles and cell fields can be reported via onDrop.

import { columnHeaders } from "../_sheet-input"
import { resolveCellInput } from "../_inline-cells"
import type { WorkbookProperties, CellInput, WorkbookWriteOptions, ColumnDef } from "../_types"
import { reportOdsCellDrops } from "./cell-drops"
import { zipStream, type ZipStreamEntry } from "../zip/stream-writer"
import { xmlEscapeAttr } from "../xml/writer"
import { validateSheetNames, validateRowSize } from "../_validate"

import {
  MIMETYPE,
  writeManifestXml,
  writeMetaXml,
  writeSettingsXml,
  writeStylesXml,
  cellToOds,
  createStyleCollector,
} from "./writer"

const encoder = /* @__PURE__ */ new TextEncoder()

/** A streamed row: positional values, each optionally carrying a formula. */

export interface OdsStreamWriteOptions extends WorkbookWriteOptions {
  /** Sheet name. Excel's limits apply — LibreOffice enforces them too. */
  name?: string
  /**
   * Column widths in characters, and the header row to emit before the
   * data. Known before the first row, which is why these can be carried
   * when per-cell styles cannot.
   */
  columns?: Array<Pick<ColumnDef, "header" | "key" | "width">>
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
    { path: "content.xml", data: contentChunks(rows, name, options) },
  ]

  validateRowSize(0, options?.columns?.length ?? 0)
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
  ' xmlns:xlink="http://www.w3.org/1999/xlink"' +
  ' xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"' +
  ' office:version="1.3">'

/** Serialize content.xml into ~64 KB encoded chunks, pulling lazily. */
async function* contentChunks(
  rows: AsyncIterable<CellInput[]> | Iterable<CellInput[]>,
  name: string,
  options?: OdsStreamWriteOptions,
): AsyncGenerator<Uint8Array> {
  const columns = options?.columns
  let rowIndex = 0
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

  const headers = columnHeaders(columns)
  if (headers) yield* push(serializeRow(headers, name, rowIndex++, options?.onDrop))

  for await (const row of rows) {
    yield* push(serializeRow(row, name, rowIndex++, options?.onDrop))
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

// Style-free ODS uses the same value/formula/link encoding as buffered
// output. A single collector stays empty because run fonts are removed;
// otherwise newly discovered styles would reference an already-emitted block.
const unstyledCollector = createStyleCollector()

function serializeRow(
  row: CellInput[],
  name: string,
  rowIndex: number,
  onDrop?: WorkbookWriteOptions["onDrop"],
): string {
  validateRowSize(rowIndex, row.length)
  const cells: string[] = []
  for (let col = 0; col < row.length; col++) {
    const cell = resolveCellInput(row[col])
    reportOdsCellDrops(cell, onDrop, name, rowIndex, col, true)
    if (cell.richText?.some((run) => run.font)) {
      cell.richText = cell.richText.map((run) => ({ text: run.text }))
    }
    cells.push(cellToOds(cell.value, { cellOverride: cell }, unstyledCollector))
  }
  return `<table:table-row>${cells.join("")}</table:table-row>`
}

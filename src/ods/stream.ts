// ── Streaming ODS Reader ─────────────────────────────────────────────
// Yields rows one at a time from an ODS file via SAX parsing.

import type { CellValue, ReadInput, StreamRow } from "../_types"
import { ParseError, ZipError } from "../errors"
import { assertNotEncrypted, readInputToUint8Array } from "../_input"
import { ZipReader } from "../zip/reader"
import { parseSax } from "../xml/parser"
import { parseOdsDateTime } from "./reader"
import { MAX_COL_INDEX, MAX_REPEAT_COUNT, MAX_ROW_INDEX } from "../limits"

// ── Helpers ──────────────────────────────────────────────────────────

function decodeUtf8(data: Uint8Array): string {
  return new TextDecoder("utf-8").decode(data)
}

/**
 * Options for {@link streamOdsRows}.
 *
 * It previously took none at all, so `maxRows` on a huge file meant
 * draining the whole thing and counting yourself.
 */
export interface OdsStreamReadOptions {
  /**
   * Restrict streaming to these sheets, by 0-based index or by name.
   * Rows from other sheets are skipped.
   */
  sheets?: Array<number | string>
  /** Stop after this many rows across all streamed sheets. */
  maxRows?: number
  /**
   * Zip-bomb ceiling for any one entry; see `ReadOptions.maxDecompressedBytes`.
   * Default: 2 GiB ({@link MAX_DECOMPRESSED_BYTES}).
   */
  maxDecompressedBytes?: number
}

/**
 * Resolve the sheet filter to a set of indices, or `undefined` for "all".
 *
 * Names cannot be resolved from the row stream alone — the SAX pass does
 * not surface table names — so only numeric entries are honoured for now
 * and a name-only filter falls back to streaming everything rather than
 * silently yielding nothing.
 */
function normalizeSheetFilter(sheets: Array<number | string> | undefined): Set<number> | undefined {
  if (!sheets || sheets.length === 0) return undefined
  const indices = sheets.filter((s): s is number => typeof s === "number")
  return indices.length > 0 ? new Set(indices) : undefined
}

// ── Row parser via SAX ──────────────────────────────────────────────

/**
 * @deprecated Use {@link StreamRow}. Kept as an alias because it was the
 * element type of a public async generator; the two shapes were
 * identical apart from `sheetIndex` now being optional.
 */
export type OdsStreamRow = StreamRow

function* parseContentRows(xml: string): Generator<StreamRow, void, undefined> {
  const completedRows: StreamRow[] = []

  let inBody = false
  let inSpreadsheet = false
  let inTable = false
  let inRow = false
  let inCell = false
  let inP = false
  let inAnnotation = false

  let sheetIndex = -1
  let currentRowIndex = -1
  let cellRepeat = 1
  let rowRepeat = 1
  let currentCells: CellValue[] = []
  let cellText = ""
  let cellValueType = ""
  let cellValue = ""
  let cellBoolValue = ""
  let cellDateValue = ""

  parseSax(xml, {
    onOpenTag(tag, attrs) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag

      switch (local) {
        case "body":
          inBody = true
          break
        case "spreadsheet":
          if (inBody) inSpreadsheet = true
          break
        case "table":
          if (inSpreadsheet) {
            inTable = true
            sheetIndex++
            currentRowIndex = -1
          }
          break
        case "table-row":
          if (inTable) {
            inRow = true
            // Clamp against a hostile huge number-rows-repeated on a row.
            rowRepeat = Math.min(
              Number(attrs["table:number-rows-repeated"] ?? "1"),
              MAX_ROW_INDEX + 1,
            )
            currentCells = []
          }
          break
        case "table-cell":
          if (inRow) {
            inCell = true
            cellRepeat = Math.min(
              Number(attrs["table:number-columns-repeated"] ?? "1"),
              MAX_COL_INDEX + 1,
            )
            cellText = ""
            cellValueType = attrs["office:value-type"] ?? attrs["calcext:value-type"] ?? ""
            cellValue = attrs["office:value"] ?? ""
            cellBoolValue = attrs["office:boolean-value"] ?? ""
            cellDateValue = attrs["office:date-value"] ?? ""
          }
          break
        case "covered-table-cell":
          if (inRow) {
            const repeat = Math.min(
              Number(attrs["table:number-columns-repeated"] ?? "1"),
              MAX_COL_INDEX + 1,
            )
            for (let i = 0; i < repeat; i++) {
              currentCells.push(null)
            }
          }
          break
        case "annotation":
          // A cell comment carries its own <text:p>. The batch reader
          // takes only direct children of the cell, so folding the
          // annotation into the value made the two readers disagree —
          // the divergence test/ods-stream-parity.test.ts exists to
          // catch. See #393.
          inAnnotation = true
          break
        case "p":
          if (inCell && !inAnnotation) inP = true
          break
        // Text content special elements — mirror collectText() in reader.ts so
        // the streaming and batch readers return the same string for a cell.
        case "s":
          if (inP && inCell) {
            // Same cap as the batch reader — an uncapped text:c reaches
            // a raw RangeError. See #363.
            const raw = Number(attrs["text:c"] ?? "1")
            const count =
              !Number.isFinite(raw) || raw < 1 ? 1 : Math.min(Math.trunc(raw), MAX_REPEAT_COUNT)
            cellText += " ".repeat(count)
          }
          break
        case "line-break":
          if (inP && inCell) cellText += "\n"
          break
        case "tab":
          if (inP && inCell) cellText += "\t"
          break
      }
    },

    onText(text) {
      if (inP && inCell) {
        cellText += text
      }
    },

    onCloseTag(tag) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag

      switch (local) {
        case "annotation":
          inAnnotation = false
          break
        case "p":
          inP = false
          break
        case "table-cell":
          if (inCell) {
            const value = resolveCellValue(
              cellValueType,
              cellValue,
              cellBoolValue,
              cellDateValue,
              cellText,
            )
            for (let i = 0; i < cellRepeat; i++) {
              currentCells.push(value)
            }
            inCell = false
          }
          break
        case "table-row":
          if (inRow) {
            // Trim trailing nulls
            while (currentCells.length > 0 && currentCells[currentCells.length - 1] === null) {
              currentCells.pop()
            }

            if (currentCells.length > 0) {
              // Cap row repeat to avoid memory explosion for empty trailing rows
              const effectiveRepeat = Math.min(rowRepeat, 1)
              for (let r = 0; r < (currentCells.length > 0 ? rowRepeat : effectiveRepeat); r++) {
                currentRowIndex++
                completedRows.push({
                  index: currentRowIndex,
                  sheetIndex,
                  values: r === 0 ? currentCells : [...currentCells],
                })
              }
            } else {
              currentRowIndex += rowRepeat
            }
            inRow = false
          }
          break
        case "table":
          inTable = false
          break
        case "spreadsheet":
          inSpreadsheet = false
          break
        case "body":
          inBody = false
          break
      }
    },
  })

  for (const row of completedRows) {
    yield row
  }
}

function resolveCellValue(
  valueType: string,
  value: string,
  boolValue: string,
  dateValue: string,
  text: string,
): CellValue {
  switch (valueType) {
    case "float":
    case "currency":
    case "percentage":
      if (value) return Number(value)
      return null
    case "boolean":
      if (boolValue === "true") return true
      if (boolValue === "false") return false
      return null
    case "date":
      // Same UTC reading as readOds — a streamed row must not disagree with
      // the same file read whole. See #415.
      if (dateValue) return parseOdsDateTime(dateValue) ?? null
      return null
    case "string":
      return text || ""
    default:
      return text || null
  }
}

// ── Main streaming reader ───────────────────────────────────────────

/**
 * Create an async iterable that yields rows one at a time from an ODS file.
 * Unzips and parses content.xml with SAX, yielding rows as they are parsed.
 */
export async function* streamOdsRows(
  input: ReadInput,
  options?: OdsStreamReadOptions,
): AsyncGenerator<StreamRow, void, undefined> {
  // Previously Uint8Array | ArrayBuffer only — a streaming reader that
  // could not take a ReadableStream, unlike streamXlsxRows. See #365.
  const data = await readInputToUint8Array(input)

  // Detect password-protected ODF workbooks (OLE2/CFB envelope) up
  // front so streamers fail fast with a typed `EncryptedFileError`
  // instead of a generic ZIP ParseError. Decryption is tracked in #156.
  assertNotEncrypted(data, "ods")

  // 1. Open ZIP archive
  let zip: ZipReader
  try {
    zip = new ZipReader(data, options?.maxDecompressedBytes)
  } catch (err) {
    if (err instanceof ZipError) throw err
    throw new ParseError("Failed to open ODS file: not a valid ZIP archive", undefined, {
      cause: err,
    })
  }

  // 2. Validate mimetype
  if (!zip.has("mimetype")) {
    throw new ParseError("Invalid ODS: missing 'mimetype' entry.")
  }

  // 3. Parse content.xml
  if (!zip.has("content.xml")) {
    throw new ParseError("Invalid ODS: missing content.xml")
  }
  const contentXml = decodeUtf8(await zip.extract("content.xml"))

  // 4. Yield rows via SAX
  // 4. Yield rows via SAX, applying the filters
  const wanted = normalizeSheetFilter(options?.sheets)
  const maxRows = options?.maxRows ?? 0
  let emitted = 0

  for (const row of parseContentRows(contentXml)) {
    if (wanted !== undefined && !wanted.has(row.sheetIndex ?? 0)) continue
    if (maxRows > 0 && emitted >= maxRows) return
    emitted++
    yield row
  }
}

// ── Ergonomic API ───────────────────────────────────────────────────
// Unified high-level functions that wrap the format-specific readers/writers.
// Auto-detects format from content (magic bytes) for reading, and dispatches
// to the correct writer based on the `format` option for writing.
// ─────────────────────────────────────────────────────────────────────

import { assertGridSize, padToRectangle } from "./_grid"
import type {
  Workbook,
  ReadOptions,
  WorkbookInput,
  XlsxWriteOptions,
  WriteOutput,
  CellValue,
  ReadInput,
  TableDefinition,
  TableColumn,
  CsvWriteOptions,
} from "./_types"
import type { JsonWriteOptions } from "./json/writer"
import type { XmlWriteOptions } from "./xml/data-writer"
import type { HtmlExportOptions } from "./export/html"
import type { MarkdownExportOptions } from "./export/markdown"
import {
  collectHeaders,
  rowsToObjects,
  selectSheet,
  type RowsToObjectsOptions,
  type ObjectsResult,
} from "./_objects"
import { readXlsx } from "./xlsx/reader"
import { readXlsb, looksLikeXlsb } from "./xlsx/xlsb/reader"
import { readXls, looksLikeXls } from "./xls/reader"
import { readCfb } from "./xlsx/crypto/cfb"
import { ZipReader } from "./zip/reader"
import { decryptAgile } from "./xlsx/crypto/agile"
import { writeXlsx } from "./xlsx/writer"
import { readOds } from "./ods/reader"
import { writeOds } from "./ods/writer"
import { EncryptedFileError, UnsupportedFormatError } from "./errors"
import { isOle2Container, readInputToUint8Array } from "./_input"
import { detectTextFormat, type TextFormat } from "./_sniff"
import { parseCsv } from "./csv/reader"
import { writeCsv } from "./csv/writer"
import { writeTsv } from "./export/tsv"
import { jsonToWorkbook, parseNdjson } from "./json/reader"
import { writeJson, writeNdjson } from "./json/writer"
import { readXml } from "./xml/data-reader"
import { writeXml } from "./xml/data-writer"
import { fromHtml } from "./export/html-import"
import { toHtml } from "./export/html"
import { toMarkdown } from "./export/markdown"
import { denseValues, sheetGrid, type GridProjectionOptions } from "./_sheet-grid"
import { toRange } from "./cell-utils"
import { toCellValues } from "./_inline-cells"
import { normalizeSheetInput } from "./_sheet-input"

// ── Format Detection ────────────────────────────────────────────────

/**
 * Detect whether a ZIP archive is XLSX or ODS by inspecting the first
 * local file entry. ODS archives store "mimetype" as the first file
 * with content "application/vnd.oasis.opendocument.spreadsheet".
 * XLSX archives are also ZIP but never have "mimetype" as the first entry.
 */
function detectFormat(data: Uint8Array): "xlsx" | "ods" {
  // Both XLSX and ODS start with PK (ZIP magic: 0x504B0304)
  if (data.length < 4 || data[0] !== 0x50 || data[1] !== 0x4b) {
    throw new UnsupportedFormatError("unknown (not a ZIP archive)")
  }

  // Read the first local file header to get the filename
  // Local file header: offset 26 = filename length (2 bytes LE), offset 30+ = filename
  if (data.length < 30) {
    throw new UnsupportedFormatError("unknown (ZIP too short)")
  }

  const filenameLen = data[26]! | (data[27]! << 8)
  if (data.length < 30 + filenameLen) {
    throw new UnsupportedFormatError("unknown (ZIP truncated)")
  }

  const decoder = new TextDecoder("utf-8")
  const firstName = decoder.decode(data.subarray(30, 30 + filenameLen))

  if (firstName === "mimetype") {
    // Read the extra field length to find where file data starts
    const extraLen = data[28]! | (data[29]! << 8)
    const dataOffset = 30 + filenameLen + extraLen

    // Read the uncompressed size from the local header (offset 22, 4 bytes LE)
    const uncompSize = data[22]! | (data[23]! << 8) | (data[24]! << 16) | (data[25]! << 24)

    if (uncompSize > 0 && data.length >= dataOffset + uncompSize) {
      const mimeContent = decoder.decode(data.subarray(dataOffset, dataOffset + uncompSize))
      if (mimeContent.trim() === "application/vnd.oasis.opendocument.spreadsheet") {
        return "ods"
      }
    }

    // Even if we couldn't read the content, "mimetype" as first entry is ODS convention
    return "ods"
  }

  // Default: assume XLSX for any other ZIP
  return "xlsx"
}

// ── Public API ──────────────────────────────────────────────────────

/**
 * Read any supported spreadsheet file. Auto-detects format from content.
 * Supports: XLSX, ODS (CSV uses parseCsv separately since it's string input).
 *
 * Input can be Uint8Array, ArrayBuffer, or ReadableStream&lt;Uint8Array&gt;.
 * ReadableStream input is buffered fully before format detection runs.
 */
export async function read(input: ReadInput, options?: ReadOptions): Promise<Workbook> {
  let data = await readInputToUint8Array(input, options?.maxInputBytes)

  // Password-protected workbooks arrive as an OLE2/CFB envelope. With a
  // password we decrypt the inner package (then `detectFormat` works on
  // the plaintext ZIP); without one we surface a typed error. The
  // container alone doesn't reveal XLSX vs ODS, so the no-password error
  // leaves `format` unset. (ODS uses a different in-ZIP scheme, so only
  // XLSX decryption is wired up — an ODS password yields a ZIP that
  // detectFormat routes to readOds, which handles its own encryption.)
  if (isOle2Container(data)) {
    // An OLE2 container is either a legacy .xls (BIFF "Workbook" stream)
    // or an encrypted OOXML/ODS package (an "EncryptionInfo" stream).
    let cfbStreams: Map<string, Uint8Array> | null = null
    try {
      cfbStreams = readCfb(data)
    } catch {
      cfbStreams = null // not a parseable CFB — treat as encrypted/unknown
    }
    if (cfbStreams && looksLikeXls(cfbStreams)) {
      return readXls(data, options)
    }
    if (options?.password) {
      data = await decryptAgile(data, options.password, options.maxSpinCount)
    } else {
      throw new EncryptedFileError()
    }
  }

  // Not a container. Every text format the library reads announces
  // itself in its first non-whitespace character or two, so `read()` no
  // longer stops at the ZIP boundary. See #469.
  if (data.length < 4 || data[0] !== 0x50 || data[1] !== 0x4b) {
    const text = detectTextFormat(data)
    if (text !== null) return readTextFormat(data, text, options)
  }

  const format = detectFormat(data)

  if (format === "ods") {
    return readOds(data, options)
  }
  // XLSX and XLSB share the ZIP shape; tell them apart by the binary
  // workbook part before dispatching. A ZIP that is neither — a .docx, a
  // plain archive — used to be handed to readXlsx and fail with a
  // ParseError from deep inside it; it is refused here by name instead.
  let zip: ZipReader | null = null
  try {
    zip = new ZipReader(data)
  } catch {
    // Not a readable ZIP here — fall through to readXlsx for a typed error.
  }
  if (zip) {
    if (looksLikeXlsb(zip)) return readXlsb(data, options)
    if (!zip.has("[Content_Types].xml")) {
      throw new UnsupportedFormatError("zip (no [Content_Types].xml — not an Office package)")
    }
    const contentTypes = new TextDecoder("utf-8").decode(await zip.extract("[Content_Types].xml"))
    if (!contentTypes.includes("spreadsheetml")) {
      throw new UnsupportedFormatError("zip (an Office package, but not a spreadsheet)")
    }
  }
  return readXlsx(data, options)
}

/**
 * Turn a detected text format into a workbook.
 *
 * Each of these readers already exists and is exported; what was missing
 * was `read()` knowing to call one. The tabular readers hand back
 * `{ data, headers }`, so the header row is put back at the top — a
 * workbook is a grid, and dropping the names would lose them.
 */
function readTextFormat(data: Uint8Array, format: TextFormat, options?: ReadOptions): Workbook {
  const limits = { maxTotalCells: options?.maxTotalCells }
  switch (format) {
    case "csv":
      // `parseCsv` keeps a short line short; a `Sheet` is a rectangle.
      return {
        sheets: [{ name: "Sheet1", rows: padToRectangle(parseCsv(data), limits.maxTotalCells) }],
      }
    case "json":
      return jsonToWorkbook(data, limits)
    case "ndjson": {
      const { data: rows, headers } = parseNdjson(data, limits)
      return {
        sheets: [{ name: "Sheet1", rows: withHeaderRow(rows, headers, limits.maxTotalCells) }],
      }
    }
    case "xml": {
      const { data: rows, headers } = readXml(data, limits)
      return {
        sheets: [{ name: "Sheet1", rows: withHeaderRow(rows, headers, limits.maxTotalCells) }],
      }
    }
    case "html":
      return { sheets: [fromHtml(new TextDecoder("utf-8").decode(data), limits)] }
  }
}

/** Put the header names back at row 0, the way a grid holds them. */
function withHeaderRow(
  rows: Array<Record<string, CellValue>>,
  headers: string[],
  limit?: number,
): CellValue[][] {
  assertGridSize(rows.length + 1, headers.length, limit)
  return [headers, ...rows.map((row) => headers.map((h) => row[h] ?? null))]
}

/** Every format {@link write} can produce. */
export type WriteFormat =
  | "xlsx"
  | "ods"
  | "csv"
  | "tsv"
  | "json"
  | "ndjson"
  | "xml"
  | "html"
  | "markdown"

/**
 * Write a workbook to the specified format.
 *
 * The union used to be `"xlsx" | "ods"` while the library could write
 * nine things, so the one function meant to be format-agnostic covered
 * two of them. See #469.
 *
 * The text formats are single-sheet by nature and take the first sheet;
 * they also carry values and not formatting, which is the same trade
 * `hucre convert` documents. The return is always bytes, so a caller can
 * hand the result to `Response` or `writeFile` without branching.
 */
export interface TextFormatOptions extends GridProjectionOptions {
  /** Options for `format: "csv"`. */
  csv?: CsvWriteOptions
  /** Options for `format: "tsv"`. The delimiter is the tab and not yours. */
  tsv?: Omit<CsvWriteOptions, "delimiter">
  /** Options for `format: "json"`. */
  json?: JsonWriteOptions
  /** Options for `format: "ndjson"`. */
  ndjson?: Pick<JsonWriteOptions, "unflatten">
  /** Options for `format: "xml"`. */
  xml?: XmlWriteOptions
  /** Options for `format: "html"`. */
  html?: HtmlExportOptions
  /** Options for `format: "markdown"`. */
  markdown?: MarkdownExportOptions
}

export interface WriteFormatOptions extends XlsxWriteOptions, TextFormatOptions {
  format?: WriteFormat
}

export async function write(
  workbook: WorkbookInput,
  options: WriteFormatOptions = {},
): Promise<WriteOutput> {
  const format = options.format ?? "xlsx"
  if (format === "xlsx") return writeXlsx(workbook, options)
  if (format === "ods") return writeOds(workbook, options)

  const sheet = workbook.sheets[0]
  if (!sheet) {
    throw new UnsupportedFormatError(`${format} needs a sheet to write, and the workbook has none.`)
  }
  const input = normalizeSheetInput(sheet, (height, width) =>
    assertGridSize(height, width, options.maxTotalCells),
  )
  const values = { ...input, rows: toCellValues(input.rows ?? []) }
  // Value exports must apply the same inline/store/cache precedence as
  // spreadsheet writers. Reading just rows silently discards sparse data.
  const records = () => {
    const grid = sheetGrid(values)
    return rowsToObjects(values, {
      skipEmptyRows: false,
      maxTotalCells: options.maxTotalCells,
      // Keep the text writer's positional names for unnamed columns.
      transformHeader: (header, col) => (grid.value(0, col) === null ? `column${col + 1}` : header),
    }).data
  }

  // Each text writer already takes an options bag; this function used to
  // call every one of them with none, so `write` — the entry #469 added
  // precisely so one call could reach all nine formats — was the only way
  // to reach seven of them that could not configure any. `bom: true` was
  // the one that mattered: it is what makes Excel open a UTF-8 CSV on a
  // non-UTF-8 locale, and #475 documents it as the answer while `write`
  // gave no way to ask for it.
  const encoder = new TextEncoder()
  switch (format) {
    case "csv":
      return encoder.encode(writeCsv(denseValues(values, options.maxTotalCells), options.csv))
    case "tsv":
      return encoder.encode(writeTsv(denseValues(values, options.maxTotalCells), options.tsv))
    case "json":
      return encoder.encode(writeJson(records(), options.json))
    case "ndjson":
      return encoder.encode(writeNdjson(records(), options.ndjson))
    case "xml":
      return encoder.encode(writeXml(records(), options.xml))
    case "html":
      return encoder.encode(
        toHtml(
          {
            name: values.name,
            rows: values.rows,
            cells: values.cells,
            merges: values.merges?.map((merge) =>
              typeof merge === "string" ? toRange(merge) : merge,
            ),
          },
          { maxTotalCells: options.maxTotalCells, ...options.html },
        ),
      )
    case "markdown":
      return encoder.encode(
        toMarkdown(values, { maxTotalCells: options.maxTotalCells, ...options.markdown }),
      )
  }
}

/**
 * Options for {@link readObjects}.
 *
 * The projection knobs are the same set — and the same defaults — as
 * `XlsxObjectsReadOptions` and `OdsObjectsReadOptions`. They are applied
 * to the workbook `read()` returns, so every one of them is honoured for
 * every format `read()` can detect (XLSX, XLSB, XLS, ODS).
 *
 * The inherited {@link ReadOptions} fields are a different story: they are
 * handed to the format reader and are honoured as unevenly as ever (see
 * #365 item 4). `sheets` is omitted because {@link sheet} supersedes it.
 */
export interface ReadObjectsOptions extends Omit<ReadOptions, "sheets">, RowsToObjectsOptions {
  /** Sheet to read from. Index (0-based) or sheet name. Default: 0. */
  sheet?: number | string
}

/**
 * Result shape for {@link readObjects} — the same `{ data, headers }`
 * every other `*Objects` reader returns.
 */
export type ReadObjectsResult<T extends Record<string, CellValue> = Record<string, CellValue>> =
  ObjectsResult<T>

/**
 * Quick helper: read a file and get a sheet as objects keyed by a header
 * row, plus the detected headers.
 *
 * Format-agnostic counterpart to `readXlsxObjects` / `readOdsObjects` —
 * same options, same defaults, same `{ data, headers }` result.
 */
export async function readObjects<T extends Record<string, CellValue> = Record<string, CellValue>>(
  input: ReadInput,
  options?: ReadObjectsOptions,
): Promise<ReadObjectsResult<T>> {
  const {
    sheet: sheetSelector = 0,
    headerRow = 0,
    skipEmptyRows = true,
    transformHeader,
    transformValue,
    maxRows,
    ...readOpts
  } = options ?? {}

  const workbook = await read(input, readOpts)
  const sheet = selectSheet(workbook, sheetSelector)

  return rowsToObjects<T>(sheet, {
    headerRow,
    skipEmptyRows,
    transformHeader,
    transformValue,
    maxRows,
    maxTotalCells: readOpts.maxTotalCells,
  })
}

/** Options for writeObjects table generation */
export interface WriteObjectsTableOption {
  /** Table name (must be unique in workbook) */
  name: string
  /** Table style (e.g. "TableStyleMedium2") */
  style?: string
  /** Show totals row */
  showTotalRow?: boolean
  /** Show auto-filter. Default: true */
  showAutoFilter?: boolean
  /** Show banded rows. Default: true */
  showRowStripes?: boolean
  /** Totals per column key: { revenue: "sum", margin: "average" } */
  totals?: Record<
    string,
    "sum" | "average" | "count" | "min" | "max" | "countNums" | "stdDev" | "var"
  >
}

/**
 * Quick helper: write an array of objects to a spreadsheet format.
 * Infers column headers from the keys of the first object.
 */
export async function writeObjects(
  data: Array<Record<string, CellValue>>,
  options?: {
    sheetName?: string
    format?: "xlsx" | "ods"
    /** Wrap output in a native Excel table (ListObject) */
    table?: WriteObjectsTableOption
  },
): Promise<WriteOutput> {
  const sheetName = options?.sheetName ?? "Sheet1"
  const format = options?.format ?? "xlsx"

  if (data.length === 0) {
    return write(
      {
        sheets: [{ name: sheetName, rows: [] }],
      },
      {
        format,
      },
    )
  }

  // Column set is the union of every record's keys, not just the first's.
  const keys = collectHeaders(data)

  // Build rows: header row + data rows
  const rows: CellValue[][] = []

  // Header row
  rows.push(keys)

  // Data rows
  for (const item of data) {
    const row: CellValue[] = keys.map((key) => {
      const val = item[key]
      return val === undefined ? null : val
    })
    rows.push(row)
  }

  // Build Excel table if requested
  let tables: TableDefinition[] | undefined
  if (options?.table) {
    const t = options.table
    const colCount = keys.length
    const rowCount = data.length + 1 // +1 for header
    const endCol = colToLetterSimple(colCount - 1)
    const range = `A1:${endCol}${rowCount + (t.showTotalRow ? 1 : 0)}`

    const tableColumns: TableColumn[] = keys.map((key) => {
      const totalFn = t.totals?.[key]
      return {
        name: key,
        ...(totalFn ? { totalFunction: totalFn } : {}),
      }
    })

    tables = [
      {
        name: t.name,
        displayName: t.name,
        range,
        columns: tableColumns,
        style: t.style,
        showAutoFilter: t.showAutoFilter,
        showRowStripes: t.showRowStripes,
        showTotalRow: t.showTotalRow,
      },
    ]
  }

  return write(
    {
      sheets: [{ name: sheetName, rows, tables }],
    },
    {
      format,
    },
  )
}

/** Simple column index to letter (0-based) */
function colToLetterSimple(col: number): string {
  let result = ""
  let n = col
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result
    n = Math.floor(n / 26) - 1
  }
  return result
}

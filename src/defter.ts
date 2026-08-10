// ── Ergonomic API ───────────────────────────────────────────────────
// Unified high-level functions that wrap the format-specific readers/writers.
// Auto-detects format from content (magic bytes) for reading, and dispatches
// to the correct writer based on the `format` option for writing.
// ─────────────────────────────────────────────────────────────────────

import type {
  Workbook,
  ReadOptions,
  WriteOptions,
  WriteOutput,
  CellValue,
  ReadInput,
  TableDefinition,
  TableColumn,
} from "./_types"
import { collectHeaders, rowsToObjects, selectSheet } from "./_objects"
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
      data = await decryptAgile(data, options.password)
    } else {
      throw new EncryptedFileError()
    }
  }

  const format = detectFormat(data)

  if (format === "ods") {
    return readOds(data, options)
  }
  // XLSX and XLSB share the ZIP shape; tell them apart by the binary
  // workbook part before dispatching.
  try {
    if (looksLikeXlsb(new ZipReader(data))) return readXlsb(data, options)
  } catch {
    // Not a readable ZIP here — fall through to readXlsx for a typed error.
  }
  return readXlsx(data, options)
}

/**
 * Write a workbook to the specified format.
 */
export async function write(
  options: WriteOptions & { format?: "xlsx" | "ods" },
): Promise<WriteOutput> {
  const format = options.format ?? "xlsx"
  if (format === "ods") {
    return writeOds(options)
  }
  return writeXlsx(options)
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
export interface ReadObjectsOptions extends Omit<ReadOptions, "sheets"> {
  /** Sheet to read from. Index (0-based) or sheet name. Default: 0. */
  sheet?: number | string
  /** 0-based row index to use as headers. Default: 0. */
  headerRow?: number
  /** Skip rows where every cell is null/empty. Default: true. */
  skipEmptyRows?: boolean
  /** Transform header values (after String/trim normalization). */
  transformHeader?: (header: string, index: number) => string
  /** Transform each cell value. */
  transformValue?: (
    value: CellValue,
    header: string,
    rowIndex: number,
    colIndex: number,
  ) => CellValue
  /**
   * Maximum number of data rows to return (after the header row).
   *
   * Shadows `ReadOptions.maxRows` — this one is applied to the projected
   * rows for every format, rather than to the XLSX parse only. The
   * format reader is never handed a `maxRows`.
   */
  maxRows?: number
}

/**
 * Result shape for {@link readObjects} — the same `{ data, headers }`
 * every other `*Objects` reader returns.
 */
export interface ReadObjectsResult<
  T extends Record<string, CellValue> = Record<string, CellValue>,
> {
  data: T[]
  headers: string[]
}

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

  return rowsToObjects<T>(sheet.rows, {
    headerRow,
    skipEmptyRows,
    transformHeader,
    transformValue,
    maxRows,
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
    return write({
      sheets: [{ name: sheetName, rows: [] }],
      format,
    })
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

  return write({
    sheets: [{ name: sheetName, rows, tables }],
    format,
  })
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

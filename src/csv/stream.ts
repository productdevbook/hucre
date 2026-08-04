// ── CSV Streaming ────────────────────────────────────────────────────
// Stream CSV rows as a synchronous generator (line by line).
//
// Two writers share the row formatter below:
//
// • `CsvStreamWriter` — incremental. Formats each row on arrival but
//   retains every line until `finish()` joins them, so peak memory is
//   O(data).
//
// • `writeCsvStream()` — genuinely streaming. Rows are pulled from a
//   source on demand and encoded lines are flushed as they accumulate,
//   so peak memory is independent of the row count.

import type { CellValue, CsvReadOptions, CsvWriteOptions } from "../_types"
import { stripBom, detectDelimiter } from "./reader"
import { escapeFormula, unescapeFormula } from "./formula"
import { inferType } from "../_infer"

const TEXT_ENCODER = /* @__PURE__ */ new TextEncoder()

/** Flush the line accumulator once it crosses this many characters. */
const CHUNK_THRESHOLD = 64 * 1024

// ── Helpers ──────────────────────────────────────────────────────────

function startsWith(str: string, prefix: string, offset: number): boolean {
  if (offset + prefix.length > str.length) return false
  for (let i = 0; i < prefix.length; i++) {
    if (str[offset + i] !== prefix[i]) return false
  }
  return true
}

// ── Streaming CSV Reader ─────────────────────────────────────────────

/**
 * Stream CSV rows as a synchronous generator.
 * Processes the string incrementally and yields one row at a time.
 */
export function* streamCsvRows(
  input: string,
  options?: CsvReadOptions,
): Generator<CellValue[], void, undefined> {
  const skipBom = options?.skipBom !== false
  const quote = options?.quote ?? '"'
  const escape = options?.escape ?? '"'
  const doTypeInference = options?.typeInference ?? false
  const skipEmptyRows = options?.skipEmptyRows ?? false
  const commentChar = options?.comment
  const isHeaderMode = options?.header ?? false
  const skipHeaderRow = options?.skipHeaderRow ?? false
  const unescapeFormulae = options?.unescapeFormulae ?? false
  // Align inferType default with parseCsv (defaults to true).
  const preserveLeadingZeros = options?.preserveLeadingZeros !== false
  const maxRows = options?.maxRows
  const skipLines = options?.skipLines ?? 0
  const transformValue = options?.transformValue
  const onRow = options?.onRow
  // Fast mode trades quote handling for speed, exactly as in parseCsv —
  // fields are split on the delimiter with no quote awareness at all.
  const fastMode = options?.fastMode ?? false

  if (skipBom) {
    input = stripBom(input)
  }

  if (input.length === 0) return

  const delimiter = options?.delimiter ?? detectDelimiter(input)
  const len = input.length

  let i = 0
  let isFirstRow = true
  let headerRow: string[] | null = null
  let physicalLine = 0 // counts every parsed physical row (for skipLines)
  let emittedDataRows = 0 // counts rows yielded (for maxRows)

  while (i < len) {
    // Parse one row
    const row: string[] = []
    let currentField = ""
    let inQuoted = false
    let rowDone = false
    // Whether the current field was opened with a quote char.
    let fieldWasQuoted = false
    // Whether the FIRST field of this row was quoted.
    let rowFirstQuoted = false

    while (i < len && !rowDone) {
      const ch = input[i]!

      if (inQuoted) {
        // Check for escape sequence
        if (ch === escape && i + 1 < len && input[i + 1] === quote) {
          currentField += quote
          i += 2
          continue
        }
        // End of quoted field
        if (ch === quote) {
          inQuoted = false
          i++
          continue
        }
        // Any other character inside quotes
        currentField += ch
        i++
        continue
      }

      // Not in quoted field
      if (startsWith(input, delimiter, i)) {
        if (row.length === 0) rowFirstQuoted = fieldWasQuoted
        row.push(currentField)
        currentField = ""
        fieldWasQuoted = false
        i += delimiter.length
        continue
      }

      // Check for line endings
      if (ch === "\r") {
        if (row.length === 0) rowFirstQuoted = fieldWasQuoted
        row.push(currentField)
        currentField = ""
        fieldWasQuoted = false
        if (i + 1 < len && input[i + 1] === "\n") {
          i += 2
        } else {
          i++
        }
        rowDone = true
        continue
      }

      if (ch === "\n") {
        if (row.length === 0) rowFirstQuoted = fieldWasQuoted
        row.push(currentField)
        currentField = ""
        fieldWasQuoted = false
        i++
        rowDone = true
        continue
      }

      // Start of quoted field. In fast mode the quote char is just another
      // character, so `inQuoted` is never entered and the branches above
      // stay dead — matching parseFast's plain split.
      if (!fastMode && ch === quote && currentField === "") {
        inQuoted = true
        fieldWasQuoted = true
        i++
        continue
      }

      currentField += ch
      i++
    }

    // End of input without trailing newline.
    // Preserve a final row whose single field was an explicit quoted-empty
    // field ("").
    if (!rowDone) {
      if (currentField !== "" || row.length > 0 || fieldWasQuoted) {
        if (row.length === 0) rowFirstQuoted = fieldWasQuoted
        row.push(currentField)
      } else {
        // Nothing left
        break
      }
    }

    // Skip leading physical lines if configured
    physicalLine++
    if (physicalLine <= skipLines) continue

    // Skip empty rows if configured
    if (row.length === 0) continue
    if (skipEmptyRows && row.every((cell) => cell === "")) continue

    // Skip comment rows — only physically-unquoted leading comment chars count.
    if (commentChar && !rowFirstQuoted && row.length > 0 && row[0].startsWith(commentChar)) {
      continue
    }

    // Undo the writer's formula escape first, so type inference and header
    // names both see the value that was written, not `'` + the value.
    const fields = unescapeFormulae ? row.map(unescapeFormula) : row

    // Capture the header row. Like parseCsv, `header: true` only marks it —
    // the row is still yielded, and is used to name columns for
    // transformValue. `skipHeaderRow` is the opt-in that consumes it.
    const isHeaderRowNow = isFirstRow && isHeaderMode
    if (isHeaderRowNow) {
      headerRow = fields
    }
    isFirstRow = false

    if (isHeaderRowNow && skipHeaderRow) continue

    // Honor maxRows (counts data rows yielded)
    if (maxRows !== undefined && maxRows >= 0 && emittedDataRows >= maxRows) {
      return
    }
    const rowIndex = emittedDataRows
    emittedDataRows++

    // Apply type inference if requested
    let outRow: CellValue[] = doTypeInference
      ? fields.map((v) => inferType(v, preserveLeadingZeros))
      : fields

    // transformValue — after type inference, matching parseCsv's ordering.
    if (transformValue) {
      outRow = outRow.map((val, colIdx) =>
        transformValue(
          val,
          headerRow ? String(headerRow[colIdx] ?? colIdx) : String(colIdx),
          rowIndex,
          colIdx,
        ),
      )
    }

    onRow?.(outRow, rowIndex)

    yield outRow
  }
}

// ── Streaming CSV Writer ─────────────────────────────────────────────

const UTF8_BOM = "\uFEFF"

/**
 * Turns row values into a CSV line. Both writers share one of these so
 * their output stays character-identical.
 */
class CsvRowFormatter {
  readonly delimiter: string
  readonly lineSeparator: string
  readonly bom: boolean
  private quote: string
  private quoteStyle: "all" | "required" | "none"
  private dateFormat: string | undefined
  private nullValue: string
  private escapeFormulae: boolean
  private comment: string | undefined

  constructor(options?: CsvWriteOptions) {
    this.delimiter = options?.delimiter ?? ","
    this.lineSeparator = options?.lineSeparator ?? "\r\n"
    this.quote = options?.quote ?? '"'
    this.quoteStyle = options?.quoteStyle ?? "required"
    this.bom = options?.bom ?? false
    this.dateFormat = options?.dateFormat
    this.nullValue = options?.nullValue ?? ""
    // Both of these were honoured by writeCsv alone until #408, so the
    // same options produced different bytes depending on which writer you
    // reached for — and one of the two was the injection escape.
    this.escapeFormulae = options?.escapeFormulae ?? false
    // "" would make startsWith() true for every value, so treat it as unset.
    this.comment = options?.comment || undefined
  }

  /** Format one row of values into a delimited line. */
  formatRow(values: CellValue[]): string {
    return values.map((v) => this.formatAndQuote(v)).join(this.delimiter)
  }

  /** Format a header line — plain strings, quoted by the same rules. */
  formatHeader(headers: string[]): string {
    return headers.map((h) => this.quoteField(h)).join(this.delimiter)
  }

  private formatAndQuote(value: CellValue): string {
    if (value === null || value === undefined) {
      if (this.quoteStyle === "all") {
        return this.quote + this.nullValue + this.quote
      }
      return this.nullValue
    }

    if (typeof value === "boolean") {
      return this.quoteField(value ? "true" : "false")
    }

    if (typeof value === "number") {
      return this.quoteField(this.formatNumber(value))
    }

    if (value instanceof Date) {
      return this.quoteField(this.formatDate(value))
    }

    const str = String(value)
    return this.quoteField(this.escapeFormulae ? escapeFormula(str) : str)
  }

  private quoteField(value: string): string {
    if (this.quoteStyle === "none") {
      return value
    }

    const needsQuoting =
      this.quoteStyle === "all" ||
      value.includes(this.delimiter) ||
      value.includes(this.quote) ||
      value.includes("\n") ||
      value.includes("\r") ||
      // A leading comment character is quoted so a reader configured with
      // `comment` keeps the row rather than dropping the line (#408).
      (this.comment !== undefined && value.startsWith(this.comment))

    if (!needsQuoting) {
      return value
    }

    const escaped = value.replaceAll(this.quote, this.quote + this.quote)
    return this.quote + escaped + this.quote
  }

  private formatNumber(n: number): string {
    if (Number.isInteger(n) && Math.abs(n) >= 1e15) {
      return n.toFixed(0)
    }
    if (Math.abs(n) > 0 && Math.abs(n) < 1e-6) {
      return n.toFixed(20).replace(/0+$/, "").replace(/\.$/, ".0")
    }
    return String(n)
  }

  private formatDate(d: Date): string {
    if (!this.dateFormat) {
      // See #364 — an unparseable Date threw a raw RangeError, and in a
      // streaming writer that lands after bytes have gone to the client.
      if (Number.isNaN(d.getTime())) return ""
      return d.toISOString()
    }

    const year = d.getFullYear()
    const month = d.getMonth() + 1
    const day = d.getDate()
    const hours = d.getHours()
    const minutes = d.getMinutes()
    const seconds = d.getSeconds()

    return this.dateFormat
      .replace("YYYY", String(year))
      .replace("MM", String(month).padStart(2, "0"))
      .replace("DD", String(day).padStart(2, "0"))
      .replace("HH", String(hours).padStart(2, "0"))
      .replace("mm", String(minutes).padStart(2, "0"))
      .replace("ss", String(seconds).padStart(2, "0"))
  }
}

// ── Incremental CSV Writer (buffered) ────────────────────────────────

/**
 * Constructor options for {@link CsvStreamWriter} — the same options
 * `writeCsv` / `writeCsvStream` take.
 *
 * Exists as a name of its own so every stream writer in the library has
 * one (`XlsxStreamWriterOptions`, `NdjsonStreamWriterOptions`).
 */
export type CsvStreamWriterOptions = CsvWriteOptions

/**
 * Incremental CSV writer.
 *
 * Each `addRow()` is formatted immediately, but every line is retained
 * until {@link CsvStreamWriter.finish} joins them, so peak memory scales
 * with the data. For constant-memory output use {@link writeCsvStream}.
 */
export class CsvStreamWriter {
  private formatter: CsvRowFormatter
  private lineSeparator: string
  private bom: boolean
  private lines: string[] = []
  private headerWritten = false
  private headers: string[] | boolean | undefined
  /** Column order for object rows, resolved on the first one seen. */
  private columns: string[] | undefined

  constructor(options?: CsvStreamWriterOptions) {
    this.formatter = new CsvRowFormatter(options)
    this.lineSeparator = this.formatter.lineSeparator
    this.bom = this.formatter.bom
    this.headers = options?.headers
    this.columns = options?.columns ?? (Array.isArray(this.headers) ? this.headers : undefined)

    // Write header row immediately if string array provided
    if (Array.isArray(this.headers) && !this.headerWritten) {
      this.lines.push(this.formatter.formatHeader(this.headers))
      this.headerWritten = true
    }
  }

  /** Add a row of values */
  addRow(values: CellValue[]): void {
    this.lines.push(this.formatter.formatRow(values))
  }

  /**
   * Add a row from an object, projected through a column order resolved
   * exactly as {@link writeCsvStream} resolves it: `columns` if given,
   * else an explicit `headers` array, else the keys of the first object.
   *
   * A header line is emitted before the first object row unless one was
   * already written or `headers: false` was passed.
   */
  addObject(item: Record<string, CellValue>): void {
    if (!this.columns) {
      this.columns = Object.keys(item)
    }
    if (!this.headerWritten && this.headers !== false) {
      this.lines.push(this.formatter.formatHeader(this.columns))
      this.headerWritten = true
    }
    this.addRow(this.columns.map((key) => item[key] ?? null))
  }

  /** Finalize and return the CSV string */
  finish(): string {
    const parts: string[] = []

    if (this.bom) {
      parts.push(UTF8_BOM)
    }

    parts.push(this.lines.join(this.lineSeparator))

    return parts.join("")
  }

  /**
   * Emit the finished CSV as a `ReadableStream<Uint8Array>`.
   *
   * **This is not a constant-memory stream.** Every row added so far is
   * still buffered; {@link finish} runs first and the whole result is
   * enqueued as one chunk. It exists so a writer can be handed to a
   * `Response` body or a file sink without a manual encode step — not to
   * bound memory. For output whose peak memory is independent of the row
   * count, use {@link writeCsvStream}, which pulls rows from an iterable
   * and flushes as it goes.
   *
   * `finish()` runs when the stream is first read, so rows added between
   * `toStream()` and the first pull are still included, and the stream
   * closes right after — no separate `finish()` call is needed (though
   * one is harmless: `finish()` is idempotent here).
   */
  toStream(): ReadableStream<Uint8Array> {
    const finish = (): Uint8Array => TEXT_ENCODER.encode(this.finish())
    return new ReadableStream<Uint8Array>({
      pull(controller) {
        controller.enqueue(finish())
        controller.close()
      },
    })
  }
}

// ── True Streaming CSV Writer ────────────────────────────────────────

/** A streamed row: positional values, or an object read through headers. */
export type CsvStreamRow = CellValue[] | Record<string, CellValue>

/**
 * Write CSV as a byte stream, pulling rows from `rows` only as the
 * consumer reads.
 *
 * Peak memory is independent of the row count — lines are formatted,
 * encoded, and flushed as they accumulate, and nothing is retained.
 *
 * ```ts
 * return new Response(writeCsvStream(rowCursor, { headers: ["id", "name"] }), {
 *   headers: { "content-type": "text/csv; charset=utf-8" },
 * })
 * ```
 *
 * Object rows are projected through a column order resolved the same way
 * {@link writeCsvObjects} resolves it: `columns` if given, else an
 * explicit `headers` array, else the keys of the first row. A header line
 * is emitted unless `headers: false`.
 */
export function writeCsvStream(
  rows: AsyncIterable<CsvStreamRow> | Iterable<CsvStreamRow>,
  options?: CsvWriteOptions,
): ReadableStream<Uint8Array> {
  const chunks = csvStreamChunks(rows, options)

  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      try {
        const { done, value } = await chunks.next()
        if (done) {
          controller.close()
          return
        }
        controller.enqueue(value)
      } catch (err) {
        controller.error(err)
      }
    },
    async cancel(reason) {
      await chunks.return?.(reason)
    },
  })
}

/** Format rows into ~64 KB encoded chunks, pulling lazily. */
async function* csvStreamChunks(
  rows: AsyncIterable<CsvStreamRow> | Iterable<CsvStreamRow>,
  options?: CsvWriteOptions,
): AsyncGenerator<Uint8Array> {
  const formatter = new CsvRowFormatter(options)
  const lineSeparator = formatter.lineSeparator

  let pending = ""
  let wroteAnyLine = false

  const push = (line: string): Uint8Array | undefined => {
    // `finish()` joins with the separator rather than terminating each
    // line, so the separator goes *before* every line but the first.
    pending += wroteAnyLine ? lineSeparator + line : line
    wroteAnyLine = true
    if (pending.length < CHUNK_THRESHOLD) return undefined
    const chunk = TEXT_ENCODER.encode(pending)
    pending = ""
    return chunk
  }

  if (formatter.bom) pending += UTF8_BOM

  const iterator: AsyncIterator<CsvStreamRow> | Iterator<CsvStreamRow> =
    Symbol.asyncIterator in Object(rows)
      ? (rows as AsyncIterable<CsvStreamRow>)[Symbol.asyncIterator]()
      : (rows as Iterable<CsvStreamRow>)[Symbol.iterator]()

  // Column order for object rows, resolved on the first one seen.
  const explicitColumns = options?.columns
  const headerOption = options?.headers
  let columns: string[] | undefined =
    explicitColumns ?? (Array.isArray(headerOption) ? headerOption : undefined)
  let headerEmitted = false

  const emitHeader = (names: string[]): Uint8Array | undefined => {
    headerEmitted = true
    if (headerOption === false) return undefined
    return push(formatter.formatHeader(names))
  }

  // An explicit column order means the header line is known up front, so
  // it goes out before the first row is even pulled.
  if (columns) {
    const chunk = emitHeader(columns)
    if (chunk) yield chunk
  }

  try {
    for (;;) {
      const result = await iterator.next()
      if (result.done) break
      const row = result.value

      let values: CellValue[]
      if (Array.isArray(row)) {
        values = row
      } else {
        if (!columns) {
          columns = Object.keys(row)
          if (!headerEmitted) {
            const chunk = emitHeader(columns)
            if (chunk) yield chunk
          }
        }
        values = columns.map((key) => row[key] ?? null)
      }

      const chunk = push(formatter.formatRow(values))
      if (chunk) yield chunk
    }
  } finally {
    await iterator.return?.()
  }

  if (pending.length > 0) yield TEXT_ENCODER.encode(pending)
}

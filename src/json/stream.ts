// ── NDJSON Streaming ─────────────────────────────────────────────────
// CF Workers / Deno / Node 18+ compatible: uses WHATWG ReadableStream only.

import type { CellValue } from "../_types"
import { InvalidArgumentError, ParseError } from "../errors"
import { flattenValue, type FlattenOptions } from "./flatten"

const TEXT_ENCODER = new TextEncoder()
const TEXT_DECODER = new TextDecoder("utf-8")

/**
 * Constructor options for {@link NdjsonStreamWriter}.
 */
export interface NdjsonStreamWriterOptions {
  // `Date` values always serialize as ISO strings. There used to be an
  // `isoDates` option and it never did anything: JSON.stringify calls
  // Date.prototype.toJSON before consulting the replacer, so a replacer
  // testing `value instanceof Date` is never reached. Removed before v1
  // rather than frozen; see the note in json/writer.ts.
  /**
   * Column names for {@link NdjsonStreamWriter.addRow}. Positional rows
   * are zipped against this list to produce each object's keys.
   *
   * NDJSON has no positional row concept of its own, so `addRow` throws
   * without it rather than guessing a key order.
   */
  columns?: string[]
}

/**
 * Incremental NDJSON writer. Each call to {@link addObject} appends one
 * JSON object terminated by `\n`. Use {@link toStream} to expose the
 * output as a `ReadableStream<Uint8Array>` for piping to a `Response`
 * body, file, or another stream.
 *
 * ```ts
 * const w = new NdjsonStreamWriter()
 * const body = w.toStream()
 * for await (const row of source) w.addObject(row)
 * w.finish() // closes the stream once it has drained
 * return new Response(body, { headers: { 'content-type': 'application/x-ndjson' } })
 * ```
 */
export class NdjsonStreamWriter {
  private buffer: string[] = []
  private done = false
  private columns: string[] | undefined

  constructor(options?: NdjsonStreamWriterOptions) {
    this.columns = options?.columns
  }

  /**
   * Append one row from an object — one NDJSON line per call.
   */
  addObject(row: Record<string, CellValue>): void {
    if (this.done) {
      throw new Error("Cannot write to NdjsonStreamWriter after finish()/end()")
    }
    this.buffer.push(JSON.stringify(row) + "\n")
  }

  /**
   * Append one row of positional values, keyed by the `columns` passed to
   * the constructor.
   *
   * Throws when no `columns` were configured — the same contract as
   * `XlsxStreamWriter.addObject`, which needs `columns[].key` to map the
   * other direction.
   */
  addRow(values: CellValue[]): void {
    if (!this.columns) {
      throw new InvalidArgumentError(
        "addRow requires `columns` — NDJSON rows are objects, so positional values need key names. Pass `new NdjsonStreamWriter({ columns: [...] })` or use addObject().",
      )
    }
    const row: Record<string, CellValue> = {}
    for (let i = 0; i < this.columns.length; i++) {
      row[this.columns[i]!] = values[i] ?? null
    }
    this.addObject(row)
  }

  /**
   * Mark the writer finished and return the buffered output.
   *
   * Subsequent `addRow`/`addObject` calls throw, and a stream handed out
   * by {@link toStream} closes once it has drained. Note that a writer
   * already drained through `toStream()` has nothing left to return here
   * — pick one drain or the other.
   */
  finish(): string {
    this.done = true
    return this.toString()
  }

  /**
   * @deprecated Renamed to {@link addObject} so all three stream writers
   * share `addRow` / `addObject` / `finish`. Same behaviour; this alias
   * will be removed in a future major.
   */
  write(row: Record<string, CellValue>): void {
    this.addObject(row)
  }

  /**
   * Mark the writer finished. Subsequent writes throw.
   *
   * @deprecated Use {@link finish}, which does the same and returns the
   * buffered output. This alias will be removed in a future major.
   */
  end(): void {
    this.done = true
  }

  /** Drain the buffered output as a single string. */
  toString(): string {
    return this.buffer.join("")
  }

  /**
   * Expose the writer as a `ReadableStream<Uint8Array>`. The stream
   * remains open until {@link finish} is called.
   *
   * Unlike `CsvStreamWriter.toStream()` / `XlsxStreamWriter.toStream()`,
   * which encode an already-finished buffer, this one is a live drain:
   * rows written while the consumer reads are delivered as they arrive.
   *
   * Consumed rows are released as they are enqueued, so a writer that is
   * only ever drained through the stream holds no more than the rows
   * written since the last pull. Note that {@link toString} then has
   * nothing left to return — pick one drain or the other.
   */
  toStream(): ReadableStream<Uint8Array> {
    const buffer = this.buffer
    const isDone = () => this.done

    return new ReadableStream<Uint8Array>({
      pull: (controller) => {
        // Detach rather than walk a cursor: holding on to already-sent
        // rows would make the stream O(total) instead of O(pending).
        if (buffer.length > 0) {
          for (const row of buffer.splice(0, buffer.length)) {
            controller.enqueue(TEXT_ENCODER.encode(row))
          }
        }
        if (isDone()) {
          controller.close()
        }
      },
    })
  }
}

/**
 * Read an NDJSON stream and yield each parsed object. Errors on malformed
 * lines throw by default; pass `onError` to skip and continue.
 */
export interface NdjsonStreamReadOptions extends FlattenOptions {
  onError?: (line: string, lineNumber: number, error: Error) => void
  /** Apply flattening to each row before yielding. Default: false. */
  flattenRows?: boolean
}

export async function* streamNdjsonRows<
  T extends Record<string, CellValue> = Record<string, CellValue>,
>(
  stream: ReadableStream<Uint8Array>,
  options?: NdjsonStreamReadOptions,
): AsyncGenerator<T, void, undefined> {
  const reader = stream.getReader()
  let buffer = ""
  let lineNumber = 0

  const flatten = options?.flattenRows ?? false
  const flatOpts: FlattenOptions = {
    flatten: options?.flatten,
    arrayJoin: options?.arrayJoin,
    maxDepth: options?.maxDepth,
  }

  try {
    while (true) {
      const { value, done } = await reader.read()
      if (value) {
        buffer += TEXT_DECODER.decode(value, { stream: true })
      }
      let newlineIdx: number
      while ((newlineIdx = buffer.indexOf("\n")) !== -1) {
        const line = buffer.slice(0, newlineIdx).replace(/\r$/, "")
        buffer = buffer.slice(newlineIdx + 1)
        lineNumber++
        if (line.trim() === "") continue
        const parsed = tryParseLine(line, lineNumber, options?.onError)
        if (parsed === SKIP) continue
        if (flatten && parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
          yield flattenValue(parsed, flatOpts) as T
        } else {
          yield parsed as T
        }
      }
      if (done) {
        // Flush trailing partial line (no newline)
        buffer += TEXT_DECODER.decode()
        const trailing = buffer.trim()
        if (trailing !== "") {
          lineNumber++
          const parsed = tryParseLine(trailing, lineNumber, options?.onError)
          if (parsed !== SKIP) {
            if (flatten && parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
              yield flattenValue(parsed, flatOpts) as T
            } else {
              yield parsed as T
            }
          }
        }
        break
      }
    }
  } finally {
    reader.releaseLock()
  }
}

const SKIP = Symbol("skip")

function tryParseLine(
  line: string,
  lineNumber: number,
  onError?: (line: string, lineNumber: number, error: Error) => void,
): unknown | typeof SKIP {
  try {
    return JSON.parse(line)
  } catch (err) {
    if (onError) {
      onError(line, lineNumber, err as Error)
      return SKIP
    }
    throw new ParseError(
      `Invalid NDJSON on line ${lineNumber}: ${(err as Error).message}`,
      { line: lineNumber },
      { cause: err },
    )
  }
}

/**
 * @deprecated Renamed to {@link streamNdjsonRows} so every streaming
 * reader in the library reads `stream*Rows`. This alias will be removed
 * in a future major.
 */
export const readNdjsonStream = streamNdjsonRows

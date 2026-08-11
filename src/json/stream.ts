// ── NDJSON Streaming ─────────────────────────────────────────────────
// CF Workers / Deno / Node 18+ compatible: uses WHATWG ReadableStream only.

import type { CellValue, SpreadsheetStreamWriter } from "../_types"
import { InvalidArgumentError, ParseError } from "../errors"
import { flattenValue, reviveDates, type FlattenOptions } from "./flatten"
import { unflattenRow } from "./unflatten"

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
  /**
   * Rebuild dot-path keys into nested objects before writing each line.
   * Default: false — same opt-in as `writeNdjson`, for the same reason;
   * `JsonWriteOptions.unflatten` in json/writer.ts carries the argument.
   */
  unflatten?: boolean
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
export class NdjsonStreamWriter implements SpreadsheetStreamWriter {
  private buffer: string[] = []
  private done = false
  private columns: string[] | undefined
  private unflatten: boolean

  constructor(options?: NdjsonStreamWriterOptions) {
    this.columns = options?.columns
    this.unflatten = options?.unflatten ?? false
  }

  /**
   * Append one row from an object — one NDJSON line per call.
   */
  addObject(row: Record<string, CellValue>): void {
    if (this.done) {
      throw new Error("Cannot write to NdjsonStreamWriter after finish()/end()")
    }
    this.buffer.push(JSON.stringify(this.unflatten ? unflattenRow(row) : row) + "\n")
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
    typeInference: options?.typeInference,
  }

  // Rows yielded without `flattenRows` never reach `flattenValue`, which is
  // where inference normally happens — so they get the same rule applied to
  // the tree directly.
  const reviveWholeRow = (options?.typeInference ?? false) && !flatten

  const emit = (parsed: unknown): T => {
    if (flatten && parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
      return flattenValue(parsed, flatOpts) as T
    }
    return (reviveWholeRow ? reviveDates(parsed, options?.maxDepth ?? 32) : parsed) as T
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
        yield emit(parsed)
      }
      if (done) {
        // Flush trailing partial line (no newline)
        buffer += TEXT_DECODER.decode()
        const trailing = buffer.trim()
        if (trailing !== "") {
          lineNumber++
          const parsed = tryParseLine(trailing, lineNumber, options?.onError)
          if (parsed !== SKIP) yield emit(parsed)
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
export const readNdjsonStream: typeof streamNdjsonRows = streamNdjsonRows

// ── True Streaming NDJSON Writer ─────────────────────────────────────

/** A streamed row: an object, or positional values read through `columns`. */
export type NdjsonStreamRow = Record<string, CellValue> | CellValue[]

/**
 * Write NDJSON as a byte stream, pulling rows from `rows` only as the
 * consumer reads.
 *
 * `NdjsonStreamWriter.toStream()` already streamed live, but it is a
 * class the caller has to drive — and `writeXlsxStream` /
 * `writeCsvStream` are the shape that reads naturally at a `Response`
 * boundary. NDJSON, of all formats, should not be the one that lacks it.
 * See #467.
 *
 * ```ts
 * return new Response(writeNdjsonStream(rowCursor), {
 *   headers: { "content-type": "application/x-ndjson" },
 * })
 * ```
 *
 * Peak memory is independent of the row count: each row is serialized,
 * encoded and enqueued on its own, and nothing is retained.
 *
 * Positional rows need `columns`, for the same reason
 * {@link NdjsonStreamWriter.addRow} does — NDJSON rows are objects, so
 * values with no key names describe nothing.
 */
export function writeNdjsonStream(
  rows: AsyncIterable<NdjsonStreamRow> | Iterable<NdjsonStreamRow>,
  options?: NdjsonStreamWriterOptions,
): ReadableStream<Uint8Array> {
  const chunks = ndjsonStreamChunks(rows, options)

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

/** Serialize rows into ~64 KB encoded chunks, pulling lazily. */
async function* ndjsonStreamChunks(
  rows: AsyncIterable<NdjsonStreamRow> | Iterable<NdjsonStreamRow>,
  options?: NdjsonStreamWriterOptions,
): AsyncGenerator<Uint8Array> {
  const columns = options?.columns
  const unflatten = options?.unflatten ?? false

  // Batching is what keeps this from being one syscall-sized write per
  // row; the same 64 KB the CSV writer uses.
  const CHUNK_BYTES = 64 * 1024
  let pending: string[] = []
  let pendingBytes = 0

  for await (const row of rows) {
    let object: Record<string, CellValue>
    if (Array.isArray(row)) {
      if (!columns) {
        throw new InvalidArgumentError(
          "writeNdjsonStream needs `columns` for positional rows — NDJSON rows " +
            "are objects, so values with no key names describe nothing. Pass " +
            "`{ columns: [...] }` or yield objects.",
        )
      }
      object = {}
      for (let i = 0; i < columns.length; i++) object[columns[i]!] = row[i] ?? null
    } else {
      object = row
    }

    const line = `${JSON.stringify(unflatten ? unflattenRow(object) : object)}\n`
    pending.push(line)
    pendingBytes += line.length

    if (pendingBytes >= CHUNK_BYTES) {
      yield TEXT_ENCODER.encode(pending.join(""))
      pending = []
      pendingBytes = 0
    }
  }

  if (pending.length > 0) yield TEXT_ENCODER.encode(pending.join(""))
}

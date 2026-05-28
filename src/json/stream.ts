// ── NDJSON Streaming ─────────────────────────────────────────────────
// CF Workers / Deno / Node 18+ compatible: uses WHATWG ReadableStream only.

import type { CellValue } from "../_types"
import { ParseError } from "../errors"
import { flattenValue, type FlattenOptions } from "./flatten"

const TEXT_ENCODER = new TextEncoder()
const TEXT_DECODER = new TextDecoder("utf-8")

/**
 * Incremental NDJSON writer. Each call to {@link write} appends one JSON
 * object terminated by `\n`. Use {@link toStream} to expose the output as
 * a `ReadableStream<Uint8Array>` for piping to a `Response` body, file,
 * or another stream.
 *
 * ```ts
 * const w = new NdjsonStreamWriter()
 * for await (const row of source) w.write(row)
 * w.end()
 * return new Response(w.toStream(), { headers: { 'content-type': 'application/x-ndjson' } })
 * ```
 */
export class NdjsonStreamWriter {
  private buffer: string[] = []
  private done = false
  private isoDates: boolean

  constructor(options?: { isoDates?: boolean }) {
    this.isoDates = options?.isoDates ?? true
  }

  /** Append one row. */
  write(row: Record<string, CellValue>): void {
    if (this.done) {
      throw new Error("Cannot write to NdjsonStreamWriter after end()")
    }
    const replacer = this.isoDates ? dateReplacer : undefined
    this.buffer.push(JSON.stringify(row, replacer) + "\n")
  }

  /** Mark the writer finished. Subsequent writes throw. */
  end(): void {
    this.done = true
  }

  /** Drain the buffered output as a single string. */
  toString(): string {
    return this.buffer.join("")
  }

  /**
   * Expose the writer as a `ReadableStream<Uint8Array>`. The stream
   * remains open until {@link end} is called.
   */
  toStream(): ReadableStream<Uint8Array> {
    const buffer = this.buffer
    const isDone = () => this.done
    let cursor = 0

    return new ReadableStream<Uint8Array>({
      pull: (controller) => {
        while (cursor < buffer.length) {
          controller.enqueue(TEXT_ENCODER.encode(buffer[cursor]!))
          cursor++
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

export async function* readNdjsonStream<
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

function dateReplacer(_key: string, value: unknown): unknown {
  if (value instanceof Date) return value.toISOString()
  return value
}

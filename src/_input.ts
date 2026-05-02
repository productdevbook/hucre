// ── ReadInput Normalization ────────────────────────────────────────────
//
// Helpers for normalizing the {@link ReadInput} union type into the byte
// buffer that ZIP-based readers (XLSX, ODS, XLSB) need.
//
// `ReadableStream<Uint8Array>` input must be fully buffered because every
// supported container format stores its central directory or directory
// equivalent at the end of the file — true streaming is not possible
// without random access. Buffering happens once and is shared by every
// reader, so the unified `read()` API and `readXlsx`/`readOds` direct
// entry points all accept streams uniformly.
// ──────────────────────────────────────────────────────────────────────

import type { ReadInput } from "./_types";
import { ParseError } from "./errors";

/**
 * Drain a {@link ReadableStream} of byte chunks into a single
 * {@link Uint8Array}. Allocates only one extra buffer when the stream
 * yields more than one chunk.
 */
export async function bufferReadableStream(
  stream: ReadableStream<Uint8Array>,
): Promise<Uint8Array> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let totalLen = 0;

  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (value) {
      chunks.push(value);
      totalLen += value.length;
    }
  }

  if (chunks.length === 0) return new Uint8Array(0);
  if (chunks.length === 1) return chunks[0]!;

  const result = new Uint8Array(totalLen);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}

/**
 * Detect whether a value is a ReadableStream of bytes. Avoids relying on
 * `instanceof ReadableStream` so the check works across realms (Node
 * worker threads, browser iframes, undici, etc.) where multiple
 * `ReadableStream` constructors may exist.
 */
function isReadableStream(value: unknown): value is ReadableStream<Uint8Array> {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as ReadableStream<Uint8Array>).getReader === "function"
  );
}

/**
 * Normalize a {@link ReadInput} into a {@link Uint8Array}. Buffers any
 * `ReadableStream<Uint8Array>` input fully. Throws {@link ParseError}
 * for unsupported input shapes.
 */
export async function readInputToUint8Array(input: ReadInput): Promise<Uint8Array> {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (isReadableStream(input)) return bufferReadableStream(input);
  throw new ParseError(
    "Unsupported input type. Expected Uint8Array, ArrayBuffer, or ReadableStream<Uint8Array>.",
  );
}

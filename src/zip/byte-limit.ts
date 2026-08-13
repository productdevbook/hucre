// ── Decompressed-size ceiling for streaming ZIP paths ────────────────
//
// The buffered decompressors count bytes as they go and stop at
// `MAX_DECOMPRESSED_BYTES` (or the entry's declared uncompressed size,
// whichever is smaller). The streaming paths hand a native
// `DecompressionStream` straight to the consumer, which counts nothing —
// so the exact same zip bomb that the buffered reader rejects expanded
// without limit as soon as the caller used a stream. This transform puts
// the two paths back on the same footing.
//
// The check has to be on the running total, not on each chunk: a bomb
// arrives as an endless sequence of perfectly ordinary 64 KiB chunks.

import { ZipError } from "../errors"

/**
 * A pass-through {@link TransformStream} that errors the stream once more
 * than `maxBytes` decompressed bytes have flowed through it. Erroring
 * (rather than truncating) means the consumer sees a typed
 * {@link ZipError} instead of a silently short read.
 */
export function byteLimitStream(maxBytes: number): TransformStream<Uint8Array, Uint8Array> {
  let total = 0
  return new TransformStream<Uint8Array, Uint8Array>({
    transform(chunk, controller) {
      total += chunk.length
      if (total > maxBytes) {
        throw new ZipError(
          `Decompressed size exceeds limit of ${maxBytes} bytes (possible zip bomb)`,
        )
      }
      controller.enqueue(chunk)
    },
  })
}

/**
 * How much of an already-in-memory entry one chunk of its stream holds.
 *
 * `DecompressionStream` hands back pieces of its own accord, but the two
 * paths that do not decompress — a STORE entry, and the fallback inflate
 * that returns a whole buffer — had nothing to divide and enqueued the
 * entry in one go. A consumer that accumulates what it is given then
 * gets the entire entry as one chunk: `parseSaxStream` does
 * `buf += decoder.decode(chunk)`, so a 589 MB stored worksheet rebuilt
 * the 512 MB string the streaming path exists to avoid, and the reader
 * for the files in #503 would have failed on exactly the ones it was
 * added for. Nothing here is copied — each chunk is a `subarray` view.
 *
 * 1 MiB is small enough to keep any one string far from the ceiling and
 * large enough that the per-chunk overhead does not show up in a
 * measurement.
 */
export const STREAM_CHUNK_BYTES: number = 1024 * 1024

/** Emit an in-memory buffer as a stream of {@link STREAM_CHUNK_BYTES} pieces. */
export function chunkedStream(data: Uint8Array): ReadableStream<Uint8Array> {
  let offset = 0
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (offset >= data.length) {
        controller.close()
        return
      }
      const end = Math.min(offset + STREAM_CHUNK_BYTES, data.length)
      controller.enqueue(data.subarray(offset, end))
      offset = end
    },
  })
}

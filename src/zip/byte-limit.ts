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

// ── Streaming inflate for a ZIP entry body ───────────────────────────
// Turns a stream of an entry's *compressed* bytes into a stream of its
// decompressed bytes, under the same ceiling the buffered readers apply.
//
// The local-header streaming reader and the random-access reader both
// pull compressed bytes from somewhere that is not memory, and both need
// exactly this: the native `DecompressionStream` where there is one, a
// buffered pure-TS inflate where there is not, and `byteLimitStream` on
// either branch so a bomb cannot stream out unbounded.

import { byteLimitStream, STREAM_CHUNK_BYTES } from "./byte-limit"
import { canInflateRaw } from "./capability"
import { inflate } from "./deflate"
import { HucreError, ZipError } from "../errors"

/**
 * Inflate `raw` (a raw-DEFLATE body) as a stream, failing with a
 * {@link ZipError} once more than `cap` bytes have come out.
 *
 * `name` is the entry path, used only in error messages.
 */
export function inflateRawStream(
  raw: ReadableStream<Uint8Array>,
  cap: number,
  name: string,
): ReadableStream<Uint8Array> {
  if (canInflateRaw()) {
    const inflated = (
      raw.pipeThrough(
        new DecompressionStream("deflate-raw") as unknown as ReadableWritablePair<
          Uint8Array,
          Uint8Array
        >,
      ) as ReadableStream<Uint8Array>
    ).pipeThrough(byteLimitStream(cap))
    return typedErrors(inflated, name)
  }
  // No DecompressionStream — buffer the compressed body and inflate once.
  // The result is handed on in pieces rather than as one enqueue: a
  // consumer that accumulates what it is given (`parseSaxStream` does
  // `buf += decoder.decode(chunk)`) would otherwise rebuild the very
  // string the streaming path exists to avoid, and this is the branch
  // the browser/Workers claims rest on. Same reason as the twin
  // paths in `reader.ts`. See #503.
  return new ReadableStream<Uint8Array>({
    async start(controller) {
      const chunks: Uint8Array[] = []
      const r = raw.getReader()
      for (;;) {
        const { value, done } = await r.read()
        if (done) break
        if (value) chunks.push(value)
      }
      const inflated = inflate(concat(chunks), cap)
      for (let at = 0; at < inflated.length; at += STREAM_CHUNK_BYTES) {
        controller.enqueue(
          inflated.subarray(at, Math.min(at + STREAM_CHUNK_BYTES, inflated.length)),
        )
      }
      controller.close()
    },
  })
}

/**
 * Re-raise what the native decompressor throws as a {@link ZipError}.
 *
 * `DecompressionStream` reports a corrupt or truncated body as a bare
 * `TypeError` — "unexpected end of file", or on Node no message at all —
 * which escapes the typed-error contract every other ZIP failure keeps.
 * It matters more once an entry's end is found by scanning for its data
 * descriptor: a wrong boundary surfaces exactly here, and the caller
 * should be told which entry and why. Errors that are already ours (the
 * size cap) pass through unchanged.
 */
function typedErrors(source: ReadableStream<Uint8Array>, name: string): ReadableStream<Uint8Array> {
  const reader = source.getReader()
  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      try {
        const { value, done } = await reader.read()
        if (done) controller.close()
        else controller.enqueue(value)
      } catch (err) {
        controller.error(
          err instanceof HucreError
            ? err
            : new ZipError(`Invalid DEFLATE data in ${name}`, { cause: err }),
        )
      }
    },
    cancel(reason) {
      return reader.cancel(reason)
    },
  })
}

export function concat(chunks: Uint8Array[]): Uint8Array {
  if (chunks.length === 1) return chunks[0]!
  let total = 0
  for (const c of chunks) total += c.length
  const out = new Uint8Array(total)
  let off = 0
  for (const c of chunks) {
    out.set(c, off)
    off += c.length
  }
  return out
}

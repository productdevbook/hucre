// ── Streaming ZIP Reader ─────────────────────────────────────────────
// Parses a ZIP archive front-to-back from a ReadableStream of bytes,
// reading each entry's *local file header* instead of the central
// directory (which lives at the end of the file and would force the whole
// archive into memory).
//
// This enables true streaming for the common case: an XLSX reader can pull
// the small metadata parts (content types, rels, workbook, shared strings,
// styles) and then pipe a single target worksheet straight into a SAX
// parser without ever holding the full archive — or the full worksheet —
// in memory.
//
// Local-header streaming only works when the entry's compressed size is
// known up front (general-purpose flag bit 3 clear) and the data isn't
// Zip64-escaped. When an entry violates those assumptions the caller can
// recover the bytes consumed so far via {@link ZipStreamReader.drainToBuffer}
// and fall back to the random-access {@link ZipReader}.

import { inflate } from "./deflate"
import { ZipError } from "../errors"
import { MAX_DECOMPRESSED_BYTES } from "../limits"

const SIG_LOCAL_FILE = 0x04034b50
const SIG_CENTRAL_DIR = 0x02014b50
const ZIP64_SENTINEL = 0xffffffff
const FLAG_DATA_DESCRIPTOR = 0x0008

const EMPTY = new Uint8Array(0)

let hasDecompressionStream: boolean | undefined
function checkDecompressionStream(): boolean {
  if (hasDecompressionStream === undefined) {
    try {
      hasDecompressionStream =
        typeof DecompressionStream !== "undefined" &&
        typeof ReadableStream !== "undefined" &&
        typeof Response !== "undefined"
    } catch {
      hasDecompressionStream = false
    }
  }
  return hasDecompressionStream
}

/** One local-file-header record surfaced by {@link ZipStreamReader.nextEntry}. */
export interface ZipStreamEntry {
  /** Entry path (e.g. `xl/worksheets/sheet1.xml`). */
  name: string
  /** Compression method — 0 (stored) or 8 (DEFLATE). */
  compressionMethod: number
  /** Compressed byte length from the local header (0 / unreliable if {@link streamable} is false). */
  compressedSize: number
  /** Uncompressed byte length from the local header. */
  uncompressedSize: number
  /** General-purpose bit flags. */
  flags: number
  /**
   * Whether this entry can be streamed by local header alone: compressed
   * size is known (no trailing data descriptor), not Zip64-escaped, and a
   * supported compression method. When false the caller should fall back.
   */
  streamable: boolean
}

/**
 * Pull-based ZIP reader over a byte {@link ReadableStream}.
 *
 * Usage: call {@link nextEntry}; for each non-null entry, consume its body
 * with exactly one of {@link readEntryBytes} / {@link entryStream} /
 * {@link skipEntry} before calling {@link nextEntry} again. {@link nextEntry}
 * returns `null` once the central directory is reached (end of entries).
 */
export class ZipStreamReader {
  private reader: ReadableStreamDefaultReader<Uint8Array>
  private leftover: Uint8Array = EMPTY
  private ended = false
  /** Raw chunks pulled from the source, retained for {@link drainToBuffer}. */
  private fallbackLog: Uint8Array[] | null = []
  /** Bytes of the current entry's body still to be consumed. */
  private pendingBody = 0
  private bodyConsumed = true

  constructor(stream: ReadableStream<Uint8Array>) {
    this.reader = stream.getReader()
  }

  /** Pull one more chunk from the source, logging it for fallback. */
  private async pull(): Promise<Uint8Array | null> {
    const { value, done } = await this.reader.read()
    if (done || !value) {
      this.ended = true
      return null
    }
    if (this.fallbackLog) this.fallbackLog.push(value)
    return value
  }

  /** Read exactly `n` bytes, or return null if the source ends first. */
  private async readExact(n: number): Promise<Uint8Array | null> {
    if (n === 0) return EMPTY
    while (this.leftover.length < n) {
      const chunk = await this.pull()
      if (!chunk) return null
      if (this.leftover.length === 0) {
        this.leftover = chunk
      } else {
        const merged = new Uint8Array(this.leftover.length + chunk.length)
        merged.set(this.leftover, 0)
        merged.set(chunk, this.leftover.length)
        this.leftover = merged
      }
    }
    const out = this.leftover.subarray(0, n)
    this.leftover = this.leftover.subarray(n)
    return out
  }

  /**
   * Read the next local file header. Returns null at the central directory
   * (no more entries). Throws if the body of the previous entry was not
   * consumed.
   */
  async nextEntry(): Promise<ZipStreamEntry | null> {
    if (!this.bodyConsumed) {
      throw new ZipError("ZipStreamReader: previous entry body not consumed")
    }
    const sigBytes = await this.readExact(4)
    if (!sigBytes) return null
    const sig = readU32(sigBytes, 0)
    if (sig === SIG_CENTRAL_DIR || sig !== SIG_LOCAL_FILE) {
      // Central directory (or anything that isn't another local header) —
      // we're past the entry list.
      return null
    }
    const rest = await this.readExact(26)
    if (!rest) throw new ZipError("ZipStreamReader: truncated local file header")
    const flags = readU16(rest, 2)
    const compressionMethod = readU16(rest, 4)
    const compressedSize = readU32(rest, 14)
    const uncompressedSize = readU32(rest, 18)
    const nameLen = readU16(rest, 22)
    const extraLen = readU16(rest, 24)
    const nameBytes = await this.readExact(nameLen)
    if (!nameBytes) throw new ZipError("ZipStreamReader: truncated entry name")
    const name = new TextDecoder("utf-8").decode(nameBytes)
    if (extraLen > 0) {
      const extra = await this.readExact(extraLen)
      if (!extra) throw new ZipError("ZipStreamReader: truncated extra field")
    }

    const streamable =
      (flags & FLAG_DATA_DESCRIPTOR) === 0 &&
      compressedSize !== ZIP64_SENTINEL &&
      uncompressedSize !== ZIP64_SENTINEL &&
      (compressionMethod === 0 || compressionMethod === 8)

    this.pendingBody = compressedSize
    this.bodyConsumed = false

    return { name, compressionMethod, compressedSize, uncompressedSize, flags, streamable }
  }

  /** Skip the current entry's body without decompressing it. */
  async skipEntry(): Promise<void> {
    if (this.bodyConsumed) return
    let remaining = this.pendingBody
    while (remaining > 0) {
      const take = await this.readExact(Math.min(remaining, 1 << 20))
      if (!take) throw new ZipError("ZipStreamReader: truncated entry body")
      remaining -= take.length
    }
    this.bodyConsumed = true
  }

  /** Read and (if DEFLATE) inflate the current entry's full body into memory. */
  async readEntryBytes(entry: ZipStreamEntry): Promise<Uint8Array> {
    if (this.bodyConsumed) throw new ZipError("ZipStreamReader: entry body already consumed")
    const compressed = (await this.readExact(this.pendingBody)) ?? EMPTY
    if (compressed.length !== this.pendingBody) {
      throw new ZipError("ZipStreamReader: truncated entry body")
    }
    this.bodyConsumed = true
    // Copy out of the leftover-backed view so later reads can't alias it.
    const owned = compressed.slice()
    if (entry.compressionMethod === 0) return owned
    const cap =
      entry.uncompressedSize > 0
        ? Math.min(entry.uncompressedSize, MAX_DECOMPRESSED_BYTES)
        : MAX_DECOMPRESSED_BYTES
    return inflate(owned, cap)
  }

  /**
   * Return a {@link ReadableStream} of the current entry's *decompressed*
   * bytes, pulled lazily from the source. Calling this commits to streaming
   * (the fallback log is released), so it must only be used once all
   * fallback conditions have been cleared.
   */
  entryStream(entry: ZipStreamEntry): ReadableStream<Uint8Array> {
    if (this.bodyConsumed) throw new ZipError("ZipStreamReader: entry body already consumed")
    this.fallbackLog = null // committed to streaming — stop retaining bytes
    let remaining = this.pendingBody
    const raw = new ReadableStream<Uint8Array>({
      pull: async (controller) => {
        if (remaining <= 0) {
          this.bodyConsumed = true
          controller.close()
          return
        }
        const take = await this.readExact(Math.min(remaining, 1 << 20))
        if (!take) {
          this.bodyConsumed = true
          controller.error(new ZipError("ZipStreamReader: truncated entry body"))
          return
        }
        remaining -= take.length
        controller.enqueue(take)
        if (remaining <= 0) {
          this.bodyConsumed = true
          controller.close()
        }
      },
    })
    if (entry.compressionMethod === 0) return raw
    if (checkDecompressionStream()) {
      return raw.pipeThrough(
        new DecompressionStream("deflate-raw") as unknown as ReadableWritablePair<
          Uint8Array,
          Uint8Array
        >,
      ) as ReadableStream<Uint8Array>
    }
    // No DecompressionStream — buffer the compressed body and inflate once.
    return new ReadableStream<Uint8Array>({
      async start(controller) {
        const chunks: Uint8Array[] = []
        const r = raw.getReader()
        for (;;) {
          const { value, done } = await r.read()
          if (done) break
          if (value) chunks.push(value)
        }
        const cap =
          entry.uncompressedSize > 0
            ? Math.min(entry.uncompressedSize, MAX_DECOMPRESSED_BYTES)
            : MAX_DECOMPRESSED_BYTES
        controller.enqueue(inflate(concat(chunks), cap))
        controller.close()
      },
    })
  }

  /**
   * Reconstruct the full archive bytes (everything pulled so far plus the
   * rest of the source) so the caller can fall back to a random-access
   * reader. Only valid while the fallback log is intact (before
   * {@link entryStream} has been called).
   */
  async drainToBuffer(): Promise<Uint8Array> {
    if (!this.fallbackLog) {
      throw new ZipError("ZipStreamReader: cannot fall back after streaming started")
    }
    const chunks = this.fallbackLog
    this.fallbackLog = null
    // Pull whatever remains from the source (stop logging — we own `chunks`).
    for (;;) {
      const { value, done } = await this.reader.read()
      if (done) break
      if (value) chunks.push(value)
    }
    return concat(chunks)
  }

  /** Release the underlying reader lock. */
  async close(): Promise<void> {
    try {
      await this.reader.cancel()
    } catch {
      // ignore
    }
  }
}

function readU16(b: Uint8Array, off: number): number {
  return b[off] | (b[off + 1] << 8)
}

function readU32(b: Uint8Array, off: number): number {
  return (b[off] | (b[off + 1] << 8) | (b[off + 2] << 16) | (b[off + 3] << 24)) >>> 0
}

function concat(chunks: Uint8Array[]): Uint8Array {
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

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
// An entry's extent comes from one of two places. Normally the local
// header states the compressed size (ZIP64 entries state it in the local
// extra field). With general-purpose bit 3 set it does not, and the sizes
// follow the body in a data descriptor; such a DEFLATE entry's end is
// found by scanning for that descriptor (see `bodyUntilDescriptor`). A
// STORE entry with bit 3, or one whose descriptor cannot be found, makes
// the entry unstreamable: the caller recovers the bytes consumed so far
// via {@link ZipStreamReader.drainToBuffer} and falls back to the
// random-access {@link ZipReader}.

import { crc32, inflate } from "./deflate"
import { concat, inflateRawStream } from "./inflate-stream"
import { decompressionCap, inputTooLarge, positiveOr } from "./bounds"
import { readZip64Extra, SIG_DATA_DESCRIPTOR, ZIP64_SENTINEL } from "./central-directory"
import { ZipError } from "../errors"
import { MAX_DECOMPRESSED_BYTES, MAX_INPUT_BYTES } from "../limits"

const SIG_LOCAL_FILE = 0x04034b50
const SIG_CENTRAL_DIR = 0x02014b50
const FLAG_DATA_DESCRIPTOR = 0x0008

const EMPTY = new Uint8Array(0)

/**
 * Bytes the descriptor scan holds back from the body it emits: the
 * longest descriptor (24, signature plus ZIP64 sizes) and the three bytes
 * of the next header's signature that might not have been checked yet.
 */
const DESCRIPTOR_HOLDBACK = 27

/** One local-file-header record surfaced by {@link ZipStreamReader.nextEntry}. */
export interface ZipStreamEntry {
  /** Entry path (e.g. `xl/worksheets/sheet1.xml`). */
  name: string
  /** Compression method — 0 (stored) or 8 (DEFLATE). */
  compressionMethod: number
  /**
   * Compressed byte length from the local header, or its ZIP64 extra
   * field. Unknown (0) when {@link dataDescriptor} is set.
   */
  compressedSize: number
  /** Uncompressed byte length, on the same terms as {@link compressedSize}. */
  uncompressedSize: number
  /** General-purpose bit flags. */
  flags: number
  /**
   * The sizes are not in the header: they follow the body in a data
   * descriptor, and the body's end is found by scanning for it.
   */
  dataDescriptor: boolean
  /**
   * Whether this entry can be streamed by local header alone: a supported
   * compression method, and an extent that is either stated in the header
   * or (for DEFLATE) recoverable from a trailing data descriptor. When
   * false the caller should fall back.
   */
  streamable: boolean
}

/** What a data descriptor says about the entry it closes. */
interface DataDescriptor {
  crc: number
  compressedSize: number
  uncompressedSize: number
  /** The 64-bit form, whose uncompressed size is not taken modulo 2^32. */
  zip64: boolean
}

/**
 * A data-descriptor entry this walk could not follow: the scan reached the
 * end of the source without finding the body's end, or, for a part read
 * whole, the boundary it found does not inflate or does not match the
 * descriptor's CRC and size. Not damage the
 * caller should see — `prepareStreaming` falls back on it while the
 * fallback log is intact, and the buffered reader, which has the central
 * directory, is the one to judge the archive.
 */
export class DataDescriptorUnresolved extends ZipError {}

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
  /** Bytes of the current entry's body still to be consumed (sized entries only). */
  private pendingBody = 0
  /** The current entry's extent is found by scanning for its descriptor. */
  private pendingDescriptor = false
  private bodyConsumed = true
  /** Bytes held in {@link fallbackLog}. */
  private logged = 0
  private readonly maxDecompressedBytes: number
  private readonly maxInputBytes: number

  constructor(
    stream: ReadableStream<Uint8Array>,
    maxDecompressedBytes: number = MAX_DECOMPRESSED_BYTES,
    maxInputBytes: number = MAX_INPUT_BYTES,
  ) {
    this.reader = stream.getReader()
    this.maxDecompressedBytes = maxDecompressedBytes
    this.maxInputBytes = positiveOr(maxInputBytes, MAX_INPUT_BYTES)
  }

  /**
   * Pull one more chunk from the source, logging it for fallback.
   *
   * The log is the buffer a fallback hands to the random-access reader, so
   * it is held to `maxInputBytes` as it grows rather than only when the
   * fallback is taken. The difference matters since data descriptors are
   * followed: the scan for one can read a stream to its end before the
   * walk decides it cannot stream, and a check made only in
   * `drainToBuffer` would then find nothing left to count.
   */
  private async pull(): Promise<Uint8Array | null> {
    const { value, done } = await this.reader.read()
    if (done || !value) {
      this.ended = true
      return null
    }
    if (this.fallbackLog) {
      this.logged += value.length
      if (this.logged > this.maxInputBytes) throw inputTooLarge(this.maxInputBytes)
      this.fallbackLog.push(value)
    }
    return value
  }

  /** Append one more chunk to `leftover`. False once the source has ended. */
  private async pullIntoLeftover(): Promise<boolean> {
    const chunk = await this.pull()
    if (!chunk) return false
    if (this.leftover.length === 0) {
      this.leftover = chunk
    } else {
      const merged = new Uint8Array(this.leftover.length + chunk.length)
      merged.set(this.leftover, 0)
      merged.set(chunk, this.leftover.length)
      this.leftover = merged
    }
    return true
  }

  /** Read exactly `n` bytes, or return null if the source ends first. */
  private async readExact(n: number): Promise<Uint8Array | null> {
    if (n === 0) return EMPTY
    while (this.leftover.length < n) {
      if (!(await this.pullIntoLeftover())) return null
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
    let compressedSize = readU32(rest, 14)
    let uncompressedSize = readU32(rest, 18)
    const nameLen = readU16(rest, 22)
    const extraLen = readU16(rest, 24)
    const nameBytes = await this.readExact(nameLen)
    if (!nameBytes) throw new ZipError("ZipStreamReader: truncated entry name")
    const name = new TextDecoder("utf-8").decode(nameBytes)
    let extra: Uint8Array = EMPTY
    if (extraLen > 0) {
      const got = await this.readExact(extraLen)
      if (!got) throw new ZipError("ZipStreamReader: truncated extra field")
      extra = got
    }

    const supported = compressionMethod === 0 || compressionMethod === 8
    const dataDescriptor = (flags & FLAG_DATA_DESCRIPTOR) !== 0
    let sized = !dataDescriptor
    // A ZIP64 entry written with its sizes known leaves the 32-bit fields
    // at the sentinel and states them in the local ZIP64 extra field.
    if (sized && (compressedSize === ZIP64_SENTINEL || uncompressedSize === ZIP64_SENTINEL)) {
      try {
        const z = readZip64Extra(
          extra,
          0,
          extra.length,
          uncompressedSize === ZIP64_SENTINEL,
          compressedSize === ZIP64_SENTINEL,
          false,
          name,
        )
        if (z.uncompressedSize !== undefined) uncompressedSize = z.uncompressedSize
        if (z.compressedSize !== undefined) compressedSize = z.compressedSize
      } catch {
        sized = false
      }
    }

    // DEFLATE ends itself, which is what lets a descriptor entry be
    // streamed: the scan's boundary is checked by the inflater finishing
    // exactly there (and, for parts read whole, by the CRC). A STORE body
    // has no end of its own, so a wrong boundary would go unnoticed; that
    // case stays on the fallback.
    const streamable = supported && (sized || (dataDescriptor && compressionMethod === 8))

    if (dataDescriptor) {
      compressedSize = 0
      uncompressedSize = 0
    }
    this.pendingBody = sized ? compressedSize : 0
    this.pendingDescriptor = dataDescriptor && streamable
    this.bodyConsumed = false

    return {
      name,
      compressionMethod,
      compressedSize,
      uncompressedSize,
      flags,
      dataDescriptor,
      streamable,
    }
  }

  /** Skip the current entry's body without decompressing it. */
  async skipEntry(): Promise<void> {
    if (this.bodyConsumed) return
    if (this.pendingDescriptor) {
      for await (const _ of this.bodyUntilDescriptor()) {
        // discarded
      }
      return
    }
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
    if (this.pendingDescriptor) return this.readDescriptorEntry(entry)
    const compressed = (await this.readExact(this.pendingBody)) ?? EMPTY
    if (compressed.length !== this.pendingBody) {
      throw new ZipError("ZipStreamReader: truncated entry body")
    }
    this.bodyConsumed = true
    // Copy out of the leftover-backed view so later reads can't alias it.
    const owned = compressed.slice()
    if (entry.compressionMethod === 0) return owned
    const cap = decompressionCap(entry.uncompressedSize, this.maxDecompressedBytes)
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
    const raw = this.pendingDescriptor ? this.descriptorBodyStream() : this.sizedBodyStream()
    if (entry.compressionMethod === 0) return raw
    // Same ceiling on both branches: the declared uncompressed size when the
    // header states one, the absolute cap otherwise. The native path used to
    // enforce neither, so a bomb streamed out unbounded.
    const cap = decompressionCap(entry.uncompressedSize, this.maxDecompressedBytes)
    return inflateRawStream(raw, cap, entry.name)
  }

  private sizedBodyStream(): ReadableStream<Uint8Array> {
    let remaining = this.pendingBody
    return new ReadableStream<Uint8Array>({
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
  }

  private descriptorBodyStream(): ReadableStream<Uint8Array> {
    const body = this.bodyUntilDescriptor()
    return new ReadableStream<Uint8Array>({
      async pull(controller) {
        try {
          const { value, done } = await body.next()
          if (done) controller.close()
          else controller.enqueue(value)
        } catch (err) {
          controller.error(err)
        }
      },
      async cancel() {
        await body.return(undefined)
      },
    })
  }

  /**
   * Read a descriptor entry whole and check it against its descriptor.
   *
   * The scan found the boundary by a size match; the CRC and uncompressed
   * size are what confirm it, and a part read whole is the one place both
   * can be checked before anything is built on the bytes.
   */
  private async readDescriptorEntry(entry: ZipStreamEntry): Promise<Uint8Array> {
    const chunks: Uint8Array[] = []
    let total = 0
    let descriptor: DataDescriptor | undefined
    const body = this.bodyUntilDescriptor((d) => (descriptor = d))
    for await (const chunk of body) {
      total += chunk.length
      // A part held whole is bounded like a buffered stream input; the
      // entry's own size is not known until its end.
      if (total > this.maxInputBytes) throw inputTooLarge(this.maxInputBytes)
      // Copy: each chunk is a view of `leftover`, which the next pull
      // replaces or merges into, so a held view could alias later bytes.
      chunks.push(chunk.slice())
    }
    const compressed = concat(chunks)
    const d = descriptor!
    const cap = decompressionCap(d.uncompressedSize, this.maxDecompressedBytes)
    // A wrong boundary usually fails here rather than at the check below:
    // the body is cut short ("Unexpected end of DEFLATE data"), or the cap
    // taken from a forged descriptor is too small. Either is the scan being
    // wrong, not a verdict on the archive, so it is Unresolved and the
    // caller falls back to the reader that has the central directory. A
    // genuine bomb loses nothing by it: that reader applies the same cap.
    let out: Uint8Array
    try {
      out = compressed.length === 0 ? EMPTY : inflate(compressed, cap)
    } catch (err) {
      throw new DataDescriptorUnresolved(
        `ZipStreamReader: ${entry.name} does not inflate up to its data descriptor`,
        { cause: err },
      )
    }
    const size = d.zip64 ? out.length : out.length >>> 0
    if (size !== d.uncompressedSize || (d.crc !== 0 && crc32(out) !== d.crc)) {
      throw new DataDescriptorUnresolved(
        `ZipStreamReader: ${entry.name} does not match its data descriptor`,
      )
    }
    return out
  }

  /**
   * Yield the current entry's compressed body up to its data descriptor,
   * then consume the descriptor.
   *
   * The body's length is not stated anywhere before it ends, so the end
   * is found from the far side: the descriptor is followed by the next
   * local header or the central directory, and it records the body's
   * compressed size. So at every `PK\x03\x04` / `PK\x01\x02` in the bytes,
   * each of the four descriptor forms is tried as ending just there —
   * with or without the optional `PK\x07\x08` signature, with 32- or
   * 64-bit sizes — and the first whose compressed size equals the bytes
   * since the body began is the boundary. A false match needs a 32-bit
   * field to equal the exact offset it sits at, as well as a header
   * signature after it.
   *
   * Feeding the inflater the bytes after the body does not work instead:
   * `DecompressionStream` rejects anything past the end of the DEFLATE
   * stream ("Trailing junk"), and does not say where that end was.
   *
   * At most {@link DESCRIPTOR_HOLDBACK} bytes are held back, so memory is
   * one chunk whatever the entry's size.
   */
  private async *bodyUntilDescriptor(
    onDescriptor?: (d: DataDescriptor) => void,
  ): AsyncGenerator<Uint8Array, void, undefined> {
    let emitted = 0
    let scanFrom = 0
    for (;;) {
      const buf = this.leftover
      for (let j = buf.indexOf(0x50, scanFrom); j !== -1 && j + 4 <= buf.length;) {
        if (isEntryBoundary(buf, j)) {
          const found = descriptorEndingAt(buf, j, emitted)
          if (found) {
            if (found.bodyEnd > 0) yield buf.subarray(0, found.bodyEnd)
            this.leftover = buf.subarray(j)
            this.bodyConsumed = true
            onDescriptor?.(found.descriptor)
            return
          }
        }
        j = buf.indexOf(0x50, j + 1)
      }
      // No boundary yet. Everything but the tail is body, whatever comes
      // next: a descriptor ending at an unchecked position starts no
      // earlier than DESCRIPTOR_HOLDBACK bytes from the end.
      const safe = buf.length - DESCRIPTOR_HOLDBACK
      if (safe > 0) {
        const out = buf.subarray(0, safe)
        this.leftover = buf.subarray(safe)
        emitted += safe
        scanFrom = Math.max(0, buf.length - 3 - safe)
        yield out
      } else {
        scanFrom = Math.max(0, buf.length - 3)
      }
      if (!(await this.pullIntoLeftover())) {
        this.bodyConsumed = true
        throw new DataDescriptorUnresolved(
          "ZipStreamReader: reached the end of the stream without finding the data descriptor",
        )
      }
    }
  }

  /**
   * Reconstruct the full archive bytes (everything pulled so far plus the
   * rest of the source) so the caller can fall back to a random-access
   * reader. Only valid while the fallback log is intact (before
   * {@link entryStream} has been called).
   *
   * Buffering is bounded by `maxBytes` (default {@link MAX_INPUT_BYTES}):
   * this is the one place the streaming reader gives up on streaming and
   * holds the whole archive, so it needs the same ceiling as the buffered
   * entry points.
   */
  async drainToBuffer(maxBytes: number = MAX_INPUT_BYTES): Promise<Uint8Array> {
    if (!this.fallbackLog) {
      throw new ZipError("ZipStreamReader: cannot fall back after streaming started")
    }
    const cap = positiveOr(maxBytes, MAX_INPUT_BYTES)
    const chunks = this.fallbackLog
    this.fallbackLog = null
    let total = this.logged
    if (total > cap) throw inputTooLarge(cap)
    // Pull whatever remains from the source (stop logging — we own `chunks`).
    for (;;) {
      if (this.ended) break
      const { value, done } = await this.reader.read()
      if (done) break
      if (value) {
        total += value.length
        if (total > cap) throw inputTooLarge(cap)
        chunks.push(value)
      }
    }
    return concat(chunks)
  }

  /**
   * Cancel the source and release the underlying reader lock. Safe to call
   * more than once, and on an already-exhausted stream. Callers must invoke
   * it when they abandon the reader (an error mid-parse, most commonly),
   * otherwise the source stays locked and whatever backs it — a socket, a
   * file handle — is never released.
   */
  async close(): Promise<void> {
    try {
      await this.reader.cancel()
    } catch {
      // ignore
    }
    try {
      this.reader.releaseLock()
    } catch {
      // ignore
    }
  }
}

/** `PK\x03\x04` (next entry) or `PK\x01\x02` (central directory) at `j`. */
function isEntryBoundary(b: Uint8Array, j: number): boolean {
  return (
    b[j + 1] === 0x4b &&
    ((b[j + 2] === 0x03 && b[j + 3] === 0x04) || (b[j + 2] === 0x01 && b[j + 3] === 0x02))
  )
}

/**
 * Try each descriptor form as ending at `j`. `emitted` is how much of the
 * body precedes `b`, so a descriptor starting at `i` closes a body of
 * `emitted + i` bytes — which its compressed-size field must equal.
 */
function descriptorEndingAt(
  b: Uint8Array,
  j: number,
  emitted: number,
): { bodyEnd: number; descriptor: DataDescriptor } | null {
  // [length, has signature, 64-bit sizes]
  const forms: Array<[number, boolean, boolean]> = [
    [16, true, false],
    [24, true, true],
    [12, false, false],
    [20, false, true],
  ]
  for (const [length, signed, zip64] of forms) {
    const i = j - length
    if (i < 0) continue
    if (signed && readU32(b, i) !== SIG_DATA_DESCRIPTOR) continue
    const at = signed ? i + 4 : i
    const compressedSize = zip64 ? readU64(b, at + 4) : readU32(b, at + 4)
    if (compressedSize !== emitted + i) continue
    return {
      bodyEnd: i,
      descriptor: {
        crc: readU32(b, at),
        compressedSize,
        uncompressedSize: zip64 ? readU64(b, at + 12) : readU32(b, at + 8),
        zip64,
      },
    }
  }
  return null
}

function readU16(b: Uint8Array, off: number): number {
  return b[off] | (b[off + 1] << 8)
}

function readU32(b: Uint8Array, off: number): number {
  return (b[off] | (b[off + 1] << 8) | (b[off + 2] << 16) | (b[off + 3] << 24)) >>> 0
}

/** A 64-bit field as a number; a high word past 2^21 cannot be a real size and reads as -1. */
function readU64(b: Uint8Array, off: number): number {
  const hi = readU32(b, off + 4)
  if (hi >= 0x200000) return -1
  return hi * 0x100000000 + readU32(b, off)
}

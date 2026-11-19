// ── Random-access ZIP reader ─────────────────────────────────────────
// Reads a ZIP archive from a source that can be read at any offset,
// without ever holding the archive.
//
// `ZipReader` needs the whole file in memory because it reads the
// central directory from the end and then indexes into the bytes.
// `ZipStreamReader` avoids that by walking local headers front to back,
// which only works when every part a reader depends on comes before the
// one it streams. Excel, ExcelJS's streaming writer and the large
// instrument exports this reader was written for each put the worksheet
// before some part it needs.
//
// A seekable source needs neither compromise: read the end records and
// the directory, read the small parts whole, and pull the big one
// through the inflater a bounded window at a time. Part order, data
// descriptors and ZIP64 then make no difference, and what is held is the
// directory, the metadata parts and one window — not a function of the
// file's size.
//
// Every size here comes from the archive, which is untrusted, and the
// source's `size` comes from the caller, which bounds nothing on a sparse
// file. So each allocation a field controls is checked before it is made:
// the directory against MAX_CENTRAL_DIRECTORY_BYTES, everything held
// against `maxBufferedBytes`, and an entry's extent against the source.

import type { RandomAccessInput } from "../_types"
import { ParseError, ZipError } from "../errors"
import { MAX_CENTRAL_DIRECTORY_BYTES, MAX_DECOMPRESSED_BYTES, MAX_INPUT_BYTES } from "../limits"
import {
  type CentralDirEntry,
  type DirectoryBounds,
  CENTRAL_DIR_ENTRY_MIN_BYTES,
  EOCD_MIN_BYTES,
  EOCD_SEARCH_BYTES,
  LOCAL_HEADER_BYTES,
  readCentralDirectory,
  readEocd,
  readZip64Eocd,
  readZip64Locator,
  SIG_LOCAL_FILE,
  ZIP64_EOCD_BYTES,
  ZIP64_LOCATOR_BYTES,
} from "./central-directory"
import { crc32 } from "./deflate"
import { decompressionCap, positiveOr } from "./bounds"
import { byteLimitStream } from "./byte-limit"
import { canInflateRaw } from "./capability"
import { inflateRawStream } from "./inflate-stream"
import { decompressDeflateRaw } from "./reader"

/**
 * The largest single read the reader makes.
 *
 * Every read goes through {@link readRange}, which splits anything larger,
 * so this is also the most one read can cost — the window the worksheet
 * streams through. 1 MiB matches `STREAM_CHUNK_BYTES`: large enough that
 * per-read overhead (a syscall, an HTTP round trip) does not dominate,
 * small enough that the window is noise next to the parsed metadata.
 */
export const RANDOM_ACCESS_WINDOW_BYTES: number = 1024 * 1024

export interface RandomAccessZipOptions {
  /** Per-entry decompression ceiling. Default {@link MAX_DECOMPRESSED_BYTES}. */
  maxDecompressedBytes?: number
  /**
   * Ceiling on everything held from the source: the directory plus every
   * part read whole. Default {@link MAX_INPUT_BYTES}.
   */
  maxBufferedBytes?: number
}

/** Whether `value` has the {@link RandomAccessInput} shape. */
export function isRandomAccessInput(value: unknown): value is RandomAccessInput {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as { getReader?: unknown }).getReader !== "function" &&
    typeof (value as RandomAccessInput).read === "function" &&
    typeof (value as RandomAccessInput).size === "number"
  )
}

/**
 * Adapt a `Blob` (or `File`) to {@link RandomAccessInput}.
 *
 * `Blob.slice` is lazy, so a read costs only the bytes asked for — and on
 * Node 24, `fs.openAsBlob(path)` returns a file-backed `Blob`, which makes
 * this the file adapter too without the library importing `node:fs`.
 */
export function randomAccessFromBlob(blob: Blob): RandomAccessInput {
  return {
    size: blob.size,
    async read(offset: number, length: number): Promise<Uint8Array> {
      return new Uint8Array(await blob.slice(offset, offset + length).arrayBuffer())
    },
  }
}

export class RandomAccessZip {
  private readonly entryMap = new Map<string, CentralDirEntry>()
  private buffered: number
  private readonly maxDecompressedBytes: number
  private readonly maxBufferedBytes: number

  private constructor(
    private readonly source: RandomAccessInput,
    entries: CentralDirEntry[],
    buffered: number,
    maxDecompressedBytes: number,
    maxBufferedBytes: number,
  ) {
    for (const e of entries) this.entryMap.set(e.fileName, e)
    this.buffered = buffered
    this.maxDecompressedBytes = maxDecompressedBytes
    this.maxBufferedBytes = maxBufferedBytes
  }

  /** Read the end records and the central directory. */
  static async open(
    source: RandomAccessInput,
    options: RandomAccessZipOptions = {},
  ): Promise<RandomAccessZip> {
    const maxDecompressedBytes = options.maxDecompressedBytes ?? MAX_DECOMPRESSED_BYTES
    const maxBufferedBytes = positiveOr(options.maxBufferedBytes, MAX_INPUT_BYTES)
    const size = sourceSize(source)
    if (size < EOCD_MIN_BYTES) {
      throw new ZipError("Data too small to be a valid ZIP archive")
    }

    // The EOCD is within the last 65,557 bytes; the ZIP64 locator, when
    // there is one, is the 20 bytes before it. One read covers both.
    const tailLength = Math.min(size, EOCD_SEARCH_BYTES + ZIP64_LOCATOR_BYTES)
    const tailStart = size - tailLength
    const tail = await readRange(source, tailStart, tailLength)
    const eocd = readEocd(tail, tailStart)

    let bounds: DirectoryBounds = eocd
    if (eocd.needsZip64) {
      const locatorStart = eocd.eocdOffset - ZIP64_LOCATOR_BYTES - tailStart
      if (locatorStart < 0) {
        throw new ZipError("ZIP64 End of Central Directory locator not found")
      }
      const recordOffset = readZip64Locator(
        tail.subarray(locatorStart, locatorStart + ZIP64_LOCATOR_BYTES),
      )
      if (recordOffset + ZIP64_EOCD_BYTES > size) {
        throw new ZipError("ZIP64 End of Central Directory record extends beyond file")
      }
      bounds = readZip64Eocd(await readRange(source, recordOffset, ZIP64_EOCD_BYTES))
    }

    const { centralDirOffset, centralDirSize, entryCount } = bounds
    if (centralDirOffset + centralDirSize > size) {
      throw new ZipError("Central Directory extends beyond file")
    }
    if (centralDirSize > MAX_CENTRAL_DIRECTORY_BYTES) {
      throw new ZipError(
        `Central Directory of ${centralDirSize} bytes exceeds the limit of ` +
          `${MAX_CENTRAL_DIRECTORY_BYTES} bytes`,
      )
    }
    // Each record is at least 46 bytes. A count the stated size cannot
    // hold is the directory contradicting itself, and believing it is
    // what would turn a small directory into a long loop.
    if (entryCount * CENTRAL_DIR_ENTRY_MIN_BYTES > centralDirSize) {
      throw new ZipError(
        `Central Directory claims ${entryCount} entries but is only ${centralDirSize} bytes`,
      )
    }
    const buffered = tailLength + centralDirSize
    checkBudget(buffered, maxBufferedBytes)

    const dir = await readRange(source, centralDirOffset, centralDirSize)
    const entries = readCentralDirectory(dir, 0, entryCount)
    return new RandomAccessZip(source, entries, buffered, maxDecompressedBytes, maxBufferedBytes)
  }

  /** Total length of the source, as it declared it. */
  get size(): number {
    return sourceSize(this.source)
  }

  /** List all entry paths in the archive. */
  entries(): string[] {
    return [...this.entryMap.keys()]
  }

  has(path: string): boolean {
    return this.entryMap.has(path)
  }

  /**
   * Whether this reader can extract `path` from the directory alone.
   *
   * Shapes it declines, all of which `ZipReader` handles with the whole
   * file in hand: a method other than STORE or DEFLATE (which it reports
   * as unsupported), and a directory record with no compressed size where
   * the directory is not the last word on it. That is a non-empty body,
   * or any entry with a data descriptor: for those `ZipReader` takes the
   * sizes from the local header, or scans for the descriptor, rather than
   * believing a zero — so "both sizes zero" means empty only when bit 3
   * is clear. The caller reads the file whole for these, so the outcome,
   * error message included, is the buffered reader's.
   */
  canExtract(path: string): boolean {
    const e = this.entryMap.get(path)
    if (!e) return false
    if (e.compressionMethod !== 0 && e.compressionMethod !== 8) return false
    if (e.compressedSize === 0 && (e.hasDataDescriptor || e.uncompressedSize !== 0)) return false
    return true
  }

  /**
   * Read and decompress one entry whole, CRC-checked. Counts against the
   * buffer budget.
   */
  async extract(path: string): Promise<Uint8Array> {
    const entry = this.require(path)
    const dataStart = await this.dataStart(entry)
    this.buffered += entry.compressedSize
    checkBudget(this.buffered, this.maxBufferedBytes)

    const compressed = await readRange(this.source, dataStart, entry.compressedSize)
    let result: Uint8Array
    if (entry.compressionMethod === 0) {
      result = compressed
    } else if (entry.compressedSize === 0 && entry.uncompressedSize === 0) {
      result = new Uint8Array(0)
    } else {
      result = await decompressDeflateRaw(compressed, this.capFor(entry))
    }

    if (entry.crc32 !== 0 && result.length > 0) {
      const actual = crc32(result)
      if (actual !== entry.crc32) {
        throw new ZipError(
          `CRC-32 mismatch for ${entry.fileName}: expected 0x${entry.crc32.toString(16)}, got 0x${actual.toString(16)}`,
        )
      }
    }
    return result
  }

  /**
   * Stream one entry's decompressed bytes, pulling its compressed body
   * from the source {@link RANDOM_ACCESS_WINDOW_BYTES} at a time, only as
   * fast as the consumer reads. Nothing is fetched until the first pull,
   * and cancelling stops the reads.
   */
  extractStream(path: string): ReadableStream<Uint8Array> {
    const entry = this.require(path)
    let position = -1
    let remaining = entry.compressedSize
    const raw = new ReadableStream<Uint8Array>(
      {
        pull: async (controller) => {
          if (position < 0) position = await this.dataStart(entry)
          if (remaining <= 0) {
            controller.close()
            return
          }
          const take = Math.min(remaining, RANDOM_ACCESS_WINDOW_BYTES)
          const chunk = await readRange(this.source, position, take)
          position += take
          remaining -= take
          controller.enqueue(chunk)
          if (remaining <= 0) controller.close()
        },
      },
      // One window in flight, not the default's read-ahead of several.
      { highWaterMark: 0 },
    )
    if (entry.compressionMethod === 0) {
      // A stored body is its own output, so the decompression ceiling is a
      // ceiling on what is read. Without it a directory claiming a huge
      // stored sheet in a sparse source would be read window by window to
      // the end of whatever the source pretends to hold.
      if (entry.compressedSize > this.maxDecompressedBytes) {
        return erroredStream(
          new ZipError(
            `Decompressed size exceeds limit of ${this.maxDecompressedBytes} bytes (possible zip bomb)`,
          ),
        )
      }
      return raw.pipeThrough(byteLimitStream(this.capFor(entry)))
    }
    if (entry.compressedSize === 0 && entry.uncompressedSize === 0) {
      return new ReadableStream<Uint8Array>({ start: (c) => c.close() })
    }
    if (!canInflateRaw()) {
      // Without the native inflater `inflateRawStream` buffers the whole
      // compressed body, whose length is the directory's claim — so it is
      // held to the same budget as a part read whole. Refused here, from
      // the claim, rather than as the body accumulates: the length is
      // known exactly and already checked against the source, so counting
      // it up front costs nothing and spares reading what would be thrown
      // away. inflateRawStream itself stays budget-free; the stream path
      // that also uses it buffered there on main too, under its own log.
      this.buffered += entry.compressedSize
      if (this.buffered > this.maxBufferedBytes) {
        return erroredStream(budgetExceeded(this.buffered, this.maxBufferedBytes))
      }
    }
    return inflateRawStream(raw, this.capFor(entry), entry.fileName)
  }

  // ── Private ─────────────────────────────────────────────────────

  private require(path: string): CentralDirEntry {
    const entry = this.entryMap.get(path)
    if (!entry) throw new ZipError(`Entry not found: ${path}`)
    if (entry.compressionMethod !== 0 && entry.compressionMethod !== 8) {
      throw new ZipError(
        `Unsupported compression method ${entry.compressionMethod} for entry: ${entry.fileName}`,
      )
    }
    return entry
  }

  /**
   * Same ceiling as `ZipReader`: the declared size when stated, the hard
   * cap otherwise.
   */
  private capFor(entry: CentralDirEntry): number {
    return decompressionCap(entry.uncompressedSize, this.maxDecompressedBytes)
  }

  /**
   * Read an entry's local header for the lengths of its name and extra
   * field — which may differ from the directory's copy — and return where
   * the body starts, having checked the body lies inside the source.
   */
  private async dataStart(entry: CentralDirEntry): Promise<number> {
    const size = this.size
    const at = entry.localHeaderOffset
    if (at + LOCAL_HEADER_BYTES > size) {
      throw new ZipError("Local file header extends beyond file")
    }
    const header = await readRange(this.source, at, LOCAL_HEADER_BYTES)
    const view = new DataView(header.buffer, header.byteOffset, header.byteLength)
    if (view.getUint32(0, true) !== SIG_LOCAL_FILE) {
      throw new ZipError(`Invalid local file header signature at offset ${at}`)
    }
    const start = at + LOCAL_HEADER_BYTES + view.getUint16(26, true) + view.getUint16(28, true)
    if (start + entry.compressedSize > size) {
      throw new ZipError(`Compressed data extends beyond file for entry: ${entry.fileName}`)
    }
    return start
  }
}

/**
 * Read `length` bytes at `offset`, never more than one window per call to
 * the source. The caller has bounded `length` before calling.
 */
export async function readRange(
  source: RandomAccessInput,
  offset: number,
  length: number,
): Promise<Uint8Array> {
  if (length <= RANDOM_ACCESS_WINDOW_BYTES) return readExact(source, offset, length)
  const out = new Uint8Array(length)
  for (let at = 0; at < length; at += RANDOM_ACCESS_WINDOW_BYTES) {
    const take = Math.min(RANDOM_ACCESS_WINDOW_BYTES, length - at)
    out.set(await readExact(source, offset + at, take), at)
  }
  return out
}

/**
 * One call to the source, held to the contract. A short read would
 * otherwise be parsed as truncation somewhere far from the cause, and a
 * long one is what an HTTP server that ignores `Range` sends: the whole
 * file, which is exactly the allocation this reader exists to avoid.
 */
async function readExact(
  source: RandomAccessInput,
  offset: number,
  length: number,
): Promise<Uint8Array> {
  if (length === 0) return new Uint8Array(0)
  const got: unknown = await source.read(offset, length)
  // `ArrayBuffer.isView` rather than `instanceof`, so bytes from another
  // realm (a worker, an iframe) are accepted.
  if (!ArrayBuffer.isView(got) || got.byteLength !== length) {
    const actual = ArrayBuffer.isView(got) ? `${got.byteLength} bytes` : typeof got
    throw new ParseError(
      `RandomAccessInput.read(${offset}, ${length}) returned ${actual}; ` +
        "it must resolve with exactly the bytes asked for",
    )
  }
  return got instanceof Uint8Array ? got : new Uint8Array(got.buffer, got.byteOffset, length)
}

function erroredStream(err: Error): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({ start: (c) => c.error(err) })
}

/** The source's declared size, checked to be a byte count. */
export function sourceSize(source: RandomAccessInput): number {
  const size = source.size
  if (!Number.isSafeInteger(size) || size < 0) {
    throw new ParseError(`RandomAccessInput.size must be a non-negative integer, got ${size}`)
  }
  return size
}

function checkBudget(buffered: number, max: number): void {
  if (buffered > max) throw budgetExceeded(buffered, max)
}

function budgetExceeded(buffered: number, max: number): ParseError {
  return new ParseError(
    `Reading this archive would hold ${buffered} bytes from the source, over the ` +
      `maximum of ${max}. Raise \`maxInputBytes\` if the file really needs it.`,
  )
}

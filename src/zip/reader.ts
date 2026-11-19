// ── ZIP Archive Reader ──────────────────────────────────────────────
// Reads ZIP files (XLSX, ODS) from Uint8Array.
// Supports STORE (method 0) and DEFLATE (method 8).

import { ZipError } from "../errors"
import { MAX_DECOMPRESSED_BYTES } from "../limits"
import { byteLimitStream, chunkedStream } from "./byte-limit"
import { crc32, inflate } from "./deflate"
import { canInflateRaw } from "./capability"
import {
  type CentralDirEntry,
  type DirectoryBounds,
  readCentralDirectory,
  readEocd,
  readZip64Eocd,
  readZip64Locator,
  SIG_DATA_DESCRIPTOR,
  SIG_LOCAL_FILE,
  ZIP64_EOCD_BYTES,
  ZIP64_LOCATOR_BYTES,
  ZIP64_SENTINEL,
} from "./central-directory"

/**
 * The most DEFLATE can expand a byte by — 1032:1, the format's ceiling
 * for a maximum-length match emitted from a minimum-length code. Used to
 * tell a declared uncompressed size that the compressed body could have
 * produced from one it could not.
 */
const MAX_DEFLATE_RATIO = 1032

// ── Decompression ───────────────────────────────────────────────────

async function decompressDeflateRaw(
  data: Uint8Array,
  maxBytes = MAX_DECOMPRESSED_BYTES,
): Promise<Uint8Array> {
  if (canInflateRaw()) {
    try {
      const ds = new DecompressionStream("deflate-raw")
      const writer = ds.writable.getWriter()
      const reader = ds.readable.getReader()

      // Write data and close. Both promises are deliberately not awaited
      // (awaiting the write before reading would deadlock on backpressure),
      // but they must still be handled: when the size cap below cancels the
      // reader, they reject with AbortError, and an unhandled rejection
      // takes the whole process down instead of surfacing the ZipError.
      void writer.write(data as unknown as BufferSource).catch(() => {})
      void writer.close().catch(() => {})

      // Read all chunks
      const chunks: Uint8Array[] = []
      let totalLen = 0

      for (;;) {
        const { done, value } = await reader.read()
        if (done) break
        totalLen += value.length
        if (totalLen > maxBytes) {
          try {
            await reader.cancel()
          } catch {
            // ignore
          }
          throw new ZipError(
            `Decompressed size exceeds limit of ${maxBytes} bytes (possible zip bomb)`,
          )
        }
        chunks.push(value)
      }

      // Combine chunks
      const result = new Uint8Array(totalLen)
      let offset = 0
      for (const chunk of chunks) {
        result.set(chunk, offset)
        offset += chunk.length
      }
      return result
    } catch (err) {
      // A size-limit breach is a hard failure — don't retry with the
      // pure-TS path (which would just OOM the same way).
      if (err instanceof ZipError) throw err
      // Otherwise fall through to pure TS implementation
    }
  }

  // Pure TypeScript fallback
  return inflate(data, maxBytes)
}

// ── ZipReader ───────────────────────────────────────────────────────

export class ZipReader {
  private view: DataView
  private centralDir: CentralDirEntry[] = []
  private entryMap: Map<string, CentralDirEntry> = new Map()

  /**
   * Absolute ceiling on what any one entry may decompress to — the
   * zip-bomb bound. Defaults to {@link MAX_DECOMPRESSED_BYTES}; readers
   * pass `ReadOptions.maxDecompressedBytes` through when the caller has
   * said the input is trusted. See #471.
   */
  private maxDecompressedBytes: number

  constructor(
    private data: Uint8Array,
    maxDecompressedBytes: number = MAX_DECOMPRESSED_BYTES,
  ) {
    this.maxDecompressedBytes = maxDecompressedBytes
    if (data.length < 22) {
      throw new ZipError("Data too small to be a valid ZIP archive")
    }
    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength)
    this.readEndOfCentralDir()
  }

  /** List all entry paths in the archive */
  entries(): string[] {
    return this.centralDir.map((e) => e.fileName)
  }

  /** Check if an entry exists */
  has(path: string): boolean {
    return this.entryMap.has(path)
  }

  /**
   * What the central directory says an entry expands to, or `undefined`
   * when it does not say.
   *
   * Lets a reader choose how to parse a part *before* paying to
   * decompress it — a worksheet past the string ceiling has to be parsed
   * as a stream, and finding that out by buffering it first costs the
   * decompression twice and the peak memory of a buffer that gets thrown
   * away. See #503.
   *
   * Zero means "not declared": it is what a producer writing a data
   * descriptor leaves behind, and it is not a small entry, so it may not
   * be reported as one — the caller is told nothing and falls back to
   * finding out the slow way. A ZIP64 size needs no handling here;
   * `readCentralDir` has already resolved the sentinel from the extra
   * field, or thrown.
   *
   * A size the compressed body could not possibly produce is also not
   * declared. Nothing verifies this field, and the answer decides whether
   * a part is read whole — CRC-32 checked — or as a stream, which has no
   * whole entry to check. Without the plausibility test a 40 KB worksheet
   * whose central directory claims 600 MB would take the streaming route
   * and skip the checksum, so a corrupt entry could buy its way out of
   * verification by lying about its size. DEFLATE cannot exceed 1032:1,
   * and a stored entry expands not at all, so a claim past that is the
   * archive contradicting itself and is not trusted.
   */
  declaredSize(path: string): number | undefined {
    const entry = this.entryMap.get(path)
    if (!entry?.uncompressedSize) return undefined
    const most =
      entry.compressionMethod === 0
        ? entry.compressedSize
        : entry.compressedSize * MAX_DEFLATE_RATIO
    if (entry.compressedSize > 0 && entry.uncompressedSize > most) return undefined
    return entry.uncompressedSize
  }

  /** Extract a single file by path */
  async extract(path: string): Promise<Uint8Array> {
    const entry = this.entryMap.get(path)
    if (!entry) {
      throw new ZipError(`Entry not found: ${path}`)
    }
    return this.extractEntry(entry)
  }

  /** Extract a single file as a ReadableStream of decompressed bytes */
  extractStream(path: string): ReadableStream<Uint8Array> {
    const entry = this.entryMap.get(path)
    if (!entry) {
      throw new ZipError(`Entry not found: ${path}`)
    }
    return this.extractEntryStream(entry)
  }

  /** Extract all files */
  async extractAll(): Promise<Map<string, Uint8Array>> {
    const result = new Map<string, Uint8Array>()
    for (const entry of this.centralDir) {
      // Skip directory entries
      if (entry.fileName.endsWith("/")) continue
      const data = await this.extractEntry(entry)
      result.set(entry.fileName, data)
    }
    return result
  }

  // ── Private ─────────────────────────────────────────────────────

  private readEndOfCentralDir(): void {
    const eocd = readEocd(this.data, 0)
    let bounds: DirectoryBounds = eocd
    if (eocd.needsZip64) bounds = this.readZip64EndOfCentralDir(eocd.eocdOffset)
    this.centralDir = readCentralDirectory(this.data, bounds.centralDirOffset, bounds.entryCount)
    for (const entry of this.centralDir) this.entryMap.set(entry.fileName, entry)
  }

  /**
   * Resolve the real directory bounds from the ZIP64 EOCD record.
   *
   * The 20-byte ZIP64 locator sits immediately before the classic EOCD and
   * points at the ZIP64 EOCD record, which repeats the same fields at 64-bit
   * width.
   */
  private readZip64EndOfCentralDir(eocdOffset: number): DirectoryBounds {
    const locatorOffset = eocdOffset - ZIP64_LOCATOR_BYTES
    if (locatorOffset < 0) {
      throw new ZipError("ZIP64 End of Central Directory locator not found")
    }
    const recordOffset = readZip64Locator(this.data.subarray(locatorOffset, eocdOffset))
    if (recordOffset + ZIP64_EOCD_BYTES > this.data.length) {
      throw new ZipError("ZIP64 End of Central Directory record extends beyond file")
    }
    return readZip64Eocd(this.data.subarray(recordOffset, recordOffset + ZIP64_EOCD_BYTES))
  }

  private async extractEntry(entry: CentralDirEntry): Promise<Uint8Array> {
    const pos = entry.localHeaderOffset

    if (pos + 30 > this.data.length) {
      throw new ZipError("Local file header extends beyond file")
    }

    const sig = this.view.getUint32(pos, true)
    if (sig !== SIG_LOCAL_FILE) {
      throw new ZipError(`Invalid local file header signature at offset ${pos}`)
    }

    const fileNameLength = this.view.getUint16(pos + 26, true)
    const extraFieldLength = this.view.getUint16(pos + 28, true)
    const dataStart = pos + 30 + fileNameLength + extraFieldLength

    // Use sizes from central directory (authoritative), not local header.
    // Local header may have zeros when data descriptor is used.
    let { compressedSize, uncompressedSize, crc32: expectedCrc } = entry

    // Handle data descriptor case (Lotus/Excel): sizes might be zero in
    // both local header and central dir in malformed files.
    if (entry.hasDataDescriptor && compressedSize === 0) {
      // Try to read from data descriptor after compressed data.
      // This is a tricky case; we rely on central dir being authoritative.
      // If central dir also has zeros, we need to find the data descriptor.
      // A sentinel here means the real value lives in the local ZIP64 extra
      // field; taking it literally would read 4 GiB past the entry.
      const localCompressedSize = this.view.getUint32(pos + 18, true)
      if (localCompressedSize > 0 && localCompressedSize !== ZIP64_SENTINEL) {
        compressedSize = localCompressedSize
      }
      const localUncompressedSize = this.view.getUint32(pos + 22, true)
      if (localUncompressedSize > 0 && localUncompressedSize !== ZIP64_SENTINEL) {
        uncompressedSize = localUncompressedSize
      }

      // If still zero, try to find data descriptor
      if (compressedSize === 0) {
        const found = this.findDataDescriptor(dataStart)
        if (found) {
          expectedCrc = found.crc
          compressedSize = found.compressedSize
          uncompressedSize = found.uncompressedSize
        }
      }
    }

    if (dataStart + compressedSize > this.data.length) {
      throw new ZipError(`Compressed data extends beyond file for entry: ${entry.fileName}`)
    }

    const compressedData = this.data.subarray(dataStart, dataStart + compressedSize)

    let result: Uint8Array

    if (entry.compressionMethod === 0) {
      // STORE — no compression
      result = compressedData
    } else if (entry.compressionMethod === 8) {
      // DEFLATE
      if (compressedSize === 0 && uncompressedSize === 0) {
        result = new Uint8Array(0)
      } else {
        // Bound output by the central-directory uncompressedSize (when
        // declared and trustworthy) as well as the absolute hard cap.
        const declaredCap =
          uncompressedSize > 0
            ? Math.min(uncompressedSize, this.maxDecompressedBytes)
            : this.maxDecompressedBytes
        result = await decompressDeflateRaw(compressedData, declaredCap)
      }
    } else {
      throw new ZipError(
        `Unsupported compression method ${entry.compressionMethod} for entry: ${entry.fileName}`,
      )
    }

    // Verify CRC-32 (skip if CRC is 0 — some generators omit it)
    if (expectedCrc !== 0 && result.length > 0) {
      const actualCrc = crc32(result)
      if (actualCrc !== expectedCrc) {
        throw new ZipError(
          `CRC-32 mismatch for ${entry.fileName}: expected 0x${expectedCrc.toString(16)}, got 0x${actualCrc.toString(16)}`,
        )
      }
    }

    return result
  }

  private extractEntryStream(entry: CentralDirEntry): ReadableStream<Uint8Array> {
    const pos = entry.localHeaderOffset

    if (pos + 30 > this.data.length) {
      throw new ZipError("Local file header extends beyond file")
    }

    const sig = this.view.getUint32(pos, true)
    if (sig !== SIG_LOCAL_FILE) {
      throw new ZipError(`Invalid local file header signature at offset ${pos}`)
    }

    const fileNameLength = this.view.getUint16(pos + 26, true)
    const extraFieldLength = this.view.getUint16(pos + 28, true)
    const dataStart = pos + 30 + fileNameLength + extraFieldLength

    let { compressedSize } = entry

    if (entry.hasDataDescriptor && compressedSize === 0) {
      // Same sentinel guard as extractEntry: 0xFFFFFFFF means the real
      // size lives in the ZIP64 extra field, so taking it literally
      // reads 4 GiB past the entry. A ZIP64 streaming producer writes
      // exactly this, and without the check extract() succeeded while
      // extractStream() threw. See #393.
      const localCompressedSize = this.view.getUint32(pos + 18, true)
      if (localCompressedSize > 0 && localCompressedSize !== ZIP64_SENTINEL) {
        compressedSize = localCompressedSize
      }
      if (compressedSize === 0) {
        const found = this.findDataDescriptor(dataStart)
        if (found) {
          compressedSize = found.compressedSize
        }
      }
    }

    if (dataStart + compressedSize > this.data.length) {
      throw new ZipError(`Compressed data extends beyond file for entry: ${entry.fileName}`)
    }

    const compressedData = this.data.subarray(dataStart, dataStart + compressedSize)

    if (entry.compressionMethod === 0) {
      // STORE — no decompression, but still emitted in pieces
      return chunkedStream(compressedData)
    }

    if (entry.compressionMethod === 8) {
      // DEFLATE — stream through DecompressionStream if available
      if (compressedSize === 0 && entry.uncompressedSize === 0) {
        return new ReadableStream<Uint8Array>({
          start(controller) {
            controller.close()
          },
        })
      }

      // Bound the output by the declared uncompressed size (when present)
      // and the absolute cap — exactly what the buffered path does. Without
      // it the native DecompressionStream below happily expanded a bomb the
      // buffered reader rejects.
      const declaredCap =
        entry.uncompressedSize > 0
          ? Math.min(entry.uncompressedSize, this.maxDecompressedBytes)
          : this.maxDecompressedBytes

      if (canInflateRaw()) {
        const inputStream = new ReadableStream({
          start(controller) {
            controller.enqueue(compressedData)
            controller.close()
          },
        })
        return (
          inputStream.pipeThrough(
            new DecompressionStream("deflate-raw"),
          ) as ReadableStream<Uint8Array>
        ).pipeThrough(byteLimitStream(declaredCap))
      }

      // Fallback: inflate synchronously and emit as stream
      return chunkedStream(inflate(compressedData, declaredCap))
    }

    throw new ZipError(
      `Unsupported compression method ${entry.compressionMethod} for entry: ${entry.fileName}`,
    )
  }

  /** Attempt to locate a data descriptor after the compressed data */
  private findDataDescriptor(
    dataStart: number,
  ): { crc: number; compressedSize: number; uncompressedSize: number } | null {
    // Scan for data descriptor signature
    for (let pos = dataStart; pos < this.data.length - 16; pos++) {
      if (this.view.getUint32(pos, true) === SIG_DATA_DESCRIPTOR) {
        return {
          crc: this.view.getUint32(pos + 4, true),
          compressedSize: this.view.getUint32(pos + 8, true),
          uncompressedSize: this.view.getUint32(pos + 12, true),
        }
      }
    }
    return null
  }
}

// ── ZIP Archive Reader ──────────────────────────────────────────────
// Reads ZIP files (XLSX, ODS) from Uint8Array.
// Supports STORE (method 0) and DEFLATE (method 8).

import { ZipError } from "../errors"
import { MAX_DECOMPRESSED_BYTES } from "../limits"
import { byteLimitStream } from "./byte-limit"
import { crc32, inflate } from "./deflate"

// ── ZIP Signatures ──────────────────────────────────────────────────

const SIG_LOCAL_FILE = 0x04034b50
const SIG_CENTRAL_DIR = 0x02014b50
const SIG_END_OF_CENTRAL_DIR = 0x06054b50
const SIG_DATA_DESCRIPTOR = 0x08074b50
const SIG_ZIP64_EOCD = 0x06064b50
const SIG_ZIP64_EOCD_LOCATOR = 0x07064b50

/** 0xFFFFFFFF marker that signals a 32-bit ZIP field overflowed into ZIP64. */
const ZIP64_SENTINEL = 0xffffffff
/** 0xFFFF marker for the 16-bit entry-count fields. */
const ZIP64_SENTINEL_16 = 0xffff
/** Header id of the ZIP64 extended information extra field. */
const ZIP64_EXTRA_ID = 0x0001

// ── Types ───────────────────────────────────────────────────────────

interface CentralDirEntry {
  fileName: string
  compressionMethod: number
  compressedSize: number
  uncompressedSize: number
  crc32: number
  localHeaderOffset: number
  /** Bit 3 of general purpose flag — data descriptor present */
  hasDataDescriptor: boolean
}

// ── Decompression ───────────────────────────────────────────────────

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

async function decompressDeflateRaw(
  data: Uint8Array,
  maxBytes = MAX_DECOMPRESSED_BYTES,
): Promise<Uint8Array> {
  if (checkDecompressionStream()) {
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

  constructor(private data: Uint8Array) {
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
    // EOCD is at least 22 bytes and located at the end of the file.
    // We need to search backwards because there may be a comment.
    const minOffset = Math.max(0, this.data.length - 65557) // 22 + 65535 max comment

    let eocdOffset = -1
    for (let i = this.data.length - 22; i >= minOffset; i--) {
      if (this.view.getUint32(i, true) === SIG_END_OF_CENTRAL_DIR) {
        eocdOffset = i
        break
      }
    }

    if (eocdOffset === -1) {
      throw new ZipError("End of Central Directory not found — not a valid ZIP file")
    }

    let centralDirSize = this.view.getUint32(eocdOffset + 12, true)
    let centralDirOffset = this.view.getUint32(eocdOffset + 16, true)
    let entryCount = this.view.getUint16(eocdOffset + 10, true)

    // ZIP64: when the entry count or central-directory size/offset overflows
    // the 16-/32-bit EOCD fields they hold a 0xFFFF / 0xFFFFFFFF sentinel and
    // the real values live in a ZIP64 EOCD record. Producers also emit these
    // records defensively on small archives, so a sentinel is not by itself
    // evidence that the archive is huge.
    if (
      entryCount === ZIP64_SENTINEL_16 ||
      centralDirSize === ZIP64_SENTINEL ||
      centralDirOffset === ZIP64_SENTINEL
    ) {
      const zip64 = this.readZip64EndOfCentralDir(eocdOffset)
      entryCount = zip64.entryCount
      centralDirSize = zip64.centralDirSize
      centralDirOffset = zip64.centralDirOffset
    }

    this.readCentralDirectory(centralDirOffset, centralDirSize, entryCount)
  }

  /**
   * Resolve the real directory bounds from the ZIP64 EOCD record.
   *
   * The 20-byte ZIP64 locator sits immediately before the classic EOCD and
   * points at the ZIP64 EOCD record, which repeats the same fields at 64-bit
   * width.
   */
  private readZip64EndOfCentralDir(eocdOffset: number): {
    entryCount: number
    centralDirSize: number
    centralDirOffset: number
  } {
    const locatorOffset = eocdOffset - 20
    if (locatorOffset < 0 || this.view.getUint32(locatorOffset, true) !== SIG_ZIP64_EOCD_LOCATOR) {
      throw new ZipError("ZIP64 End of Central Directory locator not found")
    }

    const recordOffset = this.readUint64(locatorOffset + 8, "ZIP64 EOCD offset")
    if (recordOffset + 56 > this.data.length) {
      throw new ZipError("ZIP64 End of Central Directory record extends beyond file")
    }
    if (this.view.getUint32(recordOffset, true) !== SIG_ZIP64_EOCD) {
      throw new ZipError("Invalid ZIP64 End of Central Directory signature")
    }

    return {
      entryCount: this.readUint64(recordOffset + 32, "ZIP64 entry count"),
      centralDirSize: this.readUint64(recordOffset + 40, "ZIP64 central directory size"),
      centralDirOffset: this.readUint64(recordOffset + 48, "ZIP64 central directory offset"),
    }
  }

  /** Read a 64-bit little-endian field, refusing values JS can't index with. */
  private readUint64(offset: number, what: string): number {
    if (offset + 8 > this.data.length) {
      throw new ZipError(`${what} extends beyond file`)
    }
    const value = this.view.getBigUint64(offset, true)
    if (value > BigInt(Number.MAX_SAFE_INTEGER)) {
      throw new ZipError(`${what} (${value}) exceeds the addressable range`)
    }
    return Number(value)
  }

  private readCentralDirectory(offset: number, _size: number, expectedCount: number): void {
    let pos = offset

    for (let i = 0; i < expectedCount; i++) {
      if (pos + 46 > this.data.length) {
        throw new ZipError("Central Directory entry extends beyond file")
      }

      const sig = this.view.getUint32(pos, true)
      if (sig !== SIG_CENTRAL_DIR) {
        throw new ZipError(
          `Invalid Central Directory signature at offset ${pos}: 0x${sig.toString(16)}`,
        )
      }

      const generalFlag = this.view.getUint16(pos + 8, true)
      const compressionMethod = this.view.getUint16(pos + 10, true)
      const entryCrc32 = this.view.getUint32(pos + 16, true)
      let compressedSize = this.view.getUint32(pos + 20, true)
      let uncompressedSize = this.view.getUint32(pos + 24, true)
      const fileNameLength = this.view.getUint16(pos + 28, true)
      const extraFieldLength = this.view.getUint16(pos + 30, true)
      const commentLength = this.view.getUint16(pos + 32, true)
      let localHeaderOffset = this.view.getUint32(pos + 42, true)

      const fileNameBytes = this.data.subarray(pos + 46, pos + 46 + fileNameLength)
      const fileName = new TextDecoder().decode(fileNameBytes)

      // Any field left at the sentinel is carried at 64-bit width in the
      // ZIP64 extra field instead.
      if (
        compressedSize === ZIP64_SENTINEL ||
        uncompressedSize === ZIP64_SENTINEL ||
        localHeaderOffset === ZIP64_SENTINEL
      ) {
        const zip64 = this.readZip64Extra(
          pos + 46 + fileNameLength,
          extraFieldLength,
          uncompressedSize === ZIP64_SENTINEL,
          compressedSize === ZIP64_SENTINEL,
          localHeaderOffset === ZIP64_SENTINEL,
          fileName,
        )
        if (zip64.uncompressedSize !== undefined) uncompressedSize = zip64.uncompressedSize
        if (zip64.compressedSize !== undefined) compressedSize = zip64.compressedSize
        if (zip64.localHeaderOffset !== undefined) localHeaderOffset = zip64.localHeaderOffset
      }

      const hasDataDescriptor = (generalFlag & 0x08) !== 0

      const entry: CentralDirEntry = {
        fileName,
        compressionMethod,
        compressedSize,
        uncompressedSize,
        crc32: entryCrc32,
        localHeaderOffset,
        hasDataDescriptor,
      }

      this.centralDir.push(entry)
      this.entryMap.set(fileName, entry)

      pos += 46 + fileNameLength + extraFieldLength + commentLength
    }
  }

  /**
   * Pull the overflowed fields out of a central-directory ZIP64 extra field.
   *
   * The record is positional: only the fields whose 32-bit counterpart held
   * the sentinel are present, always in the order size → compressed size →
   * local header offset → disk number.
   */
  private readZip64Extra(
    offset: number,
    length: number,
    wantUncompressed: boolean,
    wantCompressed: boolean,
    wantOffset: boolean,
    fileName: string,
  ): { uncompressedSize?: number; compressedSize?: number; localHeaderOffset?: number } {
    const end = Math.min(offset + length, this.data.length)
    let pos = offset

    while (pos + 4 <= end) {
      const headerId = this.view.getUint16(pos, true)
      const dataSize = this.view.getUint16(pos + 2, true)
      const dataStart = pos + 4

      if (headerId !== ZIP64_EXTRA_ID) {
        pos = dataStart + dataSize
        continue
      }

      const dataEnd = dataStart + dataSize
      let cursor = dataStart
      const result: {
        uncompressedSize?: number
        compressedSize?: number
        localHeaderOffset?: number
      } = {}

      const take = (what: string): number => {
        if (cursor + 8 > dataEnd) {
          throw new ZipError(`Truncated ZIP64 extra field for entry: ${fileName}`)
        }
        const value = this.readUint64(cursor, what)
        cursor += 8
        return value
      }

      if (wantUncompressed) result.uncompressedSize = take("ZIP64 uncompressed size")
      if (wantCompressed) result.compressedSize = take("ZIP64 compressed size")
      if (wantOffset) result.localHeaderOffset = take("ZIP64 local header offset")

      return result
    }

    throw new ZipError(`Missing ZIP64 extra field for entry: ${fileName}`)
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
            ? Math.min(uncompressedSize, MAX_DECOMPRESSED_BYTES)
            : MAX_DECOMPRESSED_BYTES
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
      // STORE — return raw data as a stream
      return new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(compressedData)
          controller.close()
        },
      })
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
          ? Math.min(entry.uncompressedSize, MAX_DECOMPRESSED_BYTES)
          : MAX_DECOMPRESSED_BYTES

      if (checkDecompressionStream()) {
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
      const inflated = inflate(compressedData, declaredCap)
      return new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(inflated)
          controller.close()
        },
      })
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

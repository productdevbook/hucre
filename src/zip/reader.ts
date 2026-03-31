// ── ZIP Archive Reader ──────────────────────────────────────────────
// Reads ZIP files (XLSX, ODS) from Uint8Array.
// Supports STORE (method 0) and DEFLATE (method 8).

import { ZipError } from "../errors";
import { crc32, inflate } from "./deflate";

// ── ZIP Signatures ──────────────────────────────────────────────────

const SIG_LOCAL_FILE = 0x04034b50;
const SIG_CENTRAL_DIR = 0x02014b50;
const SIG_END_OF_CENTRAL_DIR = 0x06054b50;
const SIG_DATA_DESCRIPTOR = 0x08074b50;

// ── Types ───────────────────────────────────────────────────────────

interface CentralDirEntry {
  fileName: string;
  compressionMethod: number;
  compressedSize: number;
  uncompressedSize: number;
  crc32: number;
  localHeaderOffset: number;
  /** Bit 3 of general purpose flag — data descriptor present */
  hasDataDescriptor: boolean;
}

// ── Decompression ───────────────────────────────────────────────────

let hasDecompressionStream: boolean | undefined;

function checkDecompressionStream(): boolean {
  if (hasDecompressionStream === undefined) {
    try {
      hasDecompressionStream =
        typeof DecompressionStream !== "undefined" &&
        typeof ReadableStream !== "undefined" &&
        typeof Response !== "undefined";
    } catch {
      hasDecompressionStream = false;
    }
  }
  return hasDecompressionStream;
}

async function decompressDeflateRaw(data: Uint8Array): Promise<Uint8Array> {
  if (checkDecompressionStream()) {
    try {
      const ds = new DecompressionStream("deflate-raw");
      const writer = ds.writable.getWriter();
      const reader = ds.readable.getReader();

      // Write data and close
      writer.write(data as unknown as BufferSource);
      writer.close();

      // Read all chunks
      const chunks: Uint8Array[] = [];
      let totalLen = 0;

      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(value);
        totalLen += value.length;
      }

      // Combine chunks
      const result = new Uint8Array(totalLen);
      let offset = 0;
      for (const chunk of chunks) {
        result.set(chunk, offset);
        offset += chunk.length;
      }
      return result;
    } catch {
      // Fall through to pure TS implementation
    }
  }

  // Pure TypeScript fallback
  return inflate(data);
}

// ── ZipReader ───────────────────────────────────────────────────────

export class ZipReader {
  private view: DataView;
  private centralDir: CentralDirEntry[] = [];
  private entryMap: Map<string, CentralDirEntry> = new Map();

  constructor(private data: Uint8Array) {
    if (data.length < 22) {
      throw new ZipError("Data too small to be a valid ZIP archive");
    }
    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    this.readEndOfCentralDir();
  }

  /** List all entry paths in the archive */
  entries(): string[] {
    return this.centralDir.map((e) => e.fileName);
  }

  /** Check if an entry exists */
  has(path: string): boolean {
    return this.entryMap.has(path);
  }

  /** Extract a single file by path */
  async extract(path: string): Promise<Uint8Array> {
    const entry = this.entryMap.get(path);
    if (!entry) {
      throw new ZipError(`Entry not found: ${path}`);
    }
    return this.extractEntry(entry);
  }

  /** Extract a single file as a ReadableStream of decompressed bytes */
  extractStream(path: string): ReadableStream<Uint8Array> {
    const entry = this.entryMap.get(path);
    if (!entry) {
      throw new ZipError(`Entry not found: ${path}`);
    }
    return this.extractEntryStream(entry);
  }

  /** Extract all files */
  async extractAll(): Promise<Map<string, Uint8Array>> {
    const result = new Map<string, Uint8Array>();
    for (const entry of this.centralDir) {
      // Skip directory entries
      if (entry.fileName.endsWith("/")) continue;
      const data = await this.extractEntry(entry);
      result.set(entry.fileName, data);
    }
    return result;
  }

  // ── Private ─────────────────────────────────────────────────────

  private readEndOfCentralDir(): void {
    // EOCD is at least 22 bytes and located at the end of the file.
    // We need to search backwards because there may be a comment.
    const minOffset = Math.max(0, this.data.length - 65557); // 22 + 65535 max comment

    let eocdOffset = -1;
    for (let i = this.data.length - 22; i >= minOffset; i--) {
      if (this.view.getUint32(i, true) === SIG_END_OF_CENTRAL_DIR) {
        eocdOffset = i;
        break;
      }
    }

    if (eocdOffset === -1) {
      throw new ZipError("End of Central Directory not found — not a valid ZIP file");
    }

    const centralDirSize = this.view.getUint32(eocdOffset + 12, true);
    const centralDirOffset = this.view.getUint32(eocdOffset + 16, true);
    const entryCount = this.view.getUint16(eocdOffset + 10, true);

    this.readCentralDirectory(centralDirOffset, centralDirSize, entryCount);
  }

  private readCentralDirectory(offset: number, _size: number, expectedCount: number): void {
    let pos = offset;

    for (let i = 0; i < expectedCount; i++) {
      if (pos + 46 > this.data.length) {
        throw new ZipError("Central Directory entry extends beyond file");
      }

      const sig = this.view.getUint32(pos, true);
      if (sig !== SIG_CENTRAL_DIR) {
        throw new ZipError(
          `Invalid Central Directory signature at offset ${pos}: 0x${sig.toString(16)}`,
        );
      }

      const generalFlag = this.view.getUint16(pos + 8, true);
      const compressionMethod = this.view.getUint16(pos + 10, true);
      const entryCrc32 = this.view.getUint32(pos + 16, true);
      const compressedSize = this.view.getUint32(pos + 20, true);
      const uncompressedSize = this.view.getUint32(pos + 24, true);
      const fileNameLength = this.view.getUint16(pos + 28, true);
      const extraFieldLength = this.view.getUint16(pos + 30, true);
      const commentLength = this.view.getUint16(pos + 32, true);
      const localHeaderOffset = this.view.getUint32(pos + 42, true);

      const fileNameBytes = this.data.subarray(pos + 46, pos + 46 + fileNameLength);
      const fileName = new TextDecoder().decode(fileNameBytes);

      const hasDataDescriptor = (generalFlag & 0x08) !== 0;

      const entry: CentralDirEntry = {
        fileName,
        compressionMethod,
        compressedSize,
        uncompressedSize,
        crc32: entryCrc32,
        localHeaderOffset,
        hasDataDescriptor,
      };

      this.centralDir.push(entry);
      this.entryMap.set(fileName, entry);

      pos += 46 + fileNameLength + extraFieldLength + commentLength;
    }
  }

  private async extractEntry(entry: CentralDirEntry): Promise<Uint8Array> {
    const pos = entry.localHeaderOffset;

    if (pos + 30 > this.data.length) {
      throw new ZipError("Local file header extends beyond file");
    }

    const sig = this.view.getUint32(pos, true);
    if (sig !== SIG_LOCAL_FILE) {
      throw new ZipError(`Invalid local file header signature at offset ${pos}`);
    }

    const fileNameLength = this.view.getUint16(pos + 26, true);
    const extraFieldLength = this.view.getUint16(pos + 28, true);
    const dataStart = pos + 30 + fileNameLength + extraFieldLength;

    // Use sizes from central directory (authoritative), not local header.
    // Local header may have zeros when data descriptor is used.
    let { compressedSize, uncompressedSize, crc32: expectedCrc } = entry;

    // Handle data descriptor case (Lotus/Excel): sizes might be zero in
    // both local header and central dir in malformed files.
    if (entry.hasDataDescriptor && compressedSize === 0) {
      // Try to read from data descriptor after compressed data.
      // This is a tricky case; we rely on central dir being authoritative.
      // If central dir also has zeros, we need to find the data descriptor.
      const localCompressedSize = this.view.getUint32(pos + 18, true);
      if (localCompressedSize > 0) {
        compressedSize = localCompressedSize;
      }
      const localUncompressedSize = this.view.getUint32(pos + 22, true);
      if (localUncompressedSize > 0) {
        uncompressedSize = localUncompressedSize;
      }

      // If still zero, try to find data descriptor
      if (compressedSize === 0) {
        const found = this.findDataDescriptor(dataStart);
        if (found) {
          expectedCrc = found.crc;
          compressedSize = found.compressedSize;
          uncompressedSize = found.uncompressedSize;
        }
      }
    }

    if (dataStart + compressedSize > this.data.length) {
      throw new ZipError(`Compressed data extends beyond file for entry: ${entry.fileName}`);
    }

    const compressedData = this.data.subarray(dataStart, dataStart + compressedSize);

    let result: Uint8Array;

    if (entry.compressionMethod === 0) {
      // STORE — no compression
      result = compressedData;
    } else if (entry.compressionMethod === 8) {
      // DEFLATE
      if (compressedSize === 0 && uncompressedSize === 0) {
        result = new Uint8Array(0);
      } else {
        result = await decompressDeflateRaw(compressedData);
      }
    } else {
      throw new ZipError(
        `Unsupported compression method ${entry.compressionMethod} for entry: ${entry.fileName}`,
      );
    }

    // Verify CRC-32 (skip if CRC is 0 — some generators omit it)
    if (expectedCrc !== 0 && result.length > 0) {
      const actualCrc = crc32(result);
      if (actualCrc !== expectedCrc) {
        throw new ZipError(
          `CRC-32 mismatch for ${entry.fileName}: expected 0x${expectedCrc.toString(16)}, got 0x${actualCrc.toString(16)}`,
        );
      }
    }

    return result;
  }

  private extractEntryStream(entry: CentralDirEntry): ReadableStream<Uint8Array> {
    const pos = entry.localHeaderOffset;

    if (pos + 30 > this.data.length) {
      throw new ZipError("Local file header extends beyond file");
    }

    const sig = this.view.getUint32(pos, true);
    if (sig !== SIG_LOCAL_FILE) {
      throw new ZipError(`Invalid local file header signature at offset ${pos}`);
    }

    const fileNameLength = this.view.getUint16(pos + 26, true);
    const extraFieldLength = this.view.getUint16(pos + 28, true);
    const dataStart = pos + 30 + fileNameLength + extraFieldLength;

    let { compressedSize } = entry;

    if (entry.hasDataDescriptor && compressedSize === 0) {
      const localCompressedSize = this.view.getUint32(pos + 18, true);
      if (localCompressedSize > 0) {
        compressedSize = localCompressedSize;
      }
      if (compressedSize === 0) {
        const found = this.findDataDescriptor(dataStart);
        if (found) {
          compressedSize = found.compressedSize;
        }
      }
    }

    if (dataStart + compressedSize > this.data.length) {
      throw new ZipError(`Compressed data extends beyond file for entry: ${entry.fileName}`);
    }

    const compressedData = this.data.subarray(dataStart, dataStart + compressedSize);

    if (entry.compressionMethod === 0) {
      // STORE — return raw data as a stream
      return new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(compressedData);
          controller.close();
        },
      });
    }

    if (entry.compressionMethod === 8) {
      // DEFLATE — stream through DecompressionStream if available
      if (compressedSize === 0 && entry.uncompressedSize === 0) {
        return new ReadableStream<Uint8Array>({
          start(controller) {
            controller.close();
          },
        });
      }

      if (checkDecompressionStream()) {
        const inputStream = new ReadableStream({
          start(controller) {
            controller.enqueue(compressedData);
            controller.close();
          },
        });
        return inputStream.pipeThrough(
          new DecompressionStream("deflate-raw"),
        ) as ReadableStream<Uint8Array>;
      }

      // Fallback: inflate synchronously and emit as stream
      const inflated = inflate(compressedData);
      return new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(inflated);
          controller.close();
        },
      });
    }

    throw new ZipError(
      `Unsupported compression method ${entry.compressionMethod} for entry: ${entry.fileName}`,
    );
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
        };
      }
    }
    return null;
  }
}

// ── ZIP Archive Writer ──────────────────────────────────────────────
// Creates ZIP files from entries (path → Uint8Array).
// Supports STORE (method 0) and DEFLATE (method 8).

import { crc32, deflate } from "./deflate";

// ── ZIP Signatures ──────────────────────────────────────────────────

const SIG_LOCAL_FILE = 0x04034b50;
const SIG_CENTRAL_DIR = 0x02014b50;
const SIG_END_OF_CENTRAL_DIR = 0x06054b50;

// ── Compression ─────────────────────────────────────────────────────

let hasCompressionStream: boolean | undefined;

function checkCompressionStream(): boolean {
  if (hasCompressionStream === undefined) {
    try {
      hasCompressionStream =
        typeof CompressionStream !== "undefined" &&
        typeof ReadableStream !== "undefined" &&
        typeof Response !== "undefined";
    } catch {
      hasCompressionStream = false;
    }
  }
  return hasCompressionStream;
}

async function compressDeflateRaw(data: Uint8Array): Promise<Uint8Array> {
  if (checkCompressionStream()) {
    try {
      const cs = new CompressionStream("deflate-raw");
      const writer = cs.writable.getWriter();
      const reader = cs.readable.getReader();

      writer.write(data as unknown as BufferSource);
      writer.close();

      const chunks: Uint8Array[] = [];
      let totalLen = 0;

      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(value);
        totalLen += value.length;
      }

      const result = new Uint8Array(totalLen);
      let offset = 0;
      for (const chunk of chunks) {
        result.set(chunk, offset);
        offset += chunk.length;
      }
      return result;
    } catch {
      // Fall through to pure TS
    }
  }

  return deflate(data);
}

// ── Types ───────────────────────────────────────────────────────────

interface PendingEntry {
  path: string;
  data: Uint8Array;
  compress: boolean;
}

// ── ZipWriter ───────────────────────────────────────────────────────

export class ZipWriter {
  private entries: PendingEntry[] = [];

  /** Add a file entry to the archive */
  add(path: string, data: Uint8Array, options?: { compress?: boolean }): void {
    const compress = options?.compress ?? true;
    this.entries.push({ path, data, compress });
  }

  /** Build the ZIP archive */
  async build(): Promise<Uint8Array> {
    // First pass: compress all entries
    const prepared: Array<{
      path: string;
      data: Uint8Array;
      compressedData: Uint8Array;
      method: number;
      entryCrc32: number;
    }> = [];

    for (const entry of this.entries) {
      const entryCrc32 = entry.data.length > 0 ? crc32(entry.data) : 0;
      let compressedData: Uint8Array;
      let method: number;

      if (entry.compress && entry.data.length > 0) {
        compressedData = await compressDeflateRaw(entry.data);
        method = 8; // DEFLATE

        // If compressed is larger, use STORE instead
        if (compressedData.length >= entry.data.length) {
          compressedData = entry.data;
          method = 0;
        }
      } else {
        compressedData = entry.data;
        method = 0; // STORE
      }

      prepared.push({
        path: entry.path,
        data: entry.data,
        compressedData,
        method,
        entryCrc32,
      });
    }

    // Calculate total size
    const encoder = new TextEncoder();
    const encodedPaths = prepared.map((e) => encoder.encode(e.path));

    // Local file headers + data
    let localSize = 0;
    for (let i = 0; i < prepared.length; i++) {
      localSize += 30 + encodedPaths[i].length + prepared[i].compressedData.length;
    }

    // Central directory
    let centralSize = 0;
    for (let i = 0; i < prepared.length; i++) {
      centralSize += 46 + encodedPaths[i].length;
    }

    // End of central directory
    const eocdSize = 22;

    const totalSize = localSize + centralSize + eocdSize;
    const output = new Uint8Array(totalSize);
    const view = new DataView(output.buffer);

    // Write local file headers + data
    let offset = 0;
    const localOffsets: number[] = [];

    for (let i = 0; i < prepared.length; i++) {
      const entry = prepared[i];
      const pathBytes = encodedPaths[i];

      localOffsets.push(offset);

      // Local file header
      view.setUint32(offset, SIG_LOCAL_FILE, true);
      view.setUint16(offset + 4, 20, true); // Version needed (2.0)
      view.setUint16(offset + 6, 0, true); // General purpose flag
      view.setUint16(offset + 8, entry.method, true); // Compression method
      view.setUint16(offset + 10, 0, true); // Mod time
      view.setUint16(offset + 12, 0x0021, true); // Mod date (1980-01-01)
      view.setUint32(offset + 14, entry.entryCrc32, true);
      view.setUint32(offset + 18, entry.compressedData.length, true);
      view.setUint32(offset + 22, entry.data.length, true);
      view.setUint16(offset + 26, pathBytes.length, true);
      view.setUint16(offset + 28, 0, true); // Extra field length

      output.set(pathBytes, offset + 30);
      offset += 30 + pathBytes.length;

      output.set(entry.compressedData, offset);
      offset += entry.compressedData.length;
    }

    // Write central directory
    const centralDirOffset = offset;

    for (let i = 0; i < prepared.length; i++) {
      const entry = prepared[i];
      const pathBytes = encodedPaths[i];

      view.setUint32(offset, SIG_CENTRAL_DIR, true);
      view.setUint16(offset + 4, 20, true); // Version made by (2.0)
      view.setUint16(offset + 6, 20, true); // Version needed (2.0)
      view.setUint16(offset + 8, 0, true); // General purpose flag
      view.setUint16(offset + 10, entry.method, true); // Compression method
      view.setUint16(offset + 12, 0, true); // Mod time
      view.setUint16(offset + 14, 0x0021, true); // Mod date (1980-01-01)
      view.setUint32(offset + 16, entry.entryCrc32, true);
      view.setUint32(offset + 20, entry.compressedData.length, true);
      view.setUint32(offset + 24, entry.data.length, true);
      view.setUint16(offset + 28, pathBytes.length, true);
      view.setUint16(offset + 30, 0, true); // Extra field length
      view.setUint16(offset + 32, 0, true); // File comment length
      view.setUint16(offset + 34, 0, true); // Disk number start
      view.setUint16(offset + 36, 0, true); // Internal file attributes
      view.setUint32(offset + 38, 0, true); // External file attributes
      view.setUint32(offset + 42, localOffsets[i], true);

      output.set(pathBytes, offset + 46);
      offset += 46 + pathBytes.length;
    }

    const centralDirSize = offset - centralDirOffset;

    // Write End of Central Directory
    view.setUint32(offset, SIG_END_OF_CENTRAL_DIR, true);
    view.setUint16(offset + 4, 0, true); // Disk number
    view.setUint16(offset + 6, 0, true); // Disk with central dir
    view.setUint16(offset + 8, prepared.length, true); // Entries on this disk
    view.setUint16(offset + 10, prepared.length, true); // Total entries
    view.setUint32(offset + 12, centralDirSize, true);
    view.setUint32(offset + 16, centralDirOffset, true);
    view.setUint16(offset + 20, 0, true); // Comment length

    return output;
  }
}

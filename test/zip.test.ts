import { describe, expect, it } from "vitest";
import { crc32, deflate, inflate } from "../src/zip/deflate";
import { ZipReader } from "../src/zip/reader";
import { ZipWriter } from "../src/zip/writer";

// ── Helpers ─────────────────────────────────────────────────────────

function textToBytes(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

function bytesToText(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes);
}

/** Build a minimal valid ZIP with a single STORE entry manually */
function buildManualZip(fileName: string, content: Uint8Array): Uint8Array {
  const enc = new TextEncoder();
  const nameBytes = enc.encode(fileName);
  const fileCrc = crc32(content);

  // Local file header (30 + nameLen + data)
  const localHeaderSize = 30 + nameBytes.length + content.length;
  // Central dir entry (46 + nameLen)
  const centralEntrySize = 46 + nameBytes.length;
  // EOCD (22)
  const totalSize = localHeaderSize + centralEntrySize + 22;

  const buf = new Uint8Array(totalSize);
  const view = new DataView(buf.buffer);
  let pos = 0;

  // Local file header
  view.setUint32(pos, 0x04034b50, true);
  pos += 4;
  view.setUint16(pos, 20, true);
  pos += 2; // version needed
  view.setUint16(pos, 0, true);
  pos += 2; // flags
  view.setUint16(pos, 0, true);
  pos += 2; // compression: STORE
  view.setUint16(pos, 0, true);
  pos += 2; // mod time
  view.setUint16(pos, 0x0021, true);
  pos += 2; // mod date
  view.setUint32(pos, fileCrc, true);
  pos += 4; // crc32
  view.setUint32(pos, content.length, true);
  pos += 4; // compressed size
  view.setUint32(pos, content.length, true);
  pos += 4; // uncompressed size
  view.setUint16(pos, nameBytes.length, true);
  pos += 2; // name len
  view.setUint16(pos, 0, true);
  pos += 2; // extra field len
  buf.set(nameBytes, pos);
  pos += nameBytes.length;
  buf.set(content, pos);
  pos += content.length;

  const centralDirOffset = pos;

  // Central directory entry
  view.setUint32(pos, 0x02014b50, true);
  pos += 4;
  view.setUint16(pos, 20, true);
  pos += 2; // version made by
  view.setUint16(pos, 20, true);
  pos += 2; // version needed
  view.setUint16(pos, 0, true);
  pos += 2; // flags
  view.setUint16(pos, 0, true);
  pos += 2; // compression: STORE
  view.setUint16(pos, 0, true);
  pos += 2; // mod time
  view.setUint16(pos, 0x0021, true);
  pos += 2; // mod date
  view.setUint32(pos, fileCrc, true);
  pos += 4;
  view.setUint32(pos, content.length, true);
  pos += 4;
  view.setUint32(pos, content.length, true);
  pos += 4;
  view.setUint16(pos, nameBytes.length, true);
  pos += 2;
  view.setUint16(pos, 0, true);
  pos += 2; // extra len
  view.setUint16(pos, 0, true);
  pos += 2; // comment len
  view.setUint16(pos, 0, true);
  pos += 2; // disk start
  view.setUint16(pos, 0, true);
  pos += 2; // internal attrs
  view.setUint32(pos, 0, true);
  pos += 4; // external attrs
  view.setUint32(pos, 0, true);
  pos += 4; // local header offset

  buf.set(nameBytes, pos);
  pos += nameBytes.length;

  const centralDirSize = pos - centralDirOffset;

  // End of central directory
  view.setUint32(pos, 0x06054b50, true);
  pos += 4;
  view.setUint16(pos, 0, true);
  pos += 2; // disk num
  view.setUint16(pos, 0, true);
  pos += 2; // disk with CD
  view.setUint16(pos, 1, true);
  pos += 2; // entries on disk
  view.setUint16(pos, 1, true);
  pos += 2; // total entries
  view.setUint32(pos, centralDirSize, true);
  pos += 4;
  view.setUint32(pos, centralDirOffset, true);
  pos += 4;
  view.setUint16(pos, 0, true);
  pos += 2; // comment len

  return buf;
}

// ── CRC-32 Tests ────────────────────────────────────────────────────

describe("crc32", () => {
  it("computes CRC-32 of empty data", () => {
    expect(crc32(new Uint8Array(0))).toBe(0x00000000);
  });

  it("computes CRC-32 of known string", () => {
    const data = textToBytes("123456789");
    // Known CRC-32 of "123456789" is 0xCBF43926
    expect(crc32(data)).toBe(0xcbf43926);
  });

  it("computes CRC-32 of single byte", () => {
    const data = new Uint8Array([0x00]);
    expect(crc32(data)).toBe(0xd202ef8d);
  });

  it("handles binary data", () => {
    const data = new Uint8Array([0xff, 0xfe, 0xfd, 0xfc]);
    const result = crc32(data);
    expect(result).toBeTypeOf("number");
    expect(result >>> 0).toBe(result); // Should be unsigned 32-bit
  });
});

// ── Pure TS Inflate/Deflate Tests ───────────────────────────────────

describe("inflate/deflate (pure TS)", () => {
  it("round-trips empty data", () => {
    const input = new Uint8Array(0);
    const compressed = deflate(input);
    const decompressed = inflate(compressed);
    expect(decompressed.length).toBe(0);
  });

  it("round-trips short text", () => {
    const input = textToBytes("Hello, World!");
    const compressed = deflate(input);
    const decompressed = inflate(compressed);
    expect(bytesToText(decompressed)).toBe("Hello, World!");
  });

  it("round-trips repeated data (benefits from LZ77)", () => {
    const input = textToBytes("ABCABC".repeat(100));
    const compressed = deflate(input);
    const decompressed = inflate(compressed);
    expect(bytesToText(decompressed)).toBe("ABCABC".repeat(100));
    // Compressed should be significantly smaller
    expect(compressed.length).toBeLessThan(input.length);
  });

  it("round-trips binary data", () => {
    const input = new Uint8Array(256);
    for (let i = 0; i < 256; i++) input[i] = i;
    const compressed = deflate(input);
    const decompressed = inflate(compressed);
    expect(decompressed).toEqual(input);
  });

  it("round-trips 100KB of data", () => {
    const size = 100 * 1024;
    const input = new Uint8Array(size);
    // Fill with semi-random but compressible data
    for (let i = 0; i < size; i++) {
      input[i] = (i * 7 + 13) & 0xff;
    }
    const compressed = deflate(input);
    const decompressed = inflate(compressed);
    expect(decompressed.length).toBe(input.length);
    expect(decompressed).toEqual(input);
  });

  it("round-trips highly compressible data", () => {
    // All zeros
    const input = new Uint8Array(10000);
    const compressed = deflate(input);
    const decompressed = inflate(compressed);
    expect(decompressed).toEqual(input);
    // Should compress very well
    expect(compressed.length).toBeLessThan(input.length / 10);
  });

  it("handles single byte", () => {
    const input = new Uint8Array([42]);
    const compressed = deflate(input);
    const decompressed = inflate(compressed);
    expect(decompressed).toEqual(input);
  });

  it("handles data with all byte values", () => {
    const input = new Uint8Array(512);
    for (let i = 0; i < 512; i++) input[i] = i & 0xff;
    const compressed = deflate(input);
    const decompressed = inflate(compressed);
    expect(decompressed).toEqual(input);
  });
});

// ── ZipReader Tests (hand-built ZIP) ────────────────────────────────

describe("ZipReader", () => {
  it("reads a manually-built ZIP with one STORE entry", () => {
    const content = textToBytes("Hello ZIP!");
    const zip = buildManualZip("test.txt", content);
    const reader = new ZipReader(zip);

    expect(reader.entries()).toEqual(["test.txt"]);
    expect(reader.has("test.txt")).toBe(true);
    expect(reader.has("other.txt")).toBe(false);
  });

  it("extracts a STORE entry correctly", async () => {
    const content = textToBytes("Hello ZIP!");
    const zip = buildManualZip("test.txt", content);
    const reader = new ZipReader(zip);

    const extracted = await reader.extract("test.txt");
    expect(bytesToText(extracted)).toBe("Hello ZIP!");
  });

  it("extracts all files", async () => {
    const content = textToBytes("data");
    const zip = buildManualZip("file.txt", content);
    const reader = new ZipReader(zip);

    const all = await reader.extractAll();
    expect(all.size).toBe(1);
    expect(bytesToText(all.get("file.txt")!)).toBe("data");
  });

  it("throws on invalid data", () => {
    expect(() => new ZipReader(new Uint8Array(10))).toThrow();
  });

  it("throws on non-ZIP data", () => {
    const data = new Uint8Array(100);
    data[0] = 0x50;
    data[1] = 0x44;
    data[2] = 0x46;
    expect(() => new ZipReader(data)).toThrow("End of Central Directory");
  });

  it("throws when extracting non-existent entry", async () => {
    const zip = buildManualZip("a.txt", textToBytes("a"));
    const reader = new ZipReader(zip);

    await expect(reader.extract("nope.txt")).rejects.toThrow("Entry not found");
  });
});

// ── ZipWriter Tests ─────────────────────────────────────────────────

describe("ZipWriter", () => {
  it("creates an empty archive", async () => {
    const writer = new ZipWriter();
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    expect(reader.entries()).toEqual([]);
  });

  it("creates an archive with one STORE entry", async () => {
    const writer = new ZipWriter();
    writer.add("hello.txt", textToBytes("Hello!"), { compress: false });
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    expect(reader.entries()).toEqual(["hello.txt"]);
    const content = await reader.extract("hello.txt");
    expect(bytesToText(content)).toBe("Hello!");
  });

  it("creates an archive with one DEFLATE entry", async () => {
    const writer = new ZipWriter();
    const text = "Compressible data! ".repeat(50);
    writer.add("data.txt", textToBytes(text), { compress: true });
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    const content = await reader.extract("data.txt");
    expect(bytesToText(content)).toBe(text);
  });

  it("creates an archive with multiple files", async () => {
    const writer = new ZipWriter();
    writer.add("a.txt", textToBytes("File A"));
    writer.add("b.txt", textToBytes("File B"));
    writer.add("c.txt", textToBytes("File C"));
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    expect(reader.entries()).toEqual(["a.txt", "b.txt", "c.txt"]);

    expect(bytesToText(await reader.extract("a.txt"))).toBe("File A");
    expect(bytesToText(await reader.extract("b.txt"))).toBe("File B");
    expect(bytesToText(await reader.extract("c.txt"))).toBe("File C");
  });

  it("handles empty file content", async () => {
    const writer = new ZipWriter();
    writer.add("empty.txt", new Uint8Array(0));
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    const content = await reader.extract("empty.txt");
    expect(content.length).toBe(0);
  });

  it("defaults to compress: true", async () => {
    const writer = new ZipWriter();
    const text = "Repeated data for compression. ".repeat(100);
    writer.add("file.txt", textToBytes(text));
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    const content = await reader.extract("file.txt");
    expect(bytesToText(content)).toBe(text);

    // The archive should be smaller than uncompressed data
    expect(zip.length).toBeLessThan(text.length);
  });
});

// ── Round-Trip Tests ────────────────────────────────────────────────

describe("ZIP round-trip", () => {
  it("round-trips a single file", async () => {
    const original = textToBytes("Round trip test content");

    const writer = new ZipWriter();
    writer.add("test.txt", original);
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    const extracted = await reader.extract("test.txt");
    expect(extracted).toEqual(original);
  });

  it("round-trips multiple files with mixed compression", async () => {
    const files = new Map<string, Uint8Array>();
    files.set("stored.txt", textToBytes("This is stored"));
    files.set("compressed.txt", textToBytes("This is compressed ".repeat(100)));
    files.set("binary.bin", new Uint8Array([0, 1, 2, 3, 255, 254, 253]));

    const writer = new ZipWriter();
    writer.add("stored.txt", files.get("stored.txt")!, { compress: false });
    writer.add("compressed.txt", files.get("compressed.txt")!, {
      compress: true,
    });
    writer.add("binary.bin", files.get("binary.bin")!, { compress: false });
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    const all = await reader.extractAll();

    expect(all.size).toBe(3);
    for (const [path, data] of files) {
      expect(all.get(path)).toEqual(data);
    }
  });

  it("round-trips files with subdirectories", async () => {
    const writer = new ZipWriter();
    writer.add("root.txt", textToBytes("root"));
    writer.add("dir/file.txt", textToBytes("in dir"));
    writer.add("dir/sub/deep.txt", textToBytes("deep nested"));
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    expect(reader.has("root.txt")).toBe(true);
    expect(reader.has("dir/file.txt")).toBe(true);
    expect(reader.has("dir/sub/deep.txt")).toBe(true);

    expect(bytesToText(await reader.extract("dir/sub/deep.txt"))).toBe("deep nested");
  });

  it("round-trips Unicode filenames", async () => {
    const writer = new ZipWriter();
    writer.add("日本語.txt", textToBytes("Japanese"));
    writer.add("données.txt", textToBytes("French"));
    writer.add("Ünïcödé/fïlé.txt", textToBytes("nested unicode"));
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    expect(reader.has("日本語.txt")).toBe(true);
    expect(reader.has("données.txt")).toBe(true);
    expect(reader.has("Ünïcödé/fïlé.txt")).toBe(true);

    expect(bytesToText(await reader.extract("日本語.txt"))).toBe("Japanese");
    expect(bytesToText(await reader.extract("données.txt"))).toBe("French");
  });

  it("round-trips large file (100KB+)", async () => {
    const size = 150 * 1024;
    const data = new Uint8Array(size);
    // Generate semi-random but compressible content
    for (let i = 0; i < size; i++) {
      data[i] = ((i * 31 + 17) % 128) & 0xff;
    }

    const writer = new ZipWriter();
    writer.add("large.bin", data, { compress: true });
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    const extracted = await reader.extract("large.bin");
    expect(extracted.length).toBe(data.length);
    expect(extracted).toEqual(data);
  });

  it("round-trips many small files", async () => {
    const writer = new ZipWriter();
    const fileCount = 50;

    for (let i = 0; i < fileCount; i++) {
      writer.add(`file-${i}.txt`, textToBytes(`Content of file ${i}`));
    }
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    expect(reader.entries().length).toBe(fileCount);

    for (let i = 0; i < fileCount; i++) {
      const content = await reader.extract(`file-${i}.txt`);
      expect(bytesToText(content)).toBe(`Content of file ${i}`);
    }
  });

  it("round-trips XML content (XLSX-like)", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

    const writer = new ZipWriter();
    writer.add("[Content_Types].xml", textToBytes("<Types/>"));
    writer.add("_rels/.rels", textToBytes("<Relationships/>"));
    writer.add("xl/workbook.xml", textToBytes(xml));
    writer.add("xl/worksheets/sheet1.xml", textToBytes("<worksheet><sheetData/></worksheet>"));
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    const all = await reader.extractAll();
    expect(all.size).toBe(4);
    expect(bytesToText(all.get("xl/workbook.xml")!)).toBe(xml);
  });

  it("handles files that do not compress well", async () => {
    // Random-like data that won't compress
    const data = new Uint8Array(1024);
    for (let i = 0; i < data.length; i++) {
      data[i] = (Math.sin(i * 0.1) * 128 + 128 + ((i * 7919) % 256)) & 0xff;
    }

    const writer = new ZipWriter();
    writer.add("random.bin", data, { compress: true });
    const zip = await writer.build();

    // Writer should fall back to STORE if compression doesn't help
    const reader = new ZipReader(zip);
    const extracted = await reader.extract("random.bin");
    expect(extracted).toEqual(data);
  });

  it("correctly verifies CRC-32 on extraction", async () => {
    const content = textToBytes("CRC test");
    const zip = buildManualZip("test.txt", content);

    // Corrupt one byte of the content in the ZIP
    const corruptZip = new Uint8Array(zip);
    // Content starts after local header (30 + filename length)
    const contentOffset = 30 + "test.txt".length;
    corruptZip[contentOffset] ^= 0xff;

    const reader = new ZipReader(corruptZip);
    await expect(reader.extract("test.txt")).rejects.toThrow("CRC-32 mismatch");
  });

  it("round-trips file with all byte values 0-255", async () => {
    const data = new Uint8Array(256);
    for (let i = 0; i < 256; i++) data[i] = i;

    const writer = new ZipWriter();
    writer.add("bytes.bin", data, { compress: true });
    const zip = await writer.build();

    const reader = new ZipReader(zip);
    const extracted = await reader.extract("bytes.bin");
    expect(extracted).toEqual(data);
  });
});

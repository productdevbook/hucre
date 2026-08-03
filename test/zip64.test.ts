import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { ZipReader } from "../src/zip/reader"
import { zipStreamChunks } from "../src/zip/stream-writer"
import { writeXlsxStream } from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipError } from "../src/errors"
import { crc32 } from "../src/zip/deflate"
import type { CellValue } from "../src/_types"

const enc = new TextEncoder()
const dec = new TextDecoder()

// ── Helpers ──────────────────────────────────────────────────────────

function concat(chunks: Uint8Array[]): Uint8Array {
  const total = chunks.reduce((n, c) => n + c.length, 0)
  const out = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    out.set(chunk, offset)
    offset += chunk.length
  }
  return out
}

async function buildStreamZip(
  entries: Array<{ path: string; data: Uint8Array; compress?: boolean }>,
  options?: { zip64?: boolean },
): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  for await (const chunk of zipStreamChunks(entries, options)) chunks.push(chunk)
  return concat(chunks)
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
  }
  return concat(chunks)
}

/** Locate the End-Of-Central-Directory record (signature 0x06054b50). */
function findEocd(buf: Uint8Array): number {
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  for (let i = buf.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === 0x06054b50) return i
  }
  throw new Error("EOCD not found")
}

function viewOf(buf: Uint8Array): DataView {
  return new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
}

/**
 * Hand-build a single-entry STORE archive, escaping exactly the chosen
 * fields into the ZIP64 extra field.
 *
 * Real producers escape only what actually overflows — an archive past
 * 4 GiB made of small entries escapes the offset but not the sizes. The
 * extra field is positional, so each combination exercises a different
 * read path than the writer's always-escape-everything output.
 */
function buildPartialZip64(
  name: string,
  payload: Uint8Array,
  escape: { size?: boolean; offset?: boolean },
): Uint8Array {
  const SENTINEL = 0xffffffff
  const nameBytes = enc.encode(name)
  const crc = crc32(payload)

  const escapeSize = escape.size ?? false
  const escapeOffset = escape.offset ?? false
  // Order is fixed: uncompressed, compressed, then local header offset.
  const extraFields = (escapeSize ? 2 : 0) + (escapeOffset ? 1 : 0)
  const extraSize = extraFields > 0 ? 4 + extraFields * 8 : 0

  const localSize = 30 + nameBytes.length
  const centralSize = 46 + nameBytes.length + extraSize
  const buf = new Uint8Array(localSize + payload.length + centralSize + 56 + 20 + 22)
  const view = viewOf(buf)

  // Local file header — sizes always real here; only the directory is escaped.
  view.setUint32(0, 0x04034b50, true)
  view.setUint16(4, 45, true)
  view.setUint16(6, 0, true)
  view.setUint16(8, 0, true) // STORE
  view.setUint16(12, 0x0021, true)
  view.setUint32(14, crc, true)
  view.setUint32(18, payload.length, true)
  view.setUint32(22, payload.length, true)
  view.setUint16(26, nameBytes.length, true)
  buf.set(nameBytes, 30)
  buf.set(payload, localSize)

  // Central directory
  let pos = localSize + payload.length
  const centralDirOffset = pos
  view.setUint32(pos, 0x02014b50, true)
  view.setUint16(pos + 4, 45, true)
  view.setUint16(pos + 6, 45, true)
  view.setUint16(pos + 10, 0, true) // STORE
  view.setUint16(pos + 14, 0x0021, true)
  view.setUint32(pos + 16, crc, true)
  view.setUint32(pos + 20, escapeSize ? SENTINEL : payload.length, true)
  view.setUint32(pos + 24, escapeSize ? SENTINEL : payload.length, true)
  view.setUint16(pos + 28, nameBytes.length, true)
  view.setUint16(pos + 30, extraSize, true)
  view.setUint32(pos + 42, escapeOffset ? SENTINEL : 0, true)
  buf.set(nameBytes, pos + 46)

  if (extraSize > 0) {
    let extra = pos + 46 + nameBytes.length
    view.setUint16(extra, 0x0001, true)
    view.setUint16(extra + 2, extraFields * 8, true)
    extra += 4
    if (escapeSize) {
      view.setBigUint64(extra, BigInt(payload.length), true)
      view.setBigUint64(extra + 8, BigInt(payload.length), true)
      extra += 16
    }
    if (escapeOffset) view.setBigUint64(extra, 0n, true)
  }

  pos += centralSize
  const centralDirSize = pos - centralDirOffset
  const zip64EocdOffset = pos

  view.setUint32(pos, 0x06064b50, true)
  view.setBigUint64(pos + 4, 44n, true)
  view.setUint16(pos + 12, 45, true)
  view.setUint16(pos + 14, 45, true)
  view.setBigUint64(pos + 24, 1n, true)
  view.setBigUint64(pos + 32, 1n, true)
  view.setBigUint64(pos + 40, BigInt(centralDirSize), true)
  view.setBigUint64(pos + 48, BigInt(centralDirOffset), true)
  pos += 56

  view.setUint32(pos, 0x07064b50, true)
  view.setBigUint64(pos + 8, BigInt(zip64EocdOffset), true)
  view.setUint32(pos + 16, 1, true)
  pos += 20

  view.setUint32(pos, 0x06054b50, true)
  view.setUint16(pos + 8, 0xffff, true)
  view.setUint16(pos + 10, 0xffff, true)
  view.setUint32(pos + 12, SENTINEL, true)
  view.setUint32(pos + 16, SENTINEL, true)

  return buf
}

// ═══════════════════════════════════════════════════════════════════════
// Writing ZIP64
// ═══════════════════════════════════════════════════════════════════════

describe("zipStreamChunks — ZIP64 output", () => {
  it("emits a ZIP64 EOCD record and locator", async () => {
    const buf = await buildStreamZip([{ path: "a.txt", data: enc.encode("hello") }], {
      zip64: true,
    })
    const view = viewOf(buf)
    const eocd = findEocd(buf)

    // Locator sits immediately before the classic EOCD...
    expect(view.getUint32(eocd - 20, true)).toBe(0x07064b50)
    // ...and points at the ZIP64 EOCD record.
    const recordOffset = Number(view.getBigUint64(eocd - 20 + 8, true))
    expect(view.getUint32(recordOffset, true)).toBe(0x06064b50)
    expect(Number(view.getBigUint64(recordOffset + 32, true))).toBe(1) // total entries
  })

  it("escapes the classic EOCD fields to sentinels", async () => {
    const buf = await buildStreamZip([{ path: "a.txt", data: enc.encode("hello") }], {
      zip64: true,
    })
    const view = viewOf(buf)
    const eocd = findEocd(buf)

    expect(view.getUint16(eocd + 8, true)).toBe(0xffff)
    expect(view.getUint16(eocd + 10, true)).toBe(0xffff)
    expect(view.getUint32(eocd + 12, true)).toBe(0xffffffff)
    expect(view.getUint32(eocd + 16, true)).toBe(0xffffffff)
  })

  it("writes a 24-byte data descriptor and a local ZIP64 extra field", async () => {
    const chunks: Uint8Array[] = []
    for await (const chunk of zipStreamChunks([{ path: "a.txt", data: enc.encode("x") }], {
      zip64: true,
    })) {
      chunks.push(chunk)
    }

    const header = chunks[0]!
    const hv = viewOf(header)
    expect(hv.getUint16(4, true)).toBe(45) // Version needed — 4.5
    const nameLen = hv.getUint16(26, true)
    expect(hv.getUint16(28, true)).toBe(20) // Extra field length
    expect(hv.getUint16(30 + nameLen, true)).toBe(0x0001) // ZIP64 extra id

    // Descriptor is the chunk right before the central directory.
    const descriptor = chunks[chunks.length - 2]!
    expect(descriptor.length).toBe(24)
    expect(viewOf(descriptor).getUint32(0, true)).toBe(0x08074b50)
  })

  it("stays classic (16-byte descriptor, no locator) by default", async () => {
    const buf = await buildStreamZip([{ path: "a.txt", data: enc.encode("hello") }])
    const view = viewOf(buf)
    const eocd = findEocd(buf)

    expect(view.getUint16(eocd + 8, true)).toBe(1)
    expect(view.getUint32(eocd + 12, true)).not.toBe(0xffffffff)
    expect(view.getUint16(4, true)).toBe(20) // Local header version needed — 2.0
    expect(view.getUint16(28, true)).toBe(0) // No extra field
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Reading ZIP64
// ═══════════════════════════════════════════════════════════════════════

describe("ZipReader — ZIP64 input", () => {
  it("round-trips a ZIP64 archive written by the stream writer", async () => {
    const buf = await buildStreamZip(
      [
        { path: "a.txt", data: enc.encode("hello hello hello") },
        { path: "nested/b.bin", data: new Uint8Array([1, 2, 3, 4, 5]), compress: false },
        { path: "c.txt", data: enc.encode("") },
      ],
      { zip64: true },
    )

    const zip = new ZipReader(buf)
    expect(zip.entries().sort()).toEqual(["a.txt", "c.txt", "nested/b.bin"])
    expect(dec.decode(await zip.extract("a.txt"))).toBe("hello hello hello")
    expect(Array.from(await zip.extract("nested/b.bin"))).toEqual([1, 2, 3, 4, 5])
    expect(await zip.extract("c.txt")).toHaveLength(0)
  })

  it("resolves per-entry sizes and offsets from the ZIP64 extra field", async () => {
    // Every entry in a ZIP64 archive we write carries sentinels in the
    // fixed record, so a correct read has to come from the extra field.
    const payload = enc.encode("x".repeat(5000))
    const buf = await buildStreamZip(
      [
        { path: "first.txt", data: payload },
        { path: "second.txt", data: enc.encode("tail") },
      ],
      { zip64: true },
    )

    const view = viewOf(buf)
    // Second entry's fixed central-directory offset field must be escaped —
    // otherwise this test would pass without any extra-field parsing.
    let pos = -1
    for (let i = 0; i < buf.length - 4; i++) {
      if (view.getUint32(i, true) === 0x02014b50) {
        pos = i
        break
      }
    }
    expect(pos).toBeGreaterThan(-1)
    expect(view.getUint32(pos + 42, true)).toBe(0xffffffff)

    const zip = new ZipReader(buf)
    expect(dec.decode(await zip.extract("first.txt"))).toHaveLength(5000)
    expect(dec.decode(await zip.extract("second.txt"))).toBe("tail")
  })

  it("reads a ZIP64 XLSX end to end", async () => {
    const rows: CellValue[][] = [
      [1, "Alice", true],
      [2, "Bob", false],
    ]
    const bytes = await collect(
      writeXlsxStream(
        {
          name: "Data",
          columns: [{ header: "ID" }, { header: "Name" }, { header: "Flag" }],
          zip64: true,
        },
        rows,
      ),
    )

    const workbook = await readXlsx(bytes)
    expect(workbook.sheets[0].name).toBe("Data")
    expect(workbook.sheets[0].rows).toEqual([["ID", "Name", "Flag"], ...rows])
  })

  it.each([
    ["only the local header offset escaped", { offset: true }],
    ["only the sizes escaped", { size: true }],
    ["sizes and offset escaped", { size: true, offset: true }],
    ["nothing escaped but ZIP64 records present", {}],
  ])("reads an archive with %s", async (_label, escape) => {
    const payload = enc.encode("partial zip64 payload")
    const zip = new ZipReader(buildPartialZip64("a.txt", payload, escape))
    expect(dec.decode(await zip.extract("a.txt"))).toBe("partial zip64 payload")
  })

  it("still reads a classic archive", async () => {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("hello"))
    const reader = new ZipReader(await zip.build())
    expect(dec.decode(await reader.extract("a.txt"))).toBe("hello")
  })

  it("rejects a ZIP64 escape with no locator behind it", async () => {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("hello"))
    const buf = await zip.build()

    // Forge a ZIP64 escape on an archive that has no ZIP64 records.
    const eocd = findEocd(buf)
    const view = viewOf(buf)
    view.setUint16(eocd + 8, 0xffff, true)
    view.setUint16(eocd + 10, 0xffff, true)

    expect(() => new ZipReader(buf)).toThrow(ZipError)
    expect(() => new ZipReader(buf)).toThrow(/ZIP64 End of Central Directory locator/)
  })

  it("rejects an entry whose ZIP64 extra field is missing", async () => {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("hello"))
    const buf = await zip.build()

    // Escape the compressed size without supplying an extra field.
    const view = viewOf(buf)
    let pos = -1
    for (let i = 0; i < buf.length - 4; i++) {
      if (view.getUint32(i, true) === 0x02014b50) {
        pos = i
        break
      }
    }
    view.setUint32(pos + 20, 0xffffffff, true)

    expect(() => new ZipReader(buf)).toThrow(/Missing ZIP64 extra field/)
  })
})

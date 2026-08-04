import { describe, expect, it } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { ZipStreamReader } from "../src/zip/stream-reader"
import { ZipWriter } from "../src/zip/writer"
import { crc32, deflate, inflate } from "../src/zip/deflate"
import { ParseError, ZipError } from "../src/errors"

const enc = new TextEncoder()
const dec = new TextDecoder()

const SIG_LOCAL = 0x04034b50
const SIG_CENTRAL = 0x02014b50
const SIG_EOCD = 0x06054b50
const SIG_DESCRIPTOR = 0x08074b50
const SENTINEL = 0xffffffff

// ── Archive builder ──────────────────────────────────────────────────
// Deliberately lets each field be set independently of the bytes it is
// supposed to describe: nearly every branch below is a reader defence
// against a header that disagrees with reality (truncation, a lying
// producer, a ZIP64 escape without the matching extra field).

interface EntrySpec {
  name: string
  /** Bytes stored in the archive body — already compressed when `method` is 8. */
  body: Uint8Array
  method?: number
  /** General-purpose bit flags; 0x08 marks a trailing data descriptor. */
  flags?: number
  crc?: number
  /** Sizes recorded in the central directory (default: the real ones). */
  centralSizes?: { compressed: number; uncompressed: number }
  /** Sizes recorded in the local file header (default: the real ones). */
  localSizes?: { compressed: number; uncompressed: number }
  /** Local header offset recorded in the central directory. */
  centralOffset?: number
  /** Classic 16-byte data descriptor appended after the body. */
  descriptor?: { crc: number; compressed: number; uncompressed: number }
  /** Raw extra-field bytes appended to the central-directory record. */
  centralExtra?: Uint8Array
}

function viewOf(buf: Uint8Array): DataView {
  return new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
}

function buildZip(specs: EntrySpec[]): Uint8Array {
  const parts: Uint8Array[] = []
  const offsets: number[] = []
  let offset = 0

  for (const s of specs) {
    const nameBytes = enc.encode(s.name)
    const crc = s.crc ?? ((s.method ?? 0) === 0 ? crc32(s.body) : 0)
    const local = new Uint8Array(30 + nameBytes.length + s.body.length + (s.descriptor ? 16 : 0))
    const dv = viewOf(local)
    dv.setUint32(0, SIG_LOCAL, true)
    dv.setUint16(4, 20, true)
    dv.setUint16(6, s.flags ?? 0, true)
    dv.setUint16(8, s.method ?? 0, true)
    dv.setUint16(12, 0x0021, true)
    dv.setUint32(14, crc, true)
    dv.setUint32(18, s.localSizes ? s.localSizes.compressed : s.body.length, true)
    dv.setUint32(22, s.localSizes ? s.localSizes.uncompressed : s.body.length, true)
    dv.setUint16(26, nameBytes.length, true)
    local.set(nameBytes, 30)
    local.set(s.body, 30 + nameBytes.length)
    if (s.descriptor) {
      const at = 30 + nameBytes.length + s.body.length
      dv.setUint32(at, SIG_DESCRIPTOR, true)
      dv.setUint32(at + 4, s.descriptor.crc, true)
      dv.setUint32(at + 8, s.descriptor.compressed, true)
      dv.setUint32(at + 12, s.descriptor.uncompressed, true)
    }
    offsets.push(offset)
    offset += local.length
    parts.push(local)
  }

  const centralDirOffset = offset
  specs.forEach((s, i) => {
    const nameBytes = enc.encode(s.name)
    const extra = s.centralExtra ?? new Uint8Array(0)
    const crc = s.crc ?? ((s.method ?? 0) === 0 ? crc32(s.body) : 0)
    const central = new Uint8Array(46 + nameBytes.length + extra.length)
    const dv = viewOf(central)
    dv.setUint32(0, SIG_CENTRAL, true)
    dv.setUint16(4, 20, true)
    dv.setUint16(6, 20, true)
    dv.setUint16(8, s.flags ?? 0, true)
    dv.setUint16(10, s.method ?? 0, true)
    dv.setUint16(14, 0x0021, true)
    dv.setUint32(16, crc, true)
    dv.setUint32(20, s.centralSizes ? s.centralSizes.compressed : s.body.length, true)
    dv.setUint32(24, s.centralSizes ? s.centralSizes.uncompressed : s.body.length, true)
    dv.setUint16(28, nameBytes.length, true)
    dv.setUint16(30, extra.length, true)
    dv.setUint32(42, s.centralOffset ?? offsets[i], true)
    central.set(nameBytes, 46)
    central.set(extra, 46 + nameBytes.length)
    offset += central.length
    parts.push(central)
  })

  const eocd = new Uint8Array(22)
  const ev = viewOf(eocd)
  ev.setUint32(0, SIG_EOCD, true)
  ev.setUint16(8, specs.length, true)
  ev.setUint16(10, specs.length, true)
  ev.setUint32(12, offset - centralDirOffset, true)
  ev.setUint32(16, centralDirOffset, true)
  parts.push(eocd)

  const total = parts.reduce((n, p) => n + p.length, 0)
  const out = new Uint8Array(total)
  let pos = 0
  for (const p of parts) {
    out.set(p, pos)
    pos += p.length
  }
  return out
}

/** A ZIP64 extended-information extra field with an explicitly chosen size header. */
function zip64Extra(values: bigint[], opts?: { declaredSize?: number }): Uint8Array {
  const out = new Uint8Array(4 + values.length * 8)
  const dv = viewOf(out)
  dv.setUint16(0, 0x0001, true)
  dv.setUint16(2, opts?.declaredSize ?? values.length * 8, true)
  values.forEach((v, i) => dv.setBigUint64(4 + i * 8, v, true))
  return out
}

/** An extra field the reader must skip over: 0x5455 "extended timestamp". */
function timestampExtra(): Uint8Array {
  const out = new Uint8Array(4 + 5)
  const dv = viewOf(out)
  dv.setUint16(0, 0x5455, true)
  dv.setUint16(2, 5, true)
  return out
}

function findEocd(buf: Uint8Array): number {
  const view = viewOf(buf)
  for (let i = buf.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === SIG_EOCD) return i
  }
  throw new Error("EOCD not found")
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
  }
  const total = chunks.reduce((n, c) => n + c.length, 0)
  const out = new Uint8Array(total)
  let pos = 0
  for (const c of chunks) {
    out.set(c, pos)
    pos += c.length
  }
  return out
}

/** Emit bytes as a ReadableStream in fixed-size chunks. */
function chunked(data: Uint8Array, chunkSize: number): ReadableStream<Uint8Array> {
  let offset = 0
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (offset >= data.length) {
        controller.close()
        return
      }
      const end = Math.min(offset + chunkSize, data.length)
      controller.enqueue(data.subarray(offset, end))
      offset = end
    },
  })
}

// ═══════════════════════════════════════════════════════════════════════
// ZipReader — central directory integrity
// ═══════════════════════════════════════════════════════════════════════

describe("ZipReader — central directory integrity", () => {
  it("rejects a directory that claims more entries than the file holds", () => {
    // Truncated downloads keep a plausible EOCD but lose trailing records;
    // walking off the end must be a ZipError, not a raw RangeError.
    const buf = buildZip([{ name: "a.txt", body: enc.encode("hello") }])
    viewOf(buf).setUint16(findEocd(buf) + 10, 2, true)

    expect(() => new ZipReader(buf)).toThrow(ZipError)
    expect(() => new ZipReader(buf)).toThrow(/Central Directory entry extends beyond file/)
  })

  it("rejects a central-directory record with a corrupted signature", () => {
    const buf = buildZip([{ name: "a.txt", body: enc.encode("hello") }])
    const eocd = findEocd(buf)
    const centralOffset = viewOf(buf).getUint32(eocd + 16, true)
    viewOf(buf).setUint32(centralOffset, 0xdeadbeef, true)

    expect(() => new ZipReader(buf)).toThrow(/Invalid Central Directory signature/)
  })

  it("skips directory entries when extracting everything", async () => {
    // Zip tools record folders as zero-length entries whose name ends in
    // "/". They are structure, not content, and must not appear as files.
    const buf = buildZip([
      { name: "docs/", body: new Uint8Array(0) },
      { name: "docs/a.txt", body: enc.encode("inside") },
    ])
    const all = await new ZipReader(buf).extractAll()

    expect([...all.keys()]).toEqual(["docs/a.txt"])
    expect(dec.decode(all.get("docs/a.txt")!)).toBe("inside")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ZipReader — ZIP64 escapes in the central directory
// ═══════════════════════════════════════════════════════════════════════

describe("ZipReader — ZIP64 extra field", () => {
  /** Central record with every 32-bit field escaped to the sentinel. */
  function escapedEntry(extra: Uint8Array): EntrySpec {
    return {
      name: "a.txt",
      body: enc.encode("payload"),
      centralSizes: { compressed: SENTINEL, uncompressed: SENTINEL },
      centralOffset: SENTINEL,
      centralExtra: extra,
    }
  }

  it("walks past unrelated extra fields to find the ZIP64 record", () => {
    // Info-ZIP writes an "extended timestamp" (0x5455) field ahead of the
    // ZIP64 one; the reader has to skip header ids it does not know.
    const payload = enc.encode("payload")
    const extra = new Uint8Array(timestampExtra().length + 28)
    extra.set(timestampExtra(), 0)
    extra.set(
      zip64Extra([BigInt(payload.length), BigInt(payload.length), 0n]),
      timestampExtra().length,
    )

    const zip = new ZipReader(buildZip([escapedEntry(extra)]))
    expect(zip.entries()).toEqual(["a.txt"])
  })

  it("rejects a ZIP64 extra field that stops mid-record", () => {
    // Three fields are escaped but only one 64-bit value is present.
    const buf = buildZip([escapedEntry(zip64Extra([0n]))])
    expect(() => new ZipReader(buf)).toThrow(/Truncated ZIP64 extra field for entry: a\.txt/)
  })

  it("rejects a ZIP64 extra field whose declared size runs past the file", () => {
    // The header claims 24 bytes of ZIP64 data but the extra field length
    // stops at the header, so the third value would be read out of the EOCD
    // and beyond the end of the archive.
    const extra = zip64Extra([], { declaredSize: 24 })
    expect(() => new ZipReader(buildZip([escapedEntry(extra)]))).toThrow(/extends beyond file/)
  })

  it("refuses a ZIP64 size larger than JavaScript can index", () => {
    // 2^60 bytes is a valid ZIP64 value and an unusable JS array index —
    // surfacing it as a ZipError beats silently truncating to a float.
    const extra = zip64Extra([1n << 60n, 7n, 0n])
    expect(() => new ZipReader(buildZip([escapedEntry(extra)]))).toThrow(
      /exceeds the addressable range/,
    )
  })

  it("rejects a ZIP64 locator pointing past the end of the file", async () => {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("hello"), { compress: false })
    const buf = await zip.build()
    const grown = new Uint8Array(buf.length + 42)
    grown.set(buf.subarray(0, buf.length - 22), 0)

    // Locator immediately before the EOCD, pointing at a record that would
    // start 10 bytes from the end (a ZIP64 EOCD needs 56).
    const dv = viewOf(grown)
    const locator = grown.length - 42
    dv.setUint32(locator, 0x07064b50, true)
    dv.setBigUint64(locator + 8, BigInt(grown.length - 10), true)
    grown.set(buf.subarray(buf.length - 22), grown.length - 22)
    dv.setUint16(grown.length - 22 + 8, 0xffff, true)
    dv.setUint16(grown.length - 22 + 10, 0xffff, true)

    expect(() => new ZipReader(grown)).toThrow(
      /ZIP64 End of Central Directory record extends beyond file/,
    )
  })

  it("rejects a ZIP64 EOCD record with the wrong signature", async () => {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("hello"), { compress: false })
    const buf = await zip.build()
    const grown = new Uint8Array(buf.length + 98)
    grown.set(buf.subarray(0, buf.length - 22), 0)

    const dv = viewOf(grown)
    const record = buf.length - 22
    dv.setUint32(record, 0x06064b50 ^ 0xff, true) // corrupted record signature
    const locator = record + 56
    dv.setUint32(locator, 0x07064b50, true)
    dv.setBigUint64(locator + 8, BigInt(record), true)
    grown.set(buf.subarray(buf.length - 22), locator + 20)
    dv.setUint16(locator + 20 + 8, 0xffff, true)
    dv.setUint16(locator + 20 + 10, 0xffff, true)

    expect(() => new ZipReader(grown)).toThrow(/Invalid ZIP64 End of Central Directory signature/)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ZipReader — local headers and entry bodies
// ═══════════════════════════════════════════════════════════════════════

describe("ZipReader — local header and body validation", () => {
  it("rejects an entry whose local header offset lands past the end", async () => {
    const buf = buildZip([{ name: "a.txt", body: enc.encode("hello") }])
    viewOf(buf).setUint32(findEocd(buf) - 46 - 5 + 42, buf.length - 10, true)

    const zip = new ZipReader(buf)
    await expect(zip.extract("a.txt")).rejects.toThrow(/Local file header extends beyond file/)
  })

  it("rejects an entry whose local header signature is wrong", async () => {
    const buf = buildZip([{ name: "a.txt", body: enc.encode("hello") }])
    viewOf(buf).setUint32(0, 0x504b0304, true) // byte-swapped signature

    const zip = new ZipReader(buf)
    await expect(zip.extract("a.txt")).rejects.toThrow(/Invalid local file header signature/)
  })

  it("rejects a compressed size that runs past the end of the archive", async () => {
    const buf = buildZip([
      {
        name: "a.txt",
        body: enc.encode("hello"),
        centralSizes: { compressed: 4096, uncompressed: 4096 },
      },
    ])
    const zip = new ZipReader(buf)
    await expect(zip.extract("a.txt")).rejects.toThrow(
      /Compressed data extends beyond file for entry: a\.txt/,
    )
  })

  it("rejects a compression method it cannot decode", async () => {
    // Method 12 is BZIP2 — legal in the spec, unsupported here.
    const buf = buildZip([{ name: "a.txt", body: enc.encode("hello"), method: 12, crc: 0 }])
    const zip = new ZipReader(buf)
    await expect(zip.extract("a.txt")).rejects.toThrow(
      /Unsupported compression method 12 for entry: a\.txt/,
    )
  })

  it("returns nothing for a DEFLATE entry declared empty on both sides", async () => {
    // Excel writes zero-length parts as method 8 with both sizes 0 and no
    // deflate stream at all; inflating that would throw.
    const buf = buildZip([
      {
        name: "empty.xml",
        body: new Uint8Array(0),
        method: 8,
        crc: 0,
        centralSizes: { compressed: 0, uncompressed: 0 },
      },
    ])
    expect(await new ZipReader(buf).extract("empty.xml")).toHaveLength(0)
  })

  it("reports a CRC-32 mismatch rather than handing back corrupt bytes", async () => {
    const buf = buildZip([{ name: "a.txt", body: enc.encode("hello"), crc: 0x12345678 }])
    await expect(new ZipReader(buf).extract("a.txt")).rejects.toThrow(/CRC-32 mismatch for a\.txt/)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ZipReader — data descriptors (general-purpose flag bit 3)
// ═══════════════════════════════════════════════════════════════════════

describe("ZipReader — data descriptor entries", () => {
  const payload = enc.encode("streamed content")
  const payloadCrc = crc32(payload)

  it("falls back to the local header when the directory sizes are zero", () => {
    const buf = buildZip([
      {
        name: "a.txt",
        body: payload,
        flags: 0x08,
        crc: payloadCrc,
        centralSizes: { compressed: 0, uncompressed: 0 },
      },
    ])
    return expect(new ZipReader(buf).extract("a.txt").then(dec.decode.bind(dec))).resolves.toBe(
      "streamed content",
    )
  })

  it("scans for the data descriptor when neither header carries a size", async () => {
    const buf = buildZip([
      {
        name: "a.txt",
        body: payload,
        flags: 0x08,
        crc: 0,
        centralSizes: { compressed: 0, uncompressed: 0 },
        localSizes: { compressed: 0, uncompressed: 0 },
        descriptor: { crc: payloadCrc, compressed: payload.length, uncompressed: payload.length },
      },
    ])
    expect(dec.decode(await new ZipReader(buf).extract("a.txt"))).toBe("streamed content")
  })

  it("ignores a ZIP64 sentinel in the local header instead of reading 4 GiB", async () => {
    // A ZIP64 streaming producer writes 0xFFFFFFFF in the local size fields
    // and the real value in the local extra field. Taking the sentinel
    // literally would index far past the archive.
    const buf = buildZip([
      {
        name: "a.txt",
        body: payload,
        flags: 0x08,
        crc: 0,
        centralSizes: { compressed: 0, uncompressed: 0 },
        localSizes: { compressed: SENTINEL, uncompressed: SENTINEL },
        descriptor: { crc: payloadCrc, compressed: payload.length, uncompressed: payload.length },
      },
    ])
    expect(dec.decode(await new ZipReader(buf).extract("a.txt"))).toBe("streamed content")
  })

  it("yields an empty entry when no descriptor can be located", async () => {
    const buf = buildZip([
      {
        name: "a.txt",
        body: payload,
        flags: 0x08,
        crc: 0,
        centralSizes: { compressed: 0, uncompressed: 0 },
        localSizes: { compressed: 0, uncompressed: 0 },
      },
    ])
    expect(await new ZipReader(buf).extract("a.txt")).toHaveLength(0)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ZipReader — extractStream
// ═══════════════════════════════════════════════════════════════════════

describe("ZipReader — extractStream", () => {
  it("throws for an entry that is not in the archive", () => {
    const buf = buildZip([{ name: "a.txt", body: enc.encode("hello") }])
    expect(() => new ZipReader(buf).extractStream("missing.txt")).toThrow(
      /Entry not found: missing\.txt/,
    )
  })

  it("streams a stored entry verbatim", async () => {
    const buf = buildZip([{ name: "a.txt", body: enc.encode("stored bytes") }])
    const out = await collect(new ZipReader(buf).extractStream("a.txt"))
    expect(dec.decode(out)).toBe("stored bytes")
  })

  it("streams a DEFLATE entry that declares no uncompressed size", async () => {
    // Without a declared size the stream is bounded only by the absolute
    // decompression cap.
    const body = deflate(enc.encode("compressed payload ".repeat(20)))
    const buf = buildZip([
      {
        name: "a.txt",
        body,
        method: 8,
        crc: 0,
        centralSizes: { compressed: body.length, uncompressed: 0 },
      },
    ])
    const out = await collect(new ZipReader(buf).extractStream("a.txt"))
    expect(dec.decode(out)).toBe("compressed payload ".repeat(20))
  })

  it("streams nothing for a DEFLATE entry declared empty", async () => {
    const buf = buildZip([
      {
        name: "empty.xml",
        body: new Uint8Array(0),
        method: 8,
        crc: 0,
        centralSizes: { compressed: 0, uncompressed: 0 },
      },
    ])
    expect(await collect(new ZipReader(buf).extractStream("empty.xml"))).toHaveLength(0)
  })

  it("recovers the size from the local header for a data-descriptor entry", async () => {
    const payload = enc.encode("descriptor streamed")
    const buf = buildZip([
      {
        name: "a.txt",
        body: payload,
        flags: 0x08,
        crc: crc32(payload),
        centralSizes: { compressed: 0, uncompressed: 0 },
      },
    ])
    const out = await collect(new ZipReader(buf).extractStream("a.txt"))
    expect(dec.decode(out)).toBe("descriptor streamed")
  })

  it("scans for the data descriptor when neither header carries a size", async () => {
    const payload = enc.encode("descriptor streamed")
    const buf = buildZip([
      {
        name: "a.txt",
        body: payload,
        flags: 0x08,
        crc: 0,
        centralSizes: { compressed: 0, uncompressed: 0 },
        localSizes: { compressed: 0, uncompressed: 0 },
        descriptor: {
          crc: crc32(payload),
          compressed: payload.length,
          uncompressed: payload.length,
        },
      },
    ])
    const out = await collect(new ZipReader(buf).extractStream("a.txt"))
    expect(dec.decode(out)).toBe("descriptor streamed")
  })

  it("rejects a local header that extends past the end of the archive", () => {
    const buf = buildZip([{ name: "a.txt", body: enc.encode("hello") }])
    viewOf(buf).setUint32(findEocd(buf) - 46 - 5 + 42, buf.length - 10, true)
    expect(() => new ZipReader(buf).extractStream("a.txt")).toThrow(
      /Local file header extends beyond file/,
    )
  })

  it("rejects a wrong local header signature", () => {
    const buf = buildZip([{ name: "a.txt", body: enc.encode("hello") }])
    viewOf(buf).setUint32(0, 0x504b0304, true)
    expect(() => new ZipReader(buf).extractStream("a.txt")).toThrow(
      /Invalid local file header signature/,
    )
  })

  it("rejects a compressed size that runs past the end of the archive", () => {
    const buf = buildZip([
      {
        name: "a.txt",
        body: enc.encode("hello"),
        centralSizes: { compressed: 4096, uncompressed: 5 },
      },
    ])
    expect(() => new ZipReader(buf).extractStream("a.txt")).toThrow(
      /Compressed data extends beyond file for entry: a\.txt/,
    )
  })

  it("rejects a compression method it cannot decode", () => {
    const buf = buildZip([{ name: "a.txt", body: enc.encode("hello"), method: 12, crc: 0 }])
    expect(() => new ZipReader(buf).extractStream("a.txt")).toThrow(
      /Unsupported compression method 12 for entry: a\.txt/,
    )
  })

  // BUG (reported): src/zip/reader.ts:490 adopts the local compressed size
  // without the ZIP64-sentinel guard that its buffered twin applies at
  // src/zip/reader.ts:408. The same entry therefore extracts fine through
  // `extract()` (see "ignores a ZIP64 sentinel in the local header…" above)
  // but throws "Compressed data extends beyond file" through
  // `extractStream()`, because 0xFFFFFFFF is taken as a real length.
  it("ignores a ZIP64 sentinel in the local header when streaming", async () => {
    const payload = enc.encode("streamed content")
    const buf = buildZip([
      {
        name: "a.txt",
        body: payload,
        flags: 0x08,
        crc: 0,
        centralSizes: { compressed: 0, uncompressed: 0 },
        localSizes: { compressed: SENTINEL, uncompressed: SENTINEL },
        descriptor: {
          crc: crc32(payload),
          compressed: payload.length,
          uncompressed: payload.length,
        },
      },
    ])
    const out = await collect(new ZipReader(buf).extractStream("a.txt"))
    expect(dec.decode(out)).toBe("streamed content")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Pure-TypeScript inflate/deflate
// ═══════════════════════════════════════════════════════════════════════

describe("inflate — malformed streams", () => {
  it("rejects a reserved DEFLATE block type", () => {
    // BFINAL=1, BTYPE=11 (reserved) packed into one byte.
    expect(() => inflate(new Uint8Array([0b111]))).toThrow(/Invalid DEFLATE block type: 3/)
  })

  it("rejects a stream that ends before the block header", () => {
    expect(() => inflate(new Uint8Array(0))).toThrow(/Unexpected end of DEFLATE data/)
  })

  it("rejects a stored block whose payload is truncated", () => {
    // BFINAL=1 BTYPE=00, LEN=8, NLEN, then only 2 of the 8 bytes.
    const data = new Uint8Array([0x01, 0x08, 0x00, 0xf7, 0xff, 0x41, 0x42])
    expect(() => inflate(data)).toThrow(/Unexpected end of DEFLATE data/)
  })

  it("refuses to expand past the caller's byte ceiling", () => {
    const compressed = deflate(new Uint8Array(4096))
    expect(() => inflate(compressed, 64)).toThrow(ZipError)
    expect(() => inflate(compressed, 64)).toThrow(/possible zip bomb/)
  })
})

describe("deflate — output buffer growth", () => {
  it("round-trips incompressible data larger than its initial buffer", () => {
    // Fixed-Huffman literals cost ~9 bits/byte, so random input of this size
    // overflows the `length + 512` output buffer and forces a re-allocation.
    let seed = 0x2545f491
    const data = new Uint8Array(16384).map(() => {
      seed = (seed * 1664525 + 1013904223) >>> 0
      return (seed >>> 24) & 0xff
    })
    expect([...inflate(deflate(data))]).toEqual([...data])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ZipStreamReader — local-header streaming
// ═══════════════════════════════════════════════════════════════════════

describe("ZipStreamReader — entry sequencing", () => {
  async function zipBytes(): Promise<Uint8Array> {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("first entry"), { compress: false })
    zip.add("b.txt", enc.encode("second entry ".repeat(40)))
    return zip.build()
  }

  it("refuses to advance while the previous body is unread", async () => {
    const reader = new ZipStreamReader(chunked(await zipBytes(), 64))
    await reader.nextEntry()
    await expect(reader.nextEntry()).rejects.toThrow(/previous entry body not consumed/)
    await reader.close()
  })

  it("refuses to read the same body twice", async () => {
    const reader = new ZipStreamReader(chunked(await zipBytes(), 64))
    const entry = (await reader.nextEntry())!
    await reader.readEntryBytes(entry)
    await expect(reader.readEntryBytes(entry)).rejects.toThrow(/entry body already consumed/)
    expect(() => reader.entryStream(entry)).toThrow(/entry body already consumed/)
    await reader.close()
  })

  it("skipEntry is a no-op once the body has been consumed", async () => {
    const reader = new ZipStreamReader(chunked(await zipBytes(), 64))
    const entry = (await reader.nextEntry())!
    await reader.readEntryBytes(entry)
    await expect(reader.skipEntry()).resolves.toBeUndefined()
    await reader.close()
  })

  it("reassembles headers split across single-byte chunks", async () => {
    // Every field of the local header lands on a chunk boundary here, which
    // is the whole point of the leftover buffer.
    const reader = new ZipStreamReader(chunked(await zipBytes(), 1))
    const a = (await reader.nextEntry())!
    expect(a.name).toBe("a.txt")
    expect(dec.decode(await reader.readEntryBytes(a))).toBe("first entry")
    const b = (await reader.nextEntry())!
    expect(dec.decode(await reader.readEntryBytes(b))).toBe("second entry ".repeat(40))
    expect(await reader.nextEntry()).toBeNull()
    await reader.close()
  })

  it("returns null when the source ends exactly at an entry boundary", async () => {
    // A ZIP truncated right after the last entry body has no central
    // directory left to signal the end — running out of bytes must do it.
    const full = await zipBytes()
    const view = viewOf(full)
    let centralStart = 0
    for (let i = 0; i < full.length - 4; i++) {
      if (view.getUint32(i, true) === SIG_CENTRAL) {
        centralStart = i
        break
      }
    }
    const reader = new ZipStreamReader(chunked(full.subarray(0, centralStart), 64))
    for (;;) {
      const entry = await reader.nextEntry()
      if (!entry) break
      await reader.skipEntry()
    }
    await reader.close()
  })

  it("reads and streams an entry with an empty body", async () => {
    const zip = new ZipWriter()
    zip.add("empty.txt", new Uint8Array(0))
    const bytes = await zip.build()

    const buffered = new ZipStreamReader(chunked(bytes, 16))
    const entry = (await buffered.nextEntry())!
    expect(entry.compressedSize).toBe(0)
    expect(await buffered.readEntryBytes(entry)).toHaveLength(0)
    await buffered.close()

    const streamed = new ZipStreamReader(chunked(bytes, 16))
    const entry2 = (await streamed.nextEntry())!
    expect(await collect(streamed.entryStream(entry2))).toHaveLength(0)
    await streamed.close()
  })

  it("marks an entry with an extra field and a data descriptor unstreamable", async () => {
    // Streaming writers set flag bit 3 and defer the sizes; the local header
    // alone is then not enough to find the body's end.
    const payload = enc.encode("streamed")
    const bytes = buildZip([
      {
        name: "a.txt",
        body: payload,
        flags: 0x08,
        crc: 0,
        localSizes: { compressed: 0, uncompressed: 0 },
        descriptor: {
          crc: crc32(payload),
          compressed: payload.length,
          uncompressed: payload.length,
        },
      },
    ])
    const reader = new ZipStreamReader(chunked(bytes, 8))
    const entry = (await reader.nextEntry())!
    expect(entry.streamable).toBe(false)
    await reader.close()
  })

  it("streams a stored entry without decompressing it", async () => {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("raw bytes"), { compress: false })
    const reader = new ZipStreamReader(chunked(await zip.build(), 4))
    const entry = (await reader.nextEntry())!
    expect(entry.compressionMethod).toBe(0)
    expect(dec.decode(await collect(reader.entryStream(entry)))).toBe("raw bytes")
    await reader.close()
  })
})

describe("ZipStreamReader — truncated sources", () => {
  async function truncatedAfter(bytes: number): Promise<Uint8Array> {
    const zip = new ZipWriter()
    zip.add("some/long/path/name.txt", enc.encode("body bytes here"), { compress: false })
    return (await zip.build()).subarray(0, bytes)
  }

  it("reports a local file header cut short", async () => {
    const reader = new ZipStreamReader(chunked(await truncatedAfter(12), 4))
    await expect(reader.nextEntry()).rejects.toThrow(/truncated local file header/)
    await reader.close()
  })

  it("reports an entry name cut short", async () => {
    const reader = new ZipStreamReader(chunked(await truncatedAfter(36), 4))
    await expect(reader.nextEntry()).rejects.toThrow(/truncated entry name/)
    await reader.close()
  })

  it("reports an extra field cut short", async () => {
    // Rewrite the header to declare a 16-byte extra field, then cut the file
    // where the name ends.
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("body"), { compress: false })
    const full = await zip.build()
    viewOf(full).setUint16(28, 16, true)
    const reader = new ZipStreamReader(chunked(full.subarray(0, 35), 4))
    await expect(reader.nextEntry()).rejects.toThrow(/truncated extra field/)
    await reader.close()
  })

  it("reports a body cut short while skipping", async () => {
    const reader = new ZipStreamReader(chunked(await truncatedAfter(60), 8))
    const entry = (await reader.nextEntry())!
    await expect(reader.skipEntry()).rejects.toThrow(/truncated entry body/)
    await reader.close()
    expect(entry.name).toBe("some/long/path/name.txt")
  })

  it("reports a body cut short while buffering", async () => {
    const reader = new ZipStreamReader(chunked(await truncatedAfter(60), 8))
    const entry = (await reader.nextEntry())!
    await expect(reader.readEntryBytes(entry)).rejects.toThrow(/truncated entry body/)
    await reader.close()
  })

  it("errors the entry stream when the body is cut short", async () => {
    const reader = new ZipStreamReader(chunked(await truncatedAfter(60), 8))
    const entry = (await reader.nextEntry())!
    await expect(collect(reader.entryStream(entry))).rejects.toThrow(/truncated entry body/)
    await reader.close()
  })
})

describe("ZipStreamReader — drainToBuffer", () => {
  async function zipBytes(): Promise<Uint8Array> {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("hello"), { compress: false })
    zip.add("b.txt", enc.encode("world"), { compress: false })
    return zip.build()
  }

  it("reconstructs the whole archive after inspecting the first header", async () => {
    const bytes = await zipBytes()
    const reader = new ZipStreamReader(chunked(bytes, 7))
    await reader.nextEntry()
    const drained = await reader.drainToBuffer()

    expect([...drained]).toEqual([...bytes])
    expect(new ZipReader(drained).entries()).toEqual(["a.txt", "b.txt"])
  })

  it("falls back to the default ceiling for a non-positive limit", async () => {
    const bytes = await zipBytes()
    const reader = new ZipStreamReader(chunked(bytes, 7))
    await reader.nextEntry()
    expect((await reader.drainToBuffer(0)).length).toBe(bytes.length)
  })

  it("refuses to buffer more than the ceiling allows", async () => {
    const reader = new ZipStreamReader(chunked(await zipBytes(), 7))
    await reader.nextEntry()
    await expect(reader.drainToBuffer(8)).rejects.toBeInstanceOf(ParseError)
    await reader.close()
  })

  it("refuses to fall back once streaming has started", async () => {
    const reader = new ZipStreamReader(chunked(await zipBytes(), 7))
    const entry = (await reader.nextEntry())!
    await collect(reader.entryStream(entry))
    await expect(reader.drainToBuffer()).rejects.toThrow(/cannot fall back after streaming started/)
    await reader.close()
  })
})

describe("ZipStreamReader — DEFLATE bodies", () => {
  async function deflatedZip(text: string): Promise<Uint8Array> {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode(text))
    return zip.build()
  }

  it("inflates a buffered body whose header declares no uncompressed size", async () => {
    const text = "repeated payload ".repeat(60)
    const bytes = await deflatedZip(text)
    // Zero out the local header's uncompressed size, as a streaming writer
    // that defers sizes would.
    viewOf(bytes).setUint32(22, 0, true)

    const reader = new ZipStreamReader(chunked(bytes, 64))
    const entry = (await reader.nextEntry())!
    expect(entry.compressionMethod).toBe(8)
    expect(dec.decode(await reader.readEntryBytes(entry))).toBe(text)
    await reader.close()
  })

  it("streams a body whose header declares no uncompressed size", async () => {
    const text = "repeated payload ".repeat(60)
    const bytes = await deflatedZip(text)
    viewOf(bytes).setUint32(22, 0, true)

    const reader = new ZipStreamReader(chunked(bytes, 64))
    const entry = (await reader.nextEntry())!
    expect(dec.decode(await collect(reader.entryStream(entry)))).toBe(text)
    await reader.close()
  })
})

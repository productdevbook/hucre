// ── ZipStreamReader and data descriptors ─────────────────────────────
// An entry with general-purpose bit 3 set states no sizes in its local
// header; they follow the body. The reader finds the body's end by
// scanning for the descriptor (see `bodyUntilDescriptor`). These pin the
// forms it must recognise and the ways a wrong boundary must fail loudly
// rather than yield wrong bytes.

import { describe, expect, it } from "vitest"
import { ZipError } from "../src/errors"
import { crc32 } from "../src/zip/deflate"
import { zipStream } from "../src/zip/stream-writer"
import { DataDescriptorUnresolved, ZipStreamReader } from "../src/zip/stream-reader"

const enc = new TextEncoder()
const dec = new TextDecoder()

function toStream(bytes: Uint8Array, chunk = 7): ReadableStream<Uint8Array> {
  // Small, odd-sized chunks, so descriptors and signatures straddle them.
  let at = 0
  return new ReadableStream({
    pull(c) {
      if (at >= bytes.length) return c.close()
      c.enqueue(bytes.slice(at, at + chunk))
      at += chunk
    },
  })
}

async function drain(s: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(s).arrayBuffer())
}

async function written(
  parts: Record<string, string>,
  zip64 = false,
  compress = true,
): Promise<Uint8Array> {
  const entries = Object.entries(parts).map(([path, s]) => ({
    path,
    data: enc.encode(s),
    compress,
  }))
  return drain(zipStream(entries, { zip64 }))
}

/** Read every entry: the first `whole` read whole, the rest streamed. */
async function readAll(bytes: Uint8Array): Promise<Record<string, string>> {
  const zr = new ZipStreamReader(toStream(bytes))
  const out: Record<string, string> = {}
  let n = 0
  for (let e = await zr.nextEntry(); e; e = await zr.nextEntry()) {
    expect(e.streamable).toBe(true)
    out[e.name] =
      n++ % 2 === 0
        ? dec.decode(await zr.readEntryBytes(e))
        : dec.decode(await drain(zr.entryStream(e)))
  }
  return out
}

/** Drop the optional `PK\x07\x08` from every descriptor — the other spelling. */
function withoutDescriptorSignatures(bytes: Uint8Array): Uint8Array {
  const out: number[] = []
  for (let i = 0; i < bytes.length; i++) {
    if (
      bytes[i] === 0x50 &&
      bytes[i + 1] === 0x4b &&
      bytes[i + 2] === 0x07 &&
      bytes[i + 3] === 0x08
    ) {
      i += 3
      continue
    }
    out.push(bytes[i]!)
  }
  return new Uint8Array(out)
}

const PARTS = {
  "a.xml": "<a>" + "alpha ".repeat(300) + "</a>",
  "b.xml": "<b/>",
  "empty.txt": "",
  "c.xml": "<c>" + "gamma ".repeat(3000) + "</c>",
}

describe("ZipStreamReader reads data-descriptor entries", () => {
  it("with the descriptor signature", async () => {
    expect(await readAll(await written(PARTS))).toEqual(PARTS)
  })

  it("without it", async () => {
    // The CD offsets are now wrong, which this reader never looks at.
    expect(await readAll(withoutDescriptorSignatures(await written(PARTS)))).toEqual(PARTS)
  })

  it("in the ZIP64 form, with and without the signature", async () => {
    const bytes = await written(PARTS, true)
    expect(await readAll(bytes)).toEqual(PARTS)
    expect(await readAll(withoutDescriptorSignatures(bytes))).toEqual(PARTS)
  })

  it("skips them", async () => {
    const zr = new ZipStreamReader(toStream(await written(PARTS)))
    const names: string[] = []
    for (let e = await zr.nextEntry(); e; e = await zr.nextEntry()) {
      names.push(e.name)
      if (e.name === "c.xml") expect(dec.decode(await zr.readEntryBytes(e))).toBe(PARTS["c.xml"])
      else await zr.skipEntry()
    }
    expect(names).toEqual(Object.keys(PARTS))
  })

  it("does not stream a STORE entry with a descriptor", async () => {
    // No end of its own, so a wrong boundary could not be noticed.
    const zr = new ZipStreamReader(toStream(await written(PARTS, false, false)))
    const e = await zr.nextEntry()
    expect(e?.compressionMethod).toBe(0)
    expect(e?.streamable).toBe(false)
  })
})

// ── Hand-built entries ───────────────────────────────────────────────

/** A DEFLATE stream of one stored block: the payload appears verbatim. */
function storedDeflate(payload: Uint8Array): Uint8Array {
  const out = new Uint8Array(5 + payload.length)
  out[0] = 0x01 // BFINAL, BTYPE=00
  out[1] = payload.length & 0xff
  out[2] = payload.length >>> 8
  out[3] = ~payload.length & 0xff
  out[4] = (~payload.length >>> 8) & 0xff
  out.set(payload, 5)
  return out
}

function u32(n: number): number[] {
  return [n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff]
}

/** One bit-3 DEFLATE entry, its descriptor, then the start of a central directory. */
function descriptorEntry(name: string, body: Uint8Array, d: { crc: number; usize: number }) {
  const nameBytes = enc.encode(name)
  const header = [
    ...u32(0x04034b50),
    20,
    0,
    0x08,
    0, // flags: bit 3
    8,
    0, // DEFLATE
    0,
    0,
    0,
    0,
    ...u32(0),
    ...u32(0),
    ...u32(0),
    nameBytes.length,
    0,
    0,
    0,
  ]
  const descriptor = [...u32(0x08074b50), ...u32(d.crc), ...u32(body.length), ...u32(d.usize)]
  return new Uint8Array([
    ...header,
    ...nameBytes,
    ...body,
    ...descriptor,
    ...u32(0x02014b50),
    ...Array.from({ length: 42 }, () => 0),
  ])
}

describe("ZIP64 local headers with their sizes stated", () => {
  it("streams an entry whose sizes are in the local ZIP64 extra field", async () => {
    // What a seekable ZIP64 writer leaves: bit 3 clear, the 32-bit size
    // fields at the sentinel, the real sizes in extra field 0x0001.
    // zipStream always sets bit 3, so this is built by hand.
    const payload = enc.encode("<z>" + "zeta ".repeat(200) + "</z>")
    const body = storedDeflate(payload)
    const nameBytes = enc.encode("z.xml")
    const extra = new Uint8Array(4 + 16)
    const ev = new DataView(extra.buffer)
    ev.setUint16(0, 0x0001, true)
    ev.setUint16(2, 16, true)
    ev.setBigUint64(4, BigInt(payload.length), true)
    ev.setBigUint64(12, BigInt(body.length), true)
    const header = new Uint8Array(30)
    const hv = new DataView(header.buffer)
    hv.setUint32(0, 0x04034b50, true)
    hv.setUint16(8, 8, true)
    hv.setUint32(14, crc32(payload), true)
    hv.setUint32(18, 0xffffffff, true)
    hv.setUint32(22, 0xffffffff, true)
    hv.setUint16(26, nameBytes.length, true)
    hv.setUint16(28, extra.length, true)
    const bytes = new Uint8Array([
      ...header,
      ...nameBytes,
      ...extra,
      ...body,
      ...u32(0x02014b50),
      ...Array.from({ length: 42 }, () => 0),
    ])
    for (const whole of [true, false]) {
      const zr = new ZipStreamReader(toStream(bytes))
      const e = (await zr.nextEntry())!
      expect(e).toMatchObject({ streamable: true, dataDescriptor: false })
      expect(e.compressedSize).toBe(body.length)
      const out = whole ? await zr.readEntryBytes(e) : await drain(zr.entryStream(e))
      expect(out).toEqual(payload)
      expect(await zr.nextEntry()).toBeNull()
    }
  })
})

describe("a wrong boundary fails loudly", () => {
  it("is not fooled by header signatures inside the body", async () => {
    // Stored blocks carry their payload verbatim, so the compressed body
    // itself contains `PK\x03\x04` and `PK\x07\x08` — but no size field
    // that equals its own offset.
    const payload = enc.encode("x PK\x03\x04 y PK\x07\x08 z PK\x01\x02 w")
    const body = storedDeflate(payload)
    const bytes = descriptorEntry("s.xml", body, { crc: crc32(payload), usize: payload.length })
    const zr = new ZipStreamReader(toStream(bytes, 3))
    const e = (await zr.nextEntry())!
    expect(await zr.readEntryBytes(e)).toEqual(payload)
    expect(await zr.nextEntry()).toBeNull()
  })

  it("rejects a part whose bytes do not match its descriptor", async () => {
    const payload = enc.encode("hello")
    const bytes = descriptorEntry("s.xml", storedDeflate(payload), { crc: 0x1234, usize: 5 })
    const zr = new ZipStreamReader(toStream(bytes))
    // Unresolved rather than plain damage: before the sheet the stream path
    // still holds every byte, and falls back to the buffered reader on it.
    const read = zr.readEntryBytes((await zr.nextEntry())!)
    await expect(read).rejects.toThrow(DataDescriptorUnresolved)
    await expect(read).rejects.toThrow(/does not match its data descriptor/)
  })

  it("rejects a forged descriptor inside the body instead of yielding a truncated part", async () => {
    // The payload plants a descriptor that claims to close the body at its
    // own offset, followed by a header signature — the one shape the scan
    // can be fooled by. The boundary it finds is wrong, and that has to be
    // an error, not a short read.
    const prefix = enc.encode("aaaa")
    // Body so far at the forged descriptor: 5 (block header) + prefix.
    const forgedAt = 5 + prefix.length
    const forged = new Uint8Array([
      ...u32(0x08074b50),
      ...u32(0),
      ...u32(forgedAt),
      ...u32(4),
      ...u32(0x04034b50),
    ])
    const payload = new Uint8Array([...prefix, ...forged, ...enc.encode("tail")])
    const body = storedDeflate(payload)
    const bytes = descriptorEntry("s.xml", body, { crc: crc32(payload), usize: payload.length })

    // Read whole, the cut-short body fails in the inflater before the CRC
    // can be checked. That is still the boundary being wrong, not the
    // archive, so it is Unresolved too: the stream path falls back on it.
    const whole = new ZipStreamReader(toStream(bytes))
    const read = whole.readEntryBytes((await whole.nextEntry())!)
    await expect(read).rejects.toThrow(DataDescriptorUnresolved)
    await expect(read).rejects.toMatchObject({ cause: expect.any(ZipError) })

    const streamed = new ZipStreamReader(toStream(bytes))
    const e = (await streamed.nextEntry())!
    await expect(drain(streamed.entryStream(e))).rejects.toThrow(ZipError)
  })

  it("reports a missing descriptor as DataDescriptorUnresolved, for the caller to fall back on", async () => {
    // Bit 3 set, sizes in the header, no descriptor written.
    const bytes = await written({ "a.xml": "<a/>" })
    const zr = new ZipStreamReader(toStream(withoutDescriptors(bytes)))
    const e = (await zr.nextEntry())!
    await expect(zr.readEntryBytes(e)).rejects.toThrow(DataDescriptorUnresolved)
  })
})

/** Cut every 16-byte signed descriptor out. Leaves bit 3 set and nothing after the bodies. */
function withoutDescriptors(bytes: Uint8Array): Uint8Array {
  const out: number[] = []
  for (let i = 0; i < bytes.length; i++) {
    if (
      bytes[i] === 0x50 &&
      bytes[i + 1] === 0x4b &&
      bytes[i + 2] === 0x07 &&
      bytes[i + 3] === 0x08
    ) {
      i += 15
      continue
    }
    out.push(bytes[i]!)
  }
  return new Uint8Array(out)
}

import { describe, expect, it } from "vitest"
import { ParseError, ZipError } from "../src/errors"
import { MAX_CENTRAL_DIRECTORY_BYTES } from "../src/limits"
import type { RandomAccessInput } from "../src/_types"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { zipStream } from "../src/zip/stream-writer"
import {
  RANDOM_ACCESS_WINDOW_BYTES,
  RandomAccessZip,
  randomAccessFromBlob,
} from "../src/zip/random-access"

const enc = new TextEncoder()
const dec = new TextDecoder()

/** A source over bytes in memory that records every read it serves. */
function counting(bytes: Uint8Array): RandomAccessInput & {
  reads: Array<[number, number]>
} {
  const reads: Array<[number, number]> = []
  return {
    size: bytes.length,
    reads,
    async read(offset, length) {
      reads.push([offset, length])
      return bytes.slice(offset, offset + length)
    },
  }
}

/**
 * A source that claims `size` bytes but holds only the regions planted in
 * it — zeros everywhere else. This is what a sparse file or a fabricated
 * source looks like, and why `size` cannot be what bounds an allocation.
 */
function sparse(size: number, planted: Array<[number, Uint8Array]>) {
  const reads: Array<[number, number]> = []
  return {
    size,
    reads,
    async read(offset: number, length: number): Promise<Uint8Array> {
      reads.push([offset, length])
      const out = new Uint8Array(length)
      for (const [at, bytes] of planted) {
        const from = Math.max(at, offset)
        const to = Math.min(at + bytes.length, offset + length)
        if (from < to) out.set(bytes.subarray(from - at, to - at), from - offset)
      }
      return out
    },
  }
}

/**
 * Incompressible bytes from xorshift32 with a fixed seed, so runs are
 * reproducible. (An LCG's top byte is not: DEFLATE took 200 KB of it to
 * 17 KB, which would let a window test pass without testing anything.)
 */
function noise(n: number): Uint8Array {
  const out = new Uint8Array(n)
  let x = 0x9e3779b9
  for (let i = 0; i < n; i++) {
    x ^= x << 13
    x ^= x >>> 17
    x ^= x << 5
    out[i] = x & 0xff
  }
  return out
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
  return true
}

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

function eocd(fields: { count: number; cdSize: number; cdOffset: number }): Uint8Array {
  const b = new Uint8Array(22)
  const v = new DataView(b.buffer)
  v.setUint32(0, 0x06054b50, true)
  v.setUint16(8, Math.min(fields.count, 0xffff), true)
  v.setUint16(10, Math.min(fields.count, 0xffff), true)
  v.setUint32(12, fields.cdSize, true)
  v.setUint32(16, fields.cdOffset, true)
  return b
}

/**
 * ZIP64 record + locator claiming the given bounds, then a sentinel EOCD —
 * 98 bytes, to be planted at the very end of a source of `size` bytes.
 */
function zip64Tail(
  size: number,
  fields: { count: number; cdSize: number; cdOffset: number },
): Uint8Array {
  const recordOffset = size - 98
  const record = new Uint8Array(56)
  const rv = new DataView(record.buffer)
  rv.setUint32(0, 0x06064b50, true)
  rv.setBigUint64(24, BigInt(fields.count), true)
  rv.setBigUint64(32, BigInt(fields.count), true)
  rv.setBigUint64(40, BigInt(fields.cdSize), true)
  rv.setBigUint64(48, BigInt(fields.cdOffset), true)
  const locator = new Uint8Array(20)
  const lv = new DataView(locator.buffer)
  lv.setUint32(0, 0x07064b50, true)
  lv.setBigUint64(8, BigInt(recordOffset), true)
  const end = eocd({ count: 0xffff, cdSize: 0xffffffff, cdOffset: 0xffffffff })
  const out = new Uint8Array(56 + 20 + 22)
  out.set(record, 0)
  out.set(locator, 56)
  out.set(end, 76)
  return out
}

function sampleParts(): Record<string, Uint8Array> {
  return {
    "a.txt": enc.encode("hello"),
    "dir/b.xml": enc.encode("<x>" + "row ".repeat(5000) + "</x>"),
    "empty.bin": new Uint8Array(0),
  }
}

async function viaStreamWriter(
  parts: Record<string, Uint8Array>,
  zip64 = false,
): Promise<Uint8Array> {
  const entries = Object.entries(parts).map(([path, data]) => ({ path, data }))
  return drain(zipStream(entries, { zip64 }))
}

describe("RandomAccessZip — reads what ZipReader reads", () => {
  const layouts: Array<[string, () => Promise<Uint8Array>]> = [
    [
      "sizes in the local headers",
      async () => {
        const z = new ZipWriter()
        for (const [p, d] of Object.entries(sampleParts())) z.add(p, d)
        return z.build()
      },
    ],
    ["data descriptors on every entry", () => viaStreamWriter(sampleParts())],
    ["ZIP64 records throughout", () => viaStreamWriter(sampleParts(), true)],
  ]

  for (const [label, make] of layouts) {
    it(label, async () => {
      const bytes = await make()
      const buffered = new ZipReader(bytes)
      const zip = await RandomAccessZip.open(counting(bytes))
      expect(zip.entries().sort()).toEqual(buffered.entries().sort())
      for (const path of buffered.entries()) {
        expect(await zip.extract(path)).toEqual(await buffered.extract(path))
        expect(await drain(zip.extractStream(path))).toEqual(await buffered.extract(path))
      }
    })
  }

  it("works through randomAccessFromBlob", async () => {
    const bytes = await viaStreamWriter(sampleParts())
    const zip = await RandomAccessZip.open(
      randomAccessFromBlob(new Blob([bytes as Uint8Array<ArrayBuffer>])),
    )
    expect(dec.decode(await zip.extract("a.txt"))).toBe("hello")
  })

  it("streams a large entry one window at a time, only as it is read", async () => {
    // Incompressible, so the compressed body is many windows long.
    const big = noise(RANDOM_ACCESS_WINDOW_BYTES * 5)
    const bytes = await viaStreamWriter({ "big.bin": big, "small.txt": enc.encode("s") })
    expect(bytes.length).toBeGreaterThan(big.length * 0.99)
    const src = counting(bytes)
    const zip = await RandomAccessZip.open(src)
    const opened = src.reads.length

    const reader = zip.extractStream("big.bin").getReader()
    expect(src.reads.length).toBe(opened) // nothing fetched before the first pull
    await reader.read()
    const afterFirst = src.reads.slice(opened).reduce((n, [, len]) => n + len, 0)
    expect(afterFirst).toBeLessThan(RANDOM_ACCESS_WINDOW_BYTES * 3)
    await reader.cancel()

    // A byte loop: vitest's deep equality takes seconds on 5 MiB.
    expect(sameBytes(await drain(zip.extractStream("big.bin")), big)).toBe(true)
    for (const [, len] of src.reads) expect(len).toBeLessThanOrEqual(RANDOM_ACCESS_WINDOW_BYTES)
  })
})

describe("RandomAccessZip — a hostile directory cannot buy an allocation", () => {
  const HUGE = 2 ** 50

  it("refuses a directory larger than the cap before reading it", async () => {
    const cdSize = 0xfffffff0
    const src = sparse(HUGE, [
      [HUGE - 22, eocd({ count: 10, cdSize, cdOffset: HUGE - 22 - cdSize - 100 })],
    ])
    // A 32-bit offset cannot point that far — use ZIP64 to place it.
    const tail = zip64Tail(HUGE, {
      count: 10,
      cdSize,
      cdOffset: HUGE - 4000 - cdSize,
    })
    const src64 = sparse(HUGE, [[HUGE - tail.length, tail]])
    await expect(RandomAccessZip.open(src64)).rejects.toThrow(
      new RegExp(`exceeds the limit of ${MAX_CENTRAL_DIRECTORY_BYTES} bytes`),
    )
    expect(Math.max(...src64.reads.map(([, n]) => n))).toBeLessThanOrEqual(
      RANDOM_ACCESS_WINDOW_BYTES,
    )
    expect(src64.reads.length).toBeLessThanOrEqual(2) // the tail, the ZIP64 record
    // The 32-bit form points before the start of the file, and says so.
    await expect(RandomAccessZip.open(src)).rejects.toThrow(ZipError)
  })

  it("refuses an entry count the directory's size cannot hold", async () => {
    const tail = zip64Tail(HUGE, { count: 2 ** 40, cdSize: 4096, cdOffset: HUGE - 8192 })
    const src = sparse(HUGE, [[HUGE - tail.length, tail]])
    await expect(RandomAccessZip.open(src)).rejects.toThrow(/claims 1099511627776 entries/)
    expect(src.reads.length).toBeLessThanOrEqual(2)
  })

  it("refuses a directory that runs past the end of the source", async () => {
    const end = eocd({ count: 1, cdSize: 1000, cdOffset: 500 })
    const src = sparse(1000, [[1000 - 22, end]])
    await expect(RandomAccessZip.open(src)).rejects.toThrow(/Central Directory extends beyond file/)
  })

  it("holds the directory to maxBufferedBytes too", async () => {
    const bytes = await viaStreamWriter(sampleParts())
    await expect(RandomAccessZip.open(counting(bytes), { maxBufferedBytes: 64 })).rejects.toThrow(
      /Raise `maxInputBytes`/,
    )
  })

  it("refuses an entry whose stated size would exceed the buffer budget, before reading it", async () => {
    // Bigger than the 64 KiB tail read, so the budget can admit the tail
    // and the directory and still be short of this entry.
    const bytes = await viaStreamWriter({ "a.txt": enc.encode("a"), "dir/b.xml": noise(200_000) })
    const src = counting(bytes)
    const zip = await RandomAccessZip.open(src, { maxBufferedBytes: 70_000 })
    const before = src.reads.length
    // Past the budget once the directory is counted.
    await expect(zip.extract("dir/b.xml")).rejects.toThrow(/Raise `maxInputBytes`/)
    // Only the local header was read, not the body.
    expect(src.reads.slice(before).every(([, n]) => n <= 30)).toBe(true)
  })

  it("refuses an entry whose body runs past the end of the source", async () => {
    // Patch the directory record's compressed size to something enormous.
    const bytes = await viaStreamWriter({ "a.txt": enc.encode("hello") })
    const out = bytes.slice()
    const v = new DataView(out.buffer)
    for (let i = out.length - 22; i >= 0; i--) {
      if (v.getUint32(i, true) === 0x02014b50) {
        v.setUint32(i + 20, 0x7fffffff, true)
        break
      }
    }
    const zip = await RandomAccessZip.open(counting(out))
    await expect(zip.extract("a.txt")).rejects.toThrow(/extends beyond file/)
    await expect(drain(zip.extractStream("a.txt"))).rejects.toThrow(/extends beyond file/)
  })

  it("bounds the read loop for a lying source by the entry, and stops on bad data", async () => {
    // A directory entry claiming a 1 GiB DEFLATE body in a 2^50-byte source
    // of zeros. Zeros are not valid DEFLATE, so the inflater fails on the
    // first window instead of the reader pulling a gigabyte.
    const name = enc.encode("s.xml")
    const cd = new Uint8Array(46 + name.length)
    const cv = new DataView(cd.buffer)
    cv.setUint32(0, 0x02014b50, true)
    cv.setUint16(10, 8, true)
    cv.setUint32(20, 2 ** 30, true)
    cv.setUint32(24, 2 ** 31, true)
    cv.setUint16(28, name.length, true)
    cv.setUint32(42, 0, true)
    cd.set(name, 46)
    const local = new Uint8Array(30)
    new DataView(local.buffer).setUint32(0, 0x04034b50, true)
    const cdOffset = 2 ** 31
    const tail = zip64Tail(HUGE, { count: 1, cdSize: cd.length, cdOffset })
    const src = sparse(HUGE, [
      [0, local],
      [cdOffset, cd],
      [HUGE - tail.length, tail],
    ])
    const zip = await RandomAccessZip.open(src)
    const before = src.reads.length
    await expect(drain(zip.extractStream("s.xml"))).rejects.toThrow(ZipError)
    expect(src.reads.length - before).toBeLessThan(5)
  })
})

describe("RandomAccessZip — the read contract", () => {
  it("rejects a read that returns more than asked for", async () => {
    // What an HTTP server that ignores Range sends: the whole file.
    const bytes = await viaStreamWriter(sampleParts())
    const src: RandomAccessInput = { size: bytes.length, read: async () => bytes }
    await expect(RandomAccessZip.open(src)).rejects.toThrow(ParseError)
    await expect(RandomAccessZip.open(src)).rejects.toThrow(/exactly the bytes asked for/)
  })

  it("rejects a short read", async () => {
    const bytes = await viaStreamWriter(sampleParts())
    const src: RandomAccessInput = {
      size: bytes.length,
      read: async (o, n) => bytes.slice(o, o + n - 1),
    }
    await expect(RandomAccessZip.open(src)).rejects.toThrow(/returned \d+ bytes/)
  })

  it("rejects a size that is not a byte count", async () => {
    for (const size of [-1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      const src: RandomAccessInput = { size, read: async () => new Uint8Array(0) }
      await expect(RandomAccessZip.open(src)).rejects.toThrow(/non-negative integer/)
    }
  })

  it("rejects a source too small to be a ZIP", async () => {
    await expect(RandomAccessZip.open(counting(new Uint8Array(10)))).rejects.toThrow(
      /too small to be a valid ZIP/,
    )
  })
})

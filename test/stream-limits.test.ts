import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { ZipReader } from "../src/zip/reader"
import { ZipStreamReader } from "../src/zip/stream-reader"
import { readXlsx } from "../src/xlsx/reader"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import { readXls } from "../src/xls/reader"
import { writeCfb } from "../src/xlsx/crypto/cfb"
import { parseSaxStream } from "../src/xml/parser"
import { bufferReadableStream } from "../src/_input"
import { ParseError, XmlError, ZipError } from "../src/errors"
import { MAX_INPUT_BYTES } from "../src/limits"

const enc = new TextEncoder()
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

// ── Fixtures ─────────────────────────────────────────────────────────

async function xlsxWithSheetData(sheetData: string): Promise<Uint8Array> {
  const zip = new ZipWriter()
  zip.add(
    "[Content_Types].xml",
    enc.encode(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
        `</Types>`,
    ),
  )
  zip.add(
    "_rels/.rels",
    enc.encode(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
  )
  zip.add(
    "xl/workbook.xml",
    enc.encode(
      `<?xml version="1.0"?><workbook ${NS} ${R}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
  )
  zip.add(
    "xl/_rels/workbook.xml.rels",
    enc.encode(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
  )
  zip.add(
    "xl/worksheets/sheet1.xml",
    enc.encode(
      `<?xml version="1.0"?><worksheet ${NS} ${R}><sheetData>${sheetData}</sheetData></worksheet>`,
    ),
  )
  return zip.build()
}

/** A stream that hands out `bytes` in `chunk`-sized pieces and records cancellation. */
function streamOf(bytes: Uint8Array, chunk = 4096): ReadableStream<Uint8Array> {
  let pos = 0
  return new ReadableStream<Uint8Array>({
    pull(c) {
      if (pos >= bytes.length) {
        c.close()
        return
      }
      c.enqueue(bytes.subarray(pos, pos + chunk))
      pos += chunk
    },
  })
}

/**
 * A stream of `totalMiB` MiB in 64 KiB pieces, tracking how much was
 * actually pulled and whether it was cancelled.
 *
 * Deliberately finite: the real-world shape is an unbounded `response.body`,
 * but a test that never ends would hang CI on regression instead of failing
 * it. Anything past the cap is still never pulled, which is the property
 * being asserted.
 */
function boundedStream(totalMiB: number): {
  stream: ReadableStream<Uint8Array>
  pulled: () => number
  cancelled: () => boolean
} {
  const chunk = 64 * 1024
  const chunks = (totalMiB * 1024 * 1024) / chunk
  let sent = 0
  let cancelled = false
  const stream = new ReadableStream<Uint8Array>({
    pull(c) {
      if (sent >= chunks) {
        c.close()
        return
      }
      sent++
      c.enqueue(new Uint8Array(chunk))
    },
    cancel() {
      cancelled = true
    },
  })
  return { stream, pulled: () => sent * chunk, cancelled: () => cancelled }
}

/**
 * A ZIP holding one DEFLATE entry that really expands to `realBytes` while
 * its headers claim `claimedBytes` — the shape of every zip bomb.
 */
async function lyingZip(realBytes: number, claimedBytes: number): Promise<Uint8Array> {
  const w = new ZipWriter()
  w.add("big.bin", new Uint8Array(realBytes))
  const zip = await w.build()
  const v = new DataView(zip.buffer, zip.byteOffset, zip.byteLength)
  v.setUint32(22, claimedBytes, true) // local file header
  for (let i = 0; i + 46 <= zip.length; i++) {
    if (v.getUint32(i, true) === 0x02014b50) v.setUint32(i + 24, claimedBytes, true)
  }
  return zip
}

/**
 * Buffer a stream and report only how many bytes came back. Returning the
 * buffer itself would make a *failing* assertion try to pretty-print a
 * multi-megabyte typed array, which OOMs the worker and hides the failure.
 */
async function bufferedLength(
  stream: ReadableStream<Uint8Array>,
  maxBytes?: number,
): Promise<number> {
  return (await bufferReadableStream(stream, maxBytes)).length
}

async function drainStream(stream: ReadableStream<Uint8Array>): Promise<number> {
  const r = stream.getReader()
  let total = 0
  for (;;) {
    const { done, value } = await r.read()
    if (done) break
    total += value!.length
  }
  return total
}

// ═══════════════════════════════════════════════════════════════════════
// #363, batch 3 — the remaining unbounded paths: input buffering, the two
// native DecompressionStream paths, an argument-spread loop bound, and two
// stream readers that were never released. Each test asserts *completion*
// or a *typed throw*, never a value, because the failure mode is a hang, an
// OOM, or a raw RangeError.
// ═══════════════════════════════════════════════════════════════════════

describe("input stream size ceiling", () => {
  it("stops a stream past the cap instead of buffering until the process dies", async () => {
    // Uncapped, a stream that keeps yielding is buffered forever: a probe
    // fed 72 GiB in 90 s without the reader complaining once.
    const { stream } = boundedStream(16)
    await expect(bufferedLength(stream, 1 << 20)).rejects.toThrow(ParseError)
  }, 20_000)

  it("cancels the source it gave up on, without draining it first", async () => {
    const { stream, pulled, cancelled } = boundedStream(16)
    await expect(bufferedLength(stream, 1 << 20)).rejects.toThrow(
      /exceeds the maximum of 1048576 bytes/,
    )
    expect(cancelled()).toBe(true)
    expect(pulled()).toBeLessThan(4 * 1024 * 1024)
  }, 20_000)

  it("releases the reader after a successful read", async () => {
    // The final copy used to run outside any try/finally, so a throw there
    // left the source locked for the life of the process.
    const stream = streamOf(new Uint8Array(10_000))
    const out = await bufferReadableStream(stream)
    expect(out.length).toBe(10_000)
    expect(stream.locked).toBe(false)
  }, 20_000)

  it("is reachable from a reader entry point via maxInputBytes", async () => {
    const bytes = await xlsxWithSheetData(
      `<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c></row>`,
    )
    await expect(readXlsx(streamOf(bytes), { maxInputBytes: 64 })).rejects.toThrow(ParseError)
  }, 20_000)

  it("still reads a stream that fits under the cap", async () => {
    const bytes = await xlsxWithSheetData(
      `<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c></row>`,
    )
    const workbook = await readXlsx(streamOf(bytes))
    expect(workbook.sheets[0].rows[0][0]).toBe("a")
  }, 20_000)

  it("leaves room for real files", () => {
    // A cap that rejects legitimate input is itself the bug: the biggest
    // real-world workbooks are tens of MB.
    expect(MAX_INPUT_BYTES).toBeGreaterThanOrEqual(256 * 1024 * 1024)
  })
})

describe("zip bomb ceiling on the streaming decompression paths", () => {
  // 8 MiB of zeros against a 4 KiB declaration — 2000x, from a 8 KiB archive.
  const REAL = 8 * 1024 * 1024
  const CLAIMED = 4096

  it("stops ZipReader.extractStream, like the buffered path already did", async () => {
    const zip = await lyingZip(REAL, CLAIMED)
    await expect(new ZipReader(zip).extract("big.bin")).rejects.toThrow(ZipError)
    await expect(drainStream(new ZipReader(zip).extractStream("big.bin"))).rejects.toThrow(ZipError)
  }, 30_000)

  it("stops ZipStreamReader.entryStream", async () => {
    const zip = await lyingZip(REAL, CLAIMED)
    const zr = new ZipStreamReader(streamOf(zip, 1 << 16))
    const entry = await zr.nextEntry()
    expect(entry?.name).toBe("big.bin")
    await expect(drainStream(zr.entryStream(entry!))).rejects.toThrow(
      /Decompressed size exceeds limit/,
    )
  }, 30_000)

  it("still streams an honest entry end to end", async () => {
    const w = new ZipWriter()
    w.add("ok.txt", enc.encode("hello ".repeat(50_000)))
    const zip = await w.build()
    expect(await drainStream(new ZipReader(zip).extractStream("ok.txt"))).toBe(300_000)

    const zr = new ZipStreamReader(streamOf(zip, 1 << 16))
    const entry = await zr.nextEntry()
    expect(await drainStream(zr.entryStream(entry!))).toBe(300_000)
  }, 30_000)
})

describe("XLS shared string table size", () => {
  // Minimal BIFF8 builder — just enough globals for a big SST.
  const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff]
  const u32 = (n: number): number[] => [
    n & 0xff,
    (n >> 8) & 0xff,
    (n >> 16) & 0xff,
    (n >>> 24) & 0xff,
  ]
  const rec = (sid: number, data: number[]): number[] => [...u16(sid), ...u16(data.length), ...data]
  const bof = (dt: number): number[] =>
    rec(0x0809, [...u16(0x0600), ...u16(dt), ...u16(0), ...u16(0), ...u32(0), ...u32(0)])
  const eof = (): number[] => rec(0x000a, [])

  /** SST + CONTINUE records holding `count` one-character strings. */
  function sstRecords(count: number): number[] {
    const out: number[] = []
    let body: number[] = [...u32(count), ...u32(count)]
    let budget = 8224 - body.length
    let first = true
    for (let i = 0; i < count; i++) {
      if (budget < 4) {
        out.push(...rec(first ? 0x00fc : 0x003c, body))
        first = false
        body = []
        budget = 8224
      }
      body.push(...u16(1), 0, 97 + (i % 26))
      budget -= 4
    }
    if (body.length > 0) out.push(...rec(first ? 0x00fc : 0x003c, body))
    return out
  }

  function xlsWithSharedStrings(count: number): Uint8Array {
    const sheet = [...bof(0x0010), ...eof()]
    const globals = (sheetPos: number): number[] => [
      ...bof(0x0005),
      ...sstRecords(count),
      ...rec(0x0085, [...u32(sheetPos), 0, 0, 6, 0, ...[..."Sheet1"].map((c) => c.charCodeAt(0))]),
      ...eof(),
    ]
    const stream = [...globals(globals(0).length), ...sheet]
    return writeCfb([{ name: "Workbook", data: new Uint8Array(stream) }])
  }

  it("reads a workbook with more shared strings than fit in an argument list", async () => {
    // `sst.push(...parseSst(blocks))` overflows the call stack somewhere
    // above 100k strings, and the RangeError was reported as "malformed or
    // truncated" — a valid file rejected with a misleading message.
    const workbook = await readXls(xlsWithSharedStrings(200_000))
    expect(workbook.sheets).toHaveLength(1)
  }, 30_000)

  it("still reads a small shared string table", async () => {
    const workbook = await readXls(xlsWithSharedStrings(10))
    expect(workbook.sheets[0].name).toBe("Sheet1")
  }, 20_000)
})

describe("parseSaxStream cleanup", () => {
  it("cancels the source when a handler throws", async () => {
    // The XLSX row parser aborts by throwing out of a handler, which left
    // the ZIP/decompression stream underneath locked forever.
    let cancelled = false
    const stream = new ReadableStream<Uint8Array>({
      start(c) {
        c.enqueue(enc.encode("<root><a/><b/>"))
        c.enqueue(enc.encode("<c/></root>"))
        c.close()
      },
      cancel() {
        cancelled = true
      },
    })

    await expect(
      parseSaxStream(stream, {
        onOpenTag(tag) {
          if (tag === "b") throw new XmlError("handler abort")
        },
      }),
    ).rejects.toThrow(XmlError)
    expect(stream.locked).toBe(false)
    expect(cancelled).toBe(true)
  }, 20_000)

  it("releases the reader on a clean parse too", async () => {
    const stream = new ReadableStream<Uint8Array>({
      start(c) {
        c.enqueue(enc.encode("<root>ok</root>"))
        c.close()
      },
    })
    await parseSaxStream(stream, {})
    expect(stream.locked).toBe(false)
  }, 20_000)

  it("parses one very long text run in linear time", async () => {
    // Carrying the run across chunks re-copied it every time: 4 MiB took
    // 0.4 s, 32 MiB took 25 s. The assertion is completion inside the
    // timeout — this input took ~1 s after the fix and over a minute
    // before it, so a regression to O(n^2) blows straight through.
    const size = 48 * 1024 * 1024
    const bytes = enc.encode(`<r>${"x".repeat(size)}</r>`)
    let chars = 0
    await parseSaxStream(streamOf(bytes, 64 * 1024), {
      onText: (t) => {
        chars += t.length
      },
    })
    expect(chars).toBe(size)
  }, 20_000)

  it("never splits a long run inside an entity or a surrogate pair", async () => {
    const pad = "a".repeat(400_000)
    const xml = `<r>${pad}&amp;${pad}&lt;t&gt;${"&amp;".repeat(20_000)}\u{1F600}</r>`
    let out = ""
    await parseSaxStream(streamOf(enc.encode(xml), 64 * 1024), {
      onText: (t) => {
        out += t
      },
    })
    expect(out).toBe(`${pad}&${pad}<t>${"&".repeat(20_000)}\u{1F600}`)
  }, 20_000)

  it("delivers ordinary text in a single callback", async () => {
    const calls: string[] = []
    await parseSaxStream(streamOf(enc.encode("<r>hello &amp; goodbye</r>")), {
      onText: (t) => calls.push(t),
    })
    expect(calls).toEqual(["hello & goodbye"])
  }, 20_000)
})

describe("ZipStreamReader release on failure", () => {
  it("cancels the source when streaming XLSX fails mid-archive", async () => {
    // ZipStreamReader.close() existed but nothing called it, so an error in
    // prepareStreaming left the source locked with the archive half-read.
    const w = new ZipWriter()
    w.add("[Content_Types].xml", enc.encode("<Types><Default></Wrong></Types>"))
    w.add("_rels/.rels", enc.encode("<Relationships/>"))
    w.add("xl/worksheets/sheet1.xml", enc.encode("<worksheet><sheetData/></worksheet>"))
    // Incompressible tail so the source still has bytes left when we bail.
    w.add("xl/tail.bin", new Uint8Array(512 * 1024), { compress: false })
    const bytes = await w.build()

    let cancelled = false
    let pos = 0
    const stream = new ReadableStream<Uint8Array>({
      pull(c) {
        if (pos >= bytes.length) {
          c.close()
          return
        }
        c.enqueue(bytes.subarray(pos, pos + 4096))
        pos += 4096
      },
      cancel() {
        cancelled = true
      },
    })

    const drain = async (): Promise<void> => {
      for await (const _ of streamXlsxRows(stream)) {
        // consume
      }
    }
    await expect(drain()).rejects.toThrow(XmlError)
    expect(cancelled).toBe(true)
    expect(stream.locked).toBe(false)
    expect(pos).toBeLessThan(bytes.length)
  }, 20_000)
})

// ── streamXlsxRows over a random-access source, and the layouts that ──
// used to make it buffer.
//
// The large battery-cycler and instrument exports this was written for
// (hundreds of MB) put the worksheet before the workbook's rels and the
// content types; Excel puts it before styles and shared strings;
// ExcelJS's streaming writer puts it before nearly everything and adds a
// data descriptor to every entry. Walking a ReadableStream front to back,
// streamXlsxRows could stream none of them and buffered the archive.
//
// Memory is not measured here — RSS sampling is not deterministic enough
// for CI. What is measured is the cause of it: the bytes asked of the
// source, in total, per read and before the first row.

import { openAsBlob, readdirSync, readFileSync } from "node:fs"
import { open } from "node:fs/promises"
import { join } from "node:path"
import { describe, expect, it, vi } from "vitest"
import { EncryptedFileError, ParseError, ZipError } from "../src/errors"
import type { RandomAccessInput, StreamRow } from "../src/_types"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { zipStream } from "../src/zip/stream-writer"
import { ZipWriter } from "../src/zip/writer"
import { RANDOM_ACCESS_WINDOW_BYTES, randomAccessFromBlob } from "../src/zip/random-access"
import {
  type CentralDirEntry,
  readCentralDirectory,
  readEocd,
  readZip64Eocd,
  readZip64Locator,
} from "../src/zip/central-directory"

const FIXTURES = join(import.meta.dirname, "fixtures")
const enc = new TextEncoder()

// ── Sources ──────────────────────────────────────────────────────────

interface CountingSource extends RandomAccessInput {
  reads: Array<[number, number]>
  requested(): number
}

function counting(bytes: Uint8Array): CountingSource {
  const reads: Array<[number, number]> = []
  return {
    size: bytes.length,
    reads,
    requested: () => reads.reduce((n, [, len]) => n + len, 0),
    async read(offset, length) {
      reads.push([offset, length])
      return bytes.slice(offset, offset + length)
    },
  }
}

/** A ReadableStream over `bytes` that serves 64 KiB on demand and counts what it served. */
function pulledStream(bytes: Uint8Array): { stream: ReadableStream<Uint8Array>; pulled(): number } {
  let offset = 0
  const stream = new ReadableStream<Uint8Array>(
    {
      pull(controller) {
        if (offset >= bytes.length) {
          controller.close()
          return
        }
        const end = Math.min(offset + 65536, bytes.length)
        controller.enqueue(bytes.slice(offset, end))
        offset = end
      },
    },
    { highWaterMark: 0 },
  )
  return { stream, pulled: () => offset }
}

async function collect(gen: AsyncGenerator<StreamRow>): Promise<StreamRow[]> {
  const rows: StreamRow[] = []
  for await (const row of gen) rows.push(row)
  return rows
}

// ── Repacking ────────────────────────────────────────────────────────
// The same parts in a different order, or with a different ZIP framing.
// Every repack goes through hucre's streaming ZIP writer, which puts a
// data descriptor on every entry — the ExcelJS / archiver layout.

type Order = "sheets-first" | "metadata-first"

async function repack(bytes: Uint8Array, order: Order, zip64 = false): Promise<Uint8Array> {
  const parts = await new ZipReader(bytes).extractAll()
  const isSheet = (p: string) => /^xl\/worksheets\/[^/]+\.xml$/.test(p)
  const names = [...parts.keys()]
  const sheets = names.filter(isSheet)
  const rest = names.filter((p) => !isSheet(p))
  // Sheets first, [Content_Types].xml last: the worst case of the orders
  // above, which every one of them is a milder form of. Metadata first:
  // what local-header streaming needs.
  const ordered =
    order === "sheets-first"
      ? [
          ...sheets,
          ...rest.filter((p) => p !== "[Content_Types].xml"),
          ...rest.filter((p) => p === "[Content_Types].xml"),
        ]
      : [...rest, ...sheets]
  const entries = ordered.map((path) => ({ path, data: parts.get(path)! }))
  return new Uint8Array(await new Response(zipStream(entries, { zip64 })).arrayBuffer())
}

/** Set bit 3 on every local header, leaving sizes and the directory as they are. */
function withBit3OnLocalHeaders(zip: Uint8Array): Uint8Array {
  const out = zip.slice()
  for (const e of directory(zip)) out[e.localHeaderOffset + 6]! |= 0x08
  return out
}

/** Rewrite one central directory record in place. `at` is the record's offset. */
function patchDirectoryEntry(
  zip: Uint8Array,
  name: string,
  patch: (v: DataView, at: number) => void,
): Uint8Array {
  const out = zip.slice()
  const v = new DataView(out.buffer)
  const eocd = readEocd(out, 0)
  let pos = eocd.centralDirOffset
  for (let i = 0; i < eocd.entryCount; i++) {
    const nameLen = v.getUint16(pos + 28, true)
    const entryName = new TextDecoder().decode(out.subarray(pos + 46, pos + 46 + nameLen))
    if (entryName === name) {
      patch(v, pos)
      return out
    }
    pos += 46 + nameLen + v.getUint16(pos + 30, true) + v.getUint16(pos + 32, true)
  }
  throw new Error(`no ${name} in directory`)
}

function directory(bytes: Uint8Array): CentralDirEntry[] {
  const eocd = readEocd(bytes, 0)
  let bounds: { entryCount: number; centralDirOffset: number } = eocd
  if (eocd.needsZip64) {
    const at = readZip64Locator(bytes.subarray(eocd.eocdOffset - 20, eocd.eocdOffset))
    bounds = readZip64Eocd(bytes.subarray(at, at + 56))
  }
  return readCentralDirectory(bytes, bounds.centralDirOffset, bounds.entryCount)
}

/** Where an entry's body sits, from its local header. */
function bodyRange(bytes: Uint8Array, e: CentralDirEntry): [number, number] {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const start = e.localHeaderOffset + 30 + v.getUint16(e.localHeaderOffset + 26, true)
  return [start, start + v.getUint16(e.localHeaderOffset + 28, true) + e.compressedSize]
}

// ── Parity on files hucre did not write ──────────────────────────────

const fixtureFiles = [
  ...readdirSync(FIXTURES)
    .filter((f) => f.endsWith(".xlsx"))
    .map((f) => join(FIXTURES, f)),
  ...readdirSync(join(FIXTURES, "third-party"))
    .filter((f) => f.endsWith(".xlsx"))
    .map((f) => join(FIXTURES, "third-party", f)),
]

describe("streamXlsxRows yields the same rows from every input and layout", () => {
  it("has Excel, openpyxl and ExcelJS fixtures to compare", () => {
    const names = fixtureFiles.map((f) => f.split(/[/\\]/).pop()!)
    expect(names.some((n) => n.startsWith("excel-"))).toBe(true)
    expect(names.some((n) => n.startsWith("openpyxl-"))).toBe(true)
    expect(names.some((n) => !n.startsWith("excel-") && !n.startsWith("openpyxl-"))).toBe(true)
  })

  for (const file of fixtureFiles) {
    const name = file.split(/[/\\]/).pop()!
    it(name, async () => {
      const original = new Uint8Array(readFileSync(file))
      // `sparse`, because excel-sparse.xlsx is over the dense bounding-box limit.
      const sheetCount = (await readXlsx(original, { sparse: true })).sheets.length
      const variants: Array<
        [string, () => Promise<RandomAccessInput | ReadableStream<Uint8Array>>]
      > = [
        ["random access, as written", async () => counting(original)],
        ["random access, Blob", async () => randomAccessFromBlob(new Blob([original]))],
        [
          "random access, sheets first + descriptors",
          async () => counting(await repack(original, "sheets-first")),
        ],
        [
          "random access, sheets first + ZIP64",
          async () => counting(await repack(original, "sheets-first", true)),
        ],
        [
          "stream, metadata first + descriptors",
          async () => pulledStream(await repack(original, "metadata-first")).stream,
        ],
        [
          "stream, metadata first + ZIP64",
          async () => pulledStream(await repack(original, "metadata-first", true)).stream,
        ],
      ]
      for (let sheet = 0; sheet < Math.max(sheetCount, 1); sheet++) {
        const expected = await collect(streamXlsxRows(original, { sheet }))
        for (const [label, make] of variants) {
          const got = await collect(streamXlsxRows(await make(), { sheet }))
          expect(got, `${name} sheet ${sheet}: ${label}`).toEqual(expected)
        }
      }
    })
  }

  it("works with the README's adapters: a FileHandle, and fs.openAsBlob", async () => {
    const path = join(FIXTURES, "excel-styled.xlsx")
    const expected = await collect(streamXlsxRows(new Uint8Array(readFileSync(path))))

    // The README's FileHandle adapter, verbatim apart from the loop body.
    const handle = await open(path)
    try {
      const source: RandomAccessInput = {
        size: (await handle.stat()).size,
        async read(offset: number, length: number) {
          const buf = new Uint8Array(length)
          let done = 0
          while (done < length) {
            const { bytesRead } = await handle.read(buf, done, length - done, offset + done)
            if (bytesRead === 0) break
            done += bytesRead
          }
          return buf.subarray(0, done)
        },
      }
      expect(await collect(streamXlsxRows(source))).toEqual(expected)
    } finally {
      await handle.close()
    }

    expect(await collect(streamXlsxRows(randomAccessFromBlob(await openAsBlob(path))))).toEqual(
      expected,
    )
  })

  it("applies range, maxRows, dates and dateSystem the same way", async () => {
    const original = new Uint8Array(readFileSync(join(FIXTURES, "excel-basic.xlsx")))
    const sheetsFirst = await repack(original, "sheets-first")
    for (const options of [
      { maxRows: 2 },
      { range: "B2:C4" },
      { dates: "serial" as const },
      { dateSystem: "1904" as const },
    ]) {
      expect(await collect(streamXlsxRows(counting(sheetsFirst), options))).toEqual(
        await collect(streamXlsxRows(original, options)),
      )
    }
  })
})

// ── Byte counts ──────────────────────────────────────────────────────

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

/** xorshift32: values DEFLATE cannot shrink much, so a sheet spans many windows. */
function rng(seed: number): () => number {
  let x = seed >>> 0 || 1
  return () => {
    x ^= x << 13
    x ^= x >>> 17
    x ^= x << 5
    return (x >>> 0) / 0x100000000
  }
}

/**
 * A worksheet of `rows` rows: a shared string, a date (style 1) and four
 * noisy numbers per row — the shape of a cycler's per-sample record.
 */
function sheetXml(rows: number, seed: number): string {
  const next = rng(seed)
  const out: string[] = [`<?xml version="1.0"?><worksheet ${NS} ${R}><sheetData>`]
  for (let r = 1; r <= rows; r++) {
    out.push(
      `<row r="${r}"><c r="A${r}" t="s"><v>${r % 3}</v></c><c r="B${r}" s="1"><v>${45000 + r / 86400}</v></c>` +
        `<c r="C${r}"><v>${next()}</v></c><c r="D${r}"><v>${next()}</v></c>` +
        `<c r="E${r}"><v>${next()}</v></c><c r="F${r}"><v>${next()}</v></c></row>`,
    )
  }
  out.push("</sheetData></worksheet>")
  return out.join("")
}

/**
 * Parts for a two-sheet workbook, in the order given. Sheet 2 is the
 * decoy: as large as sheet 1, and never needed when reading sheet 1.
 */
function workbookParts(rows: number): Record<string, string> {
  return {
    "xl/worksheets/sheet1.xml": sheetXml(rows, 1),
    "xl/worksheets/sheet2.xml": sheetXml(rows, 2),
    "xl/sharedStrings.xml":
      `<?xml version="1.0"?><sst ${NS} count="3" uniqueCount="3">` +
      `<si><t>charge</t></si><si><t>rest</t></si><si><t>discharge</t></si></sst>`,
    "xl/styles.xml":
      `<?xml version="1.0"?><styleSheet ${NS}><cellXfs count="2">` +
      `<xf numFmtId="0"/><xf numFmtId="22" applyNumberFormat="1"/></cellXfs></styleSheet>`,
    "xl/workbook.xml":
      `<?xml version="1.0"?><workbook ${NS} ${R}><sheets>` +
      `<sheet name="Record" sheetId="1" r:id="rId1"/><sheet name="Decoy" sheetId="2" r:id="rId2"/>` +
      `</sheets></workbook>`,
    "xl/_rels/workbook.xml.rels":
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/>` +
      `<Relationship Id="rId3" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>` +
      `<Relationship Id="rId4" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
    "_rels/.rels":
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "[Content_Types].xml":
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/></Types>`,
  }
}

async function zipParts(
  parts: Record<string, string>,
  order: string[],
  zip64 = false,
): Promise<Uint8Array> {
  const entries = order.map((path) => ({ path, data: enc.encode(parts[path]!) }))
  return new Uint8Array(await new Response(zipStream(entries, { zip64 })).arrayBuffer())
}

// ~6 MB of compressed sheet each, so "under half the sheet before the
// first row" leaves several windows of margin rather than one.
const ROWS = 100_000
// These generate and parse a few hundred thousand rows; under coverage
// instrumentation that is well past vitest's 5 s default.
const HEAVY = { timeout: 60_000 }
const parts = workbookParts(ROWS)
const SHEETS_FIRST = Object.keys(parts) // as declared: sheets, then metadata, content types last
const METADATA_FIRST = [...SHEETS_FIRST.slice(2), ...SHEETS_FIRST.slice(0, 2)]

describe("random access reads only what the sheet needs", HEAVY, () => {
  it("on an Excel-ordered workbook with data descriptors, never touches the other sheet", async () => {
    const bytes = await zipParts(parts, SHEETS_FIRST)
    const dir = directory(bytes)
    const byName = new Map(dir.map((e) => [e.fileName, e]))
    const sheet1 = byName.get("xl/worksheets/sheet1.xml")!
    const sheet2 = byName.get("xl/worksheets/sheet2.xml")!
    // The premise: data descriptors on, the sheet first, and the sheet
    // spanning many windows.
    expect(dir.every((e) => e.hasDataDescriptor)).toBe(true)
    expect(dir[0]!.fileName).toBe("xl/worksheets/sheet1.xml")
    expect(sheet1.compressedSize).toBeGreaterThan(RANDOM_ACCESS_WINDOW_BYTES * 2)

    const src = counting(bytes)
    const gen = streamXlsxRows(src)
    const first = await gen.next()
    expect(first.value).toEqual({
      index: 0,
      values: [
        "rest",
        expect.any(Date),
        expect.any(Number),
        expect.any(Number),
        expect.any(Number),
        expect.any(Number),
      ],
    })
    // Before the first row: the tail, the directory, the metadata parts
    // and at most a couple of windows of the sheet — not the sheet.
    expect(src.requested()).toBeLessThan(sheet1.compressedSize / 2)

    let count = 1
    for await (const _ of gen) count++
    expect(count).toBe(ROWS)

    for (const [, len] of src.reads) expect(len).toBeLessThanOrEqual(RANDOM_ACCESS_WINDOW_BYTES)

    // One read is the tail that finds the EOCD — the last 64 KiB, which
    // in this order ends inside the decoy. Every other read is the
    // encryption sniff or a part the sheet needs, and none touches it.
    const [decoyStart, decoyEnd] = bodyRange(bytes, sheet2)
    const tailOffset = Math.max(0, bytes.length - 65557 - 20)
    const tailReads = src.reads.filter(([offset]) => offset === tailOffset)
    expect(tailReads).toEqual([[tailOffset, bytes.length - tailOffset]])
    for (const [offset, len] of src.reads.filter(([o]) => o !== tailOffset)) {
      const overlaps = offset < decoyEnd && offset + len > decoyStart
      expect(overlaps, `read ${offset}+${len} touches the decoy sheet`).toBe(false)
    }

    // Everything asked for: the needed entries (header and body), the
    // directory and the tail read that found it. Nothing else.
    const needed = [
      "xl/worksheets/sheet1.xml",
      "xl/sharedStrings.xml",
      "xl/styles.xml",
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels",
      "_rels/.rels",
      "[Content_Types].xml",
    ]
    const neededBytes = needed.reduce((n, p) => n + byName.get(p)!.compressedSize + 30 * 2, 0)
    const tailAndDirectory = Math.min(bytes.length, 65557 + 20)
    expect(src.requested()).toBeLessThanOrEqual(neededBytes + tailAndDirectory + 4096)
    expect(src.requested()).toBeLessThan(bytes.length - sheet2.compressedSize / 2)
  })

  it("does the same with ZIP64 records", async () => {
    const bytes = await zipParts(parts, SHEETS_FIRST, true)
    const src = counting(bytes)
    const gen = streamXlsxRows(src, { sheet: "Decoy" })
    await gen.next()
    const sheet2 = directory(bytes).find((e) => e.fileName === "xl/worksheets/sheet2.xml")!
    expect(src.requested()).toBeLessThan(sheet2.compressedSize / 2)
    await gen.return()
  })
})

describe("a ReadableStream with data descriptors streams when metadata comes first", HEAVY, () => {
  it("pulls well under the file before the first row", async () => {
    const bytes = await zipParts(parts, METADATA_FIRST)
    expect(directory(bytes).every((e) => e.hasDataDescriptor)).toBe(true)

    const { stream, pulled } = pulledStream(bytes)
    const gen = streamXlsxRows(stream)
    const first = await gen.next()
    expect(first.value?.index).toBe(0)
    // The buffered fallback reads the whole file before yielding anything.
    expect(pulled()).toBeLessThan(bytes.length / 4)

    let count = 1
    for await (const _ of gen) count++
    expect(count).toBe(ROWS)
    // Sheet 1 is the first of two equal sheets, and nothing past it was needed.
    expect(pulled()).toBeLessThan(bytes.length * 0.75)
  })

  it("and with the ZIP64 descriptor form", async () => {
    const bytes = await zipParts(parts, METADATA_FIRST, true)
    const { stream, pulled } = pulledStream(bytes)
    const gen = streamXlsxRows(stream)
    await gen.next()
    expect(pulled()).toBeLessThan(bytes.length / 4)
    await gen.return()
  })

  it("yields what the buffered path yields", async () => {
    const small = workbookParts(500)
    const bytes = await zipParts(small, METADATA_FIRST)
    for (const sheet of [0, 1]) {
      expect(await collect(streamXlsxRows(pulledStream(bytes).stream, { sheet }))).toEqual(
        await collect(streamXlsxRows(bytes, { sheet })),
      )
    }
  })

  it("still buffers when the sheet comes before what it needs", async () => {
    const bytes = await zipParts(workbookParts(500), SHEETS_FIRST)
    const { stream, pulled } = pulledStream(bytes)
    const gen = streamXlsxRows(stream)
    await gen.next()
    expect(pulled()).toBe(bytes.length)
    await gen.return()
  })
})

describe('a runtime without DecompressionStream("deflate-raw")', () => {
  it("still streams the same rows from a random-access source and a stream", async () => {
    // The capability is memoized per module, so the probe runs in a fresh
    // module registry with the constructor refusing the format.
    const bytes = await zipParts(workbookParts(300), METADATA_FIRST)
    const expected = await collect(streamXlsxRows(bytes))
    class NoRaw {
      constructor(format: string) {
        throw new TypeError(`unsupported format ${format}`)
      }
    }
    vi.stubGlobal("DecompressionStream", NoRaw)
    try {
      vi.resetModules()
      const fresh = await import("../src/xlsx/stream-reader")
      expect(await collect(fresh.streamXlsxRows(counting(bytes)))).toEqual(expected)
      expect(await collect(fresh.streamXlsxRows(pulledStream(bytes).stream))).toEqual(expected)
    } finally {
      vi.unstubAllGlobals()
      vi.resetModules()
    }
  })
})

// ── Limits and fallbacks on the random-access path ───────────────────

describe("random access keeps the readers' limits", () => {
  it("honours maxDecompressedBytes on the streamed sheet", async () => {
    const bytes = await zipParts(workbookParts(2000), SHEETS_FIRST)
    await expect(
      collect(streamXlsxRows(counting(bytes), { maxDecompressedBytes: 10_000 })),
    ).rejects.toThrow(/possible zip bomb/)
    await expect(collect(streamXlsxRows(bytes, { maxDecompressedBytes: 10_000 }))).rejects.toThrow(
      /possible zip bomb/,
    )
  })

  it("honours maxDecompressedBytes on a streamed ReadableStream, which used to ignore it", async () => {
    // Sizes in the local headers and metadata first: the layout the stream
    // path always streamed, with MAX_DECOMPRESSED_BYTES in place of the
    // caller's limit. (A descriptor file used to fall back to the buffered
    // reader, which did honour it — so it would not show the difference.)
    const small = workbookParts(2000)
    const writer = new ZipWriter()
    for (const path of METADATA_FIRST) writer.add(path, enc.encode(small[path]!))
    const bytes = await writer.build()
    await expect(
      collect(streamXlsxRows(pulledStream(bytes).stream, { maxDecompressedBytes: 10_000 })),
    ).rejects.toThrow(/possible zip bomb/)
  })

  it("counts what it holds against maxInputBytes, not the size of the file", async () => {
    const bytes = await zipParts(parts, SHEETS_FIRST)
    // Far below the file, well above the directory and metadata.
    const rows = await collect(
      streamXlsxRows(counting(bytes), { maxInputBytes: 200_000, maxRows: 3 }),
    )
    expect(rows).toHaveLength(3)
    // And below the tail read, which must be held to find the directory.
    await expect(collect(streamXlsxRows(counting(bytes), { maxInputBytes: 1000 }))).rejects.toThrow(
      ParseError,
    )
  })

  it("reads an encrypted workbook whole, and reports it as encrypted", async () => {
    const ole = new Uint8Array(4096)
    ole.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
    await expect(collect(streamXlsxRows(counting(ole)))).rejects.toThrow(EncryptedFileError)
    await expect(collect(streamXlsxRows(counting(ole), { maxInputBytes: 1000 }))).rejects.toThrow(
      /Raise `maxInputBytes`/,
    )
  })

  it("requires [Content_Types].xml on every input, as read() does", async () => {
    // Validation, not resolution — but a package the buffered reader
    // rejects must not stream without complaint from a different input.
    const noTypes = { ...workbookParts(5) }
    delete (noTypes as Record<string, string>)["[Content_Types].xml"]
    const order = METADATA_FIRST.filter((p) => p !== "[Content_Types].xml")
    const bytes = await zipParts(noTypes, order)
    const message = /missing \[Content_Types\]\.xml/
    await expect(collect(streamXlsxRows(bytes))).rejects.toThrow(message)
    await expect(collect(streamXlsxRows(counting(bytes)))).rejects.toThrow(message)
    await expect(collect(streamXlsxRows(pulledStream(bytes).stream))).rejects.toThrow(message)
  })

  it("holds a ReadableStream to maxInputBytes while scanning for a descriptor", async () => {
    // Bit 3 set on every local header, sizes left in place, no descriptor
    // written. The descriptor scan reads to the end of the stream before
    // the walk falls back, so the cap has to apply as the fallback log
    // grows. It never fully did: drainToBuffer counted only the chunks it
    // pulled itself, so whatever was already logged — here, with one
    // chunk from Blob.stream(), all of it — escaped the cap on main too.
    const small = workbookParts(2000)
    const writer = new ZipWriter()
    for (const path of METADATA_FIRST) writer.add(path, enc.encode(small[path]!))
    const bytes = withBit3OnLocalHeaders(await writer.build())
    // Without the cap it still reads, through the buffered reader.
    expect(await collect(streamXlsxRows(pulledStream(bytes).stream))).toHaveLength(2000)
    await expect(
      collect(streamXlsxRows(pulledStream(bytes).stream, { maxInputBytes: 10_000 })),
    ).rejects.toThrow(/Input stream exceeds the maximum of 10000 bytes/)
  })

  it("reads whole, through the buffered reader, an entry the directory declines", async () => {
    const small = workbookParts(50)
    const writer = new ZipWriter()
    for (const path of SHEETS_FIRST) writer.add(path, enc.encode(small[path]!))
    const bytes = await writer.build()
    const expected = await collect(streamXlsxRows(bytes))

    // No compressed size in the directory, bit 3 set: ZipReader recovers
    // it from the local header; the random-access reader declines.
    const noSize = patchDirectoryEntry(bytes, "xl/worksheets/sheet1.xml", (v, at) => {
      v.setUint16(at + 8, v.getUint16(at + 8, true) | 0x08, true)
      v.setUint32(at + 20, 0, true)
    })
    const src = counting(noSize)
    expect(await collect(streamXlsxRows(src))).toEqual(expected)
    expect(src.reads.some(([o, n]) => o === 0 && n === noSize.length)).toBe(true)

    // An unsupported method: the buffered reader's error, word for word.
    const method = patchDirectoryEntry(bytes, "xl/styles.xml", (v, at) =>
      v.setUint16(at + 10, 99, true),
    )
    const buffered = await collect(streamXlsxRows(method)).catch((e: Error) => e.message)
    expect(buffered).toMatch(/Unsupported compression method 99/)
    expect(await collect(streamXlsxRows(counting(method))).catch((e: Error) => e.message)).toBe(
      buffered,
    )
  })

  it("gives a non-ZIP the buffered reader's error", async () => {
    const junk = enc.encode("not a zip file at all, just some text that goes on for a while")
    const viaBuffer = await collect(streamXlsxRows(junk)).catch((e: unknown) => e)
    const viaSource = await collect(streamXlsxRows(counting(junk))).catch((e: unknown) => e)
    expect(viaBuffer).toBeInstanceOf(ZipError)
    expect((viaSource as Error).message).toBe((viaBuffer as Error).message)
  })

  it("refuses a hostile directory in a huge source without a large read", async () => {
    // A 2^50-byte source of zeros whose ZIP64 end records claim a 4 GiB
    // directory. Too big to read whole, too big to hold the directory of.
    const size = 2 ** 50
    const tail = new Uint8Array(98)
    const v = new DataView(tail.buffer)
    v.setUint32(0, 0x06064b50, true)
    v.setBigUint64(32, 10n, true)
    v.setBigUint64(40, 2n ** 32n, true)
    v.setBigUint64(48, BigInt(size - 98 - 2 ** 32), true)
    v.setUint32(56, 0x07064b50, true)
    v.setBigUint64(64, BigInt(size - 98), true)
    v.setUint32(76, 0x06054b50, true)
    v.setUint16(84, 0xffff, true)
    v.setUint16(86, 0xffff, true)
    v.setUint32(88, 0xffffffff, true)
    v.setUint32(92, 0xffffffff, true)
    const reads: number[] = []
    const src: RandomAccessInput = {
      size,
      async read(offset, length) {
        reads.push(length)
        const out = new Uint8Array(length)
        const base = size - 98
        const from = Math.max(offset, base)
        const to = Math.min(offset + length, size)
        if (from < to) out.set(tail.subarray(from - base, to - base), from - offset)
        return out
      },
    }
    await expect(collect(streamXlsxRows(src))).rejects.toThrow(/exceeds the limit/)
    expect(Math.max(...reads)).toBeLessThanOrEqual(RANDOM_ACCESS_WINDOW_BYTES)
    expect(reads.length).toBeLessThanOrEqual(3)
  })
})

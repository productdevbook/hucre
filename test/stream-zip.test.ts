import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import { ZipStreamReader } from "../src/zip/stream-reader"
import { ZipWriter } from "../src/zip/writer"
import type { CellValue } from "../src/_types"

const encoder = new TextEncoder()
const decoder = new TextDecoder("utf-8")

/** Wrap bytes in a ReadableStream that emits fixed-size chunks and counts
 *  how many bytes have actually been pulled by the consumer. */
function chunkedStream(
  data: Uint8Array,
  chunkSize: number,
): { stream: ReadableStream<Uint8Array>; pulled: () => number } {
  let offset = 0
  let pulledBytes = 0
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) {
      if (offset >= data.length) {
        controller.close()
        return
      }
      const end = Math.min(offset + chunkSize, data.length)
      const chunk = data.subarray(offset, end)
      offset = end
      pulledBytes += chunk.length
      controller.enqueue(chunk)
    },
  })
  return { stream, pulled: () => pulledBytes }
}

async function collect(gen: AsyncGenerator<{ index: number; values: CellValue[] }>) {
  const rows: CellValue[][] = []
  for await (const row of gen) rows.push(row.values)
  return rows
}

describe("ZipStreamReader (local-header streaming)", () => {
  it("iterates entries and decompresses bodies matching the source", async () => {
    const zw = new ZipWriter()
    zw.add("a.txt", encoder.encode("hello world ".repeat(50)))
    zw.add("dir/b.xml", encoder.encode("<root>" + "x".repeat(2000) + "</root>"))
    zw.add("c.bin", new Uint8Array([1, 2, 3, 4, 5]))
    const zipped = await zw.build()

    const zr = new ZipStreamReader(chunkedStream(zipped, 7).stream)
    const seen = new Map<string, string>()
    for (;;) {
      const entry = await zr.nextEntry()
      if (!entry) break
      expect(entry.streamable).toBe(true)
      const bytes = await zr.readEntryBytes(entry)
      seen.set(entry.name, decoder.decode(bytes))
    }
    expect(seen.get("a.txt")).toBe("hello world ".repeat(50))
    expect(seen.get("dir/b.xml")).toBe("<root>" + "x".repeat(2000) + "</root>")
    expect(seen.get("c.bin")).toBe(decoder.decode(new Uint8Array([1, 2, 3, 4, 5])))
  })

  it("skipEntry advances past a body without decoding it", async () => {
    const zw = new ZipWriter()
    zw.add("skip.txt", encoder.encode("ignore me ".repeat(100)))
    zw.add("keep.txt", encoder.encode("KEEP"))
    const zr = new ZipStreamReader(chunkedStream(await zw.build(), 13).stream)

    const first = await zr.nextEntry()
    expect(first!.name).toBe("skip.txt")
    await zr.skipEntry()
    const second = await zr.nextEntry()
    expect(second!.name).toBe("keep.txt")
    expect(decoder.decode(await zr.readEntryBytes(second!))).toBe("KEEP")
  })
})

describe("streamXlsxRows over ReadableStream — true streaming (#77)", () => {
  function bookBytes() {
    return writeXlsx({
      sheets: [
        {
          name: "First",
          rows: [
            ["Name", "Qty"],
            ["alpha", 1],
            ["beta", 2],
            ["gamma", 3],
          ],
        },
        {
          name: "Second",
          rows: [
            ["City", "Pop"],
            ["NYC", 8000000],
            ["LA", 4000000],
          ],
        },
      ],
    })
  }

  it("yields identical rows to the buffered path, across tiny chunks", async () => {
    const bytes = await bookBytes()
    const buffered = await collect(streamXlsxRows(bytes))
    const { stream } = chunkedStream(bytes, 16)
    const streamed = await collect(streamXlsxRows(stream))
    expect(streamed).toEqual(buffered)
  })

  it("streams a non-first target sheet (skips earlier worksheets)", async () => {
    const bytes = await bookBytes()
    const { stream } = chunkedStream(bytes, 64)
    const rows = await collect(streamXlsxRows(stream, { sheet: "Second" }))
    expect(rows[0]).toEqual(["City", "Pop"])
    expect(rows[1]).toEqual(["NYC", 8000000])
  })

  it("honors range + maxRows over a streamed archive", async () => {
    const rows: CellValue[][] = []
    for (let i = 0; i < 50; i++) rows.push([`r${i}`, i])
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows }] })
    const { stream } = chunkedStream(bytes, 128)
    const out = await collect(streamXlsxRows(stream, { maxRows: 5 }))
    expect(out.length).toBe(5)
  })

  it("does NOT buffer the whole archive: a huge trailing sheet is never pulled", async () => {
    // Target the small FIRST sheet; a huge numeric SECOND sheet follows it
    // in the archive. Numeric cells don't go through shared strings, so the
    // second worksheet body dominates the file. A true streaming reader
    // finishes the first sheet and stops — never pulling the second.
    const small: CellValue[][] = [
      ["a", "b"],
      ["x", 1],
      ["y", 2],
    ]
    const huge: CellValue[][] = []
    for (let i = 0; i < 40000; i++) huge.push([i, i * 2, i * 3, i * 4])
    const bytes = await writeXlsx({
      sheets: [
        { name: "First", rows: small },
        { name: "Second", rows: huge },
      ],
    })

    const { stream, pulled } = chunkedStream(bytes, 4096)
    const rows = await collect(streamXlsxRows(stream, { sheet: "First" }))
    expect(rows[0]).toEqual(["a", "b"])
    // We consumed the entire target sheet, yet pulled far less than the
    // whole archive — the trailing huge sheet was never read.
    expect(pulled()).toBeLessThan(bytes.length / 2)
  })
})

import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeOds } from "../src/ods/writer"
import { streamOdsRows } from "../src/ods/stream"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import { streamCsvRows } from "../src/csv/stream"
import { writeXlsxStream } from "../src/xlsx/stream-writer"
import type { CellValue, StreamRow, WriteSheet } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  let total = 0
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
    total += value.length
  }
  const out = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    out.set(chunk, offset)
    offset += chunk.length
  }
  return out
}

async function drain<T>(gen: AsyncGenerator<T, void, undefined>): Promise<T[]> {
  const rows: T[] = []
  for await (const row of gen) rows.push(row)
  return rows
}

function streamOf(bytes: Uint8Array): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(bytes)
      controller.close()
    },
  })
}

const twoSheets: WriteSheet[] = [
  { name: "First", rows: [["a1"], ["a2"]] },
  { name: "Second", rows: [["b1"], ["b2"], ["b3"]] },
]

// ═══════════════════════════════════════════════════════════════════════
// #365 — the streaming readers had three unrelated contracts, and the
// one write helper took its arguments in the opposite order from every
// other writer.
// ═══════════════════════════════════════════════════════════════════════

describe("writeXlsxStream argument order", () => {
  it("takes (rows, options), like every other write* in the library", async () => {
    const bytes = await collect(
      writeXlsxStream(
        [
          [1, "Alice"],
          [2, "Bob"],
        ],
        { name: "Data", columns: [{ header: "ID" }, { header: "Name" }] },
      ),
    )

    const workbook = await readXlsx(bytes)
    expect(workbook.sheets[0].rows).toEqual([
      ["ID", "Name"],
      [1, "Alice"],
      [2, "Bob"],
    ])
  })
})

describe("StreamRow is one shape", () => {
  it("XLSX rows carry index and values", async () => {
    const buf = await writeXlsx({ sheets: [twoSheets[0]!] })
    const rows: StreamRow[] = await drain(streamXlsxRows(buf))

    expect(rows.map((r) => r.index)).toEqual([0, 1])
    expect(rows[0]!.values).toEqual(["a1"])
  })

  it("ODS rows carry the same fields plus sheetIndex", async () => {
    const buf = await writeOds({ sheets: twoSheets })
    const rows: StreamRow[] = await drain(streamOdsRows(buf))

    // Same property names as the XLSX reader — they used to be two
    // separate interfaces describing the same thing.
    for (const row of rows) {
      expect(row).toHaveProperty("index")
      expect(row).toHaveProperty("values")
    }
    expect(new Set(rows.map((r) => r.sheetIndex))).toEqual(new Set([0, 1]))
  })

  it("CSV still yields a bare array, on purpose", () => {
    // CSV rows are dense and positional, so an index would be ceremony —
    // and the bare array is what keeps streamCsvRows the streaming
    // mirror of parseCsv.
    const rows = Array.from(streamCsvRows("a,b\r\nc,d\r\n"))
    expect(rows).toEqual([
      ["a", "b"],
      ["c", "d"],
    ])
  })
})

describe("streamOdsRows input and options", () => {
  it("accepts a ReadableStream, like streamXlsxRows", async () => {
    const buf = await writeOds({ sheets: [twoSheets[0]!] })
    const rows = await drain(streamOdsRows(streamOf(buf)))
    expect(rows.map((r) => r.values[0])).toEqual(["a1", "a2"])
  })

  it("still accepts a Uint8Array", async () => {
    const buf = await writeOds({ sheets: [twoSheets[0]!] })
    const rows = await drain(streamOdsRows(buf))
    expect(rows).toHaveLength(2)
  })

  it("honours maxRows", async () => {
    // It previously took no options at all, so bounding a huge file meant
    // draining it and counting yourself.
    const buf = await writeOds({ sheets: twoSheets })
    const rows = await drain(streamOdsRows(buf, { maxRows: 3 }))
    expect(rows).toHaveLength(3)
  })

  it("honours a numeric sheet filter", async () => {
    const buf = await writeOds({ sheets: twoSheets })
    const rows = await drain(streamOdsRows(buf, { sheets: [1] }))

    expect(rows.every((r) => r.sheetIndex === 1)).toBe(true)
    expect(rows.map((r) => r.values[0])).toEqual(["b1", "b2", "b3"])
  })

  it("streams everything when the filter names no resolvable sheet", async () => {
    // The SAX pass does not surface table names, so a name-only filter
    // cannot be resolved. Streaming everything is the safe fallback —
    // yielding nothing would look like an empty file.
    const buf = await writeOds({ sheets: twoSheets })
    const rows = await drain(streamOdsRows(buf, { sheets: ["Second"] }))
    expect(rows.length).toBe(5)
  })

  it("combines the filters", async () => {
    const buf = await writeOds({ sheets: twoSheets })
    const rows = await drain(streamOdsRows(buf, { sheets: [1], maxRows: 2 }))
    expect(rows).toHaveLength(2)
    expect(rows.every((r) => r.sheetIndex === 1)).toBe(true)
  })
})

describe("row values survive the contract change", () => {
  it("ODS streaming still matches the batch reader", async () => {
    const rows: CellValue[][] = [
      ["text", 42],
      [true, null],
    ]
    const buf = await writeOds({ sheets: [{ name: "S", rows }] })

    const streamed = (await drain(streamOdsRows(buf))).map((r) => r.values)
    expect(streamed[0]![0]).toBe("text")
    expect(streamed[0]![1]).toBe(42)
    expect(streamed[1]![0]).toBe(true)
  })
})

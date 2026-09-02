import { describe, expect, it } from "vitest"
import { writeCsvStream, CsvStreamWriter } from "../src/csv/stream"
import { writeCsv, writeCsvObjects } from "../src/csv/writer"
import { parseCsv } from "../src/csv/reader"
import { NdjsonStreamWriter } from "../src/json/stream"
import type { CellValue, CsvWriteOptions } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

async function collect(stream: ReadableStream<Uint8Array>): Promise<string> {
  // ignoreBOM, or the decoder silently eats the BOM we are asserting on.
  const dec = new TextDecoder("utf-8", { ignoreBOM: true })
  const reader = stream.getReader()
  let out = ""
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    out += dec.decode(value, { stream: true })
  }
  out += dec.decode()
  return out
}

async function collectChunks(stream: ReadableStream<Uint8Array>): Promise<Uint8Array[]> {
  const chunks: Uint8Array[] = []
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
  }
  return chunks
}

async function* asyncRows<T>(rows: T[]): AsyncGenerator<T> {
  for (const row of rows) {
    await Promise.resolve()
    yield row
  }
}

/** Same rows through the buffered writer, for parity assertions. */
function buffered(rows: CellValue[][], options?: CsvWriteOptions): string {
  const writer = new CsvStreamWriter(options)
  for (const row of rows) writer.addRow(row)
  return writer.finishText()
}

// ═══════════════════════════════════════════════════════════════════════

describe("writeCsvStream", () => {
  it("matches writeCsv byte for byte", async () => {
    const rows: CellValue[][] = [
      ["id", "name"],
      [1, "Alice"],
      [2, "Bob"],
    ]
    expect(await collect(writeCsvStream(rows))).toBe(writeCsv(rows))
  })

  it("matches the buffered writer across option combinations", async () => {
    const rows: CellValue[][] = [
      [1, "plain"],
      [2, 'has "quotes"'],
      [3, "has,comma"],
      [4, "has\nnewline"],
      [5, null],
      [6, true],
      [7, new Date(Date.UTC(2024, 0, 15, 12, 30))],
      [8, 1e-9],
      [9, 1e16],
    ]

    const cases: CsvWriteOptions[] = [
      {},
      { delimiter: ";" },
      { lineSeparator: "\n" },
      { quoteStyle: "all" },
      { quoteStyle: "none" },
      { quote: "'" },
      { nullValue: "NULL" },
      { bom: true },
      { dateFormat: "YYYY-MM-DD HH:mm:ss" },
      { headers: ["a", "b"] },
    ]

    for (const options of cases) {
      expect(await collect(writeCsvStream(rows, options)), JSON.stringify(options)).toBe(
        buffered(rows, options),
      )
    }
  })

  it("accepts an async row source", async () => {
    const out = await collect(
      writeCsvStream(
        asyncRows([
          [1, "a"],
          [2, "b"],
        ]),
      ),
    )
    expect(out).toBe("1,a\r\n2,b")
  })

  it("emits the BOM before the first line", async () => {
    const out = await collect(writeCsvStream([[1]], { bom: true }))
    expect(out.charCodeAt(0)).toBe(0xfeff)
    expect(out).toBe("﻿1")
  })

  it("writes an explicit header array even when no rows arrive", async () => {
    expect(await collect(writeCsvStream([], { headers: ["id", "name"] }))).toBe("id,name")
  })

  it("emits nothing for an empty source", async () => {
    expect(await collect(writeCsvStream([]))).toBe("")
  })

  it("emits only the BOM for an empty source when asked", async () => {
    expect(await collect(writeCsvStream([], { bom: true }))).toBe("﻿")
  })

  describe("object rows", () => {
    const data = [
      { id: 1, name: "Alice" },
      { id: 2, name: "Bob" },
    ]

    it("derives the header from the first row's keys", async () => {
      expect(await collect(writeCsvStream(data))).toBe(writeCsvObjects(data))
    })

    it("honours an explicit column order", async () => {
      const options: CsvWriteOptions = { columns: ["name", "id"] }
      expect(await collect(writeCsvStream(data, options))).toBe(writeCsvObjects(data, options))
    })

    it("skips the header line when headers is false", async () => {
      const out = await collect(writeCsvStream(data, { headers: false }))
      expect(out).toBe("1,Alice\r\n2,Bob")
    })

    it("fills missing keys with the null value", async () => {
      const out = await collect(
        writeCsvStream([{ id: 1, name: "Alice" }, { id: 2 }] as Array<Record<string, CellValue>>),
      )
      expect(parseCsv(out)).toEqual([
        ["id", "name"],
        ["1", "Alice"],
        ["2", ""],
      ])
    })
  })

  it("round-trips through parseCsv", async () => {
    const rows: CellValue[][] = [
      ["id", "note"],
      [1, 'quoted "value"'],
      [2, "with,comma"],
      [3, "multi\nline"],
    ]
    const out = await collect(writeCsvStream(rows))
    expect(parseCsv(out)).toEqual([
      ["id", "note"],
      ["1", 'quoted "value"'],
      ["2", "with,comma"],
      ["3", "multi\nline"],
    ])
  })

  it("streams incrementally rather than emitting one blob", async () => {
    const rows: CellValue[][] = []
    for (let i = 0; i < 40_000; i++) rows.push([i, `value-${i}`, i * 1.5])

    const chunks = await collectChunks(writeCsvStream(rows))
    // Output spans many 64 KB flushes; a buffered writer would emit one.
    expect(chunks.length).toBeGreaterThan(10)
  })

  it("pulls rows lazily and stops when the consumer cancels", async () => {
    let produced = 0
    function* infinite(): Generator<CellValue[]> {
      for (;;) {
        produced++
        yield [produced, "x".repeat(100)]
      }
    }

    const reader = writeCsvStream(infinite()).getReader()
    await reader.read()
    await reader.cancel()

    const afterCancel = produced
    await new Promise((resolve) => setTimeout(resolve, 10))
    expect(produced).toBe(afterCancel)
    expect(produced).toBeLessThan(100_000)
  })

  it("closes a generator source on cancel", async () => {
    let closed = false
    function* rows(): Generator<CellValue[]> {
      try {
        for (;;) yield [1]
      } finally {
        closed = true
      }
    }

    const reader = writeCsvStream(rows()).getReader()
    await reader.read()
    await reader.cancel()
    expect(closed).toBe(true)
  })
})

// ═══════════════════════════════════════════════════════════════════════

describe("NdjsonStreamWriter.toStream", () => {
  it("releases rows once they are enqueued", async () => {
    const writer = new NdjsonStreamWriter()
    for (let i = 0; i < 50; i++) writer.addObject({ i })

    const reader = writer.toStream().getReader()
    await reader.read()

    // Nothing already sent is still held — the stream is O(pending),
    // not O(total written).
    const internal = writer as unknown as { buffer: string[] }
    expect(internal.buffer).toHaveLength(0)
    await reader.cancel()
  })

  it("still delivers everything written before finish()", async () => {
    const writer = new NdjsonStreamWriter()
    writer.addObject({ a: 1 })
    writer.addObject({ a: 2 })
    writer.finish()

    const dec = new TextDecoder()
    const reader = writer.toStream().getReader()
    let out = ""
    for (;;) {
      const { done, value } = await reader.read()
      if (done) break
      out += dec.decode(value)
    }
    expect(out).toBe('{"a":1}\n{"a":2}\n')
  })

  it("delivers rows written after the stream is already being drained", async () => {
    const writer = new NdjsonStreamWriter()
    const dec = new TextDecoder()
    const reader = writer.toStream().getReader()

    writer.addObject({ a: 1 })
    const first = await reader.read()
    expect(dec.decode(first.value)).toBe('{"a":1}\n')

    writer.addObject({ a: 2 })
    const second = await reader.read()
    expect(dec.decode(second.value)).toBe('{"a":2}\n')

    writer.finish()
    await reader.cancel()
  })
})

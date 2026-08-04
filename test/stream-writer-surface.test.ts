import { describe, expect, it } from "vitest"
import { CsvStreamWriter } from "../src/csv/stream"
import { NdjsonStreamWriter } from "../src/json/stream"
import { XlsxStreamWriter } from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"
import { InvalidArgumentError } from "../src/errors"
import type { CellValue } from "../src/_types"

// ── One surface for the three stream writers (#365 item 7) ───────────
// Before v1 these had three different vocabularies — `addRow`/`addObject`,
// `addRow` alone, and `write`/`end` — so no format-agnostic export helper
// could be written against them. These tests pin the shared surface.

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const reader = stream.getReader()
  const chunks: Uint8Array[] = []
  let total = 0
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

describe("stream writer surface parity", () => {
  it("every writer exposes addRow / addObject / finish / toStream", () => {
    const writers = [
      new XlsxStreamWriter({ name: "S", columns: [{ key: "a", header: "A" }] }),
      new CsvStreamWriter(),
      new NdjsonStreamWriter({ columns: ["a"] }),
    ]

    for (const writer of writers) {
      for (const method of ["addRow", "addObject", "finish", "toStream"] as const) {
        expect(typeof (writer as unknown as Record<string, unknown>)[method]).toBe("function")
      }
    }
  })

  it("a format-agnostic export helper can be written against the shared surface", async () => {
    // The point of the whole convergence: one function, three formats.
    interface AnyStreamWriter {
      addObject(item: Record<string, CellValue>): void
      finish(): string | Promise<Uint8Array>
      toStream(): ReadableStream<Uint8Array>
    }

    async function exportAll(
      writer: AnyStreamWriter,
      rows: Record<string, CellValue>[],
    ): Promise<Uint8Array> {
      for (const row of rows) writer.addObject(row)
      // `finish()` marks the writer done; `toStream()` then yields the
      // bytes and closes. NdjsonStreamWriter's stream is a live drain, so
      // without the finish() it would stay open waiting for more rows.
      await writer.finish()
      return drain(writer.toStream())
    }

    const rows = [
      { id: 1, name: "alpha" },
      { id: 2, name: "beta" },
    ]

    const csv = new TextDecoder().decode(
      await exportAll(new CsvStreamWriter({ lineSeparator: "\n" }), rows),
    )
    expect(csv).toBe("id,name\n1,alpha\n2,beta")

    const ndjson = new TextDecoder().decode(await exportAll(new NdjsonStreamWriter(), rows))
    expect(ndjson).toBe('{"id":1,"name":"alpha"}\n{"id":2,"name":"beta"}\n')

    const xlsx = await exportAll(
      new XlsxStreamWriter({
        name: "S",
        columns: [
          { key: "id", header: "id" },
          { key: "name", header: "name" },
        ],
      }),
      rows,
    )
    const wb = await readXlsx(xlsx)
    expect(wb.sheets[0]!.rows).toEqual([
      ["id", "name"],
      [1, "alpha"],
      [2, "beta"],
    ])
  })
})

// ── CsvStreamWriter.addObject ────────────────────────────────────────

describe("CsvStreamWriter.addObject", () => {
  it("derives the column order from the first object and emits a header", () => {
    const writer = new CsvStreamWriter({ lineSeparator: "\n" })
    writer.addObject({ name: "Alice", age: 30 })
    writer.addObject({ name: "Bob", age: 25 })
    expect(writer.finish()).toBe("name,age\nAlice,30\nBob,25")
  })

  it("projects through an explicit `columns` order", () => {
    const writer = new CsvStreamWriter({ lineSeparator: "\n", columns: ["age", "name"] })
    writer.addObject({ name: "Alice", age: 30, extra: "dropped" })
    expect(writer.finish()).toBe("age,name\n30,Alice")
  })

  it("uses an explicit `headers` array as the column order without repeating it", () => {
    const writer = new CsvStreamWriter({ lineSeparator: "\n", headers: ["age", "name"] })
    writer.addObject({ name: "Alice", age: 30 })
    expect(writer.finish()).toBe("age,name\n30,Alice")
  })

  it("emits no header line when headers: false", () => {
    const writer = new CsvStreamWriter({ lineSeparator: "\n", headers: false })
    writer.addObject({ name: "Alice", age: 30 })
    expect(writer.finish()).toBe("Alice,30")
  })

  it("fills missing keys with the null representation", () => {
    const writer = new CsvStreamWriter({ lineSeparator: "\n" })
    writer.addObject({ a: 1, b: 2 })
    writer.addObject({ a: 3 })
    expect(writer.finish()).toBe("a,b\n1,2\n3,")
  })
})

// ── toStream() on the buffering writers ──────────────────────────────

describe("buffering writers' toStream()", () => {
  it("CsvStreamWriter.toStream emits exactly what finish() returns", async () => {
    const writer = new CsvStreamWriter({ lineSeparator: "\n" })
    writer.addRow(["a", "b"])
    writer.addRow([1, 2])

    const streamed = new TextDecoder().decode(await drain(writer.toStream()))
    expect(streamed).toBe(writer.finish())
    expect(streamed).toBe("a,b\n1,2")
  })

  it("CsvStreamWriter.toStream keeps the BOM", async () => {
    const writer = new CsvStreamWriter({ lineSeparator: "\n", bom: true })
    writer.addRow(["a"])
    const bytes = await drain(writer.toStream())
    expect([bytes[0], bytes[1], bytes[2]]).toEqual([0xef, 0xbb, 0xbf])
  })

  it("XlsxStreamWriter.toStream emits a readable workbook", async () => {
    const writer = new XlsxStreamWriter({ name: "Streamed" })
    writer.addRow(["Name", "Score"])
    writer.addRow(["Alice", 95])

    const wb = await readXlsx(await drain(writer.toStream()))
    expect(wb.sheets[0]!.name).toBe("Streamed")
    expect(wb.sheets[0]!.rows).toEqual([
      ["Name", "Score"],
      ["Alice", 95],
    ])
  })

  it("XlsxStreamWriter.toStream buffers — it is not a constant-memory stream", async () => {
    // Documented explicitly because #347 was filed over a buffering writer
    // described as streaming. The whole archive arrives in one chunk.
    const writer = new XlsxStreamWriter({ name: "S" })
    for (let i = 0; i < 500; i++) writer.addRow([i, `row ${i}`])

    const reader = writer.toStream().getReader()
    const first = await reader.read()
    expect(first.done).toBe(false)
    expect(first.value!.length).toBeGreaterThan(0)
    expect((await reader.read()).done).toBe(true)
  })

  it("XlsxStreamWriter.toStream surfaces a failing finish() to the consumer", async () => {
    const writer = new XlsxStreamWriter({ name: "S" })
    writer.addRow([1])
    const boom = new Error("finish exploded")
    ;(writer as unknown as { finish: () => Promise<Uint8Array> }).finish = () =>
      Promise.reject(boom)

    await expect(drain(writer.toStream())).rejects.toThrow("finish exploded")
  })
})

// ── NdjsonStreamWriter ───────────────────────────────────────────────

describe("NdjsonStreamWriter", () => {
  it("addObject + finish returns the buffered NDJSON", () => {
    const writer = new NdjsonStreamWriter()
    writer.addObject({ a: 1 })
    writer.addObject({ a: 2 })
    expect(writer.finish()).toBe('{"a":1}\n{"a":2}\n')
  })

  it("finish() closes the writer", () => {
    const writer = new NdjsonStreamWriter()
    writer.finish()
    expect(() => writer.addObject({ a: 1 })).toThrow()
  })

  it("addRow keys positional values by the configured columns", () => {
    const writer = new NdjsonStreamWriter({ columns: ["id", "name"] })
    writer.addRow([1, "alpha"])
    writer.addRow([2])
    expect(writer.finish()).toBe('{"id":1,"name":"alpha"}\n{"id":2,"name":null}\n')
  })

  it("addRow throws without columns rather than guessing key names", () => {
    const writer = new NdjsonStreamWriter()
    expect(() => writer.addRow([1, 2])).toThrow(InvalidArgumentError)
  })

  it("finish() closes a stream that is still being drained", async () => {
    const writer = new NdjsonStreamWriter()
    const reader = writer.toStream().getReader()
    writer.addObject({ a: 1 })
    expect(new TextDecoder().decode((await reader.read()).value)).toBe('{"a":1}\n')
    writer.finish()
    expect((await reader.read()).done).toBe(true)
  })

  // ── deprecated aliases ─────────────────────────────────────────────

  it("write() still appends, as a deprecated alias of addObject", () => {
    const writer = new NdjsonStreamWriter()
    writer.write({ a: 1 })
    writer.addObject({ a: 2 })
    expect(writer.toString()).toBe('{"a":1}\n{"a":2}\n')
  })

  it("end() still closes, as a deprecated alias of finish", () => {
    const writer = new NdjsonStreamWriter()
    writer.write({ a: 1 })
    writer.end()
    expect(writer.toString()).toBe('{"a":1}\n')
    expect(() => writer.write({ a: 2 })).toThrow()
    expect(() => writer.addObject({ a: 2 })).toThrow()
  })

  it("serializes Dates identically through addObject and the write() alias", () => {
    const d = new Date("2025-04-25T00:00:00Z")

    const viaAddObject = new NdjsonStreamWriter()
    viaAddObject.addObject({ at: d })

    const viaWrite = new NdjsonStreamWriter()
    viaWrite.write({ at: d })

    expect(viaAddObject.finish()).toBe(viaWrite.finish())
    expect(viaAddObject.toString().trim()).toBe(`{"at":"${d.toISOString()}"}`)
  })
})

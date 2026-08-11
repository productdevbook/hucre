import { describe, expect, it } from "vitest"
import { writeOdsStream } from "../src/ods/stream-writer"
import { writeNdjsonStream } from "../src/json/stream"
import { readOds } from "../src/ods/reader"
import { streamOdsRows } from "../src/ods/stream"
import { parseNdjson } from "../src/json/reader"
import { ZipReader } from "../src/zip/reader"
import { InvalidArgumentError } from "../src/errors"
import type { CellValue } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #467 — constant-memory streaming is one of the library's headline
// properties and it was uneven. ODS was the one that stood out: a
// streaming *reader* and no streaming writer at all, so the format with
// the second-best support could not produce a large file without holding
// the whole thing in memory. NDJSON had the class but not the
// `write*Stream(rows, options)` shape that reads naturally at a
// `Response` boundary.
// ═══════════════════════════════════════════════════════════════════════

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
  }
  const total = chunks.reduce((n, c) => n + c.length, 0)
  const out = new Uint8Array(total)
  let at = 0
  for (const c of chunks) {
    out.set(c, at)
    at += c.length
  }
  return out
}

const text = (b: Uint8Array): string => new TextDecoder().decode(b)

describe("writeOdsStream produces a file hucre's own reader accepts", () => {
  const ROWS: CellValue[][] = [
    ["Widget", 3, true, new Date("2024-01-15T00:00:00Z")],
    ["Gadget", 7.5, false, null],
  ]

  it("round-trips values of every type", async () => {
    const bytes = await drain(writeOdsStream(ROWS, { name: "Export" }))
    const wb = await readOds(bytes)

    expect(wb.sheets[0]!.name).toBe("Export")
    expect(wb.sheets[0]!.rows[0]![0]).toBe("Widget")
    expect(wb.sheets[0]!.rows[0]![1]).toBe(3)
    expect(wb.sheets[0]!.rows[0]![2]).toBe(true)
    expect(wb.sheets[0]!.rows[0]![3]).toBeInstanceOf(Date)
    expect(wb.sheets[0]!.rows[1]![1]).toBe(7.5)
    // A trailing null is not written, so the row comes back short — the
    // same thing `writeOds` + `readOds` do, which is the bar here.
    expect(wb.sheets[0]!.rows[1]).toHaveLength(3)
  })

  it("is a valid ODF package, mimetype first and stored", async () => {
    // The one rule an ODF consumer checks before anything else.
    const bytes = await drain(writeOdsStream(ROWS))
    const entries = new ZipReader(bytes).entries()

    expect(entries[0]).toBe("mimetype")
    expect(text(await new ZipReader(bytes).extract("mimetype"))).toBe(
      "application/vnd.oasis.opendocument.spreadsheet",
    )
    for (const part of ["META-INF/manifest.xml", "content.xml", "styles.xml", "meta.xml"]) {
      expect(entries, part).toContain(part)
    }
  })

  it("the streaming reader reads what the streaming writer wrote", async () => {
    const bytes = await drain(writeOdsStream(ROWS, { name: "Export" }))
    const rows: CellValue[][] = []
    for await (const row of streamOdsRows(bytes)) rows.push(row.values)

    expect(rows[0]![0]).toBe("Widget")
    expect(rows[1]![0]).toBe("Gadget")
  })

  it("carries column widths and a header row, which are known up front", async () => {
    const bytes = await drain(
      writeOdsStream(ROWS, {
        columns: [
          { header: "name", width: 20 },
          { header: "qty", width: 10 },
          { header: "ok" },
          { header: "when" },
        ],
      }),
    )

    const content = text(await new ZipReader(bytes).extract("content.xml"))
    expect(content).toContain("style:column-width")
    expect((await readOds(bytes)).sheets[0]!.rows[0]).toEqual(["name", "qty", "ok", "when"])
  })

  it("writes a formula when a cell carries one", async () => {
    const bytes = await drain(writeOdsStream([[1, 2, { formula: "SUM(A1:B1)" }]]))
    const content = text(await new ZipReader(bytes).extract("content.xml"))

    expect(content).toContain("table:formula")
  })

  it("refuses NaN and an unparseable Date the same way writeOds does", async () => {
    // `office:value="NaN"` makes LibreOffice read garbage — a corrupt
    // file rather than an error. See #364.
    const bytes = await drain(writeOdsStream([[Number.NaN, new Date("nonsense")]]))
    const content = text(await new ZipReader(bytes).extract("content.xml"))

    expect(content).not.toContain("NaN")
  })

  it("escapes what would otherwise break the XML", async () => {
    const bytes = await drain(writeOdsStream([['<a & "b">']], { name: "A&B" }))
    const wb = await readOds(bytes)

    expect(wb.sheets[0]!.name).toBe("A&B")
    expect(wb.sheets[0]!.rows[0]![0]).toBe('<a & "b">')
  })

  it("pulls rows lazily rather than draining the source first", async () => {
    // The property the whole thing exists for. If the generator ran to
    // completion before the first chunk, this would count to 5000 before
    // a single byte came out.
    let produced = 0
    function* rows(): Generator<CellValue[]> {
      for (let i = 0; i < 5000; i++) {
        produced++
        yield [`row ${i}`, i]
      }
    }

    const reader = writeOdsStream(rows()).getReader()
    await reader.read()
    const afterFirstChunk = produced
    await reader.cancel()

    expect(afterFirstChunk).toBeLessThan(5000)
  })

  it("takes an async source", async () => {
    async function* rows(): AsyncGenerator<CellValue[]> {
      yield ["a", 1]
      yield ["b", 2]
    }

    const wb = await readOds(await drain(writeOdsStream(rows())))
    expect(wb.sheets[0]!.rows).toHaveLength(2)
  })

  it("still refuses a sheet name Excel would", async () => {
    expect(() => writeOdsStream([], { name: "a/b" })).toThrow()
  })
})

describe("writeNdjsonStream", () => {
  it("writes one JSON object per line", async () => {
    const out = text(
      await drain(
        writeNdjsonStream([
          { name: "Widget", qty: 3 },
          { name: "Gadget", qty: 7 },
        ]),
      ),
    )

    expect(out.trim().split("\n")).toHaveLength(2)
    expect(parseNdjson(out).data).toEqual([
      { name: "Widget", qty: 3 },
      { name: "Gadget", qty: 7 },
    ])
  })

  it("takes positional rows through columns", async () => {
    const out = text(
      await drain(
        writeNdjsonStream(
          [
            ["Widget", 3],
            ["Gadget", 7],
          ],
          { columns: ["name", "qty"] },
        ),
      ),
    )

    expect(parseNdjson(out).data[0]).toEqual({ name: "Widget", qty: 3 })
  })

  it("says why rather than guessing when positional rows have no columns", async () => {
    // NDJSON rows are objects; values with no key names describe nothing.
    // Same contract as NdjsonStreamWriter.addRow.
    await expect(drain(writeNdjsonStream([["a", 1]]))).rejects.toThrow(InvalidArgumentError)
  })

  it("unflattens dot paths when asked", async () => {
    const out = text(await drain(writeNdjsonStream([{ "a.b": 1 }], { unflatten: true })))

    expect(JSON.parse(out.trim())).toEqual({ a: { b: 1 } })
  })

  it("pulls rows lazily", async () => {
    // Rows big enough that the 64 KB chunk boundary lands well before the
    // source is exhausted — otherwise the first chunk is the only chunk,
    // and this would prove nothing.
    let produced = 0
    const filler = "x".repeat(200)
    function* rows(): Generator<Record<string, CellValue>> {
      for (let i = 0; i < 5000; i++) {
        produced++
        yield { i, filler }
      }
    }

    const reader = writeNdjsonStream(rows()).getReader()
    await reader.read()
    const afterFirstChunk = produced
    await reader.cancel()

    expect(afterFirstChunk).toBeLessThan(5000)
  })

  it("takes an async source", async () => {
    async function* rows(): AsyncGenerator<Record<string, CellValue>> {
      yield { a: 1 }
      yield { a: 2 }
    }

    expect(parseNdjson(text(await drain(writeNdjsonStream(rows())))).data).toHaveLength(2)
  })

  it("emits nothing for no rows, rather than a stray newline", async () => {
    expect(text(await drain(writeNdjsonStream([])))).toBe("")
  })
})

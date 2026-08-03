import { describe, expect, it } from "vitest"
import { writeXlsxStream, XlsxStreamWriter } from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { zipStreamChunks } from "../src/zip/stream-writer"
import type { CellValue } from "../src/_types"

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

async function* asyncRows(rows: CellValue[][]): AsyncGenerator<CellValue[]> {
  for (const row of rows) {
    await Promise.resolve()
    yield row
  }
}

// ═══════════════════════════════════════════════════════════════════════
// Streaming ZIP writer
// ═══════════════════════════════════════════════════════════════════════

describe("zipStreamChunks", () => {
  it("produces an archive our own reader can open", async () => {
    const encoder = new TextEncoder()
    const chunks: Uint8Array[] = []
    for await (const chunk of zipStreamChunks([
      { path: "a.txt", data: encoder.encode("hello hello hello hello") },
      { path: "b.txt", data: encoder.encode("world"), compress: false },
    ])) {
      chunks.push(chunk)
    }

    const total = chunks.reduce((n, c) => n + c.length, 0)
    const bytes = new Uint8Array(total)
    let offset = 0
    for (const chunk of chunks) {
      bytes.set(chunk, offset)
      offset += chunk.length
    }

    const zip = new ZipReader(bytes)
    const decoder = new TextDecoder()
    expect(decoder.decode(await zip.extract("a.txt"))).toBe("hello hello hello hello")
    expect(decoder.decode(await zip.extract("b.txt"))).toBe("world")
  })

  it("sets the data descriptor flag on every local header", async () => {
    const encoder = new TextEncoder()
    const chunks: Uint8Array[] = []
    for await (const chunk of zipStreamChunks([{ path: "a.txt", data: encoder.encode("x") }])) {
      chunks.push(chunk)
    }
    const header = chunks[0]!
    const view = new DataView(header.buffer, header.byteOffset, header.byteLength)
    expect(view.getUint32(0, true)).toBe(0x04034b50)
    expect(view.getUint16(6, true) & 0x0008).toBe(0x0008)
    // Sizes and CRC are deferred to the descriptor.
    expect(view.getUint32(14, true)).toBe(0)
    expect(view.getUint32(18, true)).toBe(0)
    expect(view.getUint32(22, true)).toBe(0)
  })

  it("rejects duplicate entry paths", async () => {
    const run = async () => {
      for await (const _ of zipStreamChunks([
        { path: "a.txt", data: new Uint8Array([1]) },
        { path: "a.txt", data: new Uint8Array([2]) },
      ])) {
        // drain
      }
    }
    await expect(run()).rejects.toThrow(/Duplicate ZIP entry/)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// writeXlsxStream
// ═══════════════════════════════════════════════════════════════════════

describe("writeXlsxStream", () => {
  it("writes a workbook readable by readXlsx", async () => {
    const stream = writeXlsxStream(
      { name: "Data", columns: [{ header: "ID" }, { header: "Name" }] },
      [
        [1, "Alice"],
        [2, "Bob"],
      ],
    )

    const workbook = await readXlsx(await collect(stream))
    expect(workbook.sheets).toHaveLength(1)
    expect(workbook.sheets[0].name).toBe("Data")
    expect(workbook.sheets[0].rows).toEqual([
      ["ID", "Name"],
      [1, "Alice"],
      [2, "Bob"],
    ])
  })

  it("accepts an async row source", async () => {
    const stream = writeXlsxStream({ name: "S" }, asyncRows([["a"], ["b"], ["c"]]))
    const workbook = await readXlsx(await collect(stream))
    expect(workbook.sheets[0].rows).toEqual([["a"], ["b"], ["c"]])
  })

  it("accepts object rows through column keys", async () => {
    const stream = writeXlsxStream(
      {
        name: "S",
        columns: [
          { key: "id", header: "ID" },
          { key: "name", header: "Name" },
        ],
      },
      [
        { id: 1, name: "Alice" },
        { id: 2, name: "Bob" },
      ],
    )
    const workbook = await readXlsx(await collect(stream))
    expect(workbook.sheets[0].rows).toEqual([
      ["ID", "Name"],
      [1, "Alice"],
      [2, "Bob"],
    ])
  })

  it("round-trips every cell type", async () => {
    const date = new Date(Date.UTC(2024, 0, 15))
    const stream = writeXlsxStream({ name: "S" }, [["text", 42, true, false, null, date]])
    const workbook = await readXlsx(await collect(stream))
    const row = workbook.sheets[0].rows[0]
    expect(row[0]).toBe("text")
    expect(row[1]).toBe(42)
    expect(row[2]).toBe(true)
    expect(row[3]).toBe(false)
    expect(row[5]).toBeInstanceOf(Date)
    expect((row[5] as Date).toISOString().slice(0, 10)).toBe("2024-01-15")
  })

  it("writes inline strings by default (no sharedStrings part)", async () => {
    const bytes = await collect(writeXlsxStream({ name: "S" }, [["hello"]]))
    const zip = new ZipReader(bytes)
    expect(zip.has("xl/sharedStrings.xml")).toBe(false)
    const sheet = new TextDecoder().decode(await zip.extract("xl/worksheets/sheet1.xml"))
    expect(sheet).toContain('t="inlineStr"')
    expect(sheet).toContain("<is><t>hello</t></is>")
  })

  it("uses the shared string table when asked", async () => {
    const bytes = await collect(
      writeXlsxStream({ name: "S", inlineStrings: false }, [["hello"], ["hello"]]),
    )
    const zip = new ZipReader(bytes)
    expect(zip.has("xl/sharedStrings.xml")).toBe(true)
    const shared = new TextDecoder().decode(await zip.extract("xl/sharedStrings.xml"))
    expect(shared).toContain('uniqueCount="1"')
    const workbook = await readXlsx(bytes)
    expect(workbook.sheets[0].rows).toEqual([["hello"], ["hello"]])
  })

  it("preserves significant whitespace in inline strings", async () => {
    const bytes = await collect(writeXlsxStream({ name: "S" }, [["  padded  "], ["a\nb"]]))
    const workbook = await readXlsx(bytes)
    expect(workbook.sheets[0].rows).toEqual([["  padded  "], ["a\nb"]])
  })

  it("escapes XML-hostile characters", async () => {
    const bytes = await collect(writeXlsxStream({ name: "S" }, [["<a> & \"b\" 'c'"]]))
    const workbook = await readXlsx(bytes)
    expect(workbook.sheets[0].rows[0][0]).toBe("<a> & \"b\" 'c'")
  })

  it("declares every part it references", async () => {
    const bytes = await collect(writeXlsxStream({ name: "S" }, [["x"]]))
    const zip = new ZipReader(bytes)
    const contentTypes = new TextDecoder().decode(await zip.extract("[Content_Types].xml"))

    for (const match of contentTypes.matchAll(/PartName="\/([^"]+)"/g)) {
      expect(zip.has(match[1]!), `missing part ${match[1]}`).toBe(true)
    }
  })

  it("splits into extra sheets past maxRowsPerSheet, repeating the header", async () => {
    const rows: CellValue[][] = []
    for (let i = 1; i <= 6; i++) rows.push([i])

    const stream = writeXlsxStream(
      { name: "Big", columns: [{ header: "N" }], maxRowsPerSheet: 3 },
      rows,
    )
    const workbook = await readXlsx(await collect(stream))

    expect(workbook.sheets.map((s) => s.name)).toEqual(["Big", "Big_2", "Big_3"])
    expect(workbook.sheets[0].rows).toEqual([["N"], [1], [2]])
    expect(workbook.sheets[1].rows).toEqual([["N"], [3], [4]])
    expect(workbook.sheets[2].rows).toEqual([["N"], [5], [6]])
  })

  it("does not repeat headers when told not to", async () => {
    const rows: CellValue[][] = [[1], [2], [3], [4]]
    const stream = writeXlsxStream(
      { name: "Big", columns: [{ header: "N" }], maxRowsPerSheet: 2, repeatHeaders: false },
      rows,
    )
    const workbook = await readXlsx(await collect(stream))
    expect(workbook.sheets[0].rows).toEqual([["N"], [1]])
    expect(workbook.sheets[1].rows).toEqual([[2], [3]])
  })

  it("emits a single sheet when the row count lands exactly on the cap", async () => {
    const stream = writeXlsxStream({ name: "S", maxRowsPerSheet: 3 }, [[1], [2], [3]])
    const workbook = await readXlsx(await collect(stream))
    expect(workbook.sheets).toHaveLength(1)
    expect(workbook.sheets[0].rows).toEqual([[1], [2], [3]])
  })

  it("handles an empty row source", async () => {
    const workbook = await readXlsx(await collect(writeXlsxStream({ name: "Empty" }, [])))
    expect(workbook.sheets).toHaveLength(1)
    expect(workbook.sheets[0].name).toBe("Empty")
    expect(workbook.sheets[0].rows).toEqual([])
  })

  it("rejects a maxRowsPerSheet below 2", () => {
    expect(() =>
      collect(writeXlsxStream({ name: "S", maxRowsPerSheet: 1 }, [[1]])),
    ).rejects.toThrow(/maxRowsPerSheet must be at least 2/)
  })

  it("carries freeze panes and column widths onto every sheet", async () => {
    const bytes = await collect(
      writeXlsxStream(
        {
          name: "S",
          columns: [{ header: "A", width: 25 }],
          freezePane: { rows: 1 },
          maxRowsPerSheet: 2,
        },
        [[1], [2]],
      ),
    )
    const zip = new ZipReader(bytes)
    const decoder = new TextDecoder()
    for (const path of ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]) {
      const xml = decoder.decode(await zip.extract(path))
      expect(xml).toContain('state="frozen"')
      expect(xml).toContain('width="25"')
    }
  })

  it("applies column number formats", async () => {
    const bytes = await collect(
      writeXlsxStream({ name: "S", columns: [{ header: "Amount", numFmt: "#,##0.00" }] }, [[12.5]]),
    )
    const zip = new ZipReader(bytes)
    const styles = new TextDecoder().decode(await zip.extract("xl/styles.xml"))
    expect(styles).toContain("#,##0.00")
  })

  it("honours the 1904 date system", async () => {
    const bytes = await collect(
      writeXlsxStream({ name: "S", dateSystem: "1904" }, [[new Date(Date.UTC(2024, 0, 15))]]),
    )
    const zip = new ZipReader(bytes)
    const workbookXml = new TextDecoder().decode(await zip.extract("xl/workbook.xml"))
    expect(workbookXml).toContain('date1904="1"')
    const workbook = await readXlsx(bytes)
    expect((workbook.sheets[0].rows[0][0] as Date).toISOString().slice(0, 10)).toBe("2024-01-15")
  })

  it("stores uncompressed parts when asked", async () => {
    const bytes = await collect(writeXlsxStream({ name: "S", compress: false }, [["x"]]))
    const workbook = await readXlsx(bytes)
    expect(workbook.sheets[0].rows).toEqual([["x"]])
  })

  it("streams incrementally rather than emitting one blob", async () => {
    const rows: CellValue[][] = []
    for (let i = 0; i < 20_000; i++) rows.push([i, `value-${i}`, i * 1.5])

    const chunks = await collectChunks(writeXlsxStream({ name: "S", compress: false }, rows))
    // The worksheet alone spans many 64 KB flushes, so the archive must
    // arrive as many chunks — a buffered writer would produce a handful.
    expect(chunks.length).toBeGreaterThan(20)
  })

  it("pulls rows lazily and stops when the consumer cancels", async () => {
    let produced = 0
    async function* infinite(): AsyncGenerator<CellValue[]> {
      for (;;) {
        produced++
        yield [produced]
      }
    }

    const stream = writeXlsxStream({ name: "S", compress: false }, infinite())
    const reader = stream.getReader()
    await reader.read()
    await reader.cancel()

    const afterCancel = produced
    await new Promise((resolve) => setTimeout(resolve, 10))
    // Nothing keeps running in the background once the reader is gone.
    expect(produced).toBe(afterCancel)
    // And the source was never drained eagerly.
    expect(produced).toBeLessThan(1_000_000)
  })

  it("matches the buffered writer's cell output", async () => {
    const rows: CellValue[][] = [
      [1, "Alice", true],
      [2, "Bob", false],
    ]

    const buffered = new XlsxStreamWriter({ name: "S", columns: [{ header: "ID" }] })
    for (const row of rows) buffered.addRow(row)
    const bufferedBook = await readXlsx(await buffered.finish())

    const streamedBook = await readXlsx(
      await collect(writeXlsxStream({ name: "S", columns: [{ header: "ID" }] }, rows)),
    )

    expect(streamedBook.sheets[0].rows).toEqual(bufferedBook.sheets[0].rows)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Buffered writer regression
// ═══════════════════════════════════════════════════════════════════════

describe("XlsxStreamWriter", () => {
  it("ships the theme part it declares", async () => {
    const writer = new XlsxStreamWriter({ name: "S" })
    writer.addRow(["x"])
    const zip = new ZipReader(await writer.finish())

    const contentTypes = new TextDecoder().decode(await zip.extract("[Content_Types].xml"))
    for (const match of contentTypes.matchAll(/PartName="\/([^"]+)"/g)) {
      expect(zip.has(match[1]!), `missing part ${match[1]}`).toBe(true)
    }
    expect(zip.has("xl/theme/theme1.xml")).toBe(true)
  })
})

import { describe, expect, it } from "vitest"
import {
  writeXlsxStream,
  writeXlsxStreamSheets,
  type XlsxStreamRow,
} from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { InvalidArgumentError } from "../src/errors"

// ── Streamed workbooks holding more than one sheet ───────────────────
// `writeXlsxStream` streams a single sheet, so a workbook needing two of
// them had to fall back to `writeXlsx` and build the whole object model
// first. These tests pin the multi-sheet path and the guarantees it
// inherits from the single-sheet one.

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  let total = 0
  for await (const chunk of stream) {
    chunks.push(chunk)
    total += chunk.length
  }
  const out = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    out.set(chunk, offset)
    offset += chunk.length
  }
  return out
}

async function extract(workbook: Uint8Array, path: string): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(workbook).extract(path))
}

function countChunks(stream: ReadableStream<Uint8Array>): Promise<number> {
  return (async () => {
    let chunks = 0
    for await (const _chunk of stream) chunks++
    return chunks
  })()
}

describe("writeXlsxStreamSheets", () => {
  it("writes every sheet under its own name, in the order given", async () => {
    const wb = await readXlsx(
      await drain(
        writeXlsxStreamSheets([
          { name: "Accepted", rows: [["a", 1]] },
          { name: "Rejected", rows: [["b", 2]] },
          { name: "Summary", rows: [["c", 3]] },
        ]),
      ),
    )

    expect(wb.sheets.map((sheet) => sheet.name)).toEqual(["Accepted", "Rejected", "Summary"])
    expect(wb.sheets[0]!.rows).toEqual([["a", 1]])
    expect(wb.sheets[1]!.rows).toEqual([["b", 2]])
    expect(wb.sheets[2]!.rows).toEqual([["c", 3]])
  })

  it("gives each sheet its own columns and freeze pane", async () => {
    const wb = await readXlsx(
      await drain(
        writeXlsxStreamSheets([
          {
            name: "Wide",
            rows: [[1, 2]],
            columns: [{ width: 40 }, { width: 12 }],
            freezePane: { rows: 1 },
          },
          {
            name: "Narrow",
            rows: [[3]],
            columns: [{ width: 5 }],
          },
        ]),
      ),
    )

    expect(wb.sheets[0]!.columns?.map((column) => column.width)).toEqual([40, 12])
    expect(wb.sheets[1]!.columns?.map((column) => column.width)).toEqual([5])
    expect(wb.sheets[0]!.freezePane).toEqual({ rows: 1 })
    expect(wb.sheets[1]!.freezePane).toBeUndefined()
  })

  it("resolves each sheet's header from its own columns", async () => {
    const wb = await readXlsx(
      await drain(
        writeXlsxStreamSheets([
          {
            name: "People",
            rows: [{ id: 1, name: "Ada" }],
            columns: [
              { key: "id", header: "ID" },
              { key: "name", header: "Name" },
            ],
          },
          {
            name: "Places",
            rows: [{ city: "Lisbon" }],
            columns: [{ key: "city", header: "City" }],
          },
        ]),
      ),
    )

    expect(wb.sheets[0]!.rows).toEqual([
      ["ID", "Name"],
      [1, "Ada"],
    ])
    expect(wb.sheets[1]!.rows).toEqual([["City"], ["Lisbon"]])
  })

  it("writes one style table holding the formats of every sheet", async () => {
    // xl/styles.xml is a workbook-wide part written once. A per-sheet
    // style table would leave only the last sheet's formats in it, and the
    // earlier sheets' cells would index into entries that are not there.
    const bytes = await drain(
      writeXlsxStreamSheets([
        { name: "One", rows: [[1]], columns: [{ numFmt: "0.000" }] },
        { name: "Two", rows: [[2]], columns: [{ numFmt: "#,##0.00" }] },
      ]),
    )

    const styles = await extract(bytes, "xl/styles.xml")

    expect(styles).toContain("0.000")
    expect(styles).toContain("#,##0.00")
  })

  it("writes one string table shared by every sheet", async () => {
    const bytes = await drain(
      writeXlsxStreamSheets(
        [
          { name: "One", rows: [["repeated"]] },
          { name: "Two", rows: [["repeated"]] },
        ],
        { stringMode: "shared" },
      ),
    )
    const wb = await readXlsx(bytes)

    expect(wb.sheets[0]!.rows).toEqual([["repeated"]])
    expect(wb.sheets[1]!.rows).toEqual([["repeated"]])

    // Both sheets point at the same entry, so the string is stored once.
    const table = await extract(bytes, "xl/sharedStrings.xml")

    expect(table.match(/<si>/g)).toHaveLength(1)
    expect(table).toContain('uniqueCount="1"')
  })

  it("does not touch a sheet's rows until the previous sheet runs out", async () => {
    const pulled: string[] = []

    function* track(label: string, rows: XlsxStreamRow[]): Generator<XlsxStreamRow> {
      for (const row of rows) {
        pulled.push(label)
        yield row
      }
    }

    const stream = writeXlsxStreamSheets([
      { name: "First", rows: track("first", [[1], [2]]) },
      { name: "Second", rows: track("second", [[3]]) },
    ])

    await drain(stream)

    expect(pulled).toEqual(["first", "first", "second"])
  })

  it("emits the workbook in chunks instead of one buffer", async () => {
    const rows = Array.from({ length: 4_000 }, (_, index) => [index, `row ${index}`])

    const chunks = await countChunks(
      writeXlsxStreamSheets([
        { name: "One", rows },
        { name: "Two", rows },
      ]),
    )

    expect(chunks).toBeGreaterThan(1)
  })

  it("rolls a sheet over past its cap, keeping the sheets after it intact", async () => {
    const wb = await readXlsx(
      await drain(
        writeXlsxStreamSheets(
          [
            { name: "Big", rows: Array.from({ length: 5 }, (_, index) => [index]) },
            { name: "Small", rows: [["tail"]] },
          ],
          { maxRowsPerSheet: 2, repeatHeaders: false },
        ),
      ),
    )

    expect(wb.sheets.map((sheet) => sheet.name)).toEqual(["Big", "Big_2", "Big_3", "Small"])
    expect(wb.sheets[3]!.rows).toEqual([["tail"]])
  })

  it("lets a single sheet override the workbook-wide cap", async () => {
    const wb = await readXlsx(
      await drain(
        writeXlsxStreamSheets(
          [
            { name: "Capped", rows: [[1], [2], [3], [4]] },
            { name: "Uncapped", rows: [[1], [2], [3], [4]], maxRowsPerSheet: Infinity },
          ],
          { maxRowsPerSheet: 2, repeatHeaders: false },
        ),
      ),
    )

    expect(wb.sheets.map((sheet) => sheet.name)).toEqual(["Capped", "Capped_2", "Uncapped"])
  })

  it("skips a rollover name another sheet already owns", async () => {
    const wb = await readXlsx(
      await drain(
        writeXlsxStreamSheets(
          [
            { name: "Data", rows: [[1], [2], [3]] },
            { name: "Data_2", rows: [["taken"]] },
          ],
          { maxRowsPerSheet: 2, repeatHeaders: false },
        ),
      ),
    )

    // "Data" rolls over, but "Data_2" belongs to the second sheet, so the
    // rolled part has to take the next free ordinal — Excel refuses a
    // workbook with two sheets of the same name.
    expect(wb.sheets.map((sheet) => sheet.name)).toEqual(["Data", "Data_3", "Data_2"])
    expect(wb.sheets[2]!.rows).toEqual([["taken"]])
  })

  it("rejects duplicate sheet names before returning the stream", () => {
    expect(() =>
      writeXlsxStreamSheets([
        { name: "Data", rows: [[1]] },
        { name: "data", rows: [[2]] },
      ]),
    ).toThrow(InvalidArgumentError)
  })

  it("rejects an invalid sheet name before returning the stream", () => {
    expect(() => writeXlsxStreamSheets([{ name: "with/slash", rows: [[1]] }])).toThrow(
      InvalidArgumentError,
    )
  })

  it("rejects a workbook with no sheets", () => {
    expect(() => writeXlsxStreamSheets([])).toThrow(InvalidArgumentError)
  })

  it("rejects an unusable row cap before returning the stream", () => {
    expect(() =>
      writeXlsxStreamSheets([{ name: "Data", rows: [[1]], maxRowsPerSheet: 1 }]),
    ).toThrow(InvalidArgumentError)
  })

  it("accepts async row sources", async () => {
    async function* rows(): AsyncGenerator<XlsxStreamRow> {
      yield ["a"]
      yield ["b"]
    }

    const wb = await readXlsx(await drain(writeXlsxStreamSheets([{ name: "Async", rows: rows() }])))

    expect(wb.sheets[0]!.rows).toEqual([["a"], ["b"]])
  })

  it("honours the 1904 date system for the whole workbook", async () => {
    const date = new Date(Date.UTC(2020, 0, 2))
    const wb = await readXlsx(
      await drain(
        writeXlsxStreamSheets(
          [
            { name: "One", rows: [[date]] },
            { name: "Two", rows: [[date]] },
          ],
          { dateSystem: "1904" },
        ),
      ),
    )

    expect(wb.sheets[0]!.rows[0]![0]).toEqual(date)
    expect(wb.sheets[1]!.rows[0]![0]).toEqual(date)
  })
})

describe("writeXlsxStreamSheets — formatting is per sheet", () => {
  // `rowDefs` and `merges` describe one sheet's layout, so a workbook of
  // several has to keep them apart: a height or a range declared for the
  // report must not follow the sheet after it.
  it("keeps each sheet's cell styles, row heights and merges to itself", async () => {
    const title = { font: { name: "Manrope", size: 20, bold: true } }

    const wb = await readXlsx(
      await drain(
        writeXlsxStreamSheets([
          {
            name: "Report",
            rows: [[{ value: "Q3", style: title }, null, null], ["data"]],
            rowDefs: new Map([[0, { height: 30 }]]),
            merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
          },
          {
            name: "Rejects",
            rows: [["plain"], ["also plain"]],
            rowDefs: new Map([[1, { hidden: true }]]),
          },
        ]),
      ),
      { readStyles: true },
    )

    const [report, rejects] = wb.sheets

    expect(report!.cells!.get("0,0")!.style!.font!.name).toBe("Manrope")
    expect(report!.rowDefs!.get(0)!.height).toBe(30)
    expect(report!.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }])

    // Nothing of the first sheet's layout leaks into the second.
    expect(rejects!.merges ?? []).toEqual([])
    expect(rejects!.rowDefs?.get(0)?.height).toBeUndefined()
    expect(rejects!.rowDefs!.get(1)!.hidden).toBe(true)
  })

  it("keys a sheet's rowDefs by its own rows, not the workbook's", async () => {
    const wb = await readXlsx(
      await drain(
        writeXlsxStreamSheets([
          { name: "First", rows: [["a"], ["b"]] },
          { name: "Second", rows: [["c"]], rowDefs: new Map([[0, { height: 44 }]]) },
        ]),
      ),
    )

    expect(wb.sheets[0]!.rowDefs?.get(0)?.height).toBeUndefined()
    expect(wb.sheets[1]!.rowDefs!.get(0)!.height).toBe(44)
  })
})

describe("writeXlsxStream — single-sheet behaviour is unchanged", () => {
  it("produces the same bytes as the multi-sheet writer given one sheet", async () => {
    const rows = [
      ["a", 1],
      ["b", 2],
    ]
    const columns = [{ width: 20 }, { width: 8 }]

    const single = await drain(writeXlsxStream(rows, { name: "Report", columns }))
    const multi = await drain(writeXlsxStreamSheets([{ name: "Report", rows, columns }]))

    expect(single).toEqual(multi)
  })

  it("carries row heights and merges through the delegation", async () => {
    const rowDefs = new Map([[0, { height: 30 }]])
    const merges = [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }]
    const rows = [["wide", null, null]]

    const wb = await readXlsx(
      await drain(writeXlsxStream(rows, { name: "Report", rowDefs, merges })),
    )

    expect(wb.sheets[0]!.rowDefs!.get(0)!.height).toBe(30)
    expect(wb.sheets[0]!.merges).toEqual(merges)
  })
})

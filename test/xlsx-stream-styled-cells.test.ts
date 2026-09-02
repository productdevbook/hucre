import type { CellInput } from "../src/_types"
import { describe, expect, it } from "vitest"
import { writeXlsxStream, XlsxStreamWriter } from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"
import type { CellStyle } from "../src/_types"

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

const title: CellStyle = {
  font: { name: "Manrope", size: 20, bold: true },
  alignment: { horizontal: "center" },
}
const data: CellStyle = { font: { name: "Arial", size: 8 } }

describe("writeXlsxStream — per-cell formatting", () => {
  it("styles a single cell without styling the column it sits in", async () => {
    const bytes = await collect(
      writeXlsxStream(
        [
          [{ value: "Report", style: title }, "plain"],
          ["a", "b"],
        ],
        { name: "S", columns: [{ style: data }, { style: data }] },
      ),
    )

    const wb = await readXlsx(bytes, { readStyles: true })
    const cells = wb.sheets[0]!.cells!

    expect(cells.get("0,0")!.style!.font!.name).toBe("Manrope")
    expect(cells.get("0,1")!.style!.font!.name).toBe("Arial")
    expect(cells.get("1,0")!.style!.font!.name).toBe("Arial")
  })

  it("keeps the column style for cells that do not carry one", async () => {
    const bytes = await collect(
      writeXlsxStream([[{ value: 1 }, 2]], {
        name: "S",
        columns: [{ style: data }, { style: data }],
      }),
    )

    const wb = await readXlsx(bytes, { readStyles: true })
    const cells = wb.sheets[0]!.cells!

    expect(cells.get("0,0")!.style!.font!.size).toBe(8)
    expect(cells.get("0,1")!.style!.font!.size).toBe(8)
  })

  it("writes a formula, with the cached result when one is given", async () => {
    const bytes = await collect(
      writeXlsxStream(
        [
          [
            { formula: "SUM(B1:B9)", style: data },
            { value: 3, formula: "1+2" },
          ],
        ],
        { name: "S" },
      ),
    )

    const wb = await readXlsx(bytes)
    const cells = wb.sheets[0]!.cells!

    expect(cells.get("0,0")!.formula).toBe("SUM(B1:B9)")
    expect(cells.get("0,1")!.formula).toBe("1+2")
    expect(cells.get("0,1")!.value).toBe(3)
  })

  it("applies row heights from rowDefs, including on an otherwise empty row", async () => {
    const bytes = await collect(
      writeXlsxStream([["a"], [], ["b"]], {
        name: "S",
        rowDefs: new Map([
          [0, { height: 30 }],
          [1, { height: 44 }],
        ]),
      }),
    )

    const wb = await readXlsx(bytes)
    const rowDefs = wb.sheets[0]!.rowDefs!

    expect(rowDefs.get(0)!.height).toBe(30)
    expect(rowDefs.get(1)!.height).toBe(44)
  })

  it("emits merged ranges after the streamed rows", async () => {
    const bytes = await collect(
      writeXlsxStream([[{ value: "wide", style: title }, null, null]], {
        name: "S",
        merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
      }),
    )

    const wb = await readXlsx(bytes)

    expect(wb.sheets[0]!.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }])
  })

  it("treats a Date as a value, not as a styled cell", async () => {
    const bytes = await collect(writeXlsxStream([[new Date(Date.UTC(2024, 0, 1))]], { name: "S" }))

    const wb = await readXlsx(bytes)

    expect(wb.sheets[0]!.rows[0]![0]).toBeInstanceOf(Date)
  })

  it("keeps a text or boolean formula result instead of dropping it", async () => {
    const bytes = await collect(
      writeXlsxStream(
        [
          [
            { value: "ok", formula: 'IF(1=1,"ok","no")' },
            { value: true, formula: "1=1" },
          ],
        ],
        { name: "S" },
      ),
    )

    const wb = await readXlsx(bytes)
    const cells = wb.sheets[0]!.cells!

    expect(cells.get("0,0")!.value).toBe("ok")
    expect(cells.get("0,1")!.value).toBe(true)
  })
})

describe("XlsxStreamWriter — the same options the generator takes", () => {
  it("honours styled cells, rowDefs and merges", async () => {
    const writer = new XlsxStreamWriter({
      name: "S",
      rowDefs: new Map([[0, { height: 30 }]]),
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
    })

    writer.addRow([{ value: "wide", style: title }, null, null])
    writer.addRow(["data"])

    const wb = await readXlsx(await writer.finish(), { readStyles: true })
    const sheet = wb.sheets[0]!

    expect(sheet.cells!.get("0,0")!.style!.font!.name).toBe("Manrope")
    expect(sheet.rowDefs!.get(0)!.height).toBe(30)
    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }])
  })

  it("keys rowDefs by the caller's row number across a sheet rollover", async () => {
    const writer = new XlsxStreamWriter({
      name: "S",
      maxRowsPerSheet: 2,
      repeatHeaders: false,
      rowDefs: new Map([
        [0, { height: 11 }],
        [2, { height: 33 }],
      ]),
    })

    writer.addRow(["r0"])
    writer.addRow(["r1"])
    writer.addRow(["r2"])

    const wb = await readXlsx(await writer.finish())

    // Row 0 of the first sheet, and row 2 — the first row of the second
    // sheet, which restarts the per-sheet index at 0.
    expect(wb.sheets[0]!.rowDefs!.get(0)!.height).toBe(11)
    expect(wb.sheets[1]!.rowDefs!.get(0)!.height).toBe(33)
    // The definition for row 0 must not be reapplied to the new sheet.
    expect(wb.sheets[1]!.rowDefs!.get(0)!.height).not.toBe(11)
  })

  it("drops a non-finite formula result instead of writing it", async () => {
    const bytes = await collect(
      writeXlsxStream(
        [
          [
            { value: Number.NaN, formula: "0/0" },
            { value: Number.POSITIVE_INFINITY, formula: "1/0" },
            { value: 2, formula: "1+1" },
          ],
        ],
        { name: "S" },
      ),
    )

    const wb = await readXlsx(bytes)
    const cells = wb.sheets[0]!.cells!

    expect(cells.get("0,0")!.formula).toBe("0/0")
    expect(cells.get("0,0")!.value).toBeNull()
    expect(cells.get("0,1")!.value).toBeNull()
    expect(cells.get("0,2")!.value).toBe(2)
  })

  it("serializes every rowDef property, not only the height", async () => {
    const writer = new XlsxStreamWriter({
      name: "S",
      rowDefs: new Map([
        [0, { height: 20, hidden: true }],
        [1, { outlineLevel: 2, collapsed: true }],
      ]),
    })

    writer.addRow(["a"])
    writer.addRow(["b"])

    const wb = await readXlsx(await writer.finish())
    const rowDefs = wb.sheets[0]!.rowDefs!

    expect(rowDefs.get(0)).toMatchObject({ height: 20, hidden: true })
    expect(rowDefs.get(1)).toMatchObject({ outlineLevel: 2, collapsed: true })
  })

  it("emits an otherwise empty row when its definition asks to hide it", async () => {
    const bytes = await collect(
      writeXlsxStream([["a"], [], ["c"]], {
        name: "S",
        rowDefs: new Map([[1, { hidden: true }]]),
      }),
    )

    const wb = await readXlsx(bytes)

    expect(wb.sheets[0]!.rowDefs!.get(1)!.hidden).toBe(true)
  })

  it("repeats the header as first emitted, even if the caller reuses the cell object", async () => {
    const cell: CellInput = { value: "H", style: title }
    const writer = new XlsxStreamWriter({
      name: "S",
      maxRowsPerSheet: 2,
      repeatHeaders: true,
    })

    writer.addRow([cell])
    cell.value = "mutated"
    writer.addRow([cell])
    writer.addRow(["overflow"])

    const wb = await readXlsx(await writer.finish())

    expect(wb.sheets[0]!.rows[0]![0]).toBe("H")
    expect(wb.sheets[0]!.rows[1]![0]).toBe("mutated")
    // The rolled-over sheet repeats the header as it was first emitted.
    expect(wb.sheets[1]!.rows[0]![0]).toBe("H")
  })
})

import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { insertColumns, insertRows } from "../src/sheet-ops"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { InvalidArgumentError } from "../src/errors"

// A `<col max="16384">` range is how Excel stores a format set on whole
// columns. It reads as one ColumnDef per column up to XFD, so an insert
// used to push the array past the last column and saveXlsx threw. The
// rows axis mirrors it: DV/CF ranges reaching row 1048576 used to come
// back as `A1048577` with no error at all, since rows are never
// validated the way columns are.

async function sheetXml(bytes: Uint8Array): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(bytes).extract("xl/worksheets/sheet1.xml"))
}

async function template(): Promise<Uint8Array> {
  const base = await writeXlsx({ sheets: [{ name: "S", rows: [["x"]], columns: [{ width: 20 }] }] })
  const all = await new ZipReader(base).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    const xml = new TextDecoder().decode(data)
    zw.add(
      name,
      name === "xl/worksheets/sheet1.xml"
        ? new TextEncoder().encode(
            xml.replace(
              /<cols>[\s\S]*<\/cols>/,
              '<cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="16384" width="9" customWidth="1"/></cols>',
            ),
          )
        : data,
    )
  }
  return zw.build()
}

describe("insertColumns at the grid edge", () => {
  it("drops the column definitions shifted past XFD and saves", async () => {
    const wb = await openXlsx(await template())
    insertColumns(wb.sheets[0]!, 1, 2)
    expect(wb.sheets[0]!.columns).toHaveLength(16384)
    const xml = await sheetXml(await saveXlsx(wb))
    expect(xml).toMatch(/<col min="4" max="16384" [^>]*width="9"/)
  })

  it("truncates all-blank row tails past XFD and saves", async () => {
    const wb = await openXlsx(await template())
    wb.sheets[0]!.rows[0] = Array.from({ length: 16384 }, () => null)
    insertColumns(wb.sheets[0]!, 1, 2)
    expect(wb.sheets[0]!.rows[0]).toHaveLength(16384)
    await saveXlsx(wb)
  })

  it("refuses to push nonblank cells off the sheet, before mutating it", async () => {
    const wb = await openXlsx(await template())
    const row: (string | null)[] = Array.from({ length: 16384 }, () => null)
    row[16383] = "v"
    wb.sheets[0]!.rows[0] = row
    expect(() => insertColumns(wb.sheets[0]!, 1, 2)).toThrow(/off the sheet/)
    expect(wb.sheets[0]!.rows[0]).toHaveLength(16384)
  })

  it("refuses a sparse cell pushed past XFD without mutating the sheet", async () => {
    const wb = await openXlsx(await template())
    wb.sheets[0]!.cells = new Map([["0,16383", { type: "number", value: 1 }]])
    expect(() => insertColumns(wb.sheets[0]!, 0, 1)).toThrow(InvalidArgumentError)
    expect(() => insertColumns(wb.sheets[0]!, 0, 1)).toThrow(/off the sheet/)
    expect(wb.sheets[0]!.cells!.get("0,16383")).toEqual({ type: "number", value: 1 })
    expect(wb.sheets[0]!.cells!.has("0,16384")).toBe(false)
  })

  it("accepts the last insert that fits and refuses one past it", async () => {
    const fitting = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["x"]] }] }))
    const fittingRow: (string | null)[] = Array.from({ length: 16382 }, () => null)
    fittingRow[16381] = "v"
    fitting.sheets[0]!.rows[0] = fittingRow
    insertColumns(fitting.sheets[0]!, 0, 2)
    expect(fitting.sheets[0]!.rows[0]).toHaveLength(16384)
    await saveXlsx(fitting)

    const falling = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["x"]] }] }))
    const fallingRow: (string | null)[] = Array.from({ length: 16383 }, () => null)
    fallingRow[16382] = "v"
    falling.sheets[0]!.rows[0] = fallingRow
    expect(() => insertColumns(falling.sheets[0]!, 0, 2)).toThrow(/Cannot insert 2 columns/)
  })

  it("refuses an insert position past the grid", async () => {
    const wb = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["x"]] }] }))
    expect(() => insertColumns(wb.sheets[0]!, 16385, 1)).toThrow(/only has columns/)
    expect(() => insertRows(wb.sheets[0]!, 1048577, 1)).toThrow(/only has rows/)
  })

  it("clips a conditional range to XFD instead of throwing", async () => {
    const wb = await openXlsx(await template())
    wb.sheets[0]!.conditionalRules = [
      { range: "A1:XFD1", type: "cellIs", operator: "greaterThan", formula: "1", priority: 1 },
    ]
    insertColumns(wb.sheets[0]!, 1, 2)
    expect(wb.sheets[0]!.conditionalRules![0]!.range).toBe("A1:XFD1")
    await saveXlsx(wb)
  })

  it("clips a merge pushed partly past XFD and drops one pushed wholly past", async () => {
    const wb = await openXlsx(await template())
    wb.sheets[0]!.merges = [
      { startRow: 0, startCol: 16380, endRow: 0, endCol: 16383 },
      { startRow: 1, startCol: 16382, endRow: 1, endCol: 16383 },
    ]
    insertColumns(wb.sheets[0]!, 0, 2)
    expect(wb.sheets[0]!.merges).toEqual([
      { startRow: 0, startCol: 16382, endRow: 0, endCol: 16383 },
    ])
    await saveXlsx(wb)
  })

  it("drops page breaks pushed past XFD instead of writing a bogus brk id", async () => {
    const wb = await openXlsx(await template())
    wb.sheets[0]!.colBreaks = [16383]
    insertColumns(wb.sheets[0]!, 0, 1)
    expect(wb.sheets[0]!.colBreaks).toEqual([])
    const xml = await sheetXml(await saveXlsx(wb))
    expect(xml).not.toContain("16385")
  })

  it("holds image and text-box anchors at XFD", async () => {
    const wb = await openXlsx(await template())
    wb.sheets[0]!.images = [
      { data: new Uint8Array(), type: "png", anchor: { from: { row: 0, col: 16383 } } },
    ]
    wb.sheets[0]!.textBoxes = [{ text: "x", anchor: { from: { row: 0, col: 16383 } } }]
    insertColumns(wb.sheets[0]!, 0, 1)
    expect(wb.sheets[0]!.images![0]!.anchor.from.col).toBe(16383)
    expect(wb.sheets[0]!.textBoxes![0]!.anchor.from.col).toBe(16383)
  })

  it("drops an auto filter and tables pushed off the sheet, clips the rest", async () => {
    const wb = await openXlsx(await template())
    wb.sheets[0]!.autoFilter = { range: "XFD1:XFD3" }
    wb.sheets[0]!.tables = [
      { name: "T1", columns: [{ name: "C1" }, { name: "C2" }], range: "XFD1:XFD2" },
      { name: "T2", columns: [{ name: "C1" }], range: "A1:A2" },
    ]
    insertColumns(wb.sheets[0]!, 0, 1)
    expect(wb.sheets[0]!.autoFilter).toBeUndefined()
    expect(wb.sheets[0]!.tables!.map((t) => t.range)).toEqual(["B1:B2"])
    await saveXlsx(wb)
  })
})

describe("insertRows at the grid edge", () => {
  it("clips a validation range to row 1048576 instead of writing A1048577", async () => {
    const wb = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["x"]] }] }))
    wb.sheets[0]!.dataValidations = [
      { range: "A1:A1048576", type: "whole", operator: "between", formula1: "1", formula2: "9" },
    ]
    insertRows(wb.sheets[0]!, 0, 1)
    expect(wb.sheets[0]!.dataValidations![0]!.range).toBe("A2:A1048576")
    const xml = await sheetXml(await saveXlsx(wb))
    expect(xml).not.toContain("1048577")
  })

  it("refuses a sparse cell pushed past row 1048576", async () => {
    const wb = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["x"]] }] }))
    wb.sheets[0]!.cells = new Map([["1048575,0", { type: "number", value: 1 }]])
    expect(() => insertRows(wb.sheets[0]!, 0, 1)).toThrow(/off the sheet/)
  })

  it("clips row-axis merges, conditional ranges and anchors to the bottom", async () => {
    const wb = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["x"]] }] }))
    wb.sheets[0]!.merges = [{ startRow: 1048573, startCol: 0, endRow: 1048575, endCol: 1 }]
    wb.sheets[0]!.conditionalRules = [
      { range: "A1:A1048576", type: "cellIs", operator: "greaterThan", formula: "1", priority: 1 },
    ]
    wb.sheets[0]!.images = [
      { data: new Uint8Array(), type: "png", anchor: { from: { row: 1048575, col: 0 } } },
    ]
    wb.sheets[0]!.textBoxes = [{ text: "x", anchor: { from: { row: 1048575, col: 0 } } }]
    insertRows(wb.sheets[0]!, 0, 1)
    expect(wb.sheets[0]!.merges).toEqual([
      { startRow: 1048574, startCol: 0, endRow: 1048575, endCol: 1 },
    ])
    expect(wb.sheets[0]!.conditionalRules![0]!.range).toBe("A2:A1048576")
    expect(wb.sheets[0]!.images![0]!.anchor.from.row).toBe(1048575)
    expect(wb.sheets[0]!.textBoxes![0]!.anchor.from.row).toBe(1048575)
    const xml = await sheetXml(await saveXlsx(wb))
    expect(xml).not.toContain("1048577")
  })

  it("drops row formats and breaks pushed past the bottom", async () => {
    const wb = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["x"]] }] }))
    wb.sheets[0]!.rowDefs = new Map([[1048575, { height: 20 }]])
    wb.sheets[0]!.rowBreaks = [1048575]
    insertRows(wb.sheets[0]!, 0, 1)
    expect(wb.sheets[0]!.rowDefs!.has(1048576)).toBe(false)
    expect(wb.sheets[0]!.rowBreaks).toEqual([])
    await saveXlsx(wb)
  })

  it("truncates blank rows falling off the bottom", async () => {
    const wb = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["x"]] }] }))
    wb.sheets[0]!.rows = Array.from({ length: 1048576 }, () => [])
    insertRows(wb.sheets[0]!, 0, 1)
    expect(wb.sheets[0]!.rows).toHaveLength(1048576)
  })
})

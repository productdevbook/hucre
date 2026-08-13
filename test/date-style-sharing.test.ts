import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #472 — a `Date` with no explicit number format gets one, and the
// writer built `{ ...style, numFmt: "yyyy-mm-dd" }` fresh for every such
// cell. That defeated the xf identity cache #435 added: measured on the
// 100,000 x 12 benchmark, `registerXf` was called 400,000 times — exactly
// the date-cell count — with **zero** identity hits, to produce one xf.
//
// Sharing one object per distinct input makes the cache hit on the second
// date cell and every one after: 399,999 of 400,000. What follows is the
// behaviour that must not change for it.
// ═══════════════════════════════════════════════════════════════════════

const dec = new TextDecoder()

async function stylesXml(bytes: Uint8Array): Promise<string> {
  return dec.decode(await new ZipReader(bytes).extract("xl/styles.xml"))
}

const DAY = new Date("2024-03-17T00:00:00Z")

describe("a bare date still gets its format", () => {
  it("reads back as a Date", async () => {
    const wb = await readXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [[DAY]] }] }))

    expect(wb.sheets[0]!.rows[0]![0]).toBeInstanceOf(Date)
    expect((wb.sheets[0]!.rows[0]![0] as Date).toISOString()).toBe("2024-03-17T00:00:00.000Z")
  })

  it("and many of them share one format record", async () => {
    // The point of the change, visible in the file: 500 date cells must
    // not produce 500 xfs.
    const rows = Array.from({ length: 500 }, () => [DAY, DAY, DAY])
    const xml = await stylesXml(await writeXlsx({ sheets: [{ name: "S", rows }] }))

    const xfCount = (xml.match(/<xf /g) ?? []).length
    expect(xfCount).toBeLessThan(5)
    expect((xml.match(/yyyy-mm-dd/g) ?? []).length).toBe(1)
  })
})

describe("a styled date keeps its own style", () => {
  it("the style survives alongside the added format", async () => {
    const style: CellStyle = { font: { bold: true } }
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[DAY]], cells: new Map([["0,0", { value: DAY, style }]]) }],
    })

    const cell = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")

    expect(cell?.style?.font?.bold).toBe(true)
    expect(cell?.style?.numFmt).toBe("yyyy-mm-dd")
  })

  it("two different styles do not collapse into one", async () => {
    // The risk of caching by identity: a shared derived object must be
    // shared only among cells that started from the same style.
    const bold: CellStyle = { font: { bold: true } }
    const italic: CellStyle = { font: { italic: true } }
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [[DAY, DAY]],
          cells: new Map([
            ["0,0", { value: DAY, style: bold }],
            ["0,1", { value: DAY, style: italic }],
          ]),
        },
      ],
    })

    const cells = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.cells!

    expect(cells.get("0,0")?.style?.font?.bold).toBe(true)
    expect(cells.get("0,0")?.style?.font?.italic).toBeUndefined()
    expect(cells.get("0,1")?.style?.font?.italic).toBe(true)
    expect(cells.get("0,1")?.style?.font?.bold).toBeUndefined()
  })

  it("an explicit numFmt on a date is not overwritten", async () => {
    const style: CellStyle = { numFmt: "dd/mm/yyyy" }
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[DAY]], cells: new Map([["0,0", { value: DAY, style }]]) }],
    })

    const cell = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")

    expect(cell?.style?.numFmt).toBe("dd/mm/yyyy")
  })

  it("does not mutate the caller's style object", async () => {
    // The shared object is frozen, but the caller's is not — and adding
    // the format to theirs would be a surprising side effect.
    const style: CellStyle = { font: { bold: true } }
    await writeXlsx({
      sheets: [{ name: "S", rows: [[DAY]], cells: new Map([["0,0", { value: DAY, style }]]) }],
    })

    expect(style).toEqual({ font: { bold: true } })
    expect(style.numFmt).toBeUndefined()
  })
})

describe("the sheet still writes what it always did", () => {
  it("dates mixed with everything else", async () => {
    const rows = [
      ["text", 42, DAY, true],
      [null, -1.5, new Date("2020-01-01T00:00:00Z"), false],
    ]
    const wb = await readXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }))

    expect(wb.sheets[0]!.rows[0]![0]).toBe("text")
    expect(wb.sheets[0]!.rows[0]![1]).toBe(42)
    expect((wb.sheets[0]!.rows[0]![2] as Date).toISOString()).toBe("2024-03-17T00:00:00.000Z")
    expect(wb.sheets[0]!.rows[0]![3]).toBe(true)
    expect((wb.sheets[0]!.rows[1]![2] as Date).toISOString()).toBe("2020-01-01T00:00:00.000Z")
  })

  it("under the 1904 date system", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[DAY]] }], dateSystem: "1904" })
    const wb = await readXlsx(bytes)

    expect((wb.sheets[0]!.rows[0]![0] as Date).toISOString()).toBe("2024-03-17T00:00:00.000Z")
  })
})

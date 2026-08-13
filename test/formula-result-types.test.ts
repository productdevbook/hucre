import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"

// ═══════════════════════════════════════════════════════════════════════
// #497 — `Cell.formulaResult` was assigned in exactly one place: the
// numeric arm of the cell-type switch. A formula whose cached result is
// a string, an error or a boolean set `value` and left `formulaResult`
// undefined, so the cached result survived only when it happened to be a
// number.
//
// Not a missing field — a round-trip loss. The *writer* has always been
// able to write string and boolean results back, so `readXlsx` →
// `writeXlsx` emitted `<f>` with no `<v>`, and anything opening the
// result without recalculating saw an empty cell where Excel showed
// `#DIV/0!` or `xy`.
// ═══════════════════════════════════════════════════════════════════════

const enc = new TextEncoder()
const dec = new TextDecoder()

/** A workbook whose first row is raw `<c>` elements of our choosing. */
async function withCells(cellsXml: string): Promise<Uint8Array> {
  const base = await writeXlsx({ sheets: [{ name: "S", rows: [[1]] }] })
  const all = await new ZipReader(base).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    zw.add(
      name,
      name === "xl/worksheets/sheet1.xml"
        ? enc.encode(
            dec
              .decode(data)
              .replace(
                /<sheetData>.*<\/sheetData>/,
                `<sheetData><row r="1">${cellsXml}</row></sheetData>`,
              ),
          )
        : data,
    )
  }
  return zw.build()
}

const NUMBER = '<c r="A1"><f>B1*2</f><v>24</v></c>'
const TEXT = '<c r="B1" t="str"><f>"x" &amp; "y"</f><v>xy</v></c>'
const ERROR = '<c r="C1" t="e"><f>1/0</f><v>#DIV/0!</v></c>'
const BOOLEAN = '<c r="D1" t="b"><f>1=1</f><v>1</v></c>'

async function cellsOf(xml: string) {
  const wb = await readXlsx(await withCells(xml), { readStyles: true })
  return wb.sheets[0]!.cells!
}

describe("a cached formula result survives whatever its type", () => {
  it("number, string, error and boolean all arrive", async () => {
    const cells = await cellsOf(`${NUMBER}${TEXT}${ERROR}${BOOLEAN}`)

    expect(cells.get("0,0")?.formulaResult).toBe(24)
    expect(cells.get("0,1")?.formulaResult).toBe("xy")
    expect(cells.get("0,2")?.formulaResult).toBe("#DIV/0!")
    expect(cells.get("0,3")?.formulaResult).toBe(true)
  })

  it("the formula text arrives with it", async () => {
    const cells = await cellsOf(`${NUMBER}${TEXT}${ERROR}${BOOLEAN}`)

    expect(cells.get("0,1")?.formula).toBe('"x" & "y"')
    expect(cells.get("0,2")?.formula).toBe("1/0")
  })
})

describe("the round trip that was losing them", () => {
  it("readXlsx -> writeXlsx keeps every cached result", async () => {
    // This is the whole issue: the writer could always write these back,
    // so the loss was one-sided and silent.
    const first = await readXlsx(await withCells(`${NUMBER}${TEXT}${ERROR}${BOOLEAN}`), {
      readStyles: true,
    })

    const rewritten = await writeXlsx({
      sheets: first.sheets.map((s) => ({ name: s.name, rows: s.rows, cells: s.cells })),
    })
    const second = (await readXlsx(rewritten, { readStyles: true })).sheets[0]!.cells!

    expect(second.get("0,0")?.formulaResult).toBe(24)
    expect(second.get("0,1")?.formulaResult).toBe("xy")
    expect(second.get("0,2")?.formulaResult).toBe("#DIV/0!")
    expect(second.get("0,3")?.formulaResult).toBe(true)
  })

  it("so the rewritten file has a <v> under every <f>", async () => {
    const first = await readXlsx(await withCells(`${TEXT}${ERROR}`), { readStyles: true })
    const rewritten = await writeXlsx({
      sheets: first.sheets.map((s) => ({ name: s.name, rows: s.rows, cells: s.cells })),
    })

    const sheetXml = dec.decode(await new ZipReader(rewritten).extract("xl/worksheets/sheet1.xml"))

    // An `<f>` with no `<v>` is what anything that does not recalculate
    // reads as an empty cell.
    expect(sheetXml).not.toMatch(/<f>[^<]*<\/f><\/c>/)
    expect(sheetXml).toContain("xy")
    expect(sheetXml).toContain("#DIV/0!")
  })
})

describe("the type a formula cell reports", () => {
  it("is `formula`, whatever the cached result is", async () => {
    // It used to report `error` on the way in and `formula` on the way
    // back out. Both cannot be right, and the round trip is the side
    // with a second opinion.
    const cells = await cellsOf(`${NUMBER}${TEXT}${ERROR}${BOOLEAN}`)

    expect(cells.get("0,0")?.type).toBe("formula")
    expect(cells.get("0,1")?.type).toBe("formula")
    expect(cells.get("0,2")?.type).toBe("formula")
    expect(cells.get("0,3")?.type).toBe("formula")
  })

  it("but a hard-coded error is still an error, not a formula", async () => {
    // Excel writes these for a literal error value. Nothing about them
    // changed, and spotting an error by its value works either way.
    const cells = await cellsOf('<c r="A1" t="e"><v>#REF!</v></c>')

    expect(cells.get("0,0")?.type).toBe("error")
    expect(cells.get("0,0")?.formula).toBeUndefined()
  })

  it("and a plain boolean is still a boolean", async () => {
    // No style, no formula, no comment — so there is no `cells` entry to
    // hold a type. The value is the assertion.
    const wb = await readXlsx(await withCells('<c r="A1" t="b"><v>1</v></c>'))

    expect(wb.sheets[0]!.rows[0]![0]).toBe(true)
    expect(wb.sheets[0]!.cells?.get("0,0")).toBeUndefined()
  })
})

describe("the value stays where callers look for it", () => {
  it("rows carry the cached result, not the formula text", async () => {
    const wb = await readXlsx(await withCells(`${NUMBER}${TEXT}${ERROR}${BOOLEAN}`))

    expect(wb.sheets[0]!.rows[0]).toEqual([24, "xy", "#DIV/0!", true])
  })
})

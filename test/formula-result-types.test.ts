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
    expect(sheetXml).toContain('<c r="B1" t="str"><f>"x" &amp; "y"</f><v>xy</v></c>')
    // Not just the token: written as text, `t="str"`, it is a formula
    // that returned the string "#DIV/0!" — a different cell. See #599.
    expect(sheetXml).toContain('<c r="C1" t="e"><f>1/0</f><v>#DIV/0!</v></c>')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// #599 — `formulaResult` holds an error token as a plain string, so a
// formula whose cached result is `#DIV/0!` (`t="e"`) and a formula that
// returned the *text* "#DIV/0!" (`t="str"`) read back identically, and
// the writer turned the first into the second. ST_CellType (ECMA-376
// Part 1, §18.18.11) keeps them apart; `formulaResultType` now does too.
// ═══════════════════════════════════════════════════════════════════════

const ERROR_RESULT = '<c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c>'
const TEXT_LIKE_ERROR = '<c r="B1" t="str"><f>"#DIV/0!"</f><v>#DIV/0!</v></c>'
const HARD_CODED_ERROR = '<c r="C1" t="e"><v>#N/A</v></c>'

async function sheetXmlOf(data: Uint8Array): Promise<string> {
  return dec.decode(await new ZipReader(data).extract("xl/worksheets/sheet1.xml"))
}

describe("an error result is told apart from a text result", () => {
  it('reads `formulaResultType: "error"` only on the error', async () => {
    const cells = await cellsOf(`${ERROR_RESULT}${TEXT_LIKE_ERROR}${HARD_CODED_ERROR}`)

    expect(cells.get("0,0")).toMatchObject({
      type: "formula",
      formula: "1/0",
      formulaResult: "#DIV/0!",
      formulaResultType: "error",
    })
    expect(cells.get("0,1")?.formulaResult).toBe("#DIV/0!")
    expect(cells.get("0,1")).not.toHaveProperty("formulaResultType")
    // No formula, so no formula result to qualify: `type` says it.
    expect(cells.get("0,2")).toEqual({ value: "#N/A", type: "error" })
  })

  it("is absent for every other kind of cached result", async () => {
    const cells = await cellsOf(`${NUMBER}${TEXT}${BOOLEAN}`)

    for (const key of ["0,0", "0,1", "0,3"]) {
      expect(cells.get(key)).not.toHaveProperty("formulaResultType")
    }
  })

  it('readXlsx -> writeXlsx keeps `t="e"` and `t="str"` apart', async () => {
    const first = await readXlsx(
      await withCells(`${ERROR_RESULT}${TEXT_LIKE_ERROR}${HARD_CODED_ERROR}`),
    )
    const rewritten = await writeXlsx({
      sheets: first.sheets.map((s) => ({ name: s.name, rows: s.rows, cells: s.cells })),
    })
    const sheetXml = await sheetXmlOf(rewritten)

    expect(sheetXml).toContain('<c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c>')
    expect(sheetXml).toContain('<c r="B1" t="str"><f>"#DIV/0!"</f><v>#DIV/0!</v></c>')
    expect(sheetXml).toContain('<c r="C1" t="e"><v>#N/A</v></c>')

    const second = (await readXlsx(rewritten)).sheets[0]!.cells!
    expect(second.get("0,0")?.formulaResultType).toBe("error")
    expect(second.get("0,1")).not.toHaveProperty("formulaResultType")
  })

  it('writes an authored error result as `t="e"`', async () => {
    const out = await writeXlsx({
      sheets: [
        {
          name: "S",
          cells: new Map([
            ["0,0", { formula: "1/0", formulaResult: "#DIV/0!", formulaResultType: "error" }],
            ["0,1", { formula: '"#DIV/0!"', formulaResult: "#DIV/0!" }],
          ]),
        },
      ],
    })
    const sheetXml = await sheetXmlOf(out)

    expect(sheetXml).toContain('<c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c>')
    // The token alone is not taken as a signal: without the field, a
    // string result is text, as it always was.
    expect(sheetXml).toContain('<c r="B1" t="str"><f>"#DIV/0!"</f><v>#DIV/0!</v></c>')
  })

  it("leaves a non-string result as it was", async () => {
    // An error token is text, so `t="e"` with `<v>42</v>` is not a cell
    // anything could read. The field qualifies a string result only, as
    // its doc comment says.
    const out = await writeXlsx({
      sheets: [
        {
          name: "S",
          cells: new Map([
            ["0,0", { formula: "B1*2", formulaResult: 42, formulaResultType: "error" }],
          ]),
        },
      ],
    })

    expect(await sheetXmlOf(out)).toContain('<c r="A1"><f>B1*2</f><v>42</v></c>')
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

describe("a shared formula's error result", () => {
  // The cells a shared formula covers carry no formula text of their own
  // — `<f t="shared" si="0"/>` — so the reader gives them `formula: ""`.
  // That is still a formula, and its cached error is still an error.
  const MASTER = '<c r="A1" t="e"><f t="shared" ref="A1:B1" si="0">1/0</f><v>#DIV/0!</v></c>'
  const SLAVE = '<c r="B1" t="e"><f t="shared" si="0"/><v>#DIV/0!</v></c>'

  it('reads `formulaResultType: "error"` on the covered cell too', async () => {
    const cells = await cellsOf(`${MASTER}${SLAVE}`)

    expect(cells.get("0,1")).toMatchObject({
      formula: "",
      formulaType: "shared",
      formulaResult: "#DIV/0!",
      formulaResultType: "error",
    })
  })

  it('keeps the covered cell\'s cached result and `t="e"` through a rewrite', async () => {
    const first = await readXlsx(await withCells(`${MASTER}${SLAVE}`))
    const rewritten = await writeXlsx({
      sheets: first.sheets.map((s) => ({ name: s.name, rows: s.rows, cells: s.cells })),
    })

    expect(await sheetXmlOf(rewritten)).toContain(
      '<c r="B1" t="e"><f t="shared" si="0"/><v>#DIV/0!</v></c>',
    )
  })
})

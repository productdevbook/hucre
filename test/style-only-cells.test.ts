import { cellError } from "../src/cell-error"
import { describe, expect, it } from "vitest"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { writeCsv } from "../src/csv/writer"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"

// ═══════════════════════════════════════════════════════════════════════
// #492 — Excel writes a self-closing `<c r="WVF45" s="3"/>` for every
// position formatting was ever applied to. A real packing-list workbook,
// edited by people over years, had 145,315 of them against 197 values:
// `rows` came back 45 x 16,126, and `writeCsv` of that was 727,211 bytes
// — 99.75% bare commas — from 1.8 KB of data.
//
// Under the default `readStyles: false` those cells contribute nothing:
// their styles are not read, so their only effect is null padding the
// caller cannot tell from never-written cells anyway. With
// `readStyles: true` they carry information and still count.
//
// This is not the #394 class of bug. Interior positions are untouched —
// a value at N5 still lands at rows[4][13] with nulls before it. Only
// the trailing bounding box shrinks to the last cell carrying data.
// ═══════════════════════════════════════════════════════════════════════

const enc = new TextEncoder()
const dec = new TextDecoder()

/** A workbook whose sheet body is exactly the rows given. */
async function sheetWith(rowsXml: string, dimension = "A1:WVF45"): Promise<Uint8Array> {
  const base = await writeXlsx({ sheets: [{ name: "S", rows: [["seed"]] }] })
  const all = await new ZipReader(base).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    zw.add(
      name,
      name === "xl/worksheets/sheet1.xml"
        ? enc.encode(
            dec
              .decode(data)
              .replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="${dimension}"/>`)
              .replace(/<sheetData>.*<\/sheetData>/, `<sheetData>${rowsXml}</sheetData>`),
          )
        : data,
    )
  }
  return zw.build()
}

/** The issue's repro: one value at A1, one style-only cell at WVF45. */
const REPRO =
  '<row r="1"><c r="A1" t="str"><v>hello</v></c></row>' + '<row r="45"><c r="WVF45" s="0"/></row>'

describe("a style-only cell does not inflate the grid", () => {
  it("the sheet is as wide as its data, not as its formatting", async () => {
    const wb = await readXlsx(await sheetWith(REPRO))
    const sheet = wb.sheets[0]!

    expect(sheet.rows).toHaveLength(1)
    expect(sheet.rows[0]).toEqual(["hello"])
  })

  it("so the CSV of it is the data, not 700 KB of commas", async () => {
    // The number in the issue: 727,211 bytes, ~99.75% bare commas.
    const wb = await readXlsx(await sheetWith(REPRO))

    expect(writeCsv(wb.sheets[0]!.rows).length).toBeLessThan(100)
  })

  it("counts them when readStyles is on, because then they carry something", async () => {
    const wb = await readXlsx(await sheetWith(REPRO), { readStyles: true })

    // 16,126 columns wide — WVF is column 16,125, zero-based.
    expect(wb.sheets[0]!.rows).toHaveLength(45)
    expect(wb.sheets[0]!.rows[0]).toHaveLength(16126)
    expect(wb.sheets[0]!.cells?.get("44,16125")).toBeDefined()
  })
})

describe("interior positions are untouched", () => {
  it("a gap before a value is still a gap", async () => {
    // The #394 class of bug this must not become: only the *trailing*
    // box shrinks. A value at D1 still sits at index 3.
    const wb = await readXlsx(
      await sheetWith('<row r="1"><c r="D1" t="str"><v>d</v></c></row>', "A1:Z10"),
    )

    expect(wb.sheets[0]!.rows[0]).toEqual([null, null, null, "d"])
  })

  it("an interior style-only cell keeps its column, because a later value does", async () => {
    const wb = await readXlsx(
      await sheetWith(
        '<row r="1"><c r="A1" t="str"><v>a</v></c><c r="B1" s="0"/><c r="C1" t="str"><v>c</v></c></row>',
        "A1:Z10",
      ),
    )

    expect(wb.sheets[0]!.rows[0]).toEqual(["a", null, "c"])
  })

  it("an empty row between two populated ones survives", async () => {
    const wb = await readXlsx(
      await sheetWith(
        '<row r="1"><c r="A1" t="str"><v>a</v></c></row>' +
          '<row r="3"><c r="A3" t="str"><v>c</v></c></row>',
        "A1:Z10",
      ),
    )

    expect(wb.sheets[0]!.rows).toEqual([["a"], [null], ["c"]])
  })
})

describe("what still counts as carrying data", () => {
  it("a formula with no cached result", async () => {
    const wb = await readXlsx(await sheetWith('<row r="1"><c r="C1"><f>SUM(A1:B1)</f></c></row>'))

    expect(wb.sheets[0]!.rows[0]).toHaveLength(3)
  })

  it("an error value", async () => {
    const wb = await readXlsx(await sheetWith('<row r="1"><c r="C1" t="e"><v>#REF!</v></c></row>'))

    expect(wb.sheets[0]!.rows[0]![2]).toEqual(cellError("#REF!"))
  })

  it("an inline string", async () => {
    const wb = await readXlsx(
      await sheetWith('<row r="1"><c r="C1" t="inlineStr"><is><t>x</t></is></c></row>'),
    )

    expect(wb.sheets[0]!.rows[0]![2]).toBe("x")
  })

  it("a value of zero, false or an empty string", async () => {
    // The trap in any "is it empty?" test. All three are data.
    const wb = await readXlsx(
      await sheetWith(
        '<row r="1"><c r="A1"><v>0</v></c><c r="B1" t="b"><v>0</v></c>' +
          '<c r="C1" t="inlineStr"><is><t></t></is></c><c r="D1" t="str"><v>x</v></c></row>',
      ),
    )

    expect(wb.sheets[0]!.rows[0]).toHaveLength(4)
    expect(wb.sheets[0]!.rows[0]![0]).toBe(0)
    expect(wb.sheets[0]!.rows[0]![1]).toBe(false)
  })
})

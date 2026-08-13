// A cell object written where a value goes — `rows: [[{ value, style }]]`.
//
// Before #433 the buffered writers read it as a value and emitted an
// *empty* cell: the style, the formula and the value all gone, with no
// error. `XlsxStreamWriter.addRow` had accepted the shape since it
// existed, so the two halves of the library disagreed about what a row
// may hold.

import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { write } from "../src/defter"
import type { WriteSheet } from "../src/_types"

async function part(buf: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(buf)
  return new TextDecoder().decode(await zip.extract(path))
}

const inlineSheet: WriteSheet = {
  name: "S",
  rows: [
    ["plain", { value: "wrapped", style: { alignment: { wrapText: true } } }],
    [{ value: 1234.5, style: { numFmt: "#,##0.00" } }, { formula: "A2*2" }],
  ],
}

/** The same sheet said the old way, for the writers to agree with. */
const mapSheet: WriteSheet = {
  name: "S",
  rows: [
    ["plain", "wrapped"],
    [1234.5, null],
  ],
  cells: new Map([
    ["0,1", { value: "wrapped", style: { alignment: { wrapText: true } } }],
    ["1,0", { value: 1234.5, style: { numFmt: "#,##0.00" } }],
    ["1,1", { formula: "A2*2" }],
  ]),
}

describe("cell objects written inline in rows", () => {
  it("keeps the value that used to be dropped (xlsx)", async () => {
    const buf = await writeXlsx({ sheets: [inlineSheet] })
    const wb = await readXlsx(buf, { readStyles: true })
    const rows = wb.sheets[0]!.rows

    expect(rows[0]![1]).toBe("wrapped")
    expect(rows[1]![0]).toBe(1234.5)
  })

  it("keeps the value that used to be dropped (ods)", async () => {
    const buf = await writeOds({ sheets: [inlineSheet] })
    const wb = await readOds(buf)

    expect(wb.sheets[0]!.rows[0]![1]).toBe("wrapped")
    expect(wb.sheets[0]!.rows[1]![0]).toBe(1234.5)
  })

  it("carries the style, not just the value", async () => {
    const buf = await writeXlsx({ sheets: [inlineSheet] })
    const styles = await part(buf, "xl/styles.xml")
    const sheet = await part(buf, "xl/worksheets/sheet1.xml")

    // A wrap alignment and a number format both had to be registered.
    expect(styles).toContain('applyAlignment="true"')
    expect(styles).toContain('wrapText="true"')
    expect(styles).toContain("#,##0.00")
    // …and reach the cells that asked for them.
    expect(sheet).toMatch(/<c r="B1"[^>]*s="[1-9]/)
    expect(sheet).toMatch(/<c r="A2"[^>]*s="[1-9]/)
  })

  it("carries a formula", async () => {
    const buf = await writeXlsx({ sheets: [inlineSheet] })
    expect(await part(buf, "xl/worksheets/sheet1.xml")).toContain("<f>A2*2</f>")
  })

  it("is the same document as the cells map spelling", async () => {
    const inline = await part(
      await writeXlsx({ sheets: [inlineSheet] }),
      "xl/worksheets/sheet1.xml",
    )
    const viaMap = await part(await writeXlsx({ sheets: [mapSheet] }), "xl/worksheets/sheet1.xml")
    expect(inline).toBe(viaMap)

    const odsInline = await part(await writeOds({ sheets: [inlineSheet] }), "content.xml")
    const odsMap = await part(await writeOds({ sheets: [mapSheet] }), "content.xml")
    expect(odsInline).toBe(odsMap)
  })

  it("lets an explicit cells entry win over the inline one", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [[{ value: "inline", style: { font: { bold: true } } }]],
          cells: new Map([["0,0", { value: "explicit" }]]),
        },
      ],
    })
    const wb = await readXlsx(buf)
    expect(wb.sheets[0]!.rows[0]![0]).toBe("explicit")
  })

  it("does not mistake a Date for a cell object", async () => {
    const when = new Date(Date.UTC(2020, 0, 15))
    const buf = await writeXlsx({ sheets: [{ name: "S", rows: [[when]] }] })
    const wb = await readXlsx(buf)
    expect(wb.sheets[0]!.rows[0]![0]).toBeInstanceOf(Date)
  })

  it("does not mistake a hyperlink value for a cell object", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          columns: [{ key: "a", header: "A" }],
          data: [{ a: { text: "hucre", hyperlink: "https://example.com" } }],
        },
      ],
    })
    const wb = await readXlsx(buf)
    expect(wb.sheets[0]!.rows[1]![0]).toBe("hucre")
  })

  it("reaches the round-trip writer too", async () => {
    const base = await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] })
    const wb = await openXlsx(base)
    wb.sheets[0]!.rows = [[{ value: "b", style: { font: { bold: true } } }]] as never
    const out = await saveXlsx(wb, {})
    const back = await readXlsx(out)
    expect(back.sheets[0]!.rows[0]![0]).toBe("b")
  })

  it("reduces to the value for the formats that carry only values", async () => {
    const csv = await write({ sheets: [inlineSheet], format: "csv" })
    const text = new TextDecoder().decode(csv as Uint8Array)
    expect(text).toContain("wrapped")
    expect(text).toContain("1234.5")
    expect(text).not.toContain("object Object")
  })

  it("leaves a sheet of plain values exactly as it was", async () => {
    // The scan must not copy a grid it found nothing in — the same array
    // instance is what proves it.
    const rows = [["a", 1]]
    const sheet: WriteSheet = { name: "S", rows }
    await writeXlsx({ sheets: [sheet] })
    expect(sheet.rows).toBe(rows)
    expect(sheet.cells).toBeUndefined()
  })

  it("does not mutate the caller's sheet when it does split", async () => {
    const rows = [[{ value: "x", style: { font: { bold: true } } }]]
    const sheet: WriteSheet = { name: "S", rows }
    await writeXlsx({ sheets: [sheet] })
    expect(sheet.rows).toBe(rows)
    expect(sheet.rows![0]![0]).toEqual({ value: "x", style: { font: { bold: true } } })
    expect(sheet.cells).toBeUndefined()
  })
})

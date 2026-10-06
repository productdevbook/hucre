import { describe, expect, it } from "vitest"
import { read } from "../src/defter"
import { fromHtml } from "../src/export/html-import"
import { readOds } from "../src/ods/reader"
import { writeOds } from "../src/ods/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"

// `Sheet.rows` is documented as a dense rectangle. readXlsx delivered
// one; readOds, fromHtml and the CSV path of read() returned [] for an
// empty row and left a short line short, so one sheet read three ways
// had three shapes. PARITY.md used to record the disagreement; the
// readers now agree and this pins it.
describe("every Sheet reader returns the same rectangle", () => {
  const rows = [
    ["a", 1, true],
    [null, null, null],
    ["c", null, null],
  ]

  it("xlsx, ods and csv agree through read()", async () => {
    const xlsx = await read(await writeXlsx({ sheets: [{ name: "S", rows }] }))
    const ods = await read(await writeOds({ sheets: [{ name: "S", rows }] }))
    const csv = await read(new TextEncoder().encode("a,1,true\n\nc\n"))
    expect(xlsx.sheets[0]!.rows).toEqual(rows)
    expect(ods.sheets[0]!.rows).toEqual(rows)
    // CSV carries text: 1 and true stay strings without typeInference, and
    // an empty line is one empty field — the padding is the two nulls.
    expect(csv.sheets[0]!.rows).toEqual([
      ["a", "1", "true"],
      ["", null, null],
      ["c", null, null],
    ])
  })

  it("readOds pads an empty row and a short row to what readXlsx returns", async () => {
    const fromOds = await readOds(await writeOds({ sheets: [{ name: "S", rows }] }))
    const fromXlsx = await readXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }))
    expect(fromOds.sheets[0]!.rows[1]).toEqual([null, null, null])
    expect(fromOds.sheets[0]!.rows[2]).toEqual(["c", null, null])
    expect(fromOds.sheets[0]!.rows).toEqual(fromXlsx.sheets[0]!.rows)
  })

  it("fromHtml pads a short row", () => {
    const sheet = fromHtml("<table><tr><td>a</td><td>b</td></tr><tr><td>c</td></tr></table>")
    expect(sheet.rows).toEqual([
      ["a", "b"],
      ["c", null],
    ])
  })
})

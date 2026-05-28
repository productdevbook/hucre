import { describe, expect, it } from "vitest"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { readOds } from "../src/ods/reader"
import { writeOds } from "../src/ods/writer"

async function makeXlsx(
  sheets: Array<{
    name: string
    rows: (string | number)[][]
    hidden?: boolean
    veryHidden?: boolean
  }>,
): Promise<Uint8Array> {
  return await writeXlsx({ sheets })
}

describe("ReadOptions.sheets — predicate (XLSX)", () => {
  it("includes sheets when predicate returns true", async () => {
    const buf = await makeXlsx([
      { name: "A", rows: [["a"], [1]] },
      { name: "B", rows: [["b"], [2]] },
      { name: "C", rows: [["c"], [3]] },
    ])

    const wb = await readXlsx(buf, {
      sheets: (info) => info.name !== "B",
    })

    expect(wb.sheets.map((s) => s.name)).toEqual(["A", "C"])
  })

  it("exposes 0-based index to the predicate", async () => {
    const buf = await makeXlsx([
      { name: "A", rows: [["a"]] },
      { name: "B", rows: [["b"]] },
      { name: "C", rows: [["c"]] },
    ])

    const seen: number[] = []
    await readXlsx(buf, {
      sheets: (info, i) => {
        expect(info.index).toBe(i)
        seen.push(i)
        return i === 1
      },
    })

    expect(seen).toEqual([0, 1, 2])
  })

  it('exposes hidden state from XLSX <sheet state="hidden">', async () => {
    const buf = await makeXlsx([
      { name: "Visible", rows: [["v"]] },
      { name: "Hidden", rows: [["h"]], hidden: true },
      { name: "VeryHidden", rows: [["vh"]], veryHidden: true },
    ])

    const wb = await readXlsx(buf, {
      sheets: (info) => !info.hidden && !info.veryHidden,
    })

    expect(wb.sheets.map((s) => s.name)).toEqual(["Visible"])
  })

  it("returns empty workbook when predicate rejects every sheet", async () => {
    const buf = await makeXlsx([
      { name: "A", rows: [["a"]] },
      { name: "B", rows: [["b"]] },
    ])

    const wb = await readXlsx(buf, { sheets: () => false })
    expect(wb.sheets).toEqual([])
  })

  it("array form still works (backwards compatible)", async () => {
    const buf = await makeXlsx([
      { name: "A", rows: [["a"]] },
      { name: "B", rows: [["b"]] },
      { name: "C", rows: [["c"]] },
    ])

    const wb = await readXlsx(buf, { sheets: ["A", "C"] })
    expect(wb.sheets.map((s) => s.name)).toEqual(["A", "C"])
  })

  it("empty array still reads all sheets (legacy behavior)", async () => {
    const buf = await makeXlsx([
      { name: "A", rows: [["a"]] },
      { name: "B", rows: [["b"]] },
    ])

    const wb = await readXlsx(buf, { sheets: [] })
    expect(wb.sheets).toHaveLength(2)
  })
})

describe("ReadOptions.sheets — predicate (ODS)", () => {
  it("filters ODS sheets by predicate", async () => {
    const buf = await writeOds({
      sheets: [
        { name: "Keep", rows: [["k"], [1]] },
        { name: "Skip", rows: [["s"], [2]] },
        { name: "Keep2", rows: [["k2"], [3]] },
      ],
    })

    const wb = await readOds(buf, {
      sheets: (info) => info.name.startsWith("Keep"),
    })

    expect(wb.sheets.map((s) => s.name)).toEqual(["Keep", "Keep2"])
  })

  it("ODS predicate sees name and index but no visibility metadata", async () => {
    const buf = await writeOds({
      sheets: [
        { name: "A", rows: [["a"]] },
        { name: "B", rows: [["b"]] },
      ],
    })

    let captured: { name: string; hidden?: boolean; veryHidden?: boolean } | null = null
    await readOds(buf, {
      sheets: (info) => {
        if (info.index === 0) captured = info
        return true
      },
    })

    expect(captured).not.toBeNull()
    expect(captured!.name).toBe("A")
    expect(captured!.hidden).toBeUndefined()
    expect(captured!.veryHidden).toBeUndefined()
  })
})

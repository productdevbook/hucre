import { describe, expect, it } from "vitest"
import { readOdsObjects, writeOdsObjects } from "../src/ods/index"
import { writeOds } from "../src/ods/writer"

async function makeWorkbook(
  rows: (string | number | boolean | Date | null)[][],
  sheetName = "Sheet1",
): Promise<Uint8Array> {
  return await writeOds({ sheets: [{ name: sheetName, rows }] })
}

describe("readOdsObjects", () => {
  it("returns data and headers from a simple sheet", async () => {
    const buf = await makeWorkbook([
      ["Name", "Age", "City"],
      ["Alice", 30, "Istanbul"],
      ["Bob", 25, "Ankara"],
    ])

    const result = await readOdsObjects(buf)

    expect(result.headers).toEqual(["Name", "Age", "City"])
    expect(result.data).toEqual([
      { Name: "Alice", Age: 30, City: "Istanbul" },
      { Name: "Bob", Age: 25, City: "Ankara" },
    ])
  })

  it("returns empty result for empty sheet", async () => {
    const buf = await makeWorkbook([])
    const result = await readOdsObjects(buf)
    expect(result.headers).toEqual([])
    expect(result.data).toEqual([])
  })

  it("respects custom headerRow", async () => {
    const buf = await makeWorkbook([["report"], ["Name", "Age"], ["Alice", 30]])

    const result = await readOdsObjects(buf, { headerRow: 1 })
    expect(result.headers).toEqual(["Name", "Age"])
    expect(result.data).toEqual([{ Name: "Alice", Age: 30 }])
  })

  it("selects sheet by name", async () => {
    const buf = await writeOds({
      sheets: [
        { name: "First", rows: [["a"], [1]] },
        { name: "Data", rows: [["b"], [2]] },
      ],
    })

    const result = await readOdsObjects(buf, { sheet: "Data" })
    expect(result.headers).toEqual(["b"])
    expect(result.data).toEqual([{ b: 2 }])
  })

  it("throws when sheet index is out of range", async () => {
    const buf = await makeWorkbook([["a"], [1]])
    await expect(readOdsObjects(buf, { sheet: 5 })).rejects.toThrow(/out of range/)
  })

  it("applies transformHeader and transformValue", async () => {
    const buf = await makeWorkbook([
      ["Name", "Score"],
      ["Alice", 30],
    ])

    const result = await readOdsObjects(buf, {
      transformHeader: (h) => h.toLowerCase(),
      transformValue: (v, h) => (h === "score" && typeof v === "number" ? v * 2 : v),
    })
    expect(result.headers).toEqual(["name", "score"])
    expect(result.data[0]).toEqual({ name: "Alice", score: 60 })
  })

  it("respects maxRows", async () => {
    const buf = await makeWorkbook([["A"], [1], [2], [3]])
    const result = await readOdsObjects(buf, { maxRows: 2 })
    expect(result.data).toHaveLength(2)
  })
})

describe("writeOdsObjects", () => {
  it("round-trips through readOdsObjects", async () => {
    const data = [
      { Name: "Alice", Age: 30 },
      { Name: "Bob", Age: 25 },
    ]
    const buf = await writeOdsObjects(data)

    const round = await readOdsObjects(buf)
    expect(round.headers).toEqual(["Name", "Age"])
    expect(round.data).toEqual(data)
  })

  it("uses explicit headers for column order", async () => {
    const buf = await writeOdsObjects([{ a: 1, b: 2, c: 3 }], { headers: ["c", "b", "a"] })

    const round = await readOdsObjects(buf)
    expect(round.headers).toEqual(["c", "b", "a"])
  })

  it("writes specified sheet name", async () => {
    const buf = await writeOdsObjects([{ a: 1 }], { sheetName: "Custom" })
    const round = await readOdsObjects(buf, { sheet: "Custom" })
    expect(round.data).toEqual([{ a: 1 }])
  })
})

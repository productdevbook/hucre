import { describe, expect, it } from "vitest"
import { readXlsxObjects, writeXlsxObjects } from "../src/xlsx/index"
import { writeXlsx } from "../src/xlsx/writer"

async function makeWorkbook(
  rows: (string | number | boolean | Date | null)[][],
  sheetName = "Sheet1",
): Promise<Uint8Array> {
  return await writeXlsx({ sheets: [{ name: sheetName, rows }] })
}

describe("readXlsxObjects", () => {
  it("returns data and headers from a simple sheet", async () => {
    const buf = await makeWorkbook([
      ["Name", "Age", "City"],
      ["Alice", 30, "Istanbul"],
      ["Bob", 25, "Ankara"],
    ])

    const result = await readXlsxObjects(buf)

    expect(result.headers).toEqual(["Name", "Age", "City"])
    expect(result.data).toEqual([
      { Name: "Alice", Age: 30, City: "Istanbul" },
      { Name: "Bob", Age: 25, City: "Ankara" },
    ])
  })

  it("returns empty result for empty sheet", async () => {
    const buf = await makeWorkbook([])
    const result = await readXlsxObjects(buf)
    expect(result.headers).toEqual([])
    expect(result.data).toEqual([])
  })

  it("returns headers but no data when only header row exists", async () => {
    const buf = await makeWorkbook([["A", "B", "C"]])
    const result = await readXlsxObjects(buf)
    expect(result.headers).toEqual(["A", "B", "C"])
    expect(result.data).toEqual([])
  })

  it("respects custom headerRow", async () => {
    const buf = await makeWorkbook([["report"], ["generated 2025"], ["Name", "Age"], ["Alice", 30]])

    const result = await readXlsxObjects(buf, { headerRow: 2 })
    expect(result.headers).toEqual(["Name", "Age"])
    expect(result.data).toEqual([{ Name: "Alice", Age: 30 }])
  })

  it("selects sheet by index", async () => {
    const buf = await writeXlsx({
      sheets: [
        { name: "First", rows: [["a"], [1]] },
        { name: "Second", rows: [["b"], [2]] },
      ],
    })

    const result = await readXlsxObjects(buf, { sheet: 1 })
    expect(result.headers).toEqual(["b"])
    expect(result.data).toEqual([{ b: 2 }])
  })

  it("selects sheet by name", async () => {
    const buf = await writeXlsx({
      sheets: [
        { name: "First", rows: [["a"], [1]] },
        { name: "Data", rows: [["b"], [2]] },
      ],
    })

    const result = await readXlsxObjects(buf, { sheet: "Data" })
    expect(result.headers).toEqual(["b"])
    expect(result.data).toEqual([{ b: 2 }])
  })

  it("throws when sheet index is out of range", async () => {
    const buf = await makeWorkbook([["a"], [1]])
    await expect(readXlsxObjects(buf, { sheet: 5 })).rejects.toThrow(/out of range/)
  })

  it("throws when sheet name does not exist", async () => {
    const buf = await makeWorkbook([["a"], [1]])
    await expect(readXlsxObjects(buf, { sheet: "Nope" })).rejects.toThrow(/not found/)
  })

  it("skips empty rows by default", async () => {
    const buf = await makeWorkbook([
      ["Name", "Age"],
      ["Alice", 30],
      [null, null],
      ["Bob", 25],
    ])

    const result = await readXlsxObjects(buf)
    expect(result.data).toEqual([
      { Name: "Alice", Age: 30 },
      { Name: "Bob", Age: 25 },
    ])
  })

  it("keeps empty rows when skipEmptyRows is false", async () => {
    // Construct a sheet that round-trips with an empty middle row.
    const buf = await makeWorkbook([
      ["Name", "Age"],
      ["Alice", 30],
      ["", ""],
      ["Bob", 25],
    ])

    const result = await readXlsxObjects(buf, { skipEmptyRows: false })
    // The middle empty row must be preserved.
    expect(result.data.length).toBeGreaterThanOrEqual(3)
    const alice = result.data.find((r) => r.Name === "Alice")
    const bob = result.data.find((r) => r.Name === "Bob")
    expect(alice).toBeDefined()
    expect(bob).toBeDefined()
  })

  it("applies transformHeader", async () => {
    const buf = await makeWorkbook([
      ["First Name", "Last Name"],
      ["Alice", "Smith"],
    ])

    const result = await readXlsxObjects(buf, {
      transformHeader: (h) => h.toLowerCase().replace(/ /g, "_"),
    })
    expect(result.headers).toEqual(["first_name", "last_name"])
    expect(result.data[0]).toEqual({ first_name: "Alice", last_name: "Smith" })
  })

  it("applies transformValue", async () => {
    const buf = await makeWorkbook([
      ["Name", "Score"],
      ["Alice", 30],
      ["Bob", 25],
    ])

    const result = await readXlsxObjects(buf, {
      transformValue: (v, h) => (h === "Score" && typeof v === "number" ? v * 2 : v),
    })
    expect(result.data).toEqual([
      { Name: "Alice", Score: 60 },
      { Name: "Bob", Score: 50 },
    ])
  })

  it("respects maxRows", async () => {
    const buf = await makeWorkbook([["A"], [1], [2], [3], [4]])

    const result = await readXlsxObjects(buf, { maxRows: 2 })
    expect(result.data).toHaveLength(2)
  })
})

describe("writeXlsxObjects", () => {
  it("derives headers from the first object's keys", async () => {
    const buf = await writeXlsxObjects([
      { Name: "Alice", Age: 30 },
      { Name: "Bob", Age: 25 },
    ])

    const round = await readXlsxObjects(buf)
    expect(round.headers).toEqual(["Name", "Age"])
    expect(round.data).toEqual([
      { Name: "Alice", Age: 30 },
      { Name: "Bob", Age: 25 },
    ])
  })

  it("uses explicit headers for column order", async () => {
    const buf = await writeXlsxObjects([{ a: 1, b: 2, c: 3 }], { headers: ["c", "b", "a"] })

    const round = await readXlsxObjects(buf)
    expect(round.headers).toEqual(["c", "b", "a"])
    expect(round.data[0]).toEqual({ c: 3, b: 2, a: 1 })
  })

  it("writes specified sheet name", async () => {
    const buf = await writeXlsxObjects([{ a: 1 }], { sheetName: "Custom" })
    const round = await readXlsxObjects(buf, { sheet: "Custom" })
    expect(round.data).toEqual([{ a: 1 }])
  })

  it("handles empty data array", async () => {
    const buf = await writeXlsxObjects([])
    const round = await readXlsxObjects(buf)
    expect(round.data).toEqual([])
  })

  it("treats missing keys as null", async () => {
    const buf = await writeXlsxObjects([{ a: 1, b: 2 }, { a: 3 }], { headers: ["a", "b"] })
    const round = await readXlsxObjects(buf)
    expect(round.data).toEqual([
      { a: 1, b: 2 },
      { a: 3, b: null },
    ])
  })
})

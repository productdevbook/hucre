import { describe, it, expect } from "vitest"
import { read, write, readObjects, writeObjects } from "../src/defter"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { readOds } from "../src/ods/reader"
import { writeOds } from "../src/ods/writer"
import type { CellValue } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

/** Create a minimal XLSX buffer via writeXlsx for testing */
async function makeXlsx(rows: CellValue[][], sheetName = "Sheet1"): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [{ name: sheetName, rows }],
  })
}

/** Create a minimal ODS buffer via writeOds for testing */
async function makeOds(rows: CellValue[][], sheetName = "Sheet1"): Promise<Uint8Array> {
  return writeOds({
    sheets: [{ name: sheetName, rows }],
  })
}

// ── read() ──────────────────────────────────────────────────────────

describe("read()", () => {
  it("reads XLSX input and returns a Workbook", async () => {
    const xlsx = await makeXlsx([
      ["Name", "Age"],
      ["Alice", 30],
      ["Bob", 25],
    ])

    const workbook = await read(xlsx)

    expect(workbook.sheets).toHaveLength(1)
    expect(workbook.sheets[0]!.name).toBe("Sheet1")
    expect(workbook.sheets[0]!.rows).toHaveLength(3)
    expect(workbook.sheets[0]!.rows[0]).toEqual(["Name", "Age"])
    expect(workbook.sheets[0]!.rows[1]).toEqual(["Alice", 30])
  })

  it("reads ODS input and returns a Workbook", async () => {
    const ods = await makeOds([
      ["Product", "Price"],
      ["Widget", 9.99],
    ])

    const workbook = await read(ods)

    expect(workbook.sheets).toHaveLength(1)
    expect(workbook.sheets[0]!.name).toBe("Sheet1")
    expect(workbook.sheets[0]!.rows).toHaveLength(2)
    expect(workbook.sheets[0]!.rows[0]).toEqual(["Product", "Price"])
    expect(workbook.sheets[0]!.rows[1]).toEqual(["Widget", 9.99])
  })

  it("auto-detects XLSX vs ODS format correctly", async () => {
    const xlsx = await makeXlsx([["xlsx"]])
    const ods = await makeOds([["ods"]])

    // Both should parse without error using the unified read()
    const wbXlsx = await read(xlsx)
    const wbOds = await read(ods)

    expect(wbXlsx.sheets[0]!.rows[0]![0]).toBe("xlsx")
    expect(wbOds.sheets[0]!.rows[0]![0]).toBe("ods")
  })

  it("reads from ArrayBuffer input", async () => {
    const xlsx = await makeXlsx([["Hello"]])
    const arrayBuffer = xlsx.buffer.slice(xlsx.byteOffset, xlsx.byteOffset + xlsx.byteLength)

    const workbook = await read(new Uint8Array(arrayBuffer))

    expect(workbook.sheets[0]!.rows[0]![0]).toBe("Hello")
  })

  it("passes ReadOptions through to the underlying reader", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        { name: "First", rows: [["A"]] },
        { name: "Second", rows: [["B"]] },
      ],
    })

    const workbook = await read(xlsx, { sheets: ["Second"] })

    expect(workbook.sheets).toHaveLength(1)
    expect(workbook.sheets[0]!.name).toBe("Second")
  })

  it("throws on invalid (non-ZIP) input", async () => {
    const garbage = new Uint8Array([0x00, 0x01, 0x02, 0x03])
    await expect(read(garbage)).rejects.toThrow()
  })
})

// ── write() ─────────────────────────────────────────────────────────

describe("write()", () => {
  it("writes XLSX format by default", async () => {
    const output = await write({
      sheets: [{ name: "Sheet1", rows: [["Hello", 42]] }],
    })

    // Verify it's valid XLSX by reading it back
    const workbook = await readXlsx(output)
    expect(workbook.sheets[0]!.rows[0]).toEqual(["Hello", 42])
  })

  it("writes XLSX when format is explicitly 'xlsx'", async () => {
    const output = await write({
      sheets: [{ name: "Test", rows: [["A", "B"]] }],
      format: "xlsx",
    })

    const workbook = await readXlsx(output)
    expect(workbook.sheets[0]!.name).toBe("Test")
  })

  it("writes ODS when format is 'ods'", async () => {
    const output = await write({
      sheets: [{ name: "OdsSheet", rows: [["Value", 100]] }],
      format: "ods",
    })

    const workbook = await readOds(output)
    expect(workbook.sheets[0]!.name).toBe("OdsSheet")
    expect(workbook.sheets[0]!.rows[0]).toEqual(["Value", 100])
  })

  it("supports multiple sheets", async () => {
    const output = await write({
      sheets: [
        { name: "Sheet1", rows: [["A"]] },
        { name: "Sheet2", rows: [["B"]] },
      ],
    })

    const workbook = await readXlsx(output)
    expect(workbook.sheets).toHaveLength(2)
    expect(workbook.sheets[0]!.name).toBe("Sheet1")
    expect(workbook.sheets[1]!.name).toBe("Sheet2")
  })
})

// ── readObjects() ───────────────────────────────────────────────────

describe("readObjects()", () => {
  it("returns array of objects with headers as keys", async () => {
    const xlsx = await makeXlsx([
      ["Name", "Age", "Active"],
      ["Alice", 30, true],
      ["Bob", 25, false],
    ])

    const objects = await readObjects(xlsx)

    expect(objects).toHaveLength(2)
    expect(objects[0]).toEqual({ Name: "Alice", Age: 30, Active: true })
    expect(objects[1]).toEqual({ Name: "Bob", Age: 25, Active: false })
  })

  it("returns empty array for empty sheet", async () => {
    const xlsx = await makeXlsx([])
    const objects = await readObjects(xlsx)
    expect(objects).toEqual([])
  })

  it("returns empty array for header-only sheet", async () => {
    const xlsx = await makeXlsx([["Name", "Age"]])
    const objects = await readObjects(xlsx)
    expect(objects).toEqual([])
  })

  it("handles null values in data rows", async () => {
    const xlsx = await makeXlsx([
      ["Name", "Score"],
      ["Alice", null],
      [null, 100],
    ])

    const objects = await readObjects(xlsx)

    expect(objects).toHaveLength(2)
    expect(objects[0]).toEqual({ Name: "Alice", Score: null })
    expect(objects[1]).toEqual({ Name: null, Score: 100 })
  })

  it("works with ODS input", async () => {
    const ods = await makeOds([
      ["Product", "Price"],
      ["Pen", 1.5],
    ])

    const objects = await readObjects(ods)

    expect(objects).toHaveLength(1)
    expect(objects[0]).toEqual({ Product: "Pen", Price: 1.5 })
  })

  it("skips empty-string headers", async () => {
    const xlsx = await makeXlsx([
      ["Name", "", "Age"],
      ["Alice", "ignore", 30],
    ])

    const objects = await readObjects(xlsx)

    expect(objects).toHaveLength(1)
    // Empty header key should be skipped
    expect(objects[0]!["Name"]).toBe("Alice")
    expect(objects[0]!["Age"]).toBe(30)
    expect(objects[0]![""]).toBeUndefined()
  })

  it("returns empty array when workbook has no sheets", async () => {
    // Edge case: workbook with empty sheets array
    // We can't really create a file with zero sheets, but we can create one
    // with an empty first sheet
    const xlsx = await writeXlsx({
      sheets: [{ name: "Empty", rows: [] }],
    })

    const objects = await readObjects(xlsx)
    expect(objects).toEqual([])
  })
})

// ── writeObjects() ──────────────────────────────────────────────────

describe("writeObjects()", () => {
  it("produces valid XLSX with correct data", async () => {
    const data = [
      { Name: "Alice", Age: 30 },
      { Name: "Bob", Age: 25 },
    ]

    const output = await writeObjects(data)

    // Read it back
    const workbook = await readXlsx(output)
    const rows = workbook.sheets[0]!.rows

    // Header row
    expect(rows[0]).toEqual(["Name", "Age"])
    // Data rows
    expect(rows[1]).toEqual(["Alice", 30])
    expect(rows[2]).toEqual(["Bob", 25])
  })

  it("uses custom sheet name", async () => {
    const output = await writeObjects([{ A: 1 }], { sheetName: "Custom" })

    const workbook = await readXlsx(output)
    expect(workbook.sheets[0]!.name).toBe("Custom")
  })

  it("writes ODS format when specified", async () => {
    const data = [{ X: "hello", Y: 42 }]
    const output = await writeObjects(data, { format: "ods" })

    const workbook = await readOds(output)
    expect(workbook.sheets[0]!.rows[0]).toEqual(["X", "Y"])
    expect(workbook.sheets[0]!.rows[1]).toEqual(["hello", 42])
  })

  it("handles empty data array", async () => {
    const output = await writeObjects([])

    const workbook = await readXlsx(output)
    expect(workbook.sheets[0]!.rows).toHaveLength(0)
  })

  it("handles null values in objects", async () => {
    const data: Array<Record<string, CellValue>> = [
      { Name: "Alice", Score: null },
      { Name: null, Score: 100 },
    ]

    const output = await writeObjects(data)
    const workbook = await readXlsx(output)
    const rows = workbook.sheets[0]!.rows

    expect(rows[0]).toEqual(["Name", "Score"])
    expect(rows[1]).toEqual(["Alice", null])
    expect(rows[2]).toEqual([null, 100])
  })

  it("handles Date values", async () => {
    const date = new Date("2025-06-15T10:30:00.000Z")
    const data = [{ Label: "Today", When: date }]

    const output = await writeObjects(data)
    const workbook = await readXlsx(output)

    expect(workbook.sheets[0]!.rows[0]).toEqual(["Label", "When"])
    // The date is stored and read back — it should be roughly equivalent
    const readDate = workbook.sheets[0]!.rows[1]![1]
    expect(readDate).toBeInstanceOf(Date)
  })
})

// ── Round-trip Tests ────────────────────────────────────────────────

describe("round-trip", () => {
  it("write() then read() preserves data (XLSX)", async () => {
    const rows: CellValue[][] = [
      ["Header1", "Header2"],
      ["value1", 42],
      ["value2", 99.5],
    ]

    const output = await write({
      sheets: [{ name: "Data", rows }],
      format: "xlsx",
    })

    const workbook = await read(output)
    expect(workbook.sheets[0]!.name).toBe("Data")
    expect(workbook.sheets[0]!.rows).toEqual(rows)
  })

  it("write() then read() preserves data (ODS)", async () => {
    const rows: CellValue[][] = [
      ["Col1", "Col2"],
      ["foo", 123],
    ]

    const output = await write({
      sheets: [{ name: "ODS", rows }],
      format: "ods",
    })

    const workbook = await read(output)
    expect(workbook.sheets[0]!.name).toBe("ODS")
    expect(workbook.sheets[0]!.rows).toEqual(rows)
  })

  it("writeObjects() then readObjects() preserves data", async () => {
    const original = [
      { Name: "Alice", Score: 95 },
      { Name: "Bob", Score: 87 },
      { Name: "Charlie", Score: 73 },
    ]

    const output = await writeObjects(original)
    const result = await readObjects(output)

    expect(result).toEqual(original)
  })
})

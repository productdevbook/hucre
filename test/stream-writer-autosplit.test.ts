import { describe, expect, it } from "vitest"
import { XlsxStreamWriter, XLSX_MAX_ROWS_PER_SHEET } from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"

describe("XlsxStreamWriter — auto-split (#170)", () => {
  it("exports the Excel row limit constant", () => {
    expect(XLSX_MAX_ROWS_PER_SHEET).toBe(1_048_576)
  })

  it("rolls over to a new sheet when maxRowsPerSheet is reached", async () => {
    const writer = new XlsxStreamWriter({
      name: "Big",
      maxRowsPerSheet: 5, // tiny limit so we don't have to write a million rows
      repeatHeaders: false,
    })

    for (let i = 0; i < 12; i++) {
      writer.addRow([i + 1, `row ${i + 1}`])
    }

    const buf = await writer.finish()
    const wb = await readXlsx(buf)

    expect(wb.sheets.map((s) => s.name)).toEqual(["Big", "Big_2", "Big_3"])
    expect(wb.sheets[0]!.rows.length).toBe(5)
    expect(wb.sheets[1]!.rows.length).toBe(5)
    expect(wb.sheets[2]!.rows.length).toBe(2)
  })

  it("repeats the header row on every rolled sheet by default", async () => {
    const writer = new XlsxStreamWriter({
      name: "Data",
      columns: [
        { key: "id", header: "ID" },
        { key: "name", header: "Name" },
      ],
      maxRowsPerSheet: 3, // 1 header + 2 data per sheet
    })

    for (let i = 0; i < 5; i++) {
      writer.addObject({ id: i + 1, name: `Item ${i + 1}` })
    }

    const buf = await writer.finish()
    const wb = await readXlsx(buf)

    expect(wb.sheets.length).toBe(3)
    for (const sheet of wb.sheets) {
      expect(sheet.rows[0]).toEqual(["ID", "Name"])
    }
    // First two sheets are full (header + 2 data); third has header + 1 data.
    expect(wb.sheets[0]!.rows.length).toBe(3)
    expect(wb.sheets[1]!.rows.length).toBe(3)
    expect(wb.sheets[2]!.rows.length).toBe(2)
  })

  it("does not repeat the header when repeatHeaders is false", async () => {
    const writer = new XlsxStreamWriter({
      name: "NoRepeat",
      columns: [{ key: "id", header: "ID" }],
      maxRowsPerSheet: 3,
      repeatHeaders: false,
    })

    // Constructor injects the "ID" header automatically. We then add 6 data rows.
    for (let i = 0; i < 6; i++) writer.addObject({ id: i + 1 })

    const buf = await writer.finish()
    const wb = await readXlsx(buf)

    // First sheet: ID, 1, 2 → 3 rows. Second sheet: 3, 4, 5 → 3 rows. Third: 6 → 1 row.
    expect(wb.sheets.length).toBe(3)
    expect(wb.sheets[0]!.rows[0]).toEqual(["ID"])
    expect(wb.sheets[1]!.rows[0]).toEqual([3])
    expect(wb.sheets[2]!.rows[0]).toEqual([6])
  })

  it("captures the first addRow call as the header when no columns are supplied", async () => {
    const writer = new XlsxStreamWriter({
      name: "Plain",
      maxRowsPerSheet: 3,
    })

    writer.addRow(["A", "B"])
    for (let i = 0; i < 5; i++) writer.addRow([i, i * 10])

    const buf = await writer.finish()
    const wb = await readXlsx(buf)

    expect(wb.sheets.length).toBe(3)
    for (const sheet of wb.sheets) {
      expect(sheet.rows[0]).toEqual(["A", "B"])
    }
  })

  it("does not roll over when row count stays within the limit", async () => {
    const writer = new XlsxStreamWriter({ name: "Small", maxRowsPerSheet: 1000 })
    writer.addRow(["a", "b"])
    writer.addRow([1, 2])

    const buf = await writer.finish()
    const wb = await readXlsx(buf)

    expect(wb.sheets.length).toBe(1)
    expect(wb.sheets[0]!.name).toBe("Small")
    expect(wb.sheets[0]!.rows.length).toBe(2)
  })

  it("truncates long base names so that suffixed names fit the 31-char limit", async () => {
    const longName = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcde" // exactly 31 chars
    const writer = new XlsxStreamWriter({
      name: longName,
      maxRowsPerSheet: 2,
      repeatHeaders: false,
    })

    for (let i = 0; i < 5; i++) writer.addRow([i])

    const buf = await writer.finish()
    const wb = await readXlsx(buf)

    expect(wb.sheets.length).toBe(3)
    expect(wb.sheets[0]!.name).toBe(longName)
    for (let s = 1; s < wb.sheets.length; s++) {
      expect(wb.sheets[s]!.name.length).toBeLessThanOrEqual(31)
      expect(wb.sheets[s]!.name.endsWith(`_${s + 1}`)).toBe(true)
    }
  })

  it("rejects maxRowsPerSheet < 2", () => {
    expect(() => new XlsxStreamWriter({ name: "X", maxRowsPerSheet: 1 })).toThrow(/at least 2/)
  })

  it("default maxRowsPerSheet is the Excel hard limit, so small workloads stay single-sheet", async () => {
    const writer = new XlsxStreamWriter({ name: "Default" })
    for (let i = 0; i < 50; i++) writer.addRow([i])
    const buf = await writer.finish()
    const wb = await readXlsx(buf)
    expect(wb.sheets.length).toBe(1)
  })
})

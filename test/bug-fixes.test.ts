/**
 * Tests for bug fixes:
 * - Bug #122: Missing workbookPr date1904 attribute
 * - Bug #116: ODS writer uses local time for dates
 * - Bug #105: Shared strings missing xml:space="preserve"
 */
import { describe, it, expect } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { parseXml } from "../src/xml/parser"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeWorkbookXml } from "../src/xlsx/workbook-writer"
import { createSharedStrings, writeSharedStringsXml } from "../src/xlsx/worksheet-writer"
import { XlsxStreamWriter } from "../src/xlsx/stream-writer"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8")

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data)
  const raw = await zip.extract(path)
  return decoder.decode(raw)
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

// ── Bug #122: Missing workbookPr date1904 attribute ─────────────────

describe("Bug #122: workbookPr date1904 attribute", () => {
  it("writeWorkbookXml includes workbookPr with date1904 when dateSystem is 1904", () => {
    const xml = writeWorkbookXml([{ name: "Sheet1", rows: [["test"]] }], undefined, "1904")
    const doc = parseXml(xml)

    const workbookPr = findChild(doc, "workbookPr")
    expect(workbookPr).toBeDefined()
    expect(workbookPr.attrs["date1904"]).toBe("1")
  })

  it("writeWorkbookXml omits workbookPr when dateSystem is 1900", () => {
    const xml = writeWorkbookXml([{ name: "Sheet1", rows: [["test"]] }], undefined, "1900")
    const doc = parseXml(xml)

    const workbookPr = findChild(doc, "workbookPr")
    expect(workbookPr).toBeUndefined()
  })

  it("writeWorkbookXml omits workbookPr when dateSystem is undefined", () => {
    const xml = writeWorkbookXml([{ name: "Sheet1", rows: [["test"]] }])
    const doc = parseXml(xml)

    const workbookPr = findChild(doc, "workbookPr")
    expect(workbookPr).toBeUndefined()
  })

  it("writeXlsx with dateSystem 1904 produces workbookPr in ZIP", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["hello"]] }],
      dateSystem: "1904",
    })

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const workbookPr = findChild(doc, "workbookPr")
    expect(workbookPr).toBeDefined()
    expect(workbookPr.attrs["date1904"]).toBe("1")
  })

  it("writeXlsx without dateSystem 1904 has no workbookPr", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["hello"]] }],
    })

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const workbookPr = findChild(doc, "workbookPr")
    expect(workbookPr).toBeUndefined()
  })

  it("round-trip: write 1904 dates and read them back correctly", async () => {
    const testDate = new Date(Date.UTC(2024, 5, 15, 10, 30, 0)) // June 15, 2024 10:30 UTC

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[testDate]],
          columns: [{ header: "Date", key: "date", numFmt: "yyyy-mm-dd" }],
        },
      ],
      dateSystem: "1904",
    })

    const wb = await readXlsx(data)
    expect(wb.dateSystem).toBe("1904")

    const cellValue = wb.sheets[0].rows[0][0]
    expect(cellValue).toBeInstanceOf(Date)
    const recovered = cellValue as Date
    expect(recovered.getUTCFullYear()).toBe(2024)
    expect(recovered.getUTCMonth()).toBe(5) // June
    expect(recovered.getUTCDate()).toBe(15)
  })

  it("stream writer includes workbookPr date1904 when dateSystem is 1904", async () => {
    const writer = new XlsxStreamWriter({
      name: "Sheet1",
      dateSystem: "1904",
    })
    writer.addRow(["test"])
    const data = await writer.finish()

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const workbookPr = findChild(doc, "workbookPr")
    expect(workbookPr).toBeDefined()
    expect(workbookPr.attrs["date1904"]).toBe("1")
  })

  it("stream writer omits workbookPr when dateSystem is 1900", async () => {
    const writer = new XlsxStreamWriter({
      name: "Sheet1",
      dateSystem: "1900",
    })
    writer.addRow(["test"])
    const data = await writer.finish()

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const workbookPr = findChild(doc, "workbookPr")
    expect(workbookPr).toBeUndefined()
  })
})

// ── Bug #116: ODS writer uses local time for dates ──────────────────

describe("Bug #116: ODS date values use UTC", () => {
  it("writes date values using UTC components", async () => {
    // Use a date where UTC and local time differ significantly
    // Jan 1, 2024 03:00:00 UTC
    const testDate = new Date(Date.UTC(2024, 0, 1, 3, 0, 0))

    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[testDate]] }],
    })

    const xml = await extractXml(data, "content.xml")

    // The date value should be in UTC: 2024-01-01T03:00:00
    expect(xml).toContain("2024-01-01T03:00:00")
  })

  it("writes midnight UTC dates correctly", async () => {
    const testDate = new Date(Date.UTC(2024, 11, 31, 0, 0, 0)) // Dec 31, 2024 midnight UTC

    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[testDate]] }],
    })

    const xml = await extractXml(data, "content.xml")
    expect(xml).toContain("2024-12-31T00:00:00")
  })

  it("round-trips date-only values through ODS write/read at XML level", async () => {
    // The ODS reader parses date strings with `new Date()` which treats
    // no-timezone strings as local time. We verify the written XML is correct.
    const testDate = new Date(Date.UTC(2024, 6, 4, 0, 0, 0)) // July 4, 2024 midnight UTC

    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[testDate]] }],
    })

    // Verify the raw XML contains correct UTC date
    const xml = await extractXml(data, "content.xml")
    expect(xml).toContain('office:date-value="2024-07-04T00:00:00"')

    // Also verify it round-trips through the reader (date at least parses)
    const wb = await readOds(data)
    const cellValue = wb.sheets[0].rows[0][0]
    expect(cellValue).toBeInstanceOf(Date)
  })

  it("writes time components using UTC", async () => {
    // Verify the raw XML contains UTC-based time
    const testDate = new Date(Date.UTC(2024, 6, 4, 15, 30, 45))

    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[testDate]] }],
    })

    const xml = await extractXml(data, "content.xml")
    expect(xml).toContain("2024-07-04T15:30:45")
  })
})

// ── Bug #105: Shared strings missing xml:space="preserve" ───────────

describe("Bug #105: shared strings xml:space preserve", () => {
  it("adds xml:space preserve for string with leading space", () => {
    const ss = createSharedStrings()
    ss.add(" hello")
    const xml = writeSharedStringsXml(ss)

    expect(xml).toContain('xml:space="preserve"')
    expect(xml).toContain(" hello")
  })

  it("adds xml:space preserve for string with trailing space", () => {
    const ss = createSharedStrings()
    ss.add("hello ")
    const xml = writeSharedStringsXml(ss)

    expect(xml).toContain('xml:space="preserve"')
    expect(xml).toContain("hello ")
  })

  it("adds xml:space preserve for string with leading and trailing spaces", () => {
    const ss = createSharedStrings()
    ss.add(" hello world ")
    const xml = writeSharedStringsXml(ss)

    expect(xml).toContain('xml:space="preserve"')
  })

  it("adds xml:space preserve for string with newline", () => {
    const ss = createSharedStrings()
    ss.add("line1\nline2")
    const xml = writeSharedStringsXml(ss)

    expect(xml).toContain('xml:space="preserve"')
  })

  it("adds xml:space preserve for string with tab", () => {
    const ss = createSharedStrings()
    ss.add("col1\tcol2")
    const xml = writeSharedStringsXml(ss)

    expect(xml).toContain('xml:space="preserve"')
  })

  it("does NOT add xml:space preserve for regular string without whitespace issues", () => {
    const ss = createSharedStrings()
    ss.add("hello world")
    const xml = writeSharedStringsXml(ss)

    expect(xml).not.toContain('xml:space="preserve"')
  })

  it("does NOT add xml:space preserve for empty string", () => {
    const ss = createSharedStrings()
    ss.add("")
    const xml = writeSharedStringsXml(ss)

    expect(xml).not.toContain('xml:space="preserve"')
  })

  it("handles mixed strings — only applies preserve where needed", () => {
    const ss = createSharedStrings()
    ss.add("normal")
    ss.add(" leading")
    ss.add("trailing ")
    ss.add("also normal")
    const xml = writeSharedStringsXml(ss)

    // Parse and verify individual <si> elements
    const doc = parseXml(xml)
    const siElements = findChildren(doc, "si")
    expect(siElements).toHaveLength(4)

    // First: "normal" — no preserve
    const t0 = findChild(siElements[0], "t")
    expect(t0.attrs["xml:space"]).toBeUndefined()

    // Second: " leading" — has preserve
    const t1 = findChild(siElements[1], "t")
    expect(t1.attrs["xml:space"]).toBe("preserve")

    // Third: "trailing " — has preserve
    const t2 = findChild(siElements[2], "t")
    expect(t2.attrs["xml:space"]).toBe("preserve")

    // Fourth: "also normal" — no preserve
    const t3 = findChild(siElements[3], "t")
    expect(t3.attrs["xml:space"]).toBeUndefined()
  })

  it("round-trips strings with leading/trailing spaces through XLSX", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[" hello ", "world", "  tabs\there  "]],
        },
      ],
    })

    const wb = await readXlsx(data)
    expect(wb.sheets[0].rows[0][0]).toBe(" hello ")
    expect(wb.sheets[0].rows[0][1]).toBe("world")
    expect(wb.sheets[0].rows[0][2]).toBe("  tabs\there  ")
  })
})

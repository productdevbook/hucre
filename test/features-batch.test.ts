import { describe, it, expect } from "vitest"
import { calculateRowHeight } from "../src/xlsx/auto-size"
import { parseCsv } from "../src/csv/reader"
import { writeCsv } from "../src/csv/writer"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { parseXml } from "../src/xml/parser"
import type { CellValue } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8")

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data)
  const raw = await zip.extract(path)
  return decoder.decode(raw)
}

async function parseXmlFromZip(data: Uint8Array, path: string) {
  const xml = await extractXml(data, path)
  return parseXml(xml)
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

// ══════════════════════════════════════════════════════════════════════
// #97: calculateRowHeight
// ══════════════════════════════════════════════════════════════════════

describe("calculateRowHeight", () => {
  it("returns default height (15) for single-line content", () => {
    const height = calculateRowHeight(["Hello", "World"])
    expect(height).toBe(15)
  })

  it("returns default height for no content", () => {
    const height = calculateRowHeight([null, null, ""])
    expect(height).toBe(15)
  })

  it("returns default height for empty array", () => {
    const height = calculateRowHeight([])
    expect(height).toBe(15)
  })

  it("increases height for multi-line content (explicit newlines)", () => {
    const height = calculateRowHeight(["Line 1\nLine 2\nLine 3"])
    // 3 lines * 15 = 45
    expect(height).toBe(45)
  })

  it("calculates wrapped text height when wrapText is true", () => {
    // "A".repeat(20) in a column of width 8.43 => ceil(20/8.43) = 3 lines
    const height = calculateRowHeight(["A".repeat(20)], {
      wrapText: true,
      columnWidths: [8.43],
    })
    // 3 lines * 15 = 45
    expect(height).toBe(45)
  })

  it("does not wrap text when wrapText is false (default)", () => {
    // Without wrapText, only explicit newlines count
    const height = calculateRowHeight(["A".repeat(100)])
    expect(height).toBe(15)
  })

  it("takes the maximum line count across all cells", () => {
    const height = calculateRowHeight(["Short", "Line1\nLine2\nLine3\nLine4", "Medium"], {
      wrapText: false,
    })
    // 4 lines from the second cell * 15 = 60
    expect(height).toBe(60)
  })

  it("scales line height with font size", () => {
    const defaultHeight = calculateRowHeight(["Line1\nLine2"])
    const largeFont = calculateRowHeight(["Line1\nLine2"], { fontSize: 22 })
    // 22pt font is 2x the default 11pt
    expect(largeFont).toBe(defaultHeight * 2)
  })

  it("uses per-column widths for wrap calculation", () => {
    const height = calculateRowHeight(["Short", "A".repeat(40)], {
      wrapText: true,
      columnWidths: [20, 10],
    })
    // Second cell: ceil(40/10) = 4 lines * 15 = 60
    expect(height).toBe(60)
  })

  it("handles mixed null and string values", () => {
    const height = calculateRowHeight([null, "A\nB", null, "C"])
    // Max is 2 lines from "A\nB"
    expect(height).toBe(30)
  })

  it("handles number values (converted to string for length)", () => {
    const height = calculateRowHeight([12345], {
      wrapText: true,
      columnWidths: [3],
    })
    // "12345" = 5 chars, ceil(5/3) = 2 lines => 30
    expect(height).toBe(30)
  })
})

// ══════════════════════════════════════════════════════════════════════
// #95: CSV multi-character delimiter
// ══════════════════════════════════════════════════════════════════════

describe("CSV multi-character delimiter", () => {
  it("parses with :: delimiter", () => {
    const input = "a::b::c\n1::2::3"
    const rows = parseCsv(input, { delimiter: "::" })
    expect(rows).toEqual([
      ["a", "b", "c"],
      ["1", "2", "3"],
    ])
  })

  it("parses with || delimiter", () => {
    const input = "name||age||city\nAlice||30||NYC"
    const rows = parseCsv(input, { delimiter: "||" })
    expect(rows).toEqual([
      ["name", "age", "city"],
      ["Alice", "30", "NYC"],
    ])
  })

  it("parses with \\t\\t delimiter", () => {
    const input = "a\t\tb\t\tc"
    const rows = parseCsv(input, { delimiter: "\t\t" })
    expect(rows).toEqual([["a", "b", "c"]])
  })

  it("writes with :: delimiter", () => {
    const rows: CellValue[][] = [
      ["a", "b", "c"],
      ["1", "2", "3"],
    ]
    const output = writeCsv(rows, { delimiter: "::" })
    expect(output).toBe("a::b::c\r\n1::2::3")
  })

  it("writes with || delimiter", () => {
    const rows: CellValue[][] = [
      ["name", "age"],
      ["Alice", "30"],
    ]
    const output = writeCsv(rows, { delimiter: "||" })
    expect(output).toBe("name||age\r\nAlice||30")
  })

  it("round-trips with :: delimiter", () => {
    const original: CellValue[][] = [
      ["id", "name", "value"],
      ["1", "Widget", "9.99"],
      ["2", "Gadget", "19.99"],
    ]
    const csv = writeCsv(original, { delimiter: "::" })
    const parsed = parseCsv(csv, { delimiter: "::" })
    expect(parsed).toEqual(original)
  })

  it("handles quoted fields containing the multi-char delimiter", () => {
    const input = 'a::"b::c"::d'
    const rows = parseCsv(input, { delimiter: "::" })
    expect(rows).toEqual([["a", "b::c", "d"]])
  })

  it("writes quoted fields that contain the multi-char delimiter", () => {
    const rows: CellValue[][] = [["a", "b::c", "d"]]
    const output = writeCsv(rows, { delimiter: "::" })
    // "b::c" must be quoted because it contains the delimiter
    expect(output).toBe('a::"b::c"::d')
  })
})

// ══════════════════════════════════════════════════════════════════════
// #81: Outline collapsed state
// ══════════════════════════════════════════════════════════════════════

describe("Outline collapsed — write", () => {
  it("writes collapsed attribute on row element", async () => {
    const rowDefs = new Map<number, { outlineLevel?: number; collapsed?: boolean }>()
    rowDefs.set(1, { outlineLevel: 1, collapsed: true })
    rowDefs.set(2, { outlineLevel: 1 })

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Header"], ["Detail 1"], ["Detail 2"]],
          rowDefs,
        },
      ],
    })

    const xml = await extractXml(data, "xl/worksheets/sheet1.xml")

    // Row 2 (1-based) should have collapsed="1"
    expect(xml).toContain('collapsed="1"')
    // Verify the row with outlineLevel and collapsed
    const root = await parseXmlFromZip(data, "xl/worksheets/sheet1.xml")
    const sheetData = findChild(root, "sheetData")
    const rows = findChildren(sheetData, "row")

    const row2 = rows.find((r: any) => r.attrs?.r === "2")
    expect(row2).toBeDefined()
    expect(row2.attrs?.outlineLevel).toBe("1")
    expect(row2.attrs?.collapsed).toBe("1")

    // Row 3 should have outlineLevel but no collapsed
    const row3 = rows.find((r: any) => r.attrs?.r === "3")
    expect(row3).toBeDefined()
    expect(row3.attrs?.outlineLevel).toBe("1")
    expect(row3.attrs?.collapsed).toBeUndefined()
  })

  it("writes collapsed attribute on col element", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          columns: [
            { width: 10 },
            { width: 10, outlineLevel: 1, collapsed: true },
            { width: 10, outlineLevel: 1 },
          ],
          rows: [["A", "B", "C"]],
        },
      ],
    })

    const root = await parseXmlFromZip(data, "xl/worksheets/sheet1.xml")
    const cols = findChild(root, "cols")
    expect(cols).toBeDefined()

    const colDefs = findChildren(cols, "col")
    // Column B (min=2) should have collapsed
    const colB = colDefs.find((c: any) => c.attrs?.min === "2")
    expect(colB).toBeDefined()
    expect(colB.attrs?.outlineLevel).toBe("1")
    expect(colB.attrs?.collapsed).toBe("true")

    // Column C (min=3) should not have collapsed
    const colC = colDefs.find((c: any) => c.attrs?.min === "3")
    expect(colC).toBeDefined()
    expect(colC.attrs?.outlineLevel).toBe("1")
    expect(colC.attrs?.collapsed).toBeUndefined()
  })

  it("writes outlinePr in sheetPr when outlineProperties are set", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A"]],
          outlineProperties: {
            summaryBelow: false,
            summaryRight: false,
          },
        },
      ],
    })

    const xml = await extractXml(data, "xl/worksheets/sheet1.xml")
    expect(xml).toContain("outlinePr")
    expect(xml).toContain('summaryBelow="0"')
    expect(xml).toContain('summaryRight="0"')
  })
})

describe("Outline collapsed — read round-trip", () => {
  it("reads collapsed attribute on rows", async () => {
    const rowDefs = new Map<
      number,
      { outlineLevel?: number; collapsed?: boolean; hidden?: boolean }
    >()
    rowDefs.set(1, { outlineLevel: 1, collapsed: true })
    rowDefs.set(2, { outlineLevel: 1, hidden: true })

    const written = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Header"], ["Detail 1"], ["Detail 2"]],
          rowDefs,
        },
      ],
    })

    const workbook = await readXlsx(written)
    const sheet = workbook.sheets[0]!

    expect(sheet.rowDefs).toBeDefined()
    const rd1 = sheet.rowDefs!.get(1)
    expect(rd1).toBeDefined()
    expect(rd1!.outlineLevel).toBe(1)
    expect(rd1!.collapsed).toBe(true)

    const rd2 = sheet.rowDefs!.get(2)
    expect(rd2).toBeDefined()
    expect(rd2!.outlineLevel).toBe(1)
    expect(rd2!.hidden).toBe(true)
    expect(rd2!.collapsed).toBeUndefined()
  })

  it("reads collapsed attribute on columns", async () => {
    const written = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          columns: [
            { width: 10 },
            { width: 12, outlineLevel: 1, collapsed: true },
            { width: 14, outlineLevel: 1 },
          ],
          rows: [["A", "B", "C"]],
        },
      ],
    })

    const workbook = await readXlsx(written)
    const sheet = workbook.sheets[0]!

    expect(sheet.columns).toBeDefined()
    expect(sheet.columns!.length).toBeGreaterThanOrEqual(3)

    // Column B (index 1) should be collapsed
    const colB = sheet.columns![1]!
    expect(colB.outlineLevel).toBe(1)
    expect(colB.collapsed).toBe(true)
    expect(colB.width).toBe(12)

    // Column C (index 2) should not be collapsed
    const colC = sheet.columns![2]!
    expect(colC.outlineLevel).toBe(1)
    expect(colC.collapsed).toBeUndefined()
    expect(colC.width).toBe(14)
  })
})

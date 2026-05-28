import { describe, it, expect } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import type { WriteSheet, Cell } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8")
const encoder = new TextEncoder()

function _zipEntries(data: Uint8Array): string[] {
  return new ZipReader(data).entries()
}

function zipHas(data: Uint8Array, path: string): boolean {
  return new ZipReader(data).has(path)
}

async function zipExtract(data: Uint8Array, path: string): Promise<Uint8Array> {
  return new ZipReader(data).extract(path)
}

async function zipExtractText(data: Uint8Array, path: string): Promise<string> {
  const raw = await zipExtract(data, path)
  return decoder.decode(raw)
}

/** Create a minimal valid XLSX from writeXlsx for testing */
async function createBasicXlsx(sheets: WriteSheet[]): Promise<Uint8Array> {
  return writeXlsx({ sheets })
}

/** Create a fake PNG image */
function fakePng(size = 64): Uint8Array {
  const data = new Uint8Array(size)
  data[0] = 0x89
  data[1] = 0x50
  data[2] = 0x4e
  data[3] = 0x47
  data[4] = 0x0d
  data[5] = 0x0a
  data[6] = 0x1a
  data[7] = 0x0a
  for (let i = 8; i < size; i++) {
    data[i] = i % 256
  }
  return data
}

/**
 * Inject extra ZIP entries into an existing XLSX archive.
 * Used to simulate unknown parts (charts, VBA, etc.) for round-trip testing.
 */
async function injectEntries(
  original: Uint8Array,
  extras: Array<{ path: string; data: Uint8Array }>,
): Promise<Uint8Array> {
  const zip = new ZipReader(original)
  const writer = new ZipWriter()

  // Copy all original entries
  for (const path of zip.entries()) {
    const data = await zip.extract(path)
    writer.add(path, data, { compress: false })
  }

  // Add extra entries
  for (const entry of extras) {
    writer.add(entry.path, entry.data, { compress: false })
  }

  return writer.build()
}

// ── Tests ────────────────────────────────────────────────────────────

describe("roundtrip — openXlsx", () => {
  it("returns a workbook with parsed data", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [
          ["Hello", 42],
          ["World", 99],
        ],
      },
    ])

    const wb = await openXlsx(xlsx)
    expect(wb.sheets).toHaveLength(1)
    expect(wb.sheets[0].name).toBe("Sheet1")
    expect(wb.sheets[0].rows[0][0]).toBe("Hello")
    expect(wb.sheets[0].rows[0][1]).toBe(42)
    expect(wb.sheets[0].rows[1][0]).toBe("World")
    expect(wb.sheets[0].rows[1][1]).toBe(99)
  })

  it("stores raw ZIP entries", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["A"]] }])

    const wb = await openXlsx(xlsx)
    expect(wb._rawEntries).toBeInstanceOf(Map)
    expect(wb._rawEntries.size).toBeGreaterThan(0)
    expect(wb._rawEntries.has("xl/workbook.xml")).toBe(true)
    expect(wb._rawEntries.has("xl/styles.xml")).toBe(true)
  })

  it("stores content types and root rels", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["A"]] }])

    const wb = await openXlsx(xlsx)
    // The content types XML contains the namespace with "content-types"
    expect(wb._contentTypes).toContain("content-types")
    expect(wb._rootRels).toContain("Relationships")
  })

  it("initializes _modifiedParts as empty set", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["A"]] }])

    const wb = await openXlsx(xlsx)
    expect(wb._modifiedParts).toBeInstanceOf(Set)
    expect(wb._modifiedParts.size).toBe(0)
  })

  it("accepts ArrayBuffer input", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["Test"]] }])
    const arrayBuffer = xlsx.buffer.slice(
      xlsx.byteOffset,
      xlsx.byteOffset + xlsx.byteLength,
    ) as ArrayBuffer

    const wb = await openXlsx(arrayBuffer)
    expect(wb.sheets[0].rows[0][0]).toBe("Test")
    expect(wb._rawEntries.size).toBeGreaterThan(0)
  })
})

describe("roundtrip — save without changes", () => {
  it("produces a valid XLSX when saved without modifications", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [
          ["Hello", 42],
          ["World", 99],
        ],
      },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)

    // Read back and verify data is preserved
    const wb2 = await readXlsx(saved)
    expect(wb2.sheets).toHaveLength(1)
    expect(wb2.sheets[0].name).toBe("Sheet1")
    expect(wb2.sheets[0].rows[0][0]).toBe("Hello")
    expect(wb2.sheets[0].rows[0][1]).toBe(42)
    expect(wb2.sheets[0].rows[1][0]).toBe("World")
    expect(wb2.sheets[0].rows[1][1]).toBe(99)
  })

  it("preserves multiple sheets", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Data",
        rows: [
          ["A", 1],
          ["B", 2],
        ],
      },
      { name: "Summary", rows: [["Total", 3]] },
      { name: "Notes", rows: [["Info"]] },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)

    const wb2 = await readXlsx(saved)
    expect(wb2.sheets).toHaveLength(3)
    expect(wb2.sheets[0].name).toBe("Data")
    expect(wb2.sheets[1].name).toBe("Summary")
    expect(wb2.sheets[2].name).toBe("Notes")
    expect(wb2.sheets[0].rows[1][1]).toBe(2)
    expect(wb2.sheets[1].rows[0][1]).toBe(3)
  })
})

describe("roundtrip — modify cells", () => {
  it("modifies a cell value and preserves it through round-trip", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [
          ["Hello", 42],
          ["World", 99],
        ],
      },
    ])

    const wb = await openXlsx(xlsx)
    // Modify cell value
    wb.sheets[0].rows[0][1] = 100

    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].rows[0][0]).toBe("Hello")
    expect(wb2.sheets[0].rows[0][1]).toBe(100)
    expect(wb2.sheets[0].rows[1][0]).toBe("World")
    expect(wb2.sheets[0].rows[1][1]).toBe(99)
  })

  it("adds a new row", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [
          ["A", 1],
          ["B", 2],
        ],
      },
    ])

    const wb = await openXlsx(xlsx)
    wb.sheets[0].rows.push(["C", 3])

    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].rows).toHaveLength(3)
    expect(wb2.sheets[0].rows[2][0]).toBe("C")
    expect(wb2.sheets[0].rows[2][1]).toBe(3)
  })

  it("changes a string to a number", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["status", "pending"]] }])

    const wb = await openXlsx(xlsx)
    wb.sheets[0].rows[0][1] = 42

    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)
    expect(wb2.sheets[0].rows[0][1]).toBe(42)
  })

  it("changes a cell to null", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["Hello", 42]] }])

    const wb = await openXlsx(xlsx)
    wb.sheets[0].rows[0][1] = null

    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)
    expect(wb2.sheets[0].rows[0][0]).toBe("Hello")
    // Reader returns undefined for empty cells; the value is effectively absent
    expect(wb2.sheets[0].rows[0][1] ?? null).toBe(null)
  })
})

describe("roundtrip — change sheet name", () => {
  it("renames a sheet and preserves it through round-trip", async () => {
    const xlsx = await createBasicXlsx([{ name: "OldName", rows: [["Data", 1]] }])

    const wb = await openXlsx(xlsx)
    wb.sheets[0].name = "NewName"

    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)
    expect(wb2.sheets[0].name).toBe("NewName")
    expect(wb2.sheets[0].rows[0][0]).toBe("Data")
  })
})

describe("roundtrip — preserve unknown ZIP entries", () => {
  it("preserves a fake chart entry through round-trip", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["Value", 42]] }])

    const chartXml = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?><c:chart xmlns:c="fake"><c:title>Test</c:title></c:chart>`,
    )

    const xlsxWithChart = await injectEntries(xlsx, [
      { path: "xl/charts/chart1.xml", data: chartXml },
    ])

    // Verify the injected entry exists
    expect(zipHas(xlsxWithChart, "xl/charts/chart1.xml")).toBe(true)

    const wb = await openXlsx(xlsxWithChart)
    const saved = await saveXlsx(wb)

    // Verify the chart entry survived the round-trip
    expect(zipHas(saved, "xl/charts/chart1.xml")).toBe(true)
    const preservedChart = await zipExtractText(saved, "xl/charts/chart1.xml")
    expect(preservedChart).toContain("<c:title>Test</c:title>")
  })

  it("preserves a fake VBA project binary through round-trip", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["Macro test"]] }])

    const vbaData = new Uint8Array([0xde, 0xad, 0xbe, 0xef, 0x01, 0x02, 0x03, 0x04])

    const xlsxWithVba = await injectEntries(xlsx, [{ path: "xl/vbaProject.bin", data: vbaData }])

    const wb = await openXlsx(xlsxWithVba)
    const saved = await saveXlsx(wb)

    expect(zipHas(saved, "xl/vbaProject.bin")).toBe(true)
    const preserved = await zipExtract(saved, "xl/vbaProject.bin")
    expect(preserved).toEqual(vbaData)
  })

  it("preserves theme XML through round-trip", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["Theme test"]] }])

    const themeXml = encoder.encode("<a:theme>Custom Theme</a:theme>")
    const xlsxWithTheme = await injectEntries(xlsx, [
      { path: "xl/theme/theme1.xml", data: themeXml },
    ])

    const wb = await openXlsx(xlsxWithTheme)
    const saved = await saveXlsx(wb)

    expect(zipHas(saved, "xl/theme/theme1.xml")).toBe(true)
    const preserved = await zipExtractText(saved, "xl/theme/theme1.xml")
    expect(preserved).toContain("Custom Theme")
  })

  it("preserves printer settings through round-trip", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["Print test"]] }])

    const printerData = new Uint8Array([0x01, 0x02, 0x03])
    const xlsxWithPrinter = await injectEntries(xlsx, [
      { path: "xl/printerSettings/printerSettings1.bin", data: printerData },
    ])

    const wb = await openXlsx(xlsxWithPrinter)
    const saved = await saveXlsx(wb)

    expect(zipHas(saved, "xl/printerSettings/printerSettings1.bin")).toBe(true)
  })

  it("preserves customXml through round-trip", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["Custom"]] }])

    const customXml = encoder.encode("<custom>data</custom>")
    const xlsxWithCustom = await injectEntries(xlsx, [
      { path: "customXml/item1.xml", data: customXml },
    ])

    const wb = await openXlsx(xlsxWithCustom)
    const saved = await saveXlsx(wb)

    expect(zipHas(saved, "customXml/item1.xml")).toBe(true)
    const preserved = await zipExtractText(saved, "customXml/item1.xml")
    expect(preserved).toContain("<custom>data</custom>")
  })

  it("preserves multiple unknown entries simultaneously", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["Multi"]] }])

    const xlsxWithExtras = await injectEntries(xlsx, [
      { path: "xl/charts/chart1.xml", data: encoder.encode("<chart>1</chart>") },
      { path: "xl/charts/chart2.xml", data: encoder.encode("<chart>2</chart>") },
      { path: "xl/vbaProject.bin", data: new Uint8Array([0xff]) },
      { path: "xl/theme/theme1.xml", data: encoder.encode("<theme/>") },
      { path: "customXml/item1.xml", data: encoder.encode("<custom/>") },
    ])

    const wb = await openXlsx(xlsxWithExtras)
    // Modify data to ensure regeneration happens
    wb.sheets[0].rows[0][0] = "Modified"
    const saved = await saveXlsx(wb)

    expect(zipHas(saved, "xl/charts/chart1.xml")).toBe(true)
    expect(zipHas(saved, "xl/charts/chart2.xml")).toBe(true)
    expect(zipHas(saved, "xl/vbaProject.bin")).toBe(true)
    expect(zipHas(saved, "xl/theme/theme1.xml")).toBe(true)
    expect(zipHas(saved, "customXml/item1.xml")).toBe(true)

    // Also verify the modification took effect
    const wb2 = await readXlsx(saved)
    expect(wb2.sheets[0].rows[0][0]).toBe("Modified")
  })
})

describe("roundtrip — preserve data validations", () => {
  it("preserves list validation through round-trip", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [["Status"]],
        dataValidations: [
          {
            type: "list",
            values: ["Active", "Inactive", "Draft"],
            range: "A2:A100",
            allowBlank: true,
            showInputMessage: true,
            showErrorMessage: true,
          },
        ],
      },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].dataValidations).toBeDefined()
    expect(wb2.sheets[0].dataValidations!.length).toBe(1)
    const dv = wb2.sheets[0].dataValidations![0]
    expect(dv.type).toBe("list")
    expect(dv.values).toEqual(["Active", "Inactive", "Draft"])
  })

  it("preserves whole number validation through round-trip", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [["Age"]],
        dataValidations: [
          {
            type: "whole",
            operator: "between",
            formula1: "1",
            formula2: "120",
            range: "A2:A100",
            showErrorMessage: true,
            errorTitle: "Invalid Age",
            errorMessage: "Must be between 1 and 120",
          },
        ],
      },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].dataValidations).toBeDefined()
    const dv = wb2.sheets[0].dataValidations![0]
    expect(dv.type).toBe("whole")
    expect(dv.operator).toBe("between")
    expect(dv.formula1).toBe("1")
    expect(dv.formula2).toBe("120")
  })
})

describe("roundtrip — preserve hyperlinks", () => {
  it("preserves external hyperlink through round-trip", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Click me",
      type: "string",
      hyperlink: { target: "https://example.com" },
    })

    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [["Click me"]],
        cells,
      },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].cells).toBeDefined()
    const cell = wb2.sheets[0].cells!.get("0,0")
    expect(cell).toBeDefined()
    expect(cell!.hyperlink).toBeDefined()
    expect(cell!.hyperlink!.target).toBe("https://example.com")
  })

  it("preserves internal hyperlink through round-trip", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Go to B1",
      type: "string",
      hyperlink: { target: "", location: "Sheet1!B1" },
    })

    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [["Go to B1"]],
        cells,
      },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    const cell = wb2.sheets[0].cells?.get("0,0")
    expect(cell).toBeDefined()
    expect(cell!.hyperlink).toBeDefined()
    expect(cell!.hyperlink!.location).toBe("Sheet1!B1")
  })
})

describe("roundtrip — preserve merges", () => {
  it("preserves merge ranges through round-trip", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [
          ["Merged Header", null, null],
          ["A", "B", "C"],
        ],
        merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
      },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].merges).toBeDefined()
    expect(wb2.sheets[0].merges!.length).toBe(1)
    const merge = wb2.sheets[0].merges![0]
    expect(merge.startRow).toBe(0)
    expect(merge.startCol).toBe(0)
    expect(merge.endRow).toBe(0)
    expect(merge.endCol).toBe(2)
  })

  it("preserves multiple merges through round-trip", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [
          ["Title", null, null],
          ["Sub1", null, "Sub2"],
          [1, 2, 3],
        ],
        merges: [
          { startRow: 0, startCol: 0, endRow: 0, endCol: 2 },
          { startRow: 1, startCol: 0, endRow: 1, endCol: 1 },
        ],
      },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].merges!.length).toBe(2)
  })
})

describe("roundtrip — preserve images", () => {
  it("preserves images through round-trip", async () => {
    const pngData = fakePng(128)

    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [["Image below"]],
        images: [
          {
            data: pngData,
            type: "png",
            anchor: {
              from: { row: 1, col: 0 },
              to: { row: 5, col: 3 },
            },
          },
        ],
      },
    ])

    const wb = await openXlsx(xlsx)
    expect(wb.sheets[0].images).toBeDefined()
    expect(wb.sheets[0].images!.length).toBe(1)

    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].images).toBeDefined()
    expect(wb2.sheets[0].images!.length).toBe(1)
    expect(wb2.sheets[0].images![0].type).toBe("png")
    expect(wb2.sheets[0].images![0].anchor.from.row).toBe(1)
    expect(wb2.sheets[0].images![0].anchor.from.col).toBe(0)
  })
})

describe("roundtrip — preserve comments", () => {
  it("preserves comments through round-trip", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Hello",
      type: "string",
      comment: { text: "This is a comment", author: "Tester" },
    })

    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [["Hello"]],
        cells,
      },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].cells).toBeDefined()
    const cell = wb2.sheets[0].cells!.get("0,0")
    expect(cell).toBeDefined()
    expect(cell!.comment).toBeDefined()
    expect(cell!.comment!.text).toBe("This is a comment")
    expect(cell!.comment!.author).toBe("Tester")
  })

  it("preserves multiple comments on different cells", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "A1",
      type: "string",
      comment: { text: "Comment on A1" },
    })
    cells.set("1,1", {
      value: "B2",
      type: "string",
      comment: { text: "Comment on B2", author: "Alice" },
    })

    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [
          ["A1", null],
          [null, "B2"],
        ],
        cells,
      },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    const cellA1 = wb2.sheets[0].cells!.get("0,0")
    const cellB2 = wb2.sheets[0].cells!.get("1,1")
    expect(cellA1?.comment?.text).toBe("Comment on A1")
    expect(cellB2?.comment?.text).toBe("Comment on B2")
    expect(cellB2?.comment?.author).toBe("Alice")
  })
})

describe("roundtrip — workbook structure", () => {
  it("preserves hidden sheets", async () => {
    const xlsx = await createBasicXlsx([
      { name: "Visible", rows: [["Show"]] },
      { name: "Hidden", rows: [["Hide"]], hidden: true },
    ])

    const wb = await openXlsx(xlsx)
    expect(wb.sheets[1].hidden).toBe(true)

    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)
    expect(wb2.sheets[1].hidden).toBe(true)
    expect(wb2.sheets[1].rows[0][0]).toBe("Hide")
  })

  it("preserves named ranges", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Data",
          rows: [
            ["A", 1],
            ["B", 2],
          ],
        },
      ],
      namedRanges: [{ name: "MyRange", range: "Data!$A$1:$B$2" }],
    })

    const wb = await openXlsx(xlsx)
    expect(wb.namedRanges).toBeDefined()
    expect(wb.namedRanges!.length).toBe(1)
    expect(wb.namedRanges![0].name).toBe("MyRange")

    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)
    expect(wb2.namedRanges).toBeDefined()
    expect(wb2.namedRanges!.length).toBe(1)
    expect(wb2.namedRanges![0].name).toBe("MyRange")
    expect(wb2.namedRanges![0].range).toBe("Data!$A$1:$B$2")
  })

  it("preserves freeze panes when set on workbook", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [["Header"], ["Row1"], ["Row2"]],
      },
    ])

    const wb = await openXlsx(xlsx)
    // Manually set freezePane on the workbook (simulating a modification)
    wb.sheets[0].freezePane = { rows: 1 }

    const saved = await saveXlsx(wb)
    // Verify the freeze pane XML is present in the output
    const sheetXml = await zipExtractText(saved, "xl/worksheets/sheet1.xml")
    expect(sheetXml).toContain("pane")
    expect(sheetXml).toContain("frozen")
  })

  it("preserves auto filter when set on workbook", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [
          ["Name", "Age"],
          ["Alice", 30],
          ["Bob", 25],
        ],
      },
    ])

    const wb = await openXlsx(xlsx)
    // Manually set autoFilter (simulating a modification)
    wb.sheets[0].autoFilter = { range: "A1:B3" }

    const saved = await saveXlsx(wb)
    // Verify the autoFilter XML is present in the output
    const sheetXml = await zipExtractText(saved, "xl/worksheets/sheet1.xml")
    expect(sheetXml).toContain("autoFilter")
    expect(sheetXml).toContain("A1:B3")
  })

  it("preserves document properties", async () => {
    const xlsx = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Data"]] }],
      properties: {
        title: "Test Document",
        creator: "Test Author",
        subject: "Testing",
      },
    })

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.properties).toBeDefined()
    expect(wb2.properties!.title).toBe("Test Document")
    expect(wb2.properties!.creator).toBe("Test Author")
  })
})

describe("roundtrip — complex scenarios", () => {
  it("modifies data while preserving unknown parts", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [
          ["Original", 1],
          ["Data", 2],
        ],
      },
    ])

    const xlsxWithExtras = await injectEntries(xlsx, [
      { path: "xl/charts/chart1.xml", data: encoder.encode("<chart>1</chart>") },
      { path: "xl/theme/theme1.xml", data: encoder.encode("<theme/>") },
    ])

    const wb = await openXlsx(xlsxWithExtras)

    // Modify data
    wb.sheets[0].rows[0][0] = "Modified"
    wb.sheets[0].rows.push(["New", 3])

    const saved = await saveXlsx(wb)

    // Verify modifications
    const wb2 = await readXlsx(saved)
    expect(wb2.sheets[0].rows[0][0]).toBe("Modified")
    expect(wb2.sheets[0].rows[2][0]).toBe("New")
    expect(wb2.sheets[0].rows[2][1]).toBe(3)

    // Verify preserved parts
    expect(zipHas(saved, "xl/charts/chart1.xml")).toBe(true)
    expect(zipHas(saved, "xl/theme/theme1.xml")).toBe(true)
  })

  it("adds a sheet while preserving unknown parts", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [["Original"]] }])

    const xlsxWithExtras = await injectEntries(xlsx, [
      { path: "xl/vbaProject.bin", data: new Uint8Array([0x01, 0x02]) },
    ])

    const wb = await openXlsx(xlsxWithExtras)

    // Add a new sheet by modifying the sheets array
    wb.sheets.push({
      name: "NewSheet",
      rows: [["Added"]],
    })

    const saved = await saveXlsx(wb)

    const wb2 = await readXlsx(saved)
    expect(wb2.sheets).toHaveLength(2)
    expect(wb2.sheets[1].name).toBe("NewSheet")
    expect(wb2.sheets[1].rows[0][0]).toBe("Added")

    // VBA still preserved
    expect(zipHas(saved, "xl/vbaProject.bin")).toBe(true)
  })

  it("handles empty workbook roundtrip", async () => {
    const xlsx = await createBasicXlsx([{ name: "Empty", rows: [] }])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)
    expect(wb2.sheets[0].name).toBe("Empty")
  })

  it("preserves conditional formatting through round-trip", async () => {
    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [[1], [2], [3], [4], [5]],
        conditionalRules: [
          {
            type: "cellIs",
            priority: 1,
            operator: "greaterThan",
            formula: "3",
            range: "A1:A5",
            style: {
              font: { bold: true },
            },
          },
        ],
      },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].conditionalRules).toBeDefined()
    expect(wb2.sheets[0].conditionalRules!.length).toBe(1)
    expect(wb2.sheets[0].conditionalRules![0].type).toBe("cellIs")
    expect(wb2.sheets[0].conditionalRules![0].operator).toBe("greaterThan")
  })

  it("roundtrips with styles", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Bold",
      type: "string",
      style: {
        font: { bold: true },
      },
    })

    const xlsx = await createBasicXlsx([
      {
        name: "Sheet1",
        rows: [["Bold"]],
        cells,
      },
    ])

    const wb = await openXlsx(xlsx, { readStyles: true })
    const saved = await saveXlsx(wb)

    // Verify the file is valid
    const wb2 = await readXlsx(saved, { readStyles: true })
    expect(wb2.sheets[0].rows[0][0]).toBe("Bold")
    // The style should be preserved as long as the cell has style info
    const cell = wb2.sheets[0].cells?.get("0,0")
    expect(cell).toBeDefined()
    if (cell?.style?.font) {
      expect(cell.style.font.bold).toBe(true)
    }
  })
})

describe("roundtrip — edge cases", () => {
  it("handles special characters in cell values", async () => {
    const xlsx = await createBasicXlsx([
      { name: "Sheet1", rows: [["<script>alert('xss')</script>", "A&B", '"quotes"']] },
    ])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].rows[0][0]).toBe("<script>alert('xss')</script>")
    expect(wb2.sheets[0].rows[0][1]).toBe("A&B")
    expect(wb2.sheets[0].rows[0][2]).toBe('"quotes"')
  })

  it("handles special characters in sheet names", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet & <Data>", rows: [["Test"]] }])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].name).toBe("Sheet & <Data>")
  })

  it("handles boolean values", async () => {
    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows: [[true, false]] }])

    const wb = await openXlsx(xlsx)
    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].rows[0][0]).toBe(true)
    expect(wb2.sheets[0].rows[0][1]).toBe(false)
  })

  it("handles large numbers of rows", async () => {
    const rows: Array<Array<string | number>> = []
    for (let i = 0; i < 500; i++) {
      rows.push([`Row ${i}`, i, i * 2])
    }

    const xlsx = await createBasicXlsx([{ name: "Sheet1", rows }])
    const wb = await openXlsx(xlsx)

    // Modify a row in the middle
    wb.sheets[0].rows[250][0] = "Modified"

    const saved = await saveXlsx(wb)
    const wb2 = await readXlsx(saved)

    expect(wb2.sheets[0].rows).toHaveLength(500)
    expect(wb2.sheets[0].rows[250][0]).toBe("Modified")
    expect(wb2.sheets[0].rows[0][0]).toBe("Row 0")
    expect(wb2.sheets[0].rows[499][0]).toBe("Row 499")
  })
})

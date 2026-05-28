import { describe, it, expect } from "vitest"
import { parseXml, parseSax } from "../src/xml/parser"
import { createStylesCollector } from "../src/xlsx/styles-writer"
import { createSharedStrings, writeWorksheetXml } from "../src/xlsx/worksheet-writer"
import { writeWorkbookXml } from "../src/xlsx/workbook-writer"
import { writeAppProperties } from "../src/xlsx/doc-props-writer"
import { parseStyles } from "../src/xlsx/styles"
import type { WriteSheet, WriteOptions } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

function writeXml(sheet: WriteSheet): string {
  const styles = createStylesCollector()
  const ss = createSharedStrings()
  const result = writeWorksheetXml(sheet, styles, ss)
  return result.xml
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function getChildTags(el: { children: Array<unknown> }): string[] {
  return el.children.filter((c: any) => typeof c !== "string").map((c: any) => c.local || c.tag)
}

// ── #125: dxfs element when count is 0 ────────────────────────────────

describe("#125: dxfs element when count is 0", () => {
  it("emits <dxfs count='0'/> self-closing when no dxf entries", () => {
    const styles = createStylesCollector()
    const xml = styles.toXml()
    const doc = parseXml(xml)

    const dxfs = findChild(doc, "dxfs")
    expect(dxfs).toBeDefined()
    expect(dxfs.attrs["count"]).toBe("0")
  })

  it("emits <dxfs count='N'> with children when dxf entries exist", () => {
    const styles = createStylesCollector()
    styles.addDxf({ font: { bold: true } })
    const xml = styles.toXml()
    const doc = parseXml(xml)

    const dxfs = findChild(doc, "dxfs")
    expect(dxfs).toBeDefined()
    expect(dxfs.attrs["count"]).toBe("1")
    const dxfChildren = findChildren(dxfs, "dxf")
    expect(dxfChildren.length).toBe(1)
  })
})

// ── #90: Active sheet index write ───────────────────────────────────

describe("#90: Active sheet index write", () => {
  it("writes activeTab from activeSheet parameter", () => {
    const sheets: WriteSheet[] = [
      { name: "Sheet1", rows: [["a"]] },
      { name: "Sheet2", rows: [["b"]] },
    ]

    const xml = writeWorkbookXml(sheets, undefined, undefined, 1)
    const doc = parseXml(xml)

    const bookViews = findChild(doc, "bookViews")
    expect(bookViews).toBeDefined()
    const wbView = findChild(bookViews, "workbookView")
    expect(wbView.attrs["activeTab"]).toBe("1")
  })

  it("defaults activeTab to 0 when not specified", () => {
    const sheets: WriteSheet[] = [{ name: "Sheet1", rows: [["a"]] }]

    const xml = writeWorkbookXml(sheets)
    const doc = parseXml(xml)

    const bookViews = findChild(doc, "bookViews")
    const wbView = findChild(bookViews, "workbookView")
    expect(wbView.attrs["activeTab"]).toBe("0")
  })
})

// ── #87: Tab color support ──────────────────────────────────────────

describe("#87: Tab color support", () => {
  it("writes tabColor inside sheetPr", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { tabColor: { rgb: "FF0000" } },
    }

    const xml = writeXml(sheet)
    const doc = parseXml(xml)

    const sheetPr = findChild(doc, "sheetPr")
    expect(sheetPr).toBeDefined()
    const tabColor = findChild(sheetPr, "tabColor")
    expect(tabColor).toBeDefined()
    // Should be ARGB format with FF prefix
    expect(tabColor.attrs["rgb"]).toBe("FFFF0000")
  })
})

// ── #89: Rich text writing ──────────────────────────────────────────

describe("#89: Rich text writing", () => {
  it("writes inline rich text with font properties", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [],
      cells: new Map([
        [
          "0,0",
          {
            richText: [
              { text: "Hello ", font: { bold: true } },
              { text: "World", font: { italic: true, color: { rgb: "FF0000" } } },
            ],
          },
        ],
      ]),
    }

    const xml = writeXml(sheet)
    const doc = parseXml(xml)

    const sheetData = findChild(doc, "sheetData")
    const row = findChild(sheetData, "row")
    const cell = findChild(row, "c")
    expect(cell.attrs["t"]).toBe("inlineStr")

    const is = findChild(cell, "is")
    expect(is).toBeDefined()
    const runs = findChildren(is, "r")
    expect(runs.length).toBe(2)

    // First run: bold
    const rPr1 = findChild(runs[0], "rPr")
    expect(findChild(rPr1, "b")).toBeDefined()

    // Second run: italic with color
    const rPr2 = findChild(runs[1], "rPr")
    expect(findChild(rPr2, "i")).toBeDefined()
  })
})

// ── #128: Worksheet element ordering ────────────────────────────────

describe("#128: Worksheet element ordering per OOXML spec", () => {
  it("follows correct order: sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData, ...", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [
        ["A", "B"],
        [1, 2],
      ],
      columns: [{ width: 10 }, { width: 20 }],
      view: { tabColor: { rgb: "FF0000" } },
      autoFilter: { range: "A1:B2" },
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
      pageSetup: { orientation: "landscape" },
    }

    const xml = writeXml(sheet)
    const doc = parseXml(xml)

    const tags = getChildTags(doc)

    // Verify ordering of elements
    const expectedOrder = [
      "sheetPr",
      "dimension",
      "sheetViews",
      "sheetFormatPr",
      "cols",
      "sheetData",
      "autoFilter",
      "mergeCells",
      "printOptions",
      "pageMargins",
      "pageSetup",
    ]

    // Filter to only expected tags
    const filteredTags = tags.filter((t) => expectedOrder.includes(t))
    expect(filteredTags).toEqual(expectedOrder)
  })

  it("dimension comes before sheetViews", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
    }

    const xml = writeXml(sheet)
    const doc = parseXml(xml)
    const tags = getChildTags(doc)

    const dimIdx = tags.indexOf("dimension")
    const svIdx = tags.indexOf("sheetViews")
    expect(dimIdx).toBeGreaterThan(-1)
    expect(svIdx).toBeGreaterThan(-1)
    expect(dimIdx).toBeLessThan(svIdx)
  })
})

// ── #134: DocSecurity in app.xml ────────────────────────────────────

describe("#134: DocSecurity in app.xml", () => {
  it("includes <DocSecurity>0</DocSecurity>", () => {
    const xml = writeAppProperties()
    const doc = parseXml(xml)

    const docSecurity = findChild(doc, "DocSecurity")
    expect(docSecurity).toBeDefined()
    const text = docSecurity.children.filter((c: unknown) => typeof c === "string").join("")
    expect(text).toBe("0")
  })
})

// ── #133: XML standalone attribute ──────────────────────────────────

describe("#133: XML standalone attribute", () => {
  it("uses standalone='yes' by default", () => {
    const styles = createStylesCollector()
    const xml = styles.toXml()
    expect(xml).toContain('standalone="yes"')
  })
})

// ── #121: XML parser BOM handling ───────────────────────────────────

describe("#121: XML parser BOM handling", () => {
  it("strips UTF-8 BOM from the start of XML input", () => {
    const bom = "\uFEFF"
    const xml = `${bom}<?xml version="1.0" encoding="UTF-8"?><root><child name="test"/></root>`

    const doc = parseXml(xml)
    expect(doc.tag).toBe("root")
    const child = findChild(doc, "child")
    expect(child).toBeDefined()
    expect(child.attrs["name"]).toBe("test")
  })

  it("works fine without BOM", () => {
    const xml = '<?xml version="1.0" encoding="UTF-8"?><root><child name="test"/></root>'
    const doc = parseXml(xml)
    expect(doc.tag).toBe("root")
  })

  it("handles BOM in SAX mode", () => {
    const bom = "\uFEFF"
    const xml = `${bom}<root><item val="1"/></root>`

    const tags: string[] = []
    parseSax(xml, {
      onOpenTag(tag) {
        tags.push(tag)
      },
    })

    expect(tags).toContain("root")
    expect(tags).toContain("item")
  })
})

// ── #94: Workbook calcPr ────────────────────────────────────────────

describe("#94: Workbook calcPr", () => {
  it("includes <calcPr> element in workbook.xml", () => {
    const sheets: WriteSheet[] = [{ name: "Sheet1", rows: [["a"]] }]
    const xml = writeWorkbookXml(sheets)
    const doc = parseXml(xml)

    const calcPr = findChild(doc, "calcPr")
    expect(calcPr).toBeDefined()
    expect(calcPr.attrs["calcId"]).toBe("0")
    expect(calcPr.attrs["fullCalcOnLoad"]).toBe("1")
  })

  it("calcPr comes after sheets and definedNames", () => {
    const sheets: WriteSheet[] = [{ name: "Sheet1", rows: [["a"]] }]
    const xml = writeWorkbookXml(sheets, [{ name: "MyRange", range: "Sheet1!$A$1:$A$10" }])
    const doc = parseXml(xml)

    const tags = getChildTags(doc)
    const sheetsIdx = tags.indexOf("sheets")
    const dnIdx = tags.indexOf("definedNames")
    const calcPrIdx = tags.indexOf("calcPr")

    expect(sheetsIdx).toBeGreaterThan(-1)
    expect(dnIdx).toBeGreaterThan(-1)
    expect(calcPrIdx).toBeGreaterThan(-1)
    expect(calcPrIdx).toBeGreaterThan(sheetsIdx)
    expect(calcPrIdx).toBeGreaterThan(dnIdx)
  })
})

// ── #118: Color ARGB alpha handling ─────────────────────────────────

describe("#118: Color ARGB alpha handling", () => {
  it("prepends FF alpha when writing 6-char RGB colors", () => {
    const styles = createStylesCollector()
    styles.addStyle({
      font: { color: { rgb: "FF0000" } },
    })
    const xml = styles.toXml()
    // The color element should have rgb="FFFF0000" (FF prefix + FF0000)
    expect(xml).toContain('rgb="FFFF0000"')
  })

  it("keeps 8-char ARGB colors as-is when writing", () => {
    const styles = createStylesCollector()
    styles.addStyle({
      font: { color: { rgb: "80FF0000" } },
    })
    const xml = styles.toXml()
    expect(xml).toContain('rgb="80FF0000"')
  })

  it("strips alpha prefix when reading 8-char ARGB colors", () => {
    // Parse a styles.xml with an 8-char ARGB color
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <fonts count="1">
          <font><color rgb="FFFF0000"/><sz val="11"/><name val="Calibri"/></font>
        </fonts>
        <fills count="2">
          <fill><patternFill patternType="none"/></fill>
          <fill><patternFill patternType="gray125"/></fill>
        </fills>
        <borders count="1"><border/></borders>
        <cellXfs count="1">
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        </cellXfs>
      </styleSheet>`

    const parsed = parseStyles(stylesXml)
    expect(parsed.fonts[0].color?.rgb).toBe("FF0000")
  })
})

// ── #132: ODS mimetype validation ───────────────────────────────────

describe("#132: ODS mimetype validation", () => {
  it("rejects ZIP files without mimetype entry", async () => {
    const { readOds } = await import("../src/ods/reader")
    const { ZipWriter } = await import("../src/zip/writer")

    const zip = new ZipWriter()
    zip.add("content.xml", new TextEncoder().encode("<document/>"))
    const badZip = await zip.build()

    await expect(readOds(badZip)).rejects.toThrow("missing 'mimetype' entry")
  })
})

// ── #99: Range-scoped reading ───────────────────────────────────────

describe("#99: Range-scoped reading", () => {
  it("reads only cells within the specified range", async () => {
    const { writeXlsx, readXlsx } = await import("../src/index")

    // Create a workbook with known data
    const options: WriteOptions = {
      sheets: [
        {
          name: "Test",
          rows: [
            ["A1", "B1", "C1", "D1"],
            ["A2", "B2", "C2", "D2"],
            ["A3", "B3", "C3", "D3"],
            ["A4", "B4", "C4", "D4"],
          ],
        },
      ],
    }

    const buffer = await writeXlsx(options)
    const workbook = await readXlsx(buffer, { range: "B2:C3" })

    const sheet = workbook.sheets[0]
    // Rows outside range should be empty
    // Row 0 (A1-D1): completely outside
    // Row 1 (A2-D2): B2, C2 within range
    // Row 2 (A3-D3): B3, C3 within range
    // Row 3 (A4-D4): completely outside

    // Check that B2 and C2 are present (row 1, cols 1-2)
    expect(sheet.rows[1]?.[1]).toBe("B2")
    expect(sheet.rows[1]?.[2]).toBe("C2")
    expect(sheet.rows[2]?.[1]).toBe("B3")
    expect(sheet.rows[2]?.[2]).toBe("C3")

    // Cells outside range should be null/undefined/missing
    expect(sheet.rows[0]?.[0] ?? null).toBe(null)
    expect(sheet.rows[1]?.[0] ?? null).toBe(null)
    // Row 3, col 1 should be null or not populated (no D4 data)
    expect(sheet.rows[3]?.[1] ?? null).toBe(null)
  })
})

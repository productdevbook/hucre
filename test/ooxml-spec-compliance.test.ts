import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { readXlsx } from "../src/xlsx/reader"
import { isDateStyle, parseStyles, resolveStyle } from "../src/xlsx/styles"
import { createStylesCollector } from "../src/xlsx/styles-writer"

// ── Helpers ─────────────────────────────────────────────────────────

const enc = new TextEncoder()

function textToBytes(text: string): Uint8Array {
  return enc.encode(text)
}

// ── Bug #103: OOXML Strict Namespace Support ────────────────────────

describe("Bug #103: OOXML Strict namespace support", () => {
  it("reads an XLSX with Strict namespace relationship types", async () => {
    const writer = new ZipWriter()

    // [Content_Types].xml — uses the same namespace in both modes
    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`

    writer.add("[Content_Types].xml", textToBytes(contentTypesXml), { compress: false })

    // _rels/.rels — use STRICT namespace for officeDocument relationship
    const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

    writer.add("_rels/.rels", textToBytes(rootRelsXml), { compress: false })

    // xl/workbook.xml — uses strict SpreadsheetML namespace
    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
          xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`

    writer.add("xl/workbook.xml", textToBytes(workbookXml), { compress: false })

    // xl/_rels/workbook.xml.rels — use STRICT namespace for all relationship types
    const wbRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/styles" Target="styles.xml"/>
</Relationships>`

    writer.add("xl/_rels/workbook.xml.rels", textToBytes(wbRelsXml), { compress: false })

    // xl/sharedStrings.xml
    const ssXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" count="1" uniqueCount="1">
  <si><t>Hello Strict</t></si>
</sst>`

    writer.add("xl/sharedStrings.xml", textToBytes(ssXml), { compress: false })

    // xl/styles.xml
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`

    writer.add("xl/styles.xml", textToBytes(stylesXml), { compress: false })

    // xl/worksheets/sheet1.xml
    const wsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>`

    writer.add("xl/worksheets/sheet1.xml", textToBytes(wsXml), { compress: false })

    const xlsxData = await writer.build()
    const workbook = await readXlsx(xlsxData)

    expect(workbook.sheets).toHaveLength(1)
    expect(workbook.sheets[0].name).toBe("Sheet1")
    expect(workbook.sheets[0].rows[0][0]).toBe("Hello Strict")
    expect(workbook.sheets[0].rows[0][1]).toBe(42)
  })

  it("reads an XLSX with mixed Transitional and Strict namespaces", async () => {
    const writer = new ZipWriter()

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`

    writer.add("[Content_Types].xml", textToBytes(contentTypesXml), { compress: false })

    // Root rels: Transitional
    const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

    writer.add("_rels/.rels", textToBytes(rootRelsXml), { compress: false })

    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`

    writer.add("xl/workbook.xml", textToBytes(workbookXml), { compress: false })

    // Workbook rels: Strict namespace for worksheet
    const wbRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/styles" Target="styles.xml"/>
</Relationships>`

    writer.add("xl/_rels/workbook.xml.rels", textToBytes(wbRelsXml), { compress: false })

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`

    writer.add("xl/styles.xml", textToBytes(stylesXml), { compress: false })

    const wsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>100</v></c></row>
  </sheetData>
</worksheet>`

    writer.add("xl/worksheets/sheet1.xml", textToBytes(wsXml), { compress: false })

    const xlsxData = await writer.build()
    const workbook = await readXlsx(xlsxData)

    expect(workbook.sheets).toHaveLength(1)
    expect(workbook.sheets[0].name).toBe("Data")
    expect(workbook.sheets[0].rows[0][0]).toBe(100)
  })
})

// ── Bug #104: Built-in Number Format IDs 5-8 ───────────────────────

describe("Bug #104: Built-in number format IDs 5-8", () => {
  it("format IDs 5-8 are NOT identified as date formats", () => {
    // Build a minimal styles structure with cellXfs referencing IDs 5-8
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="5" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="6" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="7" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="8" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`

    const styles = parseStyles(stylesXml)

    // IDs 5-8 are currency formats, NOT date formats
    expect(isDateStyle(styles, 1)).toBe(false) // numFmtId=5
    expect(isDateStyle(styles, 2)).toBe(false) // numFmtId=6
    expect(isDateStyle(styles, 3)).toBe(false) // numFmtId=7
    expect(isDateStyle(styles, 4)).toBe(false) // numFmtId=8
  })

  it("format IDs 14-22 ARE identified as date formats", () => {
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="22" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`

    const styles = parseStyles(stylesXml)

    expect(isDateStyle(styles, 1)).toBe(true) // numFmtId=14 (m/d/yyyy)
    expect(isDateStyle(styles, 2)).toBe(true) // numFmtId=22 (m/d/yyyy h:mm)
  })

  it("format IDs 0-4, 9-13 are NOT date formats", () => {
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="11">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="2" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="3" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="4" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="9" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="10" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="11" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="12" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="13" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="49" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`

    const styles = parseStyles(stylesXml)

    // General (0), number (1-4), percentage (9-10), scientific (11), fraction (12-13), text (49)
    expect(isDateStyle(styles, 0)).toBe(false) // numFmtId=0
    expect(isDateStyle(styles, 1)).toBe(false) // numFmtId=1
    expect(isDateStyle(styles, 2)).toBe(false) // numFmtId=2
    expect(isDateStyle(styles, 3)).toBe(false) // numFmtId=3
    expect(isDateStyle(styles, 4)).toBe(false) // numFmtId=4
    expect(isDateStyle(styles, 5)).toBe(false) // numFmtId=9
    expect(isDateStyle(styles, 6)).toBe(false) // numFmtId=10
    expect(isDateStyle(styles, 7)).toBe(false) // numFmtId=11
    expect(isDateStyle(styles, 8)).toBe(false) // numFmtId=12
    expect(isDateStyle(styles, 9)).toBe(false) // numFmtId=13
    expect(isDateStyle(styles, 10)).toBe(false) // numFmtId=49
  })

  it("resolves built-in format strings for IDs 5-8 in styles", () => {
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="5" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="6" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="7" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="8" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`

    const styles = parseStyles(stylesXml)

    // Each style should have a numFmt string (currency format)
    const style5 = resolveStyle(styles, 1)
    const style6 = resolveStyle(styles, 2)
    const style7 = resolveStyle(styles, 3)
    const style8 = resolveStyle(styles, 4)

    expect(style5.numFmt).toBeDefined()
    expect(style6.numFmt).toBeDefined()
    expect(style7.numFmt).toBeDefined()
    expect(style8.numFmt).toBeDefined()

    // They should contain dollar sign (currency format)
    expect(style5.numFmt).toContain("$")
    expect(style6.numFmt).toContain("$")
    expect(style7.numFmt).toContain("$")
    expect(style8.numFmt).toContain("$")
  })
})

// ── Bug #106: Missing cellStyles Element in styles.xml ──────────────

describe("Bug #106: Missing cellStyles element in styles.xml", () => {
  it("styles.xml output contains cellStyles element", () => {
    const collector = createStylesCollector()
    const xml = collector.toXml()

    expect(xml).toContain("<cellStyles")
    expect(xml).toContain('name="Normal"')
    expect(xml).toContain('xfId="0"')
    expect(xml).toContain('builtinId="0"')
  })

  it("styles.xml output contains dxfs element even when empty", () => {
    const collector = createStylesCollector()
    const xml = collector.toXml()

    expect(xml).toContain("dxfs")
    expect(xml).toContain('count="0"')
  })

  it("styles.xml output contains tableStyles element", () => {
    const collector = createStylesCollector()
    const xml = collector.toXml()

    expect(xml).toContain("tableStyles")
    expect(xml).toContain('defaultTableStyle="TableStyleMedium2"')
    expect(xml).toContain('defaultPivotStyle="PivotStyleLight16"')
  })

  it("styles.xml has dxfs with correct count when dxfs are added", () => {
    const collector = createStylesCollector()
    collector.addDxf({ font: { bold: true } })
    collector.addDxf({ font: { italic: true } })
    const xml = collector.toXml()

    // Should have dxfs with count="2"
    expect(xml).toMatch(/<dxfs count="2">/)
  })

  it("styles.xml preserves correct element ordering", () => {
    const collector = createStylesCollector()
    const xml = collector.toXml()

    // OOXML requires specific element ordering:
    // fonts, fills, borders, cellStyleXfs, cellXfs, cellStyles, dxfs, tableStyles
    const fontsIdx = xml.indexOf("<fonts")
    const fillsIdx = xml.indexOf("<fills")
    const bordersIdx = xml.indexOf("<borders")
    const cellStyleXfsIdx = xml.indexOf("<cellStyleXfs")
    const cellXfsIdx = xml.indexOf("<cellXfs")
    const cellStylesIdx = xml.indexOf("<cellStyles")
    const dxfsIdx = xml.indexOf("dxfs")
    const tableStylesIdx = xml.indexOf("tableStyles")

    expect(fontsIdx).toBeLessThan(fillsIdx)
    expect(fillsIdx).toBeLessThan(bordersIdx)
    expect(bordersIdx).toBeLessThan(cellStyleXfsIdx)
    expect(cellStyleXfsIdx).toBeLessThan(cellXfsIdx)
    expect(cellXfsIdx).toBeLessThan(cellStylesIdx)
    expect(cellStylesIdx).toBeLessThan(dxfsIdx)
    expect(dxfsIdx).toBeLessThan(tableStylesIdx)
  })
})

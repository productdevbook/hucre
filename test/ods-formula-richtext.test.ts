// Regression tests for #405 — the ODS writer emitted cross-sheet formulas
// ODF cannot parse, and dropped rich-text cells entirely.
//
// The formula assertions deliberately check the *emitted* attribute against
// what OpenFormula specifies, and the decoding assertions run against
// content.xml written by hand in LibreOffice's spelling. A round-trip test
// built only on hucre's own output passed happily while the writer and the
// reader shared the same wrong idea of the syntax.

import { describe, it, expect } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { parseXml } from "../src/xml/parser"
import type { Cell, WriteSheet } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

const encoder = new TextEncoder()
const decoder = new TextDecoder("utf-8")

async function extractFile(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data)
  return decoder.decode(await zip.extract(path))
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

/** Write one sheet whose A1 carries `formula`, and return A1's attributes. */
async function writeFormulaCell(formula: string): Promise<Record<string, string>> {
  const cells = new Map<string, Partial<Cell>>([["0,0", { value: 0, formula }]])
  const data = await writeOds({ sheets: [{ name: "Sheet1", rows: [[0]], cells }] })
  const doc = parseXml(await extractFile(data, "content.xml"))
  const table = findChild(findChild(findChild(doc, "body"), "spreadsheet"), "table")
  const row = findChild(table, "table-row")
  return findChild(row, "table-cell").attrs
}

/** The namespaces a real content.xml declares, for the hand-written fixtures. */
const NS = [
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"`,
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"`,
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"`,
  `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"`,
  `xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"`,
  `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"`,
  `xmlns:xlink="http://www.w3.org/1999/xlink"`,
].join(" ")

/** A minimal .ods holding one cell with the given `table:formula`. */
async function odsWithFormula(formula: string): Promise<Uint8Array> {
  const content =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content ${NS} office:version="1.3"><office:body><office:spreadsheet>` +
    `<table:table table:name="Sheet1"><table:table-row>` +
    `<table:table-cell table:formula="${formula.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;")}" ` +
    `office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell>` +
    `</table:table-row></table:table>` +
    `</office:spreadsheet></office:body></office:document-content>`

  const zip = new ZipWriter()
  zip.add("mimetype", encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), {
    compress: false,
  })
  zip.add("content.xml", encoder.encode(content))
  return await zip.build()
}

async function readFormula(odsFormula: string): Promise<string | undefined> {
  const wb = await readOds(await odsWithFormula(odsFormula))
  return wb.sheets[0]!.cells?.get("0,0")?.formula
}

// ── #405.1: cross-sheet formula references ──────────────────────────

describe("ODS #405 — cross-sheet formulas are bracketed as a whole", () => {
  it("brackets the sheet together with the cell", async () => {
    const attrs = await writeFormulaCell("Sheet2!A1+1")
    // OpenFormula: Reference ::= '[' Source? RangeAddress ']', and the
    // sheet locator lives inside the brackets.
    expect(attrs["table:formula"]).toBe("of:=[$Sheet2.A1]+1")
    // The defect emitted "of:=Sheet2.[.A1]+1", which parses as nothing.
    expect(attrs["table:formula"]).not.toContain("Sheet2.[.")
  })

  it("accepts the dot spelling of a sheet qualifier too", async () => {
    // The form used in LibreOffice's own UI, and in the report of #405.
    const attrs = await writeFormulaCell("Sheet2.A1+1")
    expect(attrs["table:formula"]).toBe("of:=[$Sheet2.A1]+1")
  })

  it("omits the sheet on the second half of a cross-sheet range", async () => {
    const attrs = await writeFormulaCell("SUM(Sheet2!A1:B2)")
    expect(attrs["table:formula"]).toBe("of:=SUM([$Sheet2.A1:.B2])")
  })

  it("keeps absolute markers inside the brackets", async () => {
    const attrs = await writeFormulaCell("Sheet2!$A$1")
    expect(attrs["table:formula"]).toBe("of:=[$Sheet2.$A$1]")
  })

  it("quotes a sheet name with a space the way ODF does", async () => {
    const attrs = await writeFormulaCell("'My Sheet'!A1*2")
    expect(attrs["table:formula"]).toBe("of:=[$'My Sheet'.A1]*2")
  })

  it("leaves local references in the plain [.A1] form", async () => {
    const attrs = await writeFormulaCell("SUM(A1:B2)+C3")
    expect(attrs["table:formula"]).toBe("of:=SUM([.A1:.B2])+[.C3]")
  })

  it("still leaves function names and string literals alone", async () => {
    const attrs = await writeFormulaCell('LOG10(A1)&"Sheet2.A1"')
    expect(attrs["table:formula"]).toBe('of:=LOG10([.A1])&"Sheet2.A1"')
  })
})

describe("ODS #405 — reading the cross-sheet forms LibreOffice writes", () => {
  it("decodes an absolute sheet locator", async () => {
    expect(await readFormula("of:=[$Sheet2.A1]+1")).toBe("Sheet2!A1+1")
  })

  it("decodes a relative sheet locator", async () => {
    expect(await readFormula("of:=[Sheet2.A1]+1")).toBe("Sheet2!A1+1")
  })

  it("decodes a range whose second half inherits the sheet", async () => {
    expect(await readFormula("of:=SUM([$Sheet2.A1:.B2])")).toBe("SUM(Sheet2!A1:B2)")
  })

  it("decodes a quoted sheet name", async () => {
    expect(await readFormula("of:=[$'My Sheet'.A1]*2")).toBe("'My Sheet'!A1*2")
  })

  it("still decodes the local forms", async () => {
    expect(await readFormula("of:=SUM([.A1:.A10])")).toBe("SUM(A1:A10)")
    expect(await readFormula("of:=[.$B$4]")).toBe("$B$4")
  })

  it("leaves a bracketed token that is not an address verbatim", async () => {
    // A broken reference, and an external one — mangling either would be
    // worse than passing it through.
    expect(await readFormula("of:=[#REF!]+1")).toBe("[#REF!]+1")
    expect(await readFormula("of:=['budget.ods'#$Sheet1.A1]")).toBe("['budget.ods'#$Sheet1.A1]")
  })

  it("leaves a bracketed token inside a string literal alone", async () => {
    expect(await readFormula('of:=IF([.A1]="[$Sheet2.A1]";1;2)')).toBe('IF(A1="[$Sheet2.A1]";1;2)')
  })

  it("round-trips a cross-sheet formula through write → read", async () => {
    const cells = new Map<string, Partial<Cell>>([["0,0", { value: 6, formula: "Sheet2!A1+1" }]])
    const sheets: WriteSheet[] = [
      { name: "Sheet1", rows: [[6]], cells },
      { name: "Sheet2", rows: [[5]] },
    ]
    const wb = await readOds(await writeOds({ sheets }))
    expect(wb.sheets[0]!.cells!.get("0,0")!.formula).toBe("Sheet2!A1+1")
  })
})

// ── #405.2: rich-text cells ─────────────────────────────────────────

describe("ODS #405 — rich-text cells keep their text", () => {
  const runs = [
    { text: "Hello ", font: { bold: true } },
    { text: "world", font: { italic: true, color: { rgb: "FF0000" } } },
  ]

  async function writeRichText(): Promise<Uint8Array> {
    const cells = new Map<string, Partial<Cell>>([["0,0", { value: null, richText: runs }]])
    return await writeOds({ sheets: [{ name: "Sheet1", rows: [[null]], cells }] })
  }

  it("writes the runs as text:span rather than an empty cell", async () => {
    const xml = await extractFile(await writeRichText(), "content.xml")
    // The defect wrote "<table:table-cell/>" and the text was nowhere in
    // the document.
    expect(xml).toContain("Hello ")
    expect(xml).toContain("world")
    expect(xml).toContain("<text:span")
    expect(xml).toContain('office:value-type="string"')
  })

  it("gives each run a text-family automatic style", async () => {
    const doc = parseXml(await extractFile(await writeRichText(), "content.xml"))
    const autoStyles = findChild(doc, "automatic-styles")
    const textStyles = findChildren(autoStyles, "style").filter(
      (s: any) => s.attrs["style:family"] === "text",
    )
    expect(textStyles).toHaveLength(2)
    const props = textStyles.map((s: any) => findChild(s, "text-properties").attrs)
    expect(props[0]["fo:font-weight"]).toBe("bold")
    expect(props[1]["fo:font-style"]).toBe("italic")
    expect(props[1]["fo:color"]).toBe("#FF0000")
  })

  it("reads back as the concatenated run text", async () => {
    // The ODS reader flattens spans (collectText), so the runs come back as
    // one string rather than as `richText` — the point is that the text
    // survives at all. It used to read back as undefined.
    const wb = await readOds(await writeRichText())
    expect(wb.sheets[0]!.rows[0]![0]).toBe("Hello world")
  })
})

// ── #405.3: hyperlinks on typed cells, display text, format sections ─

describe("ODS #405 — hyperlinks on non-string cells", () => {
  it("keeps the link and the typed value on numbers, dates and booleans", async () => {
    const date = new Date(Date.UTC(2020, 0, 2))
    const cells = new Map<string, Partial<Cell>>([
      ["0,0", { value: 42, hyperlink: { target: "https://example.com/num" } }],
      ["0,1", { value: date, hyperlink: { target: "https://example.com/date" } }],
      ["0,2", { value: true, hyperlink: { target: "https://example.com/bool" } }],
    ])
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[null, null, null]], cells }],
    })

    const wb = await readOds(data)
    const sheet = wb.sheets[0]!
    expect(sheet.cells!.get("0,0")!.hyperlink!.target).toBe("https://example.com/num")
    expect(sheet.cells!.get("0,1")!.hyperlink!.target).toBe("https://example.com/date")
    expect(sheet.cells!.get("0,2")!.hyperlink!.target).toBe("https://example.com/bool")

    // The value must survive as its own type, not degrade to text.
    expect(sheet.rows[0]![0]).toBe(42)
    // (ODS date values carry no time zone, so compare the calendar date the
    // reader parsed rather than the instant.)
    const readDate = sheet.rows[0]![1] as Date
    expect(readDate).toBeInstanceOf(Date)
    expect([readDate.getFullYear(), readDate.getMonth(), readDate.getDate()]).toEqual([2020, 0, 2])
    expect(sheet.rows[0]![2]).toBe(true)
  })

  it("writes Hyperlink.display as the anchor text", async () => {
    const cells = new Map<string, Partial<Cell>>([
      ["0,0", { value: 42, hyperlink: { target: "https://example.com", display: "Click me" } }],
    ])
    const data = await writeOds({ sheets: [{ name: "Sheet1", rows: [[null]], cells }] })

    expect(await extractFile(data, "content.xml")).toContain(">Click me</text:a>")
    const wb = await readOds(data)
    expect(wb.sheets[0]!.cells!.get("0,0")!.hyperlink!.display).toBe("Click me")
    expect(wb.sheets[0]!.rows[0]![0]).toBe(42)
  })
})

describe("ODS #405 — multi-section number formats", () => {
  async function styleXml(numFmt: string): Promise<any[]> {
    const cells = new Map<string, Partial<Cell>>([["0,0", { value: -1234.5, style: { numFmt } }]])
    const data = await writeOds({ sheets: [{ name: "Sheet1", rows: [[-1234.5]], cells }] })
    const doc = parseXml(await extractFile(data, "content.xml"))
    return findChildren(findChild(doc, "automatic-styles"), "number-style")
  }

  it("gives the negative and zero sections their own mapped data styles", async () => {
    const styles = await styleXml("#,##0.00;-#,##0.00;0.00")
    expect(styles).toHaveLength(3)

    // The style a cell points at is the one without style:volatile — the
    // negative section — and it names the others through <style:map>.
    const main = styles.find((s: any) => s.attrs["style:volatile"] !== "true")!
    const maps = findChildren(main, "map")
    expect(maps.map((m: any) => m.attrs["style:condition"])).toEqual(["value()>0", "value()=0"])
    for (const m of maps) {
      const target = styles.find(
        (s: any) => s.attrs["style:name"] === m.attrs["style:apply-style-name"],
      )
      expect(target.attrs["style:volatile"]).toBe("true")
    }

    // The negative section carries its own minus: a style reached through a
    // map formats the value as it stands.
    expect(findChild(main, "text").children.join("")).toBe("-")
  })

  it("round-trips all three sections", async () => {
    const numFmt = "#,##0.00;-#,##0.00;0.00"
    const cells = new Map<string, Partial<Cell>>([["0,0", { value: 1, style: { numFmt } }]])
    const data = await writeOds({ sheets: [{ name: "Sheet1", rows: [[1]], cells }] })
    const wb = await readOds(data, { readStyles: true })
    expect(wb.sheets[0]!.cells!.get("0,0")!.style!.numFmt).toBe(numFmt)
  })

  it("leaves a single-section format as one unmapped style", async () => {
    const styles = await styleXml("#,##0.00")
    expect(styles).toHaveLength(1)
    expect(findChildren(styles[0], "map")).toHaveLength(0)
  })
})

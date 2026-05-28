import { describe, it, expect } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { parseXml } from "../src/xml/parser"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import type { NamedRange } from "../src/_types"

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

function getElementText(el: { children: Array<unknown> }): string {
  return el.children.filter((c: unknown) => typeof c === "string").join("")
}

// ── Writing Tests ────────────────────────────────────────────────────

describe("named ranges — writing", () => {
  it("writes workbook-level named range", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["A", "B", "C"]] }],
      namedRanges: [{ name: "ProductList", range: "Sheet1!$A$1:$A$100" }],
    })

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const dn = findChild(doc, "definedNames")
    expect(dn).toBeDefined()

    const defs = findChildren(dn, "definedName")
    expect(defs.length).toBe(1)
    expect(defs[0].attrs["name"]).toBe("ProductList")
    expect(getElementText(defs[0])).toBe("Sheet1!$A$1:$A$100")
    // No localSheetId for workbook-level
    expect(defs[0].attrs["localSheetId"]).toBeUndefined()
  })

  it("writes sheet-scoped named range (localSheetId)", async () => {
    const data = await writeXlsx({
      sheets: [
        { name: "Sheet1", rows: [["Data"]] },
        { name: "Sheet2", rows: [["More"]] },
      ],
      namedRanges: [{ name: "Total", range: "Sheet1!$B$101", scope: "Sheet1" }],
    })

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const dn = findChild(doc, "definedNames")
    const defs = findChildren(dn, "definedName")
    expect(defs.length).toBe(1)
    expect(defs[0].attrs["name"]).toBe("Total")
    expect(defs[0].attrs["localSheetId"]).toBe("0")
    expect(getElementText(defs[0])).toBe("Sheet1!$B$101")
  })

  it("writes multiple named ranges", async () => {
    const data = await writeXlsx({
      sheets: [
        { name: "Sheet1", rows: [["A"]] },
        { name: "Sheet2", rows: [["B"]] },
      ],
      namedRanges: [
        { name: "Global1", range: "Sheet1!$A$1:$A$100" },
        { name: "Local1", range: "Sheet2!$B$1:$B$50", scope: "Sheet2" },
        { name: "Global2", range: "Sheet1!$C$1:$D$100" },
      ],
    })

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const dn = findChild(doc, "definedNames")
    const defs = findChildren(dn, "definedName")
    expect(defs.length).toBe(3)

    // Global1 — no localSheetId
    const global1 = defs.find((d: any) => d.attrs["name"] === "Global1")
    expect(global1).toBeDefined()
    expect(global1.attrs["localSheetId"]).toBeUndefined()

    // Local1 — localSheetId = 1 (Sheet2 is index 1)
    const local1 = defs.find((d: any) => d.attrs["name"] === "Local1")
    expect(local1).toBeDefined()
    expect(local1.attrs["localSheetId"]).toBe("1")

    // Global2 — no localSheetId
    const global2 = defs.find((d: any) => d.attrs["name"] === "Global2")
    expect(global2).toBeDefined()
    expect(global2.attrs["localSheetId"]).toBeUndefined()
  })

  it("writes named range with comment", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Val"]] }],
      namedRanges: [{ name: "Budget", range: "Sheet1!$A$1:$D$10", comment: "Annual budget data" }],
    })

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const dn = findChild(doc, "definedNames")
    const defs = findChildren(dn, "definedName")
    expect(defs[0].attrs["comment"]).toBe("Annual budget data")
  })

  it("writes print area from pageSetup", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          pageSetup: { printArea: "$A$1:$D$50" },
        },
      ],
    })

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const dn = findChild(doc, "definedNames")
    expect(dn).toBeDefined()

    const defs = findChildren(dn, "definedName")
    const printArea = defs.find((d: any) => d.attrs["name"] === "_xlnm.Print_Area")
    expect(printArea).toBeDefined()
    expect(printArea.attrs["localSheetId"]).toBe("0")
    expect(getElementText(printArea)).toBe("Sheet1!$A$1:$D$50")
  })

  it("writes print titles from pageSetup (row)", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Header"]],
          pageSetup: { printTitlesRow: "$1:$1" },
        },
      ],
    })

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const dn = findChild(doc, "definedNames")
    const defs = findChildren(dn, "definedName")
    const printTitles = defs.find((d: any) => d.attrs["name"] === "_xlnm.Print_Titles")
    expect(printTitles).toBeDefined()
    expect(printTitles.attrs["localSheetId"]).toBe("0")
    expect(getElementText(printTitles)).toBe("Sheet1!$1:$1")
  })

  it("writes print titles from pageSetup (row + column)", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Report",
          rows: [["Header"]],
          pageSetup: { printTitlesRow: "$1:$2", printTitlesColumn: "$A:$B" },
        },
      ],
    })

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const dn = findChild(doc, "definedNames")
    const defs = findChildren(dn, "definedName")
    const printTitles = defs.find((d: any) => d.attrs["name"] === "_xlnm.Print_Titles")
    expect(printTitles).toBeDefined()
    expect(getElementText(printTitles)).toBe("Report!$1:$2,Report!$A:$B")
  })

  it("does not emit definedNames when none exist", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello"]] }],
    })

    const xml = await extractXml(data, "xl/workbook.xml")
    const doc = parseXml(xml)

    const dn = findChild(doc, "definedNames")
    expect(dn).toBeUndefined()
  })
})

// ── Reading Tests ────────────────────────────────────────────────────

describe("named ranges — reading (round-trip)", () => {
  it("round-trips workbook-level named range", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["A"]] }],
      namedRanges: [{ name: "ProductList", range: "Sheet1!$A$1:$A$100" }],
    })

    const workbook = await readXlsx(data)
    expect(workbook.namedRanges).toBeDefined()
    expect(workbook.namedRanges!.length).toBe(1)

    const nr = workbook.namedRanges![0]
    expect(nr.name).toBe("ProductList")
    expect(nr.range).toBe("Sheet1!$A$1:$A$100")
    expect(nr.scope).toBeUndefined()
  })

  it("round-trips sheet-scoped named range", async () => {
    const data = await writeXlsx({
      sheets: [
        { name: "Sheet1", rows: [["A"]] },
        { name: "Sheet2", rows: [["B"]] },
      ],
      namedRanges: [{ name: "Total", range: "Sheet1!$B$101", scope: "Sheet1" }],
    })

    const workbook = await readXlsx(data)
    expect(workbook.namedRanges!.length).toBe(1)

    const nr = workbook.namedRanges![0]
    expect(nr.name).toBe("Total")
    expect(nr.range).toBe("Sheet1!$B$101")
    expect(nr.scope).toBe("Sheet1")
  })

  it("round-trips multiple named ranges", async () => {
    const ranges: NamedRange[] = [
      { name: "Global1", range: "Sheet1!$A$1:$A$100" },
      { name: "Local1", range: "Sheet2!$B$1:$B$50", scope: "Sheet2" },
      { name: "Global2", range: "Sheet1!$C$1:$D$100" },
    ]

    const data = await writeXlsx({
      sheets: [
        { name: "Sheet1", rows: [["A"]] },
        { name: "Sheet2", rows: [["B"]] },
      ],
      namedRanges: ranges,
    })

    const workbook = await readXlsx(data)
    expect(workbook.namedRanges!.length).toBe(3)

    const byName = new Map(workbook.namedRanges!.map((nr) => [nr.name, nr]))
    expect(byName.get("Global1")!.range).toBe("Sheet1!$A$1:$A$100")
    expect(byName.get("Global1")!.scope).toBeUndefined()
    expect(byName.get("Local1")!.range).toBe("Sheet2!$B$1:$B$50")
    expect(byName.get("Local1")!.scope).toBe("Sheet2")
    expect(byName.get("Global2")!.range).toBe("Sheet1!$C$1:$D$100")
  })

  it("round-trips named range with comment", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Val"]] }],
      namedRanges: [{ name: "Budget", range: "Sheet1!$A$1:$D$10", comment: "Annual budget data" }],
    })

    const workbook = await readXlsx(data)
    const nr = workbook.namedRanges![0]
    expect(nr.name).toBe("Budget")
    expect(nr.comment).toBe("Annual budget data")
  })

  it("reads print area as named range", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          pageSetup: { printArea: "$A$1:$D$50" },
        },
      ],
    })

    const workbook = await readXlsx(data)
    expect(workbook.namedRanges).toBeDefined()

    const printArea = workbook.namedRanges!.find((nr) => nr.name === "_xlnm.Print_Area")
    expect(printArea).toBeDefined()
    expect(printArea!.range).toBe("Sheet1!$A$1:$D$50")
    expect(printArea!.scope).toBe("Sheet1")
  })

  it("reads print titles as named range", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Header"]],
          pageSetup: { printTitlesRow: "$1:$1" },
        },
      ],
    })

    const workbook = await readXlsx(data)

    const printTitles = workbook.namedRanges!.find((nr) => nr.name === "_xlnm.Print_Titles")
    expect(printTitles).toBeDefined()
    expect(printTitles!.range).toBe("Sheet1!$1:$1")
    expect(printTitles!.scope).toBe("Sheet1")
  })

  it("no namedRanges property when none exist", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello"]] }],
    })

    const workbook = await readXlsx(data)
    expect(workbook.namedRanges).toBeUndefined()
  })
})

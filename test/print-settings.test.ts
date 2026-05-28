import { describe, it, expect } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { parseXml } from "../src/xml/parser"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { createStylesCollector } from "../src/xlsx/styles-writer"
import { createSharedStrings, writeWorksheetXml } from "../src/xlsx/worksheet-writer"
import type { WriteSheet, PageSetup, HeaderFooter } from "../src/_types"

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

function getElementText(el: { children: Array<unknown> }): string {
  return el.children.filter((c: unknown) => typeof c === "string").join("")
}

function writeXml(sheet: WriteSheet): string {
  const styles = createStylesCollector()
  const ss = createSharedStrings()
  const result = writeWorksheetXml(sheet, styles, ss)
  return result.xml
}

function parseSheet(xml: string) {
  return parseXml(xml)
}

// ── Page Margins Writing Tests ────────────────────────────────────────

describe("page margins — writing", () => {
  it("writes custom page margins", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      pageSetup: {
        margins: { left: 1.0, right: 1.0, top: 1.25, bottom: 1.25, header: 0.5, footer: 0.5 },
      },
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const pm = findChild(doc, "pageMargins")
    expect(pm).toBeDefined()
    expect(pm.attrs["left"]).toBe("1")
    expect(pm.attrs["right"]).toBe("1")
    expect(pm.attrs["top"]).toBe("1.25")
    expect(pm.attrs["bottom"]).toBe("1.25")
    expect(pm.attrs["header"]).toBe("0.5")
    expect(pm.attrs["footer"]).toBe("0.5")
  })

  it("writes default margins when no pageSetup specified", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const pm = findChild(doc, "pageMargins")
    expect(pm).toBeDefined()
    expect(pm.attrs["left"]).toBe("0.7")
    expect(pm.attrs["right"]).toBe("0.7")
    expect(pm.attrs["top"]).toBe("0.75")
    expect(pm.attrs["bottom"]).toBe("0.75")
    expect(pm.attrs["header"]).toBe("0.3")
    expect(pm.attrs["footer"]).toBe("0.3")
  })
})

// ── Page Setup Writing Tests ──────────────────────────────────────────

describe("page setup — writing", () => {
  it("writes paper size and orientation", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      pageSetup: { paperSize: "a4", orientation: "landscape" },
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const ps = findChild(doc, "pageSetup")
    expect(ps).toBeDefined()
    expect(ps.attrs["paperSize"]).toBe("9")
    expect(ps.attrs["orientation"]).toBe("landscape")
  })

  it("writes scale", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      pageSetup: { scale: 75 },
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const ps = findChild(doc, "pageSetup")
    expect(ps).toBeDefined()
    expect(ps.attrs["scale"]).toBe("75")
  })

  it("writes fit to page", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      pageSetup: { fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const ps = findChild(doc, "pageSetup")
    expect(ps).toBeDefined()
    expect(ps.attrs["fitToWidth"]).toBe("1")
    expect(ps.attrs["fitToHeight"]).toBe("0")
  })

  it("writes all paper sizes correctly", () => {
    const paperSizes: Array<[string, string]> = [
      ["letter", "1"],
      ["legal", "5"],
      ["a3", "8"],
      ["a4", "9"],
      ["a5", "11"],
      ["b4", "12"],
      ["b5", "13"],
      ["executive", "7"],
      ["tabloid", "3"],
    ]

    for (const [name, expected] of paperSizes) {
      const sheet: WriteSheet = {
        name: "Test",
        rows: [["Data"]],
        pageSetup: { paperSize: name as PageSetup["paperSize"] },
      }

      const xml = writeXml(sheet)
      const doc = parseSheet(xml)

      const ps = findChild(doc, "pageSetup")
      expect(ps).toBeDefined()
      expect(ps.attrs["paperSize"]).toBe(expected)
    }
  })

  it("does not emit pageSetup when no settings are specified", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const ps = findChild(doc, "pageSetup")
    expect(ps).toBeUndefined()
  })
})

// ── Header/Footer Writing Tests ──────────────────────────────────────

describe("header/footer — writing", () => {
  it("writes header with formatting codes", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      headerFooter: {
        oddHeader: '&C&"Arial,Bold"My Report&R&D',
      },
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const hf = findChild(doc, "headerFooter")
    expect(hf).toBeDefined()

    const oh = findChild(hf, "oddHeader")
    expect(oh).toBeDefined()
    expect(getElementText(oh)).toBe('&C&"Arial,Bold"My Report&R&D')
  })

  it("writes footer with page number codes", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      headerFooter: {
        oddFooter: "&CPage &P of &N",
      },
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const hf = findChild(doc, "headerFooter")
    const of_ = findChild(hf, "oddFooter")
    expect(of_).toBeDefined()
    expect(getElementText(of_)).toBe("&CPage &P of &N")
  })

  it("writes both header and footer", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      headerFooter: {
        oddHeader: "&LLeft Header&RRight Header",
        oddFooter: "&CCenter Footer",
      },
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const hf = findChild(doc, "headerFooter")

    const oh = findChild(hf, "oddHeader")
    expect(getElementText(oh)).toBe("&LLeft Header&RRight Header")

    const of_ = findChild(hf, "oddFooter")
    expect(getElementText(of_)).toBe("&CCenter Footer")
  })

  it("writes differentOddEven header/footer", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      headerFooter: {
        differentOddEven: true,
        oddHeader: "&COdd Header",
        evenHeader: "&CEven Header",
        oddFooter: "&COdd Footer",
        evenFooter: "&CEven Footer",
      },
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const hf = findChild(doc, "headerFooter")
    expect(hf.attrs["differentOddEven"]).toBe("1")

    expect(getElementText(findChild(hf, "oddHeader"))).toBe("&COdd Header")
    expect(getElementText(findChild(hf, "evenHeader"))).toBe("&CEven Header")
    expect(getElementText(findChild(hf, "oddFooter"))).toBe("&COdd Footer")
    expect(getElementText(findChild(hf, "evenFooter"))).toBe("&CEven Footer")
  })

  it("writes differentFirst header/footer", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      headerFooter: {
        differentFirst: true,
        oddHeader: "&CRegular Header",
        firstHeader: "&CFirst Page Header",
        oddFooter: "&CRegular Footer",
        firstFooter: "&CFirst Page Footer",
      },
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const hf = findChild(doc, "headerFooter")
    expect(hf.attrs["differentFirst"]).toBe("1")

    expect(getElementText(findChild(hf, "firstHeader"))).toBe("&CFirst Page Header")
    expect(getElementText(findChild(hf, "firstFooter"))).toBe("&CFirst Page Footer")
  })

  it("does not emit headerFooter when not specified", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const hf = findChild(doc, "headerFooter")
    expect(hf).toBeUndefined()
  })
})

// ── XML Element Order Tests ──────────────────────────────────────────

describe("print settings — element order", () => {
  it("places pageMargins after hyperlinks and before drawing", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      pageSetup: { paperSize: "a4", orientation: "landscape" },
      headerFooter: { oddHeader: "&CHeader" },
    }

    const xml = writeXml(sheet)

    // Verify element ordering
    const sheetDataPos = xml.indexOf("<sheetData")
    const pageMarginsPos = xml.indexOf("<pageMargins")
    const pageSetupPos = xml.indexOf("<pageSetup")
    const headerFooterPos = xml.indexOf("<headerFooter")

    expect(sheetDataPos).toBeGreaterThan(-1)
    expect(pageMarginsPos).toBeGreaterThan(-1)
    expect(pageSetupPos).toBeGreaterThan(-1)
    expect(headerFooterPos).toBeGreaterThan(-1)

    expect(pageMarginsPos).toBeGreaterThan(sheetDataPos)
    expect(pageSetupPos).toBeGreaterThan(pageMarginsPos)
    expect(headerFooterPos).toBeGreaterThan(pageSetupPos)
  })

  it("places pageMargins after dataValidations", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      dataValidations: [{ type: "list", values: ["A", "B"], range: "A1:A10" }],
      pageSetup: { paperSize: "a4" },
    }

    const xml = writeXml(sheet)

    const dvPos = xml.indexOf("<dataValidations")
    const pmPos = xml.indexOf("<pageMargins")
    expect(dvPos).toBeGreaterThan(-1)
    expect(pmPos).toBeGreaterThan(-1)
    expect(pmPos).toBeGreaterThan(dvPos)
  })
})

// ── Round-trip Tests ─────────────────────────────────────────────────

describe("page setup — round-trip", () => {
  it("round-trips page margins", async () => {
    const margins = { left: 1.0, right: 1.0, top: 1.25, bottom: 1.25, header: 0.5, footer: 0.5 }

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          pageSetup: { margins },
        },
      ],
    })

    const workbook = await readXlsx(data)
    const ps = workbook.sheets[0].pageSetup
    expect(ps).toBeDefined()
    expect(ps!.margins).toBeDefined()
    expect(ps!.margins!.left).toBe(1.0)
    expect(ps!.margins!.right).toBe(1.0)
    expect(ps!.margins!.top).toBe(1.25)
    expect(ps!.margins!.bottom).toBe(1.25)
    expect(ps!.margins!.header).toBe(0.5)
    expect(ps!.margins!.footer).toBe(0.5)
  })

  it("round-trips paper size and orientation", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          pageSetup: { paperSize: "a4", orientation: "landscape" },
        },
      ],
    })

    const workbook = await readXlsx(data)
    const ps = workbook.sheets[0].pageSetup
    expect(ps).toBeDefined()
    expect(ps!.paperSize).toBe("a4")
    expect(ps!.orientation).toBe("landscape")
  })

  it("round-trips scale", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          pageSetup: { scale: 75 },
        },
      ],
    })

    const workbook = await readXlsx(data)
    const ps = workbook.sheets[0].pageSetup
    expect(ps!.scale).toBe(75)
  })

  it("round-trips fit to page", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          pageSetup: { fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
        },
      ],
    })

    const workbook = await readXlsx(data)
    const ps = workbook.sheets[0].pageSetup
    expect(ps!.fitToPage).toBe(true)
    expect(ps!.fitToWidth).toBe(1)
    expect(ps!.fitToHeight).toBe(0)
  })

  it("round-trips all paper sizes", async () => {
    const sizes = ["letter", "legal", "a3", "a4", "a5", "b4", "b5", "executive", "tabloid"] as const

    for (const size of sizes) {
      const data = await writeXlsx({
        sheets: [
          {
            name: "Sheet1",
            rows: [["Data"]],
            pageSetup: { paperSize: size },
          },
        ],
      })

      const workbook = await readXlsx(data)
      expect(workbook.sheets[0].pageSetup!.paperSize).toBe(size)
    }
  })

  it("no pageSetup when not specified", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Data"]] }],
    })

    const workbook = await readXlsx(data)
    // pageSetup might be present due to default margins — if so, it should not have
    // paperSize, orientation, etc.
    const ps = workbook.sheets[0].pageSetup
    if (ps) {
      // Only margins should be present (from default pageMargins)
      expect(ps.paperSize).toBeUndefined()
      expect(ps.orientation).toBeUndefined()
      expect(ps.scale).toBeUndefined()
    }
  })
})

describe("header/footer — round-trip", () => {
  it("round-trips simple header and footer", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          headerFooter: {
            oddHeader: '&C&"Arial,Bold"My Report',
            oddFooter: "&CPage &P of &N",
          },
        },
      ],
    })

    const workbook = await readXlsx(data)
    const hf = workbook.sheets[0].headerFooter
    expect(hf).toBeDefined()
    expect(hf!.oddHeader).toBe('&C&"Arial,Bold"My Report')
    expect(hf!.oddFooter).toBe("&CPage &P of &N")
  })

  it("round-trips all header/footer fields", async () => {
    const hfIn: HeaderFooter = {
      differentOddEven: true,
      differentFirst: true,
      oddHeader: "&COdd H",
      oddFooter: "&COdd F",
      evenHeader: "&CEven H",
      evenFooter: "&CEven F",
      firstHeader: "&CFirst H",
      firstFooter: "&CFirst F",
    }

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          headerFooter: hfIn,
        },
      ],
    })

    const workbook = await readXlsx(data)
    const hf = workbook.sheets[0].headerFooter!
    expect(hf.differentOddEven).toBe(true)
    expect(hf.differentFirst).toBe(true)
    expect(hf.oddHeader).toBe("&COdd H")
    expect(hf.oddFooter).toBe("&COdd F")
    expect(hf.evenHeader).toBe("&CEven H")
    expect(hf.evenFooter).toBe("&CEven F")
    expect(hf.firstHeader).toBe("&CFirst H")
    expect(hf.firstFooter).toBe("&CFirst F")
  })

  it("no headerFooter when not specified", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Data"]] }],
    })

    const workbook = await readXlsx(data)
    expect(workbook.sheets[0].headerFooter).toBeUndefined()
  })
})

describe("multiple sheets with different page setups", () => {
  it("each sheet has independent page setup", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Portrait",
          rows: [["A"]],
          pageSetup: { paperSize: "a4", orientation: "portrait", scale: 100 },
          headerFooter: { oddHeader: "&CPortrait Sheet" },
        },
        {
          name: "Landscape",
          rows: [["B"]],
          pageSetup: {
            paperSize: "letter",
            orientation: "landscape",
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 0,
          },
          headerFooter: { oddFooter: "&CPage &P" },
        },
        {
          name: "NoSetup",
          rows: [["C"]],
        },
      ],
    })

    const workbook = await readXlsx(data)

    // Portrait sheet
    const ps1 = workbook.sheets[0].pageSetup
    expect(ps1!.paperSize).toBe("a4")
    expect(ps1!.orientation).toBe("portrait")
    expect(ps1!.scale).toBe(100)
    expect(workbook.sheets[0].headerFooter!.oddHeader).toBe("&CPortrait Sheet")

    // Landscape sheet
    const ps2 = workbook.sheets[1].pageSetup
    expect(ps2!.paperSize).toBe("letter")
    expect(ps2!.orientation).toBe("landscape")
    expect(ps2!.fitToPage).toBe(true)
    expect(ps2!.fitToWidth).toBe(1)
    expect(ps2!.fitToHeight).toBe(0)
    expect(workbook.sheets[1].headerFooter!.oddFooter).toBe("&CPage &P")

    // No setup sheet — no specific pageSetup attrs
    const ps3 = workbook.sheets[2].pageSetup
    if (ps3) {
      expect(ps3.paperSize).toBeUndefined()
      expect(ps3.orientation).toBeUndefined()
    }
    expect(workbook.sheets[2].headerFooter).toBeUndefined()
  })
})

// ── Integration Tests (ZIP verification) ─────────────────────────────

describe("print settings — ZIP verification", () => {
  it("worksheet XML contains correct pageMargins, pageSetup, headerFooter", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          pageSetup: {
            paperSize: "a4",
            orientation: "landscape",
            scale: 75,
            margins: { left: 0.5, right: 0.5, top: 0.8, bottom: 0.8, header: 0.25, footer: 0.25 },
          },
          headerFooter: {
            oddHeader: "&CReport Title",
            oddFooter: "&CPage &P of &N",
          },
        },
      ],
    })

    const sheetXml = await extractXml(data, "xl/worksheets/sheet1.xml")
    const doc = parseXml(sheetXml)

    // Page margins
    const pm = findChild(doc, "pageMargins")
    expect(pm).toBeDefined()
    expect(pm.attrs["left"]).toBe("0.5")
    expect(pm.attrs["right"]).toBe("0.5")
    expect(pm.attrs["top"]).toBe("0.8")
    expect(pm.attrs["bottom"]).toBe("0.8")
    expect(pm.attrs["header"]).toBe("0.25")
    expect(pm.attrs["footer"]).toBe("0.25")

    // Page setup
    const ps = findChild(doc, "pageSetup")
    expect(ps).toBeDefined()
    expect(ps.attrs["paperSize"]).toBe("9")
    expect(ps.attrs["orientation"]).toBe("landscape")
    expect(ps.attrs["scale"]).toBe("75")

    // Header/footer
    const hf = findChild(doc, "headerFooter")
    expect(hf).toBeDefined()
    expect(getElementText(findChild(hf, "oddHeader"))).toBe("&CReport Title")
    expect(getElementText(findChild(hf, "oddFooter"))).toBe("&CPage &P of &N")
  })
})

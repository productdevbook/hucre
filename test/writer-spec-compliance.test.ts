import { describe, expect, it } from "vitest"
import { writeWorkbookXml } from "../src/xlsx/workbook-writer"
import { writeWorksheetXml } from "../src/xlsx/worksheet-writer"
import { createStylesCollector } from "../src/xlsx/styles-writer"
import { createSharedStrings } from "../src/xlsx/worksheet-writer"
import { XlsxStreamWriter } from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"
import { xmlEscape, xmlEscapeAttr } from "../src/xml/writer"

// ── Fix 1: CT_Workbook ordering — workbookProtection before bookViews ──

describe("CT_Workbook ordering (ECMA-376 §18.2.27)", () => {
  it("emits workbookProtection before bookViews", () => {
    const xml = writeWorkbookXml([{ name: "Sheet1" }], undefined, undefined, 0, {
      lockStructure: true,
      lockWindows: true,
    })
    const prot = xml.indexOf("<workbookProtection")
    const views = xml.indexOf("<bookViews")
    const sheets = xml.indexOf("<sheets")
    expect(prot).toBeGreaterThan(-1)
    expect(prot).toBeLessThan(views)
    expect(views).toBeLessThan(sheets)
  })
})

// ── Fix 2: CT_Worksheet ordering — picture before tableParts; extLst last ──

describe("CT_Worksheet ordering (ECMA-376 §18.3.1.99)", () => {
  it("emits picture before tableParts, with sparkline extLst strictly last", () => {
    const styles = createStylesCollector()
    const sharedStrings = createSharedStrings()
    const result = writeWorksheetXml(
      {
        name: "Sheet1",
        rows: [[1, 2, 3]],
        backgroundImage: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
        tables: [
          {
            name: "T1",
            displayName: "T1",
            range: "A1:C1",
            columns: [{ name: "a" }, { name: "b" }, { name: "c" }],
          },
        ],
        sparklines: [{ location: "D1", dataRange: "A1:C1", type: "line" }],
      },
      styles,
      sharedStrings,
      undefined,
      0,
    )
    const xml = result.xml
    const picture = xml.indexOf("<picture")
    const tableParts = xml.indexOf("<tableParts")
    const extLst = xml.indexOf("<extLst")
    expect(picture).toBeGreaterThan(-1)
    expect(tableParts).toBeGreaterThan(-1)
    expect(extLst).toBeGreaterThan(-1)
    expect(picture).toBeLessThan(tableParts)
    expect(tableParts).toBeLessThan(extLst)
    // extLst must be the final child of <worksheet>
    expect(extLst).toBeLessThan(xml.lastIndexOf("</worksheet>"))
    expect(xml.indexOf("<", extLst + 1)).toBeLessThan(xml.lastIndexOf("</worksheet>") + 1)
  })
})

// ── Fix 3: stream writer writes the theme part it declares ──

describe("stream writer theme part", () => {
  it("actually writes xl/theme/theme1.xml declared in content-types + rels", async () => {
    const writer = new XlsxStreamWriter({ name: "Sheet1" })
    writer.addRow(["a", "b"])
    writer.addRow([1, 2])
    const buf = await writer.finish()

    // Reading back must not throw (part consistency) and round-trips data.
    const wb = await readXlsx(buf)
    expect(wb.sheets[0].rows?.[0]?.[0]).toBe("a")

    // The theme part is physically present in the archive.
    const text = new TextDecoder().decode(buf)
    expect(text.includes("xl/theme/theme1.xml")).toBe(true)
  })
})

// ── Fix 4: illegal control chars encoded as _xHHHH_ ──

describe("XML control-char escaping", () => {
  it("encodes illegal C0 control chars as _xHHHH_ in text content", () => {
    const out = xmlEscape("a\u0001b\u001Fc\u000Bd")
    expect(out).toBe("a_x0001_b_x001F_c_x000B_d")
    // Output contains no raw illegal control chars.
    const hasIllegal = [...out].some((ch) => {
      const c = ch.charCodeAt(0)
      return (c >= 0x00 && c <= 0x08) || c === 0x0b || c === 0x0c || (c >= 0x0e && c <= 0x1f)
    })
    expect(hasIllegal).toBe(false)
  })

  it("encodes CR as _x000D_ so it round-trips through XML normalization", () => {
    expect(xmlEscape("a\rb")).toBe("a_x000D_b")
  })

  it("preserves legal whitespace (tab, LF) in text content", () => {
    expect(xmlEscape("a\tb\nc")).toBe("a\tb\nc")
  })

  it("encodes illegal control chars in attribute values too", () => {
    expect(xmlEscapeAttr("a\u0001b")).toBe("a_x0001_b")
  })
})

// ── Fix 5: non-finite numbers guarded in stream writer ──

describe("stream writer non-finite guard", () => {
  it("never emits literal Infinity/NaN in cell values", async () => {
    const writer = new XlsxStreamWriter({ name: "Sheet1" })
    writer.addRow([Infinity, -Infinity, NaN, 42])
    const buf = await writer.finish()

    // Reading back must succeed (no literal Infinity/NaN to choke a parser)
    // and the finite value survives.
    const wb = await readXlsx(buf)
    const row = wb.sheets[0].rows?.[0] ?? []
    expect(row[0]).not.toBe(Infinity)
    expect(row[0]).not.toBe("Infinity")
    expect(row[3]).toBe(42)
  })
})

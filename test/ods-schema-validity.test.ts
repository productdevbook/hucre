import { describe, expect, it } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { ZipReader } from "../src/zip/reader"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// Two ways hucre wrote ODF that the OASIS grammar rejects. Both were
// invisible to every test in this repository, because every one of them
// reads the file back with hucre — and LibreOffice is lenient enough to
// open both without complaint. They were found by validating the output
// against the published RELAX NG schema with `jing`.
//
// 1. `<style:style style:family="table-cell">` has an ordered content
//    model: table-cell-properties, then paragraph-properties, then
//    text-properties. hucre wrote text first, so any cell that had both
//    a font and a fill produced a document no strict ODF consumer
//    accepts.
//
// 2. `number:min-decimal-places` does not exist in ODF 1.2 — it was
//    added in 1.3 — and hucre declares `office:version="1.2"`. #549
//    started writing it, which made every document with a number format
//    invalid against the version it claimed to be.
// ═══════════════════════════════════════════════════════════════════════

const dec = new TextDecoder()

async function contentXml(bytes: Uint8Array): Promise<string> {
  return dec.decode(await new ZipReader(bytes).extract("content.xml"))
}

const STYLE: CellStyle = {
  font: { bold: true },
  fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } },
}

describe("a cell style writes its children in the order ODF requires", () => {
  it("table-cell-properties before text-properties", async () => {
    const bytes = await writeOds({
      sheets: [
        {
          name: "S",
          rows: [["x"]],
          cells: new Map([["0,0", { value: "x", style: STYLE }]]),
        },
      ],
    })
    const xml = await contentXml(bytes)

    const cell = xml.indexOf("style:table-cell-properties")
    const text = xml.indexOf("style:text-properties")

    expect(cell, "both properties should be present").toBeGreaterThan(-1)
    expect(text, "both properties should be present").toBeGreaterThan(-1)
    expect(cell, "table-cell-properties must come first").toBeLessThan(text)
  })

  it("and the style still reads back whole", async () => {
    const bytes = await writeOds({
      sheets: [
        { name: "S", rows: [["x"]], cells: new Map([["0,0", { value: "x", style: STYLE }]]) },
      ],
    })
    const cell = (await readOds(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")

    expect(cell?.style?.font?.bold).toBe(true)
    expect(cell?.style?.fill?.type).toBe("pattern")
    if (cell?.style?.fill?.type === "pattern") {
      expect(cell.style.fill.fgColor?.rgb).toBe("FFFF00")
    }
  })
})

describe("the document declares the version it actually uses", () => {
  it("declares ODF 1.3, because it writes 1.3 attributes", async () => {
    // `number:min-decimal-places` arrived in ODF 1.3. Writing it under a
    // 1.2 declaration is what made the document invalid; the attribute is
    // worth keeping, so the declaration is what moves.
    const bytes = await writeOds({
      sheets: [
        {
          name: "S",
          rows: [[1]],
          cells: new Map([["0,0", { value: 1, style: { numFmt: "#.##" } }]]),
        },
      ],
    })
    const xml = await contentXml(bytes)

    expect(xml).toContain('office:version="1.3"')
    expect(xml).toContain("number:min-decimal-places")
  })

  it("in every part, not just content.xml", async () => {
    const bytes = await writeOds({ sheets: [{ name: "S", rows: [["x"]] }] })
    const zip = new ZipReader(bytes)

    for (const part of ["content.xml", "styles.xml", "meta.xml"]) {
      const xml = dec.decode(await zip.extract(part))
      expect(xml, part).toContain('office:version="1.3"')
    }
  })
})

import { describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import { parseStyles } from "../src/xlsx/styles"
import { ZipReader } from "../src/zip/reader"
import { readXlsx } from "../src/xlsx/reader"
import { DEFAULT_INDEXED_PALETTE } from "../src/xlsx/indexed-palette"

// ═══════════════════════════════════════════════════════════════════════
// A colour can name a palette index instead of an RGB — `indexed="2"` —
// and hucre carried the index and nothing else. A caller got
// `{ indexed: 2 }` with no way to know that means red, and a file that
// *overrode* the palette was dropped entirely: `<indexedColors>` was
// never read, so even a caller carrying its own copy of the defaults got
// the wrong answer for those files.
//
// Found by `scripts/spec-coverage.mjs` rather than by anyone noticing:
// `indexedColors` and `rgbColor` are in ECMA-376, and in the fixture
// corpus, and were nowhere in `src/`. That is the report's whole purpose
// and this is the first thing it caught.
//
// Indices 64 and 65 are the system foreground and background. They have
// no ARGB in §18.8.27 and stay unresolved here — giving them a colour the
// file did not choose would be worse than leaving them.
// ═══════════════════════════════════════════════════════════════════════

/** A stylesheet with one font colour and an optional palette override. */
function stylesXml(colorAttrs: string, palette?: string[]): string {
  const colors = palette
    ? `<colors><indexedColors>${palette
        .map((rgb) => `<rgbColor rgb="${rgb}"/>`)
        .join("")}</indexedColors></colors>`
    : ""
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<fonts count="1"><font><color ${colorAttrs}/></font></fonts>` +
    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
    '<borders count="1"><border/></borders>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>' +
    colors +
    "</styleSheet>"
  )
}

describe("an indexed colour resolves to its palette entry", () => {
  it("through the default palette", () => {
    const styles = parseStyles(stylesXml('indexed="2"'))

    expect(styles.fonts[0]!.color?.rgb).toBe("FF0000")
  })

  it("keeping the index alongside it", () => {
    // Additive: a caller reading `indexed` still finds it.
    const styles = parseStyles(stylesXml('indexed="4"'))

    expect(styles.fonts[0]!.color?.indexed).toBe(4)
    expect(styles.fonts[0]!.color?.rgb).toBe("0000FF")
  })

  it("and through a palette the file overrides", () => {
    // The case that was dropped completely. A caller with its own copy of
    // the defaults still got this wrong, because the override never
    // reached it.
    const custom = Array.from({ length: 64 }, (_, i) => (i === 2 ? "00123456" : "00000000"))
    const styles = parseStyles(stylesXml('indexed="2"', custom))

    expect(styles.fonts[0]!.color?.rgb).toBe("123456")
  })

  it("in a fill as well as a font", () => {
    const xml =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="1"><font/></fonts>' +
      '<fills count="1"><fill><patternFill patternType="solid">' +
      '<fgColor indexed="5"/><bgColor indexed="6"/>' +
      "</patternFill></fill></fills>" +
      '<borders count="1"><border/></borders>' +
      '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>' +
      "</styleSheet>"

    const fill = parseStyles(xml).fills[0]!

    expect(fill.type).toBe("pattern")
    if (fill.type === "pattern") {
      expect(fill.fgColor?.rgb).toBe("FFFF00")
      expect(fill.bgColor?.rgb).toBe("FF00FF")
    }
  })
})

describe("the indices that are not colours stay unresolved", () => {
  it("64 and 65 are the system foreground and background", () => {
    for (const indexed of [64, 65]) {
      const color = parseStyles(stylesXml(`indexed="${indexed}"`)).fonts[0]!.color

      expect(color?.indexed, `indexed ${indexed}`).toBe(indexed)
      expect(color?.rgb, `indexed ${indexed}`).toBeUndefined()
    }
  })

  it("and an index past the palette is left alone", () => {
    const color = parseStyles(stylesXml('indexed="81"')).fonts[0]!.color

    expect(color?.indexed).toBe(81)
    expect(color?.rgb).toBeUndefined()
  })

  it("an explicit rgb wins over the index", () => {
    // Both can be present. The file said the colour outright; the index
    // is the legacy spelling of it and must not overwrite.
    const color = parseStyles(stylesXml('rgb="FF00AABB" indexed="2"')).fonts[0]!.color

    expect(color?.rgb).toBe("00AABB")
  })
})

describe("the palette itself", () => {
  it("is the 64 entries §18.8.27 defines", () => {
    expect(DEFAULT_INDEXED_PALETTE).toHaveLength(64)
    // The first eight, which the spec notes are redundant of 8-15.
    expect(DEFAULT_INDEXED_PALETTE.slice(0, 8)).toEqual([
      "000000",
      "FFFFFF",
      "FF0000",
      "00FF00",
      "0000FF",
      "FFFF00",
      "FF00FF",
      "00FFFF",
    ])
    expect(DEFAULT_INDEXED_PALETTE.slice(8, 16)).toEqual(DEFAULT_INDEXED_PALETTE.slice(0, 8))
    expect(DEFAULT_INDEXED_PALETTE[63]).toBe("333333")
  })

  it("matches the one openpyxl writes, independently", async () => {
    // `openpyxl-basic.xlsx` carries the default palette written out in
    // full. Two implementations of one table: the values here came from
    // the specification text, and these came from openpyxl. If they
    // disagree, one of them is wrong and it is worth knowing which.
    const zip = new ZipReader(new Uint8Array(readFileSync("test/fixtures/openpyxl-basic.xlsx")))
    const xml = new TextDecoder().decode(await zip.extract("xl/styles.xml"))
    const block = xml.match(/<indexedColors>([\s\S]*?)<\/indexedColors>/)

    expect(block, "openpyxl-basic.xlsx should carry an indexedColors block").not.toBeNull()

    const theirs = [...block![1]!.matchAll(/rgb="([0-9A-Fa-f]{8})"/g)].map((m) =>
      m[1]!.slice(2).toUpperCase(),
    )

    expect(theirs).toHaveLength(64)
    expect(theirs).toEqual([...DEFAULT_INDEXED_PALETTE])
  })
})

describe("real files keep working", () => {
  it("excel-styled.xlsx names index 64, which stays a system colour", async () => {
    // The corpus uses 64 and 81, both outside the palette — so this fix
    // changes nothing about what those files read back as, which is the
    // point of the boundary cases above.
    const bytes = new Uint8Array(readFileSync("test/fixtures/excel-styled.xlsx"))
    const wb = await readXlsx(bytes, { readStyles: true })
    const styled = [...(wb.sheets[0]!.cells?.values() ?? [])]

    expect(styled.length).toBeGreaterThan(0)
    for (const cell of styled) {
      const color = cell.style?.font?.color
      if (color?.indexed === 64) expect(color.rgb).toBeUndefined()
    }
  })
})

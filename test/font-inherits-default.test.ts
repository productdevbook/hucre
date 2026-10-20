import { describe, expect, it } from "vitest"
import { createStylesCollector } from "../src/xlsx/styles-writer"
import { writeXlsx } from "../src/xlsx/writer"
import { writeXlsxStream } from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #594 — `defaultFont` is `fonts[0]` in styles.xml, and its doc says every
// cell format inherits from it. A cell font that gave only part of a font,
// `{ bold: true }` say, was written as its own `<font>` with only that part:
//
//   <font><sz val="10"/><name val="Arial"/></font>
//   <font><b/><color rgb="FF1F3A5F"/></font>
//
// `fontId="1"` does not layer on font 0, so the bold cell had no face or
// size at all and LibreOffice drew it in Cambria 11pt. A cell format's font
// now takes `name` and `size` from the default when it leaves them out.
// Differential formats (conditional formatting) keep only what they say:
// there a missing part means "leave the cell's own font alone".
// ═══════════════════════════════════════════════════════════════════════

const dec = new TextDecoder()

async function stylesXml(bytes: Uint8Array): Promise<string> {
  return dec.decode(await new ZipReader(bytes).extract("xl/styles.xml"))
}

function fontsOf(xml: string): string[] {
  const fonts = xml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/)?.[1] ?? ""
  return fonts.match(/<font>[\s\S]*?<\/font>/g) ?? []
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  let total = 0
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
    total += value.length
  }
  const out = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    out.set(chunk, offset)
    offset += chunk.length
  }
  return out
}

describe("a cell font inherits name and size from defaultFont (#594)", () => {
  it("fills in the default's name and size for a bold-only font", async () => {
    const bytes = await writeXlsx({
      defaultFont: { name: "Arial", size: 10 },
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [{ value: "Bold header", style: { font: { bold: true, color: { rgb: "FF1F3A5F" } } } }],
          ],
        },
      ],
    })

    expect(fontsOf(await stylesXml(bytes))).toEqual([
      '<font><sz val="10"/><name val="Arial"/></font>',
      '<font><b/><color rgb="FF1F3A5F"/><sz val="10"/><name val="Arial"/></font>',
    ])
  })

  it("keeps a name and size the cell font gives itself", () => {
    const styles = createStylesCollector({ name: "Arial", size: 10 })
    styles.addStyle({ font: { name: "Georgia", size: 14, italic: true } })
    styles.addStyle({ font: { name: "Georgia" } })
    styles.addStyle({ font: { size: 14 } })

    expect(fontsOf(styles.toXml()).slice(1)).toEqual([
      '<font><i/><sz val="14"/><name val="Georgia"/></font>',
      '<font><sz val="10"/><name val="Georgia"/></font>',
      '<font><sz val="14"/><name val="Arial"/></font>',
    ])
  })

  it("falls back to Calibri 11 when there is no defaultFont", () => {
    const styles = createStylesCollector()
    styles.addStyle({ font: { bold: true } })

    expect(fontsOf(styles.toXml())[1]).toBe('<font><b/><sz val="11"/><name val="Calibri"/></font>')
  })

  it("takes scheme and family only along with the default's name", () => {
    // A theme default — what `readXlsx` hands back as `defaultFont` for a
    // file Excel wrote. `scheme="minor"` tells Excel to draw the theme's
    // body font rather than `name`, so copying it onto a font that names
    // its own typeface would replace that typeface.
    const styles = createStylesCollector({ name: "Calibri", size: 11, family: 2, scheme: "minor" })
    styles.addStyle({ font: { bold: true } })
    styles.addStyle({ font: { name: "Georgia", bold: true } })

    const [, inherited, named] = fontsOf(styles.toXml())
    expect(inherited).toBe(
      '<font><b/><family val="2"/><scheme val="minor"/><sz val="11"/><name val="Calibri"/></font>',
    )
    expect(named).toBe('<font><b/><sz val="11"/><name val="Georgia"/></font>')
  })

  it("does not take the default's colour", () => {
    const styles = createStylesCollector({ name: "Arial", size: 10, color: { theme: 1 } })
    styles.addStyle({ font: { bold: true } })

    expect(fontsOf(styles.toXml())[1]).toBe('<font><b/><sz val="10"/><name val="Arial"/></font>')
  })

  it("leaves a conditional formatting font as given", async () => {
    const bytes = await writeXlsx({
      defaultFont: { name: "Arial", size: 10 },
      sheets: [
        {
          name: "S",
          rows: [[1], [2]],
          conditionalRules: [
            {
              type: "cellIs",
              priority: 1,
              operator: "greaterThan",
              formula: "1",
              range: "A1:A2",
              style: { font: { bold: true, color: { rgb: "FF9C0006" } } },
            },
          ],
        },
      ],
    })

    const dxfs = (await stylesXml(bytes)).match(/<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/)?.[1]
    expect(dxfs).toBe('<dxf><font><b/><color rgb="FF9C0006"/></font></dxf>')
  })

  it("does not mutate the caller's style object", async () => {
    const style: CellStyle = { font: { bold: true } }
    await writeXlsx({
      defaultFont: { name: "Arial", size: 10 },
      sheets: [{ name: "S", rows: [[{ value: "x", style }]] }],
    })

    expect(style).toEqual({ font: { bold: true } })
  })

  it("applies to the streaming writer too", async () => {
    function* rows() {
      yield [{ value: "x", style: { font: { bold: true } } }]
    }
    const bytes = await collect(writeXlsxStream(rows(), { name: "S" }))

    expect(fontsOf(await stylesXml(bytes))[1]).toBe(
      '<font><b/><sz val="11"/><name val="Calibri"/></font>',
    )
  })

  it("reads back with the inherited name and size", async () => {
    const bytes = await writeXlsx({
      defaultFont: { name: "Arial", size: 10 },
      sheets: [{ name: "S", rows: [[{ value: "x", style: { font: { bold: true } } }]] }],
    })

    const cell = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")
    expect(cell?.style?.font).toEqual({ bold: true, size: 10, name: "Arial" })
  })
})

import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { parseStyles } from "../src/xlsx/styles"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"

// Excel writes every conditional-format fill as a <patternFill> with a
// colour and no patternType, and renders it solid. Reading the absence as
// the cell-fill default (none) made saveXlsx write patternType="none", so
// the rule kept its condition and lost its colour.

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const EXCEL_DXF = '<dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'

async function part(bytes: Uint8Array, path: string): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(bytes).extract(path))
}

async function rewrite(
  bytes: Uint8Array,
  path: string,
  edit: (xml: string) => string,
): Promise<Uint8Array> {
  const all = await new ZipReader(bytes).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    zw.add(
      name,
      name === path ? new TextEncoder().encode(edit(new TextDecoder().decode(data))) : data,
    )
  }
  return zw.build()
}

async function excelStyleRule(): Promise<Uint8Array> {
  const base = await writeXlsx({
    sheets: [
      {
        name: "S",
        rows: [[5]],
        conditionalRules: [
          {
            range: "A1",
            type: "cellIs",
            operator: "greaterThan",
            formula: "1",
            priority: 1,
            style: { font: { bold: true } },
          },
        ],
      },
    ],
  })
  return rewrite(base, "xl/styles.xml", (x) =>
    x.replace(/<dxfs[^>]*>[\s\S]*<\/dxfs>/, `<dxfs count="1">${EXCEL_DXF}</dxfs>`),
  )
}

describe("dxf fill without patternType", () => {
  it("reads as solid in a dxf", () => {
    const styles = parseStyles(
      `<styleSheet xmlns="${MAIN_NS}"><dxfs count="1">${EXCEL_DXF}</dxfs></styleSheet>`,
    )
    expect(styles.dxfs[0]!.fill).toEqual({
      type: "pattern",
      pattern: "solid",
      bgColor: { rgb: "FFC7CE" },
    })
  })

  it("still reads as none in a cell fill, the schema default", () => {
    const styles = parseStyles(
      `<styleSheet xmlns="${MAIN_NS}"><fills count="1"><fill><patternFill><fgColor rgb="FFFF0000"/></patternFill></fill></fills></styleSheet>`,
    )
    expect(styles.fills[0]).toMatchObject({ type: "pattern", pattern: "none" })
  })

  it("keeps the rule's colour through openXlsx -> saveXlsx", async () => {
    const out = await saveXlsx(await openXlsx(await excelStyleRule(), { readStyles: true }))
    const dxfs = (await part(out, "xl/styles.xml")).match(/<dxfs[^>]*>[\s\S]*<\/dxfs>/)![0]
    expect(dxfs).toContain('patternType="solid"')
    expect(dxfs).toContain('rgb="FFFFC7CE"')
    expect(dxfs).not.toContain('patternType="none"')
  })
})

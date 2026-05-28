import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"

const dec = new TextDecoder()

async function unzipText(buf: Uint8Array, path: string): Promise<string | null> {
  const zip = new ZipReader(buf)
  if (!zip.has(path)) return null
  return dec.decode(await zip.extract(path))
}

async function unzipEntries(buf: Uint8Array): Promise<string[]> {
  const zip = new ZipReader(buf)
  return zip.entries()
}

describe("Excel 2024 checkbox cells (#157)", () => {
  it("emits the FeaturePropertyBag part when at least one cell is a checkbox", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Done?"], [true], [false]],
          cells: new Map([
            ["1,0", { value: true, type: "boolean", checkbox: true }],
            ["2,0", { value: false, type: "boolean", checkbox: true }],
          ]),
        },
      ],
    })

    const entries = await unzipEntries(buf)
    expect(entries).toContain("xl/featurePropertyBag/featurePropertyBag.xml")

    const bag = await unzipText(buf, "xl/featurePropertyBag/featurePropertyBag.xml")
    expect(bag).toContain('<bag type="Checkbox"/>')
    expect(bag).toContain('<bag type="XFControls">')
    expect(bag).toContain('<bag type="XFComplement">')
    expect(bag).toContain('extRef="XFComplementsMapperExtRef"')
  })

  it("registers Override + Relationship for the FeaturePropertyBag part", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [[true]],
          cells: new Map([["0,0", { value: true, type: "boolean", checkbox: true }]]),
        },
      ],
    })

    const ct = (await unzipText(buf, "[Content_Types].xml"))!
    expect(ct).toContain("/xl/featurePropertyBag/featurePropertyBag.xml")
    expect(ct).toContain("application/vnd.ms-excel.featurepropertybag+xml")

    const rels = (await unzipText(buf, "xl/_rels/workbook.xml.rels"))!
    expect(rels).toContain("featurePropertyBag/featurePropertyBag.xml")
    expect(rels).toContain("schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag")
  })

  it("emits an xfComplement extension on the checkbox xf in styles.xml", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [[true]],
          cells: new Map([["0,0", { value: true, type: "boolean", checkbox: true }]]),
        },
      ],
    })

    const styles = (await unzipText(buf, "xl/styles.xml"))!
    expect(styles).toContain("{C7286773-470A-42A8-94C5-96B5CB345126}")
    expect(styles).toContain('<xfpb:xfComplement i="0"/>')
    expect(styles).toContain("schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag")
  })

  it("does NOT emit FeaturePropertyBag when no checkbox cells exist", async () => {
    const buf = await writeXlsx({
      sheets: [{ name: "Plain", rows: [[true], [false]] }],
    })

    const entries = await unzipEntries(buf)
    expect(entries).not.toContain("xl/featurePropertyBag/featurePropertyBag.xml")

    const ct = (await unzipText(buf, "[Content_Types].xml"))!
    expect(ct).not.toContain("featurepropertybag")

    const styles = (await unzipText(buf, "xl/styles.xml"))!
    expect(styles).not.toContain("xfComplement")
  })

  it("round-trips the checkbox flag through readXlsx", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [[true], [false]],
          cells: new Map([
            ["0,0", { value: true, type: "boolean", checkbox: true }],
            ["1,0", { value: false, type: "boolean", checkbox: true }],
          ]),
        },
      ],
    })

    const wb = await readXlsx(buf)
    const sheet = wb.sheets[0]!
    expect(sheet.rows[0]![0]).toBe(true)
    expect(sheet.rows[1]![0]).toBe(false)

    expect(sheet.cells?.get("0,0")?.checkbox).toBe(true)
    expect(sheet.cells?.get("1,0")?.checkbox).toBe(true)
  })

  it("preserves checkbox alongside an explicit cell style", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "Styled",
          rows: [[true]],
          cells: new Map([
            [
              "0,0",
              {
                value: true,
                type: "boolean",
                checkbox: true,
                style: { font: { bold: true, color: { rgb: "FF0000" } } },
              },
            ],
          ]),
        },
      ],
    })

    const wb = await readXlsx(buf, { readStyles: true })
    const cell = wb.sheets[0]!.cells?.get("0,0")
    expect(cell?.checkbox).toBe(true)
    expect(cell?.style?.font?.bold).toBe(true)
  })

  it("multiple checkbox cells share a single xf complement entry", async () => {
    // XlsxWriter emits exactly one <bag type="XFComplements"> with one
    // <bagId>2</bagId> regardless of how many checkbox cells there are.
    const cells = new Map<string, { value: boolean; type: "boolean"; checkbox: true }>()
    for (let r = 0; r < 8; r++) {
      cells.set(`${r},0`, { value: r % 2 === 0, type: "boolean", checkbox: true })
    }

    const buf = await writeXlsx({
      sheets: [
        {
          name: "Many",
          rows: Array.from({ length: 8 }, (_, i) => [i % 2 === 0]),
          cells,
        },
      ],
    })

    const styles = (await unzipText(buf, "xl/styles.xml"))!
    // Exactly one xfComplement, all checkbox cells share it.
    const matches = styles.match(/xfComplement i="0"/g)
    expect(matches?.length).toBe(1)
  })
})

import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { readXlsx } from "../src/xlsx/reader"
import { MAX_COL_INDEX } from "../src/limits"

const enc = new TextEncoder()
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

/**
 * Smallest workbook that carries a `<cols>` block — the payload under
 * test is the `<col>` element, everything else is scaffolding.
 */
async function workbookWithCols(colsXml: string): Promise<Uint8Array> {
  const sheet =
    `<?xml version="1.0"?><worksheet ${NS} ${R}>` +
    `<cols>${colsXml}</cols>` +
    `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row></sheetData>` +
    `</worksheet>`

  const zip = new ZipWriter()
  zip.add(
    "[Content_Types].xml",
    enc.encode(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
        `</Types>`,
    ),
  )
  zip.add(
    "_rels/.rels",
    enc.encode(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
  )
  zip.add(
    "xl/workbook.xml",
    enc.encode(
      `<?xml version="1.0"?><workbook ${NS} ${R}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
  )
  zip.add(
    "xl/_rels/workbook.xml.rels",
    enc.encode(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
  )
  zip.add("xl/worksheets/sheet1.xml", enc.encode(sheet))

  return zip.build()
}

// ═══════════════════════════════════════════════════════════════════════
// Regression guard for #355 — an unbounded <col> range hung readXlsx
// forever on a 1.4 KB file. These must all complete, not just pass.
// ═══════════════════════════════════════════════════════════════════════

describe("<col> range bounds", () => {
  it("does not hang on max=1e999 (Infinity)", async () => {
    const buf = await workbookWithCols('<col min="1" max="1e999" width="10"/>')
    const workbook = await readXlsx(buf)
    // Data survives; the range is simply clamped.
    expect(workbook.sheets[0].rows).toEqual([["x"]])
    expect(workbook.sheets[0].columns!.length).toBeLessThanOrEqual(MAX_COL_INDEX + 1)
  }, 10_000)

  it("clamps a finite range past Excel's last column", async () => {
    const buf = await workbookWithCols('<col min="1" max="99999999" width="10"/>')
    const workbook = await readXlsx(buf)
    expect(workbook.sheets[0].columns!.length).toBe(MAX_COL_INDEX + 1)
  }, 10_000)

  it("handles a NaN max without looping", async () => {
    const buf = await workbookWithCols('<col min="1" max="abc" width="10"/>')
    const workbook = await readXlsx(buf)
    expect(workbook.sheets[0].rows).toEqual([["x"]])
  }, 10_000)

  it("handles a negative or zero range", async () => {
    for (const attrs of ['min="-5" max="-1"', 'min="0" max="0"', 'min="5" max="1"']) {
      const workbook = await readXlsx(await workbookWithCols(`<col ${attrs} width="10"/>`))
      expect(workbook.sheets[0].rows).toEqual([["x"]])
    }
  }, 10_000)

  it("still reads an ordinary column range", async () => {
    const buf = await workbookWithCols('<col min="2" max="4" width="25" hidden="1"/>')
    const workbook = await readXlsx(buf)
    const columns = workbook.sheets[0].columns!

    // 1-based min/max → 0-based indices 1..3.
    expect(columns[0]).toEqual({})
    for (const idx of [1, 2, 3]) {
      expect(columns[idx]).toMatchObject({ width: 25, hidden: true })
    }
  })

  it("truncates a fractional bound rather than rejecting it", async () => {
    const buf = await workbookWithCols('<col min="1.7" max="3.9" width="12"/>')
    const workbook = await readXlsx(buf)
    const columns = workbook.sheets[0].columns!
    // min 1.7 → 1, max 3.9 → 3, so 0-based 0..2 carry the width.
    expect(columns[0]).toMatchObject({ width: 12 })
    expect(columns[2]).toMatchObject({ width: 12 })
  })

  it("survives many oversized <col> elements without unbounded growth", async () => {
    const cols = Array.from({ length: 50 }, () => '<col min="1" max="1e999" width="10"/>').join("")
    const workbook = await readXlsx(await workbookWithCols(cols))
    // The shared array stays capped no matter how many elements ask for more.
    expect(workbook.sheets[0].columns!.length).toBeLessThanOrEqual(MAX_COL_INDEX + 1)
  }, 20_000)
})

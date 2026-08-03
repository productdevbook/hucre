import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { readOds } from "../src/ods/reader"
import { streamOdsRows } from "../src/ods/stream"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import { fromHtml } from "../src/export/html-import"
import { ParseError } from "../src/errors"
import { MAX_COL_INDEX, MAX_SPAN_CELLS } from "../src/limits"
import type { CellValue } from "../src/_types"

const enc = new TextEncoder()
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

// ── Fixtures ─────────────────────────────────────────────────────────

async function odsWithContent(bodyXml: string, stylesXml?: string): Promise<Uint8Array> {
  const zip = new ZipWriter()
  zip.add("mimetype", enc.encode("application/vnd.oasis.opendocument.spreadsheet"), {
    compress: false,
  })
  zip.add(
    "META-INF/manifest.xml",
    enc.encode(
      `<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">` +
        `<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>` +
        `</manifest:manifest>`,
    ),
  )
  zip.add(
    "content.xml",
    enc.encode(
      `<?xml version="1.0"?><office:document-content ` +
        `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
        `xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ` +
        `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">` +
        `<office:body><office:spreadsheet><table:table table:name="S">${bodyXml}` +
        `</table:table></office:spreadsheet></office:body></office:document-content>`,
    ),
  )
  if (stylesXml) zip.add("styles.xml", enc.encode(stylesXml))
  return zip.build()
}

async function xlsxWithSheetData(sheetData: string): Promise<Uint8Array> {
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
  zip.add(
    "xl/worksheets/sheet1.xml",
    enc.encode(
      `<?xml version="1.0"?><worksheet ${NS} ${R}><sheetData>${sheetData}</sheetData></worksheet>`,
    ),
  )
  return zip.build()
}

// ═══════════════════════════════════════════════════════════════════════
// #363 — values read from a file driving an allocation, a loop bound, or
// a String.repeat with no cap. Every test asserts *completion* rather
// than a value, because the failure mode is a hang or a raw RangeError.
// ═══════════════════════════════════════════════════════════════════════

describe("ODS <text:s> repeat count", () => {
  it("does not allocate a gigabyte for one cell", async () => {
    // Uncapped this reaches a raw RangeError: Invalid string length.
    const body =
      `<table:table-row><table:table-cell><text:p>` +
      `<text:s text:c="900000000"/>x</text:p></table:table-cell></table:table-row>`

    const workbook = await readOds(await odsWithContent(body))
    const value = workbook.sheets[0].rows[0][0] as string
    expect(typeof value).toBe("string")
    expect(value.length).toBeLessThan(1_000_000)
  }, 20_000)

  it("still expands an ordinary run of spaces", async () => {
    const body =
      `<table:table-row><table:table-cell><text:p>` +
      `a<text:s text:c="3"/>b</text:p></table:table-cell></table:table-row>`
    const workbook = await readOds(await odsWithContent(body))
    expect(workbook.sheets[0].rows[0][0]).toBe("a   b")
  })

  it("treats a missing or malformed count as one space", async () => {
    for (const attr of ["", ' text:c="0"', ' text:c="abc"', ' text:c="-5"']) {
      const body =
        `<table:table-row><table:table-cell><text:p>` +
        `a<text:s${attr}/>b</text:p></table:table-cell></table:table-row>`
      const workbook = await readOds(await odsWithContent(body))
      expect(workbook.sheets[0].rows[0][0], attr).toBe("a b")
    }
  })

  it("is capped in the streaming reader too", async () => {
    const body =
      `<table:table-row><table:table-cell><text:p>` +
      `<text:s text:c="900000000"/>x</text:p></table:table-cell></table:table-row>`

    const rows: CellValue[][] = []
    for await (const row of streamOdsRows(await odsWithContent(body))) {
      rows.push(row.values)
    }
    expect((rows[0][0] as string).length).toBeLessThan(1_000_000)
  }, 20_000)
})

describe("ODS number:decimal-places", () => {
  it("does not build a two-gigabyte format string", async () => {
    // Reached while parsing styles, before any cell data exists.
    const styles =
      `<?xml version="1.0"?><office:document-styles ` +
      `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
      `xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ` +
      `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">` +
      `<office:styles><number:number-style style:name="N1">` +
      `<number:number number:decimal-places="2000000000"/>` +
      `</number:number-style></office:styles></office:document-styles>`

    const body = `<table:table-row><table:table-cell><text:p>1</text:p></table:table-cell></table:table-row>`
    const workbook = await readOds(await odsWithContent(body, styles), { readStyles: true })
    expect(workbook.sheets[0].rows[0][0]).toBeDefined()
  }, 20_000)
})

describe("streaming XLSX column bounds", () => {
  it("rejects a cell reference past the last column", async () => {
    // The batch reader has always bound-checked this; the streaming one
    // allocated a 12-million-element array per row instead.
    const buf = await xlsxWithSheetData(
      `<row r="1"><c r="AAAAAA1" t="inlineStr"><is><t>x</t></is></c></row>`,
    )

    const drain = async () => {
      for await (const _ of streamXlsxRows(buf)) {
        // consume
      }
    }
    await expect(drain()).rejects.toThrow(ParseError)
    await expect(drain()).rejects.toThrow(/outside the supported sheet bounds/)
  }, 20_000)

  it("throws a typed error rather than a raw RangeError", async () => {
    // A longer reference used to overflow past Array's own limit.
    const buf = await xlsxWithSheetData(
      `<row r="1"><c r="AAAAAAAAAA1" t="inlineStr"><is><t>x</t></is></c></row>`,
    )
    const drain = async () => {
      for await (const _ of streamXlsxRows(buf)) {
        // consume
      }
    }
    await expect(drain()).rejects.toThrow(ParseError)
  }, 20_000)

  it("still streams a row using the last legal column", async () => {
    const buf = await xlsxWithSheetData(
      `<row r="1"><c r="XFD1" t="inlineStr"><is><t>x</t></is></c></row>`,
    )
    const rows: CellValue[][] = []
    for await (const row of streamXlsxRows(buf)) rows.push(row.values)
    expect(rows[0]).toHaveLength(MAX_COL_INDEX + 1)
    expect(rows[0][MAX_COL_INDEX]).toBe("x")
  })
})

describe("fromHtml span bounds", () => {
  it("does not hang on a hostile rowspan x colspan", () => {
    // 100000 x 100000 is 1e10 Set insertions. Clamping each span alone
    // still leaves 16384 x 1048576 — the product is what matters.
    const started = Date.now()
    const sheet = fromHtml('<table><tr><td rowspan="100000" colspan="100000">x</td></tr></table>')
    expect(Date.now() - started).toBeLessThan(10_000)

    const merge = sheet.merges![0]
    const cells = (merge.endRow - merge.startRow + 1) * (merge.endCol - merge.startCol + 1)
    expect(cells).toBeLessThanOrEqual(MAX_SPAN_CELLS)
  }, 20_000)

  it("leaves ordinary spans untouched", () => {
    const sheet = fromHtml(
      '<table><tr><td rowspan="2" colspan="3">x</td><td>y</td></tr><tr><td>z</td></tr></table>',
    )
    expect(sheet.merges![0]).toMatchObject({
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 2,
    })
  })

  it("treats malformed spans as 1", () => {
    const sheet = fromHtml('<table><tr><td rowspan="abc" colspan="-4">x</td></tr></table>')
    expect(sheet.merges).toBeUndefined()
    expect(sheet.rows[0]).toEqual(["x"])
  })
})

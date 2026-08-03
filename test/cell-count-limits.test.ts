import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { readXlsx } from "../src/xlsx/reader"
import { readOds } from "../src/ods/reader"
import { ParseError } from "../src/errors"
import { MAX_TOTAL_CELLS } from "../src/limits"

const enc = new TextEncoder()
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

// ── Fixtures ─────────────────────────────────────────────────────────

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

async function odsWithRows(rowsXml: string): Promise<Uint8Array> {
  const zip = new ZipWriter()
  zip.add("mimetype", enc.encode("application/vnd.oasis.opendocument.spreadsheet"), {
    compress: false,
  })
  zip.add(
    "META-INF/manifest.xml",
    enc.encode(
      `<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">` +
        `<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>` +
        `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`,
    ),
  )
  zip.add(
    "content.xml",
    enc.encode(
      `<?xml version="1.0"?><office:document-content ` +
        `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">` +
        `<office:body><office:spreadsheet><table:table table:name="S">${rowsXml}` +
        `</table:table></office:spreadsheet></office:body></office:document-content>`,
    ),
  )
  return zip.build()
}

// ═══════════════════════════════════════════════════════════════════════
// Regression guard for #363. A reader hands back `rows: CellValue[][]`,
// a dense rectangle — so a sheet costs its bounding box, not its cell
// count. Bounding each coordinate separately is not enough, and the gap
// was a fatal V8 OOM (SIGABRT, uncatchable) from a ~1.5 KB file.
//
// Every test here asserts a *catchable* ParseError, not a value.
// ═══════════════════════════════════════════════════════════════════════

describe("XLSX bounding-box limit", () => {
  it("rejects two legal cells at opposite corners instead of dying", async () => {
    // Both A1 and XFD1048576 are inside Excel's own limits, so the
    // per-coordinate check passes them. Their product is 1.7e10 slots.
    const buf = await xlsxWithSheetData(
      `<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c></row>` +
        `<row r="1048576"><c r="XFD1048576" t="inlineStr"><is><t>z</t></is></c></row>`,
    )

    await expect(readXlsx(buf)).rejects.toThrow(ParseError)
    await expect(readXlsx(buf)).rejects.toThrow(/over the .* limit/)
  }, 20_000)

  it("names the sheet and the actual dimensions", async () => {
    const buf = await xlsxWithSheetData(
      `<row r="1048576"><c r="XFD1048576" t="inlineStr"><is><t>z</t></is></c></row>`,
    )
    await expect(readXlsx(buf)).rejects.toThrow(/Sheet "S" spans 1048576 rows x 16384 columns/)
  }, 20_000)

  it("points at the options that bound the read", async () => {
    const buf = await xlsxWithSheetData(
      `<row r="1048576"><c r="XFD1048576" t="inlineStr"><is><t>z</t></is></c></row>`,
    )
    await expect(readXlsx(buf)).rejects.toThrow(/`range` or `maxRows`/)
  }, 20_000)

  it("still reads an ordinary sparse sheet", async () => {
    const buf = await xlsxWithSheetData(
      `<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c></row>` +
        `<row r="500"><c r="J500" t="inlineStr"><is><t>z</t></is></c></row>`,
    )
    const workbook = await readXlsx(buf)
    expect(workbook.sheets[0].rows).toHaveLength(500)
    expect(workbook.sheets[0].rows[0][0]).toBe("a")
    expect(workbook.sheets[0].rows[499][9]).toBe("z")
  })

  it("admits a wide-but-shallow sheet under the cap", async () => {
    // 1 row x 16,384 columns is 16,384 cells — nowhere near the limit,
    // even though it uses the very last column.
    const buf = await xlsxWithSheetData(
      `<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c>` +
        `<c r="XFD1" t="inlineStr"><is><t>z</t></is></c></row>`,
    )
    const workbook = await readXlsx(buf)
    expect(workbook.sheets[0].rows[0]).toHaveLength(16384)
  })
})

describe("ODS bounding-box limit", () => {
  it("rejects a wide row repeated across the whole sheet", async () => {
    // Each repeat attribute is individually capped; the product is what
    // was unbounded.
    const rows =
      `<table:table-row table:number-rows-repeated="1048576">` +
      `<table:table-cell table:number-columns-repeated="16384">` +
      `<text:p>x</text:p></table:table-cell></table:table-row>`

    await expect(readOds(await odsWithRows(rows))).rejects.toThrow(ParseError)
    await expect(readOds(await odsWithRows(rows))).rejects.toThrow(/over the .* limit/)
  }, 20_000)

  it("still reads an ordinary repeated row", async () => {
    const rows =
      `<table:table-row table:number-rows-repeated="3">` +
      `<table:table-cell><text:p>x</text:p></table:table-cell></table:table-row>`
    const workbook = await readOds(await odsWithRows(rows))
    expect(workbook.sheets[0].rows).toHaveLength(3)
    expect(workbook.sheets[0].rows[0][0]).toBe("x")
  })
})

describe("MAX_TOTAL_CELLS", () => {
  it("leaves room for real spreadsheets", () => {
    // Excel's row limit at 16 columns is ~16.8M cells; the cap has to sit
    // above realistic files or it becomes the bug.
    expect(MAX_TOTAL_CELLS).toBeGreaterThanOrEqual(16_777_216)
  })
})

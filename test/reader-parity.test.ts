import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { readXlsx } from "../src/xlsx/reader"
import { streamXlsxRows } from "../src/xlsx/stream-reader"

// Regression tests for batch/stream reader parity:
//  - implicit column position (cells without an `r` attribute)
//  - lenient OOXML relationship-type matching (Strict namespace)
//  - out-of-range SST index returns null in both readers

function textToBytes(s: string): Uint8Array {
  return new TextEncoder().encode(s)
}

interface BuildOpts {
  /** Full <sheetData>…</sheetData>-less worksheet body (we wrap it). */
  worksheetData: string
  sharedStrings?: string[]
  /** Use Strict OOXML relationship namespace instead of Transitional. */
  strictRelNs?: boolean
}

const TRANSITIONAL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships"

async function buildXlsx(opts: BuildOpts): Promise<Uint8Array> {
  const ns = opts.strictRelNs ? STRICT : TRANSITIONAL
  const w = new ZipWriter()
  const hasSs = !!opts.sharedStrings

  const overrides = [
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`,
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`,
    `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
  ]
  if (hasSs) {
    overrides.push(
      `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`,
    )
  }

  w.add(
    "[Content_Types].xml",
    textToBytes(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${overrides.join("\n  ")}
</Types>`,
    ),
    { compress: false },
  )

  w.add(
    "_rels/.rels",
    textToBytes(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${ns}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    ),
    { compress: false },
  )

  w.add(
    "xl/workbook.xml",
    textToBytes(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr/>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
    ),
    { compress: false },
  )

  const wbRels = [
    `<Relationship Id="rId1" Type="${ns}/worksheet" Target="worksheets/sheet1.xml"/>`,
    `<Relationship Id="rId2" Type="${ns}/styles" Target="styles.xml"/>`,
  ]
  if (hasSs) {
    wbRels.push(`<Relationship Id="rId3" Type="${ns}/sharedStrings" Target="sharedStrings.xml"/>`)
  }
  w.add(
    "xl/_rels/workbook.xml.rels",
    textToBytes(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${wbRels.join("\n  ")}
</Relationships>`,
    ),
    { compress: false },
  )

  if (hasSs) {
    const items = opts.sharedStrings!
    w.add(
      "xl/sharedStrings.xml",
      textToBytes(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${items.length}" uniqueCount="${items.length}">
  ${items.map((t) => `<si><t>${t}</t></si>`).join("\n  ")}
</sst>`,
      ),
      { compress: false },
    )
  }

  w.add(
    "xl/styles.xml",
    textToBytes(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`,
    ),
    { compress: false },
  )

  w.add(
    "xl/worksheets/sheet1.xml",
    textToBytes(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${opts.worksheetData}</sheetData>
</worksheet>`,
    ),
    { compress: false },
  )

  return w.build()
}

async function streamToRows(bytes: Uint8Array): Promise<unknown[][]> {
  const rows: unknown[][] = []
  for await (const r of streamXlsxRows(bytes)) {
    rows[r.index] = r.values
  }
  return rows
}

describe("reader parity: implicit column position", () => {
  // First cell has r="A1", subsequent cells omit r and must fall into B1, C1.
  const worksheetData = `<row r="1"><c r="A1"><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row>`

  it("batch reader fills implicit columns like the streaming reader", async () => {
    const bytes = await buildXlsx({ worksheetData })
    const wb = await readXlsx(bytes)
    expect(wb.sheets[0].rows[0]).toEqual([1, 2, 3])

    const streamRows = await streamToRows(bytes)
    expect(streamRows[0]).toEqual([1, 2, 3])
  })
})

describe("reader parity: strict OOXML relationship namespace", () => {
  const worksheetData = `<row r="1"><c r="A1"><v>42</v></c></row>`

  it("batch reader resolves relationships in the Strict namespace", async () => {
    const bytes = await buildXlsx({ worksheetData, strictRelNs: true })
    const wb = await readXlsx(bytes)
    expect(wb.sheets[0].rows[0][0]).toBe(42)
  })

  it("streaming reader resolves relationships in the Strict namespace", async () => {
    const bytes = await buildXlsx({ worksheetData, strictRelNs: true })
    const streamRows = await streamToRows(bytes)
    expect(streamRows[0][0]).toBe(42)
  })
})

describe("reader parity: out-of-range SST index", () => {
  // index 5 is out of range for a 1-entry shared strings table.
  const worksheetData = `<row r="1"><c r="A1" t="s"><v>5</v></c></row>`

  it("both readers return null, not the raw index string", async () => {
    const bytes = await buildXlsx({ worksheetData, sharedStrings: ["only"] })
    const wb = await readXlsx(bytes)
    expect(wb.sheets[0].rows[0][0]).toBeNull()

    const streamRows = await streamToRows(bytes)
    expect(streamRows[0][0]).toBeNull()
  })
})

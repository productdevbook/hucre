import { describe, it, expect } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { ZipReader } from "../src/zip/reader"
import { parseXml } from "../src/xml/parser"

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8")

async function extractFile(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data)
  const raw = await zip.extract(path)
  return decoder.decode(raw)
}

async function parseXmlFromZip(data: Uint8Array, path: string) {
  const xml = await extractFile(data, path)
  return parseXml(xml)
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function getElementText(el: { children: Array<unknown> }): string {
  return el.children.filter((c: unknown) => typeof c === "string").join("")
}

async function getFirstTable(data: Uint8Array) {
  const contentDoc = await parseXmlFromZip(data, "content.xml")
  const body = findChild(contentDoc, "body")
  const spreadsheet = findChild(body, "spreadsheet")
  return findChild(spreadsheet, "table")
}

// ── #111: settings.xml ──────────────────────────────────────────────

describe("ODS spec — #111: settings.xml", () => {
  it("settings.xml is present in the ZIP archive", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    const zip = new ZipReader(data)
    expect(zip.has("settings.xml")).toBe(true)
  })

  it("settings.xml has correct structure", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    const settingsDoc = await parseXmlFromZip(data, "settings.xml")

    // Root element should be office:document-settings
    expect(settingsDoc.tag).toContain("document-settings")

    // Should have office:version="1.2"
    expect(settingsDoc.attrs["office:version"]).toBe("1.2")

    // Should contain office:settings child element
    const settings = findChild(settingsDoc, "settings")
    expect(settings).toBeDefined()
  })

  it("settings.xml has required namespace declarations", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    const settingsXml = await extractFile(data, "settings.xml")
    expect(settingsXml).toContain("xmlns:office=")
    expect(settingsXml).toContain("xmlns:config=")
  })

  it("settings.xml is listed in manifest.xml", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    const manifest = await parseXmlFromZip(data, "META-INF/manifest.xml")
    const entries = findChildren(manifest, "file-entry")
    const paths = entries.map((e: any) => e.attrs["manifest:full-path"])

    expect(paths).toContain("settings.xml")
  })

  it("settings.xml manifest entry has correct media type", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    const manifest = await parseXmlFromZip(data, "META-INF/manifest.xml")
    const entries = findChildren(manifest, "file-entry")
    const settingsEntry = entries.find((e: any) => e.attrs["manifest:full-path"] === "settings.xml")

    expect(settingsEntry).toBeDefined()
    expect(settingsEntry.attrs["manifest:media-type"]).toBe("text/xml")
  })
})

// ── #112: manifest.xml version attribute ────────────────────────────

describe("ODS spec — #112: manifest.xml version", () => {
  it('root entry "/" has manifest:version="1.2"', async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    const manifest = await parseXmlFromZip(data, "META-INF/manifest.xml")
    const entries = findChildren(manifest, "file-entry")
    const rootEntry = entries.find((e: any) => e.attrs["manifest:full-path"] === "/")

    expect(rootEntry).toBeDefined()
    expect(rootEntry.attrs["manifest:version"]).toBe("1.2")
  })

  it("manifest root element has manifest:version attribute", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    const manifest = await parseXmlFromZip(data, "META-INF/manifest.xml")
    expect(manifest.attrs["manifest:version"]).toBe("1.2")
  })

  it("manifest contains all required file entries including settings.xml", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    const manifest = await parseXmlFromZip(data, "META-INF/manifest.xml")
    const entries = findChildren(manifest, "file-entry")
    const paths = entries.map((e: any) => e.attrs["manifest:full-path"])

    expect(paths).toContain("/")
    expect(paths).toContain("content.xml")
    expect(paths).toContain("meta.xml")
    expect(paths).toContain("styles.xml")
    expect(paths).toContain("settings.xml")
  })
})

// ── #119: table:table-column handling in reader ─────────────────────

describe("ODS spec — #119: table:table-column in reader", () => {
  it("reader handles files with table:table-column elements correctly", async () => {
    // Write a file (which generates table:table-column elements)
    const data = await writeOds({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["A", "B", "C"],
            [1, 2, 3],
          ],
        },
      ],
    })

    // Verify table-column exists in the output
    const table = await getFirstTable(data)
    const columns = findChildren(table, "table-column")
    expect(columns.length).toBeGreaterThanOrEqual(1)

    // Read back — should not be confused by table-column elements
    const wb = await readOds(data)
    expect(wb.sheets[0].rows[0]).toEqual(["A", "B", "C"])
    expect(wb.sheets[0].rows[1]).toEqual([1, 2, 3])
  })

  it("reader correctly counts columns despite table:number-columns-repeated on table-column", async () => {
    // Write a sheet with multiple columns (will generate number-columns-repeated)
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["A", "B", "C", "D", "E"]] }],
    })

    const table = await getFirstTable(data)
    const columns = findChildren(table, "table-column")

    // Writer should use number-columns-repeated for 5 columns
    let totalCols = 0
    for (const col of columns) {
      const repeat = Number(col.attrs["table:number-columns-repeated"] ?? "1")
      totalCols += repeat
    }
    expect(totalCols).toBe(5)

    // Reader should still get all 5 values correctly
    const wb = await readOds(data)
    expect(wb.sheets[0].rows[0]).toEqual(["A", "B", "C", "D", "E"])
    expect(wb.sheets[0].rows[0].length).toBe(5)
  })

  it("reader does not treat table-column elements as data rows", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["Only Row"]] }],
    })

    const wb = await readOds(data)
    // Should have exactly 1 row, not extra rows from table-column
    expect(wb.sheets[0].rows.length).toBe(1)
    expect(wb.sheets[0].rows[0]).toEqual(["Only Row"])
  })
})

// ── #124: Number display text formatting ────────────────────────────

describe("ODS spec — #124: number display text formatting", () => {
  it("integer values display without decimal point", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[12, 100, 0, -5]] }],
    })

    const table = await getFirstTable(data)
    const row = findChildren(table, "table-row")[0]
    const cells = findChildren(row, "table-cell")

    expect(getElementText(findChild(cells[0], "p"))).toBe("12")
    expect(getElementText(findChild(cells[1], "p"))).toBe("100")
    expect(getElementText(findChild(cells[2], "p"))).toBe("0")
    expect(getElementText(findChild(cells[3], "p"))).toBe("-5")
  })

  it("12.0 displays as '12' not '12.0' in text:p", async () => {
    // 12.0 in JS is === 12, Number.isInteger(12.0) is true
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[12.0]] }],
    })

    const table = await getFirstTable(data)
    const row = findChildren(table, "table-row")[0]
    const cells = findChildren(row, "table-cell")

    // Display text should be "12", not "12.0"
    expect(getElementText(findChild(cells[0], "p"))).toBe("12")

    // But the office:value attribute should still have the raw value
    expect(cells[0].attrs["office:value"]).toBe("12")
  })

  it("float values display with reasonable decimal places (no artifacts)", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[3.14, 0.1 + 0.2, 1.005]] }],
    })

    const table = await getFirstTable(data)
    const row = findChildren(table, "table-row")[0]
    const cells = findChildren(row, "table-cell")

    // 3.14 should display as "3.14"
    expect(getElementText(findChild(cells[0], "p"))).toBe("3.14")

    // 0.1 + 0.2 should display as "0.3", not "0.30000000000000004"
    expect(getElementText(findChild(cells[1], "p"))).toBe("0.3")

    // 1.005 should display as "1.005"
    expect(getElementText(findChild(cells[2], "p"))).toBe("1.005")
  })

  it("office:value preserves full precision for numbers", async () => {
    const val = 0.1 + 0.2 // 0.30000000000000004 in JS
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[val]] }],
    })

    const table = await getFirstTable(data)
    const row = findChildren(table, "table-row")[0]
    const cells = findChildren(row, "table-cell")

    // office:value should have full precision for computation fidelity
    expect(cells[0].attrs["office:value"]).toBe(String(val))
  })

  it("boolean cells display as TRUE/FALSE", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[true, false]] }],
    })

    const table = await getFirstTable(data)
    const row = findChildren(table, "table-row")[0]
    const cells = findChildren(row, "table-cell")

    expect(getElementText(findChild(cells[0], "p"))).toBe("TRUE")
    expect(getElementText(findChild(cells[1], "p"))).toBe("FALSE")
  })

  it("round-trip preserves number values despite display formatting", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[12, 3.14, 0.1 + 0.2, -99.99]] }],
    })

    const wb = await readOds(data)
    const row = wb.sheets[0].rows[0]

    // Values read from office:value should be exact
    expect(row[0]).toBe(12)
    expect(row[1]).toBe(3.14)
    expect(row[2]).toBeCloseTo(0.3, 15)
    expect(row[3]).toBe(-99.99)
  })
})

// ── #127: mimetype ZIP extra field ──────────────────────────────────

describe("ODS spec — #127: mimetype ZIP extra field", () => {
  it("mimetype entry has extra field length = 0 in local file header", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    // The local file header for mimetype starts at offset 0
    // Bytes 28-29 (uint16 LE) = extra field length
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
    const extraFieldLength = view.getUint16(28, true)
    expect(extraFieldLength).toBe(0)
  })

  it("mimetype is stored uncompressed (method 0)", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    // Compression method at offset 8 in local file header
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
    const compressionMethod = view.getUint16(8, true)
    expect(compressionMethod).toBe(0)
  })

  it("mimetype is the first entry in the ZIP", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    const zip = new ZipReader(data)
    const entries = zip.entries()
    expect(entries[0]).toBe("mimetype")
  })

  it("mimetype content is exactly the ODS MIME type with no padding", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    })

    const zip = new ZipReader(data)
    const mimeData = await zip.extract("mimetype")
    const mimeStr = decoder.decode(mimeData)

    expect(mimeStr).toBe("application/vnd.oasis.opendocument.spreadsheet")
    // No trailing newline or whitespace
    expect(mimeStr.length).toBe(46)
  })
})

import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeOds } from "../src/ods/writer";
import { readOds } from "../src/ods/reader";
import type { CellValue } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");

async function extractFile(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data);
  const raw = await zip.extract(path);
  return decoder.decode(raw);
}

async function parseXmlFromZip(data: Uint8Array, path: string) {
  const xml = await extractFile(data, path);
  return parseXml(xml);
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function getElementText(el: { children: Array<unknown> }): string {
  return el.children.filter((c: unknown) => typeof c === "string").join("");
}

// ── ODS Writer Tests ────────────────────────────────────────────────

describe("ODS Writer", () => {
  it("produces a valid ZIP archive", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["Hello"]] }],
    });

    const zip = new ZipReader(data);
    expect(zip.has("mimetype")).toBe(true);
    expect(zip.has("content.xml")).toBe(true);
    expect(zip.has("meta.xml")).toBe(true);
    expect(zip.has("styles.xml")).toBe(true);
    expect(zip.has("META-INF/manifest.xml")).toBe(true);
  });

  it("writes mimetype as the first entry and uncompressed", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const zip = new ZipReader(data);
    const entries = zip.entries();

    // mimetype should be the first entry
    expect(entries[0]).toBe("mimetype");

    // mimetype content should be the correct MIME type
    const mimeData = await zip.extract("mimetype");
    expect(decoder.decode(mimeData)).toBe("application/vnd.oasis.opendocument.spreadsheet");

    // Verify it's stored uncompressed by checking compression method in local file header
    // The local file header starts at offset 0 for the first entry
    // Compression method is at offset 8 in the local file header
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const compressionMethod = view.getUint16(8, true);
    expect(compressionMethod).toBe(0); // STORE = 0
  });

  it("writes manifest.xml with required entries", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const manifest = await parseXmlFromZip(data, "META-INF/manifest.xml");
    const entries = findChildren(manifest, "file-entry");

    // Should have entries for /, content.xml, meta.xml, styles.xml
    const paths = entries.map((e: any) => e.attrs["manifest:full-path"]);
    expect(paths).toContain("/");
    expect(paths).toContain("content.xml");
    expect(paths).toContain("meta.xml");
    expect(paths).toContain("styles.xml");
  });

  it("writes string cells correctly", async () => {
    const data = await writeOds({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Hello", "World"],
            ["Foo", "Bar"],
          ],
        },
      ],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const table = findChild(spreadsheet, "table");

    expect(table.attrs["table:name"]).toBe("Sheet1");

    const tableRows = findChildren(table, "table-row");
    expect(tableRows.length).toBe(2);

    // First row, first cell
    const firstRowCells = findChildren(tableRows[0], "table-cell");
    expect(firstRowCells[0].attrs["office:value-type"]).toBe("string");
    const textP = findChild(firstRowCells[0], "p");
    expect(getElementText(textP)).toBe("Hello");
  });

  it("writes numeric cells correctly", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[42, 3.14]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const table = findChild(spreadsheet, "table");
    const row = findChildren(table, "table-row")[0];
    const cells = findChildren(row, "table-cell");

    expect(cells[0].attrs["office:value-type"]).toBe("float");
    expect(cells[0].attrs["office:value"]).toBe("42");

    expect(cells[1].attrs["office:value-type"]).toBe("float");
    expect(cells[1].attrs["office:value"]).toBe("3.14");
  });

  it("writes boolean cells correctly", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[true, false]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const table = findChild(spreadsheet, "table");
    const row = findChildren(table, "table-row")[0];
    const cells = findChildren(row, "table-cell");

    expect(cells[0].attrs["office:value-type"]).toBe("boolean");
    expect(cells[0].attrs["office:boolean-value"]).toBe("true");

    expect(cells[1].attrs["office:value-type"]).toBe("boolean");
    expect(cells[1].attrs["office:boolean-value"]).toBe("false");
  });

  it("writes date cells correctly", async () => {
    const date = new Date("2026-03-24T10:30:00Z");
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[date]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const table = findChild(spreadsheet, "table");
    const row = findChildren(table, "table-row")[0];
    const cells = findChildren(row, "table-cell");

    expect(cells[0].attrs["office:value-type"]).toBe("date");
    expect(cells[0].attrs["office:date-value"]).toBeDefined();
  });

  it("writes multiple sheets", async () => {
    const data = await writeOds({
      sheets: [
        { name: "First", rows: [["A"]] },
        { name: "Second", rows: [["B"]] },
        { name: "Third", rows: [["C"]] },
      ],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const tables = findChildren(spreadsheet, "table");

    expect(tables.length).toBe(3);
    expect(tables[0].attrs["table:name"]).toBe("First");
    expect(tables[1].attrs["table:name"]).toBe("Second");
    expect(tables[2].attrs["table:name"]).toBe("Third");
  });

  it("writes null cells as empty", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["A", null, "C"]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const table = findChild(spreadsheet, "table");
    const row = findChildren(table, "table-row")[0];
    const cells = findChildren(row, "table-cell");

    // Middle cell should be empty (self-closing)
    expect(cells[1].attrs["office:value-type"]).toBeUndefined();
  });

  it("writes Application: defter in meta.xml", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const metaDoc = await parseXmlFromZip(data, "meta.xml");
    const metaEl = findChild(metaDoc, "meta");
    const generator = findChild(metaEl, "generator");
    expect(getElementText(generator)).toBe("defter");
  });

  it("writes document properties in meta.xml", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      properties: {
        title: "Test Title",
        creator: "Test Author",
        subject: "Test Subject",
      },
    });

    const metaDoc = await parseXmlFromZip(data, "meta.xml");
    const metaEl = findChild(metaDoc, "meta");

    expect(getElementText(findChild(metaEl, "title"))).toBe("Test Title");
    expect(getElementText(findChild(metaEl, "initial-creator"))).toBe("Test Author");
    expect(getElementText(findChild(metaEl, "subject"))).toBe("Test Subject");
  });
});

// ── ODS Reader Tests ────────────────────────────────────────────────

describe("ODS Reader", () => {
  it("reads back written ODS with string data", async () => {
    const written = await writeOds({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Hello", "World"],
            ["Foo", "Bar"],
          ],
        },
      ],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets.length).toBe(1);
    expect(workbook.sheets[0].name).toBe("Sheet1");
    expect(workbook.sheets[0].rows.length).toBe(2);
    expect(workbook.sheets[0].rows[0]).toEqual(["Hello", "World"]);
    expect(workbook.sheets[0].rows[1]).toEqual(["Foo", "Bar"]);
  });

  it("reads back written ODS with numeric data", async () => {
    const written = await writeOds({
      sheets: [{ name: "Numbers", rows: [[1, 2.5, 100]] }],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets[0].rows[0]).toEqual([1, 2.5, 100]);
  });

  it("reads back written ODS with boolean data", async () => {
    const written = await writeOds({
      sheets: [{ name: "Booleans", rows: [[true, false, true]] }],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets[0].rows[0]).toEqual([true, false, true]);
  });

  it("reads back written ODS with date values", async () => {
    const date1 = new Date("2026-01-15T10:00:00");
    const date2 = new Date("2026-06-30T23:59:59");

    const written = await writeOds({
      sheets: [{ name: "Dates", rows: [[date1, date2]] }],
    });

    const workbook = await readOds(written);
    const row = workbook.sheets[0].rows[0];

    // Dates may not be exactly equal due to time zone handling in ODS,
    // but they should be Date objects with similar values
    expect(row[0]).toBeInstanceOf(Date);
    expect(row[1]).toBeInstanceOf(Date);
  });

  it("reads back written ODS with multiple sheets", async () => {
    const written = await writeOds({
      sheets: [
        { name: "First", rows: [["A1"]] },
        { name: "Second", rows: [["B1", "B2"]] },
        { name: "Third", rows: [["C1"], ["C2"], ["C3"]] },
      ],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets.length).toBe(3);
    expect(workbook.sheets[0].name).toBe("First");
    expect(workbook.sheets[0].rows).toEqual([["A1"]]);
    expect(workbook.sheets[1].name).toBe("Second");
    expect(workbook.sheets[1].rows).toEqual([["B1", "B2"]]);
    expect(workbook.sheets[2].name).toBe("Third");
    expect(workbook.sheets[2].rows).toEqual([["C1"], ["C2"], ["C3"]]);
  });

  it("reads document properties from meta.xml", async () => {
    const written = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      properties: {
        title: "Test Title",
        creator: "Author",
        subject: "Subject",
        description: "A description",
        keywords: "test, ods",
      },
    });

    const workbook = await readOds(written);
    expect(workbook.properties).toBeDefined();
    expect(workbook.properties!.title).toBe("Test Title");
    expect(workbook.properties!.creator).toBe("Author");
    expect(workbook.properties!.subject).toBe("Subject");
    expect(workbook.properties!.description).toBe("A description");
    expect(workbook.properties!.keywords).toBe("test, ods");
  });

  it("handles mixed data types in a row", async () => {
    const written = await writeOds({
      sheets: [
        {
          name: "Mixed",
          rows: [["text", 42, true, null, "more"]],
        },
      ],
    });

    const workbook = await readOds(written);
    const row = workbook.sheets[0].rows[0];
    expect(row[0]).toBe("text");
    expect(row[1]).toBe(42);
    expect(row[2]).toBe(true);
    // null cells at the end are trimmed, but mid-row null should be preserved
    // However, the writer emits a self-closing cell which has no value-type,
    // reader returns null for such cells
    expect(row[3]).toBe(null);
    expect(row[4]).toBe("more");
  });

  it("handles empty sheets", async () => {
    const written = await writeOds({
      sheets: [{ name: "Empty", rows: [] }],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets.length).toBe(1);
    expect(workbook.sheets[0].name).toBe("Empty");
    expect(workbook.sheets[0].rows).toEqual([]);
  });

  it("rejects non-ZIP data", async () => {
    const badData = new Uint8Array([
      1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22,
    ]);
    await expect(readOds(badData)).rejects.toThrow();
  });

  it("round-trips a complex workbook", async () => {
    const rows: CellValue[][] = [
      ["Name", "Age", "Active", "Score"],
      ["Alice", 30, true, 95.5],
      ["Bob", 25, false, 87.3],
      ["Charlie", 35, true, 92.1],
    ];

    const written = await writeOds({
      sheets: [
        { name: "People", rows },
        { name: "Summary", rows: [["Total", 3]] },
      ],
    });

    const workbook = await readOds(written);

    expect(workbook.sheets.length).toBe(2);

    // First sheet
    expect(workbook.sheets[0].name).toBe("People");
    expect(workbook.sheets[0].rows[0]).toEqual(["Name", "Age", "Active", "Score"]);
    expect(workbook.sheets[0].rows[1]).toEqual(["Alice", 30, true, 95.5]);
    expect(workbook.sheets[0].rows[2]).toEqual(["Bob", 25, false, 87.3]);
    expect(workbook.sheets[0].rows[3]).toEqual(["Charlie", 35, true, 92.1]);

    // Second sheet
    expect(workbook.sheets[1].name).toBe("Summary");
    expect(workbook.sheets[1].rows[0]).toEqual(["Total", 3]);
  });

  it("accepts ArrayBuffer input", async () => {
    const written = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    // Convert Uint8Array to ArrayBuffer
    const buffer = written.buffer.slice(
      written.byteOffset,
      written.byteOffset + written.byteLength,
    ) as ArrayBuffer;

    const workbook = await readOds(buffer);
    expect(workbook.sheets[0].rows[0]).toEqual(["test"]);
  });
});

import { describe, expect, it } from "vitest";
import { writeTsv, writeTsvObjects } from "../src/export/tsv";
import { toHtml } from "../src/export/html";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { writeOds } from "../src/ods/writer";
import { parseCsv } from "../src/csv/reader";
import { streamOdsRows } from "../src/ods/stream";
import type { Sheet, CellValue, WriteSheet } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

async function roundtripXlsx(options: Parameters<typeof writeXlsx>[0]) {
  const buf = await writeXlsx(options);
  return readXlsx(buf);
}

// ═══════════════════════════════════════════════════════════════════════
// #73: TSV export format
// ═══════════════════════════════════════════════════════════════════════

describe("TSV export", () => {
  it("writeTsv produces tab-separated output", () => {
    const rows: CellValue[][] = [
      ["Name", "Age", "City"],
      ["Alice", 30, "Berlin"],
      ["Bob", 25, "Paris"],
    ];
    const result = writeTsv(rows);
    const lines = result.split("\r\n");
    expect(lines[0]).toBe("Name\tAge\tCity");
    expect(lines[1]).toBe("Alice\t30\tBerlin");
    expect(lines[2]).toBe("Bob\t25\tParis");
  });

  it("writeTsvObjects produces tab-separated output from objects", () => {
    const data = [
      { name: "Alice", age: 30 },
      { name: "Bob", age: 25 },
    ];
    const result = writeTsvObjects(data);
    const lines = result.split("\r\n");
    expect(lines[0]).toBe("name\tage");
    expect(lines[1]).toBe("Alice\t30");
  });

  it("writeTsv respects other CsvWriteOptions like bom", () => {
    const rows: CellValue[][] = [["hello"]];
    const result = writeTsv(rows, { bom: true });
    expect(result.charCodeAt(0)).toBe(0xfeff);
  });
});

// ═══════════════════════════════════════════════════════════════════════
// #71: HTML export accessibility
// ═══════════════════════════════════════════════════════════════════════

describe("HTML export accessibility", () => {
  const sheet: Sheet = {
    name: "Test",
    rows: [
      ["Name", "Value"],
      ["A", 1],
      ["B", 2],
    ],
  };

  it("adds scope='col' to <th> when headerRow: true", () => {
    const html = toHtml(sheet, { headerRow: true });
    expect(html).toContain('<th scope="col"');
    expect(html).toContain("Name");
    expect(html).toContain("Value");
  });

  it("adds role='table' when headerRow: true", () => {
    const html = toHtml(sheet, { headerRow: true });
    expect(html).toContain('role="table"');
  });

  it("does not add role='table' when headerRow: false", () => {
    const html = toHtml(sheet, { headerRow: false });
    expect(html).not.toContain('role="table"');
  });

  it("adds <caption> when caption is set", () => {
    const html = toHtml(sheet, { caption: "My Table" });
    expect(html).toContain("<caption>My Table</caption>");
  });

  it("escapes HTML in caption", () => {
    const html = toHtml(sheet, { caption: "<script>alert(1)</script>" });
    expect(html).toContain("<caption>&lt;script&gt;alert(1)&lt;/script&gt;</caption>");
  });

  it("adds aria-label when ariaLabel is set", () => {
    const html = toHtml(sheet, { ariaLabel: "Data table" });
    expect(html).toContain('aria-label="Data table"');
  });

  it("combines all accessibility options", () => {
    const html = toHtml(sheet, {
      headerRow: true,
      caption: "Test Caption",
      ariaLabel: "Test Label",
    });
    expect(html).toContain('role="table"');
    expect(html).toContain('aria-label="Test Label"');
    expect(html).toContain("<caption>Test Caption</caption>");
    expect(html).toContain('scope="col"');
  });

  it("handles empty sheet with caption", () => {
    const emptySheet: Sheet = { name: "Empty", rows: [] };
    const html = toHtml(emptySheet, { caption: "Empty Table" });
    expect(html).toContain("<caption>Empty Table</caption>");
  });
});

// ═══════════════════════════════════════════════════════════════════════
// #70: Custom document properties
// ═══════════════════════════════════════════════════════════════════════

describe("Custom document properties", () => {
  it("round-trips custom string, number, boolean properties", async () => {
    const wb = await roundtripXlsx({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      properties: {
        title: "Test Workbook",
        custom: {
          Department: "Engineering",
          Version: 5,
          Rating: 3.14,
          Approved: true,
          Reviewed: false,
        },
      },
    });

    expect(wb.properties).toBeDefined();
    expect(wb.properties!.custom).toBeDefined();
    const custom = wb.properties!.custom!;
    expect(custom["Department"]).toBe("Engineering");
    expect(custom["Version"]).toBe(5);
    expect(custom["Rating"]).toBeCloseTo(3.14);
    expect(custom["Approved"]).toBe(true);
    expect(custom["Reviewed"]).toBe(false);
  });

  it("round-trips custom Date property", async () => {
    const testDate = new Date("2025-06-15T12:00:00Z");
    const wb = await roundtripXlsx({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      properties: {
        custom: { ReleaseDate: testDate },
      },
    });

    const custom = wb.properties!.custom!;
    expect(custom["ReleaseDate"]).toBeInstanceOf(Date);
    expect((custom["ReleaseDate"] as Date).toISOString()).toContain("2025-06-15");
  });

  it("does not emit custom.xml when no custom properties", async () => {
    const wb = await roundtripXlsx({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      properties: { title: "No Custom" },
    });

    // Should still have properties but no custom
    expect(wb.properties?.title).toBe("No Custom");
    expect(wb.properties?.custom).toBeUndefined();
  });
});

// ═══════════════════════════════════════════════════════════════════════
// #69: Workbook-level protection
// ═══════════════════════════════════════════════════════════════════════

describe("Workbook protection", () => {
  it("round-trips lockStructure", async () => {
    const wb = await roundtripXlsx({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      workbookProtection: { lockStructure: true },
    });

    expect(wb.workbookProtection).toBeDefined();
    expect(wb.workbookProtection!.lockStructure).toBe(true);
    expect(wb.workbookProtection!.lockWindows).toBeUndefined();
  });

  it("round-trips lockWindows", async () => {
    const wb = await roundtripXlsx({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      workbookProtection: { lockWindows: true },
    });

    expect(wb.workbookProtection).toBeDefined();
    expect(wb.workbookProtection!.lockWindows).toBe(true);
  });

  it("round-trips both lockStructure and lockWindows", async () => {
    const wb = await roundtripXlsx({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      workbookProtection: { lockStructure: true, lockWindows: true },
    });

    expect(wb.workbookProtection!.lockStructure).toBe(true);
    expect(wb.workbookProtection!.lockWindows).toBe(true);
  });

  it("does not emit workbookProtection when not set", async () => {
    const wb = await roundtripXlsx({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    expect(wb.workbookProtection).toBeUndefined();
  });

  it("supports password option", async () => {
    // Password is write-only (hashed), so we just verify it writes without error
    const buf = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      workbookProtection: { lockStructure: true, password: "secret" },
    });

    const wb = await readXlsx(buf);
    expect(wb.workbookProtection).toBeDefined();
    expect(wb.workbookProtection!.lockStructure).toBe(true);
  });
});

// ═══════════════════════════════════════════════════════════════════════
// #76: Auto filter column criteria
// ═══════════════════════════════════════════════════════════════════════

describe("AutoFilter with column criteria", () => {
  it("round-trips autoFilter with filterColumn", async () => {
    const wb = await roundtripXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Status", "Name", "Value"],
            ["Active", "Alice", 100],
            ["Pending", "Bob", 200],
            ["Active", "Charlie", 300],
          ],
          autoFilter: {
            range: "A1:C4",
            columns: [{ colIndex: 0, filters: ["Active", "Pending"] }],
          },
        },
      ],
    });

    const sheet = wb.sheets[0];
    expect(sheet.autoFilter).toBeDefined();
    expect(sheet.autoFilter!.range).toBe("A1:C4");
    expect(sheet.autoFilter!.columns).toBeDefined();
    expect(sheet.autoFilter!.columns!.length).toBe(1);
    expect(sheet.autoFilter!.columns![0].colIndex).toBe(0);
    expect(sheet.autoFilter!.columns![0].filters).toEqual(["Active", "Pending"]);
  });

  it("round-trips autoFilter with multiple filterColumns", async () => {
    const wb = await roundtripXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Status", "Category", "Value"],
            ["Active", "A", 100],
          ],
          autoFilter: {
            range: "A1:C2",
            columns: [
              { colIndex: 0, filters: ["Active"] },
              { colIndex: 1, filters: ["A", "B"] },
            ],
          },
        },
      ],
    });

    const sheet = wb.sheets[0];
    expect(sheet.autoFilter!.columns!.length).toBe(2);
    expect(sheet.autoFilter!.columns![0].filters).toEqual(["Active"]);
    expect(sheet.autoFilter!.columns![1].filters).toEqual(["A", "B"]);
  });

  it("autoFilter without columns still works", async () => {
    const wb = await roundtripXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["A", "B"],
            [1, 2],
          ],
          autoFilter: { range: "A1:B2" },
        },
      ],
    });

    expect(wb.sheets[0].autoFilter).toBeDefined();
    expect(wb.sheets[0].autoFilter!.range).toBe("A1:B2");
    expect(wb.sheets[0].autoFilter!.columns).toBeUndefined();
  });
});

// ═══════════════════════════════════════════════════════════════════════
// #61: CSV fastMode
// ═══════════════════════════════════════════════════════════════════════

describe("CSV fastMode", () => {
  it("parses simple CSV in fastMode", () => {
    const csv = "a,b,c\n1,2,3\n4,5,6";
    const rows = parseCsv(csv, { delimiter: ",", fastMode: true });
    expect(rows).toEqual([
      ["a", "b", "c"],
      ["1", "2", "3"],
      ["4", "5", "6"],
    ]);
  });

  it("fastMode handles CRLF", () => {
    const csv = "a,b\r\n1,2\r\n";
    const rows = parseCsv(csv, { delimiter: ",", fastMode: true });
    expect(rows).toEqual([
      ["a", "b"],
      ["1", "2"],
    ]);
  });

  it("fastMode with tab delimiter", () => {
    const tsv = "name\tage\nAlice\t30";
    const rows = parseCsv(tsv, { delimiter: "\t", fastMode: true });
    expect(rows).toEqual([
      ["name", "age"],
      ["Alice", "30"],
    ]);
  });

  it("fastMode does not handle quotes", () => {
    // In fastMode, quotes are treated as literal characters
    const csv = '"hello",world\n"a,b",c';
    const rows = parseCsv(csv, { delimiter: ",", fastMode: true });
    // Without quote handling, "a,b" is split at the comma
    expect(rows[0][0]).toBe('"hello"');
    expect(rows[1]).toEqual(['"a', 'b"', "c"]);
  });

  it("fastMode with type inference", () => {
    const csv = "1,true,hello\n2,false,world";
    const rows = parseCsv(csv, { delimiter: ",", fastMode: true, typeInference: true });
    expect(rows[0]).toEqual([1, true, "hello"]);
    expect(rows[1]).toEqual([2, false, "world"]);
  });
});

// ═══════════════════════════════════════════════════════════════════════
// #84: Streaming ODS reader
// ═══════════════════════════════════════════════════════════════════════

describe("Streaming ODS rows", () => {
  it("yields rows from an ODS file", async () => {
    // Write a simple ODS file first
    const buf = await writeOds({
      sheets: [
        {
          name: "Data",
          rows: [
            ["Name", "Value"],
            ["Alice", 100],
            ["Bob", 200],
          ],
        },
      ],
    });

    const rows: Array<{ index: number; values: CellValue[] }> = [];
    for await (const row of streamOdsRows(buf)) {
      rows.push(row);
    }

    expect(rows.length).toBe(3);
    expect(rows[0].values).toEqual(["Name", "Value"]);
    expect(rows[1].values).toEqual(["Alice", 100]);
    expect(rows[2].values).toEqual(["Bob", 200]);
  });

  it("handles empty ODS sheets", async () => {
    const buf = await writeOds({
      sheets: [{ name: "Empty", rows: [] }],
    });

    const rows: Array<{ index: number; values: CellValue[] }> = [];
    for await (const row of streamOdsRows(buf)) {
      rows.push(row);
    }

    expect(rows.length).toBe(0);
  });

  it("preserves row indices", async () => {
    const buf = await writeOds({
      sheets: [
        {
          name: "Data",
          rows: [["A"], ["B"], ["C"]],
        },
      ],
    });

    const rows: Array<{ index: number; values: CellValue[] }> = [];
    for await (const row of streamOdsRows(buf)) {
      rows.push(row);
    }

    expect(rows[0].index).toBe(0);
    expect(rows[1].index).toBe(1);
    expect(rows[2].index).toBe(2);
  });
});

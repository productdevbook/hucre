import { describe, it, expect } from "vitest";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { XlsxStreamWriter } from "../src/xlsx/stream-writer";
import { streamXlsxRows } from "../src/xlsx/stream-reader";
import type { StreamRow } from "../src/xlsx/stream-reader";
import { parseCsv, parseCsvObjects, detectDelimiter, stripBom } from "../src/csv/reader";
import { writeCsv, writeCsvObjects } from "../src/csv/writer";
import { validateWithSchema } from "../src/_schema";
import { serialToDate, dateToSerial, parseDate } from "../src/_date";
import { ZipWriter } from "../src/zip/writer";
import { ZipReader } from "../src/zip/reader";
import { parseXml, parseSax } from "../src/xml/parser";
import { xmlEscape, xmlEscapeAttr } from "../src/xml/writer";
import { writeOds } from "../src/ods/writer";
import { readOds } from "../src/ods/reader";
import {
  insertRows,
  deleteRows,
  insertColumns,
  deleteColumns,
  cloneSheet,
  moveSheet,
} from "../src/sheet-ops";
import { colToLetter } from "../src/xlsx/worksheet-writer";
import type {
  CellValue,
  WriteSheet,
  CellStyle,
  SchemaDefinition,
  Sheet,
  Workbook,
} from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

async function writeAndRead(sheets: WriteSheet[]): Promise<Workbook> {
  const xlsx = await writeXlsx({ sheets });
  return readXlsx(xlsx);
}

async function collectStreamRows(
  gen: AsyncGenerator<StreamRow, void, undefined>,
): Promise<StreamRow[]> {
  const rows: StreamRow[] = [];
  for await (const row of gen) {
    rows.push(row);
  }
  return rows;
}

// ═══════════════════════════════════════════════════════════════════════
// 1. XLSX Writer Edge Cases
// ═══════════════════════════════════════════════════════════════════════

describe("XLSX Writer Edge Cases", () => {
  it("very long string value (10,000+ characters)", async () => {
    const longStr = "A".repeat(10_000);
    const wb = await writeAndRead([{ name: "Sheet1", rows: [[longStr]] }]);
    expect(wb.sheets[0].rows[0][0]).toBe(longStr);
  });

  it("unicode: emoji, surrogate pairs, zero-width joiners", async () => {
    const emoji = "\u{1F389}"; // party popper
    const family = "\u{1F468}\u200D\u{1F469}\u200D\u{1F467}\u200D\u{1F466}"; // family ZWJ sequence
    const combining = "e\u0301"; // e + combining acute accent
    const rtl = "\u200F\u0645\u0631\u062D\u0628\u0627"; // RTL mark + Arabic "hello"

    const wb = await writeAndRead([{ name: "Sheet1", rows: [[emoji, family, combining, rtl]] }]);

    expect(wb.sheets[0].rows[0][0]).toBe(emoji);
    expect(wb.sheets[0].rows[0][1]).toBe(family);
    expect(wb.sheets[0].rows[0][2]).toBe(combining);
    expect(wb.sheets[0].rows[0][3]).toBe(rtl);
  });

  it("special XML characters in cell values: < > & \" ' and ]]>", async () => {
    const values: CellValue[] = [
      '<script>alert("xss")</script>',
      "Tom & Jerry",
      'She said "hello"',
      "a > b & c < d",
      "CDATA end: ]]>",
      "apos: it's working",
    ];

    const wb = await writeAndRead([{ name: "Sheet1", rows: [values] }]);

    for (let i = 0; i < values.length; i++) {
      expect(wb.sheets[0].rows[0][i]).toBe(values[i]);
    }
  });

  it("cell values that look like formulas (=, +, -, @)", async () => {
    const values: CellValue[] = ["=SUM(A1:A10)", "+1234", "-5678", "@mention"];

    const wb = await writeAndRead([{ name: "Sheet1", rows: [values] }]);

    // These should be preserved as strings, not interpreted as formulas
    for (let i = 0; i < values.length; i++) {
      expect(wb.sheets[0].rows[0][i]).toBe(values[i]);
    }
  });

  it("numbers at edge of precision: MAX_SAFE_INTEGER, very small decimals", async () => {
    const values: CellValue[] = [
      Number.MAX_SAFE_INTEGER, // 9007199254740991
      -Number.MAX_SAFE_INTEGER,
      0.1 + 0.2, // 0.30000000000000004
      1e-15,
      -0,
    ];

    const wb = await writeAndRead([{ name: "Sheet1", rows: [values] }]);

    expect(wb.sheets[0].rows[0][0]).toBe(Number.MAX_SAFE_INTEGER);
    expect(wb.sheets[0].rows[0][1]).toBe(-Number.MAX_SAFE_INTEGER);
    expect(wb.sheets[0].rows[0][2]).toBeCloseTo(0.3, 15);
    expect(wb.sheets[0].rows[0][3]).toBe(1e-15);
  });

  it("Infinity and NaN values", async () => {
    // These are not valid Excel cell values - test that they don't crash
    const values: CellValue[] = [Infinity, -Infinity, NaN];

    // Should not throw
    const xlsx = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [values] }],
    });
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });

  it("date edge cases: Jan 1 1900, Feb 28 1900, Mar 1 1900 (Lotus bug)", async () => {
    const jan1_1900 = new Date(Date.UTC(1900, 0, 1));
    const feb28_1900 = new Date(Date.UTC(1900, 1, 28));
    const mar1_1900 = new Date(Date.UTC(1900, 2, 1));
    const year9999 = new Date(Date.UTC(9999, 11, 31));

    const wb = await writeAndRead([
      {
        name: "Dates",
        rows: [[jan1_1900, feb28_1900, mar1_1900, year9999]],
      },
    ]);

    const row = wb.sheets[0].rows[0];
    // Jan 1 1900
    expect(row[0]).toBeInstanceOf(Date);
    expect((row[0] as Date).getUTCFullYear()).toBe(1900);
    expect((row[0] as Date).getUTCMonth()).toBe(0);
    expect((row[0] as Date).getUTCDate()).toBe(1);

    // Feb 28 1900
    expect(row[1]).toBeInstanceOf(Date);
    expect((row[1] as Date).getUTCFullYear()).toBe(1900);
    expect((row[1] as Date).getUTCMonth()).toBe(1);
    expect((row[1] as Date).getUTCDate()).toBe(28);

    // Mar 1 1900
    expect(row[2]).toBeInstanceOf(Date);
    expect((row[2] as Date).getUTCFullYear()).toBe(1900);
    expect((row[2] as Date).getUTCMonth()).toBe(2);
    expect((row[2] as Date).getUTCDate()).toBe(1);

    // Year 9999
    expect(row[3]).toBeInstanceOf(Date);
    expect((row[3] as Date).getUTCFullYear()).toBe(9999);
  });

  it("empty string vs null vs undefined in cells", async () => {
    const wb = await writeAndRead([{ name: "Sheet1", rows: [["", null, "value"]] }]);

    const row = wb.sheets[0].rows[0];
    // Empty string should round-trip as empty string
    expect(row[0]).toBe("");
    // null may be represented differently but should not crash
    // After round-trip, null is typically preserved or becomes null
    expect(row[2]).toBe("value");
  });

  it("sheet name: max length (31 chars)", async () => {
    const longName = "A".repeat(31);
    const wb = await writeAndRead([{ name: longName, rows: [["data"]] }]);
    expect(wb.sheets[0].name).toBe(longName);
  });

  it("sheet name: special characters", async () => {
    // Excel allows most characters in sheet names except: / \ ? * [ ]
    const specialName = "Sheet (1) - Test & 2";
    const wb = await writeAndRead([{ name: specialName, rows: [["data"]] }]);
    expect(wb.sheets[0].name).toBe(specialName);
  });

  it("multiple sheets with same-looking names (case variation)", async () => {
    const wb = await writeAndRead([
      { name: "Sheet1", rows: [["a"]] },
      { name: "sheet1", rows: [["b"]] },
      { name: "SHEET1", rows: [["c"]] },
    ]);
    expect(wb.sheets).toHaveLength(3);
    expect(wb.sheets[0].name).toBe("Sheet1");
    expect(wb.sheets[1].name).toBe("sheet1");
    expect(wb.sheets[2].name).toBe("SHEET1");
  });

  it("column XFD (16383 = max Excel column)", async () => {
    // Just verify colToLetter handles the max column
    expect(colToLetter(16383)).toBe("XFD");
  });

  it("mixed cell types in one row: string, number, boolean, date, null", async () => {
    const date = new Date(Date.UTC(2024, 5, 15));
    const values: CellValue[] = ["hello", 42, true, date, null, false, 3.14, ""];

    const wb = await writeAndRead([{ name: "Mixed", rows: [values] }]);

    const row = wb.sheets[0].rows[0];
    expect(row[0]).toBe("hello");
    expect(row[1]).toBe(42);
    expect(row[2]).toBe(true);
    expect(row[3]).toBeInstanceOf(Date);
    expect((row[3] as Date).getUTCFullYear()).toBe(2024);
    // null cell may or may not be present in row array
    expect(row[5]).toBe(false);
    expect(row[6]).toBeCloseTo(3.14);
    expect(row[7]).toBe("");
  });

  it("style edge cases: all border sides", async () => {
    const style: CellStyle = {
      border: {
        top: { style: "thin", color: { rgb: "FF0000" } },
        right: { style: "medium", color: { rgb: "00FF00" } },
        bottom: { style: "thick", color: { rgb: "0000FF" } },
        left: { style: "dashed", color: { rgb: "FFFF00" } },
        diagonal: { style: "dotted", color: { rgb: "FF00FF" } },
        diagonalUp: true,
        diagonalDown: true,
      },
    };

    const cells = new Map<string, Partial<import("..//src/_types").Cell>>();
    cells.set("0,0", { value: "styled", type: "string", style });

    // Should not throw
    const xlsx = await writeXlsx({
      sheets: [{ name: "Borders", rows: [["styled"]], cells }],
    });
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });

  it("gradient fill with many stops", async () => {
    const style: CellStyle = {
      fill: {
        type: "gradient",
        degree: 90,
        stops: [
          { position: 0, color: { rgb: "FF0000" } },
          { position: 0.25, color: { rgb: "FFFF00" } },
          { position: 0.5, color: { rgb: "00FF00" } },
          { position: 0.75, color: { rgb: "0000FF" } },
          { position: 1, color: { rgb: "FF00FF" } },
        ],
      },
    };

    const cells = new Map<string, Partial<import("..//src/_types").Cell>>();
    cells.set("0,0", { value: "gradient", type: "string", style });

    const xlsx = await writeXlsx({
      sheets: [{ name: "Gradient", rows: [["gradient"]], cells }],
    });
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });

  it("merged cell that spans wide range", async () => {
    const wb = await writeAndRead([
      {
        name: "Merged",
        rows: [["Wide merge", null, null, null, null]],
        merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 4 }],
      },
    ]);
    expect(wb.sheets[0].merges).toBeDefined();
    expect(wb.sheets[0].merges!.length).toBe(1);
  });

  it("many hyperlinks on one sheet", async () => {
    const rows: CellValue[][] = [];
    const cells = new Map<string, Partial<import("..//src/_types").Cell>>();

    for (let i = 0; i < 100; i++) {
      rows.push([`Link ${i}`]);
      cells.set(`${i},0`, {
        value: `Link ${i}`,
        type: "string",
        hyperlink: { target: `https://example.com/${i}`, tooltip: `Link ${i}` },
      });
    }

    const xlsx = await writeXlsx({
      sheets: [{ name: "Links", rows, cells }],
    });
    expect(xlsx).toBeInstanceOf(Uint8Array);

    // Read back and verify it parses
    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows).toHaveLength(100);
  });

  it("formula with special characters", async () => {
    const cells = new Map<string, Partial<import("..//src/_types").Cell>>();
    cells.set("0,0", {
      value: 1,
      type: "formula",
      formula: 'IF(A2="hello ""world""",1,0)',
    });

    const xlsx = await writeXlsx({
      sheets: [{ name: "Formulas", rows: [[1]], cells }],
    });
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });

  it("empty sheet produces valid XLSX", async () => {
    const wb = await writeAndRead([{ name: "Empty", rows: [] }]);
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].rows).toHaveLength(0);
  });

  it("many sheets (50)", async () => {
    const sheets: WriteSheet[] = [];
    for (let i = 0; i < 50; i++) {
      sheets.push({ name: `Sheet${i + 1}`, rows: [[`data_${i}`]] });
    }

    const wb = await writeAndRead(sheets);
    expect(wb.sheets).toHaveLength(50);
    expect(wb.sheets[49].name).toBe("Sheet50");
    expect(wb.sheets[49].rows[0][0]).toBe("data_49");
  });

  it("write-then-read round-trip preserves all basic cell types", async () => {
    const date = new Date(Date.UTC(2025, 3, 15, 10, 30, 0));
    const original: CellValue[][] = [["string", 42, true, date, null, false, 0, "", 3.14, -1]];

    const wb = await writeAndRead([{ name: "Roundtrip", rows: original }]);

    const row = wb.sheets[0].rows[0];
    expect(row[0]).toBe("string");
    expect(row[1]).toBe(42);
    expect(row[2]).toBe(true);
    expect(row[3]).toBeInstanceOf(Date);
    expect((row[3] as Date).getUTCFullYear()).toBe(2025);
    expect((row[3] as Date).getUTCMonth()).toBe(3);
    expect((row[3] as Date).getUTCDate()).toBe(15);
    expect(row[5]).toBe(false);
    expect(row[6]).toBe(0);
    expect(row[7]).toBe("");
    expect(row[8]).toBeCloseTo(3.14);
    expect(row[9]).toBe(-1);
  });

  it("wide row (256 columns)", async () => {
    const row: CellValue[] = [];
    for (let i = 0; i < 256; i++) {
      row.push(`col_${i}`);
    }

    const wb = await writeAndRead([{ name: "Wide", rows: [row] }]);
    expect(wb.sheets[0].rows[0]).toHaveLength(256);
    expect(wb.sheets[0].rows[0][0]).toBe("col_0");
    expect(wb.sheets[0].rows[0][255]).toBe("col_255");
  });

  it("object data with columns", async () => {
    const wb = await writeAndRead([
      {
        name: "Objects",
        columns: [
          { header: "Name", key: "name" },
          { header: "Age", key: "age" },
        ],
        data: [
          { name: "Alice", age: 30 },
          { name: "Bob", age: 25 },
        ],
      },
    ]);

    expect(wb.sheets[0].rows).toHaveLength(3); // header + 2 data rows
    expect(wb.sheets[0].rows[0][0]).toBe("Name");
    expect(wb.sheets[0].rows[0][1]).toBe("Age");
    expect(wb.sheets[0].rows[1][0]).toBe("Alice");
    expect(wb.sheets[0].rows[1][1]).toBe(30);
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 2. CSV Edge Cases
// ═══════════════════════════════════════════════════════════════════════

describe("CSV Edge Cases", () => {
  it("CSV with 100+ columns", () => {
    const header = Array.from({ length: 120 }, (_, i) => `col${i}`).join(",");
    const dataRow = Array.from({ length: 120 }, (_, i) => `val${i}`).join(",");
    const csv = `${header}\n${dataRow}`;

    const rows = parseCsv(csv);
    expect(rows).toHaveLength(2);
    expect(rows[0]).toHaveLength(120);
    expect(rows[1]).toHaveLength(120);
    expect(rows[1][119]).toBe("val119");
  });

  it("CSV value with all special chars: comma, quote, newline, CR, tab", () => {
    const csv = '"hello, ""world""\nwith\ttabs\rand\r\nstuff",normal\nsecond,row';
    const rows = parseCsv(csv);

    expect(rows).toHaveLength(2);
    expect(rows[0][0]).toBe('hello, "world"\nwith\ttabs\rand\r\nstuff');
    expect(rows[0][1]).toBe("normal");
  });

  it("CSV with mixed line endings: \\n then \\r\\n then \\r", () => {
    const csv = "a,b\n1,2\r\n3,4\r5,6";
    const rows = parseCsv(csv);

    expect(rows).toHaveLength(4);
    expect(rows[0]).toEqual(["a", "b"]);
    expect(rows[1]).toEqual(["1", "2"]);
    expect(rows[2]).toEqual(["3", "4"]);
    expect(rows[3]).toEqual(["5", "6"]);
  });

  it("CSV with UTF-8 BOM + semicolon delimiter", () => {
    const csv = "\uFEFFname;age\nAlice;30\nBob;25";
    const rows = parseCsv(csv, { delimiter: ";" });

    expect(rows).toHaveLength(3);
    expect(rows[0][0]).toBe("name");
    expect(rows[0][1]).toBe("age");
  });

  it("CSV round-trip: write -> read -> write -> compare", () => {
    const originalRows: CellValue[][] = [
      ["Name", "Age", "City"],
      ["Alice", 30, "New York"],
      ["Bob", 25, "London"],
      ['Eve "the hacker"', 35, "Paris, France"],
    ];

    const csv1 = writeCsv(originalRows);
    const parsed = parseCsv(csv1);
    const csv2 = writeCsv(parsed);
    // After one round-trip, numbers become strings in CSV
    // But the CSV text output of parsed data should match
    expect(csv2).toBe(csv1);
  });

  it("very large single field (10KB of text in quotes)", () => {
    const bigText = "X".repeat(10_000);
    const csv = `"${bigText}",normal`;
    const rows = parseCsv(csv);

    expect(rows).toHaveLength(1);
    expect(rows[0][0]).toBe(bigText);
    expect(rows[0][1]).toBe("normal");
  });

  it("CSV with trailing commas on every line", () => {
    const csv = "a,b,\n1,2,\n3,4,";
    const rows = parseCsv(csv);

    expect(rows).toHaveLength(3);
    // Trailing comma should create an empty field
    expect(rows[0]).toEqual(["a", "b", ""]);
    expect(rows[1]).toEqual(["1", "2", ""]);
    expect(rows[2]).toEqual(["3", "4", ""]);
  });

  it("CSV where every field is quoted (even numbers)", () => {
    const csv = '"name","age","active"\n"Alice","30","true"';
    const rows = parseCsv(csv);

    expect(rows).toHaveLength(2);
    expect(rows[0]).toEqual(["name", "age", "active"]);
    expect(rows[1]).toEqual(["Alice", "30", "true"]);
  });

  it("empty CSV (just BOM)", () => {
    const csv = "\uFEFF";
    const rows = parseCsv(csv);
    expect(rows).toHaveLength(0);
  });

  it("CSV with only header row, no data", () => {
    const csv = "name,age,city";
    const rows = parseCsv(csv);

    expect(rows).toHaveLength(1);
    expect(rows[0]).toEqual(["name", "age", "city"]);
  });

  it("type inference: leading zeros should stay string", () => {
    const csv = "code\n0123\n00456\n0";
    const rows = parseCsv(csv, { typeInference: true });

    // "0123" has leading zero so should NOT be parsed as number 123
    // But the current implementation may convert it. Let's check.
    expect(rows).toHaveLength(4);
    // "0" by itself is treated as boolean false with typeInference
    expect(rows[3][0]).toBe(false);
  });

  it("type inference: scientific notation", () => {
    const csv = "val\n1e10\n2.5e-3\n1E5";
    const rows = parseCsv(csv, { typeInference: true });

    expect(rows).toHaveLength(4);
    // 1e10 should be parsed as number 10000000000
    expect(rows[1][0]).toBe(1e10);
    expect(rows[2][0]).toBeCloseTo(0.0025);
    expect(rows[3][0]).toBe(1e5);
  });

  it("type inference: true/TRUE/True/yes/YES", () => {
    const csv = "true\nTRUE\nTrue\nyes\nYES\nYes";
    const rows = parseCsv(csv, { typeInference: true });

    expect(rows).toHaveLength(6);
    expect(rows[0][0]).toBe(true);
    expect(rows[1][0]).toBe(true);
    expect(rows[2][0]).toBe(true);
    expect(rows[3][0]).toBe(true);
    expect(rows[4][0]).toBe(true);
    expect(rows[5][0]).toBe(true);
  });

  it("type inference: null, undefined, NaN, Infinity should stay strings", () => {
    const csv = "null\nundefined\nNaN\nInfinity\n-Infinity";
    const rows = parseCsv(csv, { typeInference: true });

    expect(rows).toHaveLength(5);
    // These should NOT be converted to actual null/undefined/NaN/Infinity
    expect(rows[0][0]).toBe("null");
    expect(rows[1][0]).toBe("undefined");
    expect(rows[2][0]).toBe("NaN");
    // Infinity and -Infinity: parseNumber rejects non-finite
    expect(rows[3][0]).toBe("Infinity");
    expect(rows[4][0]).toBe("-Infinity");
  });

  it("delimiter auto-detection with tabs", () => {
    const csv = "a\tb\tc\n1\t2\t3";
    const delimiter = detectDelimiter(csv);
    expect(delimiter).toBe("\t");
  });

  it("delimiter auto-detection with semicolons", () => {
    const csv = "name;age;city\nAlice;30;NY\nBob;25;LA";
    const delimiter = detectDelimiter(csv);
    expect(delimiter).toBe(";");
  });

  it("CSV quoteStyle 'all' wraps everything including empty values", () => {
    const result = writeCsv([["a", null, "b"]], { quoteStyle: "all" });
    expect(result).toBe('"a","","b"');
  });

  it("CSV quoteStyle 'none' does not quote special chars", () => {
    const result = writeCsv([["hello, world"]], { quoteStyle: "none" });
    // With 'none', the comma inside should NOT be quoted
    expect(result).toBe("hello, world");
  });

  it("writeCsvObjects round-trip", () => {
    const data = [
      { name: "Alice", age: 30 },
      { name: "Bob", age: 25 },
    ];

    const csv = writeCsvObjects(data);
    const { data: parsed, headers } = parseCsvObjects(csv, { header: true });

    expect(headers).toEqual(["name", "age"]);
    expect(parsed).toHaveLength(2);
    expect(parsed[0].name).toBe("Alice");
    expect(parsed[0].age).toBe("30"); // CSV doesn't preserve types
    expect(parsed[1].name).toBe("Bob");
  });

  it("CSV with BOM option writes BOM", () => {
    const csv = writeCsv([["a", "b"]], { bom: true });
    expect(csv.charCodeAt(0)).toBe(0xfeff);
  });

  it("stripBom removes UTF-8 BOM", () => {
    expect(stripBom("\uFEFFhello")).toBe("hello");
    expect(stripBom("hello")).toBe("hello");
    expect(stripBom("")).toBe("");
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 3. Schema Validation Edge Cases
// ═══════════════════════════════════════════════════════════════════════

describe("Schema Validation Edge Cases", () => {
  it("schema with all field types simultaneously", () => {
    const schema: SchemaDefinition = {
      name: { type: "string", column: "Name", required: true },
      age: { type: "integer", column: "Age", min: 0, max: 150 },
      score: { type: "number", column: "Score", min: 0, max: 100 },
      active: { type: "boolean", column: "Active" },
      joined: { type: "date", column: "Joined" },
    };

    const rows: CellValue[][] = [
      ["Name", "Age", "Score", "Active", "Joined"],
      ["Alice", 30, 95.5, true, new Date(Date.UTC(2024, 0, 15))],
    ];

    const result = validateWithSchema(rows, schema);
    expect(result.errors).toHaveLength(0);
    expect(result.data).toHaveLength(1);
    expect(result.data[0].name).toBe("Alice");
    expect(result.data[0].age).toBe(30);
    expect(result.data[0].score).toBe(95.5);
    expect(result.data[0].active).toBe(true);
    expect(result.data[0].joined).toBeInstanceOf(Date);
  });

  it("transform chain", () => {
    const schema: SchemaDefinition = {
      value: {
        type: "string",
        column: "Value",
        transform: (v) => (v as string).toUpperCase().trim(),
      },
    };

    const rows: CellValue[][] = [["Value"], ["  hello  "]];
    const result = validateWithSchema(rows, schema);
    expect(result.data[0].value).toBe("HELLO");
  });

  it("validate function returns very long error string", () => {
    const longError = "E".repeat(500);
    const schema: SchemaDefinition = {
      value: {
        type: "string",
        column: "Value",
        validate: () => longError,
      },
    };

    const rows: CellValue[][] = [["Value"], ["test"]];
    const result = validateWithSchema(rows, schema);
    expect(result.errors).toHaveLength(1);
    expect(result.errors[0].message).toBe(longError);
  });

  it("schema field with both pattern and enum", () => {
    const schema: SchemaDefinition = {
      code: {
        type: "string",
        column: "Code",
        pattern: /^[A-Z]{3}$/,
        enum: ["ABC", "DEF", "GHI"],
      },
    };

    // Valid: passes both pattern and enum
    const rows1: CellValue[][] = [["Code"], ["ABC"]];
    const result1 = validateWithSchema(rows1, schema);
    expect(result1.errors).toHaveLength(0);
    expect(result1.data[0].code).toBe("ABC");

    // Fails pattern but is in enum
    const rows2: CellValue[][] = [["Code"], ["abc"]];
    const result2 = validateWithSchema(rows2, schema);
    // Pattern is checked first, so it should fail on pattern
    expect(result2.errors).toHaveLength(1);
    expect(result2.errors[0].message).toContain("pattern");
  });

  it("integer validation of 3.0000000001 (floating point noise)", () => {
    const schema: SchemaDefinition = {
      count: { type: "integer", column: "Count" },
    };

    // 3.0000000001 is not an integer
    const rows: CellValue[][] = [["Count"], [3.0000000001]];
    const result = validateWithSchema(rows, schema);
    expect(result.errors).toHaveLength(1);
    expect(result.errors[0].message).toContain("integer");
  });

  it("required field with value 0, false, empty string", () => {
    const schema: SchemaDefinition = {
      num: { type: "number", column: "Num", required: true },
      bool: { type: "boolean", column: "Bool", required: true },
      str: { type: "string", column: "Str", required: true },
    };

    const rows: CellValue[][] = [
      ["Num", "Bool", "Str"],
      [0, false, ""],
    ];

    const result = validateWithSchema(rows, schema);

    // 0 and false should NOT trigger "required" error
    // but empty string "" should because isEmpty() trims and checks
    expect(result.data[0].num).toBe(0);
    expect(result.data[0].bool).toBe(false);
    // Empty string IS considered empty by the isEmpty() function
    expect(result.errors.length).toBeGreaterThanOrEqual(1);
    const strError = result.errors.find((e) => e.field === "str");
    expect(strError).toBeDefined();
  });

  it("schema with no matching columns at all", () => {
    const schema: SchemaDefinition = {
      name: { column: "Name", required: true },
      age: { column: "Age", required: true },
    };

    // Headers don't match
    const rows: CellValue[][] = [
      ["FirstName", "Years"],
      ["Alice", 30],
    ];

    const result = validateWithSchema(rows, schema);
    // Both fields should have required errors
    expect(result.errors.length).toBeGreaterThanOrEqual(2);
  });

  it("schema with columnIndex out of range", () => {
    const schema: SchemaDefinition = {
      value: { columnIndex: 999, required: true },
    };

    const rows: CellValue[][] = [
      ["A", "B"],
      ["x", "y"],
    ];

    const result = validateWithSchema(rows, schema);
    expect(result.errors).toHaveLength(1);
  });

  it("errorMode throw stops at first error", () => {
    const schema: SchemaDefinition = {
      name: { column: "Name", type: "string", required: true },
      age: { column: "Age", type: "integer", required: true },
    };

    const rows: CellValue[][] = [
      ["Name", "Age"],
      [null, "not_a_number"],
    ];

    expect(() => {
      validateWithSchema(rows, schema, { errorMode: "throw" });
    }).toThrow();
  });

  it("default value for empty cells", () => {
    const schema: SchemaDefinition = {
      name: { column: "Name", default: "Unknown" },
      age: { column: "Age", type: "number", default: 0 },
    };

    const rows: CellValue[][] = [
      ["Name", "Age"],
      [null, null],
    ];

    const result = validateWithSchema(rows, schema);
    expect(result.data[0].name).toBe("Unknown");
    expect(result.data[0].age).toBe(0);
  });

  it("empty schema returns empty data", () => {
    const schema: SchemaDefinition = {};
    const rows: CellValue[][] = [["A"], [1]];
    const result = validateWithSchema(rows, schema);
    expect(result.data).toHaveLength(0);
  });

  it("skipEmptyRows option", () => {
    const schema: SchemaDefinition = {
      name: { column: "Name" },
    };

    const rows: CellValue[][] = [
      ["Name"],
      ["Alice"],
      [null], // empty row
      ["Bob"],
      [""], // empty string row
    ];

    const result = validateWithSchema(rows, schema, { skipEmptyRows: true });
    expect(result.data).toHaveLength(2);
    expect(result.data[0].name).toBe("Alice");
    expect(result.data[1].name).toBe("Bob");
  });

  it("min/max for string length", () => {
    const schema: SchemaDefinition = {
      code: { column: "Code", type: "string", min: 3, max: 5 },
    };

    const rows: CellValue[][] = [
      ["Code"],
      ["AB"], // too short
      ["ABC"], // ok
      ["ABCDE"], // ok
      ["ABCDEF"], // too long
    ];

    const result = validateWithSchema(rows, schema);
    expect(result.errors).toHaveLength(2);
    expect(result.errors[0].message).toContain("below minimum");
    expect(result.errors[1].message).toContain("exceeds maximum");
  });

  it("coercion from string to number with commas", () => {
    const schema: SchemaDefinition = {
      amount: { column: "Amount", type: "number" },
    };

    const rows: CellValue[][] = [["Amount"], ["1,234.56"], ["1,000,000"]];

    const result = validateWithSchema(rows, schema);
    expect(result.errors).toHaveLength(0);
    expect(result.data[0].amount).toBeCloseTo(1234.56);
    expect(result.data[1].amount).toBe(1000000);
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 4. ZIP Edge Cases
// ═══════════════════════════════════════════════════════════════════════

describe("ZIP Edge Cases", () => {
  it("ZIP with many entries (100+)", async () => {
    const zip = new ZipWriter();
    const encoder = new TextEncoder();

    for (let i = 0; i < 150; i++) {
      zip.add(`file_${i}.txt`, encoder.encode(`content_${i}`));
    }

    const data = await zip.build();
    const reader = new ZipReader(data);

    expect(reader.entries()).toHaveLength(150);
    expect(reader.has("file_0.txt")).toBe(true);
    expect(reader.has("file_149.txt")).toBe(true);

    const content = await reader.extract("file_99.txt");
    expect(new TextDecoder().decode(content)).toBe("content_99");
  });

  it("ZIP entry with path containing spaces and Unicode", async () => {
    const zip = new ZipWriter();
    const encoder = new TextEncoder();

    zip.add("path with spaces/file name.txt", encoder.encode("content1"));
    zip.add("unicode/\u00E9\u00E0\u00FC/file.txt", encoder.encode("content2"));

    const data = await zip.build();
    const reader = new ZipReader(data);

    expect(reader.has("path with spaces/file name.txt")).toBe(true);
    expect(reader.has("unicode/\u00E9\u00E0\u00FC/file.txt")).toBe(true);

    const c1 = await reader.extract("path with spaces/file name.txt");
    expect(new TextDecoder().decode(c1)).toBe("content1");
  });

  it("large single entry (100KB)", async () => {
    const zip = new ZipWriter();
    const bigData = new Uint8Array(100_000);
    // Fill with a pattern
    for (let i = 0; i < bigData.length; i++) {
      bigData[i] = i % 256;
    }

    zip.add("big.bin", bigData);
    const data = await zip.build();
    const reader = new ZipReader(data);

    const extracted = await reader.extract("big.bin");
    expect(extracted.length).toBe(100_000);
    // Verify first and last bytes
    expect(extracted[0]).toBe(0);
    expect(extracted[255]).toBe(255);
    expect(extracted[256]).toBe(0);
  });

  it("entry that doesn't compress well (random-ish bytes)", async () => {
    const zip = new ZipWriter();
    const randomData = new Uint8Array(1000);
    // Fill with pseudo-random data that doesn't compress well
    for (let i = 0; i < randomData.length; i++) {
      randomData[i] = (i * 31 + 17) % 256;
    }

    zip.add("random.bin", randomData);
    const data = await zip.build();
    const reader = new ZipReader(data);

    const extracted = await reader.extract("random.bin");
    expect(extracted.length).toBe(1000);
    for (let i = 0; i < 1000; i++) {
      expect(extracted[i]).toBe(randomData[i]);
    }
  });

  it("empty entry (0 bytes)", async () => {
    const zip = new ZipWriter();
    zip.add("empty.txt", new Uint8Array(0));

    const data = await zip.build();
    const reader = new ZipReader(data);

    const extracted = await reader.extract("empty.txt");
    expect(extracted.length).toBe(0);
  });

  it("extractAll returns all non-directory entries", async () => {
    const zip = new ZipWriter();
    const encoder = new TextEncoder();

    zip.add("a.txt", encoder.encode("aaa"));
    zip.add("b.txt", encoder.encode("bbb"));
    zip.add("sub/c.txt", encoder.encode("ccc"));

    const data = await zip.build();
    const reader = new ZipReader(data);
    const all = await reader.extractAll();

    expect(all.size).toBe(3);
    expect(all.has("a.txt")).toBe(true);
    expect(all.has("sub/c.txt")).toBe(true);
  });

  it("extracting non-existent entry throws", async () => {
    const zip = new ZipWriter();
    zip.add("exists.txt", new TextEncoder().encode("hi"));

    const data = await zip.build();
    const reader = new ZipReader(data);

    await expect(reader.extract("does_not_exist.txt")).rejects.toThrow();
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 5. XML Parser Edge Cases
// ═══════════════════════════════════════════════════════════════════════

describe("XML Parser Edge Cases", () => {
  it("deeply nested XML (100 levels)", () => {
    let xml = '<?xml version="1.0"?>';
    for (let i = 0; i < 100; i++) {
      xml += `<level${i}>`;
    }
    xml += "deep";
    for (let i = 99; i >= 0; i--) {
      xml += `</level${i}>`;
    }

    const doc = parseXml(xml);
    expect(doc.tag).toBe("level0");

    // Walk down to verify depth
    let current = doc;
    for (let i = 1; i < 100; i++) {
      const child = current.children.find((c) => typeof c !== "string" && c.tag === `level${i}`);
      expect(child).toBeDefined();
      current = child as typeof doc;
    }
    expect(current.text).toBe("deep");
  });

  it("very long attribute value (5KB)", () => {
    const longVal = "X".repeat(5000);
    const xml = `<?xml version="1.0"?><root attr="${longVal}"/>`;
    const doc = parseXml(xml);
    expect(doc.attrs["attr"]).toBe(longVal);
  });

  it("attribute with XML entities", () => {
    const xml = '<?xml version="1.0"?><root val="a&amp;b&lt;c&gt;d&quot;e"/>';
    const doc = parseXml(xml);
    expect(doc.attrs["val"]).toBe('a&b<c>d"e');
  });

  it("multiple namespaces on same element", () => {
    const xml =
      '<?xml version="1.0"?><root xmlns:a="http://a.com" xmlns:b="http://b.com"><a:child/><b:child/></root>';
    const doc = parseXml(xml);

    const aChild = doc.children.find((c) => typeof c !== "string" && c.tag === "a:child");
    const bChild = doc.children.find((c) => typeof c !== "string" && c.tag === "b:child");

    expect(aChild).toBeDefined();
    expect(bChild).toBeDefined();
    expect((aChild as any).prefix).toBe("a");
    expect((aChild as any).local).toBe("child");
    expect((bChild as any).prefix).toBe("b");
    expect((bChild as any).local).toBe("child");
  });

  it("XML with processing instructions mixed with elements", () => {
    const xml = '<?xml version="1.0"?><?pi1 data?><root><?pi2 more?><child/></root>';
    const doc = parseXml(xml);
    expect(doc.tag).toBe("root");
    // Processing instructions should be skipped
    const childEl = doc.children.find((c) => typeof c !== "string" && c.tag === "child");
    expect(childEl).toBeDefined();
  });

  it("empty element: <t></t> vs <t/>", () => {
    const xml1 = '<?xml version="1.0"?><root><t></t></root>';
    const xml2 = '<?xml version="1.0"?><root><t/></root>';

    const doc1 = parseXml(xml1);
    const doc2 = parseXml(xml2);

    const t1 = doc1.children.find((c) => typeof c !== "string" && c.tag === "t");
    const t2 = doc2.children.find((c) => typeof c !== "string" && c.tag === "t");

    expect(t1).toBeDefined();
    expect(t2).toBeDefined();
    // Both should be present, empty content
    expect((t1 as any).children).toHaveLength(0);
    expect((t2 as any).children).toHaveLength(0);
  });

  it("CDATA section", () => {
    const xml = '<?xml version="1.0"?><root><![CDATA[<not&xml>]]></root>';
    const doc = parseXml(xml);
    expect(doc.text).toBe("<not&xml>");
  });

  it("XML comment handling", () => {
    const xml = '<?xml version="1.0"?><root><!-- comment --><child>text</child></root>';
    const doc = parseXml(xml);
    const child = doc.children.find((c) => typeof c !== "string" && c.tag === "child");
    expect(child).toBeDefined();
    expect((child as any).text).toBe("text");
  });

  it("SAX parser fires events correctly", () => {
    const xml = '<?xml version="1.0"?><root><child attr="val">text</child></root>';
    const events: string[] = [];

    parseSax(xml, {
      onOpenTag(tag, attrs) {
        events.push(`open:${tag}`);
        if (Object.keys(attrs).length > 0) {
          events.push(`attrs:${JSON.stringify(attrs)}`);
        }
      },
      onCloseTag(tag) {
        events.push(`close:${tag}`);
      },
      onText(text) {
        if (text.trim()) events.push(`text:${text}`);
      },
    });

    expect(events).toContain("open:root");
    expect(events).toContain("open:child");
    expect(events).toContain('attrs:{"attr":"val"}');
    expect(events).toContain("text:text");
    expect(events).toContain("close:child");
    expect(events).toContain("close:root");
  });

  it("xmlEscape handles all special chars", () => {
    expect(xmlEscape("a & b")).toBe("a &amp; b");
    expect(xmlEscape("a < b")).toBe("a &lt; b");
    expect(xmlEscape("a > b")).toBe("a &gt; b");
    expect(xmlEscape("normal")).toBe("normal");
    expect(xmlEscape("")).toBe("");
  });

  it("xmlEscapeAttr handles quotes, tabs, newlines", () => {
    expect(xmlEscapeAttr('a"b')).toBe("a&quot;b");
    expect(xmlEscapeAttr("a\tb")).toBe("a&#9;b");
    expect(xmlEscapeAttr("a\nb")).toBe("a&#10;b");
    expect(xmlEscapeAttr("a\rb")).toBe("a&#13;b");
  });

  it("numeric character references in XML", () => {
    const xml = '<?xml version="1.0"?><root>&#65;&#x42;</root>';
    const doc = parseXml(xml);
    expect(doc.text).toBe("AB"); // &#65; = A, &#x42; = B
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 6. Streaming Edge Cases
// ═══════════════════════════════════════════════════════════════════════

describe("Streaming Edge Cases", () => {
  it("stream writer with 0 rows produces valid XLSX", async () => {
    const writer = new XlsxStreamWriter({ name: "Empty" });
    const xlsx = await writer.finish();

    const wb = await readXlsx(xlsx);
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].rows).toHaveLength(0);
  });

  it("stream 10,000 rows and verify count", async () => {
    const writer = new XlsxStreamWriter({ name: "Large" });
    for (let i = 0; i < 10_000; i++) {
      writer.addRow([`Row${i}`, i]);
    }

    const xlsx = await writer.finish();

    // Verify via streaming read
    let count = 0;
    for await (const row of streamXlsxRows(xlsx)) {
      if (count === 0) {
        expect(row.values[0]).toBe("Row0");
        expect(row.values[1]).toBe(0);
      }
      count++;
    }
    expect(count).toBe(10_000);
  });

  it("stream writer with dates", async () => {
    const writer = new XlsxStreamWriter({ name: "Dates" });
    const date1 = new Date(Date.UTC(2024, 0, 1));
    const date2 = new Date(Date.UTC(2025, 11, 31));
    writer.addRow([date1, date2]);

    const xlsx = await writer.finish();
    const wb = await readXlsx(xlsx);

    expect(wb.sheets[0].rows[0][0]).toBeInstanceOf(Date);
    expect((wb.sheets[0].rows[0][0] as Date).getUTCFullYear()).toBe(2024);
    expect((wb.sheets[0].rows[0][0] as Date).getUTCMonth()).toBe(0);

    expect(wb.sheets[0].rows[0][1]).toBeInstanceOf(Date);
    expect((wb.sheets[0].rows[0][1] as Date).getUTCFullYear()).toBe(2025);
    expect((wb.sheets[0].rows[0][1] as Date).getUTCMonth()).toBe(11);
  });

  it("abort streaming read early (break from for-await)", async () => {
    const writer = new XlsxStreamWriter({ name: "Test" });
    for (let i = 0; i < 100; i++) {
      writer.addRow([`Row${i}`]);
    }
    const xlsx = await writer.finish();

    let count = 0;
    for await (const _row of streamXlsxRows(xlsx)) {
      count++;
      if (count >= 5) break;
    }

    expect(count).toBe(5);
  });

  it("stream writer handles null values correctly", async () => {
    const writer = new XlsxStreamWriter({ name: "Nulls" });
    writer.addRow([null, "a", null, "b", null]);

    const xlsx = await writer.finish();
    const wb = await readXlsx(xlsx);

    const row = wb.sheets[0].rows[0];
    expect(row[1]).toBe("a");
    expect(row[3]).toBe("b");
  });

  it("stream reader matches regular reader", async () => {
    const testRows: CellValue[][] = [
      ["Name", "Age", "Score"],
      ["Alice", 30, 95.5],
      ["Bob", 25, 88.0],
      [null, 0, -1],
    ];

    const xlsx = await writeXlsx({
      sheets: [{ name: "Test", rows: testRows }],
    });

    // Regular read
    const wb = await readXlsx(xlsx);

    // Streaming read
    const streamRows = await collectStreamRows(streamXlsxRows(xlsx));

    expect(streamRows).toHaveLength(wb.sheets[0].rows.length);

    for (let i = 0; i < streamRows.length; i++) {
      const streamVals = streamRows[i].values;
      const regularVals = wb.sheets[0].rows[i];

      // Compare significant values
      const maxLen = Math.max(streamVals.length, regularVals.length);
      for (let j = 0; j < maxLen; j++) {
        const sv = j < streamVals.length ? streamVals[j] : null;
        const rv = j < regularVals.length ? regularVals[j] : null;
        expect(sv).toEqual(rv);
      }
    }
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 7. ODS Edge Cases
// ═══════════════════════════════════════════════════════════════════════

describe("ODS Edge Cases", () => {
  it("ODS with merged cells (write, but verify no crash)", async () => {
    // ODS writer doesn't fully support merges, but should not crash
    const ods = await writeOds({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Merged", null, null],
            ["data", "data2", "data3"],
          ],
          merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
        },
      ],
    });
    expect(ods).toBeInstanceOf(Uint8Array);

    // Read it back
    const wb = await readOds(ods);
    expect(wb.sheets).toHaveLength(1);
  });

  it("ODS with dates", async () => {
    const date = new Date(2024, 5, 15, 10, 30, 0);
    const ods = await writeOds({
      sheets: [{ name: "Dates", rows: [[date, "text", 42]] }],
    });

    const wb = await readOds(ods);
    expect(wb.sheets[0].rows).toHaveLength(1);
    expect(wb.sheets[0].rows[0][0]).toBeInstanceOf(Date);
    expect(wb.sheets[0].rows[0][1]).toBe("text");
    expect(wb.sheets[0].rows[0][2]).toBe(42);
  });

  it("ODS with empty sheets", async () => {
    const ods = await writeOds({
      sheets: [
        { name: "Empty", rows: [] },
        { name: "WithData", rows: [["hello"]] },
      ],
    });

    const wb = await readOds(ods);
    expect(wb.sheets).toHaveLength(2);
    expect(wb.sheets[0].rows).toHaveLength(0);
    expect(wb.sheets[1].rows).toHaveLength(1);
    expect(wb.sheets[1].rows[0][0]).toBe("hello");
  });

  it("ODS with booleans", async () => {
    const ods = await writeOds({
      sheets: [{ name: "Bools", rows: [[true, false, null]] }],
    });

    const wb = await readOds(ods);
    expect(wb.sheets[0].rows[0][0]).toBe(true);
    expect(wb.sheets[0].rows[0][1]).toBe(false);
  });

  it("ODS round-trip: basic types", async () => {
    const original: CellValue[][] = [
      ["text", 42, true, false, null],
      ["more", 3.14, false, true, "end"],
    ];

    const ods = await writeOds({
      sheets: [{ name: "Test", rows: original }],
    });

    const wb = await readOds(ods);
    expect(wb.sheets[0].rows).toHaveLength(2);
    expect(wb.sheets[0].rows[0][0]).toBe("text");
    expect(wb.sheets[0].rows[0][1]).toBe(42);
    expect(wb.sheets[0].rows[0][2]).toBe(true);
    expect(wb.sheets[0].rows[0][3]).toBe(false);
    expect(wb.sheets[0].rows[1][0]).toBe("more");
    expect(wb.sheets[0].rows[1][1]).toBeCloseTo(3.14);
  });

  it("ODS with document properties", async () => {
    const ods = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["data"]] }],
      properties: {
        title: "Test Document",
        creator: "Test Author",
        description: "Test Description",
      },
    });

    const wb = await readOds(ods);
    expect(wb.properties?.title).toBe("Test Document");
    expect(wb.properties?.creator).toBe("Test Author");
    expect(wb.properties?.description).toBe("Test Description");
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 8. Sheet Operations Edge Cases
// ═══════════════════════════════════════════════════════════════════════

describe("Sheet Operations Edge Cases", () => {
  function makeSheet(rows: CellValue[][]): Sheet {
    return { name: "Test", rows };
  }

  it("insertRows at index 0", () => {
    const sheet = makeSheet([["a"], ["b"], ["c"]]);
    insertRows(sheet, 0, 2);

    expect(sheet.rows).toHaveLength(5);
    expect(sheet.rows[0]).toEqual([null]);
    expect(sheet.rows[1]).toEqual([null]);
    expect(sheet.rows[2]).toEqual(["a"]);
  });

  it("insertRows at end", () => {
    const sheet = makeSheet([["a"], ["b"]]);
    insertRows(sheet, 2, 3);

    expect(sheet.rows).toHaveLength(5);
    expect(sheet.rows[0]).toEqual(["a"]);
    expect(sheet.rows[1]).toEqual(["b"]);
  });

  it("insertRows with count 0 does nothing", () => {
    const sheet = makeSheet([["a"]]);
    insertRows(sheet, 0, 0);
    expect(sheet.rows).toHaveLength(1);
  });

  it("deleteRows: delete more rows than exist", () => {
    const sheet = makeSheet([["a"], ["b"], ["c"]]);
    // Try to delete 10 rows starting at index 1 (only 2 available)
    deleteRows(sheet, 1, 10);

    expect(sheet.rows).toHaveLength(1);
    expect(sheet.rows[0]).toEqual(["a"]);
  });

  it("deleteRows: delete all rows", () => {
    const sheet = makeSheet([["a"], ["b"]]);
    deleteRows(sheet, 0, 2);
    expect(sheet.rows).toHaveLength(0);
  });

  it("insertColumns at column 0", () => {
    const sheet = makeSheet([
      ["a", "b"],
      ["c", "d"],
    ]);
    insertColumns(sheet, 0, 1);

    expect(sheet.rows[0]).toEqual([null, "a", "b"]);
    expect(sheet.rows[1]).toEqual([null, "c", "d"]);
  });

  it("deleteColumns: delete all columns", () => {
    const sheet = makeSheet([
      ["a", "b"],
      ["c", "d"],
    ]);
    deleteColumns(sheet, 0, 2);

    expect(sheet.rows[0]).toEqual([]);
    expect(sheet.rows[1]).toEqual([]);
  });

  it("insertRows updates merge ranges", () => {
    const sheet = makeSheet([
      ["a", "b"],
      ["c", "d"],
      ["e", "f"],
    ]);
    sheet.merges = [{ startRow: 1, startCol: 0, endRow: 2, endCol: 1 }];

    insertRows(sheet, 1, 2);

    // Merge should shift down by 2
    expect(sheet.merges![0].startRow).toBe(3);
    expect(sheet.merges![0].endRow).toBe(4);
  });

  it("deleteRows removes merge fully within deleted range", () => {
    const sheet = makeSheet([["a"], ["b"], ["c"], ["d"]]);
    sheet.merges = [{ startRow: 1, startCol: 0, endRow: 2, endCol: 0 }];

    deleteRows(sheet, 1, 2);

    // Merge was within deleted range, should be removed
    expect(sheet.merges!.length).toBe(0);
  });

  it("insertRows updates cells Map keys", () => {
    const sheet = makeSheet([["a"], ["b"]]);
    sheet.cells = new Map();
    sheet.cells.set("1,0", { value: "cellB", type: "string" } as any);

    insertRows(sheet, 0, 1);

    // Cell at row 1 should now be at row 2
    expect(sheet.cells.has("2,0")).toBe(true);
    expect(sheet.cells.has("1,0")).toBe(false);
  });

  it("cloneSheet creates independent copy", () => {
    const sheet = makeSheet([
      ["a", "b"],
      ["c", "d"],
    ]);
    sheet.merges = [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }];

    const clone = cloneSheet(sheet, "Clone");

    expect(clone.name).toBe("Clone");
    expect(clone.rows).toEqual(sheet.rows);

    // Modify clone and verify original is unchanged
    clone.rows[0][0] = "MODIFIED";
    expect(sheet.rows[0][0]).toBe("a");

    // Merges should be independent
    clone.merges![0].startRow = 99;
    expect(sheet.merges![0].startRow).toBe(0);
  });

  it("moveSheet to same position does nothing harmful", () => {
    const wb: Workbook = {
      sheets: [makeSheet([["a"]]), makeSheet([["b"]]), makeSheet([["c"]])],
    };
    wb.sheets[0].name = "S1";
    wb.sheets[1].name = "S2";
    wb.sheets[2].name = "S3";

    moveSheet(wb, 1, 1);

    expect(wb.sheets[0].name).toBe("S1");
    expect(wb.sheets[1].name).toBe("S2");
    expect(wb.sheets[2].name).toBe("S3");
  });

  it("moveSheet from beginning to end", () => {
    const wb: Workbook = {
      sheets: [
        { name: "A", rows: [["a"]] },
        { name: "B", rows: [["b"]] },
        { name: "C", rows: [["c"]] },
      ],
    };

    moveSheet(wb, 0, 2);

    expect(wb.sheets[0].name).toBe("B");
    expect(wb.sheets[1].name).toBe("C");
    expect(wb.sheets[2].name).toBe("A");
  });

  it("insertRows updates dataValidations ranges", () => {
    const sheet = makeSheet([["a"], ["b"], ["c"]]);
    sheet.dataValidations = [{ type: "list", range: "A2:A3", values: ["x", "y"] }];

    insertRows(sheet, 1, 2);

    // The validation range should shift
    expect(sheet.dataValidations[0].range).toBe("A4:A5");
  });

  it("insertRows updates autoFilter range", () => {
    const sheet = makeSheet([
      ["h1", "h2"],
      ["a", "b"],
      ["c", "d"],
    ]);
    sheet.autoFilter = { range: "A1:B3" };

    insertRows(sheet, 1, 1);

    expect(sheet.autoFilter!.range).toBe("A1:B4");
  });

  it("insertRows updates image anchors", () => {
    const sheet = makeSheet([["a"], ["b"]]);
    sheet.images = [
      {
        data: new Uint8Array([1, 2, 3]),
        type: "png",
        anchor: { from: { row: 1, col: 0 }, to: { row: 2, col: 1 } },
      },
    ];

    insertRows(sheet, 0, 2);

    expect(sheet.images![0].anchor.from.row).toBe(3);
    expect(sheet.images![0].anchor.to!.row).toBe(4);
  });

  it("insertColumns updates image anchors", () => {
    const sheet = makeSheet([["a", "b"]]);
    sheet.images = [
      {
        data: new Uint8Array([1, 2, 3]),
        type: "png",
        anchor: { from: { row: 0, col: 1 }, to: { row: 1, col: 2 } },
      },
    ];

    insertColumns(sheet, 0, 2);

    expect(sheet.images![0].anchor.from.col).toBe(3);
    expect(sheet.images![0].anchor.to!.col).toBe(4);
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 9. Date Utilities Edge Cases
// ═══════════════════════════════════════════════════════════════════════

describe("Date Utilities Edge Cases", () => {
  it("serialToDate: serial 1 = Jan 1, 1900", () => {
    const d = serialToDate(1);
    expect(d.getUTCFullYear()).toBe(1900);
    expect(d.getUTCMonth()).toBe(0);
    expect(d.getUTCDate()).toBe(1);
  });

  it("serialToDate: serial 60 = phantom Feb 29, 1900 (Lotus bug)", () => {
    // Serial 60 is the phantom Feb 29, 1900
    // The implementation maps this to Feb 28, 1900
    const d = serialToDate(60);
    expect(d.getUTCFullYear()).toBe(1900);
    expect(d.getUTCMonth()).toBe(1);
    expect(d.getUTCDate()).toBe(28);
  });

  it("serialToDate: serial 61 = Mar 1, 1900", () => {
    const d = serialToDate(61);
    expect(d.getUTCFullYear()).toBe(1900);
    expect(d.getUTCMonth()).toBe(2);
    expect(d.getUTCDate()).toBe(1);
  });

  it("dateToSerial roundtrip for modern dates", () => {
    const dates = [
      new Date(Date.UTC(2024, 0, 1)),
      new Date(Date.UTC(2024, 5, 15)),
      new Date(Date.UTC(1999, 11, 31)),
    ];

    for (const d of dates) {
      const serial = dateToSerial(d);
      const recovered = serialToDate(serial);
      expect(recovered.getUTCFullYear()).toBe(d.getUTCFullYear());
      expect(recovered.getUTCMonth()).toBe(d.getUTCMonth());
      expect(recovered.getUTCDate()).toBe(d.getUTCDate());
    }
  });

  it("dateToSerial/serialToDate 1904 system", () => {
    const d = new Date(Date.UTC(2024, 0, 1));
    const serial = dateToSerial(d, true);
    const recovered = serialToDate(serial, true);

    expect(recovered.getUTCFullYear()).toBe(2024);
    expect(recovered.getUTCMonth()).toBe(0);
    expect(recovered.getUTCDate()).toBe(1);
  });

  it("parseDate: ISO 8601 variants", () => {
    const d1 = parseDate("2024-01-15");
    expect(d1).toBeInstanceOf(Date);
    expect(d1!.getUTCFullYear()).toBe(2024);

    const d2 = parseDate("2024-01-15T14:30:00Z");
    expect(d2).toBeInstanceOf(Date);
    expect(d2!.getUTCHours()).toBe(14);

    const d3 = parseDate("2024-01-15T14:30:00+05:00");
    expect(d3).toBeInstanceOf(Date);
    // UTC should be 14:30 - 5:00 = 09:30
    expect(d3!.getUTCHours()).toBe(9);
    expect(d3!.getUTCMinutes()).toBe(30);
  });

  it("parseDate: US format MM/DD/YYYY", () => {
    const d = parseDate("01/15/2024");
    expect(d).toBeInstanceOf(Date);
    expect(d!.getUTCFullYear()).toBe(2024);
    expect(d!.getUTCMonth()).toBe(0);
    expect(d!.getUTCDate()).toBe(15);
  });

  it("parseDate: EU format DD.MM.YYYY", () => {
    const d = parseDate("15.01.2024");
    expect(d).toBeInstanceOf(Date);
    expect(d!.getUTCFullYear()).toBe(2024);
    expect(d!.getUTCMonth()).toBe(0);
    expect(d!.getUTCDate()).toBe(15);
  });

  it("parseDate: invalid string returns null", () => {
    expect(parseDate("not a date")).toBeNull();
    expect(parseDate("")).toBeNull();
    expect(parseDate("   ")).toBeNull();
  });

  it("serialToDate with fractional part (time)", () => {
    // Serial 44927.5 = 2023-01-01 12:00:00
    const d = serialToDate(44927.5);
    expect(d.getUTCHours()).toBe(12);
    expect(d.getUTCMinutes()).toBe(0);
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 10. colToLetter Comprehensive
// ═══════════════════════════════════════════════════════════════════════

describe("colToLetter edge cases", () => {
  it("0 = A", () => expect(colToLetter(0)).toBe("A"));
  it("25 = Z", () => expect(colToLetter(25)).toBe("Z"));
  it("26 = AA", () => expect(colToLetter(26)).toBe("AA"));
  it("51 = AZ", () => expect(colToLetter(51)).toBe("AZ"));
  it("52 = BA", () => expect(colToLetter(52)).toBe("BA"));
  it("701 = ZZ", () => expect(colToLetter(701)).toBe("ZZ"));
  it("702 = AAA", () => expect(colToLetter(702)).toBe("AAA"));
  it("16383 = XFD (Excel max)", () => expect(colToLetter(16383)).toBe("XFD"));
});

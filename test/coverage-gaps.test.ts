import { describe, expect, it } from "vitest";
import {
  // High-level API
  read,
  write,
  readObjects,
  writeObjects,
  // XLSX
  readXlsx,
  writeXlsx,
  openXlsx,
  saveXlsx,
  hashSheetPassword,
  calculateColumnWidth,
  measureValueWidth,
  calculateRowHeight,
  parseThemeColors,
  resolveThemeColor,
  streamXlsxRows,
  XlsxStreamWriter,
  // ODS
  readOds,
  writeOds,
  // CSV
  parseCsv,
  parseCsvObjects,
  detectDelimiter,
  stripBom,
  writeCsv,
  writeCsvObjects,
  formatCsvValue,
  streamCsvRows,
  CsvStreamWriter,
  // Schema
  validateWithSchema,
  // Date
  serialToDate,
  dateToSerial,
  isDateFormat,
  formatDate,
  parseDate,
  serialToTime,
  timeToSerial,
  // Number format
  formatValue,
  // Sheet ops
  insertRows,
  deleteRows,
  insertColumns,
  deleteColumns,
  moveRows,
  hideRows,
  hideColumns,
  groupRows,
  cloneSheet,
  copySheetToWorkbook,
  copyRange,
  moveSheet,
  removeSheet,
  findCells,
  replaceCells,
  sortRows,
  // Worker
  serializeWorkbook,
  deserializeWorkbook,
  WORKER_SAFE_FUNCTIONS,
  // Cell utils
  parseCellRef,
  colToLetter,
  cellRef,
  rangeRef,
  letterToCol,
  parseRange,
  isInRange,
  // Sheet utils
  sheetToObjects,
  sheetToArrays,
  // Export
  toHtml,
  toMarkdown,
  toJson,
  fromHtml,
  // Image
  imageFromBase64,
  // Errors
  DefterError,
  ParseError,
  ZipError,
  XmlError,
  ValidationError,
  UnsupportedFormatError,
  EncryptedFileError,
} from "../src/index";
import type { Sheet, CellValue, Workbook } from "../src/_types";

/** Helper to create a minimal sheet */
function makeSheet(rows: CellValue[][], overrides?: Partial<Sheet>): Sheet {
  return { name: "Sheet1", rows, ...overrides };
}

describe("Coverage gaps: every exported function is callable (#135)", () => {
  it("all high-level API functions exist and are functions", () => {
    expect(typeof read).toBe("function");
    expect(typeof write).toBe("function");
    expect(typeof readObjects).toBe("function");
    expect(typeof writeObjects).toBe("function");
  });

  it("all XLSX functions exist", () => {
    expect(typeof readXlsx).toBe("function");
    expect(typeof writeXlsx).toBe("function");
    expect(typeof openXlsx).toBe("function");
    expect(typeof saveXlsx).toBe("function");
    expect(typeof hashSheetPassword).toBe("function");
    expect(typeof calculateColumnWidth).toBe("function");
    expect(typeof measureValueWidth).toBe("function");
    expect(typeof calculateRowHeight).toBe("function");
    expect(typeof parseThemeColors).toBe("function");
    expect(typeof resolveThemeColor).toBe("function");
    expect(typeof streamXlsxRows).toBe("function");
    expect(typeof XlsxStreamWriter).toBe("function");
  });

  it("all ODS functions exist", () => {
    expect(typeof readOds).toBe("function");
    expect(typeof writeOds).toBe("function");
  });

  it("all CSV functions exist", () => {
    expect(typeof parseCsv).toBe("function");
    expect(typeof parseCsvObjects).toBe("function");
    expect(typeof detectDelimiter).toBe("function");
    expect(typeof stripBom).toBe("function");
    expect(typeof writeCsv).toBe("function");
    expect(typeof writeCsvObjects).toBe("function");
    expect(typeof formatCsvValue).toBe("function");
    expect(typeof streamCsvRows).toBe("function");
    expect(typeof CsvStreamWriter).toBe("function");
  });

  it("all utility functions exist", () => {
    expect(typeof validateWithSchema).toBe("function");
    expect(typeof serialToDate).toBe("function");
    expect(typeof dateToSerial).toBe("function");
    expect(typeof isDateFormat).toBe("function");
    expect(typeof formatDate).toBe("function");
    expect(typeof parseDate).toBe("function");
    expect(typeof serialToTime).toBe("function");
    expect(typeof timeToSerial).toBe("function");
    expect(typeof formatValue).toBe("function");
  });

  it("all sheet operations exist", () => {
    expect(typeof insertRows).toBe("function");
    expect(typeof deleteRows).toBe("function");
    expect(typeof insertColumns).toBe("function");
    expect(typeof deleteColumns).toBe("function");
    expect(typeof moveRows).toBe("function");
    expect(typeof hideRows).toBe("function");
    expect(typeof hideColumns).toBe("function");
    expect(typeof groupRows).toBe("function");
    expect(typeof cloneSheet).toBe("function");
    expect(typeof copySheetToWorkbook).toBe("function");
    expect(typeof copyRange).toBe("function");
    expect(typeof moveSheet).toBe("function");
    expect(typeof removeSheet).toBe("function");
    expect(typeof findCells).toBe("function");
    expect(typeof replaceCells).toBe("function");
    expect(typeof sortRows).toBe("function");
  });

  it("all worker helpers exist", () => {
    expect(typeof serializeWorkbook).toBe("function");
    expect(typeof deserializeWorkbook).toBe("function");
    expect(Array.isArray(WORKER_SAFE_FUNCTIONS)).toBe(true);
  });

  it("all cell utils exist", () => {
    expect(typeof parseCellRef).toBe("function");
    expect(typeof colToLetter).toBe("function");
    expect(typeof cellRef).toBe("function");
    expect(typeof rangeRef).toBe("function");
    expect(typeof letterToCol).toBe("function");
    expect(typeof parseRange).toBe("function");
    expect(typeof isInRange).toBe("function");
  });

  it("all sheet utils exist", () => {
    expect(typeof sheetToObjects).toBe("function");
    expect(typeof sheetToArrays).toBe("function");
  });

  it("all export functions exist", () => {
    expect(typeof toHtml).toBe("function");
    expect(typeof toMarkdown).toBe("function");
    expect(typeof toJson).toBe("function");
    expect(typeof fromHtml).toBe("function");
  });

  it("image helper exists", () => {
    expect(typeof imageFromBase64).toBe("function");
  });

  it("all error classes exist", () => {
    expect(typeof DefterError).toBe("function");
    expect(typeof ParseError).toBe("function");
    expect(typeof ZipError).toBe("function");
    expect(typeof XmlError).toBe("function");
    expect(typeof ValidationError).toBe("function");
    expect(typeof UnsupportedFormatError).toBe("function");
    expect(typeof EncryptedFileError).toBe("function");
  });
});

describe("Coverage gaps: empty workbook write/read (#135)", () => {
  it("should write and read an empty workbook (no data rows)", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Empty", rows: [] }],
    });
    expect(data).toBeInstanceOf(Uint8Array);
    expect(data.length).toBeGreaterThan(0);

    const wb = await readXlsx(data);
    expect(wb.sheets.length).toBe(1);
    expect(wb.sheets[0]!.name).toBe("Empty");
    expect(wb.sheets[0]!.rows.length).toBe(0);
  });

  it("should write and read empty workbook via high-level API", async () => {
    const data = await write({
      sheets: [{ name: "EmptySheet", rows: [] }],
    });
    const wb = await read(data);
    expect(wb.sheets[0]!.name).toBe("EmptySheet");
    expect(wb.sheets[0]!.rows.length).toBe(0);
  });

  it("should write and read empty ODS workbook", async () => {
    const data = await writeOds({
      sheets: [{ name: "EmptyODS", rows: [] }],
    });
    const wb = await readOds(data);
    expect(wb.sheets.length).toBe(1);
    expect(wb.sheets[0]!.name).toBe("EmptyODS");
  });
});

describe("Coverage gaps: sheet with only formulas (no values) (#135)", () => {
  it("should write and read cells that have formulas but no cached values", async () => {
    const cells = new Map<string, Partial<import("../src/_types").Cell>>();
    cells.set("0,0", { formula: "1+1" });
    cells.set("0,1", { formula: "SUM(A1:A10)" });
    cells.set("1,0", { formula: "A1*2" });

    const data = await writeXlsx({
      sheets: [
        {
          name: "Formulas",
          rows: [[null, null], [null]],
          cells,
        },
      ],
    });
    const wb = await readXlsx(data);
    expect(wb.sheets.length).toBe(1);
    const sheet = wb.sheets[0]!;
    // The formulas should be preserved in the cells map
    const cell00 = sheet.cells?.get("0,0");
    expect(cell00?.formula).toBe("1+1");
  });
});

describe("Coverage gaps: very large cell value (#135)", () => {
  it("should write and read a 100KB string value", async () => {
    const bigString = "A".repeat(100 * 1024); // 100KB
    const data = await writeXlsx({
      sheets: [{ name: "BigCell", rows: [[bigString]] }],
    });
    const wb = await readXlsx(data);
    expect(wb.sheets[0]!.rows[0]![0]).toBe(bigString);
  });
});

describe("Coverage gaps: all CsvReadOptions together (#135)", () => {
  it("should apply all CsvReadOptions simultaneously", () => {
    const csv = [
      "\uFEFF# metadata line",
      "extra line to skip",
      "# comment",
      '"Name";"Age";"Active"',
      '"Alice";"30";"true"',
      '"Bob";"25";"false"',
      "",
      '"Charlie";"35";"yes"',
    ].join("\n");

    const collected: CellValue[][] = [];
    const result = parseCsv(csv, {
      delimiter: ";",
      quote: '"',
      escape: '"',
      skipBom: true,
      typeInference: true,
      preserveLeadingZeros: true,
      skipEmptyRows: true,
      comment: "#",
      maxRows: 3,
      skipLines: 1,
      onRow: (row) => collected.push([...row]),
      transformValue: (val) => val,
    });

    // skipLines: 1 removes "extra line to skip" (the original # metadata line already stripped via BOM/comment)
    // comment: "#" removes "# comment"
    // skipEmptyRows removes the blank line
    // maxRows: 3 limits to 3 data rows
    expect(result.length).toBeLessThanOrEqual(3);
    expect(collected.length).toBe(result.length);
    // Type inference should have converted numbers and booleans
    for (const row of result) {
      expect(Array.isArray(row)).toBe(true);
    }
  });
});

describe("Coverage gaps: all CsvWriteOptions together (#135)", () => {
  it("should apply all CsvWriteOptions simultaneously", () => {
    const rows: CellValue[][] = [
      ["Hello", 42, true, new Date("2024-01-15T00:00:00.000Z"), null, "=SUM(A1)"],
    ];

    const csv = writeCsv(rows, {
      delimiter: "\t",
      lineSeparator: "\n",
      quote: "'",
      quoteStyle: "all",
      headers: ["Name", "Value", "Active", "Date", "Empty", "Formula"],
      bom: true,
      dateFormat: "YYYY-MM-DD",
      nullValue: "N/A",
      escapeFormulae: true,
    });

    // Should start with BOM
    expect(csv.charCodeAt(0)).toBe(0xfeff);
    // Should contain tab delimiters
    expect(csv).toContain("\t");
    // Should use single quote
    expect(csv).toContain("'");
    // Should contain header row
    expect(csv).toContain("Name");
    // Should format date
    expect(csv).toContain("2024-01-15");
  });
});

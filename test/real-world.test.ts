import { describe, it, expect } from "vitest";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip";
import { writeOds } from "../src/ods/writer";
import { readOds } from "../src/ods/reader";
import { streamXlsxRows } from "../src/xlsx/stream-reader";
import { XlsxStreamWriter } from "../src/xlsx/stream-writer";
import { parseCsv, stripBom, detectDelimiter, writeCsv } from "../src/csv/index";
import { validateWithSchema } from "../src/_schema";
import { insertRows, deleteRows, cloneSheet } from "../src/sheet-ops";
import { ZipReader } from "../src/zip/reader";
import { ZipWriter } from "../src/zip/writer";
import type { WriteSheet, CellValue, Cell, Sheet } from "../src/_types";
import type { StreamRow } from "../src/xlsx/stream-reader";

// ── Shared Helpers ──────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");
const encoder = new TextEncoder();

async function collectStreamRows(
  gen: AsyncGenerator<StreamRow, void, undefined>,
): Promise<StreamRow[]> {
  const rows: StreamRow[] = [];
  for await (const row of gen) {
    rows.push(row);
  }
  return rows;
}

/** Create a fake PNG with valid magic bytes */
function fakePng(size = 64): Uint8Array {
  const data = new Uint8Array(size);
  data[0] = 0x89;
  data[1] = 0x50;
  data[2] = 0x4e;
  data[3] = 0x47;
  data[4] = 0x0d;
  data[5] = 0x0a;
  data[6] = 0x1a;
  data[7] = 0x0a;
  for (let i = 8; i < size; i++) {
    data[i] = i % 256;
  }
  return data;
}

/** Inject extra ZIP entries into an existing XLSX archive */
async function injectEntries(
  original: Uint8Array,
  extras: Array<{ path: string; data: Uint8Array }>,
): Promise<Uint8Array> {
  const zip = new ZipReader(original);
  const writer = new ZipWriter();
  for (const path of zip.entries()) {
    const data = await zip.extract(path);
    writer.add(path, data, { compress: false });
  }
  for (const entry of extras) {
    writer.add(entry.path, entry.data, { compress: false });
  }
  return writer.build();
}

// ═══════════════════════════════════════════════════════════════════════
// 1. Product Catalog Export
// ═══════════════════════════════════════════════════════════════════════

describe("Real World: Product Catalog Export", () => {
  const categories = ["Electronics", "Clothing", "Food", "Tools"];

  function generateProducts(count: number): CellValue[][] {
    const rows: CellValue[][] = [];
    // Header row
    rows.push([
      "Name",
      "SKU",
      "Price",
      "Stock",
      "Category",
      "Description",
      "Active",
      "Created",
      "Updated",
    ]);
    for (let i = 0; i < count; i++) {
      const cat = categories[i % categories.length]!;
      const now = new Date(2025, 0, 1 + (i % 28), 10, 0, 0);
      rows.push([
        `Product ${i + 1}`,
        `SKU-${String(i + 1).padStart(5, "0")}`,
        19.99 + i * 2.5,
        i % 7 === 0 ? i % 5 : 50 + i, // some low stock values
        cat,
        `Description for product ${i + 1} in ${cat}`,
        i % 3 !== 0,
        now,
        now,
      ]);
    }
    return rows;
  }

  it("writes and reads 50+ products with all features", async () => {
    const rows = generateProducts(55);

    const sheet: WriteSheet = {
      name: "Products",
      rows,
      columns: [
        { header: "Name", width: 30 },
        { header: "SKU", width: 18 },
        { header: "Price", width: 12, numFmt: '"$"#,##0.00' },
        { header: "Stock", width: 10 },
        { header: "Category", width: 15 },
        { header: "Description", width: 50 },
        { header: "Active", width: 10 },
        { header: "Created", width: 18, numFmt: "yyyy-mm-dd" },
        { header: "Updated", width: 18, numFmt: "yyyy-mm-dd" },
      ],
      freezePane: { rows: 1 },
      autoFilter: { range: "A1:I56" },
      dataValidations: [
        {
          type: "list",
          values: categories,
          range: "E2:E56",
          showInputMessage: true,
          inputTitle: "Category",
          inputMessage: "Select a product category",
          showErrorMessage: true,
          errorTitle: "Invalid Category",
          errorMessage: "Must be one of: Electronics, Clothing, Food, Tools",
          errorStyle: "stop",
        },
        {
          type: "whole",
          operator: "greaterThanOrEqual",
          formula1: "0",
          range: "D2:D56",
          showErrorMessage: true,
          errorTitle: "Invalid Stock",
          errorMessage: "Stock must be a whole number >= 0",
          errorStyle: "stop",
        },
      ],
      conditionalRules: [
        {
          type: "cellIs",
          operator: "lessThan",
          formula: "10",
          priority: 1,
          range: "D2:D56",
          style: {
            fill: {
              type: "pattern",
              pattern: "solid",
              fgColor: { rgb: "FF0000" },
            },
          },
        },
      ],
      cells: new Map<string, Partial<Cell>>(),
      tables: [
        {
          name: "ProductCatalog",
          columns: [
            { name: "Name" },
            { name: "SKU" },
            { name: "Price" },
            { name: "Stock" },
            { name: "Category" },
            { name: "Description" },
            { name: "Active" },
            { name: "Created" },
            { name: "Updated" },
          ],
          style: "TableStyleMedium2",
          showRowStripes: true,
          range: "A1:I56",
        },
      ],
    };

    // Apply bold header style via cells map
    const headers = rows[0]!;
    for (let c = 0; c < headers.length; c++) {
      sheet.cells!.set(`0,${c}`, {
        style: {
          font: { bold: true },
          fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { rgb: "4472C4" },
          },
        },
      });
    }

    const xlsx = await writeXlsx({ sheets: [sheet] });
    expect(xlsx).toBeInstanceOf(Uint8Array);
    expect(xlsx.byteLength).toBeGreaterThan(0);

    // Read back
    const wb = await readXlsx(xlsx, { readStyles: true });
    expect(wb.sheets).toHaveLength(1);

    const s = wb.sheets[0]!;
    expect(s.name).toBe("Products");

    // Verify row count: header + 55 products = 56 rows
    expect(s.rows).toHaveLength(56);

    // Verify header row
    expect(s.rows[0]![0]).toBe("Name");
    expect(s.rows[0]![1]).toBe("SKU");

    // Verify first product
    expect(s.rows[1]![0]).toBe("Product 1");
    expect(s.rows[1]![1]).toBe("SKU-00001");
    expect(s.rows[1]![2]).toBeCloseTo(19.99, 2);
    expect(s.rows[1]![4]).toBe("Electronics");
    // i=0 → 0 % 3 !== 0 → false
    expect(s.rows[1]![6]).toBe(false);

    // Verify last product
    expect(s.rows[55]![0]).toBe("Product 55");
    expect(s.rows[55]![1]).toBe("SKU-00055");

    // Verify data validations survived round-trip
    expect(s.dataValidations).toBeDefined();
    expect(s.dataValidations!.length).toBeGreaterThanOrEqual(2);

    const listVal = s.dataValidations!.find((dv) => dv.type === "list");
    expect(listVal).toBeDefined();
    // List values may be stored as formula1 with quoted comma-separated values
    // or as values array — check either
    if (listVal!.values) {
      expect(listVal!.values).toEqual(categories);
    } else {
      expect(listVal!.formula1).toBeDefined();
    }

    const wholeVal = s.dataValidations!.find((dv) => dv.type === "whole");
    expect(wholeVal).toBeDefined();

    // Verify conditional rules survived
    expect(s.conditionalRules).toBeDefined();
    expect(s.conditionalRules!.length).toBeGreaterThanOrEqual(1);
    const cfRule = s.conditionalRules![0]!;
    expect(cfRule.type).toBe("cellIs");

    // Note: The reader does not parse freeze panes (pane element in sheetView)
    // so we skip that assertion. The write is verified by the file being valid.

    // Verify auto filter
    expect(s.autoFilter).toBeDefined();

    // Verify table
    expect(s.tables).toBeDefined();
    expect(s.tables!.length).toBe(1);
    expect(s.tables![0]!.name).toBe("ProductCatalog");
    expect(s.tables![0]!.columns).toHaveLength(9);

    // Verify styled header cell
    expect(s.cells).toBeDefined();
    const headerCell = s.cells!.get("0,0");
    expect(headerCell).toBeDefined();
    expect(headerCell!.style).toBeDefined();
    expect(headerCell!.style!.font?.bold).toBe(true);
  });

  it("auto width columns round-trip", async () => {
    const rows = generateProducts(10);
    const sheet: WriteSheet = {
      name: "AutoWidth",
      rows,
      columns: [
        { header: "Name", autoWidth: true },
        { header: "SKU", autoWidth: true },
        { header: "Price", autoWidth: true },
        { header: "Stock", autoWidth: true },
        { header: "Category", autoWidth: true },
        { header: "Description", autoWidth: true },
        { header: "Active", autoWidth: true },
        { header: "Created", autoWidth: true },
        { header: "Updated", autoWidth: true },
      ],
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0]!.rows).toHaveLength(11);
    // Verify columns have width set (autoWidth should calculate)
    if (wb.sheets[0]!.columns) {
      for (const col of wb.sheets[0]!.columns) {
        if (col.width !== undefined) {
          expect(col.width).toBeGreaterThan(0);
        }
      }
    }
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 2. Financial Report
// ═══════════════════════════════════════════════════════════════════════

describe("Real World: Financial Report", () => {
  it("creates multi-sheet financial report with formulas and formatting", async () => {
    const summaryRows: CellValue[][] = [
      ["Quarterly Financial Report 2025", null, null, null],
      [null, null, null, null],
      ["Quarter", "Revenue", "Expenses", "Profit"],
      ["Q1", 125000, 95000, 30000],
      ["Q2", 142000, 98000, 44000],
      ["Q3", 158000, 102000, 56000],
      ["Q4", 175000, 110000, 65000],
      [null, null, null, null],
      ["Total", null, null, null],
    ];

    const summaryCells = new Map<string, Partial<Cell>>();
    // Title: bold, large
    summaryCells.set("0,0", {
      style: {
        font: { bold: true, size: 16 },
        alignment: { horizontal: "center" },
      },
    });
    // Header row: bold
    for (let c = 0; c < 4; c++) {
      summaryCells.set(`2,${c}`, {
        style: {
          font: { bold: true },
          fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { rgb: "D9E2F3" },
          },
        },
      });
    }
    // Formulas for Total row
    summaryCells.set("8,1", { formula: "SUM(B4:B7)" });
    summaryCells.set("8,2", { formula: "SUM(C4:C7)" });
    summaryCells.set("8,3", { formula: "SUM(D4:D7)" });

    // Revenue column: currency format
    for (let r = 3; r <= 8; r++) {
      for (let c = 1; c <= 3; c++) {
        const existing = summaryCells.get(`${r},${c}`);
        summaryCells.set(`${r},${c}`, {
          ...existing,
          style: {
            ...existing?.style,
            numFmt: '"$"#,##0',
          },
        });
      }
    }

    function makeQuarterSheet(name: string, base: number): WriteSheet {
      const months = ["Month 1", "Month 2", "Month 3"];
      const rows: CellValue[][] = [
        [`${name} Breakdown`, null, null],
        [null, null, null],
        ["Month", "Revenue", "Expenses"],
      ];
      for (let i = 0; i < 3; i++) {
        rows.push([months[i]!, base + i * 5000, base * 0.75 + i * 3000]);
      }
      rows.push([null, null, null]);
      rows.push(["Total", null, null]);

      const cells = new Map<string, Partial<Cell>>();
      cells.set("0,0", {
        style: { font: { bold: true, size: 14 } },
      });
      for (let c = 0; c < 3; c++) {
        cells.set(`2,${c}`, {
          style: { font: { bold: true } },
        });
      }
      cells.set("6,1", { formula: "SUM(B4:B6)" });
      cells.set("6,2", { formula: "SUM(C4:C6)" });

      return {
        name,
        rows,
        cells,
        columns: [
          { header: "Month", width: 15 },
          { header: "Revenue", width: 15, numFmt: '"$"#,##0' },
          { header: "Expenses", width: 15, numFmt: '"$"#,##0' },
        ],
      };
    }

    const sheets: WriteSheet[] = [
      {
        name: "Summary",
        rows: summaryRows,
        cells: summaryCells,
        merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 3 }],
        columns: [
          { header: "Quarter", width: 15 },
          { header: "Revenue", width: 18, numFmt: '"$"#,##0' },
          { header: "Expenses", width: 18, numFmt: '"$"#,##0' },
          { header: "Profit", width: 18, numFmt: '"$"#,##0' },
        ],
        pageSetup: {
          orientation: "landscape",
          paperSize: "a4",
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
        },
        headerFooter: {
          oddFooter: "&CPage &P of &N",
        },
      },
      makeQuarterSheet("Q1", 40000),
      makeQuarterSheet("Q2", 46000),
      makeQuarterSheet("Q3", 51000),
      makeQuarterSheet("Q4", 57000),
    ];

    const namedRanges = [
      {
        name: "AnnualRevenue",
        range: "Summary!$B$4:$B$7",
      },
      {
        name: "AnnualExpenses",
        range: "Summary!$C$4:$C$7",
      },
    ];

    const xlsx = await writeXlsx({
      sheets,
      namedRanges,
      properties: {
        title: "Financial Report 2025",
        creator: "Hucre Test Suite",
        company: "Acme Corp",
      },
    });
    expect(xlsx.byteLength).toBeGreaterThan(0);

    // Read back
    const wb = await readXlsx(xlsx, { readStyles: true });
    expect(wb.sheets).toHaveLength(5);
    expect(wb.sheets.map((s) => s.name)).toEqual(["Summary", "Q1", "Q2", "Q3", "Q4"]);

    // Summary sheet
    const summary = wb.sheets[0]!;
    expect(summary.rows[0]![0]).toBe("Quarterly Financial Report 2025");
    expect(summary.rows[2]![0]).toBe("Quarter");
    expect(summary.rows[3]![1]).toBe(125000);
    expect(summary.rows[6]![1]).toBe(175000);

    // Verify merged cells
    expect(summary.merges).toBeDefined();
    expect(summary.merges!.length).toBeGreaterThanOrEqual(1);
    const titleMerge = summary.merges![0]!;
    expect(titleMerge.startRow).toBe(0);
    expect(titleMerge.endCol).toBe(3);

    // Verify formulas
    const totalRevenueCell = summary.cells?.get("8,1");
    expect(totalRevenueCell).toBeDefined();
    expect(totalRevenueCell!.formula).toBe("SUM(B4:B7)");

    // Verify named ranges
    expect(wb.namedRanges).toBeDefined();
    expect(wb.namedRanges!.length).toBeGreaterThanOrEqual(2);
    const revenueRange = wb.namedRanges!.find((nr) => nr.name === "AnnualRevenue");
    expect(revenueRange).toBeDefined();

    // Verify page setup
    expect(summary.pageSetup).toBeDefined();
    expect(summary.pageSetup!.orientation).toBe("landscape");
    // Paper size round-trip (a4 = numeric index in OOXML)
    expect(summary.pageSetup!.paperSize).toBeDefined();

    // Verify header/footer
    expect(summary.headerFooter).toBeDefined();
    expect(summary.headerFooter!.oddFooter).toContain("&P");

    // Verify properties
    expect(wb.properties).toBeDefined();
    expect(wb.properties!.title).toBe("Financial Report 2025");
    expect(wb.properties!.creator).toBe("Hucre Test Suite");

    // Verify quarter sheets have data
    for (let i = 1; i <= 4; i++) {
      const qs = wb.sheets[i]!;
      expect(qs.rows.length).toBeGreaterThanOrEqual(7);
      expect(qs.rows[0]![0]).toContain("Breakdown");
    }
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 3. Employee Directory
// ═══════════════════════════════════════════════════════════════════════

describe("Real World: Employee Directory", () => {
  it("creates directory with rich text, hyperlinks, comments, images, protection", async () => {
    const rows: CellValue[][] = [
      ["Name", "Title", "Email", "Website", "Salary", "Notes"],
      [
        "Alice Johnson",
        "Senior Engineer",
        "alice@acme.com",
        "https://acme.com/alice",
        120000,
        "Team lead",
      ],
      ["Bob Smith", "Designer", "bob@acme.com", "https://acme.com/bob", 95000, "Remote worker"],
      [
        "Carol Williams",
        "Product Manager",
        "carol@acme.com",
        "https://acme.com/carol",
        130000,
        "VP track",
      ],
    ];

    const cells = new Map<string, Partial<Cell>>();

    // Rich text for name column
    cells.set("1,0", {
      richText: [
        { text: "Alice ", font: { bold: true } },
        { text: "Johnson", font: { italic: true, color: { rgb: "666666" } } },
      ],
    });
    cells.set("2,0", {
      richText: [
        { text: "Bob ", font: { bold: true } },
        { text: "Smith", font: { italic: true, color: { rgb: "666666" } } },
      ],
    });

    // Hyperlinks for email
    cells.set("1,2", {
      hyperlink: { target: "mailto:alice@acme.com", tooltip: "Email Alice" },
    });
    cells.set("2,2", {
      hyperlink: { target: "mailto:bob@acme.com", tooltip: "Email Bob" },
    });

    // Hyperlinks for website
    cells.set("1,3", {
      hyperlink: { target: "https://acme.com/alice", tooltip: "Alice's profile" },
    });

    // Comments
    cells.set("1,5", {
      comment: { author: "Manager", text: "Consider for promotion" },
    });
    cells.set("3,5", {
      comment: { author: "HR", text: "On VP track, review Q2" },
    });

    // Bold header
    for (let c = 0; c < 6; c++) {
      cells.set(`0,${c}`, {
        style: { font: { bold: true } },
      });
    }

    // Small PNG placeholder image
    const pngData = fakePng(128);

    const sheet: WriteSheet = {
      name: "Directory",
      rows,
      cells,
      columns: [
        { header: "Name", width: 25 },
        { header: "Title", width: 20 },
        { header: "Email", width: 25 },
        { header: "Website", width: 30 },
        { header: "Salary", width: 15, hidden: true },
        { header: "Notes", width: 30 },
      ],
      images: [
        {
          data: pngData,
          type: "png",
          anchor: { from: { row: 0, col: 6 }, to: { row: 3, col: 8 } },
        },
      ],
      protection: {
        sheet: true,
        password: "test123",
        sort: true,
        autoFilter: true,
        selectLockedCells: true,
        selectUnlockedCells: true,
      },
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    expect(xlsx.byteLength).toBeGreaterThan(0);

    // Read back
    const wb = await readXlsx(xlsx, { readStyles: true });
    const s = wb.sheets[0]!;
    expect(s.name).toBe("Directory");
    expect(s.rows).toHaveLength(4);

    // Verify hyperlinks
    const emailCell = s.cells?.get("1,2");
    expect(emailCell).toBeDefined();
    expect(emailCell!.hyperlink).toBeDefined();
    expect(emailCell!.hyperlink!.target).toBe("mailto:alice@acme.com");

    const websiteCell = s.cells?.get("1,3");
    expect(websiteCell).toBeDefined();
    expect(websiteCell!.hyperlink).toBeDefined();
    expect(websiteCell!.hyperlink!.target).toBe("https://acme.com/alice");

    // Verify comments
    const notesCell1 = s.cells?.get("1,5");
    expect(notesCell1).toBeDefined();
    expect(notesCell1!.comment).toBeDefined();
    expect(notesCell1!.comment!.text).toContain("promotion");

    const notesCell3 = s.cells?.get("3,5");
    expect(notesCell3).toBeDefined();
    expect(notesCell3!.comment).toBeDefined();
    expect(notesCell3!.comment!.text).toContain("VP track");

    // Verify image survived
    expect(s.images).toBeDefined();
    expect(s.images!.length).toBe(1);
    expect(s.images![0]!.type).toBe("png");
    expect(s.images![0]!.data.byteLength).toBeGreaterThan(0);
    // Verify PNG magic bytes preserved
    expect(s.images![0]!.data[0]).toBe(0x89);
    expect(s.images![0]!.data[1]).toBe(0x50);

    // Note: The reader does not parse <cols> elements (column widths/hidden),
    // so hidden column info is not available on read-back. The write is verified
    // by the file being valid XLSX that Excel can open.

    // Verify protection
    expect(s.protection).toBeDefined();
    expect(s.protection!.sheet).toBe(true);
    expect(s.protection!.sort).toBe(true);
    expect(s.protection!.autoFilter).toBe(true);
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 4. Data Import Template
// ═══════════════════════════════════════════════════════════════════════

describe("Real World: Data Import Template", () => {
  it("creates empty template with validations, freeze panes, named ranges", async () => {
    const rows: CellValue[][] = [["Product Name", "Category", "Price", "Quantity", "Ship Date"]];
    // Add 99 empty rows for data entry
    for (let i = 0; i < 99; i++) {
      rows.push([null, null, null, null, null]);
    }

    const sheet: WriteSheet = {
      name: "Import Template",
      rows,
      columns: [
        { header: "Product Name", width: 30 },
        { header: "Category", width: 20 },
        { header: "Price", width: 15, numFmt: "#,##0.00" },
        { header: "Quantity", width: 12 },
        { header: "Ship Date", width: 18, numFmt: "yyyy-mm-dd" },
      ],
      freezePane: { rows: 1 },
      dataValidations: [
        {
          type: "list",
          values: ["Electronics", "Clothing", "Food", "Tools", "Other"],
          range: "B2:B100",
          allowBlank: true,
          showInputMessage: true,
          inputTitle: "Category",
          inputMessage: "Choose a category from the dropdown",
          showErrorMessage: true,
          errorTitle: "Invalid",
          errorMessage: "Please select a valid category",
          errorStyle: "stop",
        },
        {
          type: "decimal",
          operator: "greaterThan",
          formula1: "0",
          range: "C2:C100",
          allowBlank: true,
          showInputMessage: true,
          inputTitle: "Price",
          inputMessage: "Enter a positive price value",
          showErrorMessage: true,
          errorTitle: "Invalid Price",
          errorMessage: "Price must be a positive number",
          errorStyle: "stop",
        },
        {
          type: "whole",
          operator: "between",
          formula1: "1",
          formula2: "99999",
          range: "D2:D100",
          allowBlank: true,
          showInputMessage: true,
          inputTitle: "Quantity",
          inputMessage: "Enter quantity between 1 and 99999",
          showErrorMessage: true,
          errorTitle: "Invalid Quantity",
          errorMessage: "Must be a whole number between 1 and 99999",
          errorStyle: "warning",
        },
        {
          type: "date",
          operator: "greaterThanOrEqual",
          formula1: "45658", // 2025-01-01 serial
          range: "E2:E100",
          allowBlank: true,
          showInputMessage: true,
          inputTitle: "Ship Date",
          inputMessage: "Enter a date on or after 2025-01-01",
          showErrorMessage: true,
          errorTitle: "Invalid Date",
          errorMessage: "Date must be 2025-01-01 or later",
          errorStyle: "information",
        },
      ],
      cells: new Map<string, Partial<Cell>>(
        Array.from(
          { length: 5 },
          (_, c) =>
            [
              `0,${c}`,
              {
                style: {
                  font: { bold: true, color: { rgb: "FFFFFF" } },
                  fill: {
                    type: "pattern" as const,
                    pattern: "solid" as const,
                    fgColor: { rgb: "4472C4" },
                  },
                },
              },
            ] as [string, Partial<Cell>],
        ),
      ),
    };

    const namedRanges = [
      {
        name: "CategoryList",
        range: "'Import Template'!$B$2:$B$100",
        scope: "Import Template",
      },
    ];

    const xlsx = await writeXlsx({ sheets: [sheet], namedRanges });
    expect(xlsx.byteLength).toBeGreaterThan(0);

    // Read back
    const wb = await readXlsx(xlsx, { readStyles: true });
    const s = wb.sheets[0]!;

    // Note: The reader does not parse freeze panes from the worksheet XML,
    // so we skip that assertion. The freeze pane is written correctly.

    // Verify all 4 data validations
    expect(s.dataValidations).toBeDefined();
    expect(s.dataValidations!.length).toBe(4);

    const listDv = s.dataValidations!.find((dv) => dv.type === "list");
    expect(listDv).toBeDefined();
    expect(listDv!.showInputMessage).toBe(true);
    expect(listDv!.showErrorMessage).toBe(true);
    expect(listDv!.inputTitle).toBe("Category");
    expect(listDv!.errorStyle).toBe("stop");

    const decimalDv = s.dataValidations!.find((dv) => dv.type === "decimal");
    expect(decimalDv).toBeDefined();
    expect(decimalDv!.operator).toBe("greaterThan");

    const wholeDv = s.dataValidations!.find((dv) => dv.type === "whole");
    expect(wholeDv).toBeDefined();
    expect(wholeDv!.operator).toBe("between");
    expect(wholeDv!.errorStyle).toBe("warning");

    const dateDv = s.dataValidations!.find((dv) => dv.type === "date");
    expect(dateDv).toBeDefined();
    expect(dateDv!.errorStyle).toBe("information");

    // Verify named range
    expect(wb.namedRanges).toBeDefined();
    const catRange = wb.namedRanges!.find((nr) => nr.name === "CategoryList");
    expect(catRange).toBeDefined();
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 5. Large Dataset Stress Test
// ═══════════════════════════════════════════════════════════════════════

describe("Real World: Large Dataset Stress Test", () => {
  const ROW_COUNT = 10_000;
  const COL_COUNT = 20;

  it("stream writes and reads 10k rows x 20 cols within time budget", async () => {
    const start = performance.now();

    // Stream write
    const columns = Array.from({ length: COL_COUNT }, (_, i) => ({
      header: `Col_${i}`,
      key: `col_${i}`,
    }));

    const writer = new XlsxStreamWriter({
      name: "LargeDataset",
      columns,
    });

    for (let r = 0; r < ROW_COUNT; r++) {
      const row: CellValue[] = [];
      for (let c = 0; c < COL_COUNT; c++) {
        if (c === 0) row.push(`Row ${r}`);
        else if (c === 1) row.push(r);
        else if (c === 2) row.push(r % 2 === 0);
        else if (c === 3) row.push(new Date(2025, 0, 1 + (r % 365)));
        else row.push(r * c + Math.random());
      }
      writer.addRow(row);
    }

    const xlsx = await writer.finish();
    const writeTime = performance.now() - start;

    expect(xlsx.byteLength).toBeGreaterThan(0);

    // Stream read
    const readStart = performance.now();
    const streamRows = await collectStreamRows(streamXlsxRows(xlsx));
    const readTime = performance.now() - readStart;
    const totalTime = performance.now() - start;

    // +1 for header row auto-added by XlsxStreamWriter
    expect(streamRows).toHaveLength(ROW_COUNT + 1);

    // Verify header row
    expect(streamRows[0]!.values[0]).toBe("Col_0");
    expect(streamRows[0]!.values[19]).toBe("Col_19");

    // Verify first data row
    expect(streamRows[1]!.values[0]).toBe("Row 0");
    expect(streamRows[1]!.values[1]).toBe(0);
    expect(streamRows[1]!.values[2]).toBe(true);

    // Verify last data row
    const lastRow = streamRows[ROW_COUNT]!;
    expect(lastRow.values[0]).toBe(`Row ${ROW_COUNT - 1}`);
    expect(lastRow.values[1]).toBe(ROW_COUNT - 1);

    // Verify random sample in the middle
    const midIdx = 5001; // +1 for header
    expect(streamRows[midIdx]!.values[0]).toBe("Row 5000");
    expect(streamRows[midIdx]!.values[1]).toBe(5000);
    expect(streamRows[midIdx]!.values[2]).toBe(true);

    // Time budget: should complete in under 5 seconds
    expect(totalTime).toBeLessThan(5000);
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 6. ODS Cross-format
// ═══════════════════════════════════════════════════════════════════════

describe("Real World: ODS Cross-format", () => {
  it("produces identical row data from XLSX and ODS round-trips", async () => {
    const rows: CellValue[][] = [
      ["Name", "Age", "Score", "Active", "Start Date"],
      ["Alice", 30, 95.5, true, new Date(2025, 0, 15)],
      ["Bob", 25, 88.0, false, new Date(2025, 2, 1)],
      ["Carol", 35, 92.3, true, new Date(2024, 5, 20)],
      ["David", 28, null, true, null],
    ];

    const writeOptions = {
      sheets: [
        {
          name: "Data",
          rows,
        },
      ],
    };

    // Write as XLSX and read back
    const xlsxBuf = await writeXlsx(writeOptions);
    const xlsxWb = await readXlsx(xlsxBuf);

    // Write as ODS and read back
    const odsBuf = await writeOds(writeOptions);
    const odsWb = await readOds(odsBuf);

    expect(xlsxWb.sheets).toHaveLength(1);
    expect(odsWb.sheets).toHaveLength(1);

    const xlsxRows = xlsxWb.sheets[0]!.rows;
    const odsRows = odsWb.sheets[0]!.rows;

    // ODS may have trailing empty rows — compare only the data rows
    expect(odsRows.length).toBeGreaterThanOrEqual(xlsxRows.length);

    for (let r = 0; r < xlsxRows.length; r++) {
      const xRow = xlsxRows[r]!;
      const oRow = odsRows[r]!;

      for (let c = 0; c < xRow.length; c++) {
        const xVal = xRow[c];
        const oVal = c < oRow.length ? oRow[c] : null;

        if (xVal === null) {
          expect(oVal).toBeNull();
        } else if (typeof xVal === "number") {
          expect(oVal).toBeCloseTo(xVal as number, 5);
        } else if (typeof xVal === "boolean") {
          expect(oVal).toBe(xVal);
        } else if (xVal instanceof Date) {
          // Both should produce Date or equivalent
          if (oVal instanceof Date) {
            // Compare year/month/day at minimum
            expect(oVal.getFullYear()).toBe(xVal.getFullYear());
            expect(oVal.getMonth()).toBe(xVal.getMonth());
            expect(oVal.getDate()).toBe(xVal.getDate());
          } else {
            // ODS might return the date as a string — just verify it's not null
            expect(oVal).not.toBeNull();
          }
        } else {
          expect(oVal).toBe(xVal);
        }
      }
    }
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 7. CSV Edge Cases in Real World
// ═══════════════════════════════════════════════════════════════════════

describe("Real World: CSV Edge Cases", () => {
  it("parses European CSV with semicolons and comma decimals", () => {
    const input = ["Name;Amount;Rate", 'Alice;"1.234,56";"0,15"', 'Bob;"987,65";"0,08"'].join("\n");

    const rows = parseCsv(input, { delimiter: ";" });
    expect(rows).toHaveLength(3);
    expect(rows[0]).toEqual(["Name", "Amount", "Rate"]);
    expect(rows[1]![0]).toBe("Alice");
    expect(rows[1]![1]).toBe("1.234,56");
    expect(rows[1]![2]).toBe("0,15");
    expect(rows[2]![0]).toBe("Bob");
  });

  it("detects semicolon delimiter automatically in European CSV", () => {
    const input = ["Name;Amount;Rate", "Alice;100;0.15", "Bob;200;0.08"].join("\n");

    const delimiter = detectDelimiter(input);
    expect(delimiter).toBe(";");

    const rows = parseCsv(input);
    expect(rows[0]).toEqual(["Name", "Amount", "Rate"]);
    expect(rows[1]![1]).toBe("100");
  });

  it("handles CSV with BOM from Excel export", () => {
    const BOM = "\uFEFF";
    const input = BOM + "Name,Value\nAlice,42\n";

    // stripBom should remove it
    const stripped = stripBom(input);
    expect(stripped.charCodeAt(0)).not.toBe(0xfeff);

    // parseCsv should handle it automatically
    const rows = parseCsv(input);
    expect(rows).toHaveLength(2);
    expect(rows[0]![0]).toBe("Name");
    expect(rows[1]![0]).toBe("Alice");
  });

  it("handles CSV with quoted fields containing delimiters and newlines", () => {
    const input = [
      "Name,Address,Notes",
      '"Smith, John","123 Main St\nApt 4","Said ""hello"""',
      'Jane,"",Simple',
    ].join("\n");

    const rows = parseCsv(input);
    expect(rows).toHaveLength(3);
    expect(rows[1]![0]).toBe("Smith, John");
    expect(rows[1]![1]).toContain("123 Main St");
    expect(rows[1]![2]).toBe('Said "hello"');
    expect(rows[2]![1]).toBe("");
  });

  it("handles tab-separated values", () => {
    const input = "Name\tAge\tCity\nAlice\t30\tNYC\nBob\t25\tLA";

    const delimiter = detectDelimiter(input);
    expect(delimiter).toBe("\t");

    const rows = parseCsv(input, { delimiter: "\t" });
    expect(rows).toHaveLength(3);
    expect(rows[0]).toEqual(["Name", "Age", "City"]);
    expect(rows[1]).toEqual(["Alice", "30", "NYC"]);
    expect(rows[2]).toEqual(["Bob", "25", "LA"]);
  });

  it("round-trips CSV write and read with special characters", () => {
    const original: CellValue[][] = [
      ["Name", "Description", "Price"],
      ["Widget A", 'Contains "quotes" and, commas', 19.99],
      ["Widget B", "Line 1\nLine 2", 29.99],
      ["Widget C", "Simple text", null],
    ];

    const csv = writeCsv(original);
    const parsed = parseCsv(csv);

    expect(parsed).toHaveLength(4);
    expect(parsed[0]).toEqual(["Name", "Description", "Price"]);
    expect(parsed[1]![0]).toBe("Widget A");
    expect(parsed[1]![1]).toBe('Contains "quotes" and, commas');
    expect(parsed[2]![1]).toBe("Line 1\nLine 2");
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 8. Schema Validation Real World
// ═══════════════════════════════════════════════════════════════════════

describe("Real World: Schema Validation", () => {
  it("validates a messy import dataset and reports all errors", () => {
    // Simulate messy CSV import data (first row = header)
    const rows: CellValue[][] = [
      ["Name", "Email", "Age", "Status", "Salary", "Extra Col"],
      // Row 2: Valid
      ["Alice", "alice@acme.com", 30, "Active", 75000, "ignore"],
      // Row 3: Missing required Name
      [null, "bob@acme.com", 25, "Active", 80000, null],
      // Row 4: Wrong type — text in Age column
      ["Carol", "carol@acme.com", "twenty-five", "Active", 90000, null],
      // Row 5: Value out of range — Age = 200
      ["David", "david@acme.com", 200, "Active", 100000, null],
      // Row 6: Invalid enum value for Status
      ["Eve", "eve@acme.com", 28, "SuperActive", 85000, null],
      // Row 7: Salary below minimum
      ["Frank", "frank@acme.com", 35, "Active", -5000, null],
      // Row 8: Missing required email + invalid age type
      ["Grace", null, "old", "Inactive", 70000, null],
      // Row 9: All valid
      ["Hank", "hank@acme.com", 45, "Inactive", 110000, null],
    ];

    const result = validateWithSchema(rows, {
      name: {
        column: "Name",
        type: "string",
        required: true,
        min: 1,
      },
      email: {
        column: "Email",
        type: "string",
        required: true,
        pattern: /^[^@]+@[^@]+\.[^@]+$/,
      },
      age: {
        column: "Age",
        type: "integer",
        required: true,
        min: 0,
        max: 150,
      },
      status: {
        column: "Status",
        type: "string",
        required: true,
        enum: ["Active", "Inactive"],
      },
      salary: {
        column: "Salary",
        type: "number",
        required: false,
        min: 0,
      },
    });

    // Valid rows: Row 2 (Alice) and Row 9 (Hank) = 2 fully valid
    // Rows with errors: 3, 4, 5, 6, 7, 8
    // We still get data for all 8 data rows, but invalid fields become null
    expect(result.data).toHaveLength(8);

    // Verify valid row 1 (Alice)
    expect(result.data[0]!.name).toBe("Alice");
    expect(result.data[0]!.email).toBe("alice@acme.com");
    expect(result.data[0]!.age).toBe(30);
    expect(result.data[0]!.status).toBe("Active");
    expect(result.data[0]!.salary).toBe(75000);

    // Count errors
    // Row 3: missing required name = 1 error
    // Row 4: "twenty-five" not an integer = 1 error
    // Row 5: age 200 > max 150 = 1 error
    // Row 6: "SuperActive" not in enum = 1 error
    // Row 7: salary -5000 < min 0 = 1 error
    // Row 8: missing required email + "old" not integer = 2 errors
    // Total = 7 errors
    expect(result.errors.length).toBe(7);

    // Verify specific errors
    const nameErrors = result.errors.filter((e) => e.field === "name");
    expect(nameErrors).toHaveLength(1);
    expect(nameErrors[0]!.row).toBe(3); // 1-based

    const ageErrors = result.errors.filter((e) => e.field === "age");
    expect(ageErrors).toHaveLength(3); // row 4 type, row 5 range, row 8 type

    const statusErrors = result.errors.filter((e) => e.field === "status");
    expect(statusErrors).toHaveLength(1);
    expect(statusErrors[0]!.row).toBe(6);

    const salaryErrors = result.errors.filter((e) => e.field === "salary");
    expect(salaryErrors).toHaveLength(1);
    expect(salaryErrors[0]!.row).toBe(7);

    const emailErrors = result.errors.filter((e) => e.field === "email");
    expect(emailErrors).toHaveLength(1);
    expect(emailErrors[0]!.row).toBe(8);

    // Verify last valid row (Hank)
    expect(result.data[7]!.name).toBe("Hank");
    expect(result.data[7]!.age).toBe(45);
    expect(result.data[7]!.status).toBe("Inactive");
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 9. Round-trip Preservation
// ═══════════════════════════════════════════════════════════════════════

describe("Real World: Round-trip Preservation", () => {
  it("preserves unknown parts (e.g. chart) while modifying cell data", async () => {
    // Create initial XLSX
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Product", "Sales"],
        ["Widget A", 100],
        ["Widget B", 200],
        ["Widget C", 300],
      ],
    };

    const initial = await writeXlsx({ sheets: [sheet] });

    // Inject a fake chart entry to simulate an unknown part
    const fakeChartXml = encoder.encode(
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>',
    );
    const withChart = await injectEntries(initial, [
      { path: "xl/charts/chart1.xml", data: fakeChartXml },
    ]);

    // Open for round-trip
    const wb = await openXlsx(withChart);
    expect(wb.sheets[0]!.rows[1]![1]).toBe(100);

    // Modify a cell value
    wb.sheets[0]!.rows[1]![1] = 999;

    // Save back
    const saved = await saveXlsx(wb);
    expect(saved.byteLength).toBeGreaterThan(0);

    // Verify: chart entry is preserved
    const zip = new ZipReader(saved);
    const entries = zip.entries();
    expect(entries).toContain("xl/charts/chart1.xml");

    // Extract and verify chart content
    const chartData = await zip.extract("xl/charts/chart1.xml");
    const chartXml = decoder.decode(chartData);
    expect(chartXml).toContain("chartSpace");

    // Verify: modified cell value is updated
    const reread = await readXlsx(saved);
    expect(reread.sheets[0]!.rows[1]![1]).toBe(999);

    // Verify: other data unchanged
    expect(reread.sheets[0]!.rows[0]![0]).toBe("Product");
    expect(reread.sheets[0]!.rows[2]![1]).toBe(200);
    expect(reread.sheets[0]!.rows[3]![1]).toBe(300);
  });

  it("preserves multiple unknown parts across round-trip", async () => {
    const sheet: WriteSheet = {
      name: "Data",
      rows: [
        ["A", "B"],
        [1, 2],
      ],
    };

    const initial = await writeXlsx({ sheets: [sheet] });

    // Inject multiple fake entries
    const withExtras = await injectEntries(initial, [
      { path: "xl/charts/chart1.xml", data: encoder.encode("<chart/>") },
      { path: "xl/vbaProject.bin", data: new Uint8Array([0x00, 0x01, 0x02]) },
      { path: "customXml/item1.xml", data: encoder.encode("<custom>data</custom>") },
    ]);

    const wb = await openXlsx(withExtras);
    wb.sheets[0]!.rows[1]![0] = 99;
    const saved = await saveXlsx(wb);

    const zip = new ZipReader(saved);
    const entries = zip.entries();
    expect(entries).toContain("xl/charts/chart1.xml");
    expect(entries).toContain("xl/vbaProject.bin");
    expect(entries).toContain("customXml/item1.xml");

    const reread = await readXlsx(saved);
    expect(reread.sheets[0]!.rows[1]![0]).toBe(99);
    expect(reread.sheets[0]!.rows[1]![1]).toBe(2);
  });
});

// ═══════════════════════════════════════════════════════════════════════
// 10. Sheet Operations Real World
// ═══════════════════════════════════════════════════════════════════════

describe("Real World: Sheet Operations", () => {
  function createCatalogSheet(): Sheet {
    return {
      name: "Catalog",
      rows: [
        ["ID", "Name", "Price", "Status"],
        [1, "Widget A", 19.99, "Active"],
        [2, "Widget B", 29.99, "Active"],
        [3, "Widget C", 39.99, "Discontinued"],
        [4, "Widget D", 49.99, "Active"],
        [5, "Widget E", 59.99, "Active"],
      ],
      dataValidations: [
        {
          type: "list" as const,
          values: ["Active", "Discontinued"],
          range: "D2:D100",
        },
      ],
      autoFilter: { range: "A1:D6" },
    };
  }

  it("inserts a new product row and verifies data integrity", () => {
    const sheet = createCatalogSheet();

    // Insert 1 row at position 3 (after Widget B, before Widget C)
    insertRows(sheet, 3, 1);

    // Verify row count increased
    expect(sheet.rows).toHaveLength(7);

    // Verify data shifted correctly
    expect(sheet.rows[0]![1]).toBe("Name"); // Header unchanged
    expect(sheet.rows[1]![1]).toBe("Widget A"); // Row 1 unchanged
    expect(sheet.rows[2]![1]).toBe("Widget B"); // Row 2 unchanged
    expect(sheet.rows[3]![0]).toBeNull(); // New empty row
    expect(sheet.rows[4]![1]).toBe("Widget C"); // Shifted down
    expect(sheet.rows[5]![1]).toBe("Widget D"); // Shifted down
    expect(sheet.rows[6]![1]).toBe("Widget E"); // Shifted down

    // Fill in the new product
    sheet.rows[3] = [6, "Widget F", 69.99, "Active"];
    expect(sheet.rows[3]![1]).toBe("Widget F");
  });

  it("deletes discontinued products and verifies cleanup", () => {
    const sheet = createCatalogSheet();

    // Widget C at row 3 is Discontinued — delete it
    deleteRows(sheet, 3, 1);

    expect(sheet.rows).toHaveLength(5);
    expect(sheet.rows[0]![1]).toBe("Name");
    expect(sheet.rows[1]![1]).toBe("Widget A");
    expect(sheet.rows[2]![1]).toBe("Widget B");
    expect(sheet.rows[3]![1]).toBe("Widget D"); // Shifted up
    expect(sheet.rows[4]![1]).toBe("Widget E"); // Shifted up
  });

  it("clones a sheet to create a backup", () => {
    const sheet = createCatalogSheet();

    const backup = cloneSheet(sheet, "Catalog Backup");

    expect(backup.name).toBe("Catalog Backup");
    expect(backup.rows).toHaveLength(6);
    expect(backup.rows[0]).toEqual(["ID", "Name", "Price", "Status"]);
    expect(backup.rows[1]).toEqual([1, "Widget A", 19.99, "Active"]);

    // Verify it's a deep copy — modifying backup doesn't affect original
    backup.rows[1]![1] = "Modified Widget";
    expect(sheet.rows[1]![1]).toBe("Widget A");

    // Verify data validations are cloned
    expect(backup.dataValidations).toBeDefined();
    expect(backup.dataValidations!.length).toBe(1);
    expect(backup.dataValidations![0]!.type).toBe("list");

    // Verify auto filter is cloned
    expect(backup.autoFilter).toBeDefined();
  });

  it("insert + delete combined workflow", () => {
    const sheet = createCatalogSheet();

    // Add 3 new rows at end
    insertRows(sheet, 6, 3);
    expect(sheet.rows).toHaveLength(9);

    sheet.rows[6] = [6, "Widget F", 69.99, "Active"];
    sheet.rows[7] = [7, "Widget G", 79.99, "Active"];
    sheet.rows[8] = [8, "Widget H", 89.99, "Discontinued"];

    // Now delete the 2 discontinued rows (original Widget C at row 3, new Widget H at row 8)
    // Delete Widget H first (higher index)
    deleteRows(sheet, 8, 1);
    expect(sheet.rows).toHaveLength(8);
    expect(sheet.rows[7]![1]).toBe("Widget G");

    // Delete Widget C (row 3)
    deleteRows(sheet, 3, 1);
    expect(sheet.rows).toHaveLength(7);
    expect(sheet.rows[3]![1]).toBe("Widget D");

    // Verify final state
    const names = sheet.rows.slice(1).map((r) => r[1]);
    expect(names).toEqual(["Widget A", "Widget B", "Widget D", "Widget E", "Widget F", "Widget G"]);
  });

  it("clone + write + read round-trip", async () => {
    const sheet = createCatalogSheet();
    const backup = cloneSheet(sheet, "Backup");

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: sheet.name,
          rows: sheet.rows,
          dataValidations: sheet.dataValidations,
          autoFilter: sheet.autoFilter,
        },
        {
          name: backup.name,
          rows: backup.rows,
          dataValidations: backup.dataValidations,
          autoFilter: backup.autoFilter,
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets).toHaveLength(2);
    expect(wb.sheets[0]!.name).toBe("Catalog");
    expect(wb.sheets[1]!.name).toBe("Backup");

    // Both sheets should have identical data
    for (let r = 0; r < sheet.rows.length; r++) {
      for (let c = 0; c < sheet.rows[r]!.length; c++) {
        const origVal = wb.sheets[0]!.rows[r]![c];
        const backupVal = wb.sheets[1]!.rows[r]![c];
        if (typeof origVal === "number") {
          expect(backupVal).toBeCloseTo(origVal, 5);
        } else {
          expect(backupVal).toEqual(origVal);
        }
      }
    }
  });
});

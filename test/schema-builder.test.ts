import { describe, expect, it } from "vitest";
import {
  writeObjects,
  readXlsx,
  fx,
  pickColumns,
  omitColumns,
  applyPreset,
  slate,
  ocean,
  forest,
  rose,
  minimal,
  WorkbookBuilder,
} from "../src/index";
import type { ColumnDef, ColumnSummary, CellValue } from "../src/index";

// ── Helpers ────────────────────────────────────────────────────────

async function roundTrip<T extends Record<string, unknown>>(
  data: T[],
  options?: { columns?: ColumnDef<T>[]; sheetName?: string },
) {
  const xlsx = await writeObjects(data, options);
  const wb = await readXlsx(xlsx, { readStyles: true });
  return wb.sheets[0]!;
}

// ── writeObjects backward compatibility ───────────────────────────

describe("writeObjects backward compatibility", () => {
  it("infers columns from object keys when no columns provided", async () => {
    const data = [
      { name: "Alice", age: 30 },
      { name: "Bob", age: 25 },
    ];
    const sheet = await roundTrip(data);
    expect(sheet.rows.length).toBe(3); // header + 2 data rows
    expect(sheet.rows[0]).toEqual(["name", "age"]);
    expect(sheet.rows[1]).toEqual(["Alice", 30]);
    expect(sheet.rows[2]).toEqual(["Bob", 25]);
  });

  it("handles empty data", async () => {
    const xlsx = await writeObjects([]);
    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0]!.rows.length).toBe(0);
  });
});

// ── Value accessors ───────────────────────────────────────────────

describe("value accessors", () => {
  it("supports key-based accessor", async () => {
    const data = [{ product: "Widget", price: 9.99 }];
    const sheet = await roundTrip(data, {
      columns: [
        { key: "product", header: "Product" },
        { key: "price", header: "Price" },
      ],
    });
    expect(sheet.rows[1]).toEqual(["Widget", 9.99]);
  });

  it("supports dot-path accessor for nested objects", async () => {
    const data = [{ user: { name: "Alice", address: { city: "Istanbul" } }, score: 95 }];
    const sheet = await roundTrip(data, {
      columns: [
        { value: "user.name", header: "Name" },
        { value: "user.address.city", header: "City" },
        { key: "score", header: "Score" },
      ],
    });
    expect(sheet.rows[1]).toEqual(["Alice", "Istanbul", 95]);
  });

  it("supports function accessor", async () => {
    const data = [{ qty: 10, price: 5 }];
    const sheet = await roundTrip(data, {
      columns: [
        { key: "qty", header: "Qty" },
        { value: (item) => (item.qty as number) * (item.price as number), header: "Total" },
      ],
    });
    expect(sheet.rows[1]).toEqual([10, 50]);
  });
});

// ── Transform + defaultValue ──────────────────────────────────────

describe("transform and defaultValue", () => {
  it("applies transform to extracted value", async () => {
    const data = [{ status: "A" }, { status: "I" }];
    const sheet = await roundTrip(data, {
      columns: [
        {
          key: "status",
          header: "Status",
          transform: (v) => (v === "A" ? "Active" : "Inactive"),
        },
      ],
    });
    expect(sheet.rows[1]).toEqual(["Active"]);
    expect(sheet.rows[2]).toEqual(["Inactive"]);
  });

  it("uses defaultValue when accessor returns null", async () => {
    const data = [{ name: "Alice", notes: null }, { name: "Bob" }];
    const sheet = await roundTrip(data, {
      columns: [
        { key: "name", header: "Name" },
        { key: "notes", header: "Notes", defaultValue: "N/A" },
      ],
    });
    expect(sheet.rows[1]![1]).toBe("N/A");
    expect(sheet.rows[2]![1]).toBe("N/A");
  });
});

// ── Formula columns ───────────────────────────────────────────────

describe("formula columns", () => {
  it("generates formulas per row", async () => {
    const data = [
      { qty: 10, price: 5 },
      { qty: 20, price: 3 },
    ];
    const xlsx = await writeObjects(data, {
      columns: [
        { key: "qty", header: "Qty" },
        { key: "price", header: "Price" },
        { header: "Total", formula: (row) => `A${row}*B${row}` },
      ],
    });
    const wb = await readXlsx(xlsx);
    const sheet = wb.sheets[0]!;
    // Formula columns should have formulas in cells
    const cells = sheet.cells!;
    // Row 2 (1-based) = index 1 in 0-based resolved rows, which is data row 0
    const formulaCell1 = cells.get("1,2");
    const formulaCell2 = cells.get("2,2");
    expect(formulaCell1?.formula).toBe("A2*B2");
    expect(formulaCell2?.formula).toBe("A3*B3");
  });
});

// ── Summary rows ──────────────────────────────────────────────────

describe("summary rows", () => {
  it("appends sum summary row", async () => {
    const data = [
      { product: "A", revenue: 100 },
      { product: "B", revenue: 200 },
    ];
    const xlsx = await writeObjects(data, {
      columns: [
        { key: "product", header: "Product", summary: { label: "Total" } },
        { key: "revenue", header: "Revenue", summary: { fn: "sum" } },
      ],
    });
    const wb = await readXlsx(xlsx);
    const sheet = wb.sheets[0]!;
    // Last row should be summary
    expect(sheet.rows.length).toBe(4); // header + 2 data + summary
    expect(sheet.rows[3]![0]).toBe("Total");
    // Check formula
    const summaryCell = sheet.cells?.get("3,1");
    expect(summaryCell?.formula).toBe("SUM(B2:B3)");
  });

  it("supports custom summary formula", async () => {
    const data = [{ val: 10 }, { val: 20 }];
    const xlsx = await writeObjects(data, {
      columns: [
        {
          key: "val",
          header: "Value",
          summary: { custom: (range) => `SUMPRODUCT(${range})` },
        },
      ],
    });
    const wb = await readXlsx(xlsx);
    const cell = wb.sheets[0]!.cells?.get("3,0");
    expect(cell?.formula).toBe("SUMPRODUCT(A2:A3)");
  });
});

// ── Conditional styles (when) ─────────────────────────────────────

describe("conditional styles", () => {
  it("applies style when condition is true", async () => {
    const data = [{ val: -5 }, { val: 10 }];
    const xlsx = await writeObjects(data, {
      columns: [
        {
          key: "val",
          header: "Value",
          when: {
            test: (v) => typeof v === "number" && v < 0,
            style: { font: { bold: true, color: { rgb: "FF0000" } } },
          },
        },
      ],
    });
    const wb = await readXlsx(xlsx, { readStyles: true });
    const negativeCell = wb.sheets[0]!.cells?.get("1,0");
    const positiveCell = wb.sheets[0]!.cells?.get("2,0");
    expect(negativeCell?.style?.font?.color?.rgb).toBe("FF0000");
    // Positive cell should NOT have the red color
    expect(positiveCell?.style?.font?.color?.rgb).not.toBe("FF0000");
  });
});

// ── Column groups (children) ──────────────────────────────────────

describe("column groups", () => {
  it("generates multi-row headers with merges", async () => {
    const data = [{ arr: 100, nrr: 0.95, dau: 500 }];
    const xlsx = await writeObjects(data, {
      columns: [
        {
          header: "Commercial",
          children: [
            { key: "arr", header: "ARR" },
            { key: "nrr", header: "NRR" },
          ],
        },
        { key: "dau", header: "DAU" },
      ],
    });
    const wb = await readXlsx(xlsx);
    const sheet = wb.sheets[0]!;
    // Should have 2 header rows + 1 data row = 3 rows
    expect(sheet.rows.length).toBe(3);
    // First row: "Commercial", null (merged), "DAU"
    expect(sheet.rows[0]![0]).toBe("Commercial");
    // Second row: "ARR", "NRR", null (merged vertically with row 0)
    expect(sheet.rows[1]![0]).toBe("ARR");
    expect(sheet.rows[1]![1]).toBe("NRR");
    // Data row
    // Data row may have trailing null from reader padding
    expect(sheet.rows[2]!.slice(0, 3)).toEqual([100, 0.95, 500]);
    // Should have merges
    expect(sheet.merges).toBeDefined();
    expect(sheet.merges!.length).toBeGreaterThan(0);
  });
});

// ── Sub-row expansion ─────────────────────────────────────────────

describe("sub-row expansion", () => {
  it("expands array values into sub-rows", async () => {
    const data = [
      {
        id: "ORD-1",
        items: [
          { name: "Widget", qty: 10 },
          { name: "Gadget", qty: 5 },
        ],
      },
      { id: "ORD-2", items: [{ name: "Sprocket", qty: 3 }] },
    ];
    const xlsx = await writeObjects(data, {
      columns: [
        { key: "id", header: "Order" },
        {
          header: "Product",
          expand: (row) => (row.items as Array<{ name: string }>).map((i) => i.name),
        },
        { header: "Qty", expand: (row) => (row.items as Array<{ qty: number }>).map((i) => i.qty) },
      ],
    });
    const wb = await readXlsx(xlsx);
    const sheet = wb.sheets[0]!;
    // Header + 2 sub-rows for ORD-1 + 1 row for ORD-2 = 4 rows
    expect(sheet.rows.length).toBe(4);
    expect(sheet.rows[1]).toEqual(["ORD-1", "Widget", 10]);
    expect(sheet.rows[2]![1]).toBe("Gadget");
    expect(sheet.rows[3]).toEqual(["ORD-2", "Sprocket", 3]);
    // ORD-1's "id" cell should be merged across 2 rows
    expect(sheet.merges).toBeDefined();
    const idMerge = sheet.merges!.find(
      (m) => m.startCol === 0 && m.startRow === 1 && m.endRow === 2,
    );
    expect(idMerge).toBeDefined();
  });
});

// ── fx helpers ────────────────────────────────────────────────────

describe("fx formula helpers", () => {
  it("sum", () => expect(fx.sum("A1:A10")).toBe("SUM(A1:A10)"));
  it("average", () => expect(fx.average("B1:B10")).toBe("AVERAGE(B1:B10)"));
  it("count", () => expect(fx.count("C1:C10")).toBe("COUNT(C1:C10)"));
  it("min", () => expect(fx.min("D1:D10")).toBe("MIN(D1:D10)"));
  it("max", () => expect(fx.max("E1:E10")).toBe("MAX(E1:E10)"));
  it("round", () => expect(fx.round("A1*B1", 2)).toBe("ROUND(A1*B1,2)"));
  it("abs", () => expect(fx.abs("A1-B1")).toBe("ABS(A1-B1)"));
  it("safeDiv", () => expect(fx.safeDiv("A1", "B1")).toBe("IF(B1=0,0,A1/B1)"));
  it("safeDiv with fallback", () => expect(fx.safeDiv("A1", "B1", '""')).toBe('IF(B1=0,"",A1/B1)'));
  it("iif", () => expect(fx.iif("A1>100", '"High"', '"Low"')).toBe('IF(A1>100,"High","Low")'));
  it("ifError", () => expect(fx.ifError("A1/B1", 0)).toBe("IFERROR(A1/B1,0)"));
  it("concat", () => expect(fx.concat("A1", '" "', "B1")).toBe('CONCATENATE(A1," ",B1)'));
  it("vlookup", () =>
    expect(fx.vlookup("A1", "Sheet2!A:C", 3)).toBe("VLOOKUP(A1,Sheet2!A:C,3,FALSE)"));
  it("col helper", () => {
    const C = fx.col("C");
    expect(C(5)).toBe("C5");
    expect(C(10)).toBe("C10");
  });
  it("and", () => expect(fx.and("A1>0", "B1>0")).toBe("AND(A1>0,B1>0)"));
  it("or", () => expect(fx.or("A1>0", "B1>0")).toBe("OR(A1>0,B1>0)"));
  it("not", () => expect(fx.not("A1>0")).toBe("NOT(A1>0)"));
  it("countA", () => expect(fx.countA("A1:A10")).toBe("COUNTA(A1:A10)"));
  it("sumIf", () => expect(fx.sumIf("A1:A10", '">0"')).toBe('SUMIF(A1:A10,">0")'));
  it("countIf", () => expect(fx.countIf("A1:A10", '"Yes"')).toBe('COUNTIF(A1:A10,"Yes")'));
});

// ── pickColumns / omitColumns ─────────────────────────────────────

describe("column utilities", () => {
  const cols: ColumnDef[] = [
    { key: "a", header: "A" },
    { key: "b", header: "B" },
    { key: "c", header: "C" },
  ];

  it("pickColumns selects by key in order", () => {
    const picked = pickColumns(cols, ["c", "a"]);
    expect(picked.map((c) => c.key)).toEqual(["c", "a"]);
  });

  it("omitColumns removes by key", () => {
    const omitted = omitColumns(cols, ["b"]);
    expect(omitted.map((c) => c.key)).toEqual(["a", "c"]);
  });

  it("pickColumns ignores missing keys", () => {
    const picked = pickColumns(cols, ["a", "z"]);
    expect(picked.length).toBe(1);
    expect(picked[0]!.key).toBe("a");
  });
});

// ── Style presets ─────────────────────────────────────────────────

describe("style presets", () => {
  it("all presets have required fields", () => {
    for (const preset of [slate, ocean, forest, rose, minimal]) {
      expect(preset.header).toBeDefined();
      expect(preset.data).toBeDefined();
    }
  });

  it("applyPreset sets headerStyle and style", () => {
    const cols: ColumnDef[] = [
      { key: "a", header: "A" },
      { key: "b", header: "B", style: { font: { bold: true } } },
    ];
    const result = applyPreset(cols, slate);
    // First column gets preset header/data
    expect(result[0]!.headerStyle).toBe(slate.header);
    expect(result[0]!.style).toBe(slate.data);
    // Second column keeps its own style
    expect(result[1]!.style).toEqual({ font: { bold: true } });
    expect(result[1]!.headerStyle).toBe(slate.header);
  });

  it("applyPreset sets summary style", () => {
    const cols: ColumnDef[] = [{ key: "a", header: "A", summary: { fn: "sum" } }];
    const result = applyPreset(cols, slate);
    expect(result[0]!.summary!.style).toBe(slate.summary);
  });
});

// ── WorkbookBuilder objectRows ────────────────────────────────────

describe("WorkbookBuilder objectRows", () => {
  it("creates sheet from object data with columns", async () => {
    const data = [
      { name: "Widget", price: 9.99 },
      { name: "Gadget", price: 24.5 },
    ];
    const xlsx = await WorkbookBuilder.create()
      .addSheet("Products")
      .objectRows(data, [
        { key: "name", header: "Name" },
        { key: "price", header: "Price", numFmt: "$#,##0.00" },
      ])
      .build();

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0]!.rows[0]).toEqual(["Name", "Price"]);
    expect(wb.sheets[0]!.rows[1]).toEqual(["Widget", 9.99]);
  });
});

// ── headerStyle ───────────────────────────────────────────────────

describe("headerStyle", () => {
  it("applies separate style to header vs data cells", async () => {
    const data = [{ val: 42 }];
    const xlsx = await writeObjects(data, {
      columns: [
        {
          key: "val",
          header: "Value",
          headerStyle: { font: { bold: true, color: { rgb: "FF0000" } } },
          style: { font: { italic: true } },
        },
      ],
    });
    const wb = await readXlsx(xlsx, { readStyles: true });
    const headerCell = wb.sheets[0]!.cells?.get("0,0");
    const dataCell = wb.sheets[0]!.cells?.get("1,0");
    expect(headerCell?.style?.font?.bold).toBe(true);
    expect(headerCell?.style?.font?.color?.rgb).toBe("FF0000");
    expect(dataCell?.style?.font?.italic).toBe(true);
  });
});

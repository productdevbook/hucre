import { describe, it, expect } from "vitest";
import { WorkbookBuilder } from "../src/builder";
import { readXlsx } from "../src/xlsx/reader";

// ── Helper ──────────────────────────────────────────────────────────

async function roundTrip(data: Uint8Array) {
  return readXlsx(data);
}

// ── Tests ────────────────────────────────────────────────────────────

describe("WorkbookBuilder — basic", () => {
  it("creates a workbook with one sheet and rows", async () => {
    const xlsx = await WorkbookBuilder.create()
      .addSheet("Sheet1")
      .row(["Hello", 42])
      .row(["World", 99])
      .build();

    const wb = await roundTrip(xlsx);
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].name).toBe("Sheet1");
    expect(wb.sheets[0].rows[0][0]).toBe("Hello");
    expect(wb.sheets[0].rows[0][1]).toBe(42);
    expect(wb.sheets[0].rows[1][0]).toBe("World");
    expect(wb.sheets[0].rows[1][1]).toBe(99);
  });

  it("static create() returns a WorkbookBuilder", () => {
    const wb = WorkbookBuilder.create();
    expect(wb).toBeInstanceOf(WorkbookBuilder);
  });
});

describe("WorkbookBuilder — method chaining", () => {
  it("fluent API chains methods", async () => {
    const xlsx = await WorkbookBuilder.create()
      .properties({ title: "Test Workbook", creator: "hucre" })
      .defaultFont({ name: "Arial", size: 11 })
      .addSheet("Data")
      .row(["A", "B", "C"])
      .row([1, 2, 3])
      .done()
      .build();

    const wb = await roundTrip(xlsx);
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].rows).toHaveLength(2);
  });

  it("build() can be called directly on SheetBuilder", async () => {
    const xlsx = await WorkbookBuilder.create().addSheet("Quick").row(["fast"]).build();

    const wb = await roundTrip(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe("fast");
  });
});

describe("WorkbookBuilder — multiple sheets", () => {
  it("creates a workbook with multiple sheets", async () => {
    const xlsx = await WorkbookBuilder.create()
      .addSheet("First")
      .row(["Sheet 1 data"])
      .done()
      .addSheet("Second")
      .row(["Sheet 2 data"])
      .done()
      .addSheet("Third")
      .row(["Sheet 3 data"])
      .done()
      .build();

    const wb = await roundTrip(xlsx);
    expect(wb.sheets).toHaveLength(3);
    expect(wb.sheets[0].name).toBe("First");
    expect(wb.sheets[1].name).toBe("Second");
    expect(wb.sheets[2].name).toBe("Third");
    expect(wb.sheets[0].rows[0][0]).toBe("Sheet 1 data");
    expect(wb.sheets[1].rows[0][0]).toBe("Sheet 2 data");
    expect(wb.sheets[2].rows[0][0]).toBe("Sheet 3 data");
  });
});

describe("WorkbookBuilder — columns", () => {
  it("sets column definitions", async () => {
    const xlsx = await WorkbookBuilder.create()
      .addSheet("WithCols")
      .columns([
        { header: "Name", width: 20 },
        { header: "Age", width: 10 },
      ])
      .row(["Alice", 30])
      .row(["Bob", 25])
      .build();

    const wb = await roundTrip(xlsx);
    expect(wb.sheets[0].columns).toBeDefined();
    expect(wb.sheets[0].columns).toHaveLength(2);
    expect(wb.sheets[0].columns![0].width).toBe(20);
    expect(wb.sheets[0].columns![1].width).toBe(10);
  });

  it("adds individual column definitions via column()", async () => {
    const xlsx = await WorkbookBuilder.create()
      .addSheet("IndivCols")
      .column({ header: "X", width: 15 })
      .column({ header: "Y", width: 12 })
      .row([1, 2])
      .build();

    const wb = await roundTrip(xlsx);
    expect(wb.sheets[0].columns).toHaveLength(2);
  });
});

describe("WorkbookBuilder — merges", () => {
  it("creates merge ranges", async () => {
    const xlsx = await WorkbookBuilder.create()
      .addSheet("Merges")
      .row(["Merged Header", null, null])
      .row(["A", "B", "C"])
      .merge(0, 0, 0, 2)
      .build();

    const wb = await roundTrip(xlsx);
    expect(wb.sheets[0].merges).toBeDefined();
    expect(wb.sheets[0].merges).toHaveLength(1);
    expect(wb.sheets[0].merges![0]).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 2,
    });
  });
});

describe("WorkbookBuilder — freeze panes", () => {
  it("freezes top row", async () => {
    const xlsx = await WorkbookBuilder.create()
      .addSheet("Frozen")
      .row(["Header A", "Header B"])
      .row([1, 2])
      .freeze(1)
      .build();

    const wb = await roundTrip(xlsx);
    expect(wb.sheets[0].freezePane).toBeDefined();
    expect(wb.sheets[0].freezePane!.rows).toBe(1);
  });

  it("freezes rows and columns", async () => {
    const xlsx = await WorkbookBuilder.create()
      .addSheet("FrozenBoth")
      .row(["A", "B", "C"])
      .freeze(1, 2)
      .build();

    const wb = await roundTrip(xlsx);
    expect(wb.sheets[0].freezePane!.rows).toBe(1);
    expect(wb.sheets[0].freezePane!.columns).toBe(2);
  });
});

describe("WorkbookBuilder — rows()", () => {
  it("adds multiple rows via rows()", async () => {
    const data = [
      ["A", 1],
      ["B", 2],
      ["C", 3],
    ] as const;

    const xlsx = await WorkbookBuilder.create()
      .addSheet("Bulk")
      .rows(data.map((r) => [...r]))
      .build();

    const wb = await roundTrip(xlsx);
    expect(wb.sheets[0].rows).toHaveLength(3);
    expect(wb.sheets[0].rows[2][0]).toBe("C");
  });
});

describe("WorkbookBuilder — validation", () => {
  it("adds data validation", async () => {
    const xlsx = await WorkbookBuilder.create()
      .addSheet("Validated")
      .row(["Status"])
      .row(["Active"])
      .validation({
        type: "list",
        values: ["Active", "Inactive", "Pending"],
        range: "A2:A100",
      })
      .build();

    const wb = await roundTrip(xlsx);
    expect(wb.sheets[0].dataValidations).toBeDefined();
    expect(wb.sheets[0].dataValidations).toHaveLength(1);
  });
});

import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import type { WriteSheet, Cell } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data);
  const raw = await zip.extract(path);
  return decoder.decode(raw);
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

/** Helper to find all <c> elements in a worksheet XML document */
function findCellElements(doc: any): any[] {
  const sheetData = findChild(doc, "sheetData");
  if (!sheetData) return [];
  const rows = findChildren(sheetData, "row");
  const cells: any[] = [];
  for (const row of rows) {
    cells.push(...findChildren(row, "c"));
  }
  return cells;
}

// ── Writing: Shared Formulas ─────────────────────────────────────────

describe("shared formula writing", () => {
  it("writes shared formula master cell with t, si, ref, and formula text", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("1,1", {
      formula: "A2*2",
      formulaType: "shared",
      formulaSharedIndex: 0,
      formulaRef: "B2:B10",
      formulaResult: 10,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Val", "Doubled"], [5]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(xlsx, "xl/worksheets/sheet1.xml");
    const doc = parseXml(xml);

    const allCells = findCellElements(doc);
    // Find B2 cell (r="B2")
    const b2 = allCells.find((c: any) => c.attrs["r"] === "B2");
    expect(b2).toBeDefined();

    const fEl = findChild(b2, "f");
    expect(fEl).toBeDefined();
    expect(fEl.attrs["t"]).toBe("shared");
    expect(fEl.attrs["si"]).toBe("0");
    expect(fEl.attrs["ref"]).toBe("B2:B10");
    expect(getElementText(fEl)).toBe("A2*2");
  });

  it("writes shared formula slave cell with t, si, and self-closing f", async () => {
    const cells = new Map<string, Partial<Cell>>();
    // Slave cell: formula text is empty string, has si but no ref
    cells.set("2,1", {
      formula: "",
      formulaType: "shared",
      formulaSharedIndex: 0,
      formulaResult: 14,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Val", "Doubled"], [5], [7]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(xlsx, "xl/worksheets/sheet1.xml");

    // The slave cell should have <f t="shared" si="0"/>  (self-closing)
    expect(xml).toContain('t="shared"');
    expect(xml).toContain('si="0"');

    const doc = parseXml(xml);
    const allCells = findCellElements(doc);
    const b3 = allCells.find((c: any) => c.attrs["r"] === "B3");
    expect(b3).toBeDefined();

    const fEl = findChild(b3, "f");
    expect(fEl).toBeDefined();
    expect(fEl.attrs["t"]).toBe("shared");
    expect(fEl.attrs["si"]).toBe("0");
    // Slave cell should not have ref attribute
    expect(fEl.attrs["ref"]).toBeUndefined();
  });

  it("writes master + slave shared formula cells together", async () => {
    const cells = new Map<string, Partial<Cell>>();
    // Master cell B2: has formula text and ref
    cells.set("1,1", {
      formula: "A2*2",
      formulaType: "shared",
      formulaSharedIndex: 0,
      formulaRef: "B2:B4",
      formulaResult: 10,
    });
    // Slave cell B3
    cells.set("2,1", {
      formula: "",
      formulaType: "shared",
      formulaSharedIndex: 0,
      formulaResult: 14,
    });
    // Slave cell B4
    cells.set("3,1", {
      formula: "",
      formulaType: "shared",
      formulaSharedIndex: 0,
      formulaResult: 20,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Val", "Doubled"], [5], [7], [10]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(xlsx, "xl/worksheets/sheet1.xml");
    const doc = parseXml(xml);
    const allCells = findCellElements(doc);

    // B2 = master
    const b2f = findChild(
      allCells.find((c: any) => c.attrs["r"] === "B2"),
      "f",
    );
    expect(b2f.attrs["t"]).toBe("shared");
    expect(b2f.attrs["ref"]).toBe("B2:B4");
    expect(getElementText(b2f)).toBe("A2*2");

    // B3, B4 = slaves
    for (const ref of ["B3", "B4"]) {
      const cellEl = allCells.find((c: any) => c.attrs["r"] === ref);
      const fEl = findChild(cellEl, "f");
      expect(fEl.attrs["t"]).toBe("shared");
      expect(fEl.attrs["si"]).toBe("0");
      expect(fEl.attrs["ref"]).toBeUndefined();
    }
  });
});

// ── Writing: Array Formulas ──────────────────────────────────────────

describe("array formula writing", () => {
  it("writes array formula with t=array and ref", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("1,1", {
      formula: "SUM(A2:A10*C2:C10)",
      formulaType: "array",
      formulaRef: "B2:B10",
      formulaResult: 100,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["A", "B"], [1]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(xlsx, "xl/worksheets/sheet1.xml");
    const doc = parseXml(xml);

    const allCells = findCellElements(doc);
    const b2 = allCells.find((c: any) => c.attrs["r"] === "B2");
    const fEl = findChild(b2, "f");

    expect(fEl.attrs["t"]).toBe("array");
    expect(fEl.attrs["ref"]).toBe("B2:B10");
    expect(getElementText(fEl)).toBe("SUM(A2:A10*C2:C10)");
    // No cm attribute for non-dynamic
    expect(fEl.attrs["cm"]).toBeUndefined();
  });

  it("writes dynamic array formula with cm=1", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("1,1", {
      formula: "SORT(A2:A10)",
      formulaType: "array",
      formulaRef: "B2:B10",
      formulaDynamic: true,
      formulaResult: 1,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["A", "B"], [3]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(xlsx, "xl/worksheets/sheet1.xml");
    const doc = parseXml(xml);

    const allCells = findCellElements(doc);
    const b2 = allCells.find((c: any) => c.attrs["r"] === "B2");
    const fEl = findChild(b2, "f");

    expect(fEl.attrs["t"]).toBe("array");
    expect(fEl.attrs["ref"]).toBe("B2:B10");
    expect(fEl.attrs["cm"]).toBe("1");
    expect(getElementText(fEl)).toBe("SORT(A2:A10)");
  });
});

// ── Writing: Normal formula (backward compat) ────────────────────────

describe("normal formula writing (backward compatibility)", () => {
  it("writes normal formula without t attribute", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,1", {
      formula: "SUM(A1:A10)",
      formulaResult: 55,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [[1]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(xlsx, "xl/worksheets/sheet1.xml");
    const doc = parseXml(xml);

    const allCells = findCellElements(doc);
    const b1 = allCells.find((c: any) => c.attrs["r"] === "B1");
    const fEl = findChild(b1, "f");

    expect(fEl).toBeDefined();
    // Normal formula: no t, si, ref, or cm attributes
    expect(fEl.attrs["t"]).toBeUndefined();
    expect(fEl.attrs["si"]).toBeUndefined();
    expect(fEl.attrs["ref"]).toBeUndefined();
    expect(fEl.attrs["cm"]).toBeUndefined();
    expect(getElementText(fEl)).toBe("SUM(A1:A10)");
  });
});

// ── Reading: Shared Formulas ─────────────────────────────────────────

describe("shared formula reading", () => {
  it("reads shared formula master cell with type, si, and ref", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("1,1", {
      formula: "A2*2",
      formulaType: "shared",
      formulaSharedIndex: 0,
      formulaRef: "B2:B4",
      formulaResult: 10,
    });
    cells.set("2,1", {
      formula: "",
      formulaType: "shared",
      formulaSharedIndex: 0,
      formulaResult: 14,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Val", "Doubled"], [5], [7]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(xlsx);
    const sheetCells = workbook.sheets[0].cells!;

    // Master cell B2 = "1,1"
    const master = sheetCells.get("1,1");
    expect(master).toBeDefined();
    expect(master!.formula).toBe("A2*2");
    expect(master!.formulaType).toBe("shared");
    expect(master!.formulaSharedIndex).toBe(0);
    expect(master!.formulaRef).toBe("B2:B4");

    // Slave cell B3 = "2,1"
    const slave = sheetCells.get("2,1");
    expect(slave).toBeDefined();
    expect(slave!.formula).toBe("");
    expect(slave!.formulaType).toBe("shared");
    expect(slave!.formulaSharedIndex).toBe(0);
    // Slave has no ref
    expect(slave!.formulaRef).toBeUndefined();
  });
});

// ── Reading: Array Formulas ──────────────────────────────────────────

describe("array formula reading", () => {
  it("reads array formula with type and ref", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("1,1", {
      formula: "SUM(A2:A10*C2:C10)",
      formulaType: "array",
      formulaRef: "B2:B10",
      formulaResult: 100,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["A", "B"], [1]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(xlsx);
    const sheetCells = workbook.sheets[0].cells!;

    const cell = sheetCells.get("1,1");
    expect(cell).toBeDefined();
    expect(cell!.formula).toBe("SUM(A2:A10*C2:C10)");
    expect(cell!.formulaType).toBe("array");
    expect(cell!.formulaRef).toBe("B2:B10");
    expect(cell!.formulaDynamic).toBeUndefined();
  });

  it("reads dynamic array formula with formulaDynamic flag", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("1,1", {
      formula: "SORT(A2:A10)",
      formulaType: "array",
      formulaRef: "B2:B10",
      formulaDynamic: true,
      formulaResult: 1,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["A", "B"], [3]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(xlsx);
    const sheetCells = workbook.sheets[0].cells!;

    const cell = sheetCells.get("1,1");
    expect(cell).toBeDefined();
    expect(cell!.formula).toBe("SORT(A2:A10)");
    expect(cell!.formulaType).toBe("array");
    expect(cell!.formulaRef).toBe("B2:B10");
    expect(cell!.formulaDynamic).toBe(true);
  });
});

// ── Round-trip: Shared Formulas ──────────────────────────────────────

describe("shared formula round-trip", () => {
  it("preserves shared formula metadata through write → read", async () => {
    const cells = new Map<string, Partial<Cell>>();
    // Master
    cells.set("1,1", {
      formula: "A2+10",
      formulaType: "shared",
      formulaSharedIndex: 0,
      formulaRef: "B2:B5",
      formulaResult: 20,
    });
    // Slaves
    for (let r = 2; r <= 4; r++) {
      cells.set(`${r},1`, {
        formula: "",
        formulaType: "shared",
        formulaSharedIndex: 0,
        formulaResult: (r + 1) * 10,
      });
    }

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Input", "Output"], [10], [20], [30], [40]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(xlsx);
    const readCells = workbook.sheets[0].cells!;

    // Verify master
    const master = readCells.get("1,1");
    expect(master!.formulaType).toBe("shared");
    expect(master!.formulaSharedIndex).toBe(0);
    expect(master!.formulaRef).toBe("B2:B5");
    expect(master!.formula).toBe("A2+10");

    // Verify slaves
    for (let r = 2; r <= 4; r++) {
      const slave = readCells.get(`${r},1`);
      expect(slave).toBeDefined();
      expect(slave!.formulaType).toBe("shared");
      expect(slave!.formulaSharedIndex).toBe(0);
      expect(slave!.formulaRef).toBeUndefined();
      expect(slave!.formula).toBe("");
    }
  });

  it("preserves multiple shared formula groups", async () => {
    const cells = new Map<string, Partial<Cell>>();
    // Group 0: B column
    cells.set("1,1", {
      formula: "A2*2",
      formulaType: "shared",
      formulaSharedIndex: 0,
      formulaRef: "B2:B3",
      formulaResult: 10,
    });
    cells.set("2,1", {
      formula: "",
      formulaType: "shared",
      formulaSharedIndex: 0,
      formulaResult: 14,
    });

    // Group 1: C column
    cells.set("1,2", {
      formula: "A2*3",
      formulaType: "shared",
      formulaSharedIndex: 1,
      formulaRef: "C2:C3",
      formulaResult: 15,
    });
    cells.set("2,2", {
      formula: "",
      formulaType: "shared",
      formulaSharedIndex: 1,
      formulaResult: 21,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Val", "x2", "x3"], [5], [7]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(xlsx);
    const readCells = workbook.sheets[0].cells!;

    // Group 0
    expect(readCells.get("1,1")!.formulaSharedIndex).toBe(0);
    expect(readCells.get("2,1")!.formulaSharedIndex).toBe(0);

    // Group 1
    expect(readCells.get("1,2")!.formulaSharedIndex).toBe(1);
    expect(readCells.get("2,2")!.formulaSharedIndex).toBe(1);
  });
});

// ── Round-trip: Array Formulas ───────────────────────────────────────

describe("array formula round-trip", () => {
  it("preserves array formula metadata through write → read", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("1,1", {
      formula: "SUM(A2:A5*C2:C5)",
      formulaType: "array",
      formulaRef: "B2:B2",
      formulaResult: 100,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["A", "B", "C"],
        [1, null, 10],
      ],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(xlsx);
    const readCells = workbook.sheets[0].cells!;

    const cell = readCells.get("1,1");
    expect(cell!.formulaType).toBe("array");
    expect(cell!.formulaRef).toBe("B2:B2");
    expect(cell!.formula).toBe("SUM(A2:A5*C2:C5)");
    expect(cell!.formulaDynamic).toBeUndefined();
  });

  it("preserves dynamic array formula through write → read", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,1", {
      formula: "UNIQUE(A1:A5)",
      formulaType: "array",
      formulaRef: "B1:B5",
      formulaDynamic: true,
      formulaResult: "alpha",
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["alpha"]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(xlsx);
    const readCells = workbook.sheets[0].cells!;

    const cell = readCells.get("0,1");
    expect(cell!.formulaType).toBe("array");
    expect(cell!.formulaRef).toBe("B1:B5");
    expect(cell!.formulaDynamic).toBe(true);
    expect(cell!.formula).toBe("UNIQUE(A1:A5)");
  });
});

// ── Normal formula round-trip (backward compat) ──────────────────────

describe("normal formula round-trip", () => {
  it("normal formula has no formulaType after round-trip", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,1", {
      formula: "A1+1",
      formulaResult: 6,
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [[5]],
      cells,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(xlsx);
    const readCells = workbook.sheets[0].cells!;

    const cell = readCells.get("0,1");
    expect(cell).toBeDefined();
    expect(cell!.formula).toBe("A1+1");
    expect(cell!.formulaType).toBeUndefined();
    expect(cell!.formulaSharedIndex).toBeUndefined();
    expect(cell!.formulaRef).toBeUndefined();
    expect(cell!.formulaDynamic).toBeUndefined();
  });
});

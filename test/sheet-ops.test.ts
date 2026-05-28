import { describe, it, expect } from "vitest"
import type { Sheet, Cell } from "../src/_types"
import {
  insertRows,
  deleteRows,
  insertColumns,
  deleteColumns,
  moveRows,
  hideRows,
  hideColumns,
  groupRows,
} from "../src/sheet-ops"

// ── Helpers ──────────────────────────────────────────────────────────

function makeSheet(overrides: Partial<Sheet> = {}): Sheet {
  return {
    name: "Sheet1",
    rows: [],
    ...overrides,
  }
}

function makeCell(value: string): Cell {
  return { value, type: "string" }
}

// ── insertRows ───────────────────────────────────────────────────────

describe("insertRows", () => {
  it("should insert 2 rows at position 0, shifting existing rows down", () => {
    const sheet = makeSheet({
      rows: [
        ["A1", "B1"],
        ["A2", "B2"],
        ["A3", "B3"],
      ],
    })

    insertRows(sheet, 0, 2)

    expect(sheet.rows).toEqual([
      [null, null],
      [null, null],
      ["A1", "B1"],
      ["A2", "B2"],
      ["A3", "B3"],
    ])
  })

  it("should insert rows in the middle, preserving data correctly", () => {
    const sheet = makeSheet({
      rows: [
        ["A1", "B1"],
        ["A2", "B2"],
        ["A3", "B3"],
      ],
    })

    insertRows(sheet, 1, 2)

    expect(sheet.rows).toEqual([
      ["A1", "B1"],
      [null, null],
      [null, null],
      ["A2", "B2"],
      ["A3", "B3"],
    ])
  })

  it("should insert rows at the end, extending the sheet", () => {
    const sheet = makeSheet({
      rows: [
        ["A1", "B1"],
        ["A2", "B2"],
      ],
    })

    insertRows(sheet, 2, 3)

    expect(sheet.rows).toEqual([
      ["A1", "B1"],
      ["A2", "B2"],
      [null, null],
      [null, null],
      [null, null],
    ])
  })

  it("should shift merge ranges correctly after insert", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
      merges: [{ startRow: 1, startCol: 0, endRow: 2, endCol: 1 }],
    })

    insertRows(sheet, 1, 2)

    expect(sheet.merges).toEqual([{ startRow: 3, startCol: 0, endRow: 4, endCol: 1 }])
  })

  it("should expand merge that spans the insertion point", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"], ["D"]],
      merges: [{ startRow: 0, startCol: 0, endRow: 2, endCol: 0 }],
    })

    insertRows(sheet, 1, 2)

    // Merge starts before insertion (row 0), ends at or after (row 2+2=4)
    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 4, endCol: 0 }])
  })

  it("should update cells Map keys correctly", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", makeCell("A1"))
    cells.set("1,0", makeCell("A2"))
    cells.set("2,1", makeCell("B3"))

    const sheet = makeSheet({
      rows: [["A1"], ["A2"], [null, "B3"]],
      cells,
    })

    insertRows(sheet, 1, 2)

    expect(sheet.cells!.get("0,0")!.value).toBe("A1")
    expect(sheet.cells!.has("1,0")).toBe(false)
    expect(sheet.cells!.get("3,0")!.value).toBe("A2")
    expect(sheet.cells!.get("4,1")!.value).toBe("B3")
  })

  it("should work on an empty sheet", () => {
    const sheet = makeSheet({ rows: [] })

    insertRows(sheet, 0, 3)

    expect(sheet.rows).toEqual([[], [], []])
  })

  it("should do nothing when count is 0", () => {
    const sheet = makeSheet({ rows: [["A"]] })
    insertRows(sheet, 0, 0)
    expect(sheet.rows).toEqual([["A"]])
  })

  it("should update data validations", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
      dataValidations: [{ type: "list", range: "A2:A3", values: ["x", "y"] }],
    })

    insertRows(sheet, 0, 1)

    expect(sheet.dataValidations![0].range).toBe("A3:A4")
  })

  it("should update conditional rules", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
      conditionalRules: [{ type: "cellIs", priority: 1, range: "A2:B3" }],
    })

    insertRows(sheet, 1, 2)

    expect(sheet.conditionalRules![0].range).toBe("A4:B5")
  })

  it("should update auto filter range", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
      autoFilter: { range: "A1:B3" },
    })

    insertRows(sheet, 0, 1)

    expect(sheet.autoFilter!.range).toBe("A2:B4")
  })

  it("should update image anchors", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"]],
      images: [
        {
          data: new Uint8Array(),
          type: "png",
          anchor: { from: { row: 1, col: 0 }, to: { row: 2, col: 1 } },
        },
      ],
    })

    insertRows(sheet, 0, 2)

    expect(sheet.images![0].anchor.from.row).toBe(3)
    expect(sheet.images![0].anchor.to!.row).toBe(4)
  })

  it("should update row defs", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
      rowDefs: new Map([
        [0, { hidden: true }],
        [2, { outlineLevel: 1 }],
      ]),
    })

    insertRows(sheet, 1, 2)

    expect(sheet.rowDefs!.get(0)).toEqual({ hidden: true })
    expect(sheet.rowDefs!.has(2)).toBe(false)
    expect(sheet.rowDefs!.get(4)).toEqual({ outlineLevel: 1 })
  })

  it("should update table ranges", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
      tables: [{ name: "Table1", columns: [{ name: "Col1" }], range: "A2:A3" }],
    })

    insertRows(sheet, 1, 1)

    expect(sheet.tables![0].range).toBe("A3:A4")
  })
})

// ── deleteRows ───────────────────────────────────────────────────────

describe("deleteRows", () => {
  it("should delete first row and shift remaining up", () => {
    const sheet = makeSheet({
      rows: [
        ["A1", "B1"],
        ["A2", "B2"],
        ["A3", "B3"],
      ],
    })

    deleteRows(sheet, 0, 1)

    expect(sheet.rows).toEqual([
      ["A2", "B2"],
      ["A3", "B3"],
    ])
  })

  it("should delete middle rows correctly", () => {
    const sheet = makeSheet({
      rows: [["A1"], ["A2"], ["A3"], ["A4"], ["A5"]],
    })

    deleteRows(sheet, 1, 2)

    expect(sheet.rows).toEqual([["A1"], ["A4"], ["A5"]])
  })

  it("should delete last row", () => {
    const sheet = makeSheet({
      rows: [["A1"], ["A2"], ["A3"]],
    })

    deleteRows(sheet, 2, 1)

    expect(sheet.rows).toEqual([["A1"], ["A2"]])
  })

  it("should remove merge within deleted range", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"], ["D"]],
      merges: [{ startRow: 1, startCol: 0, endRow: 2, endCol: 1 }],
    })

    deleteRows(sheet, 1, 2)

    expect(sheet.merges).toEqual([])
  })

  it("should adjust merge that partially overlaps from below", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"], ["D"], ["E"]],
      merges: [{ startRow: 0, startCol: 0, endRow: 3, endCol: 0 }],
    })

    // Delete rows 1 and 2
    deleteRows(sheet, 1, 2)

    // Merge was rows 0-3, deleted rows 1-2, so new merge is rows 0-1
    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 1, endCol: 0 }])
  })

  it("should shift merge entirely below deleted range", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"], ["D"], ["E"]],
      merges: [{ startRow: 3, startCol: 0, endRow: 4, endCol: 1 }],
    })

    deleteRows(sheet, 0, 2)

    expect(sheet.merges).toEqual([{ startRow: 1, startCol: 0, endRow: 2, endCol: 1 }])
  })

  it("should delete all rows leaving empty sheet", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
    })

    deleteRows(sheet, 0, 3)

    expect(sheet.rows).toEqual([])
  })

  it("should update cells Map correctly", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", makeCell("A1"))
    cells.set("1,0", makeCell("A2"))
    cells.set("2,0", makeCell("A3"))
    cells.set("3,0", makeCell("A4"))

    const sheet = makeSheet({
      rows: [["A1"], ["A2"], ["A3"], ["A4"]],
      cells,
    })

    // Delete row 1
    deleteRows(sheet, 1, 1)

    expect(sheet.cells!.get("0,0")!.value).toBe("A1")
    expect(sheet.cells!.has("1,0")).toBe(true)
    expect(sheet.cells!.get("1,0")!.value).toBe("A3")
    expect(sheet.cells!.get("2,0")!.value).toBe("A4")
  })

  it("should do nothing when count is 0", () => {
    const sheet = makeSheet({ rows: [["A"]] })
    deleteRows(sheet, 0, 0)
    expect(sheet.rows).toEqual([["A"]])
  })

  it("should update data validations and remove those fully in deleted range", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"], ["D"]],
      dataValidations: [
        { type: "list", range: "A2:A3", values: ["x"] },
        { type: "list", range: "A4:A4", values: ["y"] },
      ],
    })

    // Delete rows 1-2 (0-based), which are A2:A3 in 1-based
    deleteRows(sheet, 1, 2)

    // First validation fully within deleted range -> removed
    // Second validation at row 3 -> shifts to row 1
    expect(sheet.dataValidations!.length).toBe(1)
    expect(sheet.dataValidations![0].range).toBe("A2:A2")
  })

  it("should update conditional rules", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"], ["D"]],
      conditionalRules: [{ type: "cellIs", priority: 1, range: "A3:A4" }],
    })

    deleteRows(sheet, 0, 1)

    expect(sheet.conditionalRules![0].range).toBe("A2:A3")
  })

  it("should remove auto filter if fully within deleted range", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
      autoFilter: { range: "A2:B3" },
    })

    deleteRows(sheet, 1, 2)

    expect(sheet.autoFilter).toBeUndefined()
  })

  it("should update row defs", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"], ["D"]],
      rowDefs: new Map([
        [0, { hidden: true }],
        [1, { outlineLevel: 1 }],
        [3, { hidden: true }],
      ]),
    })

    deleteRows(sheet, 1, 2)

    expect(sheet.rowDefs!.get(0)).toEqual({ hidden: true })
    expect(sheet.rowDefs!.has(1)).toBe(true)
    expect(sheet.rowDefs!.get(1)).toEqual({ hidden: true })
    expect(sheet.rowDefs!.size).toBe(2)
  })
})

// ── insertColumns ────────────────────────────────────────────────────

describe("insertColumns", () => {
  it("should insert column at start, prepending null to all rows", () => {
    const sheet = makeSheet({
      rows: [
        ["A1", "B1"],
        ["A2", "B2"],
      ],
    })

    insertColumns(sheet, 0, 1)

    expect(sheet.rows).toEqual([
      [null, "A1", "B1"],
      [null, "A2", "B2"],
    ])
  })

  it("should insert columns in the middle", () => {
    const sheet = makeSheet({
      rows: [
        ["A1", "B1", "C1"],
        ["A2", "B2", "C2"],
      ],
    })

    insertColumns(sheet, 1, 2)

    expect(sheet.rows).toEqual([
      ["A1", null, null, "B1", "C1"],
      ["A2", null, null, "B2", "C2"],
    ])
  })

  it("should update column defs array", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C"]],
      columns: [
        { header: "Col1", width: 10 },
        { header: "Col2", width: 20 },
        { header: "Col3", width: 15 },
      ],
    })

    insertColumns(sheet, 1, 1)

    expect(sheet.columns!.length).toBe(4)
    expect(sheet.columns![0].header).toBe("Col1")
    expect(sheet.columns![1]).toEqual({})
    expect(sheet.columns![2].header).toBe("Col2")
    expect(sheet.columns![3].header).toBe("Col3")
  })

  it("should update cells Map keys", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", makeCell("A1"))
    cells.set("0,1", makeCell("B1"))
    cells.set("1,0", makeCell("A2"))

    const sheet = makeSheet({
      rows: [["A1", "B1"], ["A2"]],
      cells,
    })

    insertColumns(sheet, 1, 1)

    expect(sheet.cells!.get("0,0")!.value).toBe("A1")
    expect(sheet.cells!.get("0,2")!.value).toBe("B1")
    expect(sheet.cells!.get("1,0")!.value).toBe("A2")
  })

  it("should update merge ranges", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C"]],
      merges: [{ startRow: 0, startCol: 1, endRow: 2, endCol: 2 }],
    })

    insertColumns(sheet, 0, 1)

    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 2, endRow: 2, endCol: 3 }])
  })

  it("should expand merge spanning insertion point", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C"]],
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
    })

    insertColumns(sheet, 1, 2)

    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 4 }])
  })

  it("should update data validations", () => {
    const sheet = makeSheet({
      rows: [["A", "B"]],
      dataValidations: [{ type: "list", range: "B1:B5", values: ["x"] }],
    })

    insertColumns(sheet, 0, 1)

    expect(sheet.dataValidations![0].range).toBe("C1:C5")
  })

  it("should update auto filter range", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C"]],
      autoFilter: { range: "A1:C1" },
    })

    insertColumns(sheet, 1, 1)

    expect(sheet.autoFilter!.range).toBe("A1:D1")
  })

  it("should work on an empty sheet", () => {
    const sheet = makeSheet({ rows: [] })
    insertColumns(sheet, 0, 2)
    expect(sheet.rows).toEqual([])
  })

  it("should update image anchors", () => {
    const sheet = makeSheet({
      rows: [["A", "B"]],
      images: [
        {
          data: new Uint8Array(),
          type: "png",
          anchor: { from: { row: 0, col: 1 }, to: { row: 1, col: 2 } },
        },
      ],
    })

    insertColumns(sheet, 0, 1)

    expect(sheet.images![0].anchor.from.col).toBe(2)
    expect(sheet.images![0].anchor.to!.col).toBe(3)
  })
})

// ── deleteColumns ────────────────────────────────────────────────────

describe("deleteColumns", () => {
  it("should delete first column", () => {
    const sheet = makeSheet({
      rows: [
        ["A1", "B1", "C1"],
        ["A2", "B2", "C2"],
      ],
    })

    deleteColumns(sheet, 0, 1)

    expect(sheet.rows).toEqual([
      ["B1", "C1"],
      ["B2", "C2"],
    ])
  })

  it("should delete middle column", () => {
    const sheet = makeSheet({
      rows: [
        ["A1", "B1", "C1"],
        ["A2", "B2", "C2"],
      ],
    })

    deleteColumns(sheet, 1, 1)

    expect(sheet.rows).toEqual([
      ["A1", "C1"],
      ["A2", "C2"],
    ])
  })

  it("should update column defs after deletion", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C"]],
      columns: [
        { header: "Col1", width: 10 },
        { header: "Col2", width: 20 },
        { header: "Col3", width: 15 },
      ],
    })

    deleteColumns(sheet, 1, 1)

    expect(sheet.columns!.length).toBe(2)
    expect(sheet.columns![0].header).toBe("Col1")
    expect(sheet.columns![1].header).toBe("Col3")
  })

  it("should update cells Map correctly", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", makeCell("A1"))
    cells.set("0,1", makeCell("B1"))
    cells.set("0,2", makeCell("C1"))

    const sheet = makeSheet({
      rows: [["A1", "B1", "C1"]],
      cells,
    })

    deleteColumns(sheet, 1, 1)

    expect(sheet.cells!.get("0,0")!.value).toBe("A1")
    expect(sheet.cells!.has("0,1")).toBe(true)
    expect(sheet.cells!.get("0,1")!.value).toBe("C1")
  })

  it("should remove merge fully within deleted columns", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C", "D"]],
      merges: [{ startRow: 0, startCol: 1, endRow: 2, endCol: 2 }],
    })

    deleteColumns(sheet, 1, 2)

    expect(sheet.merges).toEqual([])
  })

  it("should shift merge below deleted columns", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C", "D", "E"]],
      merges: [{ startRow: 0, startCol: 3, endRow: 1, endCol: 4 }],
    })

    deleteColumns(sheet, 0, 2)

    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 1, endRow: 1, endCol: 2 }])
  })

  it("should update data validations and remove those fully in deleted columns", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C", "D"]],
      dataValidations: [
        { type: "list", range: "B1:B5", values: ["x"] },
        { type: "list", range: "D1:D5", values: ["y"] },
      ],
    })

    // Delete column B (index 1)
    deleteColumns(sheet, 1, 1)

    // First validation was B1:B5 (col 1), fully within deleted -> removed
    // Second validation was D1:D5 (col 3), shifts to col 2 -> C1:C5
    expect(sheet.dataValidations!.length).toBe(1)
    expect(sheet.dataValidations![0].range).toBe("C1:C5")
  })

  it("should remove auto filter if fully within deleted range", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C"]],
      autoFilter: { range: "B1:B10" },
    })

    deleteColumns(sheet, 1, 1)

    expect(sheet.autoFilter).toBeUndefined()
  })

  it("should handle deleting last columns", () => {
    const sheet = makeSheet({
      rows: [
        ["A", "B", "C"],
        ["D", "E", "F"],
      ],
    })

    deleteColumns(sheet, 2, 1)

    expect(sheet.rows).toEqual([
      ["A", "B"],
      ["D", "E"],
    ])
  })

  it("should handle deleting multiple columns", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C", "D", "E"]],
    })

    deleteColumns(sheet, 1, 3)

    expect(sheet.rows).toEqual([["A", "E"]])
  })
})

// ── moveRows ─────────────────────────────────────────────────────────

describe("moveRows", () => {
  it("should move row 0 to position 2", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
    })

    moveRows(sheet, 0, 1, 2)

    // Row "A" extracted from position 0, remaining is ["B","C"]
    // Insert at adjusted position 2-0=2 -> but after removal adjustedTo = 2-1 = 1
    // Actually: toIndex > fromIndex, so adjustedTo = 2 - 1 = 1
    // splice(1, 0, ["A"]) into ["B","C"] => ["B","A","C"]
    expect(sheet.rows).toEqual([["B"], ["A"], ["C"]])
  })

  it("should move row 2 to position 0", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
    })

    moveRows(sheet, 2, 1, 0)

    expect(sheet.rows).toEqual([["C"], ["A"], ["B"]])
  })

  it("should move multiple rows forward", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"], ["D"], ["E"]],
    })

    // Move rows 0-1 to position 3
    moveRows(sheet, 0, 2, 3)

    // Extract ["A","B"], remaining: ["C","D","E"]
    // adjustedTo = 3 - 2 = 1
    // Insert at 1: ["C", "A", "B", "D", "E"]
    expect(sheet.rows).toEqual([["C"], ["A"], ["B"], ["D"], ["E"]])
  })

  it("should move multiple rows backward", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"], ["D"], ["E"]],
    })

    // Move rows 3-4 to position 1
    moveRows(sheet, 3, 2, 1)

    // Extract ["D","E"], remaining: ["A","B","C"]
    // toIndex < fromIndex, so adjustedTo = 1
    // Insert at 1: ["A","D","E","B","C"]
    expect(sheet.rows).toEqual([["A"], ["D"], ["E"], ["B"], ["C"]])
  })

  it("should do nothing when fromIndex equals toIndex", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
    })

    moveRows(sheet, 1, 1, 1)

    expect(sheet.rows).toEqual([["A"], ["B"], ["C"]])
  })

  it("should update cells Map when moving rows", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", makeCell("A1"))
    cells.set("1,0", makeCell("B1"))
    cells.set("2,0", makeCell("C1"))

    const sheet = makeSheet({
      rows: [["A1"], ["B1"], ["C1"]],
      cells,
    })

    // Move row 0 to position 2
    moveRows(sheet, 0, 1, 2)

    // Result: ["B1"], ["A1"], ["C1"]
    expect(sheet.cells!.get("0,0")!.value).toBe("B1")
    expect(sheet.cells!.get("1,0")!.value).toBe("A1")
    expect(sheet.cells!.get("2,0")!.value).toBe("C1")
  })
})

// ── hideRows ─────────────────────────────────────────────────────────

describe("hideRows", () => {
  it("should set hidden state on rows", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
    })

    hideRows(sheet, 1, 2)

    expect(sheet.rowDefs!.get(1)!.hidden).toBe(true)
    expect(sheet.rowDefs!.get(2)!.hidden).toBe(true)
    expect(sheet.rowDefs!.has(0)).toBe(false)
  })

  it("should unhide rows when hidden=false", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"]],
      rowDefs: new Map([
        [0, { hidden: true }],
        [1, { hidden: true }],
      ]),
    })

    hideRows(sheet, 0, 1, false)

    expect(sheet.rowDefs!.get(0)!.hidden).toBe(false)
    expect(sheet.rowDefs!.get(1)!.hidden).toBe(true)
  })

  it("should preserve existing row def properties", () => {
    const sheet = makeSheet({
      rows: [["A"]],
      rowDefs: new Map([[0, { outlineLevel: 2 }]]),
    })

    hideRows(sheet, 0, 1)

    expect(sheet.rowDefs!.get(0)!.hidden).toBe(true)
    expect(sheet.rowDefs!.get(0)!.outlineLevel).toBe(2)
  })
})

// ── hideColumns ──────────────────────────────────────────────────────

describe("hideColumns", () => {
  it("should set hidden state on columns", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C"]],
      columns: [{ header: "Col1" }, { header: "Col2" }, { header: "Col3" }],
    })

    hideColumns(sheet, 1, 1)

    expect(sheet.columns![0].hidden).toBeUndefined()
    expect(sheet.columns![1].hidden).toBe(true)
    expect(sheet.columns![2].hidden).toBeUndefined()
  })

  it("should unhide columns when hidden=false", () => {
    const sheet = makeSheet({
      rows: [["A", "B"]],
      columns: [
        { header: "Col1", hidden: true },
        { header: "Col2", hidden: true },
      ],
    })

    hideColumns(sheet, 0, 1, false)

    expect(sheet.columns![0].hidden).toBe(false)
    expect(sheet.columns![1].hidden).toBe(true)
  })

  it("should create columns array if not present", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C"]],
    })

    hideColumns(sheet, 1, 1)

    expect(sheet.columns!.length).toBe(2)
    expect(sheet.columns![1].hidden).toBe(true)
  })

  it("should extend columns array if needed", () => {
    const sheet = makeSheet({
      rows: [["A"]],
      columns: [{ header: "Col1" }],
    })

    hideColumns(sheet, 3, 1)

    expect(sheet.columns!.length).toBe(4)
    expect(sheet.columns![3].hidden).toBe(true)
  })
})

// ── groupRows ────────────────────────────────────────────────────────

describe("groupRows", () => {
  it("should set outline level on rows", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"], ["D"]],
    })

    groupRows(sheet, 1, 2)

    expect(sheet.rowDefs!.get(1)!.outlineLevel).toBe(1)
    expect(sheet.rowDefs!.get(2)!.outlineLevel).toBe(1)
    expect(sheet.rowDefs!.has(0)).toBe(false)
    expect(sheet.rowDefs!.has(3)).toBe(false)
  })

  it("should set custom outline level", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
    })

    groupRows(sheet, 0, 2, 3)

    expect(sheet.rowDefs!.get(0)!.outlineLevel).toBe(3)
    expect(sheet.rowDefs!.get(1)!.outlineLevel).toBe(3)
    expect(sheet.rowDefs!.get(2)!.outlineLevel).toBe(3)
  })

  it("should ungroup by setting level to 0", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"]],
      rowDefs: new Map([
        [0, { outlineLevel: 2 }],
        [1, { outlineLevel: 2 }],
      ]),
    })

    groupRows(sheet, 0, 1, 0)

    expect(sheet.rowDefs!.get(0)!.outlineLevel).toBe(0)
    expect(sheet.rowDefs!.get(1)!.outlineLevel).toBe(0)
  })

  it("should preserve existing row def properties", () => {
    const sheet = makeSheet({
      rows: [["A"]],
      rowDefs: new Map([[0, { hidden: true }]]),
    })

    groupRows(sheet, 0, 0, 2)

    expect(sheet.rowDefs!.get(0)!.outlineLevel).toBe(2)
    expect(sheet.rowDefs!.get(0)!.hidden).toBe(true)
  })
})

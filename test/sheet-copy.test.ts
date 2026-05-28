import { describe, it, expect } from "vitest"
import type { Sheet, Workbook, Cell } from "../src/_types"
import {
  cloneSheet,
  copySheetToWorkbook,
  copyRange,
  moveSheet,
  removeSheet,
} from "../src/sheet-ops"

// ── Helpers ──────────────────────────────────────────────────────────

function makeSheet(overrides: Partial<Sheet> = {}): Sheet {
  return {
    name: "Sheet1",
    rows: [],
    ...overrides,
  }
}

function makeWorkbook(overrides: Partial<Workbook> = {}): Workbook {
  return {
    sheets: [],
    ...overrides,
  }
}

function makeCell(value: string, style?: Cell["style"]): Cell {
  return { value, type: "string", style }
}

// ── cloneSheet ──────────────────────────────────────────────────────

describe("cloneSheet", () => {
  it("should clone and preserve all row data", () => {
    const sheet = makeSheet({
      rows: [
        ["A1", "B1"],
        ["A2", "B2"],
        [1, 2, 3],
      ],
    })

    const cloned = cloneSheet(sheet, "Cloned")

    expect(cloned.rows).toEqual([
      ["A1", "B1"],
      ["A2", "B2"],
      [1, 2, 3],
    ])
  })

  it("should clone and preserve cell styles", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", {
      value: "styled",
      type: "string",
      style: {
        font: { bold: true, size: 14, color: { rgb: "FF0000" } },
        fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "00FF00" } },
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "medium" },
        },
        alignment: { horizontal: "center", wrapText: true },
        numFmt: "#,##0.00",
        protection: { locked: true },
      },
    })

    const sheet = makeSheet({ rows: [["styled"]], cells })
    const cloned = cloneSheet(sheet, "Cloned")

    const original = sheet.cells!.get("0,0")!
    const clonedCell = cloned.cells!.get("0,0")!

    expect(clonedCell.style).toEqual(original.style)
  })

  it("should clone and preserve merges", () => {
    const sheet = makeSheet({
      rows: [
        ["A", "B"],
        ["C", "D"],
      ],
      merges: [
        { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
        { startRow: 2, startCol: 0, endRow: 3, endCol: 2 },
      ],
    })

    const cloned = cloneSheet(sheet, "Cloned")

    expect(cloned.merges).toEqual([
      { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      { startRow: 2, startCol: 0, endRow: 3, endCol: 2 },
    ])
  })

  it("should clone and preserve data validations", () => {
    const sheet = makeSheet({
      rows: [["A"]],
      dataValidations: [
        { type: "list", range: "A1:A10", values: ["x", "y", "z"], allowBlank: true },
        { type: "whole", range: "B1:B10", operator: "between", formula1: "1", formula2: "100" },
      ],
    })

    const cloned = cloneSheet(sheet, "Cloned")

    expect(cloned.dataValidations).toEqual(sheet.dataValidations)
  })

  it("should clone and preserve conditional rules", () => {
    const sheet = makeSheet({
      rows: [["A"]],
      conditionalRules: [
        {
          type: "cellIs",
          priority: 1,
          range: "A1:A10",
          operator: "greaterThan",
          formula: "100",
          style: { font: { bold: true, color: { rgb: "FF0000" } } },
        },
        {
          type: "colorScale",
          priority: 2,
          range: "B1:B10",
          colorScale: {
            cfvo: [{ type: "min" }, { type: "max" }],
            colors: ["FF63BE7B", "FFF8696B"],
          },
        },
      ],
    })

    const cloned = cloneSheet(sheet, "Cloned")

    expect(cloned.conditionalRules!.length).toBe(2)
    expect(cloned.conditionalRules![0].style).toEqual(sheet.conditionalRules![0].style)
    expect(cloned.conditionalRules![1].colorScale).toEqual(sheet.conditionalRules![1].colorScale)
  })

  it("should set the new name on the cloned sheet", () => {
    const sheet = makeSheet({ name: "Original" })
    const cloned = cloneSheet(sheet, "NewName")

    expect(cloned.name).toBe("NewName")
    expect(sheet.name).toBe("Original")
  })

  it("should be independent — modifying clone does not affect original", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", makeCell("original", { font: { bold: true } }))

    const sheet = makeSheet({
      rows: [
        ["A", "B"],
        ["C", "D"],
      ],
      cells,
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
      dataValidations: [{ type: "list", range: "A1:A5", values: ["a", "b"] }],
    })

    const cloned = cloneSheet(sheet, "Cloned")

    // Modify cloned data
    cloned.rows[0][0] = "MODIFIED"
    cloned.rows.push(["new row"])
    cloned.cells!.get("0,0")!.value = "CHANGED"
    cloned.cells!.get("0,0")!.style!.font!.bold = false
    cloned.merges![0].startRow = 99
    cloned.dataValidations![0].values!.push("c")

    // Original should be unchanged
    expect(sheet.rows[0][0]).toBe("A")
    expect(sheet.rows.length).toBe(2)
    expect(sheet.cells!.get("0,0")!.value).toBe("original")
    expect(sheet.cells!.get("0,0")!.style!.font!.bold).toBe(true)
    expect(sheet.merges![0].startRow).toBe(0)
    expect(sheet.dataValidations![0].values).toEqual(["a", "b"])
  })

  it("should clone and preserve comments", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", {
      value: "with comment",
      type: "string",
      comment: {
        author: "John",
        text: "This is a comment",
        richText: [{ text: "Bold part", font: { bold: true } }],
      },
    })

    const sheet = makeSheet({ rows: [["with comment"]], cells })
    const cloned = cloneSheet(sheet, "Cloned")

    const clonedCell = cloned.cells!.get("0,0")!
    expect(clonedCell.comment!.author).toBe("John")
    expect(clonedCell.comment!.text).toBe("This is a comment")
    expect(clonedCell.comment!.richText![0].text).toBe("Bold part")
    expect(clonedCell.comment!.richText![0].font!.bold).toBe(true)

    // Verify independence
    clonedCell.comment!.text = "Modified"
    expect(sheet.cells!.get("0,0")!.comment!.text).toBe("This is a comment")
  })

  it("should clone and preserve hyperlinks", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", {
      value: "click me",
      type: "string",
      hyperlink: { target: "https://example.com", tooltip: "Visit example" },
    })

    const sheet = makeSheet({ rows: [["click me"]], cells })
    const cloned = cloneSheet(sheet, "Cloned")

    const clonedCell = cloned.cells!.get("0,0")!
    expect(clonedCell.hyperlink!.target).toBe("https://example.com")
    expect(clonedCell.hyperlink!.tooltip).toBe("Visit example")

    // Verify independence
    clonedCell.hyperlink!.target = "https://other.com"
    expect(sheet.cells!.get("0,0")!.hyperlink!.target).toBe("https://example.com")
  })

  it("should clone an empty sheet", () => {
    const sheet = makeSheet({ name: "Empty", rows: [] })
    const cloned = cloneSheet(sheet, "ClonedEmpty")

    expect(cloned.name).toBe("ClonedEmpty")
    expect(cloned.rows).toEqual([])
    expect(cloned.cells).toBeUndefined()
    expect(cloned.merges).toBeUndefined()
    expect(cloned.dataValidations).toBeUndefined()
    expect(cloned.conditionalRules).toBeUndefined()
    expect(cloned.images).toBeUndefined()
  })

  it("should clone images with new Uint8Array data", () => {
    const imageData = new Uint8Array([1, 2, 3, 4, 5])
    const sheet = makeSheet({
      rows: [["img"]],
      images: [
        {
          data: imageData,
          type: "png",
          anchor: { from: { row: 0, col: 0 }, to: { row: 2, col: 2 } },
          width: 100,
          height: 200,
        },
      ],
    })

    const cloned = cloneSheet(sheet, "Cloned")

    expect(cloned.images!.length).toBe(1)
    expect(cloned.images![0].data).toEqual(new Uint8Array([1, 2, 3, 4, 5]))
    expect(cloned.images![0].type).toBe("png")
    expect(cloned.images![0].anchor.from).toEqual({ row: 0, col: 0 })
    expect(cloned.images![0].anchor.to).toEqual({ row: 2, col: 2 })
    expect(cloned.images![0].width).toBe(100)
    expect(cloned.images![0].height).toBe(200)

    // Verify independence — modifying cloned image data doesn't affect original
    cloned.images![0].data[0] = 99
    expect(sheet.images![0].data[0]).toBe(1)
  })

  it("should clone rowDefs", () => {
    const sheet = makeSheet({
      rows: [["A"], ["B"], ["C"]],
      rowDefs: new Map([
        [0, { height: 20, hidden: false }],
        [2, { outlineLevel: 1 }],
      ]),
    })

    const cloned = cloneSheet(sheet, "Cloned")

    expect(cloned.rowDefs!.get(0)).toEqual({ height: 20, hidden: false })
    expect(cloned.rowDefs!.get(2)).toEqual({ outlineLevel: 1 })

    // Verify independence
    cloned.rowDefs!.get(0)!.height = 50
    expect(sheet.rowDefs!.get(0)!.height).toBe(20)
  })

  it("should clone columns", () => {
    const sheet = makeSheet({
      rows: [["A", "B"]],
      columns: [
        { header: "Col1", width: 10, hidden: false },
        { header: "Col2", width: 20, autoWidth: true },
      ],
    })

    const cloned = cloneSheet(sheet, "Cloned")

    expect(cloned.columns!.length).toBe(2)
    expect(cloned.columns![0]).toEqual({
      header: "Col1",
      width: 10,
      hidden: false,
      style: undefined,
    })
    expect(cloned.columns![1]).toEqual({
      header: "Col2",
      width: 20,
      autoWidth: true,
      style: undefined,
    })
  })

  it("should clone autoFilter, freezePane, protection, pageSetup, headerFooter, view", () => {
    const sheet = makeSheet({
      rows: [["A"]],
      autoFilter: { range: "A1:D10" },
      freezePane: { rows: 1, columns: 2 },
      protection: { sheet: true, password: "test" },
      pageSetup: { paperSize: "a4", orientation: "landscape", margins: { top: 1, bottom: 1 } },
      headerFooter: { oddHeader: "&CHeader" },
      view: { showGridLines: false, zoomScale: 150, tabColor: { rgb: "FF0000" } },
    })

    const cloned = cloneSheet(sheet, "Cloned")

    expect(cloned.autoFilter).toEqual({ range: "A1:D10" })
    expect(cloned.freezePane).toEqual({ rows: 1, columns: 2 })
    expect(cloned.protection).toEqual({ sheet: true, password: "test" })
    expect(cloned.pageSetup).toEqual({
      paperSize: "a4",
      orientation: "landscape",
      margins: { top: 1, bottom: 1 },
    })
    expect(cloned.headerFooter).toEqual({ oddHeader: "&CHeader" })
    expect(cloned.view).toEqual({
      showGridLines: false,
      zoomScale: 150,
      tabColor: { rgb: "FF0000" },
    })

    // Verify independence
    cloned.autoFilter!.range = "Z1:Z99"
    cloned.freezePane!.rows = 5
    cloned.pageSetup!.margins!.top = 99
    cloned.view!.tabColor!.rgb = "0000FF"

    expect(sheet.autoFilter!.range).toBe("A1:D10")
    expect(sheet.freezePane!.rows).toBe(1)
    expect(sheet.pageSetup!.margins!.top).toBe(1)
    expect(sheet.view!.tabColor!.rgb).toBe("FF0000")
  })

  it("should clone tables", () => {
    const sheet = makeSheet({
      rows: [
        ["Name", "Value"],
        ["A", 1],
      ],
      tables: [
        {
          name: "Table1",
          columns: [{ name: "Name" }, { name: "Value", totalFunction: "sum" }],
          range: "A1:B2",
          style: "TableStyleMedium2",
          showRowStripes: true,
        },
      ],
    })

    const cloned = cloneSheet(sheet, "Cloned")

    expect(cloned.tables!.length).toBe(1)
    expect(cloned.tables![0].name).toBe("Table1")
    expect(cloned.tables![0].columns).toEqual([
      { name: "Name" },
      { name: "Value", totalFunction: "sum" },
    ])

    // Verify independence
    cloned.tables![0].columns[0].name = "Modified"
    expect(sheet.tables![0].columns[0].name).toBe("Name")
  })
})

// ── copySheetToWorkbook ─────────────────────────────────────────────

describe("copySheetToWorkbook", () => {
  it("should copy sheet between workbooks", () => {
    const sourceSheet = makeSheet({
      name: "Source",
      rows: [
        ["A1", "B1"],
        ["A2", "B2"],
      ],
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
    })

    const targetWorkbook = makeWorkbook({ sheets: [makeSheet({ name: "Existing" })] })

    copySheetToWorkbook(sourceSheet, targetWorkbook, "Copied")

    expect(targetWorkbook.sheets.length).toBe(2)
    expect(targetWorkbook.sheets[1].name).toBe("Copied")
    expect(targetWorkbook.sheets[1].rows).toEqual([
      ["A1", "B1"],
      ["A2", "B2"],
    ])
    expect(targetWorkbook.sheets[1].merges).toEqual([
      { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
    ])
  })

  it("should not modify the original workbook's sheets", () => {
    const sourceSheet = makeSheet({
      name: "Source",
      rows: [["A1"]],
    })

    const sourceWorkbook = makeWorkbook({ sheets: [sourceSheet] })
    const targetWorkbook = makeWorkbook({ sheets: [] })

    copySheetToWorkbook(sourceSheet, targetWorkbook, "Copied")

    expect(sourceWorkbook.sheets.length).toBe(1)
    expect(sourceWorkbook.sheets[0].name).toBe("Source")
  })

  it("should use source sheet name when newName not provided", () => {
    const sourceSheet = makeSheet({ name: "OriginalName", rows: [["data"]] })
    const targetWorkbook = makeWorkbook({ sheets: [] })

    copySheetToWorkbook(sourceSheet, targetWorkbook)

    expect(targetWorkbook.sheets.length).toBe(1)
    expect(targetWorkbook.sheets[0].name).toBe("OriginalName")
  })

  it("should produce an independent copy in target workbook", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", makeCell("value"))

    const sourceSheet = makeSheet({
      name: "Source",
      rows: [["value"]],
      cells,
    })

    const targetWorkbook = makeWorkbook({ sheets: [] })

    copySheetToWorkbook(sourceSheet, targetWorkbook, "Copy")

    // Modify target
    targetWorkbook.sheets[0].rows[0][0] = "modified"
    targetWorkbook.sheets[0].cells!.get("0,0")!.value = "modified"

    // Source unchanged
    expect(sourceSheet.rows[0][0]).toBe("value")
    expect(sourceSheet.cells!.get("0,0")!.value).toBe("value")
  })
})

// ── copyRange ───────────────────────────────────────────────────────

describe("copyRange", () => {
  it("should copy a range of cells within the same sheet", () => {
    const sheet = makeSheet({
      rows: [
        ["A1", "B1", "C1"],
        ["A2", "B2", "C2"],
        ["A3", "B3", "C3"],
        [null, null, null],
        [null, null, null],
      ],
    })

    copyRange(
      sheet,
      { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      { startRow: 3, startCol: 0 },
    )

    expect(sheet.rows[3]).toEqual(["A1", "B1", null])
    expect(sheet.rows[4]).toEqual(["A2", "B2", null])
    // Source unchanged
    expect(sheet.rows[0]).toEqual(["A1", "B1", "C1"])
    expect(sheet.rows[1]).toEqual(["A2", "B2", "C2"])
  })

  it("should copy cell styles", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", {
      value: "styled",
      type: "string",
      style: { font: { bold: true, color: { rgb: "FF0000" } } },
    })
    cells.set("0,1", {
      value: "other",
      type: "string",
      style: { font: { italic: true } },
    })

    const sheet = makeSheet({
      rows: [
        ["styled", "other"],
        [null, null],
      ],
      cells,
    })

    copyRange(
      sheet,
      { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      { startRow: 1, startCol: 0 },
    )

    expect(sheet.cells!.get("1,0")!.value).toBe("styled")
    expect(sheet.cells!.get("1,0")!.style!.font!.bold).toBe(true)
    expect(sheet.cells!.get("1,1")!.value).toBe("other")
    expect(sheet.cells!.get("1,1")!.style!.font!.italic).toBe(true)

    // Verify independence
    sheet.cells!.get("1,0")!.style!.font!.bold = false
    expect(sheet.cells!.get("0,0")!.style!.font!.bold).toBe(true)
  })

  it("should handle overlapping source and target ranges", () => {
    const sheet = makeSheet({
      rows: [
        ["A", "B"],
        ["C", "D"],
        [null, null],
      ],
    })

    // Copy rows 0-1 to rows 1-2 (overlapping at row 1)
    copyRange(
      sheet,
      { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      { startRow: 1, startCol: 0 },
    )

    expect(sheet.rows[0]).toEqual(["A", "B"])
    expect(sheet.rows[1]).toEqual(["A", "B"])
    expect(sheet.rows[2]).toEqual(["C", "D"])
  })

  it("should copy merges within source range to target", () => {
    const sheet = makeSheet({
      rows: [
        ["A", "B", "C"],
        ["D", "E", "F"],
        [null, null, null],
        [null, null, null],
      ],
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
    })

    copyRange(
      sheet,
      { startRow: 0, startCol: 0, endRow: 1, endCol: 2 },
      { startRow: 2, startCol: 0 },
    )

    // Original merge still there
    expect(sheet.merges!).toContainEqual({ startRow: 0, startCol: 0, endRow: 0, endCol: 1 })
    // New merge at offset position
    expect(sheet.merges!).toContainEqual({ startRow: 2, startCol: 0, endRow: 2, endCol: 1 })
  })

  it("should extend rows array if target is beyond current rows", () => {
    const sheet = makeSheet({
      rows: [["A1", "B1"]],
    })

    copyRange(
      sheet,
      { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      { startRow: 3, startCol: 0 },
    )

    expect(sheet.rows.length).toBe(4)
    expect(sheet.rows[3]).toEqual(["A1", "B1"])
  })

  it("should extend row width if target column is beyond current width", () => {
    const sheet = makeSheet({
      rows: [["A1"], [null]],
    })

    copyRange(
      sheet,
      { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
      { startRow: 1, startCol: 3 },
    )

    expect(sheet.rows[1][3]).toBe("A1")
  })
})

// ── moveSheet ───────────────────────────────────────────────────────

describe("moveSheet", () => {
  it("should move first sheet to last position", () => {
    const workbook = makeWorkbook({
      sheets: [
        makeSheet({ name: "Sheet1" }),
        makeSheet({ name: "Sheet2" }),
        makeSheet({ name: "Sheet3" }),
      ],
    })

    moveSheet(workbook, 0, 2)

    expect(workbook.sheets.map((s) => s.name)).toEqual(["Sheet2", "Sheet3", "Sheet1"])
  })

  it("should move last sheet to first position", () => {
    const workbook = makeWorkbook({
      sheets: [
        makeSheet({ name: "Sheet1" }),
        makeSheet({ name: "Sheet2" }),
        makeSheet({ name: "Sheet3" }),
      ],
    })

    moveSheet(workbook, 2, 0)

    expect(workbook.sheets.map((s) => s.name)).toEqual(["Sheet3", "Sheet1", "Sheet2"])
  })

  it("should be no-op when moving to same position", () => {
    const workbook = makeWorkbook({
      sheets: [
        makeSheet({ name: "Sheet1" }),
        makeSheet({ name: "Sheet2" }),
        makeSheet({ name: "Sheet3" }),
      ],
    })

    moveSheet(workbook, 1, 1)

    expect(workbook.sheets.map((s) => s.name)).toEqual(["Sheet1", "Sheet2", "Sheet3"])
  })

  it("should move middle sheet to beginning", () => {
    const workbook = makeWorkbook({
      sheets: [
        makeSheet({ name: "A" }),
        makeSheet({ name: "B" }),
        makeSheet({ name: "C" }),
        makeSheet({ name: "D" }),
      ],
    })

    moveSheet(workbook, 2, 0)

    expect(workbook.sheets.map((s) => s.name)).toEqual(["C", "A", "B", "D"])
  })

  it("should move middle sheet to end", () => {
    const workbook = makeWorkbook({
      sheets: [
        makeSheet({ name: "A" }),
        makeSheet({ name: "B" }),
        makeSheet({ name: "C" }),
        makeSheet({ name: "D" }),
      ],
    })

    moveSheet(workbook, 1, 3)

    expect(workbook.sheets.map((s) => s.name)).toEqual(["A", "C", "D", "B"])
  })
})

// ── removeSheet ─────────────────────────────────────────────────────

describe("removeSheet", () => {
  it("should remove a middle sheet", () => {
    const workbook = makeWorkbook({
      sheets: [
        makeSheet({ name: "Sheet1" }),
        makeSheet({ name: "Sheet2" }),
        makeSheet({ name: "Sheet3" }),
      ],
    })

    removeSheet(workbook, 1)

    expect(workbook.sheets.length).toBe(2)
    expect(workbook.sheets.map((s) => s.name)).toEqual(["Sheet1", "Sheet3"])
  })

  it("should remove the last sheet", () => {
    const workbook = makeWorkbook({
      sheets: [
        makeSheet({ name: "Sheet1" }),
        makeSheet({ name: "Sheet2" }),
        makeSheet({ name: "Sheet3" }),
      ],
    })

    removeSheet(workbook, 2)

    expect(workbook.sheets.length).toBe(2)
    expect(workbook.sheets.map((s) => s.name)).toEqual(["Sheet1", "Sheet2"])
  })

  it("should adjust activeSheet when removing the active sheet", () => {
    const workbook = makeWorkbook({
      sheets: [
        makeSheet({ name: "Sheet1" }),
        makeSheet({ name: "Sheet2" }),
        makeSheet({ name: "Sheet3" }),
      ],
      activeSheet: 2,
    })

    removeSheet(workbook, 2)

    // activeSheet was 2, removed index 2, so it should clamp to 1
    expect(workbook.activeSheet).toBe(1)
  })

  it("should adjust activeSheet when removing a sheet before the active sheet", () => {
    const workbook = makeWorkbook({
      sheets: [
        makeSheet({ name: "Sheet1" }),
        makeSheet({ name: "Sheet2" }),
        makeSheet({ name: "Sheet3" }),
      ],
      activeSheet: 2,
    })

    removeSheet(workbook, 0)

    expect(workbook.activeSheet).toBe(1)
  })

  it("should not adjust activeSheet when removing a sheet after the active sheet", () => {
    const workbook = makeWorkbook({
      sheets: [
        makeSheet({ name: "Sheet1" }),
        makeSheet({ name: "Sheet2" }),
        makeSheet({ name: "Sheet3" }),
      ],
      activeSheet: 0,
    })

    removeSheet(workbook, 2)

    expect(workbook.activeSheet).toBe(0)
  })

  it("should set activeSheet to 0 when removing the only sheet", () => {
    const workbook = makeWorkbook({
      sheets: [makeSheet({ name: "Only" })],
      activeSheet: 0,
    })

    removeSheet(workbook, 0)

    expect(workbook.sheets.length).toBe(0)
    expect(workbook.activeSheet).toBe(0)
  })

  it("should handle removing the active middle sheet", () => {
    const workbook = makeWorkbook({
      sheets: [
        makeSheet({ name: "Sheet1" }),
        makeSheet({ name: "Sheet2" }),
        makeSheet({ name: "Sheet3" }),
      ],
      activeSheet: 1,
    })

    removeSheet(workbook, 1)

    // Removed index 1 which was active, min(1, length-1=1) = 1
    expect(workbook.activeSheet).toBe(1)
    expect(workbook.sheets[1].name).toBe("Sheet3")
  })

  it("should work when activeSheet is undefined", () => {
    const workbook = makeWorkbook({
      sheets: [makeSheet({ name: "Sheet1" }), makeSheet({ name: "Sheet2" })],
    })

    removeSheet(workbook, 0)

    expect(workbook.sheets.length).toBe(1)
    expect(workbook.activeSheet).toBeUndefined()
  })
})

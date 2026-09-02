import { describe, expect, it } from "vitest"
import { cloneSheet, copySheetToWorkbook } from "../src/sheet-ops"
import type { Cell, Sheet, Workbook } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #439 §N/§O — `cloneSheet` is documented as "Deep clone (all data,
// styles, merges, validations, etc.)" and enumerated the fields of
// `Sheet` by hand. It reached 19 of 30, so a clone came back with no
// sparklines, no text boxes, no page breaks, no split pane, no background
// image and no outline properties. `copySheetToWorkbook` is built on it,
// so copying a sheet between workbooks lost the same eleven.
//
// `cloneCell` had the same shape of gap: it carried `formula` but not
// `formulaType` / `formulaSharedIndex` / `formulaRef` / `formulaDynamic`
// or `checkbox`, which is worse than loss — a shared-formula slave cell
// became `{ formula: "" }`, and the writer emits that as an empty `<f/>`.
//
// `Required<…>` below is the guard: adding a field to `Sheet` or `Cell`
// fails `tsc` here until the fixture carries it, and the deep-equality
// assertion then fails until `cloneSheet` carries it too.
// ═══════════════════════════════════════════════════════════════════════

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47])

/** Every field of `Cell`, so the type stops us forgetting one. */
const FULL_CELL: Required<Cell> = {
  value: 42,
  type: "number",
  style: { font: { bold: true }, numFmt: "0.00" },
  checkbox: true,
  formula: "",
  formulaResult: 42,
  formulaType: "shared",
  formulaSharedIndex: 3,
  formulaRef: "A1:A9",
  formulaDynamic: true,
  richText: [{ text: "hi", font: { italic: true, color: { rgb: "FF0000" } } }],
  hyperlink: { target: "https://example.com", tooltip: "go" },
  comment: { text: "note", author: "Ada" },
}

/** Every field of `Sheet`, same reason. */
const FULL_SHEET: Required<Sheet> = {
  name: "Full",
  kind: "worksheet",
  rows: [
    ["a", 1],
    ["b", 2],
  ],
  cells: new Map<string, Cell>([["0,0", FULL_CELL]]),
  columns: [{ width: 12, style: { font: { name: "Arial" } } }],
  rowDefs: new Map([[0, { height: 30, hidden: true }]]),
  defaultRowHeight: 24,
  defaultColWidth: 18,
  merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
  dataValidations: [{ type: "list", range: "A1:A9", values: ["x", "y"] }],
  conditionalRules: [
    {
      type: "cellIs",
      priority: 1,
      range: "A1:A9",
      operator: "greaterThan",
      formula: ["0"],
      style: { font: { bold: true } },
    },
  ],
  autoFilter: { range: "A1:B2", columns: [{ colIndex: 0, filters: ["a"] }] },
  freezePane: { rows: 1, columns: 1 },
  splitPane: { xSplit: 2000, ySplit: 1000 },
  images: [{ data: PNG, type: "png", anchor: { from: { row: 0, col: 0 } }, altText: "logo" }],
  protection: { sheet: true, password: "pw", selectLockedCells: false },
  pageSetup: { orientation: "landscape", margins: { top: 1 }, printArea: "A1:B2" },
  headerFooter: { oddHeader: "&LLeft", oddFooter: "&P" },
  view: { showGridLines: false, zoomScale: 125, tabColor: { rgb: "00FF00" } },
  hidden: true,
  veryHidden: false,
  tables: [
    {
      name: "T1",
      range: "A1:B2",
      columns: [{ name: "a" }, { name: "b" }],
    },
  ],
  rowBreaks: [3, 7],
  colBreaks: [2],
  outlineProperties: { summaryBelow: false, summaryRight: false },
  backgroundImage: PNG,
  sparklines: [{ dataRange: "A1:A9", location: "B1", type: "line" }],
  textBoxes: [{ text: "hello", anchor: { from: { row: 1, col: 1 } } }],
  threadedComments: [
    { id: "{1}", ref: "A1", personId: "p1", text: "hi", date: "2024-01-15T00:00:00Z", done: false },
  ],
  a11y: { summary: "A full sheet", headerRow: 0 },
  pivotTables: [{ name: "P1", cacheId: 1, location: "D1:E5", fields: [] }],
  slicers: [{ name: "S1", cache: "c1", caption: "Region" }],
  timelines: [{ name: "T1", cache: "tc1", caption: "Date" }],
  charts: [{ kinds: ["bar"], seriesCount: 0, series: [], anchor: { from: { row: 0, col: 0 } } }],
}

describe("cloneSheet carries every field of Sheet", () => {
  it("produces an equal sheet", () => {
    const copy = cloneSheet(FULL_SHEET, "Copy")

    expect({ ...copy, name: FULL_SHEET.name }).toEqual(FULL_SHEET)
  })

  it("leaves nothing behind", () => {
    const copy = cloneSheet(FULL_SHEET, "Copy") as unknown as Record<string, unknown>

    const missing = Object.keys(FULL_SHEET).filter((k) => copy[k] === undefined)

    expect(missing).toEqual([])
  })

  it("detaches the collections it copied", () => {
    const copy = cloneSheet(FULL_SHEET, "Copy")

    expect(copy.rowBreaks).not.toBe(FULL_SHEET.rowBreaks)
    expect(copy.colBreaks).not.toBe(FULL_SHEET.colBreaks)
    expect(copy.splitPane).not.toBe(FULL_SHEET.splitPane)
    expect(copy.outlineProperties).not.toBe(FULL_SHEET.outlineProperties)
    expect(copy.backgroundImage).not.toBe(FULL_SHEET.backgroundImage)
    expect(copy.sparklines).not.toBe(FULL_SHEET.sparklines)
    expect(copy.sparklines![0]).not.toBe(FULL_SHEET.sparklines[0])
    expect(copy.textBoxes![0]).not.toBe(FULL_SHEET.textBoxes[0])
    expect(copy.pivotTables![0]).not.toBe(FULL_SHEET.pivotTables[0])
    expect(copy.slicers![0]).not.toBe(FULL_SHEET.slicers[0])
    expect(copy.timelines![0]).not.toBe(FULL_SHEET.timelines[0])
    expect(copy.threadedComments![0]).not.toBe(FULL_SHEET.threadedComments[0])

    copy.rowBreaks!.push(99)
    copy.sparklines![0]!.dataRange = "Z1:Z9"

    expect(FULL_SHEET.rowBreaks).toEqual([3, 7])
    expect(FULL_SHEET.sparklines[0]!.dataRange).toBe("A1:A9")
  })

  it("carries the whole cell, formula shape included", () => {
    const copy = cloneSheet(FULL_SHEET, "Copy")
    const cell = copy.cells!.get("0,0")!

    // A shared-formula slave is `{ formula: "", formulaType: "shared", si }`.
    // Losing the type and the index left `{ formula: "" }`, which the
    // writer emits as an empty <f/> — a reference replaced by nothing.
    expect(cell).toEqual(FULL_CELL)
    expect(cell.formulaType).toBe("shared")
    expect(cell.formulaSharedIndex).toBe(3)
    expect(cell.formulaRef).toBe("A1:A9")
    expect(cell.formulaDynamic).toBe(true)
    expect(cell.checkbox).toBe(true)
    expect(cell).not.toBe(FULL_CELL)
    expect(cell.style).not.toBe(FULL_CELL.style)
  })

  it("takes the new name and nothing else from the caller", () => {
    expect(cloneSheet(FULL_SHEET, "Renamed").name).toBe("Renamed")
  })
})

describe("copySheetToWorkbook carries the same fields", () => {
  it("brings the whole sheet across", () => {
    const target: Workbook = { sheets: [] }

    copySheetToWorkbook(FULL_SHEET, target, "Brought")

    const brought = target.sheets[0]!
    expect(brought.name).toBe("Brought")
    expect({ ...brought, name: FULL_SHEET.name }).toEqual(FULL_SHEET)
  })
})

describe("a Workbook is plain data and survives structuredClone", () => {
  // v1 shipped serializeWorkbook / deserializeWorkbook on the claim that
  // structured clone "does NOT handle Map". It does — Map, Date and
  // Uint8Array are all in the algorithm — so v2 removed them, and this
  // is the promise that replaces them: the model is plain data, and
  // postMessage carries it as-is. A class instance added to the model
  // would come back as a bare object here and fail.
  /** Every field of `Workbook`, so the type stops us forgetting one. */
  const FULL_WORKBOOK: Required<Workbook> = {
    sheets: [FULL_SHEET],
    properties: { title: "T", creator: "Ada", created: new Date(Date.UTC(2024, 0, 15)) },
    namedRanges: [{ name: "N", range: "Sheet1!$A$1" }],
    dateSystem: "1904",
    defaultFont: { name: "Calibri", size: 11 },
    activeSheet: 0,
    themeColors: ["FFFFFF", "000000"],
    workbookProtection: { lockStructure: true, lockWindows: false },
    persons: [{ id: "p1", displayName: "Ada", userId: "ada@example.com", providerId: "None" }],
    externalLinks: [{ target: "other.xlsx", sheetNames: ["S"], sheetData: [], definedNames: [] }],
    cellImages: [{ id: "ID_1", data: PNG, type: "png" }],
    pivotCaches: [{ cacheId: 1, sourceSheet: "Full", sourceRef: "A1:B2", fieldNames: ["a", "b"] }],
    slicerCaches: [{ name: "c1", sourceName: "Region" }],
    timelineCaches: [{ name: "tc1", sourceName: "Date" }],
  }

  it("round-trips every field of the full workbook", () => {
    expect(structuredClone(FULL_WORKBOOK)).toEqual(FULL_WORKBOOK)
  })
})

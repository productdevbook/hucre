import { describe, expect, it } from "vitest"
import type {
  Cell,
  CellStyle,
  CellValue,
  ConditionalRule,
  Sheet,
  SheetImage,
  Workbook,
} from "../src/_types"
import {
  cloneSheet,
  copyRange,
  deleteColumns,
  deleteRows,
  insertColumns,
  insertRows,
  moveRows,
  replaceCells,
  sortRows,
} from "../src/sheet-ops"

// ── Helpers ──────────────────────────────────────────────────────────

function sheet(overrides: Partial<Sheet> = {}): Sheet {
  return { name: "Sheet1", rows: [], ...overrides }
}

/** A 1×1 PNG-ish blob — the row/column ops never look inside it. */
function img(from: { row: number; col: number }, to?: { row: number; col: number }): SheetImage {
  return { data: new Uint8Array([1, 2, 3]), type: "png", anchor: { from, to } }
}

function rule(range: string): ConditionalRule {
  return { type: "cellIs", operator: "greaterThan", formula: "0", priority: 1, range }
}

/** Grid of `rows`×`cols` labelled cells, so shifts are visible in assertions. */
function grid(rows: number, cols: number): (string | null)[][] {
  return Array.from({ length: rows }, (_, r) =>
    Array.from({ length: cols }, (_, c) => `r${r}c${c}`),
  )
}

// ═══════════════════════════════════════════════════════════════════════
// Range rewriting — single-cell references
// ═══════════════════════════════════════════════════════════════════════

// Excel writes one-cell data validations and conditional rules with a bare
// `sqref="B3"` (no colon). The rewriter has to treat that as a degenerate
// range rather than dropping the end coordinate.
describe("range rewriting with colon-less references", () => {
  it("shifts a single-cell validation range and normalises it to start:end", () => {
    const s = sheet({
      rows: grid(4, 2),
      dataValidations: [{ type: "list", range: "B3", values: ["x"] }],
      conditionalRules: [rule("A2")],
      tables: [{ name: "T", range: "A1", columns: [{ name: "a" }] }],
    })

    insertRows(s, 0, 2)

    expect(s.dataValidations![0].range).toBe("B5:B5")
    expect(s.conditionalRules![0].range).toBe("A4:A4")
    expect(s.tables![0].range).toBe("A3:A3")
  })

  it("shifts a single-cell reference on column insert too", () => {
    const s = sheet({
      rows: grid(2, 4),
      dataValidations: [{ type: "list", range: "C1", values: ["x"] }],
      conditionalRules: [rule("D2")],
      tables: [{ name: "T", range: "C2", columns: [{ name: "a" }] }],
    })

    insertColumns(s, 0, 1)

    expect(s.dataValidations![0].range).toBe("D1:D1")
    expect(s.conditionalRules![0].range).toBe("E2:E2")
    expect(s.tables![0].range).toBe("D2:D2")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// insertRows
// ═══════════════════════════════════════════════════════════════════════

describe("insertRows", () => {
  // A sheet read from XLSX can declare more `<col>` entries than any row
  // actually populates (styled-but-empty trailing columns). The inserted
  // blank rows must be as wide as the widest of the two, or the new rows
  // come out narrower than the sheet.
  it("sizes new blank rows from the column defs when they are wider than any row", () => {
    const s = sheet({
      rows: [["a", "b"]],
      columns: [{ width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }],
    })

    insertRows(s, 0, 1)

    expect(s.rows[0]).toEqual([null, null, null, null])
  })

  it("rewrites conditional rules and table ranges that sit below the insertion", () => {
    const s = sheet({
      rows: grid(5, 2),
      conditionalRules: [rule("A3:B5")],
      tables: [{ name: "T", range: "A3:B5", columns: [{ name: "a" }, { name: "b" }] }],
    })

    insertRows(s, 1, 2)

    expect(s.conditionalRules![0].range).toBe("A5:B7")
    expect(s.tables![0].range).toBe("A5:B7")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// deleteRows
// ═══════════════════════════════════════════════════════════════════════

describe("deleteRows", () => {
  // A merge that starts above the cut and ends inside it survives, but its
  // bottom edge has to land on the last surviving row.
  it("clamps the bottom edge of a merge that ends inside the deleted range", () => {
    const s = sheet({
      rows: grid(6, 2),
      merges: [{ startRow: 1, startCol: 0, endRow: 3, endCol: 1 }],
    })

    deleteRows(s, 2, 3)

    expect(s.merges).toEqual([{ startRow: 1, startCol: 0, endRow: 1, endCol: 1 }])
  })

  it("drops images anchored inside the cut and lifts the ones below it", () => {
    const s = sheet({
      rows: grid(8, 2),
      images: [
        img({ row: 0, col: 0 }), // above the cut — untouched
        img({ row: 3, col: 0 }), // inside the cut — dropped
        img({ row: 6, col: 0 }, { row: 7, col: 1 }), // below — from and to lift
        img({ row: 1, col: 0 }, { row: 4, col: 1 }), // spans the cut — only `from` kept
      ],
    })

    deleteRows(s, 2, 3)

    expect(s.images!.map((i) => i.anchor.from.row)).toEqual([0, 3, 1])
    expect(s.images![1].anchor.to).toEqual({ row: 4, col: 1 })
    // `to` inside the deleted range is left alone — it is not >= deleteEnd.
    expect(s.images![2].anchor.to).toEqual({ row: 4, col: 1 })
  })

  it("drops tables fully inside the cut and lifts the ones below it", () => {
    const s = sheet({
      rows: grid(10, 2),
      tables: [
        { name: "Inside", range: "A3:B4", columns: [{ name: "a" }] },
        { name: "Below", range: "A7:B9", columns: [{ name: "a" }] },
        { name: "NoRange", columns: [{ name: "a" }] },
      ],
    })

    deleteRows(s, 2, 3)

    expect(s.tables!.map((t) => t.name)).toEqual(["Below", "NoRange"])
    expect(s.tables![0].range).toBe("A4:B6")
    expect(s.tables![1].range).toBeUndefined()
  })

  // The two halves of `shiftDeletedRangeRows`: a range whose top is inside
  // the cut collapses onto the first surviving row; a range whose bottom is
  // inside the cut collapses onto the last surviving row above it.
  it("clamps ranges that straddle the deleted rows", () => {
    const s = sheet({
      rows: grid(10, 2),
      dataValidations: [{ type: "list", range: "A4:B8", values: ["x"] }],
      conditionalRules: [rule("A1:B4")],
      autoFilter: { range: "A2:B6" },
    })

    deleteRows(s, 3, 3) // cut rows 4..6 (1-based)

    // Top was row 4 (inside) → clamps to the cut position; bottom row 8 lifts by 3.
    expect(s.dataValidations![0].range).toBe("A4:B5")
    // Bottom was row 4 (inside) → clamps to row 3, the last surviving row above.
    expect(s.conditionalRules![0].range).toBe("A1:B3")
    expect(s.autoFilter!.range).toBe("A2:B3")
  })

  it("keeps an auto filter that only partially overlaps the cut", () => {
    const s = sheet({ rows: grid(6, 2), autoFilter: { range: "A1:B6" } })

    deleteRows(s, 1, 2)

    expect(s.autoFilter!.range).toBe("A1:B4")
  })

  it("drops conditional rules that live entirely in the deleted rows", () => {
    const s = sheet({
      rows: grid(8, 2),
      conditionalRules: [
        { ...rule("A3:B4"), priority: 1 }, // fully inside → dropped
        { ...rule("A6:B8"), priority: 2 }, // below → lifts
      ],
    })

    deleteRows(s, 2, 3)

    expect(s.conditionalRules!.map((r) => r.range)).toEqual(["A3:B5"])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// insert/delete columns
// ═══════════════════════════════════════════════════════════════════════

describe("insertColumns / deleteColumns", () => {
  it("are no-ops for a non-positive count", () => {
    const s = sheet({ rows: grid(2, 2) })
    const before = JSON.stringify(s.rows)

    insertColumns(s, 0, 0)
    deleteColumns(s, 0, -1)

    expect(JSON.stringify(s.rows)).toBe(before)
  })

  it("shifts conditional rules and table ranges right on insert", () => {
    const s = sheet({
      rows: grid(2, 4),
      conditionalRules: [rule("C1:D2")],
      tables: [{ name: "T", range: "C1:D2", columns: [{ name: "a" }, { name: "b" }] }],
    })

    insertColumns(s, 1, 2)

    expect(s.conditionalRules![0].range).toBe("E1:F2")
    expect(s.tables![0].range).toBe("E1:F2")
  })

  it("clamps merges that straddle the deleted columns", () => {
    const s = sheet({
      rows: grid(3, 8),
      merges: [
        { startRow: 0, startCol: 2, endRow: 0, endCol: 6 }, // starts inside, ends after
        { startRow: 1, startCol: 0, endRow: 1, endCol: 3 }, // starts before, ends inside
        { startRow: 2, startCol: 0, endRow: 2, endCol: 6 }, // spans the whole cut
      ],
    })

    deleteColumns(s, 2, 3) // cut columns C..E

    expect(s.merges).toEqual([
      { startRow: 0, startCol: 2, endRow: 0, endCol: 3 },
      { startRow: 1, startCol: 0, endRow: 1, endCol: 1 },
      { startRow: 2, startCol: 0, endRow: 2, endCol: 3 },
    ])
  })

  it("drops conditional rules fully inside the cut and clamps the rest", () => {
    const s = sheet({
      rows: grid(2, 8),
      conditionalRules: [
        { ...rule("C1:D2"), priority: 1 }, // fully inside → dropped
        { ...rule("A1:D2"), priority: 2 }, // right edge inside → clamps
        { ...rule("C1:H2"), priority: 3 }, // left edge inside → clamps
      ],
    })

    deleteColumns(s, 2, 3)

    expect(s.conditionalRules!.map((r) => r.range)).toEqual(["A1:B2", "C1:E2"])
  })

  it("keeps an auto filter that only partially overlaps the deleted columns", () => {
    const s = sheet({ rows: grid(2, 6), autoFilter: { range: "A1:F2" } })

    deleteColumns(s, 1, 2)

    expect(s.autoFilter!.range).toBe("A1:D2")
  })

  it("drops images anchored inside the cut and pulls the ones to its right", () => {
    const s = sheet({
      rows: grid(2, 8),
      images: [
        img({ row: 0, col: 0 }),
        img({ row: 0, col: 3 }), // inside → dropped
        img({ row: 0, col: 6 }, { row: 1, col: 7 }), // right of cut → both shift
        img({ row: 0, col: 1 }, { row: 1, col: 4 }), // straddles → `to` stays
      ],
    })

    deleteColumns(s, 2, 3)

    expect(s.images!.map((i) => i.anchor.from.col)).toEqual([0, 3, 1])
    expect(s.images![1].anchor.to).toEqual({ row: 1, col: 4 })
    expect(s.images![2].anchor.to).toEqual({ row: 1, col: 4 })
  })

  it("drops tables fully inside the cut and clamps the rest", () => {
    const s = sheet({
      rows: grid(2, 8),
      tables: [
        { name: "Inside", range: "C1:D2", columns: [{ name: "a" }] },
        { name: "Right", range: "F1:H2", columns: [{ name: "a" }] },
        { name: "NoRange", columns: [{ name: "a" }] },
      ],
    })

    deleteColumns(s, 2, 3)

    expect(s.tables!.map((t) => t.name)).toEqual(["Right", "NoRange"])
    expect(s.tables![0].range).toBe("C1:E2")
    expect(s.tables![1].range).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// moveRows
// ═══════════════════════════════════════════════════════════════════════

describe("moveRows", () => {
  it("carries row definitions along with the rows they describe", () => {
    const s = sheet({
      rows: grid(4, 1),
      rowDefs: new Map([
        [0, { height: 10 }],
        [1, { height: 20 }],
        [3, { height: 40 }],
      ]),
    })

    moveRows(s, 0, 1, 3) // row 0 moves to the end

    expect(s.rows.map((r) => r[0])).toEqual(["r1c0", "r2c0", "r0c0", "r3c0"])
    expect(s.rowDefs!.get(2)).toEqual({ height: 10 }) // the moved row's def
    expect(s.rowDefs!.get(0)).toEqual({ height: 20 })
    expect(s.rowDefs!.get(3)).toEqual({ height: 40 })
  })

  // An empty override Map is indistinguishable from "no overrides", and the
  // writers branch on `sheet.cells` being present at all — so the move
  // clears it rather than leaving an empty Map behind.
  it("drops empty cell and rowDef maps instead of keeping them empty", () => {
    const s = sheet({ rows: grid(3, 1), cells: new Map(), rowDefs: new Map() })

    moveRows(s, 0, 1, 2)

    expect(s.cells).toBeUndefined()
    expect(s.rowDefs).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// cloneSheet — deep-copy coverage of the style/cell trees
// ═══════════════════════════════════════════════════════════════════════

describe("cloneSheet", () => {
  it("deep-copies a gradient fill without sharing the stop array", () => {
    const style: CellStyle = {
      fill: {
        type: "gradient",
        degree: 90,
        stops: [
          { position: 0, color: { rgb: "FFFF0000" } },
          { position: 1, color: { rgb: "FF0000FF" } },
        ],
      },
    }
    const s = sheet({
      rows: [["x"]],
      cells: new Map([["0,0", { value: "x", type: "string", style }]]),
    })

    const c = cloneSheet(s, "Copy")
    const clonedFill = c.cells!.get("0,0")!.style!.fill!
    if (clonedFill.type !== "gradient") throw new Error("expected a gradient fill")

    expect(clonedFill.degree).toBe(90)
    expect(clonedFill.stops).toHaveLength(2)
    clonedFill.stops[0].color.rgb = "FF00FF00"
    const originalFill = s.cells!.get("0,0")!.style!.fill!
    if (originalFill.type !== "gradient") throw new Error("expected a gradient fill")
    expect(originalFill.stops[0].color.rgb).toBe("FFFF0000")
  })

  it("keeps a pattern fill with no explicit colours as colourless", () => {
    const style: CellStyle = { fill: { type: "pattern", pattern: "gray125" } }
    const s = sheet({
      rows: [["x"]],
      cells: new Map([["0,0", { value: "x", type: "string", style }]]),
    })

    const clonedFill = cloneSheet(s, "Copy").cells!.get("0,0")!.style!.fill!
    if (clonedFill.type !== "pattern") throw new Error("expected a pattern fill")

    expect(clonedFill.fgColor).toBeUndefined()
    expect(clonedFill.bgColor).toBeUndefined()
  })

  // The overwhelmingly common solid fill in the wild sets only fgColor;
  // bgColor is meaningful for the striped patterns.
  it("deep-copies a pattern fill's foreground and background independently", () => {
    const both: CellStyle = {
      fill: {
        type: "pattern",
        pattern: "lightGrid",
        fgColor: { rgb: "FF000000" },
        bgColor: { rgb: "FFFFFFFF" },
      },
    }
    const fgOnly: CellStyle = {
      fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFFFF00" } },
    }
    const s = sheet({
      rows: [["x", "y"]],
      cells: new Map<string, Cell>([
        ["0,0", { value: "x", type: "string", style: both }],
        ["0,1", { value: "y", type: "string", style: fgOnly }],
      ]),
    })

    const c = cloneSheet(s, "Copy")
    const a = c.cells!.get("0,0")!.style!.fill!
    const b = c.cells!.get("0,1")!.style!.fill!
    if (a.type !== "pattern" || b.type !== "pattern") throw new Error("expected pattern fills")

    expect(b.bgColor).toBeUndefined()
    a.bgColor!.rgb = "FF123456"
    const originalA = s.cells!.get("0,0")!.style!.fill!
    if (originalA.type !== "pattern") throw new Error("expected a pattern fill")
    expect(originalA.bgColor!.rgb).toBe("FFFFFFFF")
  })

  it("deep-copies every border side, with and without a colour", () => {
    const style: CellStyle = {
      border: {
        top: { style: "thin", color: { rgb: "FF111111" } },
        right: { style: "thin" },
        bottom: { style: "medium", color: { rgb: "FF222222" } },
        left: { style: "dashed" },
        diagonal: { style: "thick", color: { rgb: "FF333333" } },
        diagonalUp: true,
      },
      alignment: { horizontal: "center" },
      numFmt: "0.00",
      protection: { locked: false },
    }
    const s = sheet({
      rows: [["x"]],
      cells: new Map([["0,0", { value: "x", type: "string", style }]]),
    })

    const c = cloneSheet(s, "Copy")
    const border = c.cells!.get("0,0")!.style!.border!

    expect(border.top!.color).toEqual({ rgb: "FF111111" })
    expect(border.right!.color).toBeUndefined()
    expect(border.left!.color).toBeUndefined()
    expect(border.diagonal!.color).toEqual({ rgb: "FF333333" })
    expect(border.diagonalUp).toBe(true)

    border.bottom!.color!.rgb = "FFFFFFFF"
    expect(s.cells!.get("0,0")!.style!.border!.bottom!.color!.rgb).toBe("FF222222")
  })

  // Partial borders are the norm — a header underline is a bottom edge and
  // nothing else. Every side has to survive being absent, and every side
  // has to survive carrying no explicit colour.
  it("preserves partial borders and colourless sides", () => {
    const mixed: CellStyle = {
      border: {
        top: { style: "thin" }, // present, no colour
        right: { style: "thin", color: { rgb: "FF111111" } },
        left: { style: "thin", color: { rgb: "FF222222" } },
        diagonal: { style: "thin" }, // present, no colour
        // no bottom at all
      },
    }
    const underlineOnly: CellStyle = { border: { bottom: { style: "medium" } } }
    const s = sheet({
      rows: [["x", "y"]],
      cells: new Map<string, Cell>([
        ["0,0", { value: "x", type: "string", style: mixed }],
        ["0,1", { value: "y", type: "string", style: underlineOnly }],
      ]),
    })

    const c = cloneSheet(s, "Copy")
    const a = c.cells!.get("0,0")!.style!.border!
    const b = c.cells!.get("0,1")!.style!.border!

    expect(a.top!.color).toBeUndefined()
    expect(a.diagonal!.color).toBeUndefined()
    expect(a.bottom).toBeUndefined()
    expect(a.right!.color).toEqual({ rgb: "FF111111" })
    expect(a.left!.color).toEqual({ rgb: "FF222222" })

    expect(b.top).toBeUndefined()
    expect(b.right).toBeUndefined()
    expect(b.left).toBeUndefined()
    expect(b.diagonal).toBeUndefined()
    expect(b.bottom).toEqual({ style: "medium", color: undefined })
  })

  it("deep-copies formulas, cached results and rich text runs", () => {
    const cell: Cell = {
      value: 3,
      type: "number",
      formula: "SUM(A1:A2)",
      formulaResult: 3,
      richText: [
        { text: "red", font: { bold: true, color: { rgb: "FFFF0000" } } },
        { text: " bold", font: { bold: true } }, // a font with no colour
        { text: " plain" }, // a run with no font at all
      ],
    }
    const s = sheet({ rows: [[3]], cells: new Map([["0,0", cell]]) })

    const cloned = cloneSheet(s, "Copy").cells!.get("0,0")!

    expect(cloned.formula).toBe("SUM(A1:A2)")
    expect(cloned.formulaResult).toBe(3)
    expect(cloned.richText![1].font!.color).toBeUndefined()
    expect(cloned.richText![2].font).toBeUndefined()
    cloned.richText![0].font!.color!.rgb = "FF00FF00"
    expect(s.cells!.get("0,0")!.richText![0].font!.color!.rgb).toBe("FFFF0000")
  })

  it("deep-copies a comment's rich text runs", () => {
    const cell: Cell = {
      value: "x",
      type: "string",
      comment: {
        text: "note",
        author: "QA",
        richText: [{ text: "red", font: { color: { rgb: "FFFF0000" } } }, { text: "plain" }],
      },
    }
    const s = sheet({ rows: [["x"]], cells: new Map([["0,0", cell]]) })

    const cloned = cloneSheet(s, "Copy").cells!.get("0,0")!

    expect(cloned.comment!.richText![1].font).toBeUndefined()
    cloned.comment!.richText![0].font!.color!.rgb = "FF00FF00"
    expect(s.cells!.get("0,0")!.comment!.richText![0].font!.color!.rgb).toBe("FFFF0000")
  })

  it("deep-copies a column-level style", () => {
    const s = sheet({
      rows: [["x"]],
      columns: [
        { width: 12, style: { font: { bold: true, color: { rgb: "FF010203" } } } },
        { width: 8 },
      ],
    })

    const c = cloneSheet(s, "Copy")

    expect(c.columns![1].style).toBeUndefined()
    c.columns![0].style!.font!.color!.rgb = "FFFFFFFF"
    expect(s.columns![0].style!.font!.color!.rgb).toBe("FF010203")
  })

  it("deep-copies data bar, icon set and multi-formula conditional rules", () => {
    const s = sheet({
      rows: [[1]],
      conditionalRules: [
        {
          type: "dataBar",
          priority: 1,
          range: "A1:A5",
          dataBar: { cfvo: [{ type: "min" }, { type: "max" }], color: "FF638EC6" },
        },
        {
          type: "iconSet",
          priority: 2,
          range: "A1:A5",
          iconSet: {
            iconSet: "3TrafficLights1",
            cfvo: [
              { type: "percent", value: "0" },
              { type: "percent", value: "67" },
            ],
            reverse: true,
          },
        },
        { type: "cellIs", operator: "between", priority: 3, range: "A1:A5", formula: ["1", "5"] },
      ],
    })

    const c = cloneSheet(s, "Copy")

    c.conditionalRules![0].dataBar!.cfvo[0].type = "num"
    c.conditionalRules![1].iconSet!.cfvo[0].value = "99"
    ;(c.conditionalRules![2].formula as string[])[0] = "42"

    expect(s.conditionalRules![0].dataBar!.cfvo[0].type).toBe("min")
    expect(s.conditionalRules![1].iconSet!.cfvo[0].value).toBe("0")
    expect((s.conditionalRules![2].formula as string[])[0]).toBe("1")
  })

  it("copies the sheet-level odds and ends: anchors, margins, tab colour and visibility", () => {
    const s = sheet({
      rows: [["x"]],
      images: [img({ row: 0, col: 0 }, { row: 2, col: 2 }), img({ row: 3, col: 0 })],
      pageSetup: { orientation: "landscape", margins: { left: 1, right: 1, top: 1, bottom: 1 } },
      view: { tabColor: { rgb: "FF00B050" }, showGridLines: false },
      hidden: true,
      veryHidden: false,
    })

    const c = cloneSheet(s, "Copy")

    expect(c.images![0].anchor.to).toEqual({ row: 2, col: 2 })
    expect(c.images![1].anchor.to).toBeUndefined()
    c.images![0].anchor.to!.row = 9
    expect(s.images![0].anchor.to!.row).toBe(2)

    c.pageSetup!.margins!.left = 5
    expect(s.pageSetup!.margins!.left).toBe(1)

    c.view!.tabColor!.rgb = "FFFFFFFF"
    expect(s.view!.tabColor!.rgb).toBe("FF00B050")

    expect(c.hidden).toBe(true)
    expect(c.veryHidden).toBe(false)
  })

  it("leaves margins and tab colour undefined when the source has none", () => {
    const s = sheet({
      rows: [["x"]],
      pageSetup: { orientation: "portrait" },
      view: { showGridLines: false },
    })

    const c = cloneSheet(s, "Copy")

    expect(c.pageSetup!.margins).toBeUndefined()
    expect(c.view!.tabColor).toBeUndefined()
    expect(c.view!.showGridLines).toBe(false)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// copyRange
// ═══════════════════════════════════════════════════════════════════════

describe("copyRange", () => {
  it("reads out-of-bounds source cells as null and clears stale target overrides", () => {
    const s = sheet({
      rows: [["a"], ["b"]],
      // The target already carries an override that the (empty) source must clear.
      cells: new Map<string, Cell>([
        ["0,0", { value: "a", type: "string", style: { font: { bold: true } } }],
        ["5,1", { value: "stale", type: "string" }],
      ]),
    })

    // Source B1:B2 is past the end of every row; target starts at row 5.
    copyRange(s, { startRow: 0, startCol: 1, endRow: 1, endCol: 1 }, { startRow: 5, startCol: 1 })

    expect(s.rows[5][1]).toBeNull()
    expect(s.rows[6][1]).toBeNull()
    expect(s.cells!.has("5,1")).toBe(false)
    expect(s.cells!.get("0,0")!.style!.font!.bold).toBe(true)
  })

  it("reads rows past the end of the sheet as null", () => {
    const s = sheet({ rows: [["a", "b"]] })

    // Row 3 does not exist at read time — it must land as null, not throw.
    copyRange(s, { startRow: 3, startCol: 0, endRow: 3, endCol: 1 }, { startRow: 0, startCol: 0 })

    expect(s.rows[0]).toEqual([null, null])
  })

  it("does not duplicate a merge that the target range already has", () => {
    const s = sheet({
      rows: grid(6, 4),
      merges: [
        { startRow: 0, startCol: 0, endRow: 0, endCol: 1 }, // the source merge
        { startRow: 3, startCol: 2, endRow: 3, endCol: 3 }, // same top row, other columns
        { startRow: 3, startCol: 0, endRow: 4, endCol: 1 }, // same corner, taller
      ],
    })

    const src = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }
    copyRange(s, src, { startRow: 3, startCol: 0 })
    const afterFirst = s.merges!.length
    copyRange(s, src, { startRow: 3, startCol: 0 })

    expect(afterFirst).toBe(4)
    expect(s.merges).toHaveLength(4) // second copy adds nothing
  })
})

// ═══════════════════════════════════════════════════════════════════════
// replaceCells
// ═══════════════════════════════════════════════════════════════════════

describe("replaceCells", () => {
  // A row built by index assignment (`row[3] = x`) has holes; the scan reads
  // a hole as null so `replaceCells(sheet, null, …)` fills it like any other
  // blank instead of comparing against `undefined`.
  it("reads holes inside a row as null", () => {
    const sparse: (string | null)[] = ["keep"]
    sparse[3] = "target"
    const s = sheet({ rows: [sparse] })

    expect(replaceCells(s, null, "filled")).toBe(2)
    expect(s.rows[0]).toEqual(["keep", "filled", "filled", "target"])
  })

  // `String.replace` only makes sense for a string replacement; a non-string
  // one replaces the whole cell so the type is not silently stringified.
  it("swaps the whole cell when a RegExp find is paired with a non-string replacement", () => {
    const s = sheet({
      rows: [
        ["N/A", "12"],
        ["N/A", "ok"],
      ],
    })

    expect(replaceCells(s, /^N\/A$/, 0)).toBe(2)
    expect(s.rows[0][0]).toBe(0)
    expect(s.rows[1][0]).toBe(0)
    expect(s.rows[0][1]).toBe("12")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// sortRows
// ═══════════════════════════════════════════════════════════════════════

describe("sortRows", () => {
  it("treats a row shorter than the sort column as blank and sorts it last", () => {
    const s = sheet({ rows: [[1], [], [3], [2]] })

    sortRows(s, 0)

    expect(s.rows.map((r) => r[0] ?? null)).toEqual([1, 2, 3, null])
  })

  it('reverses the order for "desc"', () => {
    const s = sheet({ rows: [[1], [3], [2]] })

    sortRows(s, 0, "desc")

    expect(s.rows.map((r) => r[0])).toEqual([3, 2, 1])
  })

  // BUG (see report): `sortRows` documents "nulls last" without qualifying
  // the order, and Excel always sinks blank cells to the bottom regardless
  // of sort direction. `compareCellValues` returns +1 for a null `a`, and
  // "desc" negates the whole comparison, so blanks float to the top.
  // src/sheet-ops.ts:1270 (and the mirror at 1297).
  it("keeps blanks last when sorting descending, as Excel does", () => {
    const s = sheet({ rows: [[1], [], [3], [2]] })

    sortRows(s, 0, "desc")

    expect(s.rows.map((r) => r[0] ?? null)).toEqual([3, 2, 1, null])
  })

  it("remaps the cell override map when sorting descending", () => {
    const s = sheet({
      rows: [["b"], ["c"], ["a"], []],
      cells: new Map<string, Cell>([
        ["0,0", { value: "b", type: "string" }],
        ["1,0", { value: "c", type: "string" }],
        ["2,0", { value: "a", type: "string" }],
        // A key for a row that does not exist — left exactly as it is.
        ["99,0", { value: "orphan", type: "string" }],
      ]),
    })

    sortRows(s, 0, "desc")

    // The short row carries no value, so it sorts as a blank — and blanks
    // sink in both directions now (#392). This test previously asserted
    // the pre-fix order, with the blank floating to the top.
    expect(s.rows.map((r) => r[0] ?? null)).toEqual(["c", "b", "a", null])
    expect(s.cells!.get("0,0")!.value).toBe("c")
    expect(s.cells!.get("1,0")!.value).toBe("b")
    expect(s.cells!.get("2,0")!.value).toBe("a")
    expect(s.cells!.get("99,0")!.value).toBe("orphan")
  })

  it("sorts FALSE before TRUE", () => {
    const s = sheet({ rows: [[true], [false], [true], [false]] })

    sortRows(s, 0)

    expect(s.rows.map((r) => r[0])).toEqual([false, false, true, true])
  })

  // Sheets whose rows were assembled column-by-column can be ragged and
  // hole-ridden; the sort has to read those positions as blanks on both
  // sides of the comparison rather than throwing.
  it("handles ragged and hole-ridden rows while remapping overrides", () => {
    const withHole: CellValue[] = ["x"]
    withHole[2] = "y" // index 1 is a hole
    const s = sheet({
      rows: [withHole, ["b"], [], ["a"], []],
      cells: new Map<string, Cell>([["0,0", { value: "x", type: "string" }]]),
    })

    sortRows(s, 1)

    expect(s.rows[0]![1]).toBeUndefined()
    expect(s.rows.map((r) => r[0] ?? null)).toEqual(["x", "b", null, "a", null])
    expect(s.cells!.get("0,0")!.value).toBe("x")
  })

  it("orders dates chronologically, ahead of strings and booleans", () => {
    const s = sheet({
      rows: [
        ["zeta"],
        [new Date(Date.UTC(2020, 0, 2))],
        [true],
        [new Date(Date.UTC(2019, 5, 1))],
        [false],
        [7],
      ],
    })

    sortRows(s, 0)

    const values = s.rows.map((r) => r[0])
    expect(values[0]).toBe(7)
    expect((values[1] as Date).getUTCFullYear()).toBe(2019)
    expect((values[2] as Date).getUTCFullYear()).toBe(2020)
    expect(values[3]).toBe("zeta")
    expect(values[4]).toBe(false)
    expect(values[5]).toBe(true)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// A workbook-shaped smoke test
// ═══════════════════════════════════════════════════════════════════════

describe("row/column ops on a workbook sheet", () => {
  it("keeps merges, validations, images and tables consistent through insert + delete", () => {
    const wb: Workbook = {
      sheets: [
        sheet({
          rows: grid(6, 4),
          merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 3 }],
          dataValidations: [{ type: "list", range: "A2:A6", values: ["x", "y"] }],
          autoFilter: { range: "A1:D6" },
          images: [img({ row: 4, col: 2 }, { row: 5, col: 3 })],
          tables: [{ name: "T", range: "A1:D6", columns: [{ name: "a" }] }],
        }),
      ],
    }
    const s = wb.sheets[0]

    insertRows(s, 1, 2)
    expect(s.dataValidations![0].range).toBe("A4:A8")
    expect(s.autoFilter!.range).toBe("A1:D8")
    expect(s.images![0].anchor.from.row).toBe(6)
    expect(s.tables![0].range).toBe("A1:D8")

    deleteRows(s, 1, 2)
    expect(s.dataValidations![0].range).toBe("A2:A6")
    expect(s.autoFilter!.range).toBe("A1:D6")
    expect(s.images![0].anchor.from.row).toBe(4)
    expect(s.tables![0].range).toBe("A1:D6")
    expect(s.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 3 }])
  })
})

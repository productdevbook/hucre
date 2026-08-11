import { describe, expect, it } from "vitest"
import { fieldsOf } from "./_reflect"
import { readFileSync } from "node:fs"
import { WorkbookBuilder } from "../src/builder"
import { readXlsx } from "../src/xlsx/reader"

// ═══════════════════════════════════════════════════════════════════════
// #439 §AJ — `SheetBuilder`'s entire state was eight fields:
//
//   _columns _rows _merges _freezePane _validations _cells _hidden _veryHidden
//
// out of `WriteSheet`'s 28, and there was no escape hatch. The first
// sheet needing a page setup or a conditional rule had to abandon the
// builder entirely. `WorkbookBuilder` reached four of `WriteOptions`.
//
// The named methods below cover what a builder is for; `set` covers the
// rest, so the class cannot fall behind the type again.
// ═══════════════════════════════════════════════════════════════════════

describe("the builder can express the whole model", () => {
  it("reaches every WriteSheet field", async () => {
    // `set` takes anything on the type, so this is the guarantee: whatever
    // `WriteSheet` grows, the builder can already say it.
    const builder = WorkbookBuilder.create().addSheet("S").row(["a"])

    for (const field of fieldsOf("WriteSheet")) {
      if (field === "name") continue
      expect(() => builder.set({ [field]: undefined } as never), field).not.toThrow()
    }
  })

  it("reaches every WriteOptions field", () => {
    const builder = WorkbookBuilder.create()

    for (const field of fieldsOf("WriteOptions")) {
      if (field === "sheets") continue
      expect(() => builder.set({ [field]: undefined } as never), field).not.toThrow()
    }
  })
})

describe("the named methods produce a readable workbook", () => {
  it("carries a page setup, a view, a header/footer and protection", async () => {
    const bytes = await WorkbookBuilder.create()
      .addSheet("S")
      .row(["a", 1])
      .pageSetup({ orientation: "landscape", paperSize: "a4", printArea: "A1:B1" })
      .view({ showGridLines: false, zoomScale: 125 })
      .headerFooter({ oddHeader: "&LLeft" })
      .protect({ sheet: true })
      .build()

    const sheet = (await readXlsx(bytes)).sheets[0]!

    expect(sheet.pageSetup).toMatchObject({ orientation: "landscape", paperSize: "a4" })
    expect(sheet.view).toMatchObject({ showGridLines: false, zoomScale: 125 })
    expect(sheet.headerFooter!.oddHeader).toBe("&LLeft")
    expect(sheet.protection!.sheet).toBe(true)
  })

  it("carries conditional rules, an auto-filter, a table and row definitions", async () => {
    const bytes = await WorkbookBuilder.create()
      .addSheet("S")
      .rows([
        ["h", "i"],
        [1, 2],
      ])
      .conditionalRule({
        type: "cellIs",
        priority: 1,
        range: "A2:A2",
        operator: "greaterThan",
        formula: ["0"],
      })
      .autoFilter({ range: "A1:B2" })
      .table({ name: "T1", range: "A1:B2", columns: [{ name: "h" }, { name: "i" }] })
      .rowDef(0, { height: 30 })
      .split(2000, 1000)
      .build()

    const sheet = (await readXlsx(bytes)).sheets[0]!

    expect(sheet.conditionalRules).toHaveLength(1)
    expect(sheet.autoFilter!.range).toBe("A1:B2")
    expect(sheet.tables).toHaveLength(1)
    expect(sheet.rowDefs!.get(0)!.height).toBe(30)
    expect(sheet.splitPane).toMatchObject({ xSplit: 2000, ySplit: 1000 })
  })

  it("carries workbook-level named ranges and protection", async () => {
    const bytes = await WorkbookBuilder.create()
      .namedRanges([{ name: "N", range: "S!$A$1" }])
      .protect({ lockStructure: true })
      .addSheet("S")
      .row(["a"])
      .build()

    const wb = await readXlsx(bytes)

    expect(wb.namedRanges![0]!.name).toBe("N")
    expect(wb.workbookProtection!.lockStructure).toBe(true)
  })

  it("reaches the long tail through set()", async () => {
    const bytes = await WorkbookBuilder.create()
      .addSheet("S")
      .rows([[1], [2], [3], [4]])
      .set({
        rowBreaks: [1],
        outlineProperties: { summaryBelow: false },
        a11y: { summary: "built", headerRow: 0 },
      })
      .build()

    const sheet = (await readXlsx(bytes)).sheets[0]!

    expect(sheet.rowBreaks).toEqual([1])
    expect(sheet.outlineProperties!.summaryBelow).toBe(false)
  })

  it("still builds the simple case unchanged", async () => {
    const bytes = await WorkbookBuilder.create()
      .addSheet("S")
      .columns([{ header: "A", width: 10 }])
      .row(["x"])
      .merge(0, 0, 0, 1)
      .freeze(1)
      .build()

    const sheet = (await readXlsx(bytes)).sheets[0]!

    // `columns[].header` only generates a header row on the `data[]`
    // path; with `row()` the rows are exactly what was pushed.
    expect(sheet.rows[0]).toEqual(["x"])
    expect(sheet.columns![0]!.width).toBe(10)
    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }])
    expect(sheet.freezePane).toEqual({ rows: 1 })
  })

  it("lets a named method and set() coexist, with the later call winning", async () => {
    const bytes = await WorkbookBuilder.create()
      .addSheet("S")
      .row(["a"])
      .view({ zoomScale: 100 })
      .set({ view: { zoomScale: 150 } })
      .build()

    expect((await readXlsx(bytes)).sheets[0]!.view!.zoomScale).toBe(150)
  })
})

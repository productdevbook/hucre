import { describe, expect, it } from "vitest"
import { fieldsOf } from "./_reflect"
import { prepareWorkbook, prepareSheet } from "../src/_write-model"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeOds } from "../src/ods/writer"
import type { Sheet, Workbook, WriteModelDrop } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #439 §A — `readXlsx` returns `Workbook`, `writeXlsx` takes
// `WorkbookInput`, and neither is assignable to the other. Read a file,
// change a cell, write it back and `tsc` refuses, twice: once for
// `charts` (Chart vs SheetChart) and again for `pivotTables`
// (PivotTable vs WritePivotTable).
//
// So every caller wrote the same lossy converter by hand and silently
// decided which fields to drop. `prepareWorkbook` makes the decisions once
// and makes the loss observable.
// ═══════════════════════════════════════════════════════════════════════

describe("the drop list stays exhaustive", () => {
  it("names every Sheet field with no SheetInput counterpart", () => {
    const orphans = fieldsOf("Sheet").filter((f) => !fieldsOf("WritableSheet").includes(f))
    const dropped: string[] = []
    prepareSheet({ name: "S", rows: [] }, (d) => dropped.push(d.field))

    // Nothing populated, so nothing is reported — the point of this
    // assertion is the *set* the module knows about, checked below.
    expect(dropped).toEqual([])
    // Every orphan must be one the module drops on purpose.
    for (const field of orphans) {
      const probe: string[] = []
      prepareSheet({ name: "S", rows: [], [field]: [{}] } as unknown as Sheet, (d) =>
        probe.push(d.field),
      )
      expect(
        probe,
        `Sheet.${field} has no SheetInput counterpart and is not in the drop list`,
      ).toContain(field)
    }
  })

  it("names every Workbook field with no WorkbookInput counterpart", () => {
    const orphans = fieldsOf("Workbook").filter((f) => !fieldsOf("WritableWorkbook").includes(f))

    for (const field of orphans) {
      const probe: string[] = []
      prepareWorkbook({ sheets: [], [field]: [{}] } as unknown as Workbook, (d) =>
        probe.push(d.field),
      )
      expect(
        probe,
        `Workbook.${field} has no WorkbookInput counterpart and is not in the drop list`,
      ).toContain(field)
    }
  })
})

describe("prepareWorkbook", () => {
  it("makes a read workbook writable", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [
            ["a", 1],
            ["b", 2],
          ],
          merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
          freezePane: { rows: 1 },
        },
      ],
      properties: { title: "Report" },
      namedRanges: [{ name: "N", range: "S!$A$1" }],
      dateSystem: "1904",
    })

    const wb = await readXlsx(bytes, { readStyles: true })
    wb.sheets[0]!.rows[0]![0] = "edited"

    const again = await readXlsx(await writeXlsx(wb))

    expect(again.sheets[0]!.rows[0]![0]).toBe("edited")
    expect(again.sheets[0]!.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }])
    expect(again.sheets[0]!.freezePane).toEqual({ rows: 1 })
    expect(again.properties!.title).toBe("Report")
    expect(again.namedRanges![0]!.name).toBe("N")
    expect(again.dateSystem).toBe("1904")
  })

  it("feeds writeOds as well", async () => {
    const wb = await readXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["x", 1]] }] }))

    const ods = await writeOds(wb)

    expect(ods.length).toBeGreaterThan(0)
  })

  it("reports what it dropped, with a reason", () => {
    const wb: Workbook = {
      sheets: [
        {
          name: "S",
          rows: [],
          slicers: [{ name: "sl", cache: "c", caption: "C" }],
          charts: [
            { kinds: ["bar"], seriesCount: 0, series: [], anchor: { from: { row: 0, col: 0 } } },
          ],
        },
      ],
      themeColors: ["FFFFFF"],
    }

    const drops: WriteModelDrop[] = []
    prepareWorkbook(wb, (d) => drops.push(d))

    expect(drops.map((d) => d.field).sort()).toEqual(["charts", "slicers", "themeColors"])
    expect(drops.find((d) => d.field === "slicers")!.sheet).toBe("S")
    expect(drops.find((d) => d.field === "themeColors")!.sheet).toBeUndefined()
    for (const drop of drops) {
      expect(drop.reason.length).toBeGreaterThan(20)
    }
  })

  it("says nothing when nothing was dropped", () => {
    const drops: WriteModelDrop[] = []
    prepareWorkbook({ sheets: [{ name: "S", rows: [["a"]] }] }, (d) => drops.push(d))

    expect(drops).toEqual([])
  })

  it("does not report an empty collection as a loss", () => {
    const drops: WriteModelDrop[] = []
    prepareWorkbook({ sheets: [{ name: "S", rows: [], slicers: [], charts: [] }] }, (d) =>
      drops.push(d),
    )

    expect(drops).toEqual([])
  })

  it("carries the cells map across, Cell being a valid Partial<Cell>", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[1]], cells: new Map([["0,0", { value: 1, formula: "1+0" }]]) }],
    })
    const wb = await readXlsx(bytes)

    const again = await readXlsx(await writeXlsx(wb))

    expect(again.sheets[0]!.cells!.get("0,0")!.formula).toBe("1+0")
  })
})

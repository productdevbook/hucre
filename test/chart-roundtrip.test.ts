import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { addChart, getCharts } from "../src/xlsx/chart-helpers"
import { cloneChart } from "../src/xlsx/chart-clone"
import { copySheetToWorkbook } from "../src/sheet-ops"
import { ZipReader } from "../src/zip/reader"
import type { SheetChart, WriteSheet } from "../src/_types"

const decoder = new TextDecoder("utf-8")

function entries(data: Uint8Array): string[] {
  return new ZipReader(data).entries()
}

async function readPart(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data)
  const bytes = await zip.extract(path)
  return decoder.decode(bytes)
}

function dataRows(): WriteSheet {
  return {
    name: "Data",
    rows: [
      ["Quarter", "Revenue", "Forecast"],
      ["Q1", 12000, 11500],
      ["Q2", 15500, 15000],
      ["Q3", 14000, 14500],
      ["Q4", 17800, 17200],
    ],
  }
}

function chart(type: SheetChart["type"], title: string, row = 6): SheetChart {
  return {
    type,
    title,
    series: [{ name: "Revenue", values: "B2:B5", categories: "A2:A5", color: "1F77B4" }],
    anchor: { from: { row, col: 0 }, to: { row: row + 14, col: 7 } },
  }
}

describe("issue #136 — model charts survive the roundtrip (saveXlsx)", () => {
  it("a chart written fresh survives openXlsx -> saveXlsx -> reread", async () => {
    const sheet = dataRows()
    sheet.charts = [chart("column", "Quarterly Revenue")]
    const fresh = await writeXlsx({ sheets: [sheet] })

    // Open it (roundtrip model) and save it back unchanged.
    const wb = await openXlsx(fresh)
    const saved = await saveXlsx(wb)

    const re = await openXlsx(saved)
    const charts = getCharts(re)
    expect(charts.length).toBe(1)
    expect(charts[0].chart.title).toBe("Quarterly Revenue")
    expect(charts[0].chart.kinds).toContain("bar") // column maps to bar/<c:barChart> with col dir
  })

  it("a chart added to an opened workbook on a NEW sheet is emitted by saveXlsx", async () => {
    const base = await writeXlsx({ sheets: [dataRows()] })
    const wb = await openXlsx(base)

    // Append a brand-new sheet carrying a model chart.
    const newSheet = {
      name: "Dashboard",
      rows: [
        ["Region", "Sales"],
        ["North", 100],
        ["South", 200],
        ["East", 150],
        ["West", 175],
      ],
      charts: [chart("line", "Sales by Region")],
    }
    // @ts-expect-error roundtrip Sheet.charts is the read model; the writer
    // accepts write-model SheetChart entries here (issue #136 bridge).
    wb.sheets.push(newSheet)

    const saved = await saveXlsx(wb)
    const names = entries(saved)
    // A chart + drawing part pair was emitted.
    expect(names.some((n) => /^xl\/charts\/chart\d+\.xml$/.test(n))).toBe(true)
    expect(names.some((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n))).toBe(true)
    // Content types declares the chart Override.
    const ct = await readPart(saved, "[Content_Types].xml")
    expect(ct).toContain("drawingml.chart+xml")

    const re = await openXlsx(saved)
    const titles = getCharts(re).map((c) => c.chart.title)
    expect(titles).toContain("Sales by Region")
  })

  it("preserved original charts and newly added model charts coexist without index collisions", async () => {
    // Workbook A: sheet 1 has an original chart (preserved as raw parts).
    const a = dataRows()
    a.charts = [chart("bar", "Original")]
    const baseBytes = await writeXlsx({ sheets: [a] })
    const wb = await openXlsx(baseBytes)

    // Append a new sheet with its own model chart.
    const extra = {
      name: "Extra",
      rows: [
        ["X", "Y"],
        ["a", 1],
        ["b", 2],
        ["c", 3],
      ],
      charts: [chart("pie", "Added")],
    }
    // @ts-expect-error see note above
    wb.sheets.push(extra)

    const saved = await saveXlsx(wb)

    // Two distinct chart parts, no overwrite.
    const chartParts = entries(saved).filter((n) => /^xl\/charts\/chart\d+\.xml$/.test(n))
    expect(new Set(chartParts).size).toBe(chartParts.length) // unique
    expect(chartParts.length).toBeGreaterThanOrEqual(2)

    const re = await openXlsx(saved)
    const titles = getCharts(re)
      .map((c) => c.chart.title)
      .sort()
    expect(titles).toEqual(["Added", "Original"])
  })

  it("copySheetToWorkbook carries charts across workbooks and saveXlsx emits them", async () => {
    const src = dataRows()
    src.charts = [chart("column", "Carried")]
    const srcWb = await openXlsx(await writeXlsx({ sheets: [src] }))
    expect(getCharts(srcWb).length).toBe(1)

    const target = await openXlsx(await writeXlsx({ sheets: [{ name: "Blank", rows: [["x"]] }] }))
    copySheetToWorkbook(srcWb.sheets[0], target, "CopiedData")

    // In-memory: the copied sheet carries the chart model.
    expect(getCharts(target).length).toBe(1)

    // And it survives the save round-trip.
    const saved = await saveXlsx(target)
    const re = await openXlsx(saved)
    const charts = getCharts(re)
    expect(charts.length).toBe(1)
    expect(charts[0].chart.title).toBe("Carried")
  })

  it("the dashboard composition flow from the issue works end-to-end", async () => {
    // Template workbook with one of several chart kinds.
    const kinds: SheetChart["type"][] = ["line", "bar", "pie", "area"]
    const template = await openXlsx(
      await writeXlsx({
        sheets: [
          {
            ...dataRows(),
            name: "Template",
            charts: kinds.map((k, i) => chart(k, `T-${k}`, 6 + i * 16)),
          },
        ],
      }),
    )
    const parsed = getCharts(template)
    expect(parsed.length).toBe(kinds.length)

    // Compose a dashboard: clone each template chart onto a fresh sheet.
    const dashboard: WriteSheet = {
      name: "Dashboard",
      rows: [
        ["Region", "Sales"],
        ["N", 1],
        ["S", 2],
        ["E", 3],
        ["W", 4],
      ],
    }
    parsed.forEach(({ chart: c }, i) => {
      const cloned = cloneChart(c, {
        title: `Panel ${i}`,
        series: [{ name: "Sales", values: "B2:B5", categories: "A2:A5", color: "00C586" }],
        anchor: { from: { row: 6 + i * 16, col: 0 }, to: { row: 20 + i * 16, col: 7 } },
      })
      addChart(dashboard, cloned)
    })

    const out = await openXlsx(await writeXlsx({ sheets: [dashboard] }))
    const outCharts = getCharts(out)
    expect(outCharts.length).toBe(kinds.length)
    expect(outCharts.every((c) => (c.chart.title ?? "").startsWith("Panel "))).toBe(true)
  })

  it("a non-writable chart kind is skipped, not fatal, on save", async () => {
    const base = await writeXlsx({ sheets: [dataRows()] })
    const wb = await openXlsx(base)
    const newSheet = {
      name: "Mixed",
      rows: [
        ["A", "B"],
        ["x", 1],
        ["y", 2],
      ],
      // A bogus/non-writable kind alongside a valid one — only the valid
      // chart should land; the save must not throw.
      charts: [
        { kinds: ["radar"], seriesCount: 0 }, // read-model shape, non-writable
        chart("column", "Valid"),
      ],
    }
    // @ts-expect-error mixed read/write model entries to exercise the guard
    wb.sheets.push(newSheet)

    const saved = await saveXlsx(wb) // must not throw
    const re = await openXlsx(saved)
    const titles = getCharts(re).map((c) => c.chart.title)
    expect(titles).toContain("Valid")
  })
})

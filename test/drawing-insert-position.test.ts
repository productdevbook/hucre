import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { writeWorksheetXml, createSharedStrings } from "../src/xlsx/worksheet-writer"
import { createStylesCollector } from "../src/xlsx/styles-writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { getCharts } from "../src/xlsx/chart-helpers"
import { ZipReader } from "../src/zip/reader"
import type { SheetChart, WriteSheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #474 — a preserved chart drawing needs a `<drawing r:id>` in the
// regenerated worksheet, and the body is serialized before the rId is
// known, so one has to go in afterwards. That was done by searching the
// finished string for thirteen candidate successor tags and inserting
// before the first one present.
//
// Safe, because the input is hucre's own output and cell text is escaped
// — but a heuristic standing in for something the writer knows exactly.
// The writer now reports the offset; this pins where it lands.
// ═══════════════════════════════════════════════════════════════════════

const decoder = new TextDecoder("utf-8")

async function sheetXml(bytes: Uint8Array, n = 1): Promise<string> {
  return decoder.decode(await new ZipReader(bytes).extract(`xl/worksheets/sheet${n}.xml`))
}

const CHART: SheetChart = {
  type: "column",
  title: "Revenue",
  series: [{ name: "R", values: "B2:B5", categories: "A2:A5" }],
  anchor: { from: { row: 7, col: 0 }, to: { row: 21, col: 7 } },
}

function base(): WriteSheet {
  return {
    name: "Data",
    rows: [
      ["Quarter", "Revenue"],
      ["Q1", 12000],
      ["Q2", 15500],
      ["Q3", 14000],
      ["Q4", 17800],
    ],
    charts: [CHART],
  }
}

/** Index of an element in the worksheet body, or -1. */
function at(xml: string, tag: string): number {
  return xml.indexOf(`<${tag}`)
}

describe("the inserted <drawing> sits where CT_Worksheet says", () => {
  it("goes after the elements that precede it", async () => {
    // pageMargins, pageSetup, rowBreaks and colBreaks all come before
    // `drawing` in the schema's sequence.
    const sheet = base()
    sheet.pageSetup = { orientation: "landscape", margins: { top: 1 } }
    sheet.rowBreaks = [2]
    sheet.colBreaks = [1]

    const xml = await sheetXml(await saveXlsx(await openXlsx(await writeXlsx({ sheets: [sheet] }))))
    const drawing = at(xml, "drawing ")

    expect(drawing).toBeGreaterThan(0)
    expect(drawing).toBeGreaterThan(at(xml, "pageMargins"))
    expect(drawing).toBeGreaterThan(at(xml, "pageSetup"))
    expect(drawing).toBeGreaterThan(at(xml, "rowBreaks"))
    expect(drawing).toBeGreaterThan(at(xml, "colBreaks"))
  })

  it("goes before the elements that follow it", async () => {
    // `tableParts` is the case the old search existed for: it is one of
    // the thirteen, and a drawing after it is schema-invalid.
    const sheet = base()
    sheet.tables = [
      {
        name: "T1",
        displayName: "T1",
        range: "A1:B5",
        columns: [{ name: "Quarter" }, { name: "Revenue" }],
      },
    ]

    const xml = await sheetXml(await saveXlsx(await openXlsx(await writeXlsx({ sheets: [sheet] }))))

    expect(at(xml, "drawing ")).toBeGreaterThan(0)
    expect(at(xml, "drawing ")).toBeLessThan(at(xml, "tableParts"))
  })

  it("lands correctly on a sheet with none of the trailing siblings", async () => {
    // The old code's fallback path: nothing to insert before, so the
    // element goes last. Still has to be inside <worksheet>.
    const xml = await sheetXml(
      await saveXlsx(await openXlsx(await writeXlsx({ sheets: [base()] }))),
    )

    expect(at(xml, "drawing ")).toBeGreaterThan(0)
    expect(at(xml, "drawing ")).toBeLessThan(xml.indexOf("</worksheet>"))
    expect(at(xml, "drawing ")).toBeGreaterThan(at(xml, "sheetData"))
  })

  it("emits exactly one", async () => {
    const xml = await sheetXml(
      await saveXlsx(await openXlsx(await writeXlsx({ sheets: [base()] }))),
    )

    expect(xml.match(/<drawing /g)).toHaveLength(1)
  })
})

describe("the reference resolves to a real part", () => {
  it("the rId names a drawing relationship that exists", async () => {
    const saved = await saveXlsx(await openXlsx(await writeXlsx({ sheets: [base()] })))

    const xml = await sheetXml(saved)
    const rId = /<drawing r:id="(rId\d+)"/.exec(xml)?.[1]
    expect(rId).toBeDefined()

    const rels = decoder.decode(
      await new ZipReader(saved).extract("xl/worksheets/_rels/sheet1.xml.rels"),
    )
    const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)?.[1]
    expect(target).toBeDefined()

    const entries = new ZipReader(saved).entries()
    expect(entries).toContain(`xl/${target!.replace(/^\.\.\//, "")}`)
  })

  it("and the chart still reads back", async () => {
    // The point of the insertion in the first place.
    const charts = getCharts(
      await openXlsx(await saveXlsx(await openXlsx(await writeXlsx({ sheets: [base()] })))),
    )

    expect(charts).toHaveLength(1)
    expect(charts[0].chart.title).toBe("Revenue")
  })
})

// ── The contract itself ──────────────────────────────────────────────
//
// The tests above pass against the old string search too — this is a
// refactor, and both put the element in a schema-valid place. What is
// new is that the position is *reported* rather than rediscovered, so
// that is what this asserts directly.

describe("drawingInsertOffset is where the writer's own <drawing> goes", () => {
  it("points exactly at the element on a sheet that has one", () => {
    // A sheet with an image makes the writer emit its own `<drawing>`.
    // The reported offset has to be that element's index — same slot,
    // filled rather than empty.
    const result = writeWorksheetXml(
      {
        name: "S",
        rows: [["a"]],
        images: [
          {
            data: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
            type: "png",
            anchor: { from: { row: 0, col: 2 } },
          },
        ],
      },
      createStylesCollector(),
      createSharedStrings(),
    )

    expect(result.drawingRId).not.toBeNull()
    expect(result.drawingInsertOffset).toBe(result.xml.indexOf("<drawing "))
  })

  it("points at the gap where one would go on a sheet that has none", () => {
    const result = writeWorksheetXml(
      { name: "S", rows: [["a"]] },
      createStylesCollector(),
      createSharedStrings(),
    )

    expect(result.drawingRId).toBeNull()
    expect(result.drawingInsertOffset).toBeGreaterThan(result.xml.indexOf("<sheetData"))
    expect(result.drawingInsertOffset).toBeLessThanOrEqual(result.xml.indexOf("</worksheet>"))

    // And splicing there produces a body whose drawing is in sequence.
    const spliced =
      result.xml.slice(0, result.drawingInsertOffset) +
      '<drawing r:id="rId1"/>' +
      result.xml.slice(result.drawingInsertOffset)

    expect(spliced.indexOf("<drawing ")).toBeGreaterThan(spliced.indexOf("<sheetData"))
    expect(spliced.indexOf("<drawing ")).toBeLessThan(spliced.indexOf("</worksheet>"))
  })
})

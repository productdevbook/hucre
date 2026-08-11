import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { getCharts } from "../src/xlsx/chart-helpers"
import { ZipReader } from "../src/zip/reader"
import type { SheetChart, WriteSheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #465 — a worksheet carries exactly one `<drawing>` element, so images
// and charts on the same sheet have to share one drawing part. The
// roundtrip built that part twice over: once for the images, and once,
// separately, for the charts — and then skipped the chart pass entirely
// for any sheet that had already produced an image drawing.
//
// The visible result was that the chart vanished. Not the anchor, not
// the styling: the whole chart, silently, on save. This proves both live
// in the one regenerated drawing, with their anchors intact.
// ═══════════════════════════════════════════════════════════════════════

const decoder = new TextDecoder("utf-8")

async function readPart(data: Uint8Array, path: string): Promise<string> {
  return decoder.decode(await new ZipReader(data).extract(path))
}

/** A 1×1 PNG — the smallest thing that makes a sheet own a drawing. */
const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
])

function dataSheet(): WriteSheet {
  return {
    name: "Data",
    rows: [
      ["Quarter", "Revenue"],
      ["Q1", 12000],
      ["Q2", 15500],
      ["Q3", 14000],
      ["Q4", 17800],
    ],
  }
}

const CHART: SheetChart = {
  type: "column",
  title: "Quarterly Revenue",
  series: [{ name: "Revenue", values: "B2:B5", categories: "A2:A5" }],
  anchor: { from: { row: 8, col: 1 }, to: { row: 22, col: 9 } },
}

/** A workbook whose one sheet owns a hucre-managed image. */
async function withImage(): Promise<Uint8Array> {
  const sheet = dataSheet()
  sheet.images = [{ data: PNG, type: "png", anchor: { from: { row: 0, col: 4 } } }]
  return writeXlsx({ sheets: [sheet] })
}

/** Attach a model chart to an opened workbook's sheet. */
function attachChart(wb: Awaited<ReturnType<typeof openXlsx>>, index = 0): void {
  // The roundtrip Sheet is the read model; the writer accepts write-model
  // SheetChart entries here — the same bridge issue #136 established.
  ;(wb.sheets[index] as unknown as { charts: SheetChart[] }).charts = [CHART]
}

describe("a sheet with both an image and a chart keeps both", () => {
  it("emits the chart part at all", async () => {
    const wb = await openXlsx(await withImage())
    attachChart(wb)

    const saved = await saveXlsx(wb)
    const names = new ZipReader(saved).entries()

    expect(names.some((n) => /^xl\/charts\/chart\d+\.xml$/.test(n))).toBe(true)
    expect(names.some((n) => /^xl\/charts\/_rels\/chart\d+\.xml\.rels$/.test(n))).toBe(true)
    // The image is still there too — this is not a trade.
    expect(names.some((n) => /^xl\/media\/image\d+\.png$/.test(n))).toBe(true)
  })

  it("puts both in the one drawing the worksheet points at", async () => {
    const wb = await openXlsx(await withImage())
    attachChart(wb)

    const saved = await saveXlsx(wb)

    // A worksheet carries a single <drawing r:id>. Whatever it resolves
    // to has to hold the picture and the graphicFrame together.
    const sheetXml = await readPart(saved, "xl/worksheets/sheet1.xml")
    const rId = /<drawing r:id="(rId\d+)"/.exec(sheetXml)?.[1]
    expect(rId).toBeDefined()

    const rels = await readPart(saved, "xl/worksheets/_rels/sheet1.xml.rels")
    const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)?.[1]
    expect(target).toBeDefined()

    const drawing = await readPart(saved, `xl/${target!.replace(/^\.\.\//, "")}`)
    expect(drawing).toContain("<xdr:pic>")
    expect(drawing).toContain("<xdr:graphicFrame")
  })

  it("keeps the chart's anchor rather than dropping it", async () => {
    const wb = await openXlsx(await withImage())
    attachChart(wb)

    const saved = await saveXlsx(wb)
    const re = await openXlsx(saved)
    const charts = getCharts(re)

    expect(charts.length).toBe(1)
    expect(charts[0].chart.title).toBe("Quarterly Revenue")

    // The anchor is the thing #465 named. Row 8 / col 1 is where it was
    // put, and a rebuilt-without-anchor graphicFrame reads back as 0,0.
    const anchor = charts[0].chart.anchor
    expect(anchor).toBeDefined()
    expect(anchor!.from).toEqual({ row: 8, col: 1 })
    expect(anchor!.to).toEqual({ row: 22, col: 9 })
  })

  it("declares the chart in [Content_Types].xml", async () => {
    const wb = await openXlsx(await withImage())
    attachChart(wb)

    const ct = await readPart(await saveXlsx(wb), "[Content_Types].xml")

    expect(ct).toContain("drawingml.chart+xml")
    expect(ct).toContain('Extension="png"')
  })
})

describe("the case as reported", () => {
  it("a file that already had both survives a plain open-and-save", async () => {
    // No editing, no appending: open the file and write it straight back.
    // On the old code the chart came back as `[]` — the sheet had images,
    // so hucre regenerated the drawing, and the chart pass then skipped
    // the sheet for exactly that reason. Nothing warned.
    const sheet = dataSheet()
    sheet.images = [{ data: PNG, type: "png", anchor: { from: { row: 0, col: 6 } } }]
    sheet.charts = [CHART]
    const original = await writeXlsx({ sheets: [sheet] })

    const before = getCharts(await openXlsx(original))
    const after = getCharts(await openXlsx(await saveXlsx(await openXlsx(original))))

    expect(before.length).toBe(1)
    expect(after.length).toBe(1)
    expect(after[0].chart.title).toBe(before[0].chart.title)
    expect(after[0].chart.anchor).toEqual(before[0].chart.anchor)
  })
})

describe("the media numbering survives the second pass", () => {
  it("numbers the image the same whether or not a chart joins it", async () => {
    // The chart pass rebuilds the image drawing. Handing writeDrawing a
    // different start index there would point the drawing at a media
    // part that was never written.
    const plain = await openXlsx(await withImage())
    const plainNames = new ZipReader(await saveXlsx(plain))
      .entries()
      .filter((n) => n.startsWith("xl/media/"))

    const withChart = await openXlsx(await withImage())
    attachChart(withChart)
    const savedNames = new ZipReader(await saveXlsx(withChart))
      .entries()
      .filter((n) => n.startsWith("xl/media/"))

    expect(savedNames).toEqual(plainNames)

    // And the drawing references a part that exists.
    const saved = await saveXlsx(withChart)
    const all = new ZipReader(saved).entries()
    const rels = await readPart(saved, "xl/drawings/_rels/drawing1.xml.rels")
    for (const m of rels.matchAll(/Target="\.\.\/(media\/[^"]+)"/g)) {
      expect(all).toContain(`xl/${m[1]}`)
    }
  })
})

describe("a preserved foreign drawing is still left alone", () => {
  it("does not rebuild a drawing hucre did not author", async () => {
    // The SAFETY boundary: hucre extends only drawings it generated this
    // run. A sheet whose chart came from the opened package keeps its
    // original parts byte-for-byte.
    const sheet = dataSheet()
    sheet.charts = [CHART]
    const original = await writeXlsx({ sheets: [sheet] })

    const wb = await openXlsx(original)
    const saved = await saveXlsx(wb)

    expect(await readPart(saved, "xl/drawings/drawing1.xml")).toBe(
      await readPart(original, "xl/drawings/drawing1.xml"),
    )
  })
})

import { describe, it, expect } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { insertRows, insertColumns } from "../src/sheet-ops"
import type { WriteSheet } from "../src/_types"

// #567: every xdr:colOff / xdr:rowOff was a hardcoded "0", so two images
// could not sit side by side in one cell.

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3])

async function drawingXml(sheet: WriteSheet): Promise<string> {
  const xlsx = await writeXlsx({ sheets: [sheet] })
  return new TextDecoder().decode(await new ZipReader(xlsx).extract("xl/drawings/drawing1.xml"))
}

function oneImage(from: {
  row: number
  col: number
  colOff?: number
  rowOff?: number
}): WriteSheet {
  return {
    name: "S",
    rows: [["x"]],
    images: [{ data: PNG, type: "png", anchor: { from }, width: 20, height: 20 }],
  }
}

describe("drawing anchor offsets (#567)", () => {
  it("writes colOff / rowOff in EMU for a oneCellAnchor image", async () => {
    const xml = await drawingXml(oneImage({ row: 1, col: 2, colOff: 10, rowOff: 4 }))
    expect(xml).toContain("<xdr:colOff>95250</xdr:colOff>")
    expect(xml).toContain("<xdr:rowOff>38100</xdr:rowOff>")
  })

  it("two images in one cell read back side by side", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["x"]],
          images: [
            { data: PNG, type: "png", anchor: { from: { row: 0, col: 0 } }, width: 20, height: 20 },
            {
              data: PNG,
              type: "png",
              anchor: { from: { row: 0, col: 0, colOff: 25 } },
              width: 20,
              height: 20,
            },
          ],
        },
      ],
    })
    const [a, b] = (await readXlsx(xlsx)).sheets[0].images!
    expect(a.anchor.from).toEqual({ row: 0, col: 0 })
    expect(b.anchor.from).toEqual({ row: 0, col: 0, colOff: 25 })
  })

  it("round-trips from and to offsets on an image, a text box and a chart", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a", 1]],
          images: [
            {
              data: PNG,
              type: "png",
              anchor: {
                from: { row: 2, col: 1, colOff: 3, rowOff: 5 },
                to: { row: 6, col: 4, colOff: 7, rowOff: 9 },
              },
            },
          ],
          textBoxes: [
            {
              text: "t",
              anchor: {
                from: { row: 8, col: 0, colOff: 1, rowOff: 2 },
                to: { row: 10, col: 3, colOff: 4, rowOff: 6 },
              },
            },
          ],
          charts: [
            {
              type: "column",
              series: [{ name: "s", values: "S!$B$1:$B$1" }],
              anchor: {
                from: { row: 12, col: 0, colOff: 11, rowOff: 12 },
                to: { row: 20, col: 6, colOff: 13, rowOff: 14 },
              },
            },
          ],
        },
      ],
    })
    const sheet = (await readXlsx(xlsx)).sheets[0]
    expect(sheet.images![0].anchor).toEqual({
      from: { row: 2, col: 1, colOff: 3, rowOff: 5 },
      to: { row: 6, col: 4, colOff: 7, rowOff: 9 },
    })
    expect(sheet.textBoxes![0].anchor).toEqual({
      from: { row: 8, col: 0, colOff: 1, rowOff: 2 },
      to: { row: 10, col: 3, colOff: 4, rowOff: 6 },
    })
    expect(sheet.charts![0].anchor).toEqual({
      from: { row: 12, col: 0, colOff: 11, rowOff: 12 },
      to: { row: 20, col: 6, colOff: 13, rowOff: 14 },
    })
  })

  it("writes the same bytes as before when no offset is given", async () => {
    // Literal fragments, not new-vs-new: any change to the default output fails.
    const from = (col: number, row: number) =>
      `<xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`
    const to = (col: number, row: number) =>
      `<xdr:to><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`
    const oneCell = await drawingXml(oneImage({ row: 2, col: 1 }))
    expect(oneCell).toContain(from(1, 2))
    expect(oneCell).not.toContain("<xdr:to>")
    expect(await drawingXml(oneImage({ row: 2, col: 1, colOff: 0, rowOff: 0 }))).toBe(oneCell)

    const xml = await drawingXml({
      name: "S",
      rows: [["x", 1]],
      images: [
        {
          data: PNG,
          type: "png",
          anchor: { from: { row: 2, col: 1 }, to: { row: 5, col: 4 } },
        },
      ],
      textBoxes: [{ text: "t", anchor: { from: { row: 7, col: 6 } } }],
      charts: [
        {
          type: "column",
          series: [{ name: "s", values: "S!$B$1:$B$1" }],
          anchor: { from: { row: 9, col: 8 }, to: { row: 12, col: 11 } },
        },
      ],
    })
    expect(xml).toContain(from(1, 2) + to(4, 5)) // image, two-cell
    expect(xml).toContain(from(6, 7) + to(9, 10)) // text box, default to
    expect(xml).toContain(from(8, 9) + to(11, 12)) // chart
  })

  it("keeps a fractional pixel offset instead of rounding it", async () => {
    // 14288 EMU is not a multiple of 9525 (1.5001... px). It is kept
    // fractional on read and the writer returns exactly 14288.
    const xml = await drawingXml(oneImage({ row: 0, col: 0, colOff: 14288 / 9525 }))
    expect(xml).toContain("<xdr:colOff>14288</xdr:colOff>")
    const xlsx = await writeXlsx({ sheets: [oneImage({ row: 0, col: 0, colOff: 1.5 })] })
    const from = (await readXlsx(xlsx)).sheets[0].images![0].anchor.from
    expect(from.colOff).toBeCloseTo(1.5, 3)
    expect(from.colOff).not.toBe(2)
  })

  it("carries a negative offset through both ways and writes a non-finite one as 0", async () => {
    const xml = await drawingXml(oneImage({ row: 1, col: 1, colOff: -3, rowOff: Number.NaN }))
    expect(xml).toContain("<xdr:colOff>-28575</xdr:colOff>")
    expect(xml).toContain("<xdr:rowOff>0</xdr:rowOff>")
    const xlsx = await writeXlsx({ sheets: [oneImage({ row: 1, col: 1, colOff: -3 })] })
    expect((await readXlsx(xlsx)).sheets[0].images![0].anchor.from.colOff).toBe(-3)
  })

  it("moves offsets with the anchor when rows and columns are inserted", async () => {
    const xlsx = await writeXlsx({
      sheets: [oneImage({ row: 1, col: 1, colOff: 6, rowOff: 8 })],
    })
    const sheet = (await readXlsx(xlsx)).sheets[0]
    insertRows(sheet, 0, 2)
    insertColumns(sheet, 0, 1)
    expect(sheet.images![0].anchor.from).toEqual({ row: 3, col: 2, colOff: 6, rowOff: 8 })
  })
})

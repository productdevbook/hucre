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
    const a = await drawingXml(oneImage({ row: 1, col: 1 }))
    expect(a).toContain("<xdr:colOff>0</xdr:colOff>")
    expect(a).toContain("<xdr:rowOff>0</xdr:rowOff>")
    expect(await drawingXml(oneImage({ row: 1, col: 1, colOff: 0, rowOff: 0 }))).toBe(a)
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

  it("treats a negative or non-finite offset as 0", async () => {
    const xml = await drawingXml(oneImage({ row: 0, col: 0, colOff: -3, rowOff: Number.NaN }))
    expect(xml).toContain("<xdr:colOff>0</xdr:colOff>")
    expect(xml).toContain("<xdr:rowOff>0</xdr:rowOff>")
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

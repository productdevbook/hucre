import { describe, it, expect } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { writeDrawing } from "../src/xlsx/drawing-writer"
import { parseXml } from "../src/xml/parser"
import type { WriteSheet } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

const encoder = new TextEncoder()
const decoder = new TextDecoder()

async function drawingXmlOf(xlsx: Uint8Array): Promise<string> {
  return decoder.decode(await new ZipReader(xlsx).extract("xl/drawings/drawing1.xml"))
}

/** Local names of an element's child elements, in document order. */
function childNames(el: { children: Array<unknown> }): string[] {
  return el.children
    .filter((c): c is { local?: string; tag: string } => typeof c !== "string")
    .map((c) => c.local || c.tag)
}

/** Create a simple fake PNG-like image */
function fakePng(size = 64): Uint8Array {
  const data = new Uint8Array(size)
  data[0] = 0x89
  data[1] = 0x50
  data[2] = 0x4e
  data[3] = 0x47
  data[4] = 0x0d
  data[5] = 0x0a
  data[6] = 0x1a
  data[7] = 0x0a
  for (let i = 8; i < size; i++) {
    data[i] = i % 256
  }
  return data
}

/**
 * Build a minimal XLSX with a oneCellAnchor drawing, by taking a valid XLSX
 * and replacing its drawing XML with a oneCellAnchor variant.
 */
async function buildXlsxWithOneCellAnchor(
  imageData: Uint8Array,
  fromRow: number,
  fromCol: number,
  extCx: number,
  extCy: number,
): Promise<Uint8Array> {
  // First, write a normal XLSX with a twoCellAnchor image
  const sheet: WriteSheet = {
    name: "Sheet1",
    rows: [["Data"]],
    images: [
      {
        data: imageData,
        type: "png",
        anchor: { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } },
      },
    ],
  }

  const baseXlsx = await writeXlsx({ sheets: [sheet] })

  // Now rewrite the drawing XML to use oneCellAnchor instead
  const oneCellDrawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>${fromCol}</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>${fromRow}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="${extCx}" cy="${extCy}"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Picture 1"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${extCx}" cy="${extCy}"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>`

  // Read the base XLSX, replace drawing1.xml, rewrite
  const baseZip = new ZipReader(baseXlsx)
  const newZip = new ZipWriter()

  for (const entry of baseZip.entries()) {
    if (entry === "xl/drawings/drawing1.xml") {
      newZip.add(entry, encoder.encode(oneCellDrawingXml))
    } else {
      const data = await baseZip.extract(entry)
      newZip.add(entry, data)
    }
  }

  return newZip.build()
}

// ── twoCellAnchor still works ─────────────────────────────────────────

describe("twoCellAnchor (existing, verify still works)", () => {
  it("writes and reads back a twoCellAnchor image", async () => {
    const imageData = fakePng(100)
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Hello"]],
      images: [
        {
          data: imageData,
          type: "png",
          anchor: { from: { row: 1, col: 2 }, to: { row: 8, col: 6 } },
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    expect(workbook.sheets[0].images).toHaveLength(1)
    const img = workbook.sheets[0].images![0]
    expect(img.type).toBe("png")
    expect(img.data).toEqual(imageData)
    expect(img.anchor.from).toEqual({ row: 1, col: 2 })
    expect(img.anchor.to).toEqual({ row: 8, col: 6 })
  })
})

// ── oneCellAnchor reading ─────────────────────────────────────────────

describe("oneCellAnchor image reading", () => {
  it("reads image from oneCellAnchor", async () => {
    const imageData = fakePng(128)
    const fromRow = 3
    const fromCol = 2
    const extCx = 3000000 // ~315 pixels
    const extCy = 2000000 // ~210 pixels

    const xlsx = await buildXlsxWithOneCellAnchor(imageData, fromRow, fromCol, extCx, extCy)
    const workbook = await readXlsx(xlsx)

    expect(workbook.sheets[0].images).toBeDefined()
    expect(workbook.sheets[0].images).toHaveLength(1)

    const img = workbook.sheets[0].images![0]
    expect(img.type).toBe("png")
    expect(img.data).toEqual(imageData)
  })

  it("parses correct from position", async () => {
    const imageData = fakePng(64)
    const xlsx = await buildXlsxWithOneCellAnchor(imageData, 5, 7, 1000000, 500000)
    const workbook = await readXlsx(xlsx)

    const img = workbook.sheets[0].images![0]
    expect(img.anchor.from.row).toBe(5)
    expect(img.anchor.from.col).toBe(7)
  })

  it("does not have anchor.to for oneCellAnchor", async () => {
    const imageData = fakePng(64)
    const xlsx = await buildXlsxWithOneCellAnchor(imageData, 0, 0, 2000000, 1000000)
    const workbook = await readXlsx(xlsx)

    const img = workbook.sheets[0].images![0]
    // oneCellAnchor has no "to" element
    expect(img.anchor.to).toBeUndefined()
  })

  it("converts ext dimensions from EMU to pixels", async () => {
    const imageData = fakePng(64)
    // 9525 EMU = 1 pixel, so 952500 EMU = 100 pixels
    const extCx = 952500 // 100 pixels
    const extCy = 476250 // 50 pixels

    const xlsx = await buildXlsxWithOneCellAnchor(imageData, 0, 0, extCx, extCy)
    const workbook = await readXlsx(xlsx)

    const img = workbook.sheets[0].images![0]
    expect(img.width).toBe(100)
    expect(img.height).toBe(50)
  })

  it("handles large dimensions correctly", async () => {
    const imageData = fakePng(64)
    // 1920 pixels wide, 1080 pixels tall
    const extCx = 1920 * 9525 // 18288000
    const extCy = 1080 * 9525 // 10287000

    const xlsx = await buildXlsxWithOneCellAnchor(imageData, 0, 0, extCx, extCy)
    const workbook = await readXlsx(xlsx)

    const img = workbook.sheets[0].images![0]
    expect(img.width).toBe(1920)
    expect(img.height).toBe(1080)
  })
})

// ── Mixed anchors ────────────────────────────────────────────────────

describe("mixed anchor types", () => {
  it("reads both twoCellAnchor and oneCellAnchor from same drawing", async () => {
    const imageData1 = fakePng(80)

    // Build an XLSX with a twoCellAnchor image first
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [
        {
          data: imageData1,
          type: "png",
          anchor: { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } },
        },
      ],
    }

    const baseXlsx = await writeXlsx({ sheets: [sheet] })

    // Replace drawing XML with one that has both twoCellAnchor and oneCellAnchor
    const mixedDrawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="1905000" cy="952500"/>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 2"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>`

    const baseZip = new ZipReader(baseXlsx)
    const newZip = new ZipWriter()

    for (const entry of baseZip.entries()) {
      if (entry === "xl/drawings/drawing1.xml") {
        newZip.add(entry, encoder.encode(mixedDrawingXml))
      } else {
        const data = await baseZip.extract(entry)
        newZip.add(entry, data)
      }
    }

    const modifiedXlsx = await newZip.build()
    const workbook = await readXlsx(modifiedXlsx)

    expect(workbook.sheets[0].images).toHaveLength(2)

    // First image: twoCellAnchor
    const img1 = workbook.sheets[0].images![0]
    expect(img1.anchor.from).toEqual({ row: 0, col: 0 })
    expect(img1.anchor.to).toEqual({ row: 5, col: 3 })

    // Second image: oneCellAnchor
    const img2 = workbook.sheets[0].images![1]
    expect(img2.anchor.from).toEqual({ row: 10, col: 5 })
    expect(img2.anchor.to).toBeUndefined()
    // 1905000 / 9525 = 200, 952500 / 9525 = 100
    expect(img2.width).toBe(200)
    expect(img2.height).toBe(100)
  })
})

// ── oneCellAnchor writing (#598) ─────────────────────────────────────

describe("image without anchor.to is written as oneCellAnchor (#598)", () => {
  // The issue's reproduction: a 200×40 px logo anchored at B2 with no `to`.
  const logoSheet = (): WriteSheet => ({
    name: "Sheet1",
    rows: [["x"]],
    images: [
      {
        data: fakePng(),
        type: "png",
        anchor: { from: { row: 1, col: 1 } },
        width: 200,
        height: 40,
      },
    ],
  })

  it("emits oneCellAnchor with from, ext, pic, clientData and no to", async () => {
    const xml = await drawingXmlOf(await writeXlsx({ sheets: [logoSheet()] }))

    expect(xml).not.toContain("twoCellAnchor")
    expect(xml).not.toContain("<xdr:to>")
    const doc = parseXml(xml)
    expect(childNames(doc)).toEqual(["oneCellAnchor"])
    const anchor = doc.children.find((c) => typeof c !== "string") as { children: Array<unknown> }
    // CT_OneCellAnchor is a sequence: from, ext, object, clientData.
    expect(childNames(anchor)).toEqual(["from", "ext", "pic", "clientData"])
    // 200 px × 9525 EMU/px, 40 px × 9525 EMU/px
    expect(xml).toContain('<xdr:ext cx="1905000" cy="381000"/>')
  })

  it("reads back with only `from` and the written size", async () => {
    const wb = await readXlsx(await writeXlsx({ sheets: [logoSheet()] }))
    const img = wb.sheets[0].images![0]
    expect(img.anchor).toEqual({ from: { row: 1, col: 1 } })
    expect(img.width).toBe(200)
    expect(img.height).toBe(40)
  })

  it("uses the default size for ext when width/height are absent", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["x"]],
          images: [{ data: fakePng(), type: "png", anchor: { from: { row: 0, col: 0 } } }],
        },
      ],
    })
    const xml = await drawingXmlOf(xlsx)
    expect(xml).toContain("<xdr:oneCellAnchor>")
    expect(xml).toContain('<xdr:ext cx="3000000" cy="2000000"/>')

    // The default is what the reader reports, as SheetImage documents.
    const img = (await readXlsx(xlsx)).sheets[0].images![0]
    expect(img.anchor).toEqual({ from: { row: 0, col: 0 } })
    expect(img.width).toBe(315)
    expect(img.height).toBe(210)
  })

  it("keeps altText and title on the picture", async () => {
    const sheet = logoSheet()
    sheet.images![0].altText = "Company logo"
    sheet.images![0].title = "Logo"
    const xlsx = await writeXlsx({ sheets: [sheet] })

    const xml = await drawingXmlOf(xlsx)
    expect(xml).toContain("<xdr:oneCellAnchor>")
    expect(xml).toContain('<xdr:cNvPr id="2" name="Picture 1" title="Logo" descr="Company logo"/>')
    const img = (await readXlsx(xlsx)).sheets[0].images![0]
    expect(img.altText).toBe("Company logo")
    expect(img.title).toBe("Logo")
  })

  it("still writes twoCellAnchor, unchanged, when `to` is given", () => {
    const result = writeDrawing(
      [
        {
          data: fakePng(),
          type: "png",
          anchor: { from: { row: 1, col: 1 }, to: { row: 5, col: 3 } },
          width: 200,
          height: 40,
          altText: "a<b",
          title: "T",
        },
      ],
      1,
    )
    // Byte-for-byte what the writer produced before #598.
    expect(result.drawingXml).toContain(
      "<xdr:twoCellAnchor>" +
        "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
        "<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
        '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" title="T" descr="a&lt;b"/>' +
        '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
        '<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
        '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="381000"/></a:xfrm>' +
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>' +
        "<xdr:clientData/></xdr:twoCellAnchor>",
    )
    expect(result.drawingXml).not.toContain("oneCellAnchor")
  })

  it("mixes both anchor kinds on one sheet in declaration order", async () => {
    const a = fakePng(70)
    const b = fakePng(90)
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["x"]],
          images: [
            { data: a, type: "png", anchor: { from: { row: 0, col: 0 }, to: { row: 4, col: 2 } } },
            { data: b, type: "png", anchor: { from: { row: 6, col: 3 } }, width: 64, height: 32 },
          ],
        },
      ],
    })
    expect(childNames(parseXml(await drawingXmlOf(xlsx)))).toEqual([
      "twoCellAnchor",
      "oneCellAnchor",
    ])

    const images = (await readXlsx(xlsx)).sheets[0].images!
    expect(images.map((i) => i.data)).toEqual([a, b])
    expect(images[0].anchor).toEqual({ from: { row: 0, col: 0 }, to: { row: 4, col: 2 } })
    expect(images[1].anchor).toEqual({ from: { row: 6, col: 3 } })
    expect(images[1].width).toBe(64)
    expect(images[1].height).toBe(32)
  })

  it("stays a oneCellAnchor through openXlsx / saveXlsx", async () => {
    const saved = await saveXlsx(await openXlsx(await writeXlsx({ sheets: [logoSheet()] })))
    expect(await drawingXmlOf(saved)).toContain('<xdr:ext cx="1905000" cy="381000"/>')
    const img = (await readXlsx(saved)).sheets[0].images![0]
    expect(img.anchor).toEqual({ from: { row: 1, col: 1 } })
    expect(img.width).toBe(200)
    expect(img.height).toBe(40)
  })
})

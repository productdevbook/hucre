import { describe, it, expect } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { parseXml } from "../src/xml/parser"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeDrawing } from "../src/xlsx/drawing-writer"
import { writeContentTypes } from "../src/xlsx/content-types-writer"
import type { WriteSheet, SheetImage } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8")

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data)
  const raw = await zip.extract(path)
  return decoder.decode(raw)
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function zipHas(data: Uint8Array, path: string): boolean {
  const zip = new ZipReader(data)
  return zip.has(path)
}

async function zipExtract(data: Uint8Array, path: string): Promise<Uint8Array> {
  const zip = new ZipReader(data)
  return zip.extract(path)
}

/** Create a simple fake PNG-like image (minimal header + bytes) */
function fakePng(size = 64): Uint8Array {
  const data = new Uint8Array(size)
  // PNG magic bytes
  data[0] = 0x89
  data[1] = 0x50 // P
  data[2] = 0x4e // N
  data[3] = 0x47 // G
  data[4] = 0x0d
  data[5] = 0x0a
  data[6] = 0x1a
  data[7] = 0x0a
  // Fill rest with arbitrary data
  for (let i = 8; i < size; i++) {
    data[i] = i % 256
  }
  return data
}

/** Create a simple fake JPEG-like image */
function fakeJpeg(size = 64): Uint8Array {
  const data = new Uint8Array(size)
  // JPEG magic bytes (SOI marker)
  data[0] = 0xff
  data[1] = 0xd8
  data[2] = 0xff
  data[3] = 0xe0
  for (let i = 4; i < size; i++) {
    data[i] = i % 256
  }
  return data
}

/** Create a simple fake GIF-like image */
function fakeGif(size = 64): Uint8Array {
  const data = new Uint8Array(size)
  // GIF89a magic
  const magic = [0x47, 0x49, 0x46, 0x38, 0x39, 0x61] // "GIF89a"
  for (let i = 0; i < magic.length; i++) {
    data[i] = magic[i]
  }
  for (let i = 6; i < size; i++) {
    data[i] = i % 256
  }
  return data
}

function makeImage(
  type: SheetImage["type"],
  from: { row: number; col: number },
  to?: { row: number; col: number },
  opts?: { width?: number; height?: number },
): SheetImage {
  let data: Uint8Array
  if (type === "png") data = fakePng()
  else if (type === "jpeg") data = fakeJpeg()
  else data = fakeGif()

  return {
    data,
    type,
    anchor: { from, to },
    width: opts?.width,
    height: opts?.height,
  }
}

// ── writeDrawing unit tests ──────────────────────────────────────────

describe("writeDrawing", () => {
  it("generates drawing XML with correct namespaces", () => {
    const images: SheetImage[] = [makeImage("png", { row: 0, col: 0 }, { row: 5, col: 3 })]
    const result = writeDrawing(images, 1)

    expect(result.drawingXml).toContain("xdr:wsDr")
    expect(result.drawingXml).toContain(
      'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"',
    )
    expect(result.drawingXml).toContain(
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
    )
    expect(result.drawingXml).toContain(
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    )
  })

  it("generates correct anchor positions", () => {
    const images: SheetImage[] = [makeImage("png", { row: 2, col: 1 }, { row: 8, col: 5 })]
    const result = writeDrawing(images, 1)

    // Check from element
    expect(result.drawingXml).toContain("<xdr:col>1</xdr:col>")
    expect(result.drawingXml).toContain("<xdr:row>2</xdr:row>")
    // Check to element
    expect(result.drawingXml).toContain("<xdr:col>5</xdr:col>")
    expect(result.drawingXml).toContain("<xdr:row>8</xdr:row>")
  })

  it("returns image file entries with correct paths", () => {
    const images: SheetImage[] = [
      makeImage("png", { row: 0, col: 0 }),
      makeImage("jpeg", { row: 5, col: 0 }),
    ]
    const result = writeDrawing(images, 1)

    expect(result.images).toHaveLength(2)
    expect(result.images[0].path).toBe("xl/media/image1.png")
    expect(result.images[0].contentType).toBe("image/png")
    expect(result.images[1].path).toBe("xl/media/image2.jpeg")
    expect(result.images[1].contentType).toBe("image/jpeg")
  })

  it("uses global image index for naming", () => {
    const images: SheetImage[] = [makeImage("png", { row: 0, col: 0 })]
    const result = writeDrawing(images, 5)

    expect(result.images[0].path).toBe("xl/media/image5.png")
  })

  it("generates relationship entries with correct rIds", () => {
    const images: SheetImage[] = [
      makeImage("png", { row: 0, col: 0 }),
      makeImage("jpeg", { row: 5, col: 0 }),
    ]
    const result = writeDrawing(images, 1)

    expect(result.drawingRels).toContain('Id="rId1"')
    expect(result.drawingRels).toContain('Id="rId2"')
    expect(result.drawingRels).toContain("../media/image1.png")
    expect(result.drawingRels).toContain("../media/image2.jpeg")
  })

  it("embeds blip references with correct rIds", () => {
    const images: SheetImage[] = [makeImage("png", { row: 0, col: 0 })]
    const result = writeDrawing(images, 1)

    expect(result.drawingXml).toContain('r:embed="rId1"')
  })

  it("uses default to anchor when not specified", () => {
    const images: SheetImage[] = [makeImage("png", { row: 2, col: 1 })]
    const result = writeDrawing(images, 1)

    // Default to: from.col + 3, from.row + 5
    const doc = parseXml(result.drawingXml)
    const anchor = findChild(doc, "twoCellAnchor")
    expect(anchor).toBeTruthy()

    const toEl = findChild(anchor, "to")
    const toCol = findChild(toEl, "col")
    const toRow = findChild(toEl, "row")
    expect(toCol.children[0]).toBe("4") // 1 + 3
    expect(toRow.children[0]).toBe("7") // 2 + 5
  })
})

// ── Full XLSX write + read integration ──────────────────────────────

describe("XLSX image writing", () => {
  it("writes a single PNG image to the ZIP", async () => {
    const imageData = fakePng(128)
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Hello"]],
      images: [
        {
          data: imageData,
          type: "png",
          anchor: { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } },
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    // Verify ZIP contains expected files
    expect(zipHas(xlsx, "xl/media/image1.png")).toBe(true)
    expect(zipHas(xlsx, "xl/drawings/drawing1.xml")).toBe(true)
    expect(zipHas(xlsx, "xl/drawings/_rels/drawing1.xml.rels")).toBe(true)
    expect(zipHas(xlsx, "xl/worksheets/_rels/sheet1.xml.rels")).toBe(true)

    // Verify image data is preserved
    const extracted = await zipExtract(xlsx, "xl/media/image1.png")
    expect(extracted).toEqual(imageData)
  })

  it("includes drawing reference in worksheet XML", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Hello"]],
      images: [makeImage("png", { row: 0, col: 0 }, { row: 5, col: 3 })],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const wsXml = await extractXml(xlsx, "xl/worksheets/sheet1.xml")

    // Worksheet should contain <drawing r:id="rId1"/>
    expect(wsXml).toContain("drawing")
    expect(wsXml).toContain('r:id="rId1"')
  })

  it("includes drawing relationship in sheet rels", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Hello"]],
      images: [makeImage("png", { row: 0, col: 0 })],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const relsXml = await extractXml(xlsx, "xl/worksheets/_rels/sheet1.xml.rels")

    expect(relsXml).toContain("drawing")
    expect(relsXml).toContain("../drawings/drawing1.xml")
    expect(relsXml).toContain(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    )
  })

  it("writes multiple images on one sheet", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [
        makeImage("png", { row: 0, col: 0 }, { row: 3, col: 3 }),
        makeImage("jpeg", { row: 5, col: 0 }, { row: 10, col: 3 }),
        makeImage("gif", { row: 12, col: 0 }, { row: 15, col: 3 }),
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    // All three images in media
    expect(zipHas(xlsx, "xl/media/image1.png")).toBe(true)
    expect(zipHas(xlsx, "xl/media/image2.jpeg")).toBe(true)
    expect(zipHas(xlsx, "xl/media/image3.gif")).toBe(true)

    // Single drawing file with all three anchors
    const drawingXml = await extractXml(xlsx, "xl/drawings/drawing1.xml")
    const drawingDoc = parseXml(drawingXml)
    const anchors = findChildren(drawingDoc, "twoCellAnchor")
    expect(anchors).toHaveLength(3)

    // Drawing rels should have three relationships
    const drawingRels = await extractXml(xlsx, "xl/drawings/_rels/drawing1.xml.rels")
    expect(drawingRels).toContain("image1.png")
    expect(drawingRels).toContain("image2.jpeg")
    expect(drawingRels).toContain("image3.gif")
  })

  it("writes images on different sheets", async () => {
    const sheet1: WriteSheet = {
      name: "Sheet1",
      rows: [["First"]],
      images: [makeImage("png", { row: 0, col: 0 })],
    }
    const sheet2: WriteSheet = {
      name: "Sheet2",
      rows: [["Second"]],
      images: [makeImage("jpeg", { row: 1, col: 1 })],
    }

    const xlsx = await writeXlsx({ sheets: [sheet1, sheet2] })

    // Each sheet has its own drawing
    expect(zipHas(xlsx, "xl/drawings/drawing1.xml")).toBe(true)
    expect(zipHas(xlsx, "xl/drawings/drawing2.xml")).toBe(true)

    // Images have unique global indices
    expect(zipHas(xlsx, "xl/media/image1.png")).toBe(true)
    expect(zipHas(xlsx, "xl/media/image2.jpeg")).toBe(true)

    // Each sheet has its own rels
    expect(zipHas(xlsx, "xl/worksheets/_rels/sheet1.xml.rels")).toBe(true)
    expect(zipHas(xlsx, "xl/worksheets/_rels/sheet2.xml.rels")).toBe(true)
  })

  it("does not create drawing parts for sheets without images", async () => {
    const sheet1: WriteSheet = {
      name: "WithImages",
      rows: [["Has image"]],
      images: [makeImage("png", { row: 0, col: 0 })],
    }
    const sheet2: WriteSheet = {
      name: "NoImages",
      rows: [["No image"]],
    }

    const xlsx = await writeXlsx({ sheets: [sheet1, sheet2] })

    // Sheet1 has drawing
    expect(zipHas(xlsx, "xl/drawings/drawing1.xml")).toBe(true)
    // Sheet2 does not have drawing
    expect(zipHas(xlsx, "xl/drawings/drawing2.xml")).toBe(false)

    // Sheet2 worksheet XML should not have a drawing element
    const ws2Xml = await extractXml(xlsx, "xl/worksheets/sheet2.xml")
    expect(ws2Xml).not.toContain("<drawing")
  })

  it("empty images array does not create drawing parts", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    expect(zipHas(xlsx, "xl/drawings/drawing1.xml")).toBe(false)
    expect(zipHas(xlsx, "xl/media/image1.png")).toBe(false)
  })
})

// ── Content Types ────────────────────────────────────────────────────

describe("content types with images", () => {
  it("includes image extension defaults and drawing overrides", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [makeImage("png", { row: 0, col: 0 }), makeImage("jpeg", { row: 5, col: 0 })],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const ctXml = await extractXml(xlsx, "[Content_Types].xml")

    // Image extension defaults
    expect(ctXml).toContain('Extension="png"')
    expect(ctXml).toContain('ContentType="image/png"')
    expect(ctXml).toContain('Extension="jpeg"')
    expect(ctXml).toContain('ContentType="image/jpeg"')

    // Drawing override
    expect(ctXml).toContain("/xl/drawings/drawing1.xml")
    expect(ctXml).toContain("application/vnd.openxmlformats-officedocument.drawing+xml")
  })

  it("includes gif extension default", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [makeImage("gif", { row: 0, col: 0 })],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const ctXml = await extractXml(xlsx, "[Content_Types].xml")

    expect(ctXml).toContain('Extension="gif"')
    expect(ctXml).toContain('ContentType="image/gif"')
  })

  it("does not include image extensions when no images", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const ctXml = await extractXml(xlsx, "[Content_Types].xml")

    expect(ctXml).not.toContain('Extension="png"')
    expect(ctXml).not.toContain('Extension="jpeg"')
    expect(ctXml).not.toContain('Extension="gif"')
    expect(ctXml).not.toContain("drawing")
  })

  it("writeContentTypes function with drawing options", () => {
    const xml = writeContentTypes({
      sheetCount: 2,
      hasSharedStrings: false,
      drawingIndices: [1, 2],
      imageExtensions: new Set(["png", "jpeg"]),
    })

    expect(xml).toContain('Extension="png"')
    expect(xml).toContain('Extension="jpeg"')
    expect(xml).toContain("/xl/drawings/drawing1.xml")
    expect(xml).toContain("/xl/drawings/drawing2.xml")
    expect(xml).toContain("application/vnd.openxmlformats-officedocument.drawing+xml")
  })

  it("writeContentTypes backward compatibility with (number, boolean) signature", () => {
    const xml = writeContentTypes(2, true)

    expect(xml).toContain("/xl/worksheets/sheet1.xml")
    expect(xml).toContain("/xl/worksheets/sheet2.xml")
    expect(xml).toContain("/xl/sharedStrings.xml")
    expect(xml).not.toContain("drawing")
    expect(xml).not.toContain('Extension="png"')
  })
})

// ── Image anchor positions ───────────────────────────────────────────

describe("image anchor positions", () => {
  it("correctly encodes from/to cell positions in drawing XML", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [makeImage("png", { row: 3, col: 2 }, { row: 10, col: 7 })],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const drawingXml = await extractXml(xlsx, "xl/drawings/drawing1.xml")
    const doc = parseXml(drawingXml)

    const anchor = findChild(doc, "twoCellAnchor")
    const from = findChild(anchor, "from")
    const to = findChild(anchor, "to")

    const fromCol = findChild(from, "col")
    const fromRow = findChild(from, "row")
    const toCol = findChild(to, "col")
    const toRow = findChild(to, "row")

    expect(fromCol.children[0]).toBe("2")
    expect(fromRow.children[0]).toBe("3")
    expect(toCol.children[0]).toBe("7")
    expect(toRow.children[0]).toBe("10")
  })

  it("preserves pixel dimensions as EMU in spPr", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [
        makeImage("png", { row: 0, col: 0 }, { row: 5, col: 3 }, { width: 200, height: 150 }),
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const drawingXml = await extractXml(xlsx, "xl/drawings/drawing1.xml")

    // 200 * 9525 = 1905000, 150 * 9525 = 1428750
    expect(drawingXml).toContain('cx="1905000"')
    expect(drawingXml).toContain('cy="1428750"')
  })
})

// ── JPEG and GIF image types ─────────────────────────────────────────

describe("image type handling", () => {
  it("handles JPEG images correctly", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [makeImage("jpeg", { row: 0, col: 0 })],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    expect(zipHas(xlsx, "xl/media/image1.jpeg")).toBe(true)
    expect(zipHas(xlsx, "xl/media/image1.png")).toBe(false)
  })

  it("handles GIF images correctly", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [makeImage("gif", { row: 0, col: 0 })],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    expect(zipHas(xlsx, "xl/media/image1.gif")).toBe(true)
    expect(zipHas(xlsx, "xl/media/image1.png")).toBe(false)
  })

  it("handles mixed image types on one sheet", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [
        makeImage("png", { row: 0, col: 0 }),
        makeImage("jpeg", { row: 5, col: 0 }),
        makeImage("gif", { row: 10, col: 0 }),
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    expect(zipHas(xlsx, "xl/media/image1.png")).toBe(true)
    expect(zipHas(xlsx, "xl/media/image2.jpeg")).toBe(true)
    expect(zipHas(xlsx, "xl/media/image3.gif")).toBe(true)

    // Content types should have all three extensions
    const ctXml = await extractXml(xlsx, "[Content_Types].xml")
    expect(ctXml).toContain('Extension="png"')
    expect(ctXml).toContain('Extension="jpeg"')
    expect(ctXml).toContain('Extension="gif"')
  })
})

// ── Round-trip (write + read) ────────────────────────────────────────

describe("image round-trip", () => {
  it("reads back a single image written to XLSX", async () => {
    const imageData = fakePng(128)
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Hello"]],
      images: [
        {
          data: imageData,
          type: "png",
          anchor: { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } },
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    expect(workbook.sheets).toHaveLength(1)
    expect(workbook.sheets[0].images).toBeDefined()
    expect(workbook.sheets[0].images).toHaveLength(1)

    const img = workbook.sheets[0].images![0]
    expect(img.type).toBe("png")
    expect(img.data).toEqual(imageData)
    expect(img.anchor.from.row).toBe(0)
    expect(img.anchor.from.col).toBe(0)
    expect(img.anchor.to?.row).toBe(5)
    expect(img.anchor.to?.col).toBe(3)
  })

  it("reads back multiple images from one sheet", async () => {
    const pngData = fakePng(100)
    const jpegData = fakeJpeg(100)

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [
        {
          data: pngData,
          type: "png",
          anchor: { from: { row: 0, col: 0 }, to: { row: 3, col: 3 } },
        },
        {
          data: jpegData,
          type: "jpeg",
          anchor: { from: { row: 5, col: 1 }, to: { row: 10, col: 4 } },
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    expect(workbook.sheets[0].images).toHaveLength(2)

    const img1 = workbook.sheets[0].images![0]
    expect(img1.type).toBe("png")
    expect(img1.data).toEqual(pngData)
    expect(img1.anchor.from).toEqual({ row: 0, col: 0 })
    expect(img1.anchor.to).toEqual({ row: 3, col: 3 })

    const img2 = workbook.sheets[0].images![1]
    expect(img2.type).toBe("jpeg")
    expect(img2.data).toEqual(jpegData)
    expect(img2.anchor.from).toEqual({ row: 5, col: 1 })
    expect(img2.anchor.to).toEqual({ row: 10, col: 4 })
  })

  it("reads back images from multiple sheets", async () => {
    const pngData = fakePng(80)
    const jpegData = fakeJpeg(80)

    const sheet1: WriteSheet = {
      name: "Sheet1",
      rows: [["First"]],
      images: [
        {
          data: pngData,
          type: "png",
          anchor: { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } },
        },
      ],
    }
    const sheet2: WriteSheet = {
      name: "Sheet2",
      rows: [["Second"]],
      images: [
        {
          data: jpegData,
          type: "jpeg",
          anchor: { from: { row: 1, col: 1 }, to: { row: 8, col: 6 } },
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet1, sheet2] })
    const workbook = await readXlsx(xlsx)

    expect(workbook.sheets).toHaveLength(2)

    expect(workbook.sheets[0].images).toHaveLength(1)
    expect(workbook.sheets[0].images![0].type).toBe("png")
    expect(workbook.sheets[0].images![0].data).toEqual(pngData)

    expect(workbook.sheets[1].images).toHaveLength(1)
    expect(workbook.sheets[1].images![0].type).toBe("jpeg")
    expect(workbook.sheets[1].images![0].data).toEqual(jpegData)
  })

  it("sheet without images has no images array after read", async () => {
    const sheet1: WriteSheet = {
      name: "WithImages",
      rows: [["Has image"]],
      images: [makeImage("png", { row: 0, col: 0 }, { row: 5, col: 3 })],
    }
    const sheet2: WriteSheet = {
      name: "NoImages",
      rows: [["No image"]],
    }

    const xlsx = await writeXlsx({ sheets: [sheet1, sheet2] })
    const workbook = await readXlsx(xlsx)

    expect(workbook.sheets[0].images).toHaveLength(1)
    expect(workbook.sheets[1].images).toBeUndefined()
  })

  it("reads back GIF images", async () => {
    const gifData = fakeGif(96)
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [
        {
          data: gifData,
          type: "gif",
          anchor: { from: { row: 2, col: 1 }, to: { row: 7, col: 4 } },
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    expect(workbook.sheets[0].images).toHaveLength(1)
    expect(workbook.sheets[0].images![0].type).toBe("gif")
    expect(workbook.sheets[0].images![0].data).toEqual(gifData)
  })
})

// ── Edge cases ───────────────────────────────────────────────────────

describe("image edge cases", () => {
  it("coexists with hyperlinks in sheet rels", async () => {
    const cells = new Map<string, Partial<import("../src/_types").Cell>>()
    cells.set("0,0", {
      value: "Link",
      hyperlink: { target: "https://example.com" },
    })

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Link"]],
      cells,
      images: [makeImage("png", { row: 2, col: 0 }, { row: 7, col: 3 })],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const relsXml = await extractXml(xlsx, "xl/worksheets/_rels/sheet1.xml.rels")

    // Should contain both hyperlink and drawing relationships
    expect(relsXml).toContain("hyperlink")
    expect(relsXml).toContain("drawing")
    expect(relsXml).toContain("https://example.com")
    expect(relsXml).toContain("../drawings/drawing1.xml")

    // rIds should not conflict: hyperlink is rId1, drawing is rId2
    const doc = parseXml(relsXml)
    const rels = findChildren(doc, "Relationship")
    const ids = rels.map((r: any) => r.attrs["Id"])
    expect(new Set(ids).size).toBe(ids.length) // All unique
  })

  it("handles images alongside other sheet features (merges, data validations)", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Merged", null, null],
        [1, 2, 3],
      ],
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
      images: [makeImage("png", { row: 3, col: 0 }, { row: 8, col: 3 })],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    // Both features present
    const wsXml = await extractXml(xlsx, "xl/worksheets/sheet1.xml")
    expect(wsXml).toContain("mergeCells")
    expect(wsXml).toContain("drawing")

    // Image round-trips
    const workbook = await readXlsx(xlsx)
    expect(workbook.sheets[0].images).toHaveLength(1)
    expect(workbook.sheets[0].merges).toHaveLength(1)
  })
})

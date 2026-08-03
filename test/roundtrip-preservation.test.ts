import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { ZipReader } from "../src/zip/reader"
import type { WriteSheet, WriteOptions } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

/** write → open → save → read, i.e. one full preservation cycle. */
async function cycle(options: WriteOptions) {
  const original = await writeXlsx(options)
  const opened = await openXlsx(original)
  const saved = await saveXlsx(opened)
  return { saved, workbook: await readXlsx(saved) }
}

function sheetWith(extra: Partial<WriteSheet>): WriteSheet {
  return {
    name: "S",
    rows: [
      ["a", "b"],
      [1, 2],
    ],
    ...extra,
  }
}

// ═══════════════════════════════════════════════════════════════════════
// Regression guard for #359 — saveXlsx rebuilt each sheet from a
// hand-written field map that omitted nine things the reader and writer
// both understood, so open → save destroyed them.
// ═══════════════════════════════════════════════════════════════════════

describe("openXlsx → saveXlsx preserves sheet features", () => {
  it("keeps split panes", async () => {
    const { workbook } = await cycle({
      sheets: [sheetWith({ splitPane: { xSplit: 2000, ySplit: 1500 } })],
    })
    expect(workbook.sheets[0].splitPane).toMatchObject({ xSplit: 2000, ySplit: 1500 })
  })

  it("keeps manual page breaks", async () => {
    const { workbook } = await cycle({
      sheets: [sheetWith({ rowBreaks: [5, 10], colBreaks: [3] })],
    })
    expect(workbook.sheets[0].rowBreaks).toEqual([5, 10])
    expect(workbook.sheets[0].colBreaks).toEqual([3])
  })

  it("keeps the background image and its relationship", async () => {
    const png = new Uint8Array([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
      0x52,
    ])
    const { saved, workbook } = await cycle({ sheets: [sheetWith({ backgroundImage: png })] })
    expect(workbook.sheets[0].backgroundImage).toEqual(png)

    const zip = new ZipReader(saved)
    const sheetXml = new TextDecoder().decode(await zip.extract("xl/worksheets/sheet1.xml"))
    expect(sheetXml).toContain("<picture")
    expect(zip.has("xl/worksheets/_rels/sheet1.xml.rels")).toBe(true)
    expect(zip.has("xl/media/image1.png")).toBe(true)
  })

  it("keeps sparklines", async () => {
    const { workbook } = await cycle({
      sheets: [
        sheetWith({
          sparklines: [{ type: "column", location: "D1", dataRange: "S!A1:C1", color: "FF0000" }],
        }),
      ],
    })
    const sparklines = workbook.sheets[0].sparklines
    expect(sparklines).toHaveLength(1)
    expect(sparklines![0]).toMatchObject({
      type: "column",
      location: "D1",
      dataRange: "S!A1:C1",
      color: "FF0000",
    })
  })

  it("keeps text boxes", async () => {
    const { workbook } = await cycle({
      sheets: [
        sheetWith({
          textBoxes: [
            { text: "Note", anchor: { from: { col: 3, row: 1 }, to: { col: 6, row: 4 } } },
          ],
        }),
      ],
    })
    expect(workbook.sheets[0].textBoxes).toHaveLength(1)
    expect(workbook.sheets[0].textBoxes![0].text).toBe("Note")
  })

  it("keeps several of them at once", async () => {
    const { workbook } = await cycle({
      sheets: [
        sheetWith({
          rowBreaks: [4],
          colBreaks: [2],
          sparklines: [{ type: "column", location: "D2", dataRange: "S!A2:C2" }],
          splitPane: { xSplit: 1000, ySplit: 1000 },
        }),
      ],
    })
    const sheet = workbook.sheets[0]
    expect(sheet.rowBreaks).toEqual([4])
    expect(sheet.colBreaks).toEqual([2])
    expect(sheet.sparklines).toHaveLength(1)
    expect(sheet.splitPane).toBeDefined()
  })
})

describe("openXlsx → saveXlsx preserves workbook-level state", () => {
  it("keeps workbook protection", async () => {
    // The security-relevant one: a structurally locked workbook used to
    // come back unlocked, with nothing in the output to say so.
    const { workbook } = await cycle({
      sheets: [sheetWith({})],
      workbookProtection: { lockStructure: true, lockWindows: true },
    })
    expect(workbook.workbookProtection).toMatchObject({
      lockStructure: true,
      lockWindows: true,
    })
  })

  it("keeps a structure lock that carries a password hash", async () => {
    const { workbook } = await cycle({
      sheets: [sheetWith({})],
      workbookProtection: { lockStructure: true, password: "secret" },
    })
    expect(workbook.workbookProtection?.lockStructure).toBe(true)
  })

  it("does not invent protection where there was none", async () => {
    const { workbook } = await cycle({ sheets: [sheetWith({})] })
    expect(workbook.workbookProtection).toBeUndefined()
  })

  it("keeps Excel 2024 checkboxes", async () => {
    const { saved, workbook } = await cycle({
      sheets: [
        {
          name: "S",
          rows: [["flag"]],
          cells: new Map([["1,0", { value: true, type: "boolean", checkbox: true }]]),
        },
      ],
    })

    // The cell keeps its checkbox flag...
    expect(workbook.sheets[0].cells?.get("1,0")?.checkbox).toBe(true)

    // ...and the part it depends on is actually in the archive, not just
    // declared. A dangling declaration is what Excel calls corrupt.
    const zip = new ZipReader(saved)
    expect(zip.has("xl/featurePropertyBag/featurePropertyBag.xml")).toBe(true)
  })

  it("does not emit a featurePropertyBag part when nothing uses one", async () => {
    const { saved } = await cycle({ sheets: [sheetWith({})] })
    const zip = new ZipReader(saved)
    expect(zip.has("xl/featurePropertyBag/featurePropertyBag.xml")).toBe(false)
  })
})

describe("saveXlsx output integrity", () => {
  it("ships every part its content types declare", async () => {
    const { saved } = await cycle({
      sheets: [
        sheetWith({
          rowBreaks: [3],
          sparklines: [{ type: "line", location: "D1", dataRange: "S!A1:C1" }],
          textBoxes: [{ text: "x", anchor: { from: { col: 3, row: 1 }, to: { col: 5, row: 3 } } }],
          cells: new Map([["2,0", { value: false, type: "boolean", checkbox: true }]]),
        }),
      ],
      workbookProtection: { lockStructure: true },
    })

    const zip = new ZipReader(saved)
    const contentTypes = new TextDecoder().decode(await zip.extract("[Content_Types].xml"))

    for (const match of contentTypes.matchAll(/PartName="\/([^"]+)"/g)) {
      expect(zip.has(match[1]!), `missing part ${match[1]}`).toBe(true)
    }
  })
})

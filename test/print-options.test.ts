import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { ZipReader } from "../src/zip/reader"
import type { PageSetup, WriteOptions } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

async function roundTrip(pageSetup: PageSetup): Promise<PageSetup | undefined> {
  const options: WriteOptions = { sheets: [{ name: "S", rows: [["a"]], pageSetup }] }
  const workbook = await readXlsx(await writeXlsx(options))
  return workbook.sheets[0].pageSetup
}

async function sheetXml(pageSetup: PageSetup): Promise<string> {
  const buf = await writeXlsx({ sheets: [{ name: "S", rows: [["a"]], pageSetup }] })
  const zip = new ZipReader(buf)
  return new TextDecoder().decode(await zip.extract("xl/worksheets/sheet1.xml"))
}

// ═══════════════════════════════════════════════════════════════════════
// #360 — <printOptions> was written unconditionally with hardcoded
// zeros whenever any pageSetup existed, so setting a page margin turned
// printed gridlines and headings off. The two type fields backing it
// were never parsed, making them write-only and inert.
// ═══════════════════════════════════════════════════════════════════════

describe("print gridlines and headings", () => {
  it("survives a write → read cycle", async () => {
    const result = await roundTrip({ showGridLines: true, showRowColHeaders: true })
    expect(result?.showGridLines).toBe(true)
    expect(result?.showRowColHeaders).toBe(true)
  })

  it("is not turned off by an unrelated page setting", async () => {
    // The original bug in one line: ask for landscape, lose your
    // gridlines.
    const xml = await sheetXml({ orientation: "landscape" })
    expect(xml).not.toContain('gridLines="0"')
    expect(xml).not.toContain('headings="0"')
  })

  it("emits nothing when both are at their OOXML default", async () => {
    const xml = await sheetXml({ orientation: "landscape", scale: 80 })
    expect(xml).not.toContain("<printOptions")
  })

  it("emits only the attributes that were set", async () => {
    const xml = await sheetXml({ showGridLines: true })
    expect(xml).toContain('gridLines="1"')
    expect(xml).not.toContain("headings=")
  })

  it("reads a file that has printOptions before pageSetup", async () => {
    // Both orders are valid in a worksheet, so the reader merges rather
    // than assigns.
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"]],
          pageSetup: { showGridLines: true, orientation: "landscape" },
        },
      ],
    })
    const workbook = await readXlsx(buf)
    expect(workbook.sheets[0].pageSetup).toMatchObject({
      showGridLines: true,
      orientation: "landscape",
    })
  })
})

describe("print centering", () => {
  it("is written on printOptions, where ECMA-376 puts it", async () => {
    // hucre used to read and write these on <pageSetup>. That
    // round-tripped through hucre only because it was consistently wrong
    // in both directions — Excel ignored them entirely.
    const xml = await sheetXml({ horizontalCentered: true, verticalCentered: true })

    const printOptions = xml.slice(xml.indexOf("<printOptions"))
    const printOptionsTag = printOptions.slice(0, printOptions.indexOf(">") + 1)
    expect(printOptionsTag).toContain('horizontalCentered="1"')
    expect(printOptionsTag).toContain('verticalCentered="1"')

    const pageSetupTag = xml.slice(xml.indexOf("<pageSetup"))
    expect(pageSetupTag.slice(0, pageSetupTag.indexOf(">") + 1)).not.toContain("Centered")
  })

  it("survives a write → read cycle", async () => {
    const result = await roundTrip({ horizontalCentered: true, verticalCentered: true })
    expect(result?.horizontalCentered).toBe(true)
    expect(result?.verticalCentered).toBe(true)
  })

  it("still accepts the attributes from pageSetup, for files hucre wrote before", async () => {
    const legacy = (await sheetXml({ orientation: "landscape" })).replace(
      "<pageSetup ",
      '<pageSetup horizontalCentered="1" verticalCentered="1" ',
    )

    // Rebuild the workbook around the patched sheet.
    const original = await writeXlsx({
      sheets: [{ name: "S", rows: [["a"]], pageSetup: { orientation: "landscape" } }],
    })
    const zip = new ZipReader(original)
    const { ZipWriter } = await import("../src/zip/writer")
    const rebuilt = new ZipWriter()
    for (const path of zip.entries()) {
      rebuilt.add(
        path,
        path === "xl/worksheets/sheet1.xml"
          ? new TextEncoder().encode(legacy)
          : await zip.extract(path),
      )
    }

    const workbook = await readXlsx(await rebuilt.build())
    expect(workbook.sheets[0].pageSetup?.horizontalCentered).toBe(true)
    expect(workbook.sheets[0].pageSetup?.verticalCentered).toBe(true)
  })
})

describe("openXlsx → saveXlsx", () => {
  it("preserves print options rather than resetting them", async () => {
    const original = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"]],
          pageSetup: {
            orientation: "landscape",
            showGridLines: true,
            showRowColHeaders: true,
            horizontalCentered: true,
          },
        },
      ],
    })

    const saved = await saveXlsx(await openXlsx(original))
    const workbook = await readXlsx(saved)

    expect(workbook.sheets[0].pageSetup).toMatchObject({
      showGridLines: true,
      showRowColHeaders: true,
      horizontalCentered: true,
    })
  })
})

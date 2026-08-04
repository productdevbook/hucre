import { describe, expect, it } from "vitest"
import { readXml } from "../src/xml/data-reader"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import type { CellValue } from "../src/_types"

const FEED = `<?xml version="1.0"?><products>
  <product><sku>A1</sku><qty>2</qty></product>
  <product><sku>B2</sku><qty>5</qty></product>
</products>`

// ═══════════════════════════════════════════════════════════════════════
// #384 — readXml's transformValue was the only one in the library
// missing colIndex, and TypeScript accepts a callback that takes fewer
// arguments, so nothing flagged a four-argument function silently
// losing its last parameter here.
// ═══════════════════════════════════════════════════════════════════════

describe("readXml transformValue", () => {
  it("passes the column index", () => {
    const seen: Array<[string, number, number]> = []
    readXml(FEED, {
      transformValue: (value, header, rowIndex, colIndex) => {
        seen.push([header, rowIndex, colIndex])
        return value
      },
    })

    expect(seen).toEqual([
      ["sku", 0, 0],
      ["qty", 0, 1],
      ["sku", 1, 0],
      ["qty", 1, 1],
    ])
  })

  it("matches the signature the other readers use", () => {
    // A callback written for parseCsv / readXlsxObjects must work here
    // unchanged — that is the whole point of the fix.
    const byColumn = (value: CellValue, _h: string, _r: number, colIndex: number): CellValue =>
      colIndex === 1 ? Number(value) * 10 : value

    const { data } = readXml(FEED, { transformValue: byColumn })
    expect(data).toEqual([
      { sku: "A1", qty: 20 },
      { sku: "B2", qty: 50 },
    ])
  })

  it("still transforms without using the new argument", () => {
    const { data } = readXml(FEED, {
      transformValue: (value) => (typeof value === "string" ? value.toLowerCase() : value),
    })
    expect(data[0]).toMatchObject({ sku: "a1" })
  })
})

describe("readXml input", () => {
  it("accepts a Uint8Array, like parseJson", () => {
    const bytes = new TextEncoder().encode(FEED)
    const { data, headers } = readXml(bytes)

    expect(headers).toEqual(["sku", "qty"])
    expect(data).toHaveLength(2)
  })

  it("returns the same result for bytes and string", () => {
    const fromString = readXml(FEED)
    const fromBytes = readXml(new TextEncoder().encode(FEED))
    expect(fromBytes).toEqual(fromString)
  })

  it("handles empty bytes", () => {
    expect(readXml(new Uint8Array()).data).toEqual([])
  })

  it("decodes multi-byte characters correctly", () => {
    // A naive byte-per-char decode would mangle these.
    const xml = `<rows><row><city>İstanbul</city></row></rows>`
    const { data } = readXml(new TextEncoder().encode(xml))
    expect(data[0]).toEqual({ city: "İstanbul" })
  })
})

// ═══════════════════════════════════════════════════════════════════════
// #359 — outlineProperties was write-only: the type and the writer
// existed, nothing ever parsed <outlinePr>, so the field was always
// undefined after a read and open → save could not preserve it.
// ═══════════════════════════════════════════════════════════════════════

describe("outlineProperties", () => {
  it("survives a write → read cycle", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"]],
          outlineProperties: { summaryBelow: false, summaryRight: false },
        },
      ],
    })

    const workbook = await readXlsx(buf)
    expect(workbook.sheets[0].outlineProperties).toEqual({
      summaryBelow: false,
      summaryRight: false,
    })
  })

  it("reads the true case too", async () => {
    const buf = await writeXlsx({
      sheets: [
        { name: "S", rows: [["a"]], outlineProperties: { summaryBelow: true, summaryRight: true } },
      ],
    })
    expect((await readXlsx(buf)).sheets[0].outlineProperties).toEqual({
      summaryBelow: true,
      summaryRight: true,
    })
  })

  it("stays undefined when the sheet has no outlinePr", async () => {
    const buf = await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] })
    expect((await readXlsx(buf)).sheets[0].outlineProperties).toBeUndefined()
  })

  it("survives openXlsx → saveXlsx", async () => {
    // The field was already in the roundtrip map (#366), but forwarding
    // an always-undefined value did nothing. It carries a value now.
    const original = await writeXlsx({
      sheets: [{ name: "S", rows: [["a"]], outlineProperties: { summaryBelow: false } }],
    })

    const saved = await saveXlsx(await openXlsx(original))
    expect((await readXlsx(saved)).sheets[0].outlineProperties).toEqual({ summaryBelow: false })
  })
})

import { describe, expect, it } from "vitest"
import { XlsxStreamWriter } from "../src/xlsx/stream-writer"
import { CsvStreamWriter } from "../src/csv/stream"
import { NdjsonStreamWriter } from "../src/json/stream"
import { OdsStreamWriter } from "../src/ods/incremental-writer"
import { readOds } from "../src/ods/reader"
import { readXlsx } from "../src/xlsx/reader"
import { parseCsv } from "../src/csv/reader"
import type { CellValue, SpreadsheetStreamWriter } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// The four incremental writers share one interface, and this is the test
// that a helper written against the interface — no instanceof, no
// per-format branch — produces real output from every one of them.
//
// v1's interface said `finish(): string | Promise<Uint8Array>` and carried
// a `toStream()` that buffered on three of the four. v2 makes `finish()`
// bytes everywhere; the text writers add `finishText()`.
// ═══════════════════════════════════════════════════════════════════════

const KEYS = ["name", "qty"]
const COLUMN_DEFS = [
  { header: "Name", key: "name" },
  { header: "Qty", key: "qty" },
]

const ROWS: Array<Record<string, CellValue>> = [
  { name: "Widget", qty: 3 },
  { name: "Gadget", qty: 7 },
]

async function exportAll(writer: SpreadsheetStreamWriter): Promise<Uint8Array> {
  for (const row of ROWS) writer.addObject(row)
  return await writer.finish()
}

function writers(): Array<[string, SpreadsheetStreamWriter]> {
  return [
    ["XlsxStreamWriter", new XlsxStreamWriter({ name: "Sheet1", columns: COLUMN_DEFS })],
    ["CsvStreamWriter", new CsvStreamWriter({ columns: KEYS, headers: ["Name", "Qty"] })],
    ["NdjsonStreamWriter", new NdjsonStreamWriter({ columns: KEYS })],
    [
      "OdsStreamWriter",
      new OdsStreamWriter({
        name: "S",
        columns: [
          { header: "Name", key: "name" },
          { header: "Qty", key: "qty" },
        ],
      }),
    ],
  ]
}

describe("the four writers satisfy one interface", () => {
  it("every one of them is assignable to it", () => {
    for (const [name, writer] of writers()) {
      expect(typeof writer.addRow, name).toBe("function")
      expect(typeof writer.addObject, name).toBe("function")
      expect(typeof writer.finish, name).toBe("function")
    }
  })

  it("the one helper produces real output from each, as bytes", async () => {
    const out = new Map<string, Uint8Array>()
    for (const [name, writer] of writers()) out.set(name, await exportAll(writer))
    for (const [name, bytes] of out) expect(bytes, name).toBeInstanceOf(Uint8Array)

    const wb = await readXlsx(out.get("XlsxStreamWriter")!)
    expect(wb.sheets[0]!.rows).toEqual([
      ["Name", "Qty"],
      ["Widget", 3],
      ["Gadget", 7],
    ])

    expect(parseCsv(out.get("CsvStreamWriter")!)).toEqual([
      ["Name", "Qty"],
      ["Widget", "3"],
      ["Gadget", "7"],
    ])

    expect((await readOds(out.get("OdsStreamWriter")!)).sheets[0]!.rows).toEqual([
      ["Name", "Qty"],
      ["Widget", 3],
      ["Gadget", 7],
    ])

    expect(
      new TextDecoder()
        .decode(out.get("NdjsonStreamWriter")!)
        .trim()
        .split("\n")
        .map((l) => JSON.parse(l)),
    ).toEqual(ROWS)
  })

  it("a cell object is accepted by every writer, and reduces to its value where styles cannot go", async () => {
    const styled = { value: "Widget", style: { font: { bold: true } } }
    const csv = new CsvStreamWriter({ headers: false })
    csv.addRow([styled, 3])
    expect(csv.finishText()).toBe("Widget,3")

    const ndjson = new NdjsonStreamWriter({ columns: KEYS })
    ndjson.addRow([styled, 3])
    expect(ndjson.finishText()).toBe('{"name":"Widget","qty":3}\n')

    const xlsx = new XlsxStreamWriter({ name: "S" })
    xlsx.addRow([styled, 3])
    const wb = await readXlsx(await xlsx.finish(), { readStyles: true })
    expect(wb.sheets[0]!.cells?.get("0,0")?.style?.font?.bold).toBe(true)
  })
})

describe("the text writers also give their output as a string", () => {
  it("finishText() is finish() decoded", async () => {
    const a = new CsvStreamWriter({ columns: KEYS, headers: ["Name", "Qty"] })
    const b = new CsvStreamWriter({ columns: KEYS, headers: ["Name", "Qty"] })
    for (const row of ROWS) {
      a.addObject(row)
      b.addObject(row)
    }
    expect(new TextDecoder().decode(await a.finish())).toBe(b.finishText())
  })
})

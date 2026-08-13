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
// #468 — the README has promised since before v1 that the three
// incremental writers "share one vocabulary, so a format-agnostic export
// helper can be written once". There was no `implements` anywhere in
// `src/` and no shared interface, so nothing held them to it — and they
// had already drifted: #436 widened `XlsxStreamWriter.addRow` to accept
// `StreamStyledCell` and nothing failed.
//
// The interface is the enforcement. This file is the claim: the helper
// the README describes, actually written once, actually run against all
// three.
// ═══════════════════════════════════════════════════════════════════════

// Construction is NOT part of the shared vocabulary and this is where
// that shows: XLSX wants `ColumnDef[]` and a sheet name, the two text
// writers want a plain key list. The interface covers the four methods,
// which is what "the helper is written once" actually means — each
// writer is still built its own way. `tsc` catches a mix-up here, which
// is how this comment came to exist.
const KEYS = ["name", "qty"]
const COLUMN_DEFS = [
  { header: "Name", key: "name" },
  { header: "Qty", key: "qty" },
]

const ROWS: Array<Record<string, CellValue>> = [
  { name: "Widget", qty: 3 },
  { name: "Gadget", qty: 7 },
]

/**
 * The helper. Written against the interface and nothing else — no
 * `instanceof`, no per-format branch.
 *
 * `finish()` is the one place the three genuinely differ, and the
 * interface says so: `string | Promise<Uint8Array>`. `await` covers both,
 * and the caller narrows.
 */
async function exportAll(writer: SpreadsheetStreamWriter): Promise<string | Uint8Array> {
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

describe("the three writers satisfy one interface", () => {
  it("every one of them is assignable to it", () => {
    // The assignment itself is the assertion — this file fails `tsc`
    // before it fails vitest if any of the three drifts.
    for (const [name, writer] of writers()) {
      expect(typeof writer.addRow, name).toBe("function")
      expect(typeof writer.addObject, name).toBe("function")
      expect(typeof writer.finish, name).toBe("function")
      expect(typeof writer.toStream, name).toBe("function")
    }
  })

  it("the one helper produces real output from each", async () => {
    const out = new Map<string, string | Uint8Array>()
    for (const [name, writer] of writers()) out.set(name, await exportAll(writer))

    // XLSX: bytes, and a workbook that reads back.
    const xlsxBytes = out.get("XlsxStreamWriter")
    expect(xlsxBytes).toBeInstanceOf(Uint8Array)
    const wb = await readXlsx(xlsxBytes as Uint8Array)
    expect(wb.sheets[0]!.rows).toEqual([
      ["Name", "Qty"],
      ["Widget", 3],
      ["Gadget", 7],
    ])

    // CSV: text, and rows that parse back.
    const csv = out.get("CsvStreamWriter")
    expect(typeof csv).toBe("string")
    expect(parseCsv(csv as string)).toEqual([
      ["Name", "Qty"],
      ["Widget", "3"],
      ["Gadget", "7"],
    ])

    // ODS: bytes, and a document that reads back.
    const odsBytes = out.get("OdsStreamWriter")
    expect(odsBytes).toBeInstanceOf(Uint8Array)
    expect((await readOds(odsBytes as Uint8Array)).sheets[0]!.rows).toEqual([
      ["Name", "Qty"],
      ["Widget", 3],
      ["Gadget", 7],
    ])

    // NDJSON: one object per line.
    const ndjson = out.get("NdjsonStreamWriter")
    expect(typeof ndjson).toBe("string")
    expect(
      (ndjson as string)
        .trim()
        .split("\n")
        .map((l) => JSON.parse(l)),
    ).toEqual(ROWS)
  })

  it("toStream is the same shape on all three", async () => {
    for (const [name, writer] of writers()) {
      writer.addRow(["a", 1])
      const stream = writer.toStream()
      expect(stream, name).toBeInstanceOf(ReadableStream)

      const reader = stream.getReader()
      const first = await reader.read()
      expect(first.done, name).toBe(false)
      expect(first.value, name).toBeInstanceOf(Uint8Array)
      await reader.cancel()
    }
  })
})

describe("what the interface deliberately does not promise", () => {
  it("XlsxStreamWriter still takes more than the interface asks for", async () => {
    // Contravariance: a writer may accept more than the interface
    // promises, never less. The interface narrowing to `CellValue[]` is
    // what makes the helper portable; it does not take styling away.
    const w = new XlsxStreamWriter({ name: "Sheet1", columns: COLUMN_DEFS })
    w.addRow(["Widget", { value: 3, style: { font: { bold: true } } }])

    const wb = await readXlsx(await w.finish(), { readStyles: true })
    expect(wb.sheets[0]!.rows[1]).toEqual(["Widget", 3])
    expect(wb.sheets[0]!.cells?.get("1,1")?.style?.font?.bold).toBe(true)
  })

  it("finish() is not one type, and the interface says so rather than lying", async () => {
    // Converging these is a real API decision and a breaking one. What
    // the interface buys today is that a caller who `await`s and narrows
    // is correct for all three, forever.
    const xlsx = await new XlsxStreamWriter({ name: "Sheet1", columns: COLUMN_DEFS }).finish()
    const csv = new CsvStreamWriter({ columns: KEYS, headers: ["Name", "Qty"] }).finish()

    expect(xlsx).toBeInstanceOf(Uint8Array)
    expect(typeof csv).toBe("string")
  })
})

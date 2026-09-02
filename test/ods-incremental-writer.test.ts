import { describe, expect, it } from "vitest"
import { OdsStreamWriter } from "../src/ods/incremental-writer"
import { writeOdsStream } from "../src/ods/stream-writer"
import { readOds } from "../src/ods/reader"
import { ZipReader } from "../src/zip/reader"
import type { CellValue, SpreadsheetStreamWriter } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #467's last empty cell. `writeOdsStream` covers the constant-memory
// case and carries values only, because ODF puts
// `<office:automatic-styles>` before the body — a style first seen on row
// 900,000 has nowhere to be declared once the body has gone out.
//
// A buffering writer does not have that problem: it holds the serialized
// rows until `finish()`, so the style block is written from everything it
// saw. That is the trade, and it is the same one XLSX already offers —
// `writeXlsxStream` for constant memory, `XlsxStreamWriter` for a buffer
// you can style.
// ═══════════════════════════════════════════════════════════════════════

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
  }
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0))
  let at = 0
  for (const c of chunks) {
    out.set(c, at)
    at += c.length
  }
  return out
}

describe("it writes a document ODS readers accept", () => {
  it("values of every type, in order", async () => {
    const w = new OdsStreamWriter({ name: "Report" })
    w.addRow(["Widget", 3, true, new Date("2024-03-17T00:00:00Z")])
    w.addRow(["Gadget", -7.5, false, null])

    const wb = await readOds(await w.finish())
    const sheet = wb.sheets[0]!

    expect(sheet.name).toBe("Report")
    expect(sheet.rows[0]![0]).toBe("Widget")
    expect(sheet.rows[0]![1]).toBe(3)
    expect(sheet.rows[0]![2]).toBe(true)
    expect(sheet.rows[0]![3]).toBeInstanceOf(Date)
    expect(sheet.rows[1]![1]).toBe(-7.5)
  })

  it("a valid ODF package, mimetype first and stored", async () => {
    const w = new OdsStreamWriter()
    w.addRow(["a"])
    const bytes = await w.finish()

    const entries = new ZipReader(bytes).entries()
    expect(entries[0]).toBe("mimetype")
    for (const part of ["META-INF/manifest.xml", "content.xml", "styles.xml", "meta.xml"]) {
      expect(entries, part).toContain(part)
    }
  })
})

describe("the thing writeOdsStream cannot do", () => {
  const styled = async (): Promise<Uint8Array> => {
    const w = new OdsStreamWriter({ name: "S" })
    w.addRow([
      { value: "bold", style: { font: { bold: true } } },
      { value: 1234.5, style: { numFmt: "#,##0.00" } },
      {
        value: "filled",
        style: { fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } } },
      },
    ])
    return w.finish()
  }

  it("carries per-cell styles", async () => {
    const sheet = (await readOds(await styled(), { readStyles: true })).sheets[0]!

    expect(sheet.cells?.get("0,0")?.style?.font?.bold).toBe(true)
    expect(sheet.cells?.get("0,1")?.style?.numFmt).toBe("#,##0.00")

    const fill = sheet.cells?.get("0,2")?.style?.fill
    expect(fill?.type === "pattern" ? fill.fgColor?.rgb : undefined).toBe("FFFF00")
  })

  it("which the streaming writer does not, and that is the trade", async () => {
    // Not a defect in `writeOdsStream` — the style block has already been
    // written by the time row 1 is serialized. This test exists so the
    // difference between the two is pinned rather than folklore.
    const streamed = await drain(writeOdsStream([["bold"]], { name: "S" }))
    const sheet = (await readOds(streamed, { readStyles: true })).sheets[0]!

    expect(sheet.rows[0]![0]).toBe("bold")
    expect(sheet.cells?.get("0,0")?.style?.font?.bold).toBeUndefined()
  })

  it("and column widths, which it also cannot", async () => {
    const w = new OdsStreamWriter({ name: "S", columns: [{ width: 30 }, {}] })
    w.addRow(["a", "b"])
    const content = new TextDecoder().decode(
      await new ZipReader(await w.finish()).extract("content.xml"),
    )

    expect(content).toContain("style:column-width")
  })
})

describe("the shared vocabulary", () => {
  it("satisfies SpreadsheetStreamWriter", async () => {
    // The assignment is the assertion — this file fails `tsc` before it
    // fails vitest if the class drifts from the other three.
    const writer: SpreadsheetStreamWriter = new OdsStreamWriter({
      name: "S",
      columns: [{ header: "Name", key: "name" }],
    })

    writer.addObject({ name: "Widget" })
    const bytes = await writer.finish()

    expect(bytes).toBeInstanceOf(Uint8Array)
  })

  it("writes a header row from columns, like the others do", async () => {
    const w = new OdsStreamWriter({
      name: "S",
      columns: [
        { header: "Name", key: "name" },
        { header: "Qty", key: "qty" },
      ],
    })
    w.addObject({ name: "Widget", qty: 3 })

    const rows = (await readOds(await w.finish())).sheets[0]!.rows
    expect(rows[0]).toEqual(["Name", "Qty"])
    expect(rows[1]).toEqual(["Widget", 3])
  })

  it("addObject needs keys, and says so rather than guessing", async () => {
    const w = new OdsStreamWriter({ name: "S" })

    expect(() => w.addObject({ a: 1 })).toThrow(/columns/)
  })

  it("finish hands back the finished document", async () => {
    const w = new OdsStreamWriter({ name: "S" })
    w.addRow(["a", 1])

    const wb = await readOds(await w.finish())
    expect(wb.sheets[0]!.rows[0]).toEqual(["a", 1])
  })

  it("refuses writes after finish", async () => {
    const w = new OdsStreamWriter({ name: "S" })
    w.addRow(["a"])
    await w.finish()

    expect(() => w.addRow(["b"])).toThrow(/after finish/)
  })
})

describe("details that are easy to get wrong", () => {
  it("a cell's own style beats its column's", async () => {
    const w = new OdsStreamWriter({
      name: "S",
      columns: [{ style: { font: { italic: true } } }],
    })
    w.addRow([{ value: "x", style: { font: { bold: true } } }])

    const style = (await readOds(await w.finish(), { readStyles: true })).sheets[0]!.cells?.get(
      "0,0",
    )?.style

    expect(style?.font?.bold).toBe(true)
    expect(style?.font?.italic).toBeUndefined()
  })

  it("a column style applies where the cell has none", async () => {
    const w = new OdsStreamWriter({ name: "S", columns: [{ style: { font: { italic: true } } }] })
    w.addRow(["x"])

    const style = (await readOds(await w.finish(), { readStyles: true })).sheets[0]!.cells?.get(
      "0,0",
    )?.style

    expect(style?.font?.italic).toBe(true)
  })

  it("carries a formula", async () => {
    const w = new OdsStreamWriter({ name: "S" })
    w.addRow([1, 2, { formula: "SUM(A1:B1)" }])

    const content = new TextDecoder().decode(
      await new ZipReader(await w.finish()).extract("content.xml"),
    )
    expect(content).toContain("table:formula")
  })

  it("reuses one style record for a repeated style", async () => {
    // The collector is shared across rows, which is what makes this a
    // writer rather than a document builder.
    const style = { font: { bold: true } }
    const w = new OdsStreamWriter({ name: "S" })
    for (let i = 0; i < 50; i++) w.addRow([{ value: i, style }])

    const content = new TextDecoder().decode(
      await new ZipReader(await w.finish()).extract("content.xml"),
    )
    expect((content.match(/style:family="table-cell"/g) ?? []).length).toBe(1)
  })

  it("escapes what would otherwise break the XML", async () => {
    const w = new OdsStreamWriter({ name: "A&B" })
    w.addRow(['<a & "b">'])

    const wb = await readOds(await w.finish())
    expect(wb.sheets[0]!.name).toBe("A&B")
    expect(wb.sheets[0]!.rows[0]![0]).toBe('<a & "b">')
  })

  it("still refuses a sheet name Excel would", () => {
    expect(() => new OdsStreamWriter({ name: "a/b" })).toThrow()
  })

  it("takes rows of differing length", async () => {
    const w = new OdsStreamWriter({ name: "S" })
    w.addRow(["a"])
    w.addRow(["b", "c", "d"])

    const rows = (await readOds(await w.finish())).sheets[0]!.rows
    expect(rows[0]![0]).toBe("a")
    expect(rows[1]).toEqual(["b", "c", "d"])
  })
})

describe("it agrees with writeOds on values", () => {
  it("the same rows produce the same values", async () => {
    const rows: CellValue[][] = [
      ["name", "qty"],
      ["Widget", 3],
      ["Gadget", -7.5],
    ]

    const { writeOds } = await import("../src/ods/writer")
    const buffered = (await readOds(await writeOds({ sheets: [{ name: "S", rows }] }))).sheets[0]!

    const w = new OdsStreamWriter({ name: "S" })
    for (const row of rows) w.addRow(row)
    const incremental = (await readOds(await w.finish())).sheets[0]!

    expect(incremental.rows).toEqual(buffered.rows)
  })
})

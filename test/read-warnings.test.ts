import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import type { ReadWarning } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #439 §S — the readers are lenient on purpose, and that is the right
// default for a format you receive rather than control. But leniency was
// the *only* mode: a cell pointing at a shared string that is not there
// came back as `null`, indistinguishable from a cell that was genuinely
// empty, and nothing said which.
//
// `onWarning` is a side channel, not part of the document — which is why
// it is a callback rather than a field on `Workbook`.
// ═══════════════════════════════════════════════════════════════════════

/** Rebuild an archive with one part rewritten, standing in for a damaged file. */
async function damage(
  bytes: Uint8Array,
  path: string,
  edit: (xml: string) => string,
): Promise<Uint8Array> {
  const all = await new ZipReader(bytes).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    zw.add(
      name,
      name === path ? new TextEncoder().encode(edit(new TextDecoder().decode(data))) : data,
    )
  }
  return zw.build()
}

const SHEET = "xl/worksheets/sheet1.xml"

async function baseline(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      { name: "Data", rows: [["hello", 1]], columns: [{ style: { font: { bold: true } } }] },
    ],
  })
}

describe("a dropped shared string is reported", () => {
  it("names the cell, the index, and what the file actually has", async () => {
    const bytes = await damage(await baseline(), SHEET, (xml) =>
      xml.replace("<v>0</v>", "<v>9999</v>"),
    )

    const warnings: ReadWarning[] = []
    const wb = await readXlsx(bytes, { onWarning: (w) => warnings.push(w) })

    // Still lenient: the value is null, not an exception.
    expect(wb.sheets[0]!.rows[0]![0]).toBeNull()

    expect(warnings).toHaveLength(1)
    expect(warnings[0]!.code).toBe("unresolved-shared-string")
    expect(warnings[0]!.sheet).toBe("Data")
    expect(warnings[0]!.row).toBe(0)
    expect(warnings[0]!.col).toBe(0)
    expect(warnings[0]!.message).toContain("9999")
    expect(warnings[0]!.message).toContain("A1")
  })
})

describe("a dropped cell format is reported", () => {
  it("says which format the cell asked for", async () => {
    const bytes = await damage(await baseline(), SHEET, (xml) =>
      xml.replace(/ s="\d+"/g, ' s="9999"'),
    )

    const warnings: ReadWarning[] = []
    const wb = await readXlsx(bytes, { readStyles: true, onWarning: (w) => warnings.push(w) })

    expect(wb.sheets[0]!.cells?.get("0,0")?.style).toBeUndefined()
    expect(warnings.some((w) => w.code === "unresolved-style")).toBe(true)
    expect(warnings.find((w) => w.code === "unresolved-style")!.message).toContain("9999")
  })

  it("says nothing when styles are not being read", async () => {
    const bytes = await damage(await baseline(), SHEET, (xml) =>
      xml.replace(/ s="\d+"/g, ' s="9999"'),
    )

    const warnings: ReadWarning[] = []
    await readXlsx(bytes, { onWarning: (w) => warnings.push(w) })

    expect(warnings.filter((w) => w.code === "unresolved-style")).toEqual([])
  })

  it("does not report a format that simply resolves to nothing", async () => {
    // xf 0 is the default and carries no formatting. That is not damage.
    const warnings: ReadWarning[] = []
    await readXlsx(await baseline(), { readStyles: true, onWarning: (w) => warnings.push(w) })

    expect(warnings).toEqual([])
  })
})

describe("a clean file says nothing", () => {
  it("reports no warnings for a workbook hucre wrote", async () => {
    const warnings: ReadWarning[] = []

    await readXlsx(await baseline(), { readStyles: true, onWarning: (w) => warnings.push(w) })

    expect(warnings).toEqual([])
  })

  it("changes nothing when the option is omitted", async () => {
    const bytes = await damage(await baseline(), SHEET, (xml) =>
      xml.replace("<v>0</v>", "<v>9999</v>"),
    )

    const withSink = await readXlsx(bytes, { onWarning: () => {} })
    const without = await readXlsx(bytes)

    expect(without.sheets[0]!.rows).toEqual(withSink.sheets[0]!.rows)
  })
})

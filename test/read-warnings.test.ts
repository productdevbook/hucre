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

// ═══════════════════════════════════════════════════════════════════════
// #474 — `onWarning` was wired at the two sites measured as silent. Three
// more follow the same shape, each with a test rather than a speculative
// call. One of them, `unresolved-dxf`, was already in the `ReadWarning`
// union and emitted from nowhere.
// ═══════════════════════════════════════════════════════════════════════

describe("a conditional rule's formatting that resolves to nothing", () => {
  async function ruled(): Promise<Uint8Array> {
    return writeXlsx({
      sheets: [
        {
          name: "Data",
          rows: [[1], [2]],
          conditionalRules: [
            {
              type: "cellIs",
              priority: 1,
              range: "A1:A2",
              operator: "greaterThan",
              formula: "1",
              style: { fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } } },
            },
          ],
        },
      ],
    })
  }

  it("names the dxfId and says what the rule keeps", async () => {
    const bytes = await damage(await ruled(), SHEET, (xml) =>
      xml.replace(/dxfId="\d+"/, 'dxfId="42"'),
    )

    const warnings: ReadWarning[] = []
    const wb = await readXlsx(bytes, { readStyles: true, onWarning: (w) => warnings.push(w) })

    // Still lenient: the rule survives without its formatting, because a
    // rule that paints nothing is closer to the file than no rule at all.
    expect(wb.sheets[0]!.conditionalRules).toHaveLength(1)
    expect(wb.sheets[0]!.conditionalRules![0]!.style).toBeUndefined()

    const dxf = warnings.find((w) => w.code === "unresolved-dxf")
    expect(dxf).toBeDefined()
    expect(dxf!.message).toContain("42")
    expect(dxf!.message).toContain("A1:A2")
    expect(dxf!.sheet).toBe("Data")
  })

  it("says nothing when the rule's format is there", async () => {
    const warnings: ReadWarning[] = []
    await readXlsx(await ruled(), { readStyles: true, onWarning: (w) => warnings.push(w) })

    expect(warnings.filter((w) => w.code === "unresolved-dxf")).toEqual([])
  })
})

describe("a hyperlink pointing at a relationship that is not there", () => {
  async function linked(): Promise<Uint8Array> {
    return writeXlsx({
      sheets: [
        {
          name: "Data",
          rows: [["click"]],
          cells: new Map([
            ["0,0", { value: "click", hyperlink: { target: "https://example.com" } }],
          ]),
        },
      ],
    })
  }

  it("names the cell and the rId", async () => {
    // Drop the relationship, keep the reference — a real shape, since the
    // two live in different parts of the package.
    const bytes = await damage(await linked(), "xl/worksheets/_rels/sheet1.xml.rels", (xml) =>
      xml.replace(/<Relationship [^>]*hyperlink[^>]*\/>/, ""),
    )

    const warnings: ReadWarning[] = []
    const wb = await readXlsx(bytes, { onWarning: (w) => warnings.push(w) })

    // Lenient: an empty target, not an exception.
    expect(wb.sheets[0]!.cells?.get("0,0")?.hyperlink?.target).toBe("")

    const warning = warnings.find((w) => w.code === "unresolved-hyperlink")
    expect(warning).toBeDefined()
    expect(warning!.message).toContain("A1")
    expect(warning!.row).toBe(0)
    expect(warning!.col).toBe(0)
    expect(warning!.sheet).toBe("Data")
  })

  it("says nothing about a link that resolves", async () => {
    const warnings: ReadWarning[] = []
    await readXlsx(await linked(), { onWarning: (w) => warnings.push(w) })

    expect(warnings.filter((w) => w.code === "unresolved-hyperlink")).toEqual([])
  })
})

describe("a paper size that is not a usable code", () => {
  async function printed(): Promise<Uint8Array> {
    return writeXlsx({
      sheets: [{ name: "Data", rows: [["a"]], pageSetup: { paperSize: "a4" } }],
    })
  }

  it("names what the file said", async () => {
    const bytes = await damage(await printed(), SHEET, (xml) =>
      xml.replace(/paperSize="\d+"/, 'paperSize="0"'),
    )

    const warnings: ReadWarning[] = []
    const wb = await readXlsx(bytes, { onWarning: (w) => warnings.push(w) })

    expect(wb.sheets[0]!.pageSetup?.paperSize).toBeUndefined()

    const warning = warnings.find((w) => w.code === "unusable-paper-size")
    expect(warning).toBeDefined()
    expect(warning!.message).toContain('"0"')
    expect(warning!.sheet).toBe("Data")
  })

  it("says nothing about a code with no name, which is still a size", async () => {
    // 999 has no name in hucre's table; it round-trips as the number
    // rather than vanishing, so there is nothing to report.
    const bytes = await damage(await printed(), SHEET, (xml) =>
      xml.replace(/paperSize="\d+"/, 'paperSize="999"'),
    )

    const warnings: ReadWarning[] = []
    const wb = await readXlsx(bytes, { onWarning: (w) => warnings.push(w) })

    expect(wb.sheets[0]!.pageSetup?.paperSize).toBe(999)
    expect(warnings.filter((w) => w.code === "unusable-paper-size")).toEqual([])
  })
})

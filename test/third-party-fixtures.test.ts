import { describe, expect, it } from "vitest"
import { readFileSync, readdirSync } from "node:fs"
import { readXlsx } from "../src/xlsx/reader"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import { ZipReader } from "../src/zip/reader"
import type { CellValue } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #464 — the suite parsed 9,934 assertions' worth of hucre's own output
// and not one byte from anywhere else, so a writer bug the reader
// mirrored was invisible. Three of the defects fixed in the #439 round
// were exactly that shape, all sitting in `main` at 98.8% coverage.
//
// These fixtures are written by **ExcelJS** — an independent
// implementation with its own element ordering, its own defaults, and
// its own idea of what a minimal workbook contains. It is not Excel, and
// #464 asks for Excel and LibreOffice too; those still need someone with
// those tools. What this gives is the thing that was missing entirely:
// bytes hucre did not write.
//
// See test/fixtures/third-party/README.md for provenance and licensing.
// ═══════════════════════════════════════════════════════════════════════

const DIR = new URL("./fixtures/third-party/", import.meta.url)

function load(name: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(name, DIR)))
}

const NAMES = readdirSync(DIR)
  .filter((f: string) => f.endsWith(".xlsx"))
  .sort()

describe("the corpus is there and is not ours", () => {
  it("has every fixture the generator writes", () => {
    expect(NAMES).toEqual([
      "basic-values.xlsx",
      "conditional.xlsx",
      "errors-and-blanks.xlsx",
      "hyperlinks-and-comments.xlsx",
      "layout.xlsx",
      "multi-sheet.xlsx",
      "properties.xlsx",
      "styled.xlsx",
      "unicode.xlsx",
      "whitespace-strings.xlsx",
      "wide-and-tall.xlsx",
    ])
  })

  it("was written by something other than hucre", async () => {
    // The load-bearing assertion of this whole file. If someone ever
    // regenerates these with hucre, every test below becomes a round
    // trip against itself again and proves nothing — this is the line
    // that notices.
    const app = new TextDecoder().decode(
      await new ZipReader(load("basic-values.xlsx")).extract("docProps/app.xml"),
    )

    // hucre writes `<Application>hucre</Application>` and stops there.
    expect(app).not.toContain("hucre")
    // ExcelJS declares itself as Excel — it aims at Excel's output rather
    // than at a dialect of its own, which is a point in these fixtures'
    // favour. `HeadingPairs` and `AppVersion` are parts of that shape
    // hucre does not write at all.
    expect(app).toContain("<Application>Microsoft Excel</Application>")
    expect(app).toContain("HeadingPairs")
    expect(app).toContain("AppVersion")
  })

  it("carries package details hucre does not produce", async () => {
    // ZIP directory entries. hucre omits them — they are optional per the
    // spec and no reader looks at them — so their presence here is proof
    // these bytes came from elsewhere, and exercise for a reader that has
    // to skip them.
    const entries = new ZipReader(load("basic-values.xlsx")).entries()

    expect(entries.some((p) => p.endsWith("/"))).toBe(true)
  })
})

describe("every fixture reads", () => {
  for (const name of NAMES) {
    it(name, async () => {
      const wb = await readXlsx(load(name), { readStyles: true })

      expect(wb.sheets.length).toBeGreaterThan(0)
      expect(wb.sheets[0]!.rows.length).toBeGreaterThan(0)
    })
  }
})

// ── Golden models ────────────────────────────────────────────────────

describe("basic-values", () => {
  it("reads all five types, and the formula's cached result", async () => {
    const rows = (await readXlsx(load("basic-values.xlsx"))).sheets[0]!.rows

    expect(rows[0]).toEqual(["text", "number", "date", "boolean", "formula"])
    expect(rows[1]![0]).toBe("Widget")
    expect(rows[1]![1]).toBe(42)
    expect(rows[1]![2]).toBeInstanceOf(Date)
    expect((rows[1]![2] as Date).toISOString()).toBe("2024-01-15T00:00:00.000Z")
    expect(rows[1]![3]).toBe(true)
    // Formulas arrive as their cached value; hucre has no engine.
    expect(rows[1]![4]).toBe(84)
    expect(rows[2]![1]).toBe(-7.25)
    expect(rows[2]![3]).toBe(false)
  })
})

describe("whitespace-strings", () => {
  it("keeps every space another tool bothered to preserve", async () => {
    // The #441 shape, from the other side. hucre's reader does not trim,
    // so a hucre-written file agreed with itself whether or not the
    // writer emitted `xml:space`. Only a third tool can tell you.
    const rows = (await readXlsx(load("whitespace-strings.xlsx"))).sheets[0]!.rows

    expect(rows[0]![0]).toBe("  leading")
    expect(rows[1]![0]).toBe("trailing  ")
    expect(rows[2]![0]).toBe("  both  ")
    expect(rows[3]![0]).toBe("inner   gap")
    expect(rows[4]![0]).toBe("line\nbreak")
    expect(rows[5]![0]).toBe(" ")
  })
})

describe("styled", () => {
  it("reads fonts, fills, borders and number formats another tool wrote", async () => {
    const sheet = (await readXlsx(load("styled.xlsx"), { readStyles: true })).sheets[0]!
    const cell = (key: string) => sheet.cells?.get(key)?.style

    expect(cell("0,0")?.font?.bold).toBe(true)
    expect(cell("0,1")?.font).toMatchObject({ italic: true, size: 14, name: "Georgia" })
    const fill = cell("0,2")?.fill
    expect(fill).toMatchObject({ type: "pattern", pattern: "solid" })
    expect(fill?.type === "pattern" ? fill.fgColor?.rgb : undefined).toBe("FFFF00")
    expect(cell("0,3")?.border?.left).toMatchObject({ style: "medium" })
    expect(cell("0,3")?.border?.left?.color?.rgb).toBe("FF0000")
    expect(cell("0,3")?.border?.bottom?.style).toBe("double")
    expect(cell("1,4")?.numFmt).toBe("#,##0.00")
  })

  it("reads a column width another tool wrote", async () => {
    const sheet = (await readXlsx(load("styled.xlsx"), { readStyles: true })).sheets[0]!

    expect(sheet.columns?.[0]?.width).toBe(22)
  })
})

describe("layout", () => {
  it("reads merges, both freeze splits, and a column-level format", async () => {
    const sheet = (await readXlsx(load("layout.xlsx"), { readStyles: true })).sheets[0]!

    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }])
    expect(sheet.freezePane).toEqual({ columns: 1, rows: 2 })
    // A format on the column rather than the cell — the case #450 added.
    expect(sheet.columns?.[2]?.style?.numFmt).toBe('"$"#,##0.00')
  })
})

describe("conditional", () => {
  it("reads the rule and the differential format it points at", async () => {
    const sheet = (await readXlsx(load("conditional.xlsx"), { readStyles: true })).sheets[0]!
    const rule = sheet.conditionalRules?.[0]

    expect(rule).toMatchObject({
      type: "cellIs",
      operator: "greaterThan",
      range: "A2:A4",
      formula: "10",
    })
    expect(rule?.style?.fill).toBeDefined()
  })
})

describe("multi-sheet", () => {
  it("reads every sheet, in order, with its visibility", async () => {
    const wb = await readXlsx(load("multi-sheet.xlsx"))

    expect(wb.sheets.map((s) => s.name)).toEqual(["First", "Second", "Hidden"])
    expect(wb.sheets[0]!.hidden).toBeUndefined()
    expect(wb.sheets[2]!.hidden).toBe(true)
  })
})

describe("errors-and-blanks", () => {
  it("reads error values as their token, and gaps as null", async () => {
    const rows = (await readXlsx(load("errors-and-blanks.xlsx"))).sheets[0]!.rows

    expect(rows[1]![0]).toBe("#DIV/0!")
    expect(rows[1]![2]).toBe("#N/A")
    expect(rows[2]).toEqual([null, null, null])
    expect(rows[3]![0]).toBe("gap above")
  })
})

describe("hyperlinks-and-comments", () => {
  it("reads a link's target and a note's text and author", async () => {
    const sheet = (await readXlsx(load("hyperlinks-and-comments.xlsx"), { readStyles: true }))
      .sheets[0]!

    expect(sheet.cells?.get("1,0")?.hyperlink?.target).toBe("https://example.com")
    expect(sheet.cells?.get("2,0")?.comment?.text).toBe("a note from another tool")
  })
})

describe("wide-and-tall", () => {
  it("reads past the single-letter column boundary", async () => {
    // `AA`, `AB`, `AD` — references a narrower fixture never exercises.
    const rows = (await readXlsx(load("wide-and-tall.xlsx"))).sheets[0]!.rows

    expect(rows[0]).toHaveLength(30)
    expect(rows[0]![26]).toBe("c27")
    expect(rows[0]![29]).toBe("c30")
    expect(rows[5]![29]).toBe(4 * 30 + 29)
  })
})

describe("unicode", () => {
  it("reads every script, including astral-plane emoji", async () => {
    const rows = (await readXlsx(load("unicode.xlsx"))).sheets[0]!.rows

    expect(rows[0]).toEqual(["Türkçe", "şehir"])
    expect(rows[1]).toEqual(["日本語", "テスト"])
    expect(rows[2]).toEqual(["Ελληνικά", "δοκιμή"])
    // A surrogate pair, which a UTF-8 decoder that counts wrong will split.
    expect(rows[3]![1]).toBe("😀🎉")
    expect(rows[4]![1]).toBe("مرحبا")
  })
})

describe("properties", () => {
  it("reads the document properties another tool wrote", async () => {
    const props = (await readXlsx(load("properties.xlsx"))).properties

    expect(props?.creator).toBe("hucre fixture generator")
    expect(props?.title).toBe("Third-party fixture")
    expect(props?.subject).toBe("Testing")
    expect(props?.created?.toISOString()).toBe("2024-06-01T12:00:00.000Z")
  })
})

// ── The sharper tests ────────────────────────────────────────────────

describe("the streaming reader agrees with the buffering one", () => {
  for (const name of NAMES) {
    it(name, async () => {
      const buffered = (await readXlsx(load(name))).sheets[0]!.rows
      const streamed: CellValue[][] = []
      for await (const row of streamXlsxRows(load(name))) streamed.push(row.values)

      // The streaming reader skips entirely empty rows and keeps the true
      // index, so compare only the rows it produced.
      expect(streamed.length).toBeGreaterThan(0)
      expect(streamed[0]).toEqual(buffered[0])
    })
  }
})

describe("openXlsx -> saveXlsx keeps a foreign package intact", () => {
  for (const name of NAMES) {
    it(name, async () => {
      const original = load(name)
      const saved = await saveXlsx(await openXlsx(original))

      // Values first: a roundtrip that preserved every part and changed
      // the data would still be broken.
      const before = (await readXlsx(original)).sheets.map((s) => s.rows)
      const after = (await readXlsx(saved)).sheets.map((s) => s.rows)
      expect(after).toEqual(before)

      // Then the parts. ZIP *directory* entries are dropped — ExcelJS
      // writes them, hucre does not, they are optional per the spec and
      // no reader looks at them. Every real part has to survive.
      const originalParts = new ZipReader(original)
        .entries()
        .filter((p) => !p.endsWith("/"))
        .sort()
      const savedParts = new ZipReader(saved).entries().sort()

      for (const part of originalParts) {
        expect(savedParts, `${name} lost ${part}`).toContain(part)
      }
    })
  }
})

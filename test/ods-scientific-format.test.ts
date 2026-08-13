import { describe, expect, it } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// `docs/PARITY.md` lists number format among the six style facets ODS
// carries. A scientific one did not survive: `0.00E+00` was written as a
// plain `<number:number number:decimal-places="2"/>` and read back as
// `0.00`, the exponent gone.
//
// That is not only a round-trip loss. ODF has `<number:scientific-number>`
// for exactly this, so what hucre wrote *displayed* as a plain decimal in
// LibreOffice too — 1.2345E+03 shown as 1234.50.
//
// Found by a property test over random styles, checking each facet PARITY
// claims against what came back.
// ═══════════════════════════════════════════════════════════════════════

const dec = new TextDecoder()

async function styleOf(numFmt: string): Promise<CellStyle | undefined> {
  const bytes = await writeOds({
    sheets: [
      {
        name: "S",
        rows: [[1234.5]],
        cells: new Map([["0,0", { value: 1234.5, style: { numFmt } }]]),
      },
    ],
  })
  return (await readOds(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")?.style
}

async function contentXml(numFmt: string): Promise<string> {
  const bytes = await writeOds({
    sheets: [
      {
        name: "S",
        rows: [[1234.5]],
        cells: new Map([["0,0", { value: 1234.5, style: { numFmt } }]]),
      },
    ],
  })
  return dec.decode(await new ZipReader(bytes).extract("content.xml"))
}

describe("a scientific format survives ODS", () => {
  it("round-trips the common spellings", async () => {
    for (const numFmt of ["0.00E+00", "0.000E+00", "0.0E+00", "0E+00"]) {
      expect((await styleOf(numFmt))?.numFmt, numFmt).toBe(numFmt)
    }
  })

  it("as ODF's own element, so LibreOffice shows it too", async () => {
    // The assertion that matters to every other consumer. Reconstructing
    // the code on read would fix hucre's round trip and leave the file
    // still displaying a plain decimal everywhere else.
    const xml = await contentXml("0.00E+00")

    expect(xml).toContain("<number:scientific-number")
    expect(xml).toContain('number:min-exponent-digits="2"')
    expect(xml).toContain('number:decimal-places="2"')
  })

  it("with the exponent width it was given", async () => {
    expect(await contentXml("0.0E+0")).toContain('number:min-exponent-digits="1"')
    expect(await contentXml("0.0E+000")).toContain('number:min-exponent-digits="3"')
  })

  it("and XLSX still carries it, as it always did", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [[1234.5]],
          cells: new Map([["0,0", { value: 1234.5, style: { numFmt: "0.00E+00" } }]]),
        },
      ],
    })
    const back = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")

    expect(back?.style?.numFmt).toBe("0.00E+00")
  })
})

describe("the formats around it are unchanged", () => {
  it("plain, grouped, percentage, currency and date still round-trip", async () => {
    // A scientific check that fires on a bare `E` would catch these, and
    // `General` has one.
    for (const numFmt of [
      "0.00",
      "#,##0",
      "#,##0.00",
      "0%",
      "0.00%",
      "yyyy-mm-dd",
      "hh:mm:ss",
      '"$"#,##0.00',
    ]) {
      expect((await styleOf(numFmt))?.numFmt, numFmt).toBe(numFmt)
    }
  })

  it("a literal E in a quoted section is not an exponent", async () => {
    // `"E"0.00` is the letter E followed by a number, not scientific.
    const xml = await contentXml('"EUR "#,##0.00')

    expect(xml).not.toContain("<number:scientific-number")
  })
})

describe("the number-format codes ODS still cannot carry", () => {
  // Chasing the scientific gap turned up three more. None has an ODF
  // spelling, so they are losses rather than defects — but `PARITY.md`
  // listed number format as carried without saying which codes, and a
  // documented loss nobody checks is a claim, not a fact. These are the
  // three the doc now names.

  it("drops the width reserved by _", async () => {
    // `_)` reserves the width of a `)` so positives line up with
    // parenthesised negatives. ODF has nothing for it.
    expect((await styleOf("0.00_);(0.00)"))?.numFmt).toBe("0.00;(0.00)")
  })

  it("has no data style for the text format", async () => {
    // A cell already holding a string is unaffected — this only matters
    // for a number one asked to display as text.
    expect((await styleOf("@"))?.numFmt).toBeUndefined()
  })

  it("makes an optional digit a mandatory one", async () => {
    // `#` means "a digit if there is one" and `0` means "always a digit".
    // ODF counts digits, so `#.##` displays 1234.50 where Excel showed
    // 1234.5.
    expect((await styleOf("#"))?.numFmt).toBe("0")
    expect((await styleOf("#.##"))?.numFmt).toBe("0.00")
  })

  it("keeps the elapsed marker on hours but not on minutes or seconds", async () => {
    expect((await styleOf("[hh]:mm"))?.numFmt).toBe("[hh]:mm")
    expect((await styleOf("[mm]:ss"))?.numFmt).toBe("mm:ss")
    expect((await styleOf("[ss]"))?.numFmt).toBe("ss")
  })

  it("drops a colour tag, which is what keeps [White] out of the time branch", async () => {
    expect((await styleOf("0.00;[Red]-0.00"))?.numFmt).toBe("0.00;-0.00")
  })

  it("drops empty trailing sections", async () => {
    expect((await styleOf("0.00;;"))?.numFmt).toBe("0.00")
  })

  it("and General is no loss — it is the absence of a data style", async () => {
    expect((await styleOf("General"))?.numFmt).toBeUndefined()
  })

  it("flattens engineering notation to plain scientific", async () => {
    // `##0.0E+0` steps the integer part in threes. ODF spells that with
    // `number:exponent-interval`, which the writer does not emit, so the
    // exponent survives and the stepping does not.
    expect((await styleOf("##0.0E+0"))?.numFmt).toBe("0.0E+0")
  })
})

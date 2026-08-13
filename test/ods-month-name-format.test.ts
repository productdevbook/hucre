import { describe, expect, it } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { isDateFormat } from "../src/_date"
import { ZipReader } from "../src/zip/reader"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// `mmm` and `mmmm` — the month-name formats — were not recognised as
// dates by the ODS writer. They fell through to the plain-number branch,
// which has no `#` or `0` to work with, so the whole code became literal
// text: a cell formatted `mmm` displayed the letters "mmm" instead of
// "Mar".
//
// There are two implementations of this question in the repo, and only
// one had the rule. `src/_date.ts` — the shared one, used by every reader
// to tell a date cell from a number — ends with:
//
//   // "mmm" (abbreviated month name) or "mmmm" (full month name)
//   // are always date formats.
//   if (/m{3,}/.test(lower)) return true
//
// The ODS writer's local copy checked only for `y` and `d`. A lone `m` or
// `mm` is genuinely ambiguous — minutes after an `h` — and both agree it
// is not a date on its own. Three or more never are.
// ═══════════════════════════════════════════════════════════════════════

const dec = new TextDecoder()

async function odsNumFmt(numFmt: string): Promise<string | undefined> {
  const cells = new Map([["0,0", { value: 45373, style: { numFmt } as CellStyle }]])
  const bytes = await writeOds({ sheets: [{ name: "S", rows: [[45373]], cells }] })
  return (await readOds(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")?.style?.numFmt
}

async function contentXml(numFmt: string): Promise<string> {
  const cells = new Map([["0,0", { value: 45373, style: { numFmt } as CellStyle }]])
  const bytes = await writeOds({ sheets: [{ name: "S", rows: [[45373]], cells }] })
  return dec.decode(await new ZipReader(bytes).extract("content.xml"))
}

describe("a month-name format is a date format", () => {
  it("round-trips through ODS", async () => {
    expect(await odsNumFmt("mmm")).toBe("mmm")
    expect(await odsNumFmt("mmmm")).toBe("mmmm")
  })

  it("as a date style, not as literal text", async () => {
    // The failure mode was a cell displaying the letters "mmm". The
    // element is what decides that, in LibreOffice as much as on re-read.
    const xml = await contentXml("mmm")

    expect(xml).toContain("<number:date-style")
    expect(xml).toContain('number:textual="true"')
    expect(xml).not.toContain("<number:text>m</number:text>")
  })

  it("agreeing with the shared isDateFormat, which always said so", () => {
    // The two answers to one question, now the same.
    expect(isDateFormat("mmm")).toBe(true)
    expect(isDateFormat("mmmm")).toBe(true)
  })

  it("and in combination, as it already did", async () => {
    for (const numFmt of ["mmm-yy", "d-mmm-yy", "mmmm d, yyyy"]) {
      expect(await odsNumFmt(numFmt), numFmt).toBe(numFmt)
    }
  })
})

describe("a lone m is still not a date, in either implementation", () => {
  it("because after an h it means minutes", () => {
    // Deliberate in `_date.ts`, and the reason is written there. The ODS
    // writer has to agree or a duration becomes a month.
    expect(isDateFormat("m")).toBe(false)
    expect(isDateFormat("mm")).toBe(false)
  })

  it("and a time format keeps its minutes", async () => {
    for (const numFmt of ["hh:mm", "hh:mm:ss", "mm:ss"]) {
      expect(await odsNumFmt(numFmt), numFmt).toBe(numFmt)
    }
  })

  it("including the elapsed-hour one", async () => {
    expect(await odsNumFmt("[hh]:mm")).toBe("[hh]:mm")
  })
})

describe("nothing else moved", () => {
  it("dates, numbers, percentages and currency still round-trip", async () => {
    for (const numFmt of [
      "yyyy-mm-dd",
      "dd/mm/yyyy",
      "0.00",
      "#,##0",
      "0%",
      '"$"#,##0.00',
      '#,##0" days"',
    ]) {
      expect(await odsNumFmt(numFmt), numFmt).toBe(numFmt)
    }
  })
})

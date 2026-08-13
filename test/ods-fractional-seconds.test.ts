import { describe, expect, it } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { ZipReader } from "../src/zip/reader"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// `hh:mm:ss.0` came back as `hh:mm:ss."0"` — the fractional part became
// literal text, so the cell displayed the digit rather than a fraction of
// a second.
//
// Excel spells the precision by writing the places after the token;
// ODF puts the count on the element as `number:decimal-places`. The
// tokeniser hands `.` and each `0` over separately, and nothing consumed
// them, so they reached the literal-text branch.
//
// Timing and instrument exports use these. Found by a sweep of number
// formats through the ODS round trip.
// ═══════════════════════════════════════════════════════════════════════

const dec = new TextDecoder()

async function odsNumFmt(numFmt: string): Promise<string | undefined> {
  const cells = new Map([["0,0", { value: 0.5, style: { numFmt } as CellStyle }]])
  const bytes = await writeOds({ sheets: [{ name: "S", rows: [[0.5]], cells }] })
  return (await readOds(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")?.style?.numFmt
}

async function contentXml(numFmt: string): Promise<string> {
  const cells = new Map([["0,0", { value: 0.5, style: { numFmt } as CellStyle }]])
  const bytes = await writeOds({ sheets: [{ name: "S", rows: [[0.5]], cells }] })
  return dec.decode(await new ZipReader(bytes).extract("content.xml"))
}

describe("fractional seconds survive", () => {
  it("round-trip at each precision", async () => {
    for (const numFmt of ["hh:mm:ss.0", "hh:mm:ss.00", "hh:mm:ss.000", "mm:ss.0", "ss.00"]) {
      expect(await odsNumFmt(numFmt), numFmt).toBe(numFmt)
    }
  })

  it("as decimal-places on the element, not as literal text", async () => {
    // What LibreOffice reads. A cell showing the digit `0` after the
    // seconds is wrong there too, so the round trip alone is not enough.
    const xml = await contentXml("hh:mm:ss.00")

    expect(xml).toContain('number:decimal-places="2"')
    expect(xml).not.toContain("<number:text>0</number:text>")
  })

  it("and the seconds element keeps its own style", async () => {
    // `ss` is two-digit, `s` is not — the precision must not overwrite it.
    expect(await odsNumFmt("hh:mm:s.0")).toBe("hh:mm:s.0")
    expect(await odsNumFmt("hh:mm:ss.0")).toBe("hh:mm:ss.0")
  })
})

describe("seconds without a fraction are untouched", () => {
  it("the ordinary time formats", async () => {
    for (const numFmt of ["hh:mm:ss", "mm:ss", "hh:mm", "h:mm AM/PM", "[hh]:mm"]) {
      expect(await odsNumFmt(numFmt), numFmt).toBe(numFmt)
    }
  })

  it("a trailing dot with no digits is left alone", async () => {
    // The consumer only fires on `.` followed by at least one `0`.
    const back = await odsNumFmt("hh:mm:ss.")

    expect(back).toContain("ss")
  })

  it("and a date is still a date", async () => {
    for (const numFmt of ["yyyy-mm-dd", "yyyy-mm-dd hh:mm:ss", "mmm"]) {
      expect(await odsNumFmt(numFmt), numFmt).toBe(numFmt)
    }
  })
})

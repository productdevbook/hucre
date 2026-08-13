import { describe, expect, it } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { ZipReader } from "../src/zip/reader"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// In an Excel number format `0` is a digit that is always shown and `#`
// is one shown only if there is something to show. `#.##` displays 1234.5
// as "1234.5"; `0.00` displays it as "1234.50".
//
// #535 recorded that distinction as a loss ODS could not carry — "ODF
// counts digits, so the distinction goes". It does not. ODF carries both
// counts: `number:decimal-places` is the maximum and
// `number:min-decimal-places` the minimum, and `number:min-integer-digits`
// does the same job on the other side of the point. LibreOffice writes
// all three.
//
// Found the same way as the text format in #548 — by crossing the ODF
// grammar with a LibreOffice document rather than by re-reading the page
// that said it was impossible.
// ═══════════════════════════════════════════════════════════════════════

const dec = new TextDecoder()

async function bytesFor(numFmt: string): Promise<Uint8Array> {
  return writeOds({
    sheets: [
      {
        name: "S",
        rows: [[1234.5]],
        cells: new Map([["0,0", { value: 1234.5, style: { numFmt } as CellStyle }]]),
      },
    ],
  })
}

async function roundTrip(numFmt: string): Promise<string | undefined> {
  const wb = await readOds(await bytesFor(numFmt), { readStyles: true })
  return wb.sheets[0]!.cells?.get("0,0")?.style?.numFmt
}

describe("an optional decimal stays optional", () => {
  it("round-trips the forms that differ only in # versus 0", async () => {
    for (const numFmt of ["0.00", "#.##", "0.0#", "#,##0.00", "#,##0.##"]) {
      expect(await roundTrip(numFmt), numFmt).toBe(numFmt)
    }
  })

  it("writes both counts, which is how ODF says it", async () => {
    const xml = dec.decode(await new ZipReader(await bytesFor("0.0#")).extract("content.xml"))

    // Two decimals at most, one always shown.
    expect(xml).toContain('number:decimal-places="2"')
    expect(xml).toContain('number:min-decimal-places="1"')
  })

  it("a fully mandatory format says so", async () => {
    const xml = dec.decode(await new ZipReader(await bytesFor("0.00")).extract("content.xml"))

    expect(xml).toContain('number:decimal-places="2"')
    expect(xml).toContain('number:min-decimal-places="2"')
  })
})

describe("an optional integer digit too", () => {
  it("keeps # distinct from 0", async () => {
    expect(await roundTrip("#")).toBe("#")
    expect(await roundTrip("0")).toBe("0")
  })

  it("and keeps grouping when there is no mandatory digit", async () => {
    // `#,###` lost its separator entirely, coming back as `0`.
    expect(await roundTrip("#,###")).toBe("#,###")
  })

  it("writes min-integer-digits 0 for a wholly optional integer part", async () => {
    const xml = dec.decode(await new ZipReader(await bytesFor("#.##")).extract("content.xml"))

    expect(xml).toContain('number:min-integer-digits="0"')
  })
})

describe("the formats around them are unchanged", () => {
  it("still round-trip", async () => {
    for (const numFmt of [
      "#,##0",
      "0%",
      "0.00%",
      "yyyy-mm-dd",
      "hh:mm:ss",
      '"$"#,##0.00',
      "0.00E+00",
      "@",
    ]) {
      expect(await roundTrip(numFmt), numFmt).toBe(numFmt)
    }
  })
})

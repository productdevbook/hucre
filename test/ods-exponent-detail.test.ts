import { describe, expect, it } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { ZipReader } from "../src/zip/reader"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// Two details of a scientific format that #529 left behind.
//
// `E+` forces a sign on a positive exponent; `E-` shows one only when the
// exponent is negative. hucre wrote both as `E+` and read every one back
// that way, so `0.00E-00` became `0.00E+00`. ODF spells the difference
// `number:forced-exponent-sign`.
//
// `##0.0E+0` is engineering notation: the exponent moves in steps of
// three so the integer part stays between 1 and 999. #529 recorded that
// as a loss and named its ODF spelling — `number:exponent-interval` — in
// the same sentence, then did not write it. LibreOffice writes both
// attributes on every scientific style it produces.
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

describe("the exponent sign is carried", () => {
  it("keeps E+ and E- apart", async () => {
    expect(await roundTrip("0.00E+00")).toBe("0.00E+00")
    expect(await roundTrip("0.00E-00")).toBe("0.00E-00")
  })

  it("says which it is in the file", async () => {
    const forced = dec.decode(
      await new ZipReader(await bytesFor("0.00E+00")).extract("content.xml"),
    )
    const plain = dec.decode(await new ZipReader(await bytesFor("0.00E-00")).extract("content.xml"))

    expect(forced).toContain('number:forced-exponent-sign="true"')
    expect(plain).not.toContain('number:forced-exponent-sign="true"')
  })
})

describe("engineering notation is carried", () => {
  it("round-trips ##0.0E+0", async () => {
    expect(await roundTrip("##0.0E+0")).toBe("##0.0E+0")
  })

  it("as an exponent interval of three", async () => {
    const xml = dec.decode(await new ZipReader(await bytesFor("##0.0E+0")).extract("content.xml"))

    expect(xml).toContain('number:exponent-interval="3"')
  })

  it("and a plain scientific format does not claim one", async () => {
    const xml = dec.decode(await new ZipReader(await bytesFor("0.00E+00")).extract("content.xml"))

    expect(xml).not.toContain('number:exponent-interval="3"')
  })
})

describe("the rest of the scientific handling is unchanged", () => {
  it("widths still round-trip", async () => {
    for (const numFmt of ["0.00E+00", "0.000E+00", "0.0E+00", "0E+00", "0.0E+0"]) {
      expect(await roundTrip(numFmt), numFmt).toBe(numFmt)
    }
  })

  it("and a literal E is still not an exponent", async () => {
    const xml = dec.decode(
      await new ZipReader(await bytesFor('"EUR "#,##0.00')).extract("content.xml"),
    )

    expect(xml).not.toContain("<number:scientific-number")
  })
})

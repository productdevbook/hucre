import { describe, expect, it } from "vitest"
import { decodePart, tooLargeToDecode, MAX_STRING_LENGTH } from "../src/_decode"
import { ParseError, HucreError } from "../src/errors"

// ═══════════════════════════════════════════════════════════════════════
// #503 — a worksheet part over V8's 512 MB string ceiling made the
// buffered readers die with
//
//   Error: Cannot create a string longer than 0x1fffffe8 characters
//
// which is not a `ParseError`, names no part, says nothing about
// spreadsheets, and arrives after however long the decompression took.
// Three files in a corpus of ~600 hit it — instrument logs at Excel's
// row limit, 56–99 MB compressed and 607 MB expanded.
//
// It is not "hucre cannot read big files": `streamXlsxRows` reads exactly
// those files, in ~30s at a flat 944 MB. The buffered reader had a hard
// ceiling it did not know about.
//
// The trigger is not reproducible here and that is worth stating rather
// than working around: it needs a part larger than this repository.
// Faking the decode failure would test the wrapper rather than the
// condition, so what is tested is what has logic in it — the decision
// and the message.
// ═══════════════════════════════════════════════════════════════════════

describe("the error a part over the ceiling produces", () => {
  const error = tooLargeToDecode("xl/worksheets/sheet1.xml", 607_505_144)

  it("is one of ours, so a caller catching HucreError sees it", () => {
    expect(error).toBeInstanceOf(ParseError)
    expect(error).toBeInstanceOf(HucreError)
  })

  it("names the part", () => {
    expect(error.message).toContain("xl/worksheets/sheet1.xml")
  })

  it("names the measurement and the bound, as the maxTotalCells error does", () => {
    expect(error.message).toContain("607,505,144")
    expect(error.message).toContain(MAX_STRING_LENGTH.toLocaleString("en-US"))
  })

  it("names the way out", () => {
    expect(error.message).toContain("streamXlsxRows")
  })

  it("says the workbook is not damaged, because it is not", () => {
    // Everything else that throws from a reader means the file is wrong.
    // This one means the file is large, and a caller deciding whether to
    // retry or to reject needs to know which.
    expect(error.message).toContain("not damaged")
  })
})

describe("decodePart", () => {
  it("decodes ordinary bytes unchanged", () => {
    expect(decodePart(new TextEncoder().encode("hello şehir"), "x.xml")).toBe("hello şehir")
  })

  it("passes through an error that is not the ceiling", () => {
    // A RangeError is the ceiling; anything else is not ours to
    // reinterpret, and swallowing it would hide a real bug.
    const notBytes = { byteLength: 1, length: 1 } as unknown as Uint8Array

    expect(() => decodePart(notBytes, "x.xml")).toThrow()
    expect(() => decodePart(notBytes, "x.xml")).not.toThrow(ParseError)
  })

  it("reports the ceiling as the number V8 actually uses", () => {
    // 0x1fffffe8. Hard-coded here so a change to the constant is a
    // decision rather than a typo.
    expect(MAX_STRING_LENGTH).toBe(536_870_888)
  })
})

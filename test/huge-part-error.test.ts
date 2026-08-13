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
// The first version of this file said the trigger was not reproducible —
// that it needed a part larger than the repository — and tested only the
// message factory. That was wrong, and the gap it left was a real bug:
// the guard tested `instanceof RangeError`, Node throws a plain `Error`
// with `code: "ERR_STRING_TOO_LONG"`, so on Node the conversion never
// happened and the raw error reached the caller exactly as before (#516).
//
// The condition needs a large *buffer*, not a large *file*. One
// allocation of MAX_STRING_LENGTH + 1 bytes reproduces it in about a
// second, and that test is below. A test that could not fail for the
// right reason is worth less than the paragraph explaining why it does
// not exist.
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

  it("turns the real ceiling into a ParseError", () => {
    // The test that was missing. ~537 MB, outside the JS heap, about a
    // second — and it is the only thing here that exercises the `catch`
    // rather than the message it builds.
    const tooBig = new Uint8Array(MAX_STRING_LENGTH + 1)

    expect(() => decodePart(tooBig, "xl/worksheets/sheet1.xml")).toThrow(ParseError)
    expect(() => decodePart(tooBig, "xl/worksheets/sheet1.xml")).toThrow(/streamXlsxRows/)
  }, 60_000)

  it("recognises it whatever the engine calls the error", () => {
    // Node throws `Error` + `code: "ERR_STRING_TOO_LONG"`; other engines
    // throw `RangeError`. The byte length is the backstop that does not
    // depend on guessing which.
    expect(() => decodePart(new Uint8Array(MAX_STRING_LENGTH + 1), "p.xml")).toThrow(ParseError)
  }, 60_000)

  it("reports the ceiling as the number V8 actually uses", () => {
    // 0x1fffffe8. Hard-coded here so a change to the constant is a
    // decision rather than a typo.
    expect(MAX_STRING_LENGTH).toBe(536_870_888)
  })
})

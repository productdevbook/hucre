import { valuesOf } from "./_stream"
import { describe, expect, it } from "vitest"
import { parseCsv, parseCsvObjects } from "../src/csv/reader"
import { streamCsvRows } from "../src/csv/stream"
import { detectBom, decodeCsvInput } from "../src/csv/encoding"
import { InvalidArgumentError } from "../src/errors"

// ═══════════════════════════════════════════════════════════════════════
// #475 — `parseCsv` took a string while every other reader took bytes, so
// the byte→string step was the caller's alone. That is where the actual
// difficulty of CSV lives: Excel on a Turkish Windows writes
// windows-1254, and "Save as Unicode Text" writes UTF-16LE with a BOM.
//
// What this does is decode, not guess. A byte-order mark is a statement
// the file makes about itself and is honoured; anything else has to be
// named, because telling windows-1254 from windows-1252 by byte frequency
// is a guess wrong often enough to be worse than asking.
// ═══════════════════════════════════════════════════════════════════════

/** `ad,şehir\nÖzgür,İstanbul` as Excel-on-Turkish-Windows writes it. */
const WINDOWS_1254 = new Uint8Array([
  0x61,
  0x64,
  0x2c,
  0xfe,
  0x65,
  0x68,
  0x69,
  0x72,
  0x0a, // ad,şehir
  0xd6,
  0x7a,
  0x67,
  0xfc,
  0x72,
  0x2c,
  0xdd,
  0x73,
  0x74,
  0x61,
  0x6e,
  0x62,
  0x75,
  0x6c, // Özgür,İstanbul
])

const TEXT = "ad,şehir\nÖzgür,İstanbul"
const EXPECTED = [
  ["ad", "şehir"],
  ["Özgür", "İstanbul"],
]

function utf8(text: string, bom = false): Uint8Array {
  return new TextEncoder().encode(bom ? `﻿${text}` : text)
}

function utf16(text: string, littleEndian: boolean): Uint8Array {
  const withBom = `﻿${text}`
  const out = new Uint8Array(withBom.length * 2)
  const view = new DataView(out.buffer)
  for (let i = 0; i < withBom.length; i++) {
    view.setUint16(i * 2, withBom.charCodeAt(i), littleEndian)
  }
  return out
}

describe("detectBom", () => {
  it("reads the three marks", () => {
    expect(detectBom(utf8(TEXT, true))).toEqual({ encoding: "utf-8", length: 3 })
    expect(detectBom(utf16(TEXT, true))).toEqual({ encoding: "utf-16le", length: 2 })
    expect(detectBom(utf16(TEXT, false))).toEqual({ encoding: "utf-16be", length: 2 })
  })

  it("says nothing when there is no mark", () => {
    expect(detectBom(utf8(TEXT))).toBeNull()
    expect(detectBom(WINDOWS_1254)).toBeNull()
    expect(detectBom(new Uint8Array([]))).toBeNull()
    expect(detectBom(new Uint8Array([0xef]))).toBeNull()
  })
})

describe("parseCsv takes bytes", () => {
  it("reads plain UTF-8", () => {
    expect(parseCsv(utf8(TEXT))).toEqual(EXPECTED)
  })

  it("reads UTF-8 with a BOM without leaving it in the first header", () => {
    const rows = parseCsv(utf8(TEXT, true))

    expect(rows).toEqual(EXPECTED)
    expect(rows[0]![0]).toBe("ad")
  })

  it('reads UTF-16LE, which is what Excel\'s "Unicode Text" writes', () => {
    // Decoded as UTF-8 this is a run of NUL-separated letters that a CSV
    // parser reads as data rather than rejecting.
    expect(parseCsv(utf16(TEXT, true))).toEqual(EXPECTED)
  })

  it("reads UTF-16BE", () => {
    expect(parseCsv(utf16(TEXT, false))).toEqual(EXPECTED)
  })

  it("reads a named legacy encoding no mark can declare", () => {
    expect(parseCsv(WINDOWS_1254, { encoding: "windows-1254" })).toEqual(EXPECTED)
  })

  it("does not guess at that encoding", () => {
    // Without being told, these bytes are read as UTF-8 and come back as
    // replacement characters. That is the honest answer — there is nothing
    // in the file that says otherwise.
    expect(parseCsv(WINDOWS_1254)[0]![1]).not.toBe("şehir")
  })

  it("lets an explicit encoding override the mark", () => {
    // A file can carry a mark that is simply wrong; the caller knows.
    expect(parseCsv(utf8(TEXT, true), { encoding: "utf-8" })).toEqual(EXPECTED)
  })

  it("takes an ArrayBuffer as well as a Uint8Array", () => {
    const bytes = utf8(TEXT)
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)

    expect(parseCsv(buffer as ArrayBuffer)).toEqual(EXPECTED)
  })

  it("leaves a string exactly as it was", () => {
    expect(parseCsv(TEXT)).toEqual(EXPECTED)
  })

  it("refuses an encoding label TextDecoder cannot use", () => {
    expect(() => parseCsv(utf8(TEXT), { encoding: "not-an-encoding" })).toThrow(
      InvalidArgumentError,
    )
  })
})

describe("the other readers take bytes too", () => {
  it("parseCsvObjects", () => {
    const { data, headers } = parseCsvObjects(WINDOWS_1254, {
      header: true,
      encoding: "windows-1254",
    })

    expect(headers).toEqual(["ad", "şehir"])
    expect(data).toEqual([{ ad: "Özgür", şehir: "İstanbul" }])
  })

  it("streamCsvRows", async () => {
    expect(await valuesOf(streamCsvRows(utf16(TEXT, true)))).toEqual(EXPECTED)
    expect(await valuesOf(streamCsvRows(WINDOWS_1254, { encoding: "windows-1254" }))).toEqual(
      EXPECTED,
    )
  })

  it("streamCsvRows still takes a string", async () => {
    expect(await valuesOf(streamCsvRows(TEXT))).toEqual(EXPECTED)
  })
})

describe("decodeCsvInput", () => {
  it("prefers the named encoding, then the mark, then utf-8", () => {
    expect(decodeCsvInput(WINDOWS_1254, "windows-1254")).toBe(TEXT)
    expect(decodeCsvInput(utf16(TEXT, true))).toBe(TEXT)
    expect(decodeCsvInput(utf8(TEXT))).toBe(TEXT)
  })

  it("strips the mark whichever path decoded it", () => {
    // TextDecoder drops a UTF-8 BOM itself but not a UTF-16 one when the
    // caller named the encoding, so the three paths would otherwise differ.
    expect(decodeCsvInput(utf16(TEXT, true), "utf-16le")).toBe(TEXT)
    expect(decodeCsvInput(utf8(TEXT, true), "utf-8")).toBe(TEXT)
  })
})

import { describe, expect, it } from "vitest"
import { deflate, inflate } from "../src/zip/deflate"
import { parseWorksheet } from "../src/xlsx/worksheet"
import type { WorksheetContext } from "../src/xlsx/worksheet"
import { parseSst } from "../src/xls/biff"
import { readCfb } from "../src/xlsx/crypto/cfb"
import { decryptAgile } from "../src/xlsx/crypto/agile"
import { toHtml } from "../src/export/html"
import { ParseError, ZipError, DecryptionError } from "../src/errors"
import type { CellValue, Sheet } from "../src/_types"

const ctx: WorksheetContext = {
  sharedStrings: [],
  styles: null,
  readStyles: false,
  dateSystem: "1900",
}

describe("DoS hardening — zip bomb (inflate cap)", () => {
  it("throws ZipError when decompressed output exceeds the cap", () => {
    // Highly compressible input that inflates to ~10 KiB; cap at 256 bytes.
    const raw = new Uint8Array(10_000) // all zeros — compresses tiny
    const compressed = deflate(raw)
    expect(() => inflate(compressed, 256)).toThrow(ZipError)
  })

  it("still inflates normally under the cap", () => {
    const raw = new Uint8Array(1000).fill(7)
    const compressed = deflate(raw)
    const out = inflate(compressed, 1_000_000)
    expect(out.length).toBe(1000)
    expect(out[0]).toBe(7)
  })
})

describe("DoS hardening — cell-ref OOM (worksheet bounds)", () => {
  it("rejects a cell reference beyond Excel's row limit", () => {
    const xml = `<worksheet><sheetData><row r="2000000"><c r="A2000000"><v>1</v></c></row></sheetData></worksheet>`
    expect(() => parseWorksheet(xml, "S", ctx)).toThrow(ParseError)
  })

  it("rejects a cell reference beyond Excel's column limit", () => {
    // Column XFE = 16385 (> 16384 max).
    const xml = `<worksheet><sheetData><row r="1"><c r="XFE1"><v>1</v></c></row></sheetData></worksheet>`
    expect(() => parseWorksheet(xml, "S", ctx)).toThrow(ParseError)
  })

  it("accepts an in-bounds reference", () => {
    const xml = `<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`
    const sheet = parseWorksheet(xml, "S", ctx)
    expect(sheet.rows[0][0]).toBe(1)
  })
})

describe("DoS hardening — XLS SST loop bound", () => {
  it("does not spin on an inflated cstUnique", () => {
    // cstTotal (4) + cstUnique (4 = 0xFFFFFFFF) with no string data.
    const block = new Uint8Array(8)
    const dv = new DataView(block.buffer)
    dv.setUint32(0, 0xffffffff, true) // cstTotal
    dv.setUint32(4, 0xffffffff, true) // cstUnique — hostile
    const out = parseSst([block])
    // Bounded by available bytes; returns quickly without OOM/hang.
    expect(Array.isArray(out)).toBe(true)
    expect(out.length).toBe(0)
  })
})

describe("DoS hardening — CFB DIFAT / mini-stream bounds", () => {
  it("does not allocate from a hostile mini-stream size", () => {
    // A minimal-but-corrupt CFB: valid signature, but the DIFAT/mini walks
    // must terminate by file size rather than header-claimed counts.
    const data = new Uint8Array(512)
    // CFB signature
    const sig = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
    data.set(sig, 0)
    const dv = new DataView(data.buffer)
    dv.setUint16(30, 9, true) // sectorShift -> 512
    dv.setUint16(32, 6, true) // miniSectorShift -> 64
    dv.setUint32(44, 1, true) // numFatSectors
    dv.setUint32(72, 0xffffffff, true) // numDifatSectors — hostile, must be bounded
    // readCfb should fail with a typed ParseError rather than hanging/OOMing.
    expect(() => readCfb(data)).toThrow(ParseError)
  })
})

describe("DoS hardening — Agile crypto spinCount cap", () => {
  it("rejects an absurd spinCount before deriving the key", async () => {
    // Build a CFB with EncryptionInfo whose spinCount is enormous. Decrypt
    // must reject via DecryptionError instead of spinning for minutes.
    const xml =
      `<?xml version="1.0"?><encryption><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>` +
      `<keyEncryptors><keyEncryptor><p:encryptedKey spinCount="999999999" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA==" encryptedVerifierHashInput="AAAAAAAAAAAAAAAAAAAAAA==" encryptedVerifierHashValue="AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" encryptedKeyValue="AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="/></keyEncryptor></keyEncryptors></encryption>`
    const { writeCfb } = await import("../src/xlsx/crypto/cfb")
    const infoHeader = new Uint8Array(8)
    const dv = new DataView(infoHeader.buffer)
    dv.setUint16(0, 4, true) // major
    dv.setUint16(2, 4, true) // minor
    const info = new Uint8Array([...infoHeader, ...new TextEncoder().encode(xml)])
    const pkg = new Uint8Array(8 + 16) // dummy
    const cfb = writeCfb([
      { name: "EncryptionInfo", data: info },
      { name: "EncryptedPackage", data: pkg },
    ])
    await expect(decryptAgile(cfb, "x")).rejects.toThrow(DecryptionError)
  })
})

describe("HTML export XSS — style attribute escaping", () => {
  it("escapes quotes/brackets injected via font name", () => {
    const sheet: Sheet = {
      name: "S",
      rows: [["x"]] as CellValue[][],
      cells: new Map([
        [
          "0,0",
          {
            value: "x",
            type: "string",
            style: { font: { name: `Arial";}</style><script>alert(1)</script>` } },
          },
        ],
      ]),
    } as unknown as Sheet
    const html = toHtml(sheet, { styles: true })
    expect(html).not.toContain("<script>alert(1)</script>")
    expect(html).toContain("&quot;")
  })
})

// ── read() on an OLE2 container that is not an intact encrypted package ──
//
// `read()` used to treat every OLE2 container it could not parse as an
// encrypted workbook, so a truncated legacy .xls surfaced as
// `EncryptedFileError` ("File is password-protected") — and with a
// password it went on to try decrypting garbage. A container `readCfb`
// cannot open is damaged, not encrypted; an encrypted OOXML package is a
// parseable CFB that carries `EncryptionInfo` / `EncryptedPackage`.
// ──────────────────────────────────────────────────────────────────────

import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { read } from "../src/defter"
import { readXls } from "../src/xls/reader"
import { writeCfb } from "../src/xlsx/crypto/cfb"
import { writeXlsx } from "../src/xlsx/writer"
import {
  DecryptionError,
  EncryptedFileError,
  ParseError,
  UnsupportedFormatError,
} from "../src/errors"

const xls = new Uint8Array(readFileSync(new URL("./fixtures/excel-basic.xls", import.meta.url)))

/** `bytes` cut to `percent`% of their length. */
const cut = (bytes: Uint8Array, percent: number): Uint8Array =>
  bytes.subarray(0, Math.floor((bytes.length * percent) / 100))

/** A genuine encrypted OOXML package, written by hucre's own `writeXlsx`. */
const encryptedPackage = (): Promise<Uint8Array> =>
  writeXlsx({
    sheets: [{ name: "S", rows: [["a"]] }],
    encryption: { password: "pw", spinCount: 64 },
  })

const PERCENTS = Array.from({ length: 99 }, (_, i) => i + 1)

async function errorOf(p: Promise<unknown>): Promise<unknown> {
  try {
    await p
  } catch (err) {
    return err
  }
  throw new Error("expected the read to reject")
}

describe("read() — truncated legacy .xls", () => {
  it("still reads the intact fixture", async () => {
    const wb = await read(xls)
    expect(wb.sheets.length).toBeGreaterThan(0)
  })

  it("rejects every cut from 1% to 99% with a ParseError, not EncryptedFileError", async () => {
    const wrong: string[] = []
    for (const p of PERCENTS) {
      const err = await errorOf(read(cut(xls, p)))
      // EncryptedFileError is not a ParseError, so this one check excludes it.
      if (!(err instanceof ParseError)) {
        wrong.push(`${p}%: ${(err as Error).name}`)
      }
    }
    expect(wrong).toEqual([])
  })

  it("names the damage and keeps the container error as the cause", async () => {
    const err = await errorOf(read(cut(xls, 50)))
    expect(err).toBeInstanceOf(ParseError)
    expect((err as Error).message).toMatch(/truncated or damaged/)
    expect((err as Error).cause).toBeInstanceOf(Error)
  })

  it("does not try to decrypt a truncated file when a password is given", async () => {
    for (const p of [1, 25, 50, 75, 99]) {
      const err = await errorOf(read(cut(xls, p), { password: "pw" }))
      expect(err).toBeInstanceOf(ParseError)
    }
  })

  it("agrees with readXls on the error class", async () => {
    const direct = await errorOf(readXls(cut(xls, 50)))
    expect(direct).toBeInstanceOf(ParseError)
    expect(await errorOf(read(cut(xls, 50)))).toBeInstanceOf(ParseError)
  })
})

describe("read() — other OLE2 containers", () => {
  it("rejects the bare OLE2 magic padded with zeros as damaged", async () => {
    const data = new Uint8Array(264)
    data.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
    const err = await errorOf(read(data))
    expect(err).toBeInstanceOf(ParseError)
  })

  it("rejects an intact CFB with neither a workbook nor an encrypted package as unsupported", async () => {
    const data = writeCfb([{ name: "SummaryInformation", data: new Uint8Array(64) }])
    for (const options of [undefined, { password: "pw" }]) {
      const err = await errorOf(read(data, options))
      expect(err).toBeInstanceOf(UnsupportedFormatError)
      expect((err as Error).message).toBe(
        "Unsupported format: unknown (OLE2 container with neither a Workbook stream nor an encrypted package)",
      )
    }
  })

  it("still reports an intact encrypted package as encrypted", async () => {
    const err = await errorOf(read(await encryptedPackage()))
    expect(err).toBeInstanceOf(EncryptedFileError)
    expect((err as EncryptedFileError).format).toBeUndefined()
  })

  it("treats a container with only one of the encryption streams as encrypted", async () => {
    for (const name of ["EncryptionInfo", "EncryptedPackage"]) {
      const data = writeCfb([{ name, data: new Uint8Array(64) }])
      expect(await errorOf(read(data))).toBeInstanceOf(EncryptedFileError)
      const withPassword = await errorOf(read(data, { password: "pw" }))
      expect(withPassword).toBeInstanceOf(DecryptionError)
      expect((withPassword as Error).message).toMatch(/missing EncryptionInfo\/EncryptedPackage/)
    }
  })

  it("reports a truncated encrypted package as damaged, not encrypted", async () => {
    const enc = await encryptedPackage()
    for (const p of [10, 50, 90]) {
      const err = await errorOf(read(cut(enc, p)))
      expect(err).toBeInstanceOf(ParseError)
    }
  })
})

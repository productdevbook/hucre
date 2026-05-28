import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import { read, readObjects } from "../src/defter"
import { isOle2Container } from "../src/_input"
import { DecryptionError, EncryptedFileError } from "../src/errors"
import { readCfb, writeCfb } from "../src/xlsx/crypto/cfb"
import { decryptAgile, encryptAgile } from "../src/xlsx/crypto/agile"
import type { CellValue } from "../src/_types"

const FAST = { spinCount: 64 }

function book() {
  return {
    sheets: [
      {
        name: "Sheet1",
        rows: [
          ["Name", "Score"],
          ["Ada", 95],
          ["Linus", 88],
        ] as CellValue[][],
      },
    ],
  }
}

describe("CFB container", () => {
  it("round-trips mini (small) and regular (large) streams", () => {
    const small = new Uint8Array(1300).map((_, i) => i & 0xff)
    const big = new Uint8Array(20000).map((_, i) => (i * 31) & 0xff)
    const streams = readCfb(
      writeCfb([
        { name: "EncryptionInfo", data: small },
        { name: "EncryptedPackage", data: big },
      ]),
    )
    expect([...streams.get("EncryptionInfo")!]).toEqual([...small])
    expect([...streams.get("EncryptedPackage")!]).toEqual([...big])
  })
})

describe("agile crypto primitive", () => {
  it("encrypt → decrypt recovers arbitrary bytes", async () => {
    const payload = new TextEncoder().encode("PK" + "payload ".repeat(900))
    const enc = await encryptAgile(payload, "hunter2", FAST)
    expect(isOle2Container(enc)).toBe(true)
    expect([...(await decryptAgile(enc, "hunter2"))]).toEqual([...payload])
  })

  it("wrong password rejects with DecryptionError", async () => {
    const enc = await encryptAgile(new TextEncoder().encode("data ".repeat(2000)), "right", FAST)
    await expect(decryptAgile(enc, "wrong")).rejects.toBeInstanceOf(DecryptionError)
  })

  it("interoperates at Excel's default spin count", { timeout: 30000 }, async () => {
    const payload = new TextEncoder().encode("z".repeat(9000))
    const enc = await encryptAgile(payload, "pw") // default 100000
    expect([...(await decryptAgile(enc, "pw"))]).toEqual([...payload])
  })
})

describe("writeXlsx encryption ↔ readXlsx decryption", () => {
  it("encrypts on write and decrypts on read with the password", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    expect(isOle2Container(enc)).toBe(true) // output is an encrypted OLE2 container, not a ZIP

    const wb = await readXlsx(enc, { password: "pw" })
    expect(wb.sheets[0].name).toBe("Sheet1")
    expect(wb.sheets[0].rows[1]).toEqual(["Ada", 95])
  })

  it("reading without a password throws EncryptedFileError", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    await expect(readXlsx(enc)).rejects.toBeInstanceOf(EncryptedFileError)
  })

  it("reading with the wrong password throws DecryptionError", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    await expect(readXlsx(enc, { password: "nope" })).rejects.toBeInstanceOf(DecryptionError)
  })
})

describe("decryption across the read entry points", () => {
  it("read() auto-detects and decrypts", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    const wb = await read(enc, { password: "pw" })
    expect(wb.sheets[0].rows[2]).toEqual(["Linus", 88])
  })

  it("readObjects() decrypts", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    const objs = await readObjects<{ Name: string; Score: number }>(enc, { password: "pw" })
    expect(objs).toEqual([
      { Name: "Ada", Score: 95 },
      { Name: "Linus", Score: 88 },
    ])
  })

  it("streamXlsxRows() decrypts", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    const rows: CellValue[][] = []
    for await (const row of streamXlsxRows(enc, { password: "pw" })) rows.push(row.values)
    expect(rows[0]).toEqual(["Name", "Score"])
    expect(rows[1]).toEqual(["Ada", 95])
  })

  it("streamXlsxRows() without a password throws EncryptedFileError", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    await expect(async () => {
      for await (const _ of streamXlsxRows(enc)) void _
    }).rejects.toBeInstanceOf(EncryptedFileError)
  })
})

describe("roundtrip open → save with encryption", () => {
  it("opens an encrypted workbook and re-saves it encrypted", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "first", spinCount: 64 } })
    const wb = await openXlsx(enc, { password: "first" })
    expect(wb.sheets[0].rows[1]).toEqual(["Ada", 95])

    const resaved = await saveXlsx(wb, { encryption: { password: "second", spinCount: 64 } })
    expect(isOle2Container(resaved)).toBe(true)

    const reopened = await readXlsx(resaved, { password: "second" })
    expect(reopened.sheets[0].rows[2]).toEqual(["Linus", 88])
    // old password no longer works on the re-encrypted file
    await expect(readXlsx(resaved, { password: "first" })).rejects.toBeInstanceOf(DecryptionError)
  })

  it("saveXlsx without an encryption option produces a plain ZIP", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    const wb = await openXlsx(enc, { password: "pw" })
    const plain = await saveXlsx(wb)
    expect(isOle2Container(plain)).toBe(false)
    expect((await readXlsx(plain)).sheets[0].rows[1]).toEqual(["Ada", 95])
  })
})

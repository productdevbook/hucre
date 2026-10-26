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

  // The header names 109 FAT sectors, enough for 109 × 128 sectors of 512
  // bytes — just under 7 MB. A larger container lists the rest of its FAT
  // in DIFAT sectors, which the writer never wrote: what it produced read
  // back cut off at 7,143,936 bytes, whatever had gone in.
  it.each([
    ["one DIFAT sector", 8 * 1024 * 1024, 1],
    ["two chained DIFAT sectors", 16 * 1024 * 1024, 2],
  ])("round-trips a stream whose FAT needs %s", (_, size, difatSectors) => {
    const data = patterned(size)
    const file = writeCfb([{ name: "EncryptedPackage", data }])

    const back = readCfb(file).get("EncryptedPackage")!
    expect(back.length).toBe(data.length)
    expect(firstDifference(back, data)).toBe(-1)
    expect(new DataView(file.buffer).getUint32(72, true)).toBe(difatSectors)
  })
})

/** Bytes that differ from sector to sector, so a misordered chain cannot read back equal. */
function patterned(size: number): Uint8Array {
  const out = new Uint8Array(size)
  for (let i = 0; i < size; i++) out[i] = (i * 31 + (i >> 9)) & 0xff
  return out
}

/** The first index at which the two differ, or -1; cheaper than a deep equal at this size. */
function firstDifference(a: Uint8Array, b: Uint8Array): number {
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return i
  return -1
}

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

  // 100,000 spins on each side is 200,000 SHA-512 iterations, and the
  // budget has to cover the worst case rather than the quiet one: vitest
  // runs files in parallel, so this shares a machine with the rest of the
  // suite. 30s was enough in isolation and not under load, so it failed
  // intermittently — on a test whose subject is not speed. Keeping the
  // real default matters: this is the only assertion that Excel's own
  // spin count works end to end.
  it("interoperates at Excel's default spin count", { timeout: 180_000 }, async () => {
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

  // A package past the ~7 MB the CFB header's own FAT list can map: it was
  // written, and reading it back failed with "Not a valid encrypted
  // workbook container". Random text barely compresses, so 80k cells of it
  // make a package of ~8.6 MB.
  it(
    "round-trips a workbook whose encrypted package is over 7 MB",
    { timeout: 30_000 },
    async () => {
      const rows = noiseRows(80_000, 128)
      const enc = await writeXlsx({
        sheets: [{ name: "Sheet1", rows }],
        encryption: { password: "pw", spinCount: 64 },
      })
      expect(enc.length).toBeGreaterThan(7_500_000)

      const wb = await readXlsx(enc, { password: "pw" })
      expect(wb.sheets[0].rows).toHaveLength(rows.length)
      expect(wb.sheets[0].rows.at(-1)).toEqual(rows.at(-1))
    },
  )
})

/** Rows of one cell of pseudo-random text each — deterministic, and nearly incompressible. */
function noiseRows(count: number, width: number): CellValue[][] {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  let seed = 1
  const rows: CellValue[][] = []
  for (let r = 0; r < count; r++) {
    let text = ""
    for (let i = 0; i < width; i++) {
      seed = (Math.imul(seed, 1103515245) + 12345) >>> 0
      text += alphabet[seed >>> 26]
    }
    rows.push([text])
  }
  return rows
}

describe("decryption across the read entry points", () => {
  it("read() auto-detects and decrypts", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    const wb = await read(enc, { password: "pw" })
    expect(wb.sheets[0].rows[2]).toEqual(["Linus", 88])
  })

  it("readObjects() decrypts", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    const { data } = await readObjects<{ Name: string; Score: number }>(enc, { password: "pw" })
    expect(data).toEqual([
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

import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { ZipReader } from "../src/zip/reader"
import { ZipError } from "../src/errors"

const enc = new TextEncoder()

/** Locate the End-Of-Central-Directory record (signature 0x06054b50). */
function findEocd(buf: Uint8Array): number {
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  for (let i = buf.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === 0x06054b50) return i
  }
  throw new Error("EOCD not found")
}

describe("ZipReader — malformed ZIP64 escapes", () => {
  it("throws instead of mis-parsing an escape with no ZIP64 records behind it", async () => {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("hello"))
    const buf = await zip.build()

    // Forge a ZIP64 escape on an archive that has no ZIP64 EOCD record.
    // Well-formed ZIP64 archives are covered in zip64.test.ts.
    const eocd = findEocd(buf)
    const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
    view.setUint16(eocd + 8, 0xffff, true) // total entries on this disk
    view.setUint16(eocd + 10, 0xffff, true) // total entries

    expect(() => new ZipReader(buf)).toThrow(ZipError)
    expect(() => new ZipReader(buf)).toThrow(/ZIP64/)
  })

  it("still reads a normal (non-ZIP64) archive", async () => {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("hello"))
    const buf = await zip.build()
    const reader = new ZipReader(buf)
    expect(new TextDecoder().decode(await reader.extract("a.txt"))).toBe("hello")
  })
})

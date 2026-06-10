import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { ZipReader, ZipError } from "../src/zip/reader"

const enc = new TextEncoder()

/** Locate the End-Of-Central-Directory record (signature 0x06054b50). */
function findEocd(buf: Uint8Array): number {
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  for (let i = buf.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === 0x06054b50) return i
  }
  throw new Error("EOCD not found")
}

describe("ZipReader — ZIP64 sentinel detection", () => {
  it("throws a clear error instead of mis-parsing a ZIP64-escaped entry count", async () => {
    const zip = new ZipWriter()
    zip.add("a.txt", enc.encode("hello"))
    const buf = await zip.build()

    // Forge a ZIP64 escape: set the EOCD entry-count fields to 0xFFFF.
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

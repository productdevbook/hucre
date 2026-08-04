import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { ZipReader } from "../src/zip/reader"
import { sniffImageFormat } from "../src/xlsx/background-image"

// ═══════════════════════════════════════════════════════════════════════
// #427 — a background image was written to xl/media/imageN.png and
// declared image/png whatever it actually was, on both the authoring and
// the round-trip path. Excel sniffs image bytes and renders it anyway,
// which is why it went unnoticed; the package still declared a content
// type that did not match the part it pointed at.
//
// `Sheet.backgroundImage` is a bare Uint8Array and the reader discards
// the source extension, so the bytes are the only thing left to be
// faithful to — hence sniffing rather than a new type field.
// ═══════════════════════════════════════════════════════════════════════

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
])
const JPEG = new Uint8Array([
  0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01,
  0x00, 0x01, 0x00, 0x00, 0xff, 0xd9,
])
const GIF = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00])
const WEBP = new Uint8Array([
  0x52, 0x49, 0x46, 0x46, 0x24, 0x00, 0x00, 0x00, 0x57, 0x45, 0x42, 0x50, 0x56, 0x50, 0x38, 0x20,
])
const SVG = new TextEncoder().encode(
  '<?xml version="1.0"?>\n<svg xmlns="http://www.w3.org/2000/svg"/>',
)

const media = (zip: ZipReader): string[] => zip.entries().filter((e) => e.startsWith("xl/media/"))

const defaults = async (zip: ZipReader): Promise<string> =>
  new TextDecoder().decode(await zip.extract("[Content_Types].xml"))

describe("sniffImageFormat", () => {
  const cases: Array<[string, Uint8Array, string]> = [
    ["PNG", PNG, "png"],
    ["JPEG", JPEG, "jpeg"],
    ["GIF", GIF, "gif"],
    ["WebP", WEBP, "webp"],
    ["SVG behind an XML declaration", SVG, "svg"],
    ["SVG with no declaration", new TextEncoder().encode("<svg viewBox='0 0 1 1'/>"), "svg"],
  ]

  for (const [label, bytes, expected] of cases) {
    it(`identifies ${label}`, () => {
      expect(sniffImageFormat(bytes)).toBe(expected)
    })
  }

  it("falls back to png for bytes it does not recognise", () => {
    // The same thing the code did unconditionally before, so an exotic
    // format is no worse off than it already was.
    expect(sniffImageFormat(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))).toBe("png")
    expect(sniffImageFormat(new Uint8Array(0))).toBe("png")
  })

  it("does not mistake a binary file that merely contains '<svg' for SVG", () => {
    const data = new Uint8Array(2048)
    data.set(new TextEncoder().encode("<svg "), 1500)
    expect(sniffImageFormat(data)).toBe("png")
  })

  it("does not read past the sniff window looking for it", () => {
    // A 5 MB blob must not be scanned end to end on the off-chance.
    const big = new Uint8Array(5 * 1024 * 1024)
    big.set(JPEG, 0)
    expect(sniffImageFormat(big)).toBe("jpeg")
  })
})

describe("writeXlsx names the media part after the real format", () => {
  const cases: Array<[string, Uint8Array, string, string]> = [
    ["png", PNG, "xl/media/image1.png", "image/png"],
    ["jpeg", JPEG, "xl/media/image1.jpeg", "image/jpeg"],
    ["gif", GIF, "xl/media/image1.gif", "image/gif"],
    ["webp", WEBP, "xl/media/image1.webp", "image/webp"],
    ["svg", SVG, "xl/media/image1.svg", "image/svg+xml"],
  ]

  for (const [label, bytes, path, contentType] of cases) {
    it(`stores a ${label} background as ${path}`, async () => {
      const buf = await writeXlsx({
        sheets: [{ name: "S", rows: [["a"]], backgroundImage: bytes }],
      })
      const zip = new ZipReader(buf)
      expect(media(zip)).toEqual([path])
      expect(await defaults(zip)).toContain(`ContentType="${contentType}"`)
      expect([...(await zip.extract(path))]).toEqual([...bytes])
    })
  }

  it("points the picture relationship at the part it wrote", async () => {
    const buf = await writeXlsx({
      sheets: [{ name: "S", rows: [["a"]], backgroundImage: JPEG }],
    })
    const zip = new ZipReader(buf)
    const rels = new TextDecoder().decode(await zip.extract("xl/worksheets/_rels/sheet1.xml.rels"))
    expect(rels).toContain("../media/image1.jpeg")
    expect(rels).not.toContain(".png")
  })
})

describe("saveXlsx keeps the format across a round trip", () => {
  it("does not turn a JPEG into a part named .png", async () => {
    const original = await writeXlsx({
      sheets: [{ name: "S", rows: [["a"]], backgroundImage: JPEG }],
    })
    const saved = await saveXlsx(await openXlsx(original))
    const zip = new ZipReader(saved)

    expect(media(zip)).toEqual(["xl/media/image1.jpeg"])
    expect(await defaults(zip)).toContain('ContentType="image/jpeg"')

    const back = await readXlsx(saved)
    expect([...(back.sheets[0].backgroundImage ?? [])]).toEqual([...JPEG])
  })

  it("keeps two sheets' backgrounds apart when their formats differ", async () => {
    const original = await writeXlsx({
      sheets: [
        { name: "One", rows: [["a"]], backgroundImage: PNG },
        { name: "Two", rows: [["b"]], backgroundImage: GIF },
      ],
    })
    const saved = await saveXlsx(await openXlsx(original))
    const zip = new ZipReader(saved)

    expect(media(zip).sort()).toEqual(["xl/media/image1.png", "xl/media/image2.gif"].sort())
    const ct = await defaults(zip)
    expect(ct).toContain('ContentType="image/png"')
    expect(ct).toContain('ContentType="image/gif"')

    const back = await readXlsx(saved)
    expect([...(back.sheets[0].backgroundImage ?? [])]).toEqual([...PNG])
    expect([...(back.sheets[1].backgroundImage ?? [])]).toEqual([...GIF])
  })

  it("still shares the media counter with drawing images", async () => {
    // Both land in xl/media, so numbering them independently would
    // collide and one would overwrite the other.
    const original = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"]],
          images: [{ data: PNG, type: "png", anchor: { from: { row: 0, col: 0 } } }],
          backgroundImage: JPEG,
        },
      ],
    })
    const saved = await saveXlsx(await openXlsx(original))
    expect(new Set(media(new ZipReader(saved))).size).toBe(2)
  })
})

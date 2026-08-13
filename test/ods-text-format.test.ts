import { describe, expect, it } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { ZipReader } from "../src/zip/reader"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// `@` is Excel's text format: display this cell as text, whatever it
// holds. #535 recorded it as a loss ODS could not carry —
//
//   | `@` | _(General)_ | the text format has no data style to write |
//
// — which was wrong, and wrong in the way that is hardest to notice: it
// was written down, so nobody looked again. ODF has
// `<number:text-style>`, and LibreOffice writes one into every document
// it saves.
//
// Found by `scripts/spec-coverage.mjs` after #547 crossed the ODF half
// with the corpus: `number:text-style` is in the grammar, in
// `libreoffice-basic.ods`, and was nowhere in `src/`.
// ═══════════════════════════════════════════════════════════════════════

const dec = new TextDecoder()

async function bytesFor(numFmt: string): Promise<Uint8Array> {
  return writeOds({
    sheets: [
      {
        name: "S",
        rows: [["hello"]],
        cells: new Map([["0,0", { value: "hello", style: { numFmt } as CellStyle }]]),
      },
    ],
  })
}

async function roundTrip(numFmt: string): Promise<string | undefined> {
  const wb = await readOds(await bytesFor(numFmt), { readStyles: true })
  return wb.sheets[0]!.cells?.get("0,0")?.style?.numFmt
}

describe("the text format survives ODS", () => {
  it("round-trips as @", async () => {
    expect(await roundTrip("@")).toBe("@")
  })

  it("as ODF's own element, so LibreOffice shows it too", async () => {
    const xml = dec.decode(await new ZipReader(await bytesFor("@")).extract("content.xml"))

    expect(xml).toContain("<number:text-style")
    expect(xml).toContain("<number:text-content/>")
  })

  it("and the cell keeps its value", async () => {
    const wb = await readOds(await bytesFor("@"))

    expect(wb.sheets[0]!.rows[0]![0]).toBe("hello")
  })
})

describe("what General does is unchanged", () => {
  it("General still writes no data style", async () => {
    // `General` is the *absence* of a format. It must not become a text
    // style just because `@` now does.
    const xml = dec.decode(await new ZipReader(await bytesFor("General")).extract("content.xml"))

    expect(xml).not.toContain("<number:text-style")
    expect(await roundTrip("General")).toBeUndefined()
  })

  it("and the ordinary formats are untouched", async () => {
    for (const numFmt of ["0.00", "#,##0", "0%", "yyyy-mm-dd", "hh:mm:ss", '"$"#,##0.00']) {
      expect(await roundTrip(numFmt), numFmt).toBe(numFmt)
    }
  })
})

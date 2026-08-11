import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { HucreError, ParseError, ZipError } from "../src/errors"

// ═══════════════════════════════════════════════════════════════════════
// #473 — a corpus of deliberately-broken files, each with one assertion:
// *this throws a typed error*, or *this returns a partial result*.
//
// Built in code rather than committed as binaries. A directory of opaque
// .xlsx files tells a reader nothing about what is wrong with each one,
// and cannot be adjusted when a part's layout changes. Every entry below
// says how it is broken in the line that breaks it.
//
// This is where the known cases belong — they were scattered through
// `edge-cases*.test.ts` — and the boundary it draws is the one that
// matters: **short is fine, wrong is not.** A file missing a shared
// string reads as empty; a file missing `xl/workbook.xml` is not a
// workbook and says so.
// ═══════════════════════════════════════════════════════════════════════

const enc = new TextEncoder()
const dec = new TextDecoder()

async function clean(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Data",
        rows: [
          ["name", "qty"],
          ["Widget", 3],
        ],
      },
    ],
  })
}

/** Rebuild the archive with parts rewritten, replaced or removed. */
async function rebuild(
  bytes: Uint8Array,
  edits: Record<string, ((xml: string) => string) | null>,
): Promise<Uint8Array> {
  const all = await new ZipReader(bytes).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    const edit = edits[name]
    if (edit === null) continue // removed
    zw.add(name, edit ? enc.encode(edit(dec.decode(data))) : data)
  }
  return zw.build()
}

// ── Not a workbook at all ────────────────────────────────────────────

describe("structural damage throws, because the answer would be wrong", () => {
  it("empty input", async () => {
    await expect(readXlsx(new Uint8Array(0))).rejects.toThrow(HucreError)
  })

  it("not a ZIP", async () => {
    await expect(readXlsx(enc.encode("this is just text, at some length"))).rejects.toThrow(
      HucreError,
    )
  })

  it("a ZIP with no end-of-central-directory record", async () => {
    const bytes = await clean()
    await expect(readXlsx(bytes.subarray(0, bytes.length - 40))).rejects.toThrow(ZipError)
  })

  it("a valid ZIP that is not a workbook", async () => {
    const zw = new ZipWriter()
    zw.add("hello.txt", enc.encode("hi"))

    await expect(readXlsx(await zw.build())).rejects.toThrow(HucreError)
  })

  it("no xl/workbook.xml", async () => {
    await expect(
      readXlsx(await rebuild(await clean(), { "xl/workbook.xml": null })),
    ).rejects.toThrow(HucreError)
  })

  it("a worksheet part the workbook declares and the archive does not have", async () => {
    await expect(
      readXlsx(await rebuild(await clean(), { "xl/worksheets/sheet1.xml": null })),
    ).rejects.toThrow(HucreError)
  })

  it("workbook.xml that is not XML", async () => {
    await expect(
      readXlsx(await rebuild(await clean(), { "xl/workbook.xml": () => "<<<<not xml" })),
    ).rejects.toThrow(HucreError)
  })

  it("corrupt compressed data", async () => {
    // The case the fuzzer found: this used to throw a bare `Error`, so a
    // caller catching HucreError missed it. See #473.
    const bytes = await clean()
    const copy = new Uint8Array(bytes)
    // Byte 87 is inside the first entry's DEFLATE stream.
    copy[87] ^= 0xff

    await expect(readXlsx(copy)).rejects.toThrow(HucreError)
  })
})

// ── Damaged but readable ─────────────────────────────────────────────

describe("content damage reads short, because a partial answer is useful", () => {
  it("a cell pointing at a shared string that is not there", async () => {
    const wb = await readXlsx(
      await rebuild(await clean(), {
        "xl/worksheets/sheet1.xml": (x) => x.replace("<v>0</v>", "<v>9999</v>"),
      }),
    )

    expect(wb.sheets[0]!.rows[0]![0]).toBeNull()
    expect(wb.sheets[0]!.rows[1]![1]).toBe(3)
  })

  it("no sharedStrings.xml at all", async () => {
    const wb = await readXlsx(await rebuild(await clean(), { "xl/sharedStrings.xml": null }))

    // Numbers live in the sheet and survive; text lived in the part that
    // is gone.
    expect(wb.sheets[0]!.rows[1]![1]).toBe(3)
  })

  it("no styles.xml", async () => {
    const wb = await readXlsx(await rebuild(await clean(), { "xl/styles.xml": null }), {
      readStyles: true,
    })

    expect(wb.sheets[0]!.rows[1]![1]).toBe(3)
  })

  it("a cell reference with no row number", async () => {
    const wb = await readXlsx(
      await rebuild(await clean(), {
        "xl/worksheets/sheet1.xml": (x) => x.replace(/ r="B2"/, ' r="B"'),
      }),
    )

    expect(wb.sheets).toHaveLength(1)
  })

  it("a dimension that lies about the sheet's size", async () => {
    // `A1:XFD1048576` is 17 billion slots. The reader uses the cells it
    // finds, not the claim.
    const wb = await readXlsx(
      await rebuild(await clean(), {
        "xl/worksheets/sheet1.xml": (x) =>
          x.replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:XFD1048576"/>'),
      }),
    )

    expect(wb.sheets[0]!.rows.length).toBeLessThan(10)
  })

  it("sheetData that ends mid-element", async () => {
    const wb = await readXlsx(
      await rebuild(await clean(), {
        "xl/worksheets/sheet1.xml": (x) => x.replace("</sheetData>", ""),
      }),
    )

    expect(wb.sheets[0]!.rows[0]![0]).toBe("name")
  })

  it("an unknown element in the middle of the sheet", async () => {
    const wb = await readXlsx(
      await rebuild(await clean(), {
        "xl/worksheets/sheet1.xml": (x) =>
          x.replace("<sheetData>", "<sheetData><wat><nested/></wat>"),
      }),
    )

    expect(wb.sheets[0]!.rows[0]![0]).toBe("name")
  })

  it("a style index past the end of the table", async () => {
    const wb = await readXlsx(
      await rebuild(await clean(), {
        "xl/worksheets/sheet1.xml": (x) => x.replace(/ s="\d+"/g, ' s="9999"'),
      }),
      { readStyles: true },
    )

    expect(wb.sheets[0]!.rows[1]![1]).toBe(3)
    expect(wb.sheets[0]!.cells?.get("1,1")?.style).toBeUndefined()
  })

  it("a merge range whose end is before its start", async () => {
    const wb = await readXlsx(
      await rebuild(await clean(), {
        "xl/worksheets/sheet1.xml": (x) =>
          x.replace(
            "<sheetData>",
            '<mergeCells count="1"><mergeCell ref="D4:A1"/></mergeCells><sheetData>',
          ),
      }),
    )

    expect(wb.sheets[0]!.rows[0]![0]).toBe("name")
  })
})

// ── The line between them ────────────────────────────────────────────

describe("the boundary is short-versus-wrong", () => {
  it("a missing part that only holds content reads short", async () => {
    // Nothing here throws, and that is the decision: half a sheet is
    // usually more useful than an exception.
    const wb = await readXlsx(await rebuild(await clean(), { "xl/sharedStrings.xml": null }))

    expect(wb.sheets).toHaveLength(1)
  })

  it("a missing part that holds structure throws", async () => {
    // Without workbook.xml there is no sheet list, so a "partial" answer
    // would be a fabricated one.
    await expect(
      readXlsx(await rebuild(await clean(), { "xl/workbook.xml": null })),
    ).rejects.toThrow(ParseError)
  })
})

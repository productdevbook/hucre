import { cellError } from "../src/cell-error"
import { describe, expect, it } from "vitest"
import { readFileSync, readdirSync } from "node:fs"
import { readOds } from "../src/ods/reader"
import { streamOdsRows } from "../src/ods/stream"
import { ZipReader } from "../src/zip/reader"
import type { CellValue } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #464 — the XLSX side got a third-party corpus (Excel, openpyxl,
// ExcelJS) and the **ODS reader got none**. Every ODS byte the suite
// parsed was one hucre had just produced: a reader that misreads
// `office:value-type` checked against a writer that misreads it the same
// way, green forever.
//
// Most are written by **SheetJS** (`xlsx`, Apache-2.0), an independent
// implementation with its own element order, style names, and idea of a
// minimal ODF document. The remaining fixture is LibreOffice's conversion
// of the Excel-authored basic workbook. It covers the two shapes SheetJS
// will not emit:
//
//   * `table:number-columns-repeated`, which LibreOffice uses for every
//     run of like cells and is the sharpest trap in the format.
//   * error cells — SheetJS writes an error as an empty
//     `<table:table-cell/>`, so there is no error in the file to read.
//
// The repeated tail found a real defect: five populated cells became
// 16,384 because the batch reader treated an unresolved style name as
// data even when styles were not requested.
//
// See test/fixtures/third-party/README.md for provenance and licensing.
// ═══════════════════════════════════════════════════════════════════════

const DIR = new URL("./fixtures/third-party/", import.meta.url)

function load(name: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(name, DIR)))
}

const NAMES = readdirSync(DIR)
  .filter((f: string) => f.endsWith(".ods"))
  .sort()

async function rowsOf(name: string, sheet = 0): Promise<CellValue[][]> {
  const wb = await readOds(load(name))
  return wb.sheets[sheet]!.rows
}

function trimTrailingNulls<T>(row: T[]): T[] {
  let end = row.length
  while (end > 0 && row[end - 1] === null) end--
  return row.slice(0, end)
}

describe("the corpus is there and is not ours", () => {
  it("has every fixture the generator writes", () => {
    expect(NAMES).toEqual([
      "libreoffice-basic.ods",
      "sheetjs-basic.ods",
      "sheetjs-dates.ods",
      "sheetjs-empty.ods",
      "sheetjs-formulas.ods",
      "sheetjs-multi-sheet.ods",
      "sheetjs-numbers.ods",
      "sheetjs-sparse.ods",
      "sheetjs-unicode.ods",
      "sheetjs-whitespace.ods",
    ])
  })

  it("was written by something other than hucre", async () => {
    // The load-bearing assertion of this file. If these are ever
    // regenerated with hucre, every test below becomes a round trip
    // against itself and proves nothing — this is the line that notices.
    const meta = new TextDecoder().decode(
      await new ZipReader(load("sheetjs-basic.ods")).extract("meta.xml"),
    )

    expect(meta).toContain("<meta:generator>SheetJS")
    expect(meta).not.toContain("hucre")
  })

  it("does not put mimetype first, and hucre reads it anyway", async () => {
    // ODF §3.3 requires `mimetype` to be the archive's first entry and
    // stored uncompressed. hucre's own writer honours that. SheetJS does
    // not — it lands third — and a reader that enforced the rule would
    // reject a file every other tool opens. Worth pinning: the leniency
    // is deliberate, not an accident nobody noticed.
    const names = new ZipReader(load("sheetjs-basic.ods")).entries()

    expect(names[0]).not.toBe("mimetype")
    expect(names).toContain("mimetype")
    expect((await readOds(load("sheetjs-basic.ods"))).sheets).toHaveLength(1)
  })

  it("ships a part hucre never writes", async () => {
    // `manifest.rdf` is SheetJS's, not ours. An unknown part must be
    // ignored rather than tripping anything.
    const names = new ZipReader(load("sheetjs-basic.ods")).entries()

    expect(names).toContain("manifest.rdf")
  })
})

describe("LibreOffice", () => {
  it("is the producer, and writes the repeated default-style tail", async () => {
    const zip = new ZipReader(load("libreoffice-basic.ods"))
    const meta = new TextDecoder().decode(await zip.extract("meta.xml"))
    const content = new TextDecoder().decode(await zip.extract("content.xml"))

    expect(meta).toContain("<meta:generator>LibreOffice")
    expect(meta).not.toContain("hucre")
    expect(content).toContain(
      '<table:table-cell table:style-name="Default" table:number-columns-repeated="16379"/>',
    )
  })

  it.each([false, true])(
    "does not turn five cells into 16,384 when readStyles is %s",
    async (readStyles) => {
      const rows = (await readOds(load("libreoffice-basic.ods"), { readStyles })).sheets[0]!.rows

      expect(rows.map((row) => row.length)).toEqual([5, 5, 5, 5, 5])
      expect(rows[4]!.slice(0, 3)).toEqual(["Broken", cellError("#DIV/0!"), "xy"])
    },
  )
})

describe("the four ODF value types", () => {
  it("come back as the right JavaScript types", async () => {
    const rows = await rowsOf("sheetjs-basic.ods")

    expect(rows[0]).toEqual(["text", "number", "date", "boolean"])
    expect(rows[1]![0]).toBe("Widget")
    expect(rows[1]![1]).toBe(42)
    expect(rows[1]![2]).toBeInstanceOf(Date)
    expect(rows[1]![3]).toBe(true)
    expect(rows[2]![1]).toBe(-7.25)
    expect(rows[2]![3]).toBe(false)
  })
})

describe("dates", () => {
  it("survive SheetJS's zone designator", async () => {
    // SheetJS writes `office:date-value="2024-03-17T00:00:00.000Z"`.
    // LibreOffice writes `2024-03-17`. The trailing `Z` is SheetJS's own
    // habit and a reader anchored on the other spelling drops all five.
    const rows = await rowsOf("sheetjs-dates.ods")

    expect(rows.map((r) => (r[0] as Date).toISOString())).toEqual([
      "2024-03-17T00:00:00.000Z",
      "2024-03-17T13:45:30.000Z",
      "1899-12-30T00:00:00.000Z",
      "1900-01-01T00:00:00.000Z",
      "2000-02-29T00:00:00.000Z",
    ])
  })

  it("including the serial-0 epoch and the 1900 leap-year trap", async () => {
    // 1899-12-30 is serial 0 and 2000-02-29 is a real date that the
    // 1900-is-a-leap-year fiction gets wrong by a day. ODF carries the
    // literal date, so a reader doing serial arithmetic of its own is
    // the only way to be wrong here — which is the #415 family.
    const rows = await rowsOf("sheetjs-dates.ods")

    expect((rows[2]![0] as Date).getTime()).toBe(Date.UTC(1899, 11, 30))
    expect((rows[4]![0] as Date).getTime()).toBe(Date.UTC(2000, 1, 29))
  })
})

describe("strings", () => {
  it("keep their whitespace without needing an attribute", async () => {
    // The #441 shape in ODF's spelling. XLSX needs `xml:space="preserve"`
    // and a writer can forget it; ODF keeps the spaces by the format, so
    // a reader that trims is wrong with nothing to blame.
    const rows = await rowsOf("sheetjs-whitespace.ods")

    expect(rows.map((r) => r[0])).toEqual([
      "  leading",
      "trailing  ",
      "  both  ",
      "inner   gap",
      "line\nbreak",
    ])
  })

  it("carry unicode through a second encoder intact", async () => {
    const rows = await rowsOf("sheetjs-unicode.ods")

    expect(rows.map((r) => r[0])).toEqual([
      "ünïcödé",
      "日本語のテキスト",
      "🎉 emoji 🚀",
      "Ω≈ç√∫˜µ",
      "e\u0301 combining",
      "\u200bzero width",
    ])
  })
})

describe("numbers", () => {
  it("parse from the decimal string without losing a bit", async () => {
    // `office:value` is text, so the reader's parse is the only thing
    // between the file and the double. The subnormal end of this list is
    // where #485 found the CSV writer flattening values to zero.
    const rows = await rowsOf("sheetjs-numbers.ods")

    expect(rows.map((r) => r[0])).toEqual([
      0.1 + 0.2,
      1e21,
      1e-7,
      Number.MAX_SAFE_INTEGER,
      -0.000001,
      12345678.9,
    ])
  })
})

describe("holes", () => {
  it("read as null rather than shifting the row left", async () => {
    // SheetJS spells a gap `<table:table-cell/>`. A reader that skips
    // those instead of counting them puts "far" in column B.
    const rows = await rowsOf("sheetjs-sparse.ods")

    expect(rows[0]).toHaveLength(11)
    expect(rows[0]![0]).toBe("a")
    expect(rows[0]![10]).toBe("far")
    expect(rows[0]!.slice(1, 10).every((v) => v === null)).toBe(true)
    expect(rows[1]!.slice(0, 3)).toEqual([1, null, 2])
    expect(rows[1]).toHaveLength(11)
    expect(rows[3]![3]).toBe("island")
  })

  it("keep an entirely empty row in place", async () => {
    const rows = await rowsOf("sheetjs-sparse.ods")

    expect(rows[2]).toHaveLength(11)
    expect(rows[2]!.every((v) => v === null)).toBe(true)
    expect(rows).toHaveLength(4)
  })
})

describe("formulas", () => {
  it("arrive as their cached value", async () => {
    // ODF formulas are `of:=[.A2]*2`, nothing like XLSX's `A2*2`. hucre
    // has no formula engine, so what has to survive is the cached result.
    const rows = await rowsOf("sheetjs-formulas.ods")

    expect(rows[1]).toEqual([21, 42])
    expect(rows[2]).toEqual([50, 100])
  })
})

describe("the degenerate and the plural", () => {
  it("an empty sheet is an empty sheet, not a throw", async () => {
    const wb = await readOds(load("sheetjs-empty.ods"))

    expect(wb.sheets).toHaveLength(1)
    expect(wb.sheets[0]!.name).toBe("Empty")
    expect(wb.sheets[0]!.rows).toEqual([])
  })

  it("three sheets keep their order and their names", async () => {
    const wb = await readOds(load("sheetjs-multi-sheet.ods"))

    expect(wb.sheets.map((s) => s.name)).toEqual(["First", "İkinci Sayfa", "Third & Last"])
    expect(wb.sheets[1]!.rows[1]).toEqual([2])
    expect(wb.sheets[2]!.rows[1]).toEqual([3])
  })
})

describe("the streaming reader sees the same thing", () => {
  it("agrees with the buffered reader on every row of every fixture", async () => {
    // Two independent parsers over the same bytes. They have disagreed
    // before, and only a file neither of them wrote can show it.
    //
    // The comparison is by `index`, not by position, because the two do
    // not emit the same *number* of rows: `streamOdsRows` skips an
    // entirely empty row and advances the index past it, where `readOds`
    // materialises it into the grid. `sheetjs-sparse.ods` has one in the
    // middle, so it comes out of the stream as rows 0, 1, 3.
    //
    // That is the streaming contract rather than a defect — a sheet
    // whose `table:number-rows-repeated` covers a million blank rows must
    // not yield a million rows — and it holds only because the index is
    // still right. This is what checks that it is.
    for (const name of NAMES) {
      const sheets = (await readOds(load(name))).sheets
      const seen = sheets.map(() => new Set<number>())

      for await (const row of streamOdsRows(load(name), { sheet: "all" })) {
        const at = row.sheet
        // The buffered reader pads every row to the sheet's width; the
        // streaming one cannot know the width yet, so compare up to the
        // last value.
        expect(row.values, `${name} sheet ${at} row ${row.index}`).toEqual(
          trimTrailingNulls(sheets[at]!.rows[row.index]!),
        )
        seen[at]!.add(row.index)
      }

      // Everything the stream left out has to have been empty.
      sheets.forEach((sheet, at) => {
        sheet.rows.forEach((values: CellValue[], index: number) => {
          if (seen[at]!.has(index)) return
          expect(trimTrailingNulls(values), `${name} sheet ${at} row ${index} was skipped`).toEqual(
            [],
          )
        })
      })
    }
  })
})

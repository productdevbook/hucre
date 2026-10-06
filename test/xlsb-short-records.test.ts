import { cellError } from "../src/cell-error"
import { describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import { readXlsb } from "../src/xlsx/xlsb/reader"
import { readXls } from "../src/xls/reader"
import type { CellValue, Workbook } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// hucre's XLSB reader handled the full-form cell records — `BrtCellRk`,
// `BrtCellSt`, `BrtCellReal`, `BrtCellBool`, `BrtCellError`,
// `BrtCellIsst`, `BrtCellBlank` — and **none of the `BrtShort*` forms**.
//
// They are the same records **without the column**. [MS-XLSB] §2.5.9
// `Cell` is a 4-byte column then a 4-byte style field; §2.5.10
// `ShortCell` is the style alone, and the column is the previous cell's
// plus one. Excel writes the full form every time, so a corpus of Excel
// output cannot see this; SheetJS writes the short form for every cell
// after the first in a row.
//
// Which field is missing is not something the byte lengths settle — a
// `BrtShortSt` for "col2" is 16 bytes against 20 for a `BrtCellSt` for
// "col1", and that fits either reading. Reading it as "the style is
// missing" put every cell of a row at column 0, which is how I know.
//
// The result was silent and severe: a twelve-column sheet read back one
// column wide. No error, no warning, and every cell after the first of
// each row simply absent.
//
//     sheetjs-wide.xls    col1 | col2 | col3 | … | col12
//     sheetjs-wide.xlsb   col1
//
// `test/fixtures/PROVENANCE.md` called this out before it happened: "The
// XLS and XLSB readers are the sharp end — they exist only to consume
// other tools' output and, until this directory, had never seen any."
// The directory fixed half of it; openpyxl writes `.xlsx` only, so the
// binary readers still had exactly one source until these fixtures.
// ═══════════════════════════════════════════════════════════════════════

const DIR = new URL("./fixtures/third-party/", import.meta.url)

function load(name: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(name, DIR)))
}

const STEMS = [
  "sheetjs-values",
  "sheetjs-wide",
  "sheetjs-unicode",
  "sheetjs-dates",
  "sheetjs-sparse",
]

/** Dates and numbers compared on the same footing. */
function shape(wb: Workbook): unknown {
  return wb.sheets.map((s) => ({
    name: s.name,
    rows: s.rows.map((row: CellValue[]) =>
      row.map((v) => (v instanceof Date ? `D:${v.toISOString()}` : v)),
    ),
  }))
}

describe("the two binary readers agree on the same workbook", () => {
  // The strongest assertion available: one logical workbook, written by
  // one producer into two formats, read by two hucre readers that share
  // no code. Neither can be checked against the other's mistakes.
  for (const stem of STEMS) {
    it(`${stem}: xls and xlsb read the same thing`, async () => {
      const fromXls = shape(await readXls(load(`${stem}.xls`)))
      const fromXlsb = shape(await readXlsb(load(`${stem}.xlsb`)))

      expect(fromXlsb).toEqual(fromXls)
    })
  }
})

describe("a cell written with the short record form is still a cell", () => {
  it("reads every column of a wide sheet, not just the first", async () => {
    const wb = await readXlsb(load("sheetjs-wide.xlsb"))
    const rows = wb.sheets[0]!.rows

    expect(rows[0]).toHaveLength(12)
    expect(rows[0]![0]).toBe("col1")
    expect(rows[0]![11]).toBe("col12")
    expect(rows[1]![0]).toBe("r0c0")
    expect(rows[1]![1]).toBe(1)
    expect(rows[1]![11]).toBe(11)
  })

  it("carries each cell type through its short form", async () => {
    // BrtShortSt, BrtShortRk, BrtShortReal, BrtShortBool, BrtShortBlank
    // — one row exercises all five.
    const rows = (await readXlsb(load("sheetjs-values.xlsb"))).sheets[0]!.rows

    expect(rows[0]).toEqual(["text", "int", "float", "bool", "blank"])
    expect(rows[1]![0]).toBe("Widget")
    expect(rows[1]![1]).toBe(42)
    expect(rows[1]![2]).toBe(-7.25)
    expect(rows[1]![3]).toBe(true)
    expect(rows[2]![1]).toBe(0)
    expect(rows[2]![2]).toBe(0.1 + 0.2)
    expect(rows[2]![3]).toBe(false)
  })

  it("including a float that does not fit an RK", async () => {
    // `1e-7` cannot be an RK, so it is a `BrtShortReal` — a different
    // payload from the integer beside it.
    const rows = (await readXlsb(load("sheetjs-values.xlsb"))).sheets[0]!.rows

    expect(rows[3]![1]).toBe(1000000)
    expect(rows[3]![2]).toBe(1e-7)
  })
})

describe("strings that live in the shared table", () => {
  // Every other fixture here has SheetJS's default: strings written
  // inline, so `BrtCellIsst` and `BrtShortIsst` — the records that carry
  // an *index* into `sharedStrings.bin` rather than the text — were the
  // pair no real file exercised. This one is generated with `bookSST`,
  // and its sheet uses record ids 7 and 18 for the strings and 3 and 14
  // for the errors: the four the short-record fix touched that nothing
  // else reaches.
  it("resolve through the shared table, short form included", async () => {
    const rows = (await readXlsb(load("sheetjs-shared-strings.xlsb"))).sheets[0]!.rows

    expect(rows[0]).toEqual(["alpha", "beta", "gamma"])
    expect(rows[1]).toEqual(["alpha", "delta", "beta"])
    expect(rows[2]).toEqual(["gamma", "alpha", "delta"])
  })

  it("and an error cell keeps its code", async () => {
    const rows = (await readXlsb(load("sheetjs-shared-strings.xlsb"))).sheets[0]!.rows

    expect(rows[3]).toEqual([cellError("#DIV/0!"), cellError("#N/A"), cellError("#REF!")])
  })
})

describe("what already worked still does", () => {
  it("unicode, including astral characters", async () => {
    const rows = (await readXlsb(load("sheetjs-unicode.xlsb"))).sheets[0]!.rows

    expect(rows.map((r) => r[0])).toEqual([
      "ünïcödé",
      "日本語のテキスト",
      "🎉 emoji 🚀",
      "  padded  ",
      'quote"inside',
    ])
  })

  it("a styled cell, which uses the full record form", async () => {
    // Dates carry a number format, so SheetJS writes them as `BrtCellRk`
    // with a style — the path that always worked. This is what notices
    // if the fix is applied to the wrong records.
    //
    // The *instant* is not asserted. SheetJS converts a `Date` to a
    // serial through local time, so these bytes carry whatever offset the
    // machine that generated them was on; hucre reads the serial it was
    // given, faithfully, and both readers agree. Pinning an instant here
    // would pin my timezone into the suite.
    const rows = (await readXlsb(load("sheetjs-dates.xlsb"))).sheets[0]!.rows
    const fromXls = (await readXls(load("sheetjs-dates.xls"))).sheets[0]!.rows

    expect(rows[0]![0]).toBeInstanceOf(Date)
    expect(rows).toHaveLength(3)
    for (let i = 0; i < rows.length; i++) {
      expect((rows[i]![0] as Date).getTime()).toBe((fromXls[i]![0] as Date).getTime())
    }
  })

  it("holes stay holes", async () => {
    const rows = (await readXlsb(load("sheetjs-sparse.xlsb"))).sheets[0]!.rows

    expect(rows[0]![0]).toBe("a")
    expect(rows[0]![4]).toBe("far")
    expect(rows[0]![1]).toBeNull()
    expect(rows[3]![3]).toBe("island")
  })
})

import { describe, expect, it } from "vitest"
import { fixtureBytes, fixtureJson } from "../support/fixtures"
import { projectWorkbook, type Flat, type WorkbookModel } from "../support/workbook-model"
import { read, readXls, readXlsb, readXlsx, streamXlsxRows } from "../../src/index"
import type { ReadOptions, Workbook } from "../../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #464 — files hucre did not write.
//
// Hand-built inputs alone form a closed loop: a reader that misunderstands
// a record can be tested against a record with the same misunderstanding.
// Independent producer files exercise the reader against another tool's
// interpretation of the format, especially for XLS and XLSB.
//
// So: test/fixtures/ holds fourteen workbooks written by Microsoft Excel
// 16.0. The content is synthetic and authored for this purpose by
// scripts/fixtures/make-fixtures.vbs; see test/fixtures/PROVENANCE.md.
//
// The expected models below are hand-written from what was authored in
// the .vbs and from what the raw file actually contains (unzip the xlsx,
// dump the BIFF records) — NOT from running hucre and pasting the
// result. Pasting hucre's output would rebuild exactly the closed loop
// this file exists to break.
//
// For the same reason the projection is deliberately narrow. It carries
// what a person can check against Excel — values, types, merges, frozen
// panes, hidden rows and columns, number formats, page setup — and drops
// what a person cannot, notably Excel's autofit column widths
// (10.33203125) and its theme palette. Asserting those would be
// asserting that hucre agrees with hucre.
//
// Expected values are independent of hucre. A mismatch must be checked
// against the authored input before changing the golden.
// ═══════════════════════════════════════════════════════════════════════

const bytes = fixtureBytes
const golden = (name: string): WorkbookModel => fixtureJson(`${name}.golden.json`)

const modelOf = async (
  name: string,
  reader: (b: Uint8Array, o?: ReadOptions) => Promise<Workbook>,
): Promise<WorkbookModel> => {
  const warnings: string[] = []
  const wb = await reader(bytes(name), {
    readStyles: true,
    onWarning: (w) => warnings.push(w.message),
  })
  return projectWorkbook(wb, warnings)
}

// ── the corpus ──────────────────────────────────────────────────────

const FIXTURES: Array<{ file: string; reader: typeof readXlsx }> = [
  { file: "excel-basic.xlsx", reader: readXlsx },
  { file: "excel-basic.xls", reader: readXls },
  { file: "excel-basic.xlsb", reader: readXlsb },
  { file: "excel-strings.xlsx", reader: readXlsx },
  { file: "excel-strings.xlsb", reader: readXlsb },
  { file: "excel-styled.xlsx", reader: readXlsx },
  { file: "excel-layout.xlsx", reader: readXlsx },
  { file: "excel-pagesetup.xlsx", reader: readXlsx },
  { file: "excel-styleonly.xlsx", reader: readXlsx },
  { file: "excel-dates.xls", reader: readXls },
  { file: "excel-empty.xlsx", reader: readXlsx },
  { file: "excel-features.xlsx", reader: readXlsx },
  // A second producer. openpyxl is not Excel-with-a-different-icon: it
  // emits formulas with no cached result, ISO-8601 `t="d"` date cells,
  // `date1904`, and inline strings with no shared string table — four
  // shapes Excel never writes, so a corpus of Excel output alone cannot
  // reach them. It writes .xlsx only, hence no .xls/.xlsb siblings.
  { file: "openpyxl-basic.xlsx", reader: readXlsx },
  { file: "openpyxl-formulas.xlsx", reader: readXlsx },
  { file: "openpyxl-isodates.xlsx", reader: readXlsx },
  { file: "openpyxl-1904.xlsx", reader: readXlsx },
  { file: "openpyxl-inline-strings.xlsx", reader: readXlsx },
  { file: "openpyxl-styled.xlsx", reader: readXlsx },
]

describe("workbooks written by Excel, not by this test suite", () => {
  for (const { file, reader } of FIXTURES) {
    describe(file, () => {
      it("matches its hand-written model through the explicit reader", async () => {
        expect(await modelOf(file, reader)).toEqual(golden(file))
      })

      it("reads the same way through read()'s format detection", async () => {
        expect(await modelOf(file, read)).toEqual(golden(file))
      })
    })
  }

  // excel-basic.{xlsx,xls,xlsb} are one authored sheet saved three ways,
  // which is the whole reason all three are in the corpus: whatever the
  // three readers disagree about is a place where the container is
  // leaking into the model. PARITY.md already says XLS and XLSB carry
  // less than XLSX — styles, dimensions, properties — so the comparison
  // is limited to the values, which all three do claim to carry.
  describe("one sheet, three containers", () => {
    const rowsOf = async (file: string, reader: typeof readXlsx): Promise<Flat[][]> =>
      (await modelOf(file, reader)).sheets[0]?.rows ?? []

    it("#494 — reads the same values whichever container Excel saved it in", async () => {
      const xlsx = await rowsOf("excel-basic.xlsx", readXlsx)
      expect(await rowsOf("excel-basic.xls", readXls)).toEqual(xlsx)
      expect(await rowsOf("excel-basic.xlsb", readXlsb)).toEqual(xlsx)
    })

    it("agrees on every cell that all three readers return", async () => {
      const xlsx = await rowsOf("excel-basic.xlsx", readXlsx)
      for (const [file, reader] of [
        ["excel-basic.xls", readXls],
        ["excel-basic.xlsb", readXlsb],
      ] as const) {
        const other = await rowsOf(file, reader)
        expect(other.length).toBe(xlsx.length)
        other.forEach((row, r) => {
          row.forEach((cell, c) => expect([file, r, c, cell]).toEqual([file, r, c, xlsx[r]?.[c]]))
        })
      }
    })
  })

  // Two producers, one authored sheet. This is what a second producer
  // buys that a bigger Excel corpus cannot: if hucre and Excel happened
  // to share a misunderstanding, a corpus made only of Excel output
  // would agree with itself about it. openpyxl read the same spec
  // independently, so where the two files disagree, one of them is
  // wrong and neither gets a free pass.
  //
  // Column E is compared as values: Excel cached its formula results,
  // openpyxl cannot evaluate, so the Python writes the same numbers as
  // literals. The formula-with-no-cached-result case is
  // openpyxl-formulas.xlsx, where it is the subject rather than noise.
  it("agrees between Excel and openpyxl on the same authored sheet", async () => {
    const excel = await modelOf("excel-basic.xlsx", readXlsx)
    const python = await modelOf("openpyxl-basic.xlsx", readXlsx)
    // excel-basic has a fifth row of formula results openpyxl has no
    // way to produce; compare the four rows both files author.
    expect(python.sheets[0]?.rows).toEqual(excel.sheets[0]?.rows.slice(0, 4))
    expect(python.sheets[0]?.name).toBe(excel.sheets[0]?.name)
    expect(python.dateSystem).toBe(excel.dateSystem)
  })

  // A chart sheet — a chart on its own tab — is not a worksheet, and
  // xl/workbook.xml's <sheets> lists it anyway; the relationship type is
  // what separates the kinds.
  //
  // This one was not hypothetical: it was the single largest failure in a
  // corpus of real instrument-exported workbooks, 52 files out of 538.
  // Fixed in #499 — the block below is what it does now.
  describe("excel-chartsheet.xlsx", () => {
    it("#499 — the workbook reads, and the worksheet beside it survives", async () => {
      const wb = await readXlsx(bytes("excel-chartsheet.xlsx"))

      expect(wb.sheets.some((s) => s.rows.length > 0)).toBe(true)
    })

    it("keeps the chart sheet as a tab, so the indices are Excel's", async () => {
      // Skipping it would renumber every sheet after it, which is a
      // quieter kind of wrong than throwing.
      const wb = await readXlsx(bytes("excel-chartsheet.xlsx"))

      expect(wb.sheets).toHaveLength(2)
      expect(wb.sheets.map((s) => s.kind)).toEqual(["chartsheet", undefined])
      expect(wb.sheets[0]?.rows).toEqual([])
    })

    it("streams the worksheet by name", async () => {
      const rows: unknown[][] = []
      for await (const row of streamXlsxRows(bytes("excel-chartsheet.xlsx"), { sheet: 1 })) {
        rows.push(row.values)
      }

      expect(rows.length).toBeGreaterThan(0)
    })

    it("streams a chart sheet as nothing, rather than throwing", async () => {
      const rows: unknown[][] = []
      for await (const row of streamXlsxRows(bytes("excel-chartsheet.xlsx"), { sheet: 0 })) {
        rows.push(row.values)
      }

      expect(rows).toEqual([])
    })
  })

  // Sparse, not large. `Sheet.rows` is a dense rectangle, so the cost of
  // a read is the bounding box and not the cell count — about thirty
  // values placed out to column 15,312 describe 30.6 million slots. The
  // real file behind this had 76,277 values over 507 columns, a
  // 305,612,208-slot box at 0.03% fill, and Excel opens it without
  // complaint.
  //
  // The default read still refuses it, and should: the grid genuinely
  // cannot be built. What #501 changed is that there is now a way out,
  // and that the error names it.
  describe("excel-sparse.xlsx", () => {
    it("#501 — reads with `sparse: true`, which builds no grid", async () => {
      const wb = await readXlsx(bytes("excel-sparse.xlsx"), { sparse: true })

      expect(wb.sheets[0]?.name).toBe("Sparse")
      expect(wb.sheets[0]?.rows).toEqual([])
      expect(wb.sheets[0]?.cells?.size ?? 0).toBeGreaterThan(0)
    })

    it("and streams, which is the other answer", async () => {
      let count = 0
      for await (const _row of streamXlsxRows(bytes("excel-sparse.xlsx"))) count++

      expect(count).toBeGreaterThan(0)
    })

    it("the default read still refuses it, and says how empty the box is", async () => {
      await expect(readXlsx(bytes("excel-sparse.xlsx"))).rejects.toThrow(
        /spans 2000 rows x 15312 columns .* over the \d+ limit/,
      )
      await expect(readXlsx(bytes("excel-sparse.xlsx"))).rejects.toThrow(/% of them filled/)
    })

    it("the error names both ways out", async () => {
      // It used to name three options and none of them worked here.
      await expect(readXlsx(bytes("excel-sparse.xlsx"))).rejects.toThrow(/streamXlsxRows/)
      await expect(readXlsx(bytes("excel-sparse.xlsx"))).rejects.toThrow(/sparse: true/)
    })

    it("can still be read when the caller already knows where the data is", async () => {
      const wb = await readXlsx(bytes("excel-sparse.xlsx"), { range: "A1:C1" })

      expect(wb.sheets[0]?.rows[0]?.[0]).toBe("left edge")
    })
  })

  // A test that reads no bytes would pass just as green. This one fails
  // loudly if test/fixtures/ ever stops being on disk.
  it("is actually reading the files", () => {
    for (const { file } of [
      ...FIXTURES,
      { file: "excel-chartsheet.xlsx" },
      { file: "excel-sparse.xlsx" },
    ]) {
      expect(bytes(file).byteLength).toBeGreaterThan(1000)
    }
  })
})

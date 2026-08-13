import { describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import { read, readXls, readXlsb, readXlsx, streamXlsxRows } from "../src/index"
import type { Cell, CellStyle, ReadOptions, Sheet, Workbook } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #464 — files hucre did not write.
//
// Every other binary input under test/ is assembled byte-by-byte by the
// test that reads it. That is a closed loop: a reader that misunderstands
// a record is tested against a hand-built record that misunderstands it
// the same way, and the suite stays green. The XLS and XLSB readers are
// the sharp end — they exist only to consume other tools' output, and
// until this file they had never seen any.
//
// So: test/fixtures/ holds thirteen workbooks written by Microsoft Excel
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
// Anything here that fails is triaged as defect-vs-wrong-model before it
// is touched. Known defects are marked with `it.fails` and an issue
// number rather than fixed in this file: this is the corpus PR, and each
// fix gets its own change with its own failing-test-first.
// ═══════════════════════════════════════════════════════════════════════

// ── the projection ──────────────────────────────────────────────────

/** A cell value flattened to something a JSON golden can hold. */
type Flat = string | number | boolean | null

interface SheetModel {
  name: string
  /** Dense rows. Dates become `D:<iso>` so they cannot pass as strings. */
  rows: Flat[][]
  /** Merged ranges in A1 notation, sorted. */
  merges: string[]
  /** `"<rows>x<cols>"` frozen, or `null`. */
  freeze: string | null
  /** 1-based row numbers, as Excel shows them. */
  hiddenRows: number[]
  /** Column letters. */
  hiddenColumns: string[]
  /** A1 → compact style description, for cells carrying any style. */
  styles: Record<string, string>
  /** Column letter → compact style description. */
  columnStyles: Record<string, string>
  /** One line per conditional-formatting rule. */
  conditional: string[]
  /** The autofilter range in A1 notation, or `null`. */
  autoFilter: string | null
  /** One line per data-validation rule. */
  validations: string[]
  /** Which protection flags are on, sorted; empty when unprotected. */
  protection: string[]
  /** `A1: text` per cell comment, sorted. */
  comments: string[]
  /**
   * A1 → `<formula text> => <cached result>`, for cells carrying a
   * formula. The cached result is the interesting half: Excel always
   * writes one, openpyxl never does, and `openpyxl-formulas.xlsx` is in
   * the corpus precisely because a formula with no cached value is a
   * shape Excel cannot produce.
   */
  formulas: Record<string, string>
  /** Page setup, minus the margins every sheet gets by default. */
  pageSetup: Record<string, unknown> | null
}

/**
 * A field the corpus proved hucre gets wrong. It is recorded rather than
 * fixed: this is the corpus PR, and every fix lands separately with its
 * own failing test first (CONTRIBUTING.md). `expected` is what the file
 * says it should be, `actual` is what hucre returns today — so the bulk
 * of the fixture keeps asserting, and the one wrong field is both
 * excluded from the green comparison and asserted by an `it.fails`.
 */
interface KnownDefect {
  issue: string
  /** Dotted path into the model, e.g. `sheets.0.rows.7.0`. */
  path: string
  expected: unknown
  actual: unknown
  why: string
}

interface WorkbookModel {
  sheets: SheetModel[]
  dateSystem: string | null
  /**
   * Whether docProps carries an author. The fixtures are generated with
   * `RemovePersonalInformation = True`, so this must stay false — it is
   * the licence-and-privacy guarantee of the corpus, checked by the same
   * suite that reads it rather than only by a note in a markdown file.
   */
  /** `name=range` per defined name, sorted. Excel's built-in print
   * names are excluded — they are page setup, asserted separately. */
  namedRanges: string[]
  hasAuthor: boolean
  warnings: string[]
  knownDefects?: KnownDefect[]
}

const colName = (index: number): string => {
  let n = index + 1
  let out = ""
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - rem) / 26)
  }
  return out
}

const flat = (v: unknown): Flat => {
  if (v instanceof Date) return `D:${v.toISOString()}`
  if (v === undefined || v === null) return null
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v
  return `?:${String(v)}`
}

const colorOf = (c: { rgb?: string; theme?: number; indexed?: number } | undefined): string => {
  if (!c) return "-"
  if (c.rgb !== undefined) return c.rgb
  if (c.theme !== undefined) return `theme${c.theme}`
  if (c.indexed !== undefined) return `indexed${c.indexed}`
  return "-"
}

/**
 * A style as one short line. Excel's default font (Aptos Narrow 11,
 * theme colour 1) is the baseline every cell inherits, so it is elided —
 * what is left is what the fixture actually set.
 */
const styleLine = (s: CellStyle | undefined): string => {
  if (!s) return ""
  const parts: string[] = []
  const f = s.font
  if (f) {
    const bits: string[] = []
    if (f.bold) bits.push("bold")
    if (f.italic) bits.push("italic")
    if (f.underline) bits.push("underline")
    if (f.strikethrough) bits.push("strike")
    if (f.size !== undefined && f.size !== 11) bits.push(`size=${f.size}`)
    if (f.name !== undefined && f.name !== "Aptos Narrow") bits.push(`name=${f.name}`)
    const c = colorOf(f.color)
    if (c !== "theme1" && c !== "-") bits.push(`color=${c}`)
    if (bits.length > 0) parts.push(`font(${bits.join(",")})`)
  }
  if (s.fill?.type === "pattern") parts.push(`fill(${s.fill.pattern},${colorOf(s.fill.fgColor)})`)
  else if (s.fill) parts.push(`fill(gradient)`)
  if (s.border) {
    const sides = (["left", "right", "top", "bottom"] as const)
      .filter((k) => s.border?.[k]?.style)
      .map((k) => `${k}=${s.border?.[k]?.style}`)
    if (sides.length > 0) parts.push(`border(${sides.join(",")})`)
  }
  if (s.alignment) {
    const bits: string[] = []
    if (s.alignment.horizontal) bits.push(`h=${s.alignment.horizontal}`)
    if (s.alignment.vertical) bits.push(`v=${s.alignment.vertical}`)
    if (s.alignment.wrapText) bits.push("wrap")
    if (bits.length > 0) parts.push(`align(${bits.join(",")})`)
  }
  if (s.numFmt) parts.push(`numFmt(${s.numFmt})`)
  return parts.join(" ")
}

const projectSheet = (sheet: Sheet): SheetModel => {
  const styles: Record<string, string> = {}
  const formulas: Record<string, string> = {}
  for (const [key, cell] of sheet.cells ?? new Map<string, Cell>()) {
    const [r, c] = key.split(",").map(Number)
    const a1 = `${colName(c as number)}${(r as number) + 1}`
    const line = styleLine(cell.style)
    if (line !== "") styles[a1] = line
    if (cell.formula !== undefined) {
      formulas[a1] = `${cell.formula} => ${JSON.stringify(flat(cell.formulaResult))}`
    }
  }

  const columnStyles: Record<string, string> = {}
  const hiddenColumns: string[] = []
  sheet.columns?.forEach((col, i) => {
    if (!col) return
    if (col.hidden) hiddenColumns.push(colName(i))
    const line = styleLine(col.style)
    if (line !== "") columnStyles[colName(i)] = line
  })

  const hiddenRows: number[] = []
  for (const [row, def] of sheet.rowDefs ?? new Map()) {
    if (def.hidden) hiddenRows.push(row + 1)
  }

  const pageSetup = { ...sheet.pageSetup } as Record<string, unknown>
  // Every sheet Excel writes carries the same Normal margins; only the
  // one fixture that changes them says anything by having them.
  delete pageSetup.margins
  const left = sheet.pageSetup?.margins?.left
  if (left !== undefined && left !== 0.7) pageSetup.leftMargin = left

  return {
    name: sheet.name,
    rows: (sheet.rows ?? []).map((row) => row.map(flat)),
    merges: (sheet.merges ?? [])
      .map((m) => `${colName(m.startCol)}${m.startRow + 1}:${colName(m.endCol)}${m.endRow + 1}`)
      .sort(),
    freeze: sheet.freezePane
      ? `${sheet.freezePane.rows ?? 0}x${sheet.freezePane.columns ?? 0}`
      : null,
    hiddenRows: hiddenRows.sort((a, b) => a - b),
    hiddenColumns: hiddenColumns.sort(),
    styles,
    columnStyles,
    autoFilter: sheet.autoFilter?.range ?? null,
    validations: (sheet.dataValidations ?? [])
      .map(
        (v) =>
          `${v.range} ${v.type}` +
          `${v.operator ? ` ${v.operator}` : ""}` +
          `${v.values ? ` [${v.values.join("|")}]` : ""}` +
          `${v.formula1 ? ` f1=${v.formula1}` : ""}`,
      )
      .sort(),
    protection: Object.entries(sheet.protection ?? {})
      .filter(([k, v]) => v === true && k !== "password")
      .map(([k]) => k)
      .sort(),
    comments: [...(sheet.cells ?? new Map<string, Cell>())]
      .filter(([, c]) => c.comment)
      .map(([key, c]) => {
        const [r, cc] = key.split(",").map(Number)
        return `${colName(cc as number)}${(r as number) + 1}: ${c.comment?.text ?? ""}`
      })
      .sort(),
    conditional: (sheet.conditionalRules ?? [])
      .map(
        (r) =>
          `${r.range} ${r.type} ${r.operator ?? "-"} ${r.formula ?? "-"}` +
          `${r.stopIfTrue ? " stopIfTrue" : ""}` +
          ` fill=${colorOf(r.style?.fill?.type === "pattern" ? r.style.fill.bgColor : undefined)}`,
      )
      .sort(),
    formulas,
    pageSetup: Object.keys(pageSetup).length > 0 ? pageSetup : null,
  }
}

const project = (wb: Workbook, warnings: string[]): WorkbookModel => ({
  sheets: wb.sheets.map(projectSheet),
  dateSystem: wb.dateSystem ?? null,
  namedRanges: (wb.namedRanges ?? [])
    .filter((n) => !n.name.startsWith("_xlnm."))
    .map((n) => `${n.name}=${n.range}${n.scope ? ` @${n.scope}` : ""}`)
    .sort(),
  hasAuthor: Boolean(wb.properties?.creator || wb.properties?.lastModifiedBy),
  warnings,
})

const at = (model: unknown, path: string): unknown =>
  path.split(".").reduce<unknown>((acc, k) => (acc as Record<string, unknown>)?.[k], model)

const setAt = (model: unknown, path: string, value: unknown): void => {
  const keys = path.split(".")
  const last = keys.pop() as string
  const parent = keys.reduce<unknown>(
    (acc, k) => (acc as Record<string, unknown>)?.[k],
    model,
  ) as Record<string, unknown>
  parent[last] = value
}

// ── loading ─────────────────────────────────────────────────────────

const bytes = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)))

const golden = (name: string): WorkbookModel =>
  JSON.parse(
    readFileSync(new URL(`./fixtures/${name}.golden.json`, import.meta.url), "utf-8"),
  ) as WorkbookModel

const modelOf = async (
  name: string,
  reader: (b: Uint8Array, o?: ReadOptions) => Promise<Workbook>,
): Promise<WorkbookModel> => {
  const warnings: string[] = []
  const wb = await reader(bytes(name), {
    readStyles: true,
    onWarning: (w) => warnings.push(w.message),
  })
  return project(wb, warnings)
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

/**
 * The golden with its known-defect fields replaced by what hucre returns
 * today, so the rest of the fixture still asserts. Each swap is named by
 * its issue number in the golden JSON and re-asserted, unswapped, by the
 * `it.fails` below.
 */
const expectedNow = (file: string): WorkbookModel => {
  const g = golden(file)
  for (const d of g.knownDefects ?? []) setAt(g, d.path, d.actual)
  delete g.knownDefects
  return g
}

describe("workbooks written by Excel, not by this test suite", () => {
  for (const { file, reader } of FIXTURES) {
    describe(file, () => {
      it("matches its hand-written model through the explicit reader", async () => {
        expect(await modelOf(file, reader)).toEqual(expectedNow(file))
      })

      it("reads the same way through read()'s format detection", async () => {
        expect(await modelOf(file, read)).toEqual(expectedNow(file))
      })

      for (const d of golden(file).knownDefects ?? []) {
        it.fails(`${d.issue} — ${d.why}`, async () => {
          expect(at(await modelOf(file, reader), d.path)).toEqual(d.expected)
        })
      }
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
  // values placed out to column 15,312 describe 30.6 million slots and
  // the workbook is refused. The real file behind this had 76,277 values
  // over 507 columns, a 305,612,208-slot box at 0.03% fill, and Excel
  // opens it without complaint.
  describe("excel-sparse.xlsx", () => {
    it.fails("#501 — a sparse sheet can be read without raising any bound", async () => {
      const wb = await readXlsx(bytes("excel-sparse.xlsx"))
      expect(wb.sheets[0]?.name).toBe("Sparse")
    })

    it("fails the way #501 describes, and not some other way", async () => {
      await expect(readXlsx(bytes("excel-sparse.xlsx"))).rejects.toThrow(
        /spans 2000 rows x 15312 columns .* over the \d+ limit/,
      )
    })

    // The escape hatch the error message offers should at least work.
    // It does — which is what makes the message misleading rather than
    // wrong: `range` is usable only if you already know where the data
    // is, and on the real file the used columns were scattered across
    // 507 of 15,312.
    it("can be read when the caller already knows where the data is", async () => {
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

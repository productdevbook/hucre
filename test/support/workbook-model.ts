import { isCellError } from "../../src/cell-error"
import type { Cell, CellStyle, Sheet, Workbook } from "../../src/_types"

/** A cell value flattened to something a JSON golden can hold. */
export type Flat = string | number | boolean | null

interface SheetModel {
  name: string
  /** Dense rows. Dates become `D:<iso>` and errors `E:<token>` so neither can pass as a string. */
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

export interface WorkbookModel {
  sheets: SheetModel[]
  dateSystem: string | null
  /**
   * Whether docProps carries an author. The fixtures are generated with
   * `RemovePersonalInformation = True`, so this must stay false — it is
   * the licence-and-privacy guarantee of the corpus, checked by the same
   * suite that reads it rather than only by a note in a markdown file.
   */
  hasAuthor: boolean
  /** `name=range` per defined name, sorted. Excel's built-in print
   * names are excluded — they are page setup, asserted separately. */
  namedRanges: string[]
  warnings: string[]
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

export const flat = (v: unknown): Flat => {
  if (v instanceof Date) return `D:${v.toISOString()}`
  if (isCellError(v)) return `E:${v.error}`
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

export const projectWorkbook = (wb: Workbook, warnings: string[]): WorkbookModel => ({
  sheets: wb.sheets.map(projectSheet),
  dateSystem: wb.dateSystem ?? null,
  namedRanges: (wb.namedRanges ?? [])
    .filter((n) => !n.name.startsWith("_xlnm."))
    .map((n) => `${n.name}=${n.range}${n.scope ? ` @${n.scope}` : ""}`)
    .sort(),
  hasAuthor: Boolean(wb.properties?.creator || wb.properties?.lastModifiedBy),
  warnings,
})

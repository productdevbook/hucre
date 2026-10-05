// Value edits share this boundary: rows are the dense cache, CellStore
// carries sparse values and metadata, and a writer prefers an override.
import type { Cell, CellType, CellValue, Sheet } from "./_types"
import { cellEntries, getCell } from "./cell-store"
import { isCellError } from "./cell-error"

export type ValuePredicate =
  | CellValue
  | RegExp
  | ((value: CellValue, row: number, col: number) => boolean)

export function valueType(value: CellValue): CellType {
  if (value === null) return "empty"
  if (value instanceof Date) return "date"
  if (typeof value === "string") return "string"
  if (typeof value === "number") return "number"
  if (typeof value === "boolean") return "boolean"
  return "error"
}

function hasFormula(cell: Cell): boolean {
  return cell.formula !== undefined || cell.type === "formula" || cell.formulaType !== undefined
}

function effectiveValue(cell: Cell | undefined, fallback: CellValue): CellValue {
  if (!cell) return fallback
  if (hasFormula(cell) && cell.formulaResult !== undefined) return cell.formulaResult
  return cell.value !== undefined ? cell.value : fallback
}

/** Visit each materialized coordinate once; sparse cells never grow a grid. */
export function visitValues(
  sheet: Sheet,
  visit: (value: CellValue, row: number, col: number, cell: Cell | undefined) => void,
): void {
  for (let row = 0; row < sheet.rows.length; row++) {
    const values = sheet.rows[row]!
    for (let col = 0; col < values.length; col++) {
      const cell = getCell(sheet.cells, row, col)
      visit(effectiveValue(cell, values[col] ?? null), row, col, cell)
    }
  }
  for (const [row, col, cell] of cellEntries(sheet.cells)) {
    if (col < (sheet.rows[row]?.length ?? 0)) continue
    visit(effectiveValue(cell, null), row, col, cell)
  }
}

/** Update value, type and formula cache together, retaining unrelated metadata. */
export function setValue(
  sheet: Sheet,
  row: number,
  col: number,
  value: CellValue,
  cell?: Cell,
): void {
  const values = sheet.rows[row]
  if (values && col < values.length) values[col] = value
  if (!cell) return
  const formula = hasFormula(cell)
  // Runs encode the previous text and outrank `value` in XLSX serialization.
  // Retaining them after an edit silently writes the old text back out.
  if (value !== effectiveValue(cell, null)) delete cell.richText
  cell.value = value
  cell.type = formula ? "formula" : cell.richText?.length ? "richText" : valueType(value)
  if (formula) cell.formulaResult = value
}

export function matchesValue(
  value: CellValue,
  predicate: ValuePredicate,
  row: number,
  col: number,
): boolean {
  if (typeof predicate === "function") return predicate(value, row, col)
  if (predicate instanceof RegExp) {
    // A /g or /y expression must not carry state between cells, including
    // state left by the caller before this operation starts.
    predicate.lastIndex = 0
    const matched = typeof value === "string" && predicate.test(value)
    predicate.lastIndex = 0
    return matched
  }
  return isCellError(value) && isCellError(predicate)
    ? value.error === predicate.error
    : value === predicate
}

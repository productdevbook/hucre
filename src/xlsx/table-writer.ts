// ── Table Writer ──────────────────────────────────────────────────────
// Generates xl/tables/tableN.xml for Excel Table (ListObject) support.

import type { TableDefinition } from "../_types"
import { parseRange } from "../cell-utils"
import { xmlDocument, xmlElement, xmlSelfClose, xmlEscape } from "../xml/writer"

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

// ── Types ────────────────────────────────────────────────────────────

export interface TableResult {
  /** The table XML content (xl/tables/tableN.xml) */
  tableXml: string
  /** The table id attribute (unique across workbook) */
  tableId: number
}

// ── Writer ───────────────────────────────────────────────────────────

/**
 * Generate a table XML file for an Excel Table (ListObject).
 *
 * @param table - The table definition
 * @param tableId - Unique table ID (1-based, unique across the workbook)
 * @param globalTableIndex - Global table index for file naming (1-based)
 * @returns TableResult with XML content and table ID
 */
export function writeTable(
  table: TableDefinition,
  tableId: number,
  _globalTableIndex: number,
): TableResult {
  const displayName = table.displayName ?? table.name
  const ref = table.range ?? ""
  const showAutoFilter = table.showAutoFilter !== false
  const showTotalRow = table.showTotalRow === true

  // Root <table> attributes
  const tableAttrs: Record<string, string | number> = {
    xmlns: NS_SPREADSHEET,
    id: tableId,
    name: table.name,
    displayName,
    ref,
  }

  if (!showTotalRow) {
    tableAttrs["totalsRowShown"] = 0
  } else {
    tableAttrs["totalsRowCount"] = 1
  }

  const children: string[] = []

  // <autoFilter> — covers the data range (excluding total row)
  if (showAutoFilter) {
    // For autoFilter ref, we need to exclude the totals row if present
    const autoFilterRef = showTotalRow ? removeLastRow(ref) : ref
    children.push(xmlSelfClose("autoFilter", { ref: autoFilterRef }))
  }

  // <tableColumns>
  const colElements: string[] = []
  for (let i = 0; i < table.columns.length; i++) {
    const col = table.columns[i]
    const colAttrs: Record<string, string | number> = {
      id: i + 1,
      name: col.name,
    }

    if (showTotalRow && col.totalFunction) {
      colAttrs["totalsRowFunction"] = col.totalFunction
    }

    if (showTotalRow && col.totalLabel) {
      colAttrs["totalsRowLabel"] = col.totalLabel
    }

    // Custom formula in total row
    if (showTotalRow && col.totalFunction === "custom" && col.totalFormula) {
      const formulaChild = xmlElement("totalsRowFormula", undefined, xmlEscape(col.totalFormula))
      colElements.push(xmlElement("tableColumn", colAttrs, [formulaChild]))
    } else {
      colElements.push(xmlSelfClose("tableColumn", colAttrs))
    }
  }

  children.push(xmlElement("tableColumns", { count: table.columns.length }, colElements))

  // <tableStyleInfo>
  const styleName = table.style ?? "TableStyleMedium2"
  const showRowStripes = table.showRowStripes !== false
  const showColumnStripes = table.showColumnStripes === true

  children.push(
    xmlSelfClose("tableStyleInfo", {
      name: styleName,
      showFirstColumn: 0,
      showLastColumn: 0,
      showRowStripes: showRowStripes ? 1 : 0,
      showColumnStripes: showColumnStripes ? 1 : 0,
    }),
  )

  const tableXml = xmlDocument("table", tableAttrs, children)

  return { tableXml, tableId }
}

// ── Total Row Cells ──────────────────────────────────────────────────

/**
 * `SUBTOTAL` function numbers for each `totalsRowFunction`, as Excel
 * writes them. The 10x forms skip rows hidden by the table's filter.
 */
const SUBTOTAL_FUNCTION_NUM: Record<string, number> = {
  average: 101,
  countNums: 102,
  count: 103,
  max: 104,
  min: 105,
  stdDev: 107,
  sum: 109,
  var: 110,
}

/** One cell of a table's total row, 0-based. */
export interface TotalRowCell {
  row: number
  col: number
  value?: string
  formula?: string
}

/**
 * The cells a table's total row needs on the sheet.
 *
 * `totalsRowFunction`, `totalsRowFormula` and `totalsRowLabel` in the
 * table part do not put anything in the grid. Excel writes the total row
 * twice, once in the table part and once as ordinary cells, and reads the
 * cells. Without them a built-in function shows an empty total row, and
 * a custom formula with no matching cell makes Excel repair the file.
 *
 * The formulas copy what Excel writes. They use the table's
 * `displayName`, the name formulas see, and escape `[ ] # '` in the
 * column name with a `'`, so `a[b]` is written `Table1[a'[b']]`.
 * See #604.
 */
export function totalRowCells(table: TableDefinition, range: string): TotalRowCell[] {
  if (table.showTotalRow !== true) return []
  const { startCol, endRow } = parseRange(range)
  const displayName = table.displayName ?? table.name
  const cells: TotalRowCell[] = []

  for (let i = 0; i < table.columns.length; i++) {
    const col = table.columns[i]!
    const fn = col.totalFunction
    const at = { row: endRow, col: startCol + i }

    if (fn === "custom") {
      if (col.totalFormula) cells.push({ ...at, formula: col.totalFormula })
    } else if (fn !== undefined && Object.hasOwn(SUBTOTAL_FUNCTION_NUM, fn)) {
      const ref = `${displayName}[${escapeStructuredRefColumn(col.name)}]`
      cells.push({ ...at, formula: `SUBTOTAL(${SUBTOTAL_FUNCTION_NUM[fn]},${ref})` })
    } else if (col.totalLabel) {
      cells.push({ ...at, value: col.totalLabel })
    }
  }

  return cells
}

/** Escape a column name for use inside a structured reference's brackets. */
function escapeStructuredRefColumn(name: string): string {
  return name.replace(/[[\]#']/g, "'$&")
}

// ── Helpers ──────────────────────────────────────────────────────────

/**
 * Remove the last row from a range reference.
 * E.g. "A1:D10" → "A1:D9"
 */
function removeLastRow(ref: string): string {
  const colonIdx = ref.indexOf(":")
  if (colonIdx === -1) return ref

  const endPart = ref.slice(colonIdx + 1)
  // Parse trailing number
  let numStart = endPart.length
  while (
    numStart > 0 &&
    endPart.charCodeAt(numStart - 1) >= 48 &&
    endPart.charCodeAt(numStart - 1) <= 57
  ) {
    numStart--
  }

  const colLetters = endPart.slice(0, numStart)
  const rowNum = parseInt(endPart.slice(numStart), 10)
  if (isNaN(rowNum) || rowNum <= 1) return ref

  return `${ref.slice(0, colonIdx + 1)}${colLetters}${rowNum - 1}`
}

import type { Cell, WorkbookWriteOptions } from "../_types"
import { cellRef } from "../cell-utils"

// A new Cell field needs a capability decision instead of disappearing
// silently in an adapter. Values/types are inferred; these fields are not
// represented by the shared ODS cell serializer.
const DROPS: Record<
  Exclude<
    keyof Cell,
    "value" | "type" | "style" | "formula" | "formulaResult" | "richText" | "hyperlink"
  >,
  string
> = {
  checkbox: "ODS writes the boolean value without the Excel checkbox control.",
  formulaType: "ODS writes ordinary formulas without XLSX shared/array metadata.",
  formulaSharedIndex: "ODS does not encode XLSX shared-formula indexes.",
  formulaRef: "ODS does not rebuild XLSX formula ranges.",
  formulaDynamic: "ODS does not encode the Excel dynamic-array flag.",
  comment: "The ODS writer does not emit cell annotations.",
}

export function reportOdsCellDrops(
  cell: Partial<Cell>,
  onDrop: WorkbookWriteOptions["onDrop"],
  sheet: string,
  row: number,
  col: number,
  unstyled = false,
): void {
  if (!onDrop) return
  const report = (field: string, reason: string) =>
    onDrop({ field: `cells.${field}`, sheet, cell: cellRef(row, col), reason })
  for (const [field, reason] of Object.entries(DROPS)) {
    const value = cell[field as keyof typeof DROPS]
    if (value !== undefined && value !== false) report(field, reason)
  }
  if (unstyled) {
    const reason =
      "True ODS streaming emits styles before rows; use writeOds or OdsStreamWriter to retain formatting."
    if (cell.style) report("style", reason)
    if (cell.richText?.some((run) => run.font)) report("richText.font", reason)
  }
}

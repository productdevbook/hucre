import type { Sheet, CellValue, MergeRange } from "../_types"
import { parseSax } from "../xml/parser"
import { MAX_COL_INDEX, MAX_ROW_INDEX, MAX_SPAN_CELLS } from "../limits"

/**
 * Bound a `colspan` / `rowspan` attribute. Non-numeric and non-positive
 * values collapse to 1, which is the HTML default; anything past the
 * sheet's own limit is clamped, since a span wider than the grid cannot
 * mean anything useful.
 */
function clampSpan(raw: string | undefined, max: number): number {
  const value = Number(raw ?? "1")
  if (!Number.isFinite(value) || value < 1) return 1
  return Math.min(Math.trunc(value), max)
}

/**
 * Parse an HTML table string into a Sheet.
 *
 * This is a best-effort parser that handles well-formed table HTML.
 * It uses the SAX XML parser internally, which handles basic HTML structure.
 *
 * Supports: `<table>`, `<thead>`, `<tbody>`, `<tfoot>`, `<tr>`, `<td>`, `<th>`,
 * `colspan`, `rowspan` attributes.
 */
export function fromHtml(html: string, options?: { sheetName?: string }): Sheet {
  const rows: CellValue[][] = []
  const merges: MergeRange[] = []

  // Track which cells are occupied by rowspan from previous rows.
  // Key: "row,col" → true
  const occupied = new Set<string>()

  let inTable = false
  let inRow = false
  let inCell = false
  let currentRowCells: CellValue[] = []
  let currentCellText = ""
  let currentCellColspan = 1
  let currentCellRowspan = 1

  // We need to track the actual grid column for each cell due to rowspan reservations
  let currentRow = -1

  parseSax(html, {
    onOpenTag(tag, attrs) {
      const local = tagLocal(tag)

      if (local === "table") {
        inTable = true
        return
      }

      if (!inTable) return

      if (local === "tr") {
        inRow = true
        currentRow++
        currentRowCells = []
        return
      }

      if ((local === "td" || local === "th") && inRow) {
        inCell = true
        currentCellText = ""
        // fromHtml exists to consume markup the caller did not write, so
        // spans are bounded: rowspan x colspan drives a nested loop of Set
        // insertions, and 100000 x 100000 is 1e10 of them. See #363.
        currentCellColspan = clampSpan(attrs.colspan, MAX_COL_INDEX + 1)
        currentCellRowspan = clampSpan(attrs.rowspan, MAX_ROW_INDEX + 1)
        // Clamping each span alone is not enough: the reservation below
        // is a nested loop, so the cost is their product. Shrink the
        // rowspan until the rectangle fits.
        if (currentCellColspan * currentCellRowspan > MAX_SPAN_CELLS) {
          currentCellRowspan = Math.max(1, Math.floor(MAX_SPAN_CELLS / currentCellColspan))
        }
      }
    },

    onText(text) {
      if (inCell) {
        currentCellText += text
      }
    },

    onCloseTag(tag) {
      const local = tagLocal(tag)

      if ((local === "td" || local === "th") && inCell) {
        inCell = false

        // Find the next available column in this row
        let col = currentRowCells.length
        while (occupied.has(`${currentRow},${col}`)) {
          // Push null for occupied cells in our row array
          currentRowCells.push(null)
          col = currentRowCells.length
        }

        const value = parseValue(currentCellText.trim())

        // Place the value
        currentRowCells.push(value)

        // Fill the remaining colspan slots with null. Track the actual grid
        // column via currentRowCells.length so any cells reserved by an
        // earlier row's rowspan (occupied) are skipped correctly — the old
        // arithmetic drifted as nulls were pushed.
        for (let c = 1; c < currentCellColspan; c++) {
          while (occupied.has(`${currentRow},${currentRowCells.length}`)) {
            currentRowCells.push(null)
          }
          currentRowCells.push(null)
        }

        // Record merge if colspan > 1 or rowspan > 1
        if (currentCellColspan > 1 || currentCellRowspan > 1) {
          merges.push({
            startRow: currentRow,
            startCol: col,
            endRow: currentRow + currentCellRowspan - 1,
            endCol: col + currentCellColspan - 1,
          })
        }

        // Reserve cells for rowspan in subsequent rows
        if (currentCellRowspan > 1) {
          for (let r = 1; r < currentCellRowspan; r++) {
            for (let c = 0; c < currentCellColspan; c++) {
              occupied.add(`${currentRow + r},${col + c}`)
            }
          }
        }

        return
      }

      if (local === "tr" && inRow) {
        inRow = false
        rows.push(currentRowCells)
        return
      }

      if (local === "table") {
        inTable = false
      }
    },
  })

  return {
    name: options?.sheetName ?? "Sheet1",
    rows,
    merges: merges.length > 0 ? merges : undefined,
  }
}

/** Extract the local tag name (strip namespace prefix) */
function tagLocal(tag: string): string {
  const colon = tag.indexOf(":")
  return (colon === -1 ? tag : tag.slice(colon + 1)).toLowerCase()
}

/** Try to parse a string value as a number or return it as-is */
function parseValue(text: string): CellValue {
  if (text === "") return null

  // Try number
  const num = Number(text)
  if (!Number.isNaN(num) && text !== "") {
    return num
  }

  return text
}

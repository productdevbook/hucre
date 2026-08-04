import type { Sheet, CellValue, MergeRange, SheetA11y } from "../_types"
import { parseSax } from "../xml/parser"
import { XmlError, ParseError } from "../errors"
import { inferType } from "../_infer"
import { MAX_COL_INDEX, MAX_ROW_INDEX, MAX_SPAN_CELLS, MAX_TOTAL_CELLS } from "../limits"

/** Type marker `toHtml` writes as a CSS class on a cell. */
type DeclaredType = "num" | "bool" | "date" | "null"

export interface HtmlImportOptions {
  /** Name for the returned sheet. Default: "Sheet1" */
  sheetName?: string
  /**
   * Infer numbers, booleans and ISO dates from cell text. Default: true.
   *
   * The name and the behaviour are `parseCsv`'s; the **default** is not.
   * `parseCsv` defaults to `false` because a CSV field is already a
   * string and quoting can say so. An HTML table has no such convention —
   * every cell is text, `fromHtml` has always coerced, and returning
   * `"42"` for `<td>42</td>` would silently restring every existing
   * caller's data. Pass `false` for cell text exactly as written.
   */
  typeInference?: boolean
  /**
   * Keep strings with leading zeros ("0123", "007") as strings instead of
   * converting them to numbers. Default: true, as in `parseCsv`.
   */
  preserveLeadingZeros?: boolean
  /**
   * Honour the type classes {@link toHtml} writes (`hucre-num`,
   * `hucre-bool`, `hucre-date`, `hucre-null`). Default: true.
   *
   * These are declarations, not guesses, so they apply even with
   * `typeInference: false`. Set this to `false` for markup that happens to
   * use the same class names for something else.
   */
  classes?: boolean
  /** Class prefix those type classes were written with. Default: "hucre" */
  classPrefix?: string
}

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
 * Best-effort, and specific about what that means: markup the XML scanner
 * cannot finish (a `<!--` with no `-->`, a truncated tag) ends the parse
 * where it broke and returns the rows collected so far rather than
 * throwing — `fromHtml` consumes third-party HTML, where a stray
 * construct at the bottom of a page should not cost you the table above
 * it. Input that would exhaust memory still throws a {@link ParseError}:
 * that is a resource bound, not a syntax problem, and answering it with
 * half a sheet would be worse than saying so.
 *
 * Supports: `<table>`, `<thead>`, `<tbody>`, `<tfoot>`, `<tr>`, `<td>`,
 * `<th>`, `<caption>`, `colspan`, `rowspan`.
 *
 * What comes back beyond `rows` and `merges`:
 * - a `<thead>` row, or a row made entirely of `<th>`, sets
 *   `sheet.a11y.headerRow` to that row's index
 * - `<caption>` text becomes `sheet.a11y.summary`
 *
 * This is **not** the inverse of {@link toHtml}. Inline styles, the
 * `<style>` block, `role` and `aria-label` are presentation and are not
 * read back; see the round-trip note in the README.
 *
 * Cell text is trimmed, because whitespace around a `<td>` is how markup
 * is indented rather than something the author typed. `toHtml` does not
 * trim on the way out, so a value that really is padded survives a write
 * but not the read back.
 */
export function fromHtml(html: string, options?: HtmlImportOptions): Sheet {
  const typeInference = options?.typeInference !== false
  const preserveLeadingZeros = options?.preserveLeadingZeros !== false
  const useClasses = options?.classes !== false
  const classPrefix = options?.classPrefix ?? "hucre"

  const rows: CellValue[][] = []
  const merges: MergeRange[] = []

  // Track which cells are occupied by rowspan from previous rows.
  // Key: "row,col" → true
  const occupied = new Set<string>()

  // Nesting depth, not a flag: a table inside a `<td>` is ordinary HTML,
  // and treating its `</table>` as the end of the outer one silently threw
  // away every remaining row of the table the caller asked for. Only depth
  // 1 is the table being read; deeper markup contributes its text to the
  // cell that contains it, which is what a browser shows there.
  let tableDepth = 0
  let inRow = false
  let inCell = false
  let inCaption = false
  let currentRowCells: CellValue[] = []
  let currentCellText = ""
  let currentCellColspan = 1
  let currentCellRowspan = 1
  let currentCellType: DeclaredType | null = null
  let caption = ""

  // Header detection: a row inside <thead>, or one whose cells are all
  // <th>. Only the first such row is reported — a11y.headerRow is a single
  // index, and the first is the one screen readers announce.
  let inThead = false
  let headerRow: number | undefined
  let rowCellCount = 0
  let rowHeaderCellCount = 0

  // We need to track the actual grid column for each cell due to rowspan reservations
  let currentRow = -1

  // Slots this table has materialized: placed values plus colspan padding.
  // A single `<td colspan="16384">` costs 30 bytes of markup and 16,384
  // array entries, so 175 KB of them reached 82 million entries and five
  // seconds before this counter existed.
  let gridCells = 0

  function spend(): void {
    if (++gridCells > MAX_TOTAL_CELLS) {
      throw new ParseError(`HTML table spans over ${MAX_TOTAL_CELLS} cells`)
    }
  }

  /** Append one slot to the current row, dropping anything past the last column. */
  function pushSlot(value: CellValue): void {
    if (currentRowCells.length > MAX_COL_INDEX) return
    spend()
    currentRowCells.push(value)
  }

  function startRow(): void {
    inRow = true
    currentRow++
    currentRowCells = []
    rowCellCount = 0
    rowHeaderCellCount = 0
  }

  function endRow(): void {
    inRow = false
    if (rows.length <= MAX_ROW_INDEX) rows.push(currentRowCells)
    if (
      headerRow === undefined &&
      rowCellCount > 0 &&
      (inThead || rowHeaderCellCount === rowCellCount)
    ) {
      headerRow = currentRow
    }
  }

  /**
   * Finish the cell being read. Called from `</td>` and also from every
   * place that proves the cell ended without one — the next `<td>`, the
   * end of the row, the end of the table. An unclosed cell used to be
   * dropped outright, taking its row with it when nothing closed the row
   * either.
   */
  function closeCell(): void {
    inCell = false
    rowCellCount++

    // Find the next available column in this row
    let col = currentRowCells.length
    while (col <= MAX_COL_INDEX && occupied.has(`${currentRow},${col}`)) {
      // Push null for occupied cells in our row array
      pushSlot(null)
      col = currentRowCells.length
    }

    if (col > MAX_COL_INDEX) return

    const value = parseValue(currentCellText.trim(), currentCellType)

    // Place the value
    pushSlot(value)

    // Fill the remaining colspan slots with null. Track the actual grid
    // column via currentRowCells.length so any cells reserved by an
    // earlier row's rowspan (occupied) are skipped correctly — the old
    // arithmetic drifted as nulls were pushed.
    for (let c = 1; c < currentCellColspan; c++) {
      while (
        currentRowCells.length <= MAX_COL_INDEX &&
        occupied.has(`${currentRow},${currentRowCells.length}`)
      ) {
        pushSlot(null)
      }
      pushSlot(null)
    }

    // Record merge if colspan > 1 or rowspan > 1
    const endCol = Math.min(col + currentCellColspan - 1, MAX_COL_INDEX)
    const endRowIndex = Math.min(currentRow + currentCellRowspan - 1, MAX_ROW_INDEX)
    if (endCol > col || endRowIndex > currentRow) {
      merges.push({ startRow: currentRow, startCol: col, endRow: endRowIndex, endCol })
    }

    // Reserve cells for rowspan in subsequent rows
    if (currentCellRowspan > 1) {
      // MAX_SPAN_CELLS bounds one cell's rectangle; the set it feeds is
      // bounded document-wide by the same number, or a page of 3,000
      // `rowspan="1000000"` cells would each pay the per-cell price.
      const reserved = (currentCellRowspan - 1) * currentCellColspan
      if (occupied.size + reserved > MAX_SPAN_CELLS) {
        throw new ParseError(`HTML table reserves over ${MAX_SPAN_CELLS} cells through rowspan`)
      }
      for (let r = 1; r < currentCellRowspan; r++) {
        for (let c = 0; c < currentCellColspan; c++) {
          occupied.add(`${currentRow + r},${col + c}`)
        }
      }
    }
  }

  function parseValue(text: string, declared: DeclaredType | null): CellValue {
    if (declared === "null") return null
    if (text === "") return null

    if (declared !== null) {
      // A type class is the writer stating the type, not the reader
      // guessing at it, so it counts even with typeInference off. Leading
      // zeros need no protection here: a cell the writer called a number
      // never had any to lose.
      const typed = inferType(text, false)
      if (declared === "num" && typeof typed === "number") return typed
      if (declared === "bool" && typeof typed === "boolean") return typed
      if (declared === "date" && typed instanceof Date) return typed
      // Class and text disagree — the text is what is actually there.
    }

    if (!typeInference) return text
    return inferType(text, preserveLeadingZeros)
  }

  try {
    parseSax(html, {
      onOpenTag(tag, attrs) {
        const local = tagLocal(tag)

        if (local === "table") {
          tableDepth++
          return
        }

        // Depth 0 is markup around the table; depth 2+ is a nested table,
        // whose text keeps flowing into the enclosing cell through onText.
        if (tableDepth !== 1) return

        if (local === "caption" && !inCell) {
          inCaption = true
          return
        }

        if (local === "thead") {
          inThead = true
          return
        }

        if (local === "tr") {
          if (inCell) closeCell()
          if (inRow) endRow()
          startRow()
          return
        }

        if ((local === "td" || local === "th") && inRow) {
          // A cell already open means the previous one was never closed.
          if (inCell) closeCell()
          inCell = true
          if (local === "th") rowHeaderCellCount++
          currentCellText = ""
          currentCellType = useClasses ? declaredType(attrs.class, classPrefix) : null
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
          return
        }
        if (inCaption) caption += text
      },

      onCloseTag(tag) {
        const local = tagLocal(tag)

        if (local === "table") {
          if (tableDepth === 1) {
            // Nothing after this can close them, so flush what is open.
            if (inCell) closeCell()
            if (inRow) endRow()
            inThead = false
          }
          if (tableDepth > 0) tableDepth--
          return
        }

        if (tableDepth !== 1) return

        if ((local === "td" || local === "th") && inCell) {
          closeCell()
          return
        }

        if (local === "caption") {
          inCaption = false
          return
        }

        if (local === "tr" && inRow) {
          if (inCell) closeCell()
          endRow()
          return
        }

        if (local === "thead") {
          inThead = false
        }
      },
    })
  } catch (error) {
    // Only the scanner giving up on malformed markup is survivable. A
    // resource bound (ParseError, thrown above) means the answer would be
    // wrong, not merely short, so it propagates.
    if (!(error instanceof XmlError)) throw error
  }

  // Truncated input, or markup that broke mid-table: emit what was read.
  if (inCell) closeCell()
  if (inRow) endRow()

  // The two things a table says about itself that a sheet has somewhere to
  // put: which row is the header, and what the table is called. Both land
  // on a11y because that is where hucre already keeps them, and both are
  // what `audit` looks for.
  const summary = caption.trim()
  const a11y: SheetA11y = {}
  if (summary !== "") a11y.summary = summary
  if (headerRow !== undefined) a11y.headerRow = headerRow
  const described = summary !== "" || headerRow !== undefined

  return {
    name: options?.sheetName ?? "Sheet1",
    rows,
    merges: merges.length > 0 ? merges : undefined,
    a11y: described ? a11y : undefined,
  }
}

/** Extract the local tag name (strip namespace prefix) */
function tagLocal(tag: string): string {
  const colon = tag.indexOf(":")
  return (colon === -1 ? tag : tag.slice(colon + 1)).toLowerCase()
}

/** Read the type class `toHtml` writes, if this class list carries one. */
function declaredType(classAttr: string | undefined, prefix: string): DeclaredType | null {
  if (!classAttr) return null
  for (const name of classAttr.split(/\s+/)) {
    if (!name.startsWith(`${prefix}-`)) continue
    const kind = name.slice(prefix.length + 1)
    if (kind === "num" || kind === "bool" || kind === "date" || kind === "null") return kind
  }
  return null
}

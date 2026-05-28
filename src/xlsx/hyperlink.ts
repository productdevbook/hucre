// ── Hyperlink helpers ────────────────────────────────────────────────

import type { CellValue, HyperlinkValue } from "../_types"

/**
 * Narrow a row-data value to a {@link HyperlinkValue}. Plain scalars are the
 * common case; a hyperlink is always an object carrying `text` and `hyperlink`
 * strings.
 */
export function isHyperlinkValue(v: unknown): v is HyperlinkValue {
  return (
    typeof v === "object" &&
    v !== null &&
    typeof (v as HyperlinkValue).hyperlink === "string" &&
    typeof (v as HyperlinkValue).text === "string"
  )
}

/** Reduce a rich data-row value to its scalar display value (for non-XLSX consumers). */
export function unwrapCellValue(v: CellValue | HyperlinkValue): CellValue {
  return isHyperlinkValue(v) ? v.text : v
}

/**
 * Build a rich {@link HyperlinkValue} for inline use in a {@link WriteSheet.data}
 * row object. The returned object can also be written by hand — this helper is
 * purely ergonomic sugar.
 *
 * @param text Display text shown in the cell.
 * @param hyperlink Link destination — an external URL, or an internal ref
 *   prefixed with `#` (e.g. `"#Sheet2!A1"`).
 * @param tooltip Optional hover tooltip.
 *
 * @example
 * ```ts
 * import { writeXlsx, link } from "hucre/xlsx"
 *
 * await writeXlsx({
 *   sheets: [{
 *     name: "Summary",
 *     columns: [{ header: "Link", key: "link" }, { header: "ID", key: "id" }],
 *     data: [{ link: link("Open", "https://example.com/items/abc"), id: "abc" }],
 *   }],
 * })
 * ```
 */
export function link(text: string, hyperlink: string, tooltip?: string): HyperlinkValue {
  return tooltip === undefined ? { text, hyperlink } : { text, hyperlink, tooltip }
}

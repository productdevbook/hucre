import { visitValues, setValue } from "./_sheet-values"
// ── Template Engine ──────────────────────────────────────────────────
// Fill {{placeholder}} patterns in workbook cells with data values.
// Works with round-trip: openXlsx -> fillTemplate -> saveXlsx.

import { isCellError } from "./cell-error"
import type { Workbook, CellValue } from "./_types"

/** Regex matching `{{key}}` placeholders (non-greedy, trims inner whitespace). */
const PLACEHOLDER_RE = /\{\{\s*([^}\s]+)\s*\}\}/g

/**
 * Fill template placeholders in a workbook with data values.
 *
 * Scans all cells for `{{key}}` patterns and replaces them with the
 * corresponding value from the `data` record. If a placeholder key
 * is not found in `data`, it is left as-is.
 *
 * When a cell contains only a single placeholder and the replacement
 * value is a non-string type (number, boolean, Date), the cell value
 * is set to that typed value directly. When a cell contains multiple
 * placeholders or mixed text, the result is always a string.
 *
 * @example
 * ```ts
 * const wb = await openXlsx(templateBytes);
 * const filled = fillTemplate(wb, {
 *   name: "Acme Corp",
 *   total: 12500,
 *   date: new Date("2025-01-15"),
 * });
 * const output = await saveXlsx(filled);
 * ```
 */
export function fillTemplate(workbook: Workbook, data: Record<string, CellValue>): Workbook {
  // Every lookup below uses Object.hasOwn rather than `key in data`.
  // `in` walks the prototype chain, so a template containing
  // `{{toString}}` used to resolve to Object.prototype.toString and put
  // a *function* into a cell — outside CellValue entirely. `constructor`,
  // `valueOf`, `hasOwnProperty` and `__proto__` behaved the same way.
  for (const sheet of workbook.sheets) {
    visitValues(sheet, (value, row, col, cell) => {
      if (typeof value !== "string" || !value.includes("{{")) return
      const single = /^\{\{\s*([^}\s]+)\s*\}\}$/.exec(value)
      if (single) {
        if (Object.hasOwn(data, single[1]!)) setValue(sheet, row, col, data[single[1]!]!, cell)
        return
      }
      const next = value.replace(PLACEHOLDER_RE, (match, key: string) => {
        if (!Object.hasOwn(data, key)) return match
        const replacement = data[key]
        if (replacement === null) return ""
        if (replacement instanceof Date) return replacement.toISOString()
        return isCellError(replacement) ? replacement.error : String(replacement)
      })
      if (next !== value) setValue(sheet, row, col, next, cell)
    })
  }
  return workbook
}

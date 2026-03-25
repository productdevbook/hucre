// ── Template Engine ──────────────────────────────────────────────────
// Fill {{placeholder}} patterns in workbook cells with data values.
// Works with round-trip: openXlsx -> fillTemplate -> saveXlsx.

import type { Workbook, CellValue } from "./_types";

/** Regex matching `{{key}}` placeholders (non-greedy, trims inner whitespace). */
const PLACEHOLDER_RE = /\{\{\s*([^}\s]+)\s*\}\}/g;

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
  for (const sheet of workbook.sheets) {
    for (let r = 0; r < sheet.rows.length; r++) {
      const row = sheet.rows[r]!;
      for (let c = 0; c < row.length; c++) {
        const val = row[c];
        if (typeof val !== "string") continue;

        // Check if this cell has any placeholders
        if (!val.includes("{{")) continue;

        // Check if the entire cell is a single placeholder
        const singleMatch = val.match(/^\{\{\s*([^}\s]+)\s*\}\}$/);
        if (singleMatch) {
          const key = singleMatch[1]!;
          if (key in data) {
            row[c] = data[key]!;
          }
          // If key not in data, leave as-is
          continue;
        }

        // Multiple placeholders or mixed text: string replacement
        const replaced = val.replace(PLACEHOLDER_RE, (match, key: string) => {
          if (key in data) {
            const replacement = data[key];
            if (replacement === null) return "";
            if (replacement instanceof Date) return replacement.toISOString();
            return String(replacement);
          }
          return match; // leave unmatched placeholders as-is
        });

        row[c] = replaced;
      }
    }

    // Also process the cells Map if present (for rich cell data)
    if (sheet.cells) {
      for (const [key, cell] of sheet.cells) {
        if (typeof cell.value !== "string") continue;
        if (!cell.value.includes("{{")) continue;

        const singleMatch = cell.value.match(/^\{\{\s*([^}\s]+)\s*\}\}$/);
        if (singleMatch) {
          const dataKey = singleMatch[1]!;
          if (dataKey in data) {
            cell.value = data[dataKey]!;
            // Update cell type based on value
            if (typeof cell.value === "number") cell.type = "number";
            else if (typeof cell.value === "boolean") cell.type = "boolean";
            else if (cell.value instanceof Date) cell.type = "date";
            else cell.type = "string";
          }
          continue;
        }

        cell.value = cell.value.replace(PLACEHOLDER_RE, (match, k: string) => {
          if (k in data) {
            const replacement = data[k];
            if (replacement === null) return "";
            if (replacement instanceof Date) return replacement.toISOString();
            return String(replacement);
          }
          return match;
        });
      }
    }
  }

  return workbook;
}

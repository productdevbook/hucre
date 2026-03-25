// ── TSV Export ───────────────────────────────────────────────────────
// Convenience wrappers around CSV writer with tab delimiter.

import type { CellValue, CsvWriteOptions } from "../_types";
import { writeCsv, writeCsvObjects } from "../csv/writer";

/**
 * Write a 2D array of cell values to a TSV string (tab-separated values).
 */
export function writeTsv(
  rows: CellValue[][],
  options?: Omit<CsvWriteOptions, "delimiter">,
): string {
  return writeCsv(rows, { ...options, delimiter: "\t" });
}

/**
 * Write an array of objects to a TSV string (tab-separated values).
 */
export function writeTsvObjects(
  data: Array<Record<string, CellValue>>,
  options?: Omit<CsvWriteOptions, "delimiter">,
): string {
  return writeCsvObjects(data, { ...options, delimiter: "\t" });
}

import type { CsvReadOptions, CellValue } from "../_types";
import { parseCsv } from "./reader";

/** Fetch a CSV from a URL and parse it (requires fetch API — Node.js/Deno) */
export async function fetchCsv(url: string, options?: CsvReadOptions): Promise<CellValue[][]> {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Failed to fetch: ${response.status}`);
  const text = await response.text();
  return parseCsv(text, options);
}

// ── Column Utilities ────────────────────────────────────────────────
// Filter column definition arrays by key.

import type { ColumnDef } from "./_types";

/** Pick specific columns by key (or header if no key). Preserves the order of `keys`. */
export function pickColumns<T>(columns: ColumnDef<T>[], keys: string[]): ColumnDef<T>[] {
  const map = new Map<string, ColumnDef<T>>();
  for (const col of columns) {
    const k = col.key ?? col.header;
    if (k) map.set(k, col);
  }
  return keys.map((k) => map.get(k)).filter((c): c is ColumnDef<T> => c !== undefined);
}

/** Omit specific columns by key (or header if no key). */
export function omitColumns<T>(columns: ColumnDef<T>[], keys: string[]): ColumnDef<T>[] {
  const keySet = new Set(keys);
  return columns.filter((col) => {
    const k = col.key ?? col.header;
    return !k || !keySet.has(k);
  });
}

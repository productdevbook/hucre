import type { StreamRow } from "../src/_types"

/** Drain a `stream*Rows` generator and keep only the row values. */
export async function valuesOf<T>(gen: AsyncIterable<StreamRow<T>>): Promise<T[]> {
  const out: T[] = []
  for await (const row of gen) out.push(row.values)
  return out
}

/** Drain a `stream*Rows` generator, rows and all. */
export async function rowsOf<T>(gen: AsyncIterable<StreamRow<T>>): Promise<Array<StreamRow<T>>> {
  const out: Array<StreamRow<T>> = []
  for await (const row of gen) out.push(row)
  return out
}

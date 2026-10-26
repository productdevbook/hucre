import { expect } from "vitest"
import type { CellValue, ReadOptions, StreamRow } from "../src/_types"
import { readXls } from "../src/xls/reader"
import { streamXlsRows } from "../src/xls/stream-reader"

/** Every row a stream yields, collected. */
export async function streamed(
  xls: Uint8Array,
  options?: ReadOptions & { sheet?: number | string },
): Promise<StreamRow[]> {
  const out: StreamRow[] = []
  for await (const row of streamXlsRows(xls, options)) out.push(row)
  return out
}

/**
 * A dense sheet as a stream yields it: the empty rows left out, each row
 * cut after its last value. A value is never null, so trailing nulls are
 * only the rectangle's padding.
 */
export function asStreamed(rows: CellValue[][]): StreamRow[] {
  const out: StreamRow[] = []
  rows.forEach((row, index) => {
    let end = row.length
    while (end > 0 && row[end - 1] === null) end--
    if (end > 0) out.push({ index, values: row.slice(0, end) })
  })
  return out
}

/** Each sheet, streamed, holds exactly what readXls reads for it. */
export async function expectStreamMatchesReadXls(
  xls: Uint8Array,
  options?: ReadOptions,
): Promise<void> {
  const workbook = await readXls(xls, options)
  for (const [sheet, { rows }] of workbook.sheets.entries()) {
    expect(await streamed(xls, { ...options, sheet })).toEqual(asStreamed(rows))
  }
}

// ── Streaming XLS Reader ─────────────────────────────────────────────
// Yields one sheet of a BIFF5/BIFF8 .xls a row at a time. The container
// is read whole — a CFB file's directory and allocation tables can sit
// anywhere in it — but the sheet is never materialized: values are
// decoded from the records as the cursor reaches them, by the same code
// `readXls` uses.

import type { CellValue, ReadOptions, StreamRow } from "../_types"
import { ParseError } from "../errors"
import { MAX_TOTAL_CELLS } from "../limits"
import { applyRangeFilter, parseRangeFilter, resolveTargetSheet } from "../_stream-rows"
import { RecordCursor, recordStarts, SID } from "./biff"
import {
  asParseError,
  assertInSheetBounds,
  decodeCells,
  openWorkbookStream,
  placeCell,
  readGlobals,
  type XlsGlobals,
} from "./reader"

/**
 * Stream one sheet of a BIFF5 or BIFF8 .xls workbook, a row at a time.
 *
 * The first sheet, or the one `sheet` names by 0-based index or by name;
 * a sheet that is not there yields nothing. `range` and `maxRows` mean
 * what they do for `streamXlsxRows`, and reading stops as soon as no
 * more rows can be emitted. As in the other stream readers, an empty row
 * is skipped and `index` is the row's own — a gap in the indexes is the
 * signal — and `values` run to the row's last cell, nulls in between.
 *
 * The input is read whole, because a CFB container needs random access;
 * the rows are not, so memory follows the file and its shared strings
 * rather than the sheet. Streaming needs the cells in row order — the
 * order BIFF lays its row blocks out in, and the one Excel writes. A sheet
 * whose records go back a row is buffered instead and yielded in order:
 * its cells are held the way `readXls` holds them, under the same
 * `maxTotalCells` bound.
 */
export async function* streamXlsRows(
  input: Uint8Array | ArrayBuffer | ReadableStream<Uint8Array>,
  options?: ReadOptions & { sheet?: number | string },
): AsyncGenerator<StreamRow, void, undefined> {
  const stream = await openWorkbookStream(input, options)
  let rows: Iterable<StreamRow>
  try {
    const globals = readGlobals(stream, options)
    const target = resolveTargetSheet(globals.boundSheets, options?.sheet)
    // An offset that misses a record boundary is an empty sheet to readXls too.
    if (!target || !recordStarts(stream, [target.pos]).has(target.pos)) return
    rows = inRowOrder(stream, target.pos)
      ? streamedRows(stream, target.pos, globals)
      : bufferedRows(stream, target.pos, globals, options?.maxTotalCells ?? MAX_TOTAL_CELLS)
  } catch (err) {
    throw asParseError(err)
  }

  const range = options?.range ? parseRangeFilter(options.range) : undefined
  const maxRows = typeof options?.maxRows === "number" && options.maxRows > 0 ? options.maxRows : 0
  let emitted = 0
  for (const row of rows) {
    // Rows ascend, so past the end of the range nothing more can match.
    if (range && row.index > range.endRow) return
    const filtered = range ? applyRangeFilter(row, range) : row
    if (!filtered) continue
    yield filtered
    if (maxRows > 0 && ++emitted >= maxRows) return
  }
}

/** The records that carry a cell value — the ones whose rows have to ascend for a sheet to stream. */
const VALUE_RECORDS: ReadonlySet<number> = new Set([
  SID.LABELSST,
  SID.RK,
  SID.NUMBER,
  SID.MULRK,
  SID.BOOLERR,
  SID.LABEL,
  SID.RSTRING,
  SID.FORMULA,
])

/**
 * Whether no value record goes back a row. Read from the record headers
 * alone, before anything is decoded, so a sheet that needs buffering is
 * known before its first row is yielded.
 */
function inRowOrder(stream: Uint8Array, pos: number): boolean {
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength)
  const cursor = new RecordCursor(stream, pos)
  cursor.skip() // the sheet's BOF
  let last = 0
  for (let at = cursor.skip(); at >= 0; at = cursor.skip()) {
    const id = view.getUint16(at, true)
    if (id === SID.EOF) break
    // A value record too short to name its row is malformed, which decoding it reports.
    if (!VALUE_RECORDS.has(id) || view.getUint16(at + 2, true) < 2 || at + 6 > stream.length) {
      continue
    }
    const row = view.getUint16(at + 4, true)
    if (row < last) return false
    last = row
  }
  return true
}

/** The sheet's rows as the records reach them: a row is complete when a later row's cell arrives. */
function* streamedRows(
  stream: Uint8Array,
  pos: number,
  globals: XlsGlobals,
): Generator<StreamRow, void, undefined> {
  try {
    const cursor = new RecordCursor(stream, pos)
    cursor.skip() // the sheet's BOF
    let index = -1
    let values: CellValue[] = []
    // One record's cells share a row, so a record completes at most one row.
    const ready: StreamRow[] = []
    const set = (row: number, col: number, value: CellValue): void => {
      assertInSheetBounds(row, col)
      if (row !== index) {
        if (index >= 0) ready.push({ index, values })
        index = row
        values = []
      }
      placeCell(values, col, value)
    }
    for (let rec = cursor.next(); rec; rec = cursor.next()) {
      if (rec.id === SID.EOF) break
      decodeCells(rec, cursor, globals, set)
      for (const row of ready) yield row
      ready.length = 0
    }
    if (index >= 0) yield { index, values }
  } catch (err) {
    throw asParseError(err)
  }
}

/**
 * A sheet whose records go back a row, held whole and then yielded in
 * order, so no row is yielded twice or split in two. Its cells are kept
 * per row rather than as a rectangle, but bounded as `readXls` bounds the
 * rectangle it would allocate.
 */
function* bufferedRows(
  stream: Uint8Array,
  pos: number,
  globals: XlsGlobals,
  cellLimit: number,
): Generator<StreamRow, void, undefined> {
  const byRow = new Map<number, CellValue[]>()
  try {
    const cursor = new RecordCursor(stream, pos)
    cursor.skip() // the sheet's BOF
    let rowCount = 0
    let widestCol = 0
    const set = (row: number, col: number, value: CellValue): void => {
      assertInSheetBounds(row, col)
      if (row >= rowCount) rowCount = row + 1
      if (col >= widestCol) widestCol = col + 1
      if (rowCount * widestCol > cellLimit) {
        throw new ParseError(
          `Sheet spans ${rowCount * widestCol} cells, over the ${cellLimit} limit. ` +
            "Raise `maxTotalCells` if the sheet really is this large.",
        )
      }
      let values = byRow.get(row)
      if (!values) byRow.set(row, (values = []))
      placeCell(values, col, value)
    }
    for (let rec = cursor.next(); rec; rec = cursor.next()) {
      if (rec.id === SID.EOF) break
      decodeCells(rec, cursor, globals, set)
    }
  } catch (err) {
    throw asParseError(err)
  }
  const ordered = [...byRow.entries()].sort((a, b) => a[0] - b[0])
  for (const [index, values] of ordered) yield { index, values }
}

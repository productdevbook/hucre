// ── Resource limits (DoS hardening) ─────────────────────────────────
// Central place for the bounds used to defend against malicious /
// malformed input (zip bombs, billion-laughs cell refs, etc.).

/**
 * Absolute hard cap on the number of bytes any single entry may
 * decompress to. Defends against zip bombs that claim a small
 * compressed size but expand to gigabytes. Default ~2 GiB.
 */
export const MAX_DECOMPRESSED_BYTES: number = 2 * 1024 * 1024 * 1024

/** Maximum row index (0-based) — Excel supports 1,048,576 rows. */
export const MAX_ROW_INDEX = 1_048_575

/** Maximum column index (0-based) — Excel supports 16,384 columns. */
export const MAX_COL_INDEX = 16_383

/**
 * Cap on the number of cells a single sheet may be normalized into.
 *
 * Readers hand back `rows: CellValue[][]`, a dense rectangle, so a sheet's
 * cost is its bounding box rather than its cell count. Bounding each
 * coordinate separately is not enough: two entirely legal cells at `A1`
 * and `XFD1048576` describe a 1,048,576 x 16,384 rectangle — 1.7e10 slots
 * — from a few hundred bytes of XML, which V8 answers with a fatal OOM
 * that no caller can catch.
 *
 * 20 million comfortably covers real spreadsheets (Excel's own row limit
 * at 20 columns is 21 million) while refusing the sparse-corner case with
 * a typed error instead of killing the process.
 */
export const MAX_TOTAL_CELLS = 20_000_000

/**
 * Cap on any `String.repeat` count taken from file content.
 *
 * ODS expresses runs of spaces as `<text:s text:c="N"/>` and decimal
 * precision as `number:decimal-places="N"`, both free-form integers. A
 * single `text:c="900000000"` reaches a raw `RangeError: Invalid string
 * length`; half that succeeds and allocates a gigabyte for one cell.
 *
 * 100,000 is far past any legitimate run of spaces or decimal places
 * while keeping the worst case at a few hundred KB.
 */
export const MAX_REPEAT_COUNT = 100_000

/**
 * Cap on the number of grid cells one HTML table cell may reserve
 * through `rowspan` x `colspan`.
 *
 * The reservation is a nested loop of Set insertions, so the cost is the
 * product — clamping each span on its own still leaves 16,384 x
 * 1,048,576. `fromHtml` exists to consume markup the caller did not
 * write, which makes this the cheapest hang in the library to trigger.
 */
export const MAX_SPAN_CELLS = 1_000_000

/**
 * Upper bound on the password-derivation spin count accepted from an
 * encrypted workbook. Office uses 100,000; we allow a generous ceiling
 * so a hostile file cannot pin a CPU for minutes.
 */
export const MAX_SPIN_COUNT = 10_000_000

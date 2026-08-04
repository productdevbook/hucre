// ── Resource limits (DoS hardening) ─────────────────────────────────
// Central place for the bounds used to defend against malicious /
// malformed input (zip bombs, billion-laughs cell refs, etc.).

/**
 * Absolute hard cap on the number of bytes any single entry may
 * decompress to. Defends against zip bombs that claim a small
 * compressed size but expand to gigabytes. Default ~2 GiB.
 */
export const MAX_DECOMPRESSED_BYTES: number = 2 * 1024 * 1024 * 1024

/**
 * Cap on the number of bytes buffered out of a `ReadableStream` input.
 *
 * Every format entry point funnels stream input through
 * `readInputToUint8Array`, which used to accumulate chunks with no ceiling
 * at all — so `read(response.body)` grew until V8 died, long before the
 * ZIP layer's {@link MAX_DECOMPRESSED_BYTES} could apply to anything.
 * Bounding each chunk would be useless here: chunks are individually tiny,
 * it is the running total that has to be checked.
 *
 * 1 GiB is far above any real workbook (the largest spreadsheets in the
 * wild are tens of MB, and a container this size already decompresses past
 * what the readers can hold), while keeping the peak cost of buffering —
 * chunks plus the single joined copy — inside a default Node heap. Callers
 * with genuinely larger input can raise it with `maxInputBytes`.
 */
export const MAX_INPUT_BYTES: number = 1024 * 1024 * 1024

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
 * Cap on the number of grid cells reserved through `rowspan` x `colspan`
 * while reading an HTML table — both by one cell and by the document as a
 * whole.
 *
 * The reservation is a nested loop of Set insertions, so the cost is the
 * product — clamping each span on its own still leaves 16,384 x
 * 1,048,576. `fromHtml` exists to consume markup the caller did not
 * write, which makes this the cheapest hang in the library to trigger.
 * The per-cell bound alone was not enough either: a page of three
 * thousand `rowspan="1000000"` cells pays it three thousand times over,
 * so the live reservation set is held to the same ceiling.
 */
export const MAX_SPAN_CELLS = 1_000_000

/**
 * Upper bound on the password-derivation spin count accepted from an
 * encrypted workbook. Office uses 100,000; we allow a generous ceiling
 * so a hostile file cannot pin a CPU for minutes.
 */
export const MAX_SPIN_COUNT = 10_000_000

// ── Read model → write model ─────────────────────────────────────────
//
// `readXlsx` returns a `Workbook`. `writeXlsx` takes `WriteOptions`. The
// two are not the same type and never were — read a file, change a cell,
// write it back, and `tsc` says:
//
//   Type 'Sheet[]' is not assignable to type 'WriteSheet[]'.
//     Types of property 'charts' are incompatible.
//       Property 'type' is missing in type 'Chart' but required in
//       type 'SheetChart'.
//
// and, once `charts` is stripped, the same again for `pivotTables`.
//
// Without this function every caller writes the same `map(({ charts,
// pivotTables, ...rest }) => rest)` by hand, and each one silently
// decides which fields to drop. The decisions are the same every time,
// so they belong here — stated once, in the open, with a way to see what
// went. See #439 §A.

import type { Sheet, Workbook, WriteOptions, WriteSheet } from "./_types"

/** What {@link toWriteOptions} could not carry, and from where. */
export interface WriteModelDrop {
  /** The field's name on `Sheet` or `Workbook`. */
  field: string
  /** The sheet it came from, or `undefined` for a workbook-level field. */
  sheet?: string
  /** Why the authoring model cannot take it. */
  reason: string
}

export interface ToWriteOptionsOptions {
  /**
   * Called once per field the authoring model cannot carry. Use it to log
   * or to refuse the conversion; the drops happen either way.
   */
  onDrop?: (drop: WriteModelDrop) => void
}

const NO_AUTHORING_API =
  "read and round-tripped only — there is no authoring API. Use openXlsx/saveXlsx to edit a file that has one."

/** Sheet fields with no `WriteSheet` counterpart at all. */
const SHEET_DROPS: Record<string, string> = {
  slicers: NO_AUTHORING_API,
  timelines: NO_AUTHORING_API,
  threadedComments: NO_AUTHORING_API,
  charts:
    "the read model (`Chart`) and the write model (`SheetChart`) are different types, and only 7 of the 16 readable kinds can be authored. Convert with `cloneChart` if you need them.",
  pivotTables:
    "`PivotTable` (what the reader produces) and `WritePivotTable` (what the writer takes) are disjoint by design — a pivot table cannot round-trip through the model.",
}

/** Workbook fields with no `WriteOptions` counterpart at all. */
const WORKBOOK_DROPS: Record<string, string> = {
  themeColors: "`writeXlsx` always emits the standard Office theme.",
  externalLinks: NO_AUTHORING_API,
  cellImages: NO_AUTHORING_API,
  persons: NO_AUTHORING_API,
  pivotCaches:
    "follows `pivotTables` — `WritePivotTable` builds its own cache from the source range.",
  slicerCaches: NO_AUTHORING_API,
  timelineCaches: NO_AUTHORING_API,
}

/**
 * Convert a `Workbook` read from a file into `WriteOptions` you can hand
 * to `writeXlsx` or `writeOds`.
 *
 * ```ts
 * const wb = await readXlsx(bytes)
 * wb.sheets[0].rows[0][0] = "edited"
 * const out = await writeXlsx(toWriteOptions(wb))
 * ```
 *
 * **This is the authoring path, so it is lossy by definition** — see
 * `docs/PARITY.md`. Everything the authoring model has no field for is
 * dropped, and `onDrop` is called once per field so the loss is
 * observable rather than silent:
 *
 * ```ts
 * toWriteOptions(wb, { onDrop: (d) => console.warn(d.field, d.reason) })
 * ```
 *
 * If you are **editing someone else's file** rather than producing a new
 * one, use `openXlsx` / `saveXlsx` instead: that path copies the parts
 * hucre does not regenerate byte-for-byte, so none of this applies.
 */
export function toWriteOptions(workbook: Workbook, options?: ToWriteOptionsOptions): WriteOptions {
  const onDrop = options?.onDrop

  const out: WriteOptions = {
    sheets: workbook.sheets.map((sheet) => toWriteSheet(sheet, onDrop)),
  }

  if (workbook.properties) out.properties = workbook.properties
  if (workbook.namedRanges) out.namedRanges = workbook.namedRanges
  if (workbook.defaultFont) out.defaultFont = workbook.defaultFont
  if (workbook.dateSystem) out.dateSystem = workbook.dateSystem
  if (workbook.activeSheet !== undefined) out.activeSheet = workbook.activeSheet
  if (workbook.workbookProtection) out.workbookProtection = workbook.workbookProtection

  if (onDrop) {
    const held = workbook as unknown as Record<string, unknown>
    for (const [field, reason] of Object.entries(WORKBOOK_DROPS)) {
      if (held[field] !== undefined) onDrop({ field, reason })
    }
  }

  return out
}

/**
 * The per-sheet half of {@link toWriteOptions}. Exported because a caller
 * assembling a workbook from sheets of several sources needs it on its
 * own.
 */
export function toWriteSheet(sheet: Sheet, onDrop?: (drop: WriteModelDrop) => void): WriteSheet {
  const {
    slicers: _slicers,
    timelines: _timelines,
    threadedComments: _threadedComments,
    charts: _charts,
    pivotTables: _pivotTables,
    // `Sheet.cells` is `Map<string, Cell>` and `WriteSheet.cells` is
    // `Map<string, Partial<Cell>>` — every `Cell` is a valid `Partial<Cell>`,
    // so this one carries across as it is.
    ...rest
  } = sheet

  if (onDrop) {
    const held = sheet as unknown as Record<string, unknown>
    for (const [field, reason] of Object.entries(SHEET_DROPS)) {
      const value = held[field]
      if (value !== undefined && (!Array.isArray(value) || value.length > 0)) {
        onDrop({ field, sheet: sheet.name, reason })
      }
    }
  }

  return rest
}

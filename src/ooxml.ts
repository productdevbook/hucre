// ── hucre/ooxml entry point ───────────────────────────────────────────
//
// Low-level parsers for individual OOXML parts. These take a raw XML
// string from inside an .xlsx package — `xl/charts/chart1.xml`,
// `xl/pivotTables/pivotTable1.xml`, and so on — and return hucre's
// internal model of it.
//
// ## Stability
//
// **This entry point is explicitly excluded from hucre's v1 stability
// commitment.** Its shapes mirror the OOXML parse pipeline, so they move
// when that pipeline moves. Freezing them at v1 would mean the internals
// could never change without a major bump.
//
// The rest of the library — `hucre`, `hucre/xlsx`, `hucre/csv`,
// `hucre/ods`, `hucre/json`, `hucre/xml` — is stable. If you only need
// to read or write spreadsheets, you do not need anything here.
//
// The raw-XML parsers below are exported from here only. The chart
// helpers that take a model (`cloneChart`, `addChart`, `getCharts`) are
// also on the root, because they are not part of the parse pipeline.

// ── Charts ─────────────────────────────────────────────────────────
export { parseChart } from "./xlsx/chart-reader"
export { cloneChart, chartKindToWriteKind } from "./xlsx/chart-clone"
export type { CloneChartOptions, CloneChartSeriesOverride } from "./xlsx/chart-clone"
export { addChart, getCharts } from "./xlsx/chart-helpers"
export type { ChartLocation } from "./xlsx/chart-helpers"

// ── Pivot tables ───────────────────────────────────────────────────
export {
  parsePivotTable,
  parsePivotCacheDefinition,
  attachPivotCacheFields,
} from "./xlsx/pivot-reader"

// ── Slicers and timelines ──────────────────────────────────────────
export {
  parseSlicers,
  parseSlicerCache,
  parseTimelines,
  parseTimelineCache,
} from "./xlsx/slicer-reader"

// ── Comments, links, images, theme ─────────────────────────────────
export { parsePersons, parseThreadedComments } from "./xlsx/threaded-comments-reader"
export { parseExternalLink } from "./xlsx/external-link-reader"
export { parseCellImages, assembleCellImages, REL_CELL_IMAGES } from "./xlsx/cell-images-reader"
export type { ParsedCellImageRef } from "./xlsx/cell-images-reader"
export { parseThemeColors, resolveThemeColor } from "./xlsx/theme"

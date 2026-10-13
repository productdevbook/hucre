// Normalize once at the writer boundary. Keeping metadata on one model
// removes the per-field copies that previously lost new Sheet fields.
import { validateRowSize } from "./_validate"
import { normalizeSheetInput } from "./_sheet-input"
import type {
  SheetInput,
  SheetChart,
  WritePivotTable,
  WorkbookInput,
  WritableSheet,
  WritableWorkbook,
  WriteModelDrop,
} from "./_types"

const PRESERVE_PARTS =
  "Read and preserved through openXlsx/saveXlsx; the authoring writer cannot rebuild this feature."

const SHEET_DROPS = {
  kind: "The authoring writer creates worksheets; non-worksheet tabs need openXlsx/saveXlsx.",
  slicers: PRESERVE_PARTS,
  timelines: PRESERVE_PARTS,
  threadedComments: PRESERVE_PARTS,
} as const

const WORKBOOK_DROPS = {
  themeColors: "The authoring writer emits the standard Office theme.",
  externalLinks: PRESERVE_PARTS,
  cellImages: PRESERVE_PARTS,
  persons: PRESERVE_PARTS,
  pivotCaches: "Authored pivot tables build their own cache from source data.",
  slicerCaches: PRESERVE_PARTS,
  timelineCaches: PRESERVE_PARTS,
} as const

// Compile-time exhaustive: a new model field needs an ODS capability
// decision. ODS authoring carries values, cell styles/formulas, column
// value/style definitions, merges, properties and named ranges.
const ODS_SHEET_DROPS: Record<
  Exclude<keyof SheetInput, "name" | "rows" | "data" | "cells" | "columns" | "merges">,
  true
> = {
  kind: true,
  rowDefs: true,
  defaultRowHeight: true,
  defaultColWidth: true,
  dataValidations: true,
  conditionalRules: true,
  autoFilter: true,
  freezePane: true,
  splitPane: true,
  images: true,
  protection: true,
  pageSetup: true,
  headerFooter: true,
  view: true,
  hidden: true,
  veryHidden: true,
  tables: true,
  rowBreaks: true,
  colBreaks: true,
  outlineProperties: true,
  backgroundImage: true,
  sparklines: true,
  textBoxes: true,
  threadedComments: true,
  a11y: true,
  slicers: true,
  timelines: true,
  charts: true,
  pivotTables: true,
}
const ODS_WORKBOOK_DROPS: Record<
  Exclude<keyof WorkbookInput, "sheets" | "properties" | "namedRanges">,
  true
> = {
  dateSystem: true,
  defaultFont: true,
  activeSheet: true,
  themeColors: true,
  workbookProtection: true,
  persons: true,
  externalLinks: true,
  cellImages: true,
  pivotCaches: true,
  slicerCaches: true,
  timelineCaches: true,
}

const ODS_REASON =
  "The ODS authoring writer does not rebuild this metadata. Use XLSX part preservation when this feature must survive."

type OnDrop = ((drop: WriteModelDrop) => void) | undefined

function populated(value: unknown): boolean {
  return value !== undefined && (!Array.isArray(value) || value.length > 0)
}

export function prepareSheet(sheet: SheetInput, onDrop?: OnDrop): WritableSheet {
  // Object expansion is checked before lookups; positional rows are
  // checked before lifting rich cells or invoking authoring-loss callbacks.
  validateRowSize(Math.max(0, (sheet.rows?.length ?? 0) - 1), sheet.columns?.length ?? 0)
  for (const row of sheet.rows ?? []) validateRowSize(0, row.length)
  sheet = normalizeSheetInput(sheet, (height, width) =>
    validateRowSize(Math.max(0, height - 1), width),
  )
  const {
    kind,
    slicers: _slicers,
    timelines: _timelines,
    threadedComments: _threadedComments,
    charts,
    pivotTables,
    ...rest
  } = sheet
  for (const [field, reason] of Object.entries(SHEET_DROPS)) {
    if (field === "kind" && (!kind || kind === "worksheet")) continue
    if (populated(sheet[field as keyof typeof SHEET_DROPS])) {
      onDrop?.({ field, sheet: sheet.name, reason })
    }
  }
  const authoredCharts = charts?.filter(
    (chart): chart is SheetChart => chart !== null && typeof chart === "object" && "type" in chart,
  )
  const authoredPivots = pivotTables?.filter(
    (pivot): pivot is WritePivotTable =>
      pivot !== null && typeof pivot === "object" && "values" in pivot,
  )
  if (charts && authoredCharts?.length !== charts.length) {
    onDrop?.({
      field: "charts",
      sheet: sheet.name,
      reason:
        "Reader Chart records are inspection data. Supply a SheetChart or preserve the original parts with openXlsx/saveXlsx.",
    })
  }
  if (pivotTables && authoredPivots?.length !== pivotTables.length) {
    onDrop?.({
      field: "pivotTables",
      sheet: sheet.name,
      reason:
        "Reader PivotTable records have no source data. Supply a WritePivotTable or use openXlsx/saveXlsx.",
    })
  }
  return {
    ...rest,
    ...(authoredCharts?.length ? { charts: authoredCharts } : {}),
    ...(authoredPivots?.length ? { pivotTables: authoredPivots } : {}),
  }
}

export function prepareWorkbook(workbook: WorkbookInput, onDrop?: OnDrop): WritableWorkbook {
  const {
    sheets,
    themeColors: _themeColors,
    externalLinks: _externalLinks,
    cellImages: _cellImages,
    persons: _persons,
    pivotCaches: _pivotCaches,
    slicerCaches: _slicerCaches,
    timelineCaches: _timelineCaches,
    ...rest
  } = workbook
  for (const [field, reason] of Object.entries(WORKBOOK_DROPS)) {
    if (populated(workbook[field as keyof WorkbookInput])) {
      onDrop?.({ field, reason })
    }
  }
  return {
    ...rest,
    sheets: sheets.map((sheet) => prepareSheet(sheet, onDrop)),
  }
}

// A separate entry lets XLSX-only bundles omit the ODS capability
// register while both writers still share the same model normalization.
export function prepareOdsWorkbook(workbook: WorkbookInput, onDrop?: OnDrop): WritableWorkbook {
  for (const field of Object.keys(ODS_WORKBOOK_DROPS)) {
    if (populated(workbook[field as keyof typeof ODS_WORKBOOK_DROPS])) {
      onDrop?.({ field, reason: ODS_REASON })
    }
  }
  for (const sheet of workbook.sheets) {
    for (const field of Object.keys(ODS_SHEET_DROPS)) {
      if (field === "kind" && (!sheet.kind || sheet.kind === "worksheet")) continue
      if (populated(sheet[field as keyof typeof ODS_SHEET_DROPS])) {
        onDrop?.({ field, sheet: sheet.name, reason: ODS_REASON })
      }
    }
  }
  return prepareWorkbook(workbook)
}

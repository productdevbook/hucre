// ── Chart Helpers ────────────────────────────────────────────────────
// High-level convenience helpers around the chart read/write data
// model. `getCharts(workbook)` enumerates every chart anchored on the
// workbook's sheets, in workbook order, with its sheet context attached
// so callers don't have to walk `workbook.sheets[].charts` themselves.
// `addChart(sheet, chart)` pushes a `SheetChart` onto a `WriteSheet`'s
// chart list, creating the array on the fly.
//
// These mirror the `getCharts(workbook)` / `addChart(sheet, ...)`
// shorthand sketched in the dashboard composition issue (#136).

import type { Chart, Sheet, SheetChart, WriteSheet, Workbook } from "../_types";

// ── getCharts ────────────────────────────────────────────────────────

/**
 * One entry returned by {@link getCharts}. Carries enough sheet context
 * (name + 0-based index) for callers to look up siblings on the same
 * sheet without walking the workbook a second time.
 *
 * `chartIndex` is 0-based inside the sheet's `charts` array, in the
 * order the chart parts were resolved off the drawing — typically the
 * authoring order in Excel.
 */
export interface ChartLocation {
  /** Reference to the sheet that owns the chart. Same identity as `workbook.sheets[sheetIndex]`. */
  sheet: Sheet;
  /** Sheet name as declared in `xl/workbook.xml`. */
  sheetName: string;
  /** 0-based position of the sheet inside `workbook.sheets`. */
  sheetIndex: number;
  /** The parsed chart record. Same identity as `sheet.charts[chartIndex]`. */
  chart: Chart;
  /** 0-based position of the chart inside `sheet.charts`. */
  chartIndex: number;
}

/**
 * Enumerate every chart anchored on the workbook's sheets.
 *
 * Visits sheets in workbook order; within a sheet, visits charts in
 * the order surfaced by the reader. Sheets without charts are skipped.
 * Returns an empty array when the workbook has no charts at all.
 *
 * @example
 * ```ts
 * import { openXlsx, getCharts } from "hucre";
 *
 * const wb = await openXlsx(bytes);
 * for (const { sheetName, chart } of getCharts(wb)) {
 *   console.log(sheetName, chart.kinds, chart.title);
 * }
 * ```
 */
export function getCharts(workbook: Workbook): ChartLocation[] {
  const out: ChartLocation[] = [];
  for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex++) {
    const sheet = workbook.sheets[sheetIndex];
    const charts = sheet.charts;
    if (!charts || charts.length === 0) continue;
    for (let chartIndex = 0; chartIndex < charts.length; chartIndex++) {
      out.push({
        sheet,
        sheetName: sheet.name,
        sheetIndex,
        chart: charts[chartIndex],
        chartIndex,
      });
    }
  }
  return out;
}

// ── addChart ─────────────────────────────────────────────────────────

/**
 * Append a {@link SheetChart} to a {@link WriteSheet}'s `charts` list,
 * lazily creating the array on the first call. Returns the same chart
 * object so callers can inline declarations:
 *
 * @example
 * ```ts
 * import { addChart, writeXlsx } from "hucre";
 *
 * const sheet = {
 *   name: "Dashboard",
 *   rows: [
 *     ["Quarter", "Revenue"],
 *     ["Q1", 12000],
 *     ["Q2", 15500],
 *   ],
 * };
 *
 * addChart(sheet, {
 *   type: "column",
 *   title: "Revenue",
 *   series: [{ name: "Revenue", values: "B2:B3", categories: "A2:A3" }],
 *   anchor: { from: { row: 5, col: 0 } },
 * });
 *
 * await writeXlsx({ sheets: [sheet] });
 * ```
 *
 * Equivalent to:
 *
 * ```ts
 * (sheet.charts ??= []).push(chart);
 * ```
 */
export function addChart(sheet: WriteSheet, chart: SheetChart): SheetChart {
  if (!chart || typeof chart !== "object") {
    throw new TypeError("addChart: chart is required");
  }
  if (!chart.type) {
    throw new TypeError("addChart: chart.type is required");
  }
  if (!Array.isArray(chart.series) || chart.series.length === 0) {
    throw new TypeError("addChart: chart.series must contain at least one entry");
  }
  if (!chart.anchor || !chart.anchor.from) {
    throw new TypeError("addChart: chart.anchor.from is required");
  }
  const list = sheet.charts ?? (sheet.charts = []);
  list.push(chart);
  return chart;
}

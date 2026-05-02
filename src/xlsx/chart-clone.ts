// ── Chart Clone ──────────────────────────────────────────────────────
// Bridges the read-side `Chart` metadata produced by `parseChart` to the
// write-side `SheetChart` shape consumed by `writeXlsx`.
//
// Use case (issue #136): a template workbook stores one of each chart
// flavor; at export time the caller pulls a chart out, swaps its data
// ranges and labels, and re-emits it (often several times) into a new
// workbook. The two type families overlap — `ChartSeriesInfo` already
// mirrors `ChartSeries` — but the read side has no anchor and supports
// kinds the write side cannot author yet, so a dedicated converter
// keeps the type-narrowing explicit.

import type {
  Chart,
  ChartDataLabels,
  ChartDataLabelsInfo,
  ChartKind,
  ChartSeries,
  ChartSeriesInfo,
  SheetChart,
  WriteChartKind,
} from "../_types";

// ── Public API ───────────────────────────────────────────────────────

/**
 * Per-series override applied on top of the source chart's series.
 *
 * Each field defaults to the value carried by the source series at the
 * matching position. Pass `null` to drop the source value entirely
 * (e.g. `color: null` removes a series tint inherited from the
 * template).
 */
export interface CloneChartSeriesOverride {
  name?: string | null;
  /** A1 range for `<c:val>` / `<c:yVal>`. Required when the source has none. */
  values?: string;
  /** A1 range for `<c:cat>` / `<c:xVal>`. */
  categories?: string | null;
  /** 6-digit RGB hex (e.g. `"1F77B4"`). */
  color?: string | null;
  /**
   * Per-series data label override. `undefined` (or omitted) inherits
   * the source series' `dataLabels`; `null` drops the inherited block;
   * `false` suppresses labels for this series alone (overriding any
   * chart-level default); a `ChartDataLabels` object replaces the
   * inherited block wholesale.
   */
  dataLabels?: ChartDataLabels | false | null;
}

/**
 * Options accepted by {@link cloneChart}.
 *
 * `anchor` is required because the read-side `Chart` does not capture
 * placement — drawings live in a separate part. Every other field
 * defaults to the source chart.
 */
export interface CloneChartOptions {
  /**
   * Cell anchor for the cloned chart. `to` defaults to a 6×15 area
   * below `from`, mirroring `SheetChart.anchor`.
   */
  anchor: SheetChart["anchor"];
  /**
   * Override the chart family. When omitted, the source's first
   * write-compatible kind is used. An explicit value lets callers
   * narrow a combo chart down to one renderable type or coerce a
   * `doughnut` template into a plain `pie`.
   */
  type?: WriteChartKind;
  /** Override the chart title. Pass `null` to drop the source title. */
  title?: string | null;
  /** Replace the entire series array (skips per-series merging). */
  series?: ChartSeries[];
  /**
   * Per-series overrides. Indices line up with the source's
   * {@link Chart.series}. Use this to remap data ranges without
   * rewriting every other field.
   */
  seriesOverrides?: ReadonlyArray<CloneChartSeriesOverride | undefined>;
  /** Override `SheetChart.legend`. */
  legend?: SheetChart["legend"];
  /** Override `SheetChart.barGrouping`. */
  barGrouping?: SheetChart["barGrouping"];
  /** Override `SheetChart.showTitle`. */
  showTitle?: boolean;
  /** Override `SheetChart.altText`. */
  altText?: string;
  /** Override `SheetChart.frameTitle`. */
  frameTitle?: string;
  /**
   * Override the chart-level data labels. `undefined` (or omitted)
   * inherits the source's `dataLabels`; `null` drops the inherited
   * block; a `ChartDataLabels` object replaces it.
   */
  dataLabels?: ChartDataLabels | null;
  /**
   * Per-axis title overrides. Each field accepts a string to replace,
   * or `null` to drop the source value (the cloned chart will render
   * without that axis label even if the template carried one). Omit a
   * field to inherit the source.
   *
   * Ignored when the resolved chart type is `pie` since pie has no
   * axes; the writer drops the entire `axes` object in that case.
   */
  axes?: {
    x?: { title?: string | null };
    y?: { title?: string | null };
  };
}

/**
 * Convert a parsed {@link Chart} into a {@link SheetChart} ready for
 * `writeXlsx`. Series formula references (`valuesRef`, `categoriesRef`)
 * become `values` / `categories` on the new chart; per-series colors
 * carry over.
 *
 * @throws {Error} when the source chart kinds cannot be authored on
 *   the write side and no `options.type` override is provided.
 * @throws {Error} when a non-overridden series has no `valuesRef` —
 *   `SheetChart.series[].values` is mandatory.
 *
 * @example
 * ```ts
 * import { parseChart, cloneChart } from "hucre";
 *
 * const source = parseChart(templateChartXml)!;
 * const clone = cloneChart(source, {
 *   anchor: { from: { row: 14, col: 0 } },
 *   title: "Revenue",
 *   seriesOverrides: [{ values: "Dashboard!$B$2:$B$13", color: "1070CA" }],
 * });
 * ```
 */
export function cloneChart(source: Chart, options: CloneChartOptions): SheetChart {
  if (!options || !options.anchor) {
    throw new Error("cloneChart: options.anchor is required");
  }

  const type = options.type ?? pickWritableKind(source);

  // Pick a base title: explicit override (including `null` meaning drop)
  // wins over the source's title.
  const title = resolveTitle(source.title, options.title);

  // Build the series array.
  let series: ChartSeries[];
  if (options.series) {
    series = options.series.map((s) => ({ ...s }));
  } else {
    series = buildSeriesFromSource(source, options.seriesOverrides);
  }

  if (series.length === 0) {
    throw new Error(
      "cloneChart: produced 0 series; pass `series` or ensure the source has at least one series with a valuesRef",
    );
  }

  const out: SheetChart = {
    type,
    series,
    anchor: options.anchor,
  };
  if (title !== undefined) out.title = title;

  // Legend / bar grouping carry over from the source when the caller
  // does not supply an override. Bar grouping only round-trips for
  // bar/column targets — applying a stacked grouping to a line/pie
  // template clone would be silently ignored by the writer.
  const legend = options.legend !== undefined ? options.legend : source.legend;
  if (legend !== undefined) out.legend = legend;

  const barGrouping = options.barGrouping !== undefined ? options.barGrouping : source.barGrouping;
  if (barGrouping !== undefined && (type === "bar" || type === "column")) {
    out.barGrouping = barGrouping;
  }

  if (options.showTitle !== undefined) out.showTitle = options.showTitle;
  if (options.altText !== undefined) out.altText = options.altText;
  if (options.frameTitle !== undefined) out.frameTitle = options.frameTitle;

  const resolvedDataLabels = resolveChartDataLabels(source.dataLabels, options.dataLabels);
  if (resolvedDataLabels !== undefined) out.dataLabels = resolvedDataLabels;

  // Pie has no axes, so silently skip carrying over axis titles even
  // when the source declared them or the caller passed an override.
  if (type !== "pie") {
    const axes = resolveAxes(source.axes, options.axes);
    if (axes !== undefined) out.axes = axes;
  }

  return out;
}

// ── Internals ────────────────────────────────────────────────────────

/**
 * Map a read-side {@link ChartKind} to the writer's
 * {@link WriteChartKind}, or `undefined` when no equivalent exists.
 *
 * The writer authors the six families covered in chart writer Phase 1
 * (issue #152). 3D variants collapse onto their 2D counterparts;
 * `doughnut` collapses to `pie`. Kinds with no analog (`bubble`,
 * `radar`, `surface`, `stock`, `ofPie`) return `undefined` and force
 * the caller to pass an explicit `type` override.
 */
export function chartKindToWriteKind(kind: ChartKind): WriteChartKind | undefined {
  switch (kind) {
    case "bar":
    case "bar3D":
      // Read-side `bar` covers both `<c:barChart barDir="bar">` and
      // `<c:barChart barDir="col">`; the parser does not split them.
      // Default to the more common vertical orientation; callers who
      // need horizontal pass `type: "bar"` explicitly.
      return "column";
    case "line":
    case "line3D":
      return "line";
    case "pie":
    case "pie3D":
    case "doughnut":
      return "pie";
    case "area":
    case "area3D":
      return "area";
    case "scatter":
      return "scatter";
    case "bubble":
    case "radar":
    case "surface":
    case "surface3D":
    case "stock":
    case "ofPie":
      return undefined;
  }
}

function pickWritableKind(source: Chart): WriteChartKind {
  if (source.kinds.length === 0) {
    throw new Error("cloneChart: source chart has no kinds; pass `options.type` explicitly");
  }
  for (const k of source.kinds) {
    const mapped = chartKindToWriteKind(k);
    if (mapped) return mapped;
  }
  throw new Error(
    `cloneChart: source kind${source.kinds.length > 1 ? "s" : ""} ${source.kinds
      .map((k) => `"${k}"`)
      .join(
        ", ",
      )} cannot be authored on the write side; pass \`options.type\` to coerce a renderable kind`,
  );
}

function resolveTitle(
  sourceTitle: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return sourceTitle;
  if (override === null) return undefined;
  return override;
}

function buildSeriesFromSource(
  source: Chart,
  overrides: ReadonlyArray<CloneChartSeriesOverride | undefined> | undefined,
): ChartSeries[] {
  const sourceSeries = source.series ?? [];
  // The override array can be longer than the source (caller wants to
  // append a fully-specified series). Walk the union of both lengths.
  const length = Math.max(sourceSeries.length, overrides?.length ?? 0);
  const out: ChartSeries[] = [];

  for (let i = 0; i < length; i++) {
    const src: ChartSeriesInfo | undefined = sourceSeries[i];
    const ov = overrides?.[i];
    const merged = mergeSeries(src, ov, i);
    out.push(merged);
  }

  return out;
}

function mergeSeries(
  src: ChartSeriesInfo | undefined,
  ov: CloneChartSeriesOverride | undefined,
  index: number,
): ChartSeries {
  // Resolve `values` first — it's the only mandatory field.
  const values = ov?.values ?? src?.valuesRef;
  if (!values) {
    throw new Error(
      `cloneChart: series #${index} has no values reference; provide \`seriesOverrides[${index}].values\``,
    );
  }

  const out: ChartSeries = { values };

  const name = applyOverride(src?.name, ov?.name);
  if (name !== undefined) out.name = name;

  const categories = applyOverride(src?.categoriesRef, ov?.categories);
  if (categories !== undefined) out.categories = categories;

  const color = applyOverride(src?.color, ov?.color);
  if (color !== undefined) out.color = color;

  const dataLabels = resolveSeriesDataLabels(src?.dataLabels, ov?.dataLabels);
  if (dataLabels !== undefined) out.dataLabels = dataLabels;

  return out;
}

/**
 * Resolve a chart-level data-labels override.
 *
 * `undefined` → inherit the source's parsed `dataLabels` (downcast from
 * the read-side {@link ChartDataLabelsInfo} to the write-side
 * {@link ChartDataLabels} shape — they share field semantics).
 * `null`      → drop the inherited block.
 * object      → replace.
 */
function resolveChartDataLabels(
  sourceLabels: ChartDataLabelsInfo | undefined,
  override: ChartDataLabels | null | undefined,
): ChartDataLabels | undefined {
  if (override === undefined) {
    return sourceLabels ? { ...sourceLabels } : undefined;
  }
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a per-series data-labels override.
 *
 * `undefined` → inherit the source series' `dataLabels`.
 * `null`      → drop the inherited block (series will fall back to
 *               whatever the chart-level default is at write time).
 * `false`     → suppress labels on this series alone.
 * object      → replace.
 */
function resolveSeriesDataLabels(
  sourceLabels: ChartDataLabelsInfo | undefined,
  override: ChartDataLabels | false | null | undefined,
): ChartDataLabels | false | undefined {
  if (override === undefined) {
    return sourceLabels ? { ...sourceLabels } : undefined;
  }
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a "source value + optional override" pair where the override
 * may be `undefined` (no override), `null` (drop the source value), or
 * a string (replace).
 */
function applyOverride(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Merge the source chart's `axes` block with per-axis overrides. The
 * result mirrors the writer's {@link SheetChart.axes} shape — missing
 * fields are dropped so the writer doesn't emit empty `<c:title>`
 * elements.
 */
function resolveAxes(
  sourceAxes: Chart["axes"],
  overrides: CloneChartOptions["axes"],
): SheetChart["axes"] | undefined {
  const xTitle = applyOverride(sourceAxes?.x?.title, overrides?.x?.title);
  const yTitle = applyOverride(sourceAxes?.y?.title, overrides?.y?.title);

  const out: NonNullable<SheetChart["axes"]> = {};
  if (xTitle !== undefined) out.x = { title: xTitle };
  if (yTitle !== undefined) out.y = { title: yTitle };

  return out.x || out.y ? out : undefined;
}

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
  ChartAxisGridlines,
  ChartAxisLabelAlign,
  ChartAxisNumberFormat,
  ChartAxisScale,
  ChartAxisTickLabelPosition,
  ChartAxisTickMark,
  ChartDataLabels,
  ChartDataLabelsInfo,
  ChartDisplayBlanksAs,
  ChartKind,
  ChartLineStroke,
  ChartMarker,
  ChartScatterStyle,
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
  /**
   * Smoothed-line override. `undefined` (or omitted) inherits the source
   * series' `smooth`; `null` drops the inherited flag (the cloned series
   * renders straight); a `boolean` replaces it wholesale. Only meaningful
   * for `line` and `scatter` clones — silently dropped from the output
   * when the resolved chart type is anything else.
   */
  smooth?: boolean | null;
  /**
   * Line stroke override. `undefined` (or omitted) inherits the source
   * series' `stroke`; `null` drops the inherited block (the cloned
   * series falls back to Excel's per-series default); a
   * {@link ChartLineStroke} object replaces the inherited block
   * wholesale (no per-field merge — pass the full shape you want).
   * Only meaningful for `line` and `scatter` clones — silently dropped
   * from the output when the resolved chart type is anything else.
   */
  stroke?: ChartLineStroke | null;
  /**
   * Marker override. `undefined` (or omitted) inherits the source
   * series' `marker`; `null` drops the inherited block (the cloned
   * series falls back to Excel's series-rotation default); a
   * {@link ChartMarker} object replaces the inherited block wholesale
   * (no per-field merging — pass every field you want preserved).
   * Only meaningful for `line` and `scatter` clones — silently dropped
   * from the output when the resolved chart type is anything else.
   */
  marker?: ChartMarker | null;
  /**
   * Invert-if-negative override. `undefined` (or omitted) inherits the
   * source series' `invertIfNegative`; `null` drops the inherited flag
   * (the cloned series renders negatives in the series fill color);
   * a `boolean` replaces it wholesale. Only meaningful for `bar` and
   * `column` clones — silently dropped from the output when the
   * resolved chart type is anything else.
   */
  invertIfNegative?: boolean | null;
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
   * narrow a combo chart down to one renderable type or flatten a
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
  /**
   * Override the chart-level legend-overlay flag. `undefined` (or
   * omitted) inherits the source's parsed value; `null` drops the
   * inherited value (the writer falls back to the OOXML `false` default
   * — the legend reserves its own slot, no overlap with the plot area);
   * a `boolean` replaces it.
   *
   * The override is silently dropped from the cloned `SheetChart` when
   * the resolved legend is `false` (no legend element will be emitted)
   * — there is no overlay flag to set on a hidden legend, so leaking
   * the value into the output would carry a toggle Excel never reads.
   */
  legendOverlay?: boolean | null;
  /** Override `SheetChart.barGrouping`. */
  barGrouping?: SheetChart["barGrouping"];
  /**
   * Override `SheetChart.gapWidth` (only meaningful for `bar` /
   * `column`). Dropped silently when the resolved chart type is
   * neither — a gap-width hint inherited from a column template never
   * leaks into a line / pie clone.
   */
  gapWidth?: number;
  /**
   * Override `SheetChart.overlap` (only meaningful for `bar` /
   * `column`). Dropped silently when the resolved chart type is
   * neither.
   */
  overlap?: number;
  /** Override `SheetChart.lineGrouping`. */
  lineGrouping?: SheetChart["lineGrouping"];
  /** Override `SheetChart.areaGrouping`. */
  areaGrouping?: SheetChart["areaGrouping"];
  /**
   * Override `SheetChart.holeSize` (only meaningful for `doughnut`).
   * When the resolved chart type is not `doughnut`, the field is
   * dropped from the output so it does not leak into a cloned pie or
   * column chart.
   */
  holeSize?: number;
  /**
   * Override `SheetChart.firstSliceAng` (the pie / doughnut starting
   * angle in degrees, clockwise from 12 o'clock). Only meaningful for
   * `pie` and `doughnut`; dropped silently when the resolved chart
   * type is anything else, so a rotation hint inherited from a
   * doughnut template never leaks into a column or scatter clone.
   */
  firstSliceAng?: number;
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
   * Override how the chart renders missing / blank cells. `undefined`
   * (or omitted) inherits the source's `dispBlanksAs`; `null` drops
   * the inherited value (the writer falls back to the OOXML `"gap"`
   * default); a {@link ChartDisplayBlanksAs} value replaces it. Useful
   * when a template uses `"span"` to bridge gaps but the cloned
   * dashboard chart should render the gaps explicitly (or vice versa).
   */
  dispBlanksAs?: ChartDisplayBlanksAs | null;
  /**
   * Override `<c:varyColors>` (the per-point unique-color toggle).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * `varyColors`. `null` drops the inherited value so the writer falls
   * back to the per-family default (`true` for pie / doughnut, `false`
   * everywhere else). A `boolean` replaces it — useful for collapsing
   * a doughnut to a single color (`false`) or painting each bar of a
   * single-series column chart in a different color (`true`).
   */
  varyColors?: boolean | null;
  /**
   * Override `<c:plotVisOnly>` (the "hide hidden cells" toggle).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * `plotVisOnly`. `null` drops the inherited value so the writer
   * falls back to the OOXML `true` default (hidden cells drop out of
   * the chart). A `boolean` replaces it — useful for keeping hidden
   * helper rows in the rendered chart (`false`) or restoring the
   * default behavior on a clone whose template overrode it (`true`).
   *
   * The grammar mirrors `dispBlanksAs` / `varyColors` so the
   * chart-level toggles compose the same way at the call site.
   */
  plotVisOnly?: boolean | null;
  /**
   * Override `<c:roundedCorners>` (the chart-frame rounded-edge toggle).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * `roundedCorners`. `null` drops the inherited value so the writer
   * falls back to the OOXML `false` default (square chart frame). A
   * `boolean` replaces it — useful for matching a dashboard whose
   * other charts already carry the rounded look from a template, or
   * for squaring off a clone whose template was rounded.
   *
   * The grammar mirrors `plotVisOnly` / `varyColors` so the
   * chart-frame toggles compose the same way at the call site.
   */
  roundedCorners?: boolean | null;
  /**
   * Override `<c:scatterStyle>` (the chart-level XY-scatter preset).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * `scatterStyle`. `null` drops the inherited value so the writer
   * falls back to its `"lineMarker"` default. A {@link ChartScatterStyle}
   * value replaces it — useful when a smoothed-line scatter template
   * should clone as a marker-only or straight-line variant.
   *
   * Only meaningful when the resolved chart type is `scatter`; the
   * field is silently dropped when the clone targets any other family
   * since the OOXML schema places `<c:scatterStyle>` exclusively on
   * `<c:scatterChart>`.
   */
  scatterStyle?: ChartScatterStyle | null;
  /**
   * Per-axis overrides. Each field accepts a value to replace the
   * source's, or `null` to drop the source value (the cloned chart
   * will render without that axis label / gridline even if the
   * template carried one). Omit a field to inherit the source.
   *
   * Ignored when the resolved chart type is `pie` or `doughnut` since
   * neither has axes; the writer drops the entire `axes` object in
   * those cases.
   */
  axes?: {
    x?: {
      title?: string | null;
      gridlines?: ChartAxisGridlines | null;
      scale?: ChartAxisScale | null;
      numberFormat?: ChartAxisNumberFormat | null;
      /**
       * Override the major tick-mark style. `undefined` (or omitted)
       * inherits the source axis' parsed value; `null` drops it (the
       * writer falls back to the OOXML default `"out"`); a value
       * replaces it.
       */
      majorTickMark?: ChartAxisTickMark | null;
      /**
       * Override the minor tick-mark style. `undefined` (or omitted)
       * inherits the source axis' parsed value; `null` drops it (the
       * writer falls back to the OOXML default `"none"`); a value
       * replaces it.
       */
      minorTickMark?: ChartAxisTickMark | null;
      /**
       * Override the tick-label position. `undefined` (or omitted)
       * inherits the source axis' parsed value; `null` drops it (the
       * writer falls back to the OOXML default `"nextTo"`); a value
       * replaces it.
       */
      tickLblPos?: ChartAxisTickLabelPosition | null;
      /**
       * Override the reverse-axis flag. `undefined` (or omitted)
       * inherits the source axis' parsed value; `null` drops it (the
       * writer falls back to the OOXML default `"minMax"` — forward
       * orientation); `true` reverses, `false` forces forward.
       */
      reverse?: boolean | null;
      /**
       * Override `SheetChart.axes.x.tickLblSkip`. `undefined` (or
       * omitted) inherits the source axis's skip; `null` drops the
       * inherited value (Excel falls back to showing every label); a
       * positive integer replaces it. Only meaningful for resolved
       * chart types whose X axis is `<c:catAx>` (bar / column / line
       * / area); silently dropped on scatter and pie / doughnut.
       */
      tickLblSkip?: number | null;
      /**
       * Override `SheetChart.axes.x.tickMarkSkip`. Same grammar and
       * scope rules as {@link tickLblSkip}.
       */
      tickMarkSkip?: number | null;
      /**
       * Override `SheetChart.axes.x.lblOffset`. `undefined` (or
       * omitted) inherits the source axis's label offset; `null`
       * drops the inherited value (the writer falls back to Excel's
       * default `100`); a number in the `0..1000` band replaces it.
       * Only meaningful for resolved chart types whose X axis is
       * `<c:catAx>` (bar / column / line / area); silently dropped
       * on scatter and pie / doughnut.
       */
      lblOffset?: number | null;
      /**
       * Override `SheetChart.axes.x.lblAlgn`. `undefined` (or
       * omitted) inherits the source axis's label alignment; `null`
       * drops the inherited value (the writer falls back to Excel's
       * default `"ctr"`); a {@link ChartAxisLabelAlign} token replaces
       * it. Unknown tokens collapse to `undefined` rather than
       * fabricate a value the writer would never emit. Only
       * meaningful for resolved chart types whose X axis is
       * `<c:catAx>` (bar / column / line / area); silently dropped
       * on scatter and pie / doughnut.
       */
      lblAlgn?: ChartAxisLabelAlign | null;
      /**
       * Override `SheetChart.axes.x.noMultiLvlLbl`. `undefined` (or
       * omitted) inherits the source axis's flag; `null` drops the
       * inherited value (the writer falls back to the OOXML `false`
       * default — multi-level labels enabled); a `boolean` replaces
       * it. Only meaningful for resolved chart types whose X axis is
       * `<c:catAx>` (bar / column / line / area); silently dropped on
       * scatter and pie / doughnut.
       */
      noMultiLvlLbl?: boolean | null;
      /**
       * Override `SheetChart.axes.x.hidden`. `undefined` (or omitted)
       * inherits the source axis's flag; `null` drops the inherited
       * value (the writer falls back to the OOXML `false` default —
       * axis visible); a `boolean` replaces it. Useful when porting a
       * "hide axis" template to a chart that should reveal its axis,
       * or vice versa.
       *
       * Silently dropped when the resolved chart type is `pie` /
       * `doughnut` since neither has axes.
       */
      hidden?: boolean | null;
    };
    y?: {
      title?: string | null;
      gridlines?: ChartAxisGridlines | null;
      scale?: ChartAxisScale | null;
      numberFormat?: ChartAxisNumberFormat | null;
      /** See {@link CloneChartOptions.axes.x.majorTickMark}. */
      majorTickMark?: ChartAxisTickMark | null;
      /** See {@link CloneChartOptions.axes.x.minorTickMark}. */
      minorTickMark?: ChartAxisTickMark | null;
      /** See {@link CloneChartOptions.axes.x.tickLblPos}. */
      tickLblPos?: ChartAxisTickLabelPosition | null;
      /** See {@link CloneChartOptions.axes.x.hidden}. */
      hidden?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.reverse}. */
      reverse?: boolean | null;
    };
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

  // `<c:smooth>`, `<a:ln>` (stroke), and `<c:marker>` all render
  // meaningfully only on line / scatter series; drop them from every
  // other resolved type so a doughnut → column flatten (or any other
  // coercion) does not leak the fields into a chart kind whose schema
  // rejects them.
  if (type !== "line" && type !== "scatter") {
    for (const s of series) {
      if (s.smooth !== undefined) delete s.smooth;
      if (s.stroke !== undefined) delete s.stroke;
      if (s.marker !== undefined) delete s.marker;
    }
  }

  // `<c:invertIfNegative>` lives exclusively on bar / column series
  // (CT_BarSer / CT_Bar3DSer); drop the field from every other
  // resolved type so a column → line flatten (or any other coercion)
  // does not leak the flag into a chart kind whose schema rejects it.
  if (type !== "bar" && type !== "column") {
    for (const s of series) {
      if (s.invertIfNegative !== undefined) delete s.invertIfNegative;
    }
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

  // Legend / per-family grouping carry over from the source when the
  // caller does not supply an override. Each grouping only round-trips
  // for the matching target family — applying a stacked grouping to a
  // family that does not support it would be silently ignored by the
  // writer, so we drop the inherited value to keep the model honest.
  const legend = options.legend !== undefined ? options.legend : source.legend;
  if (legend !== undefined) out.legend = legend;

  // `legendOverlay` only renders inside `<c:legend>`, so a clone whose
  // resolved legend is `false` (legend hidden) drops the inherited
  // overlay flag — there is no `<c:overlay>` slot on a hidden legend
  // for the writer to populate. The override wins over the source's
  // parsed value; absence inherits, `null` drops, a `boolean` replaces.
  if (legend !== false) {
    const resolvedLegendOverlay = resolveLegendOverlay(source.legendOverlay, options.legendOverlay);
    if (resolvedLegendOverlay !== undefined) out.legendOverlay = resolvedLegendOverlay;
  }

  const barGrouping = options.barGrouping !== undefined ? options.barGrouping : source.barGrouping;
  if (barGrouping !== undefined && (type === "bar" || type === "column")) {
    out.barGrouping = barGrouping;
  }

  // Bar / column gap width and overlap only make sense on bar-family
  // targets — flattening a column template into a line clone drops
  // the inherited values so they do not leak into a chart kind that
  // has no `<c:barChart>` element to host them. The override wins over
  // the source's parsed value.
  if (type === "bar" || type === "column") {
    const gapWidth = options.gapWidth !== undefined ? options.gapWidth : source.gapWidth;
    if (gapWidth !== undefined) out.gapWidth = gapWidth;
    const overlap = options.overlap !== undefined ? options.overlap : source.overlap;
    if (overlap !== undefined) out.overlap = overlap;
  }

  const lineGrouping =
    options.lineGrouping !== undefined ? options.lineGrouping : source.lineGrouping;
  if (lineGrouping !== undefined && type === "line") {
    out.lineGrouping = lineGrouping;
  }

  const areaGrouping =
    options.areaGrouping !== undefined ? options.areaGrouping : source.areaGrouping;
  if (areaGrouping !== undefined && type === "area") {
    out.areaGrouping = areaGrouping;
  }

  // Doughnut hole size only makes sense when the resolved type is
  // doughnut; flattening to pie (or any other kind) drops the hint so
  // the writer does not silently ignore it. The override wins over the
  // source's parsed `holeSize`.
  if (type === "doughnut") {
    const holeSize = options.holeSize !== undefined ? options.holeSize : source.holeSize;
    if (holeSize !== undefined) out.holeSize = holeSize;
  }

  // First slice angle round-trips for both pie and doughnut — the
  // OOXML schema places the element on `<c:pieChart>` and
  // `<c:doughnutChart>` alike. A doughnut template flattened to pie
  // therefore keeps its rotation; coercion into a non-pie family drops
  // the inherited value so it never leaks into a chart kind that has
  // no rotation knob.
  if (type === "pie" || type === "doughnut") {
    const firstSliceAng =
      options.firstSliceAng !== undefined ? options.firstSliceAng : source.firstSliceAng;
    if (firstSliceAng !== undefined) out.firstSliceAng = firstSliceAng;
  }

  if (options.showTitle !== undefined) out.showTitle = options.showTitle;
  if (options.altText !== undefined) out.altText = options.altText;
  if (options.frameTitle !== undefined) out.frameTitle = options.frameTitle;

  const resolvedDataLabels = resolveChartDataLabels(source.dataLabels, options.dataLabels);
  if (resolvedDataLabels !== undefined) out.dataLabels = resolvedDataLabels;

  const resolvedDispBlanks = resolveDispBlanksAs(source.dispBlanksAs, options.dispBlanksAs);
  if (resolvedDispBlanks !== undefined) out.dispBlanksAs = resolvedDispBlanks;

  const resolvedVaryColors = resolveVaryColors(source.varyColors, options.varyColors);
  if (resolvedVaryColors !== undefined) out.varyColors = resolvedVaryColors;

  const resolvedPlotVisOnly = resolvePlotVisOnly(source.plotVisOnly, options.plotVisOnly);
  if (resolvedPlotVisOnly !== undefined) out.plotVisOnly = resolvedPlotVisOnly;

  const resolvedRoundedCorners = resolveRoundedCorners(
    source.roundedCorners,
    options.roundedCorners,
  );
  if (resolvedRoundedCorners !== undefined) out.roundedCorners = resolvedRoundedCorners;

  // `<c:scatterStyle>` only renders inside `<c:scatterChart>`. Drop the
  // field on every other resolved type so a scatter template flattened
  // to line / column does not leak the preset into a chart kind whose
  // schema rejects it. Override wins over the source's parsed value.
  if (type === "scatter") {
    const resolvedScatterStyle = resolveScatterStyle(source.scatterStyle, options.scatterStyle);
    if (resolvedScatterStyle !== undefined) out.scatterStyle = resolvedScatterStyle;
  }

  // Pie and doughnut have no axes, so silently skip carrying over axis
  // titles even when the source declared them or the caller passed an
  // override.
  if (type !== "pie" && type !== "doughnut") {
    const axes = resolveAxes(source.axes, options.axes, type);
    if (axes !== undefined) out.axes = axes;
  }

  return out;
}

// ── Internals ────────────────────────────────────────────────────────

/**
 * Map a read-side {@link ChartKind} to the writer's
 * {@link WriteChartKind}, or `undefined` when no equivalent exists.
 *
 * 3D variants collapse onto their 2D counterparts; `doughnut` keeps
 * its own write-side kind so a doughnut template round-trips with the
 * hole intact. Kinds with no analog (`bubble`, `radar`, `surface`,
 * `stock`, `ofPie`) return `undefined` and force the caller to pass
 * an explicit `type` override.
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
      return "pie";
    case "doughnut":
      return "doughnut";
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

  const smooth = resolveSmooth(src?.smooth, ov?.smooth);
  if (smooth !== undefined) out.smooth = smooth;

  const stroke = resolveStroke(src?.stroke, ov?.stroke);
  if (stroke !== undefined) out.stroke = stroke;

  const marker = resolveMarker(src?.marker, ov?.marker);
  if (marker !== undefined) out.marker = marker;

  const invertIfNegative = resolveInvertIfNegative(src?.invertIfNegative, ov?.invertIfNegative);
  if (invertIfNegative !== undefined) out.invertIfNegative = invertIfNegative;

  return out;
}

/**
 * Resolve a per-series line-stroke override.
 *
 * `undefined` → inherit the source series' `stroke` (a fresh shallow
 *               copy so the caller cannot mutate the parsed source).
 * `null`      → drop the inherited block.
 * object      → replace the inherited block wholesale (no per-field
 *               merge; pass the full shape you want).
 *
 * An empty stroke block (no dash, no width) collapses to `undefined`
 * so the writer can elide the element rather than emit a bare
 * `<a:ln/>` that Excel paints with the inherited default.
 */
function resolveStroke(
  sourceStroke: ChartLineStroke | undefined,
  override: ChartLineStroke | null | undefined,
): ChartLineStroke | undefined {
  if (override === undefined) {
    if (!sourceStroke) return undefined;
    return cloneStroke(sourceStroke);
  }
  if (override === null) return undefined;
  return cloneStroke(override);
}

function cloneStroke(source: ChartLineStroke): ChartLineStroke | undefined {
  const out: ChartLineStroke = {};
  if (source.dash !== undefined) out.dash = source.dash;
  if (typeof source.width === "number" && Number.isFinite(source.width)) out.width = source.width;
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Resolve a per-series smooth-line override.
 *
 * `undefined` → inherit the source series' `smooth`.
 * `null`      → drop the inherited flag (the cloned series renders straight).
 * `boolean`   → replace.
 *
 * Only the `true` outcome materializes on the result — `false` collapses
 * to `undefined` so absence and the OOXML default round-trip identically
 * (the writer emits straight segments either way).
 */
function resolveSmooth(
  sourceSmooth: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return sourceSmooth === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve a per-series invert-if-negative override.
 *
 * `undefined` → inherit the source series' `invertIfNegative`.
 * `null`      → drop the inherited flag (the cloned series renders
 *               negatives in the series fill color).
 * `boolean`   → replace.
 *
 * Only the `true` outcome materializes on the result — `false` collapses
 * to `undefined` so absence and the OOXML default round-trip identically
 * (the writer omits `<c:invertIfNegative>` either way).
 */
function resolveInvertIfNegative(
  sourceFlag: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return sourceFlag === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve a per-series marker override.
 *
 * `undefined` → inherit the source series' `marker` (a fresh shallow
 * copy so the caller cannot mutate the parsed source).
 * `null`      → drop the inherited block (the cloned series falls back
 *               to Excel's series-rotation default).
 * object      → replace the inherited block wholesale.
 *
 * An empty marker block (no symbol, size, or color) collapses to
 * `undefined` so the writer can elide the element rather than emit a
 * bare `<c:marker/>` that Excel paints with the inherited default.
 */
function resolveMarker(
  sourceMarker: ChartMarker | undefined,
  override: ChartMarker | null | undefined,
): ChartMarker | undefined {
  if (override === undefined) {
    if (!sourceMarker) return undefined;
    return cloneMarker(sourceMarker);
  }
  if (override === null) return undefined;
  return cloneMarker(override);
}

function cloneMarker(source: ChartMarker): ChartMarker | undefined {
  const out: ChartMarker = {};
  if (source.symbol !== undefined) out.symbol = source.symbol;
  if (typeof source.size === "number" && Number.isFinite(source.size)) out.size = source.size;
  if (typeof source.fill === "string" && source.fill.length > 0) out.fill = source.fill;
  if (typeof source.line === "string" && source.line.length > 0) out.line = source.line;
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Resolve a `dispBlanksAs` override.
 *
 * `undefined` → inherit the source's parsed `dispBlanksAs`.
 * `null`      → drop the inherited value (the writer falls back to
 *               the OOXML `"gap"` default).
 * value       → replace.
 *
 * Unknown strings are ignored (treated as `undefined`); only the three
 * OOXML-defined tokens propagate through to the writer.
 */
function resolveDispBlanksAs(
  sourceValue: ChartDisplayBlanksAs | undefined,
  override: ChartDisplayBlanksAs | null | undefined,
): ChartDisplayBlanksAs | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `varyColors` override.
 *
 * `undefined` → inherit the source's parsed `varyColors`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               per-family default — `true` for pie / doughnut, `false`
 *               everywhere else).
 * `boolean`   → replace.
 *
 * The override grammar mirrors `dispBlanksAs` so the two chart-level
 * toggles compose the same way at the call site.
 */
function resolveVaryColors(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `plotVisOnly` override.
 *
 * `undefined` → inherit the source's parsed `plotVisOnly`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `true` default — hidden cells drop out of the chart).
 * `boolean`   → replace.
 *
 * The grammar mirrors `dispBlanksAs` / `varyColors` so the chart-level
 * toggles compose the same way at the call site.
 */
function resolvePlotVisOnly(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `roundedCorners` override.
 *
 * `undefined` → inherit the source's parsed `roundedCorners`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `false` default — square chart frame).
 * `boolean`   → replace.
 *
 * The grammar mirrors `plotVisOnly` / `varyColors` so the chart-frame
 * toggles compose the same way at the call site.
 */
function resolveRoundedCorners(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `legendOverlay` override.
 *
 * `undefined` → inherit the source's parsed `legendOverlay`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `false` default — the legend reserves its own
 *               slot, no overlap with the plot area).
 * `boolean`   → replace.
 *
 * The grammar mirrors `plotVisOnly` / `roundedCorners` so the chart-
 * level toggles compose the same way at the call site. Callers should
 * gate the result on the resolved legend visibility — when no legend
 * is emitted, the overlay flag has no slot in the rendered chart.
 */
function resolveLegendOverlay(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `scatterStyle` override.
 *
 * `undefined` → inherit the source's parsed `scatterStyle`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               default `"lineMarker"`).
 * value       → replace.
 *
 * The grammar mirrors `dispBlanksAs` / `varyColors` so the chart-level
 * toggles compose the same way at the call site.
 */
function resolveScatterStyle(
  sourceValue: ChartScatterStyle | undefined,
  override: ChartScatterStyle | null | undefined,
): ChartScatterStyle | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
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
 * elements or redundant gridline blocks.
 */
function resolveAxes(
  sourceAxes: Chart["axes"],
  overrides: CloneChartOptions["axes"],
  type: WriteChartKind,
): SheetChart["axes"] | undefined {
  const xTitle = applyOverride(sourceAxes?.x?.title, overrides?.x?.title);
  const yTitle = applyOverride(sourceAxes?.y?.title, overrides?.y?.title);
  const xGridlines = applyGridlinesOverride(sourceAxes?.x?.gridlines, overrides?.x?.gridlines);
  const yGridlines = applyGridlinesOverride(sourceAxes?.y?.gridlines, overrides?.y?.gridlines);
  const xScale = applyScaleOverride(sourceAxes?.x?.scale, overrides?.x?.scale);
  const yScale = applyScaleOverride(sourceAxes?.y?.scale, overrides?.y?.scale);
  const xNumFmt = applyNumberFormatOverride(
    sourceAxes?.x?.numberFormat,
    overrides?.x?.numberFormat,
  );
  const yNumFmt = applyNumberFormatOverride(
    sourceAxes?.y?.numberFormat,
    overrides?.y?.numberFormat,
  );
  const xMajorTickMark = applyTickMarkOverride(
    sourceAxes?.x?.majorTickMark,
    overrides?.x?.majorTickMark,
  );
  const yMajorTickMark = applyTickMarkOverride(
    sourceAxes?.y?.majorTickMark,
    overrides?.y?.majorTickMark,
  );
  const xMinorTickMark = applyTickMarkOverride(
    sourceAxes?.x?.minorTickMark,
    overrides?.x?.minorTickMark,
  );
  const yMinorTickMark = applyTickMarkOverride(
    sourceAxes?.y?.minorTickMark,
    overrides?.y?.minorTickMark,
  );
  const xTickLblPos = applyTickLblPosOverride(sourceAxes?.x?.tickLblPos, overrides?.x?.tickLblPos);
  const yTickLblPos = applyTickLblPosOverride(sourceAxes?.y?.tickLblPos, overrides?.y?.tickLblPos);
  const xReverse = applyReverseOverride(sourceAxes?.x?.reverse, overrides?.x?.reverse);
  const yReverse = applyReverseOverride(sourceAxes?.y?.reverse, overrides?.y?.reverse);
  // `tickLblSkip` / `tickMarkSkip` only render on category axes
  // (`<c:catAx>`). Scatter charts use two value axes, so the X axis
  // skip would be silently dropped by the writer anyway — collapse it
  // to undefined here so the cloned `SheetChart` accurately reflects
  // what the chart will paint.
  const isCatAxisX = type !== "scatter";
  const xTickLblSkip = isCatAxisX
    ? applySkipOverride(sourceAxes?.x?.tickLblSkip, overrides?.x?.tickLblSkip)
    : undefined;
  const xTickMarkSkip = isCatAxisX
    ? applySkipOverride(sourceAxes?.x?.tickMarkSkip, overrides?.x?.tickMarkSkip)
    : undefined;
  // `lblOffset` is also category-axis-only (CT_CatAx / CT_DateAx) per
  // the OOXML schema. Same scope rule as the skip elements above.
  const xLblOffset = isCatAxisX
    ? applyLblOffsetOverride(sourceAxes?.x?.lblOffset, overrides?.x?.lblOffset)
    : undefined;
  // `lblAlgn` is category-axis-only as well (CT_CatAx / CT_DateAx
  // per ECMA-376 §21.2.2). Same scope as `lblOffset`.
  const xLblAlgn = isCatAxisX
    ? applyLblAlgnOverride(sourceAxes?.x?.lblAlgn, overrides?.x?.lblAlgn)
    : undefined;
  // `noMultiLvlLbl` is even tighter — `CT_CatAx` only (no `<c:dateAx>`
  // slot per ECMA-376 §21.2.2). Reuse the catAx scope rule above; the
  // resolved chart type still funnels through `<c:catAx>` for every
  // bar / column / line / area family the writer supports.
  const xNoMultiLvlLbl = isCatAxisX
    ? applyNoMultiLvlLblOverride(sourceAxes?.x?.noMultiLvlLbl, overrides?.x?.noMultiLvlLbl)
    : undefined;
  // `<c:delete>` lives on every axis flavour — both `<c:catAx>` and
  // `<c:valAx>` accept it — so the hidden flag carries through every
  // chart family that has axes. Pie / doughnut have no axes at all
  // and the caller already short-circuited those above.
  const xHidden = applyHiddenOverride(sourceAxes?.x?.hidden, overrides?.x?.hidden);
  const yHidden = applyHiddenOverride(sourceAxes?.y?.hidden, overrides?.y?.hidden);

  const out: NonNullable<SheetChart["axes"]> = {};
  if (
    xTitle !== undefined ||
    xGridlines !== undefined ||
    xScale !== undefined ||
    xNumFmt !== undefined ||
    xMajorTickMark !== undefined ||
    xMinorTickMark !== undefined ||
    xTickLblPos !== undefined ||
    xReverse !== undefined ||
    xTickLblSkip !== undefined ||
    xTickMarkSkip !== undefined ||
    xLblOffset !== undefined ||
    xLblAlgn !== undefined ||
    xNoMultiLvlLbl !== undefined ||
    xHidden !== undefined
  ) {
    out.x = {};
    if (xTitle !== undefined) out.x.title = xTitle;
    if (xGridlines !== undefined) out.x.gridlines = xGridlines;
    if (xScale !== undefined) out.x.scale = xScale;
    if (xNumFmt !== undefined) out.x.numberFormat = xNumFmt;
    if (xMajorTickMark !== undefined) out.x.majorTickMark = xMajorTickMark;
    if (xMinorTickMark !== undefined) out.x.minorTickMark = xMinorTickMark;
    if (xTickLblPos !== undefined) out.x.tickLblPos = xTickLblPos;
    if (xReverse !== undefined) out.x.reverse = xReverse;
    if (xTickLblSkip !== undefined) out.x.tickLblSkip = xTickLblSkip;
    if (xTickMarkSkip !== undefined) out.x.tickMarkSkip = xTickMarkSkip;
    if (xLblOffset !== undefined) out.x.lblOffset = xLblOffset;
    if (xLblAlgn !== undefined) out.x.lblAlgn = xLblAlgn;
    if (xNoMultiLvlLbl !== undefined) out.x.noMultiLvlLbl = xNoMultiLvlLbl;
    if (xHidden !== undefined) out.x.hidden = xHidden;
  }
  if (
    yTitle !== undefined ||
    yGridlines !== undefined ||
    yScale !== undefined ||
    yNumFmt !== undefined ||
    yMajorTickMark !== undefined ||
    yMinorTickMark !== undefined ||
    yTickLblPos !== undefined ||
    yHidden !== undefined ||
    yReverse !== undefined
  ) {
    out.y = {};
    if (yTitle !== undefined) out.y.title = yTitle;
    if (yGridlines !== undefined) out.y.gridlines = yGridlines;
    if (yScale !== undefined) out.y.scale = yScale;
    if (yNumFmt !== undefined) out.y.numberFormat = yNumFmt;
    if (yMajorTickMark !== undefined) out.y.majorTickMark = yMajorTickMark;
    if (yMinorTickMark !== undefined) out.y.minorTickMark = yMinorTickMark;
    if (yTickLblPos !== undefined) out.y.tickLblPos = yTickLblPos;
    if (yHidden !== undefined) out.y.hidden = yHidden;
    if (yReverse !== undefined) out.y.reverse = yReverse;
  }

  return out.x || out.y ? out : undefined;
}

/**
 * Resolve a `tickLblSkip` / `tickMarkSkip` override using the same
 * `undefined` (inherit) / `null` (drop) / value (replace) grammar as
 * the other axis helpers. Out-of-range / non-positive values collapse
 * to `undefined` so they cannot leak into the writer (which would
 * silently drop them anyway via {@link normalizeAxisSkip}).
 */
function applySkipOverride(
  source: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) {
    if (typeof source !== "number" || !Number.isFinite(source)) return undefined;
    const rounded = Math.round(source);
    if (rounded < 1 || rounded > 32767 || rounded === 1) return undefined;
    return rounded;
  }
  if (override === null) return undefined;
  if (typeof override !== "number" || !Number.isFinite(override)) return undefined;
  const rounded = Math.round(override);
  if (rounded < 1 || rounded > 32767 || rounded === 1) return undefined;
  return rounded;
}

/**
 * Resolve an `lblOffset` override using the same `undefined` (inherit)
 * / `null` (drop) / value (replace) grammar as the other axis helpers.
 * Out-of-range / non-numeric values collapse to `undefined` so they
 * cannot leak into the writer (which would silently drop them anyway
 * via {@link normalizeAxisLblOffset}). The OOXML default `100` also
 * collapses to `undefined` so absence and the default round-trip
 * identically — symmetric with the parser-side default-collapse.
 */
function applyLblOffsetOverride(
  source: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) {
    if (typeof source !== "number" || !Number.isFinite(source)) return undefined;
    const rounded = Math.round(source);
    if (rounded < 0 || rounded > 1000 || rounded === 100) return undefined;
    return rounded;
  }
  if (override === null) return undefined;
  if (typeof override !== "number" || !Number.isFinite(override)) return undefined;
  const rounded = Math.round(override);
  if (rounded < 0 || rounded > 1000 || rounded === 100) return undefined;
  return rounded;
}

/**
 * Resolve an `lblAlgn` override using the same `undefined` (inherit)
 * / `null` (drop) / value (replace) grammar as the other axis helpers.
 * Unknown tokens collapse to `undefined` so they cannot leak into the
 * writer (which would silently drop them anyway via
 * {@link normalizeAxisLblAlgn}). The OOXML default `"ctr"` also
 * collapses to `undefined` so absence and the default round-trip
 * identically — symmetric with the parser-side default-collapse.
 */
function applyLblAlgnOverride(
  source: ChartAxisLabelAlign | undefined,
  override: ChartAxisLabelAlign | null | undefined,
): ChartAxisLabelAlign | undefined {
  if (override === undefined) {
    if (source !== "l" && source !== "r" && source !== "ctr") return undefined;
    return source === "ctr" ? undefined : source;
  }
  if (override === null) return undefined;
  if (override !== "l" && override !== "r" && override !== "ctr") return undefined;
  return override === "ctr" ? undefined : override;
}

/**
 * Resolve a `noMultiLvlLbl` override using the same `undefined`
 * (inherit) / `null` (drop) / `boolean` (replace) grammar as the
 * other axis helpers. Only `true` surfaces (the writer treats `false`
 * and absence identically — both produce `<c:noMultiLvlLbl val="0"/>`),
 * so an override of `false` collapses to `undefined` to keep the
 * cloned `SheetChart` shape minimal. Non-boolean inputs fall through
 * the type guard to `undefined`.
 */
function applyNoMultiLvlLblOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return source === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve an axis `hidden` override using the same `undefined`
 * (inherit) / `null` (drop) / `boolean` (replace) grammar as the
 * other axis helpers. Only `true` surfaces (the writer treats `false`
 * and absence identically — both produce `<c:delete val="0"/>`), so
 * an override of `false` collapses to `undefined` to keep the cloned
 * `SheetChart` shape minimal. Non-boolean inputs fall through the
 * type guard to `undefined`.
 */
function applyHiddenOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return source === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve gridlines using the same `undefined` (inherit) / `null`
 * (drop) / object (replace) grammar as the other axis overrides.
 * Returns `undefined` when neither source nor override declares a
 * non-empty gridline configuration.
 */
function applyGridlinesOverride(
  source: ChartAxisGridlines | undefined,
  override: ChartAxisGridlines | null | undefined,
): ChartAxisGridlines | undefined {
  if (override === undefined) {
    if (!source) return undefined;
    const out: ChartAxisGridlines = {};
    if (source.major) out.major = true;
    if (source.minor) out.minor = true;
    return out.major || out.minor ? out : undefined;
  }
  if (override === null) return undefined;
  const out: ChartAxisGridlines = {};
  if (override.major === true) out.major = true;
  if (override.minor === true) out.minor = true;
  return out.major || out.minor ? out : undefined;
}

/**
 * Resolve a scale override using the same `undefined` / `null` /
 * object grammar as {@link applyGridlinesOverride}. The override
 * replaces the source wholesale rather than merging field-by-field —
 * a partial template scale `{ min: 0 }` plus an override
 * `{ max: 100 }` yields `{ max: 100 }`, not `{ min: 0, max: 100 }`.
 * Per-field merges proved confusing in the dashboard composition flow
 * (callers expected the override to fully describe the target scale),
 * so wholesale replacement is the simpler contract.
 */
function applyScaleOverride(
  source: ChartAxisScale | undefined,
  override: ChartAxisScale | null | undefined,
): ChartAxisScale | undefined {
  if (override === undefined) {
    if (!source) return undefined;
    return cloneScale(source);
  }
  if (override === null) return undefined;
  return cloneScale(override);
}

function cloneScale(source: ChartAxisScale): ChartAxisScale | undefined {
  const out: ChartAxisScale = {};
  if (typeof source.min === "number" && Number.isFinite(source.min)) out.min = source.min;
  if (typeof source.max === "number" && Number.isFinite(source.max)) out.max = source.max;
  if (
    typeof source.majorUnit === "number" &&
    Number.isFinite(source.majorUnit) &&
    source.majorUnit > 0
  ) {
    out.majorUnit = source.majorUnit;
  }
  if (
    typeof source.minorUnit === "number" &&
    Number.isFinite(source.minorUnit) &&
    source.minorUnit > 0
  ) {
    out.minorUnit = source.minorUnit;
  }
  if (typeof source.logBase === "number" && Number.isFinite(source.logBase)) {
    out.logBase = source.logBase;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Resolve a number-format override. Same grammar as the other
 * per-axis helpers: `undefined` inherits, `null` drops, an object
 * replaces.
 */
function applyNumberFormatOverride(
  source: ChartAxisNumberFormat | undefined,
  override: ChartAxisNumberFormat | null | undefined,
): ChartAxisNumberFormat | undefined {
  if (override === undefined) {
    if (!source) return undefined;
    if (typeof source.formatCode !== "string" || source.formatCode.length === 0) return undefined;
    const out: ChartAxisNumberFormat = { formatCode: source.formatCode };
    if (source.sourceLinked === true) out.sourceLinked = true;
    return out;
  }
  if (override === null) return undefined;
  if (typeof override.formatCode !== "string" || override.formatCode.length === 0) return undefined;
  const out: ChartAxisNumberFormat = { formatCode: override.formatCode };
  if (override.sourceLinked === true) out.sourceLinked = true;
  return out;
}

/** Recognized values of `<c:majorTickMark>` / `<c:minorTickMark>`. */
const VALID_TICK_MARK_VALUES: ReadonlySet<ChartAxisTickMark> = new Set([
  "none",
  "in",
  "out",
  "cross",
]);

/**
 * Resolve a tick-mark override using the same `undefined` (inherit) /
 * `null` (drop) / value (replace) grammar as the other axis helpers.
 * Unknown / typo'd inputs collapse to `undefined` so the writer never
 * emits a value the OOXML `ST_TickMark` enum rejects.
 */
function applyTickMarkOverride(
  source: ChartAxisTickMark | undefined,
  override: ChartAxisTickMark | null | undefined,
): ChartAxisTickMark | undefined {
  if (override === undefined) {
    if (source === undefined) return undefined;
    return VALID_TICK_MARK_VALUES.has(source) ? source : undefined;
  }
  if (override === null) return undefined;
  return VALID_TICK_MARK_VALUES.has(override) ? override : undefined;
}

/** Recognized values of `<c:tickLblPos>`. */
const VALID_TICK_LBL_POS_VALUES: ReadonlySet<ChartAxisTickLabelPosition> = new Set([
  "nextTo",
  "low",
  "high",
  "none",
]);

/**
 * Resolve a tick-label-position override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Unknown / typo'd inputs collapse to `undefined` so
 * the writer never emits a value the OOXML `ST_TickLblPos` enum
 * rejects.
 */
function applyTickLblPosOverride(
  source: ChartAxisTickLabelPosition | undefined,
  override: ChartAxisTickLabelPosition | null | undefined,
): ChartAxisTickLabelPosition | undefined {
  if (override === undefined) {
    if (source === undefined) return undefined;
    return VALID_TICK_LBL_POS_VALUES.has(source) ? source : undefined;
  }
  if (override === null) return undefined;
  return VALID_TICK_LBL_POS_VALUES.has(override) ? override : undefined;
}

/**
 * Resolve a reverse-axis override using the same `undefined` (inherit) /
 * `null` (drop) / value (replace) grammar as the other axis helpers.
 *
 * Only `true` round-trips meaningfully — `false` is the OOXML default
 * (`orientation="minMax"`) so it collapses to `undefined` to keep the
 * cloned shape minimal. A source carrying `false` (e.g. an over-eager
 * parser that surfaced the default) collapses to `undefined` on
 * inherit; an explicit `false` override likewise drops the field. The
 * writer's per-axis `reverse: false` default already produces a forward
 * orientation, so the dropped state is indistinguishable from a literal
 * `false`.
 */
function applyReverseOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return source === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

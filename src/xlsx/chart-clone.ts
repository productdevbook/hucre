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
  ChartAxisCrossBetween,
  ChartAxisCrosses,
  ChartAxisDispUnit,
  ChartAxisDispUnits,
  ChartAxisGridlines,
  ChartAxisLabelAlign,
  ChartAxisNumberFormat,
  ChartAxisScale,
  ChartAxisTickLabelPosition,
  ChartAxisTickMark,
  ChartBorderDash,
  ChartDataLabels,
  ChartDataLabelsInfo,
  ChartDataPoint,
  ChartDataTable,
  ChartDisplayBlanksAs,
  ChartErrorBars,
  ChartKind,
  ChartLegendEntry,
  ChartLineStroke,
  ChartManualLayout,
  ChartMarker,
  ChartProtection,
  ChartScatterStyle,
  ChartSeries,
  ChartSeriesInfo,
  ChartShape3D,
  ChartTrendline,
  ChartView3D,
  SheetChart,
  WriteChartKind,
} from "../_types";
import {
  clampStrokeWidthPt as normalizeBorderWidthPt,
  normalizeBorderDash,
  normalizeRgbHex,
  resolveBorderDash,
  resolveBorderWidthPt,
} from "./chart/shape";
import {
  resolveBackWallThickness,
  resolveFloorThickness,
  resolveSideWallThickness,
  resolveView3D,
} from "./chart/walls";
import { normalizeLayoutCoordinate } from "./chart/layout";
import {
  normalizeTitleBold,
  normalizeTitleColor,
  normalizeTitleFontFamily,
  normalizeTitleFontSize,
  normalizeTitleItalic,
  normalizeTitleRotation,
  normalizeTitleStrike,
  normalizeTitleUnderline,
  resolveCloneTitleBold,
  resolveCloneTitleBorderColor,
  resolveCloneTitleBorderWidth,
  resolveCloneTitleColor,
  resolveCloneTitleFillColor,
  resolveCloneTitleFontFamily,
  resolveCloneTitleFontSize,
  resolveCloneTitleItalic,
  resolveCloneTitleLayout,
  resolveCloneTitleOverlay,
  resolveCloneTitleRotation,
  resolveCloneTitleStrike,
  resolveCloneTitleUnderline,
} from "./chart/title";
import {
  normalizeLegendBold,
  normalizeLegendBorderWidth,
  normalizeLegendFontFamily,
  normalizeLegendItalic,
  normalizeLegendLayout,
  normalizeLegendStrikethrough,
  normalizeLegendUnderline,
  resolveCloneLegendBold,
  resolveCloneLegendBorderColor,
  resolveCloneLegendBorderWidth,
  resolveCloneLegendEntries,
  resolveCloneLegendFillColor,
  resolveCloneLegendFontColor,
  resolveCloneLegendFontFamily,
  resolveCloneLegendFontSize,
  resolveCloneLegendItalic,
  resolveCloneLegendLayout,
  resolveCloneLegendOverlay,
  resolveCloneLegendStrikethrough,
  resolveCloneLegendUnderline,
} from "./chart/legend";
import {
  buildSeriesFromSource,
  cloneMarker,
  cloneStroke,
  resolveExplosion,
  resolveInvertIfNegative,
  resolveMarker,
  resolveShowLineMarkers,
  resolveSmooth,
  resolveStroke,
} from "./chart/series";
import { applyOverride } from "./chart/util";
import { resolveAxes, resolveCloneAutoTitleDeleted } from "./chart/axis";
import {
  resolveCloneDropLines,
  resolveCloneHiLowLines,
  resolveClonePlotAreaBorderColor,
  resolveClonePlotAreaBorderWidth,
  resolveClonePlotAreaFillColor,
  resolveClonePlotAreaLayout,
  resolveCloneScatterStyle,
  resolveCloneSerLines,
  resolveCloneUpDownBars,
  resolveCloneUpDownBarsGapWidth,
  resolveCloneVaryColors,
} from "./chart/plotArea";
import { resolveCloneDataTable } from "./chart/dataTable";
import { resolveChartDataLabels, resolveSeriesDataLabels } from "./chart/dataLabels";
import {
  resolveCloneChartSpaceBorderColor,
  resolveCloneChartSpaceFillColor,
  resolveCloneDate1904,
  resolveCloneDispBlanksAs,
  resolveCloneLang,
  resolveClonePlotVisOnly,
  resolveCloneProtection,
  resolveCloneRoundedCorners,
  resolveCloneShowDLblsOverMax,
  resolveCloneStyle,
} from "./chart/chartSpace";

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
  /**
   * Slice-explosion override (in percent of the radius). `undefined`
   * (or omitted) inherits the source series' `explosion`; `null` drops
   * the inherited value (the cloned series falls back to the OOXML
   * default `0`); a finite `number` replaces it wholesale (clamped to
   * the 0..400% band Excel's UI exposes; `0` collapses to absence).
   * Only meaningful for `pie` and `doughnut` clones — silently dropped
   * from the output when the resolved chart type is anything else.
   */
  explosion?: number | null;
  /**
   * Per-data-point override array. `undefined` inherits the source
   * series' `dataPoints`; `null` drops them; an array replaces.
   */
  dataPoints?: ChartDataPoint[] | null;
  /**
   * Per-series trendline override array. `undefined` inherits the
   * source series' `trendlines`; `null` drops them; an array replaces.
   * Silently dropped on pie / doughnut clones.
   */
  trendlines?: ChartTrendline[] | null;
  /**
   * Per-series error-bar override array. `undefined` inherits the
   * source series' `errorBars`; `null` drops them; an array replaces.
   * Silently dropped on pie / doughnut clones.
   */
  errorBars?: ChartErrorBars[] | null;
  /**
   * Bubble-size A1 range override. `undefined` inherits the source
   * series' parsed `bubbleSizeRef`; `null` drops it; a string replaces.
   * Silently dropped on every family except `bubble`.
   */
  bubbleSize?: string | null;
  /**
   * 3D shape variant override. `undefined` inherits the source series'
   * `shape3D`; `null` drops it; a {@link ChartShape3D} replaces.
   * Silently dropped on every family except `bar3D`.
   */
  shape3D?: ChartShape3D | null;
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
  /**
   * Override the chart-level per-series legend-entry overrides.
   * `undefined` (or omitted) inherits the source's parsed list; `null`
   * drops every inherited entry (the writer emits no `<c:legendEntry>`
   * children); a `ChartLegendEntry[]` replaces the inherited list
   * outright.
   *
   * The override is silently dropped from the cloned `SheetChart` when
   * the resolved legend is `false` (no `<c:legend>` element will be
   * emitted) — there is no slot to host the entries on a hidden legend,
   * so leaking the value into the output would carry a list Excel never
   * reads.
   *
   * Replacement semantics matter when the cloned chart's series count
   * differs from the source's: an entry whose `idx` no longer points
   * at a real series still emits — Excel renders it harmlessly — but a
   * caller can pass `null` (or an empty array) to start fresh.
   */
  legendEntries?: ChartLegendEntry[] | null;
  /**
   * Override `SheetChart.legendFontSize`. `undefined` (or omitted)
   * inherits the source's parsed `legendFontSize`; `null` drops the
   * inherited size (the writer emits no `<c:txPr>` block on the
   * legend, falling back to Excel's theme-default 9pt); a number
   * replaces. Out-of-range / non-numeric / non-finite overrides
   * collapse to `undefined` (inherit) so a typed escape from an
   * untyped caller cannot pin a value the writer would silently elide
   * back to absence. Fractional inputs round to the nearest 0.5pt
   * (Excel's UI granularity).
   *
   * The override is silently dropped from the cloned `SheetChart` when
   * the resolved legend is `false` (no `<c:legend>` element will be
   * emitted) — there is no `<c:txPr>` slot to host the size on a
   * hidden legend, so leaking the value into the output would carry a
   * pin Excel never reads.
   *
   * The grammar mirrors `titleFontSize` / `axes.x.axisTitleFontSize` /
   * `axes.x.labelFontSize` so the typography knobs compose the same way
   * at the call site. Bridges another typography-customization gap for
   * the dashboard composition flow tracked in #136 — a templated
   * dashboard chart whose user pinned a custom legend size now
   * round-trips that pin through the parse / clone / write loop.
   */
  legendFontSize?: number | null;
  /**
   * Override `SheetChart.legendBold`. `undefined` (or omitted) inherits
   * the source's parsed `legendBold`; `null` drops the inherited flag
   * (the writer emits no `<c:txPr>` block on the legend, falling back
   * to the OOXML default — no `b` attribute, equivalent to non-bold);
   * a `boolean` replaces. Non-boolean overrides (typed escapes from an
   * untyped caller) collapse to `undefined` so a typed escape cannot
   * pin a value the writer would silently elide.
   *
   * The override is silently dropped from the cloned `SheetChart` when
   * the resolved legend is `false` (no `<c:legend>` element will be
   * emitted) — there is no `<c:txPr>` slot to host the flag on a
   * hidden legend, so leaking the value into the output would carry a
   * pin Excel never reads.
   *
   * The grammar mirrors `titleBold` / `axes.x.axisTitleBold` /
   * `axes.x.labelBold` so the typography knobs compose the same way
   * at the call site. Bridges another typography-customization gap
   * for the dashboard composition flow tracked in #136 — a templated
   * dashboard chart whose user pinned a custom legend bold flag now
   * round-trips that pin through the parse / clone / write loop.
   */
  legendBold?: boolean | null;
  /**
   * Override `SheetChart.legendItalic`. `undefined` (or omitted)
   * inherits the source's parsed `legendItalic`; `null` drops the
   * inherited flag (the writer emits no `<c:txPr>` block on the
   * legend, falling back to the OOXML default — no `i` attribute,
   * equivalent to non-italic); a `boolean` replaces. Non-boolean
   * overrides (typed escapes from an untyped caller) collapse to
   * `undefined` so a typed escape cannot pin a value the writer
   * would silently elide.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved legend is `false` (no `<c:legend>` element will
   * be emitted) — there is no `<c:txPr>` slot to host the flag on a
   * hidden legend, so leaking the value into the output would carry
   * a pin Excel never reads.
   *
   * The grammar mirrors `titleItalic` / `axes.x.axisTitleItalic` /
   * `axes.x.labelItalic` so the typography knobs compose the same way
   * at the call site. Bridges another typography-customization gap
   * for the dashboard composition flow tracked in #136 — a templated
   * dashboard chart whose user pinned a custom legend italic flag now
   * round-trips that pin through the parse / clone / write loop.
   */
  legendItalic?: boolean | null;
  /**
   * Override `SheetChart.legendUnderline`. `undefined` (or omitted)
   * inherits the source's parsed `legendUnderline`; `null` drops the
   * inherited flag (the writer emits no `u` attribute on the legend's
   * `<a:defRPr>`, falling back to the OOXML default — non-underlined);
   * a `boolean` replaces. Non-boolean overrides (typed escapes from an
   * untyped caller) collapse to `undefined` so a typed escape cannot
   * pin a value the writer would silently elide.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved legend is `false` (no `<c:legend>` element will
   * be emitted) — there is no `<c:txPr>` slot to host the flag on a
   * hidden legend, so leaking the value into the output would carry
   * a pin Excel never reads.
   *
   * The grammar mirrors `titleUnderline` /
   * `axes.x.axisTitleUnderline` / `axes.x.labelUnderline` so the
   * typography knobs compose the same way at the call site.
   */
  legendUnderline?: boolean | null;
  /**
   * Override `SheetChart.legendStrikethrough`. `undefined` (or omitted)
   * inherits the source's parsed `legendStrikethrough`; `null` drops
   * the inherited flag (the writer emits no `strike` attribute on the
   * legend's `<a:defRPr>`, falling back to the OOXML default — no
   * strikethrough); a `boolean` replaces. Non-boolean overrides (typed
   * escapes from an untyped caller) collapse to `undefined` so a typed
   * escape cannot pin a value the writer would silently elide.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved legend is `false` (no `<c:legend>` element will
   * be emitted) — there is no `<c:txPr>` slot to host the flag on a
   * hidden legend, so leaking the value into the output would carry
   * a pin Excel never reads.
   *
   * The grammar mirrors `titleStrikethrough` /
   * `axes.x.axisTitleStrike` / `axes.x.labelStrikethrough` so the
   * typography knobs compose the same way at the call site.
   */
  legendStrikethrough?: boolean | null;
  /**
   * Override `SheetChart.legendFontColor`. `undefined` (or omitted)
   * inherits the source's parsed `legendFontColor`; `null` drops the
   * inherited fill (the writer emits no `<a:solidFill>` block on the
   * legend's `<a:defRPr>`, falling back to the theme text color); a
   * hex string replaces. Accepts the color either with or without a
   * leading `#` and in any case (`"FF0000"`, `"#FF0000"`, `"ff0000"`
   * all collapse to the OOXML uppercase canonical form `"FF0000"`);
   * malformed inputs (wrong length, non-hex characters, alpha-channel
   * forms, non-string escapes from an untyped caller) collapse to
   * `undefined` so a typed escape cannot pin a value the writer would
   * silently elide.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved legend is `false` (no `<c:legend>` element will
   * be emitted) — there is no `<c:txPr>` slot to host the fill on a
   * hidden legend, so leaking the value into the output would carry
   * a pin Excel never reads.
   *
   * The grammar mirrors `titleColor` / `axes.x.axisTitleColor` /
   * `axes.x.labelColor` so the typography knobs compose the same way
   * at the call site.
   */
  legendFontColor?: string | null;
  /**
   * Override `SheetChart.legendFontFamily`. `undefined` (or omitted)
   * inherits the source's parsed `legendFontFamily`; `null` drops the
   * inherited typeface (the writer emits no `<a:latin>` element on
   * the legend's `<a:defRPr>`, falling back to the theme typeface);
   * a non-empty string replaces it. The override is trimmed; empty /
   * whitespace-only strings and non-string overrides (typed escapes
   * from an untyped caller) collapse to `undefined` so the cloned
   * `SheetChart` always carries a value the writer will accept.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved legend is `false` (no `<c:legend>` element will
   * be emitted) — there is no `<c:txPr>` slot to host the typeface on
   * a hidden legend, so leaking the value into the output would carry
   * a pin Excel never reads.
   *
   * The grammar mirrors `titleFontFamily` /
   * `axes.x.axisTitleFontFamily` / `axes.x.labelFontFamily` so the
   * typography knobs compose the same way at the call site.
   */
  legendFontFamily?: string | null;
  /**
   * Override `SheetChart.legendLayout`. `undefined` (or omitted)
   * inherits the source's parsed `legendLayout`; `null` drops the
   * inherited layout (the writer emits no `<c:layout>` block on the
   * legend, falling back to Excel's auto-layout position); a
   * {@link ChartManualLayout} replaces it.
   *
   * Each of {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
   * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} runs
   * through the writer-side normalizer — coordinates outside the
   * `0..1` band, `NaN`, `Infinity`, and non-numeric overrides all
   * collapse to omitting the matching `<c:x>` / `<c:y>` / `<c:w>` /
   * `<c:h>` slot so the cloned `SheetChart` always carries a value the
   * writer will accept. An override whose every coordinate dropped on
   * normalization collapses the entire layout to `undefined` so the
   * writer skips the `<c:layout>` block entirely.
   *
   * The override is silently dropped from the cloned `SheetChart` when
   * the resolved legend is `false` (no `<c:legend>` element will be
   * emitted) — there is no slot to host the layout on a hidden legend,
   * so leaking the value into the output would carry a pin Excel never
   * reads. The grammar mirrors `legendOverlay` / `legendEntries` /
   * `legendFontSize` so the legend knobs compose the same way at the
   * call site.
   */
  legendLayout?: ChartManualLayout | null;
  /**
   * Override `SheetChart.legendFillColor`. `undefined` (or omitted)
   * inherits the source's parsed `legendFillColor`; `null` drops the
   * inherited fill (the writer emits no `<c:spPr>` block on
   * `<c:legend>`, falling back to the theme default — typically a
   * transparent background); a hex string replaces it. The override
   * is normalized through the writer-side hex-color path — accepts
   * `"FFFF00"` / `"#FFFF00"` / `"ffff00"` and collapses malformed
   * tokens (wrong length, non-hex characters, alpha-channel forms,
   * empty / whitespace-only strings, non-string escapes from an
   * untyped caller) to `undefined`. The cloned `SheetChart` always
   * carries a value the writer will accept; malformed source values
   * likewise collapse on the resolver path.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved legend is `false` (no `<c:legend>` element will
   * be emitted) — there is no slot to host the fill on a hidden
   * legend, so leaking the value into the output would carry a pin
   * Excel never reads.
   *
   * The grammar mirrors `plotAreaFillColor` / `titleColor` /
   * `axes.x.axisTitleColor` / `axes.x.labelColor` /
   * `legendFontColor` so the fill / color knobs compose the same way
   * at the call site. The override lands on the legend's `<c:spPr>`
   * block and composes independently with `legendFontColor` (which
   * lands on the legend's `<c:txPr>` block) — the two knobs target
   * different children of `<c:legend>` so a caller can pin both
   * without conflict.
   */
  legendFillColor?: string | null;
  /**
   * Override `SheetChart.legendBorderColor`. `undefined` (or omitted)
   * inherits the source's parsed `legendBorderColor`; `null` drops the
   * inherited stroke (the writer falls back to the auto-stroke Excel
   * picks from the chart's theme — no `<a:ln>` block on the legend's
   * `<c:spPr>`); a 6-digit RGB hex string replaces it.
   *
   * The override runs through the same sRGB normalizer as the writer —
   * the leading `#` and case are accepted, then the value collapses to
   * the OOXML canonical uppercase form. Malformed tokens (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` so the cloned
   * `SheetChart` drops the field rather than carry a value the writer
   * would silently elide back to absence.
   *
   * The grammar mirrors `legendFillColor` so the legend `<c:spPr>`
   * knobs compose the same way at the call site. Composes
   * independently with `legendFillColor` — the two knobs land on the
   * same `<c:spPr>` block but on different children (`<a:solidFill>`
   * for fill, `<a:ln>` for stroke), and the writer emits `<c:spPr>`
   * whenever either knob is set. The override is silently dropped from
   * the cloned `SheetChart` when the resolved legend is `false` (no
   * `<c:legend>` element will be emitted) — there is no slot to host
   * the stroke on a hidden legend.
   */
  legendBorderColor?: string | null;
  /**
   * Override `SheetChart.legendBorderWidth`. `undefined` (or omitted)
   * inherits the source's parsed `legendBorderWidth`; `null` drops the
   * inherited width (the writer emits `<a:ln>` without a `w` attribute,
   * the line keeps Excel's auto-thickness — typically 0.75 pt); a
   * finite point value (e.g. `1.5`) replaces it.
   *
   * The override runs through the same clamp / snap as the writer —
   * values are clamped to the `0.25..13.5` pt band Excel's UI exposes
   * and snapped to the 0.25 pt grid so a parsed-then-written width does
   * not drift across round-trips. Non-finite / non-numeric tokens
   * (`NaN`, `Infinity`, strings, `null` from an untyped caller) collapse
   * to `undefined` so the cloned `SheetChart` drops the field rather
   * than carry a value the writer would silently elide back to absence.
   *
   * Composes independently with `legendBorderColor` — both knobs land
   * on the same `<a:ln>` element but on a different slot (the color's
   * `<a:solidFill>` child versus the line's `w` attribute). A caller
   * can pin a width without a color (the border picks Excel's
   * auto-color), pin a color without a width (the border picks Excel's
   * auto-thickness), or pin both. The override is silently dropped from
   * the cloned `SheetChart` when the resolved legend is `false` (no
   * `<c:legend>` element will be emitted) — there is no slot to host
   * the stroke on a hidden legend. Mirrors `plotAreaBorderWidth` —
   * same accept-finite-number / clamp / snap grammar — but on the
   * legend's own `<c:spPr>` block.
   */
  legendBorderWidth?: number | null;
  /**
   * Override `SheetChart.legendBorderDash`. `undefined` (or omitted)
   * inherits the source's parsed dash; `null` drops the inherited
   * dash (the writer renders solid); a {@link ChartBorderDash} value
   * replaces it.
   *
   * Composes independently with `legendBorderColor` and
   * `legendBorderWidth`. Silently dropped from the cloned `SheetChart`
   * when the resolved legend is `false`.
   */
  legendBorderDash?: ChartBorderDash | null;
  /**
   * Override `SheetChart.plotAreaLayout`. `undefined` (or omitted)
   * inherits the source's parsed `plotAreaLayout`; `null` drops the
   * inherited layout (the writer emits the bare `<c:layout/>`
   * placeholder, falling back to Excel's auto-layout position); a
   * {@link ChartManualLayout} replaces it.
   *
   * Each of {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
   * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} runs
   * through the writer-side normalizer — coordinates outside the
   * `0..1` band, `NaN`, `Infinity`, and non-numeric overrides all
   * collapse to omitting the matching `<c:x>` / `<c:y>` / `<c:w>` /
   * `<c:h>` slot so the cloned `SheetChart` always carries a value the
   * writer will accept. An override whose every coordinate dropped on
   * normalization collapses the entire layout to `undefined` so the
   * writer emits the bare `<c:layout/>` placeholder.
   *
   * The grammar mirrors `legendLayout` so the manual-layout knobs
   * compose the same way at the call site. Unlike `legendLayout`, the
   * plot-area layout is never gated on a visibility flag — every chart
   * has a `<c:plotArea>` element to host the layout.
   */
  plotAreaLayout?: ChartManualLayout | null;
  /**
   * Override `SheetChart.plotAreaFillColor`. `undefined` (or omitted)
   * inherits the source's parsed `plotAreaFillColor`; `null` drops the
   * inherited fill (the writer falls back to the auto-fill Excel picks
   * from the chart's theme — no `<c:spPr>` block on the plot area); a
   * 6-digit RGB hex string replaces it.
   *
   * The override runs through the same sRGB normalizer as the writer —
   * the leading `#` and case are accepted, then the value collapses to
   * the OOXML canonical uppercase form. Malformed tokens (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` so the cloned
   * `SheetChart` drops the field rather than carry a value the writer
   * would silently elide back to absence.
   *
   * The grammar mirrors `titleColor` / `axes.x.axisTitleColor` /
   * `axes.x.labelColor` so the chart `<a:srgbClr>` knobs compose the
   * same way at the call site. Unlike the title / axis-title color
   * knobs, the plot-area fill is never gated on a visibility flag —
   * every chart has a `<c:plotArea>` element to host the fill.
   */
  plotAreaFillColor?: string | null;
  /**
   * Override `SheetChart.plotAreaBorderColor`. `undefined` (or omitted)
   * inherits the source's parsed `plotAreaBorderColor`; `null` drops
   * the inherited stroke (the writer falls back to the auto-stroke
   * Excel picks from the chart's theme — no `<a:ln>` block on the plot
   * area's `<c:spPr>`); a 6-digit RGB hex string replaces it.
   *
   * The override runs through the same sRGB normalizer as the writer —
   * the leading `#` and case are accepted, then the value collapses to
   * the OOXML canonical uppercase form. Malformed tokens (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` so the cloned
   * `SheetChart` drops the field rather than carry a value the writer
   * would silently elide back to absence.
   *
   * The grammar mirrors `plotAreaFillColor` so the chart `<c:spPr>`
   * knobs compose the same way at the call site. Composes
   * independently with `plotAreaFillColor` — the two knobs land on
   * the same `<c:spPr>` block but on different children
   * (`<a:solidFill>` for fill, `<a:ln>` for stroke), and the writer
   * emits `<c:spPr>` whenever either knob is set. Like the fill knob,
   * the border is never gated on a visibility flag — every chart has
   * a `<c:plotArea>` element to host the stroke.
   */
  plotAreaBorderColor?: string | null;
  /**
   * Override `SheetChart.plotAreaBorderWidth`. `undefined` (or omitted)
   * inherits the source's parsed `plotAreaBorderWidth`; `null` drops
   * the inherited width (the writer emits `<a:ln>` without a `w`
   * attribute, the line keeps Excel's auto-thickness — typically
   * 0.75 pt); a finite point value (e.g. `1.5`) replaces it.
   *
   * The override runs through the same clamp / snap as the writer —
   * values are clamped to the `0.25..13.5` pt band Excel's UI exposes
   * and snapped to the 0.25 pt grid so a parsed-then-written width does
   * not drift across round-trips. Non-finite / non-numeric tokens
   * (`NaN`, `Infinity`, strings, `null` from an untyped caller) collapse
   * to `undefined` so the cloned `SheetChart` drops the field rather
   * than carry a value the writer would silently elide back to absence.
   *
   * Composes independently with `plotAreaBorderColor` — both knobs land
   * on the same `<a:ln>` element but on a different slot (the color's
   * `<a:solidFill>` child versus the line's `w` attribute). A caller
   * can pin a width without a color (the border picks Excel's
   * auto-color), pin a color without a width (the border picks Excel's
   * auto-thickness), or pin both. Like the color knob, the width is
   * never gated on a visibility flag — every chart has a `<c:plotArea>`
   * element to host the stroke.
   */
  plotAreaBorderWidth?: number | null;
  /**
   * Override `SheetChart.plotAreaBorderDash`. `undefined` (or omitted)
   * inherits the source's parsed dash; `null` drops the inherited
   * dash (the writer renders solid); a {@link ChartBorderDash} value
   * replaces it. Unrecognized tokens (and the OOXML default `"solid"`)
   * collapse to `undefined` so the cloned `SheetChart` drops the field.
   *
   * Composes independently with `plotAreaBorderColor` and
   * `plotAreaBorderWidth` — all three knobs share the same `<a:ln>`
   * element. Like the color / width knobs, the dash is never gated on
   * a visibility flag — every chart has a `<c:plotArea>` element.
   */
  plotAreaBorderDash?: ChartBorderDash | null;
  /**
   * Override `SheetChart.chartSpaceFillColor`. `undefined` (or omitted)
   * inherits the source's parsed `chartSpaceFillColor`; `null` drops
   * the inherited fill (the writer emits no `<c:spPr>` block on
   * `<c:chartSpace>`, falling back to the auto-fill Excel picks from
   * the workbook theme — typically opaque white); a 6-digit RGB hex
   * string replaces it.
   *
   * The override runs through the same sRGB normalizer as the writer —
   * the leading `#` and case are accepted, then the value collapses to
   * the OOXML canonical uppercase form. Malformed tokens (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` so the cloned
   * `SheetChart` drops the field rather than carry a value the writer
   * would silently elide back to absence.
   *
   * The grammar mirrors `plotAreaFillColor` / `legendFillColor` /
   * `titleColor` so the chart `<a:srgbClr>` fill / color knobs compose
   * the same way at the call site. Unlike the title / legend color
   * knobs, the chart-space fill is never gated on a visibility flag —
   * every chart has a `<c:chartSpace>` document root to host the fill.
   * Composes independently with `plotAreaFillColor` — the two knobs
   * land on different host elements (`<c:chartSpace>` for the entire
   * frame, `<c:plotArea>` for the inner band that hosts the series),
   * so a caller can pin both without conflict.
   */
  chartSpaceFillColor?: string | null;
  /**
   * Override `SheetChart.chartSpaceBorderColor`. `undefined` (or
   * omitted) inherits the source's parsed `chartSpaceBorderColor`;
   * `null` drops the inherited stroke (the writer emits no `<a:ln>`
   * block on `<c:chartSpace>`'s `<c:spPr>`, falling back to the auto-
   * stroke Excel picks from the workbook theme — typically a
   * translucent gray border or no border depending on the theme); a
   * 6-digit RGB hex string replaces it.
   *
   * The override runs through the same sRGB normalizer as the writer —
   * the leading `#` and case are accepted, then the value collapses to
   * the OOXML canonical uppercase form. Malformed tokens (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` so the cloned
   * `SheetChart` drops the field rather than carry a value the writer
   * would silently elide back to absence.
   *
   * Composes independently with `chartSpaceFillColor` — the two knobs
   * land on the same `<c:spPr>` block but on different children
   * (`<a:solidFill>` for fill, `<a:ln>` for stroke), and the writer
   * emits `<c:spPr>` whenever either knob is set. Like the fill knob,
   * the border is never gated on a visibility flag — every chart has
   * a `<c:chartSpace>` document root to host the stroke.
   */
  chartSpaceBorderColor?: string | null;
  /**
   * Override `SheetChart.chartSpaceBorderWidth`. `undefined` (or
   * omitted) inherits the source's parsed width; `null` drops the
   * inherited width (the writer emits `<a:ln>` without a `w` attribute,
   * the line keeps Excel's auto-thickness — typically 0.75 pt); a
   * finite point value (e.g. `1.5`) replaces it. Values are clamped
   * to the `0.25..13.5` pt band Excel's UI exposes and snapped to the
   * 0.25 pt grid; non-finite / non-numeric overrides collapse to
   * `undefined`. Composes independently with `chartSpaceBorderColor`
   * and `chartSpaceBorderDash` — all three knobs share the same
   * `<a:ln>` element.
   */
  chartSpaceBorderWidth?: number | null;
  /**
   * Override `SheetChart.chartSpaceBorderDash`. `undefined` (or
   * omitted) inherits the source's parsed dash style; `null` drops
   * the inherited dash (the writer renders solid); a
   * {@link ChartBorderDash} value replaces it. Unrecognized tokens
   * (and the OOXML default `"solid"`) collapse to `undefined`.
   */
  chartSpaceBorderDash?: ChartBorderDash | null;
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
   * Override `SheetChart.dropLines`. `undefined` (or omitted) inherits
   * the source's parsed flag; `null` drops the inherited value (the
   * writer falls back to the OOXML default of no `<c:dropLines>`); a
   * `boolean` replaces it. Only meaningful when the resolved chart type
   * is `line` or `area`; silently dropped on every other family.
   */
  dropLines?: boolean | null;
  /**
   * Override `SheetChart.hiLowLines`. `undefined` (or omitted) inherits
   * the source's parsed flag; `null` drops the inherited value (the
   * writer falls back to the OOXML default of no `<c:hiLowLines>`); a
   * `boolean` replaces it. Only meaningful when the resolved chart type
   * is `line`; silently dropped on every other family (`<c:hiLowLines>`
   * has no slot on `<c:areaChart>` per OOXML).
   */
  hiLowLines?: boolean | null;
  /**
   * Override `SheetChart.serLines`. `undefined` (or omitted) inherits
   * the source's parsed flag; `null` drops the inherited value (the
   * writer falls back to the OOXML default of no `<c:serLines>`); a
   * `boolean` replaces it. Only meaningful when the resolved chart type
   * is `bar` or `column`; silently dropped on every other family
   * (`<c:serLines>` has no slot on `<c:lineChart>` / `<c:areaChart>` /
   * `<c:pieChart>` / `<c:scatterChart>` per OOXML).
   */
  serLines?: boolean | null;
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
  /**
   * Override the chart-level title-overlay flag. `undefined` (or
   * omitted) inherits the source's parsed value; `null` drops the
   * inherited value (the writer falls back to the OOXML `false` default
   * — the title reserves its own slot above the plot area, no overlap);
   * a `boolean` replaces it.
   *
   * The override is silently dropped from the cloned `SheetChart` when
   * the resolved chart renders no title (`title` resolved to `undefined`
   * or `showTitle === false`) — there is no `<c:title>` block to host
   * the overlay flag in either case.
   */
  titleOverlay?: boolean | null;
  /**
   * Override the chart-level title rotation in whole degrees.
   * `undefined` (or omitted) inherits the source's parsed value;
   * `null` drops the inherited rotation so the writer falls back to
   * the OOXML default `0` (horizontal); a `number` replaces it.
   *
   * Out-of-range overrides clamp to the `-90..90` band Excel's UI
   * exposes; non-integer overrides round to the nearest whole degree;
   * `0`, `NaN`, `Infinity`, and non-numeric overrides collapse to a
   * drop (the writer's normalization band) so the cloned `SheetChart`
   * always carries a value the writer will accept.
   *
   * The override is silently dropped from the cloned `SheetChart` when
   * the resolved chart renders no title (`title` resolved to `undefined`
   * or `showTitle === false`) — there is no `<c:title>` block to host
   * the rotation in either case.
   */
  titleRotation?: number | null;
  /**
   * Override the chart-level title font size in whole or half points.
   * `undefined` (or omitted) inherits the source's parsed value;
   * `null` drops the inherited size so the writer falls back to
   * Excel's default 14pt; a `number` replaces it.
   *
   * Out-of-range overrides (outside the `1..400`pt band the OOXML
   * `ST_TextFontSize` schema exposes) collapse to a drop (the writer's
   * normalization band) so the cloned `SheetChart` always carries a
   * value the writer will accept. Fractional overrides round to the
   * nearest 0.5pt (Excel's UI granularity); `NaN`, `Infinity`, and
   * non-numeric overrides also collapse to a drop.
   *
   * The override is silently dropped from the cloned `SheetChart` when
   * the resolved chart renders no title (`title` resolved to `undefined`
   * or `showTitle === false`) — there is no `<c:title>` block to host
   * the size in either case. The grammar mirrors `titleRotation` /
   * `titleOverlay` so the chart-level title knobs compose the same
   * way at the call site.
   */
  titleFontSize?: number | null;
  /**
   * Override the chart-level title bold flag.
   * `undefined` (or omitted) inherits the source's parsed value;
   * `null` drops the inherited flag so the writer falls back to the
   * OOXML default `b="0"` (non-bold); a `boolean` replaces it.
   *
   * Non-boolean overrides (typed escapes from an untyped caller)
   * collapse to a drop so the cloned `SheetChart` always carries a
   * value the writer will accept.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved chart renders no title (`title` resolved to
   * `undefined` or `showTitle === false`) — there is no `<c:title>`
   * block to host the flag in either case. The grammar mirrors
   * `titleFontSize` / `titleRotation` / `titleOverlay` so the
   * chart-level title knobs compose the same way at the call site.
   */
  titleBold?: boolean | null;
  /**
   * Override the chart-level title italic flag.
   * `undefined` (or omitted) inherits the source's parsed value;
   * `null` drops the inherited flag so the writer falls back to the
   * OOXML default (no `i` attribute, equivalent to non-italic);
   * a `boolean` replaces it.
   *
   * Non-boolean overrides (typed escapes from an untyped caller)
   * collapse to a drop so the cloned `SheetChart` always carries a
   * value the writer will accept.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved chart renders no title (`title` resolved to
   * `undefined` or `showTitle === false`) — there is no `<c:title>`
   * block to host the flag in either case. The grammar mirrors
   * `titleBold` / `titleFontSize` / `titleRotation` / `titleOverlay`
   * so the chart-level title knobs compose the same way at the call
   * site.
   */
  titleItalic?: boolean | null;
  /**
   * Override the chart-level title font color.
   * `undefined` (or omitted) inherits the source's parsed value;
   * `null` drops the inherited fill so the writer falls back to the
   * theme text color (no `<a:solidFill>` element on the title's
   * default-paragraph properties);
   * a 6-character hex string (with or without a leading `#`)
   * replaces it.
   *
   * Malformed overrides (wrong length, non-hex characters,
   * alpha-channel forms, non-string escapes) collapse to a drop so
   * the cloned `SheetChart` always carries a value the writer will
   * accept.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved chart renders no title (`title` resolved to
   * `undefined` or `showTitle === false`) — there is no `<c:title>`
   * block to host the fill in either case. The grammar mirrors
   * `titleBold` / `titleItalic` / `titleFontSize` / `titleRotation`
   * / `titleOverlay` so the chart-level title knobs compose the same
   * way at the call site.
   */
  titleColor?: string | null;
  /**
   * Override the chart-level title strikethrough flag.
   * `undefined` (or omitted) inherits the source's parsed `titleStrike`.
   * `null` drops the inherited flag (the writer falls back to the
   * OOXML default — no `strike` attribute, equivalent to no
   * strikethrough).
   * A `boolean` replaces it: `true` emits `strike="sngStrike"`
   * (Excel's UI "Strikethrough" — single line); `false` pins the
   * non-default omission (functionally identical to dropping).
   *
   * Non-boolean overrides (typed escapes from an untyped caller)
   * collapse to a drop so the cloned `SheetChart` always carries a
   * value the writer will accept.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved chart renders no title (`title` resolved to
   * `undefined` or `showTitle === false`) — there is no `<c:title>`
   * block to host the flag in either case. The grammar mirrors
   * `titleBold` / `titleItalic` / `titleColor` / `titleFontSize` /
   * `titleRotation` / `titleOverlay` so the chart-level title knobs
   * compose the same way at the call site.
   */
  titleStrike?: boolean | null;
  /**
   * Override the chart-level title underline flag.
   * `undefined` (or omitted) inherits the source's parsed `titleUnderline`.
   * `null` drops the inherited flag (the writer falls back to the
   * OOXML default — no `u` attribute, equivalent to no underline).
   * A `boolean` replaces it: `true` emits `u="sng"` (Excel's UI
   * "Underline" — single line); `false` pins the non-default omission
   * (functionally identical to dropping).
   *
   * Non-boolean overrides (typed escapes from an untyped caller)
   * collapse to a drop so the cloned `SheetChart` always carries a
   * value the writer will accept.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved chart renders no title (`title` resolved to
   * `undefined` or `showTitle === false`) — there is no `<c:title>`
   * block to host the flag in either case. The grammar mirrors
   * `titleBold` / `titleItalic` / `titleStrike` / `titleColor` /
   * `titleFontSize` / `titleRotation` / `titleOverlay` so the
   * chart-level title knobs compose the same way at the call site.
   */
  titleUnderline?: boolean | null;
  /**
   * Override the chart-level title font family / typeface.
   * `undefined` (or omitted) inherits the source's parsed
   * `titleFontFamily`; `null` drops the inherited typeface so the
   * writer falls back to the OOXML default (no `<a:latin>` element,
   * the title inherits the theme typeface); a non-empty string
   * replaces it.
   *
   * The override is trimmed; empty / whitespace-only strings and
   * non-string overrides (typed escapes from an untyped caller)
   * collapse to a drop so the cloned `SheetChart` always carries a
   * value the writer will accept.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved chart renders no title (`title` resolved to
   * `undefined` or `showTitle === false`) — there is no `<c:title>`
   * block to host the typeface in either case. The grammar mirrors
   * `titleColor` (the other string-typed knob) and `titleBold` /
   * `titleItalic` / `titleStrike` / `titleUnderline` /
   * `titleFontSize` / `titleRotation` / `titleOverlay` so the
   * chart-level title knobs compose the same way at the call site.
   */
  titleFontFamily?: string | null;
  /**
   * Override the chart-level title manual layout. `undefined` (or
   * omitted) inherits the source's parsed `titleLayout`; `null` drops
   * the inherited layout so the writer falls back to Excel's auto-
   * layout position (the title renders above the plot area, no
   * `<c:layout>` block on `<c:title>`); a {@link ChartManualLayout}
   * replaces it.
   *
   * Each of {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
   * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} runs
   * through the writer-side `0..1` band — out-of-range / non-finite /
   * non-numeric coordinates collapse to `undefined` on the matching
   * axis so the cloned `SheetChart` always carries a value the writer
   * will accept; an override whose every axis dropped collapses to
   * `undefined` so the writer skips the `<c:layout>` block entirely.
   *
   * The override is silently dropped from the cloned `SheetChart` when
   * the resolved chart renders no title (`title` resolved to `undefined`
   * or `showTitle === false`) — there is no `<c:title>` block to host
   * the layout in either case. The grammar mirrors `legendLayout` so
   * the chart-level manual-layout knobs compose the same way at the
   * call site.
   */
  titleLayout?: ChartManualLayout | null;
  /**
   * Override `SheetChart.titleFillColor`. `undefined` (or omitted)
   * inherits the source's parsed `titleFillColor`; `null` drops the
   * inherited fill (the writer emits no `<c:spPr>` block on
   * `<c:title>`, falling back to the theme default — typically a
   * transparent title background); a hex string replaces it. The
   * override is normalized through the writer-side hex-color path —
   * accepts `"FFFF00"` / `"#FFFF00"` / `"ffff00"` and collapses
   * malformed tokens (wrong length, non-hex characters, alpha-channel
   * forms, empty / whitespace-only strings, non-string escapes from
   * an untyped caller) to `undefined`. The cloned `SheetChart` always
   * carries a value the writer will accept; malformed source values
   * likewise collapse on the resolver path.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved chart renders no title (`title` resolved to
   * `undefined` or `showTitle === false`) — there is no `<c:title>`
   * block to host the fill on a hidden title, so leaking the value
   * into the output would carry a pin Excel never reads.
   *
   * The grammar mirrors `plotAreaFillColor` / `legendFillColor` /
   * `titleColor` / `axes.x.axisTitleColor` so the fill / color knobs
   * compose the same way at the call site. The override lands on the
   * title's `<c:spPr>` block and composes independently with
   * `titleColor` (which lands on the title's
   * `<c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>` slot) — the
   * two knobs target different children of `<c:title>` so a caller
   * can pin both without conflict.
   */
  titleFillColor?: string | null;
  /**
   * Override `SheetChart.titleBorderColor`. `undefined` (or omitted)
   * inherits the source's parsed `titleBorderColor`; `null` drops the
   * inherited stroke (the writer emits no `<a:ln>` block on
   * `<c:title><c:spPr>`, falling back to the theme default — typically
   * no visible border); a hex string replaces it. The override is
   * normalized through the writer-side hex-color path — accepts
   * `"1F77B4"` / `"#1F77B4"` / `"1f77b4"` and collapses malformed
   * tokens (wrong length, non-hex characters, alpha-channel forms,
   * empty / whitespace-only strings, non-string escapes from an
   * untyped caller) to `undefined`. The cloned `SheetChart` always
   * carries a value the writer will accept; malformed source values
   * likewise collapse on the resolver path.
   *
   * The override is silently dropped from the cloned `SheetChart`
   * when the resolved chart renders no title (`title` resolved to
   * `undefined` or `showTitle === false`) — there is no `<c:title>`
   * block to host the stroke on a hidden title, so leaking the
   * value into the output would carry a pin Excel never reads.
   *
   * The grammar mirrors `plotAreaBorderColor` so the chart `<c:spPr>`
   * stroke knobs compose the same way at the call site. The override
   * lands on the title's `<c:spPr><a:ln>` block and composes
   * independently with `titleFillColor` — the two knobs share the
   * `<c:spPr>` host but land on different children
   * (`<a:solidFill>` for fill, `<a:ln>` for stroke), and the writer
   * authors a `<c:spPr>` whenever either knob resolves to a value.
   */
  titleBorderColor?: string | null;
  /**
   * Override `SheetChart.titleBorderWidth`. `undefined` (or omitted)
   * inherits the source's parsed `titleBorderWidth`; `null` drops the
   * inherited width (the writer emits `<a:ln>` without a `w` attribute,
   * the line keeps Excel's auto-thickness — typically 0.75 pt); a
   * finite point value (e.g. `1.5`) replaces it.
   *
   * The override runs through the same clamp / snap as the writer —
   * values are clamped to the `0.25..13.5` pt band Excel's UI exposes
   * and snapped to the 0.25 pt grid so a parsed-then-written width does
   * not drift across round-trips. Non-finite / non-numeric tokens
   * (`NaN`, `Infinity`, strings, `null` from an untyped caller) collapse
   * to `undefined` so the cloned `SheetChart` drops the field rather
   * than carry a value the writer would silently elide back to absence.
   *
   * Composes independently with `titleBorderColor` — both knobs land
   * on the same `<a:ln>` element but on a different slot (the color's
   * `<a:solidFill>` child versus the line's `w` attribute). A caller
   * can pin a width without a color (the border picks Excel's
   * auto-color), pin a color without a width (the border picks Excel's
   * auto-thickness), or pin both. The override is silently dropped
   * from the cloned `SheetChart` when the resolved chart renders no
   * title (`title` resolved to `undefined` or `showTitle === false`) —
   * there is no `<c:title>` block to host the stroke on a hidden
   * title. Mirrors `plotAreaBorderWidth` and `legendBorderWidth` —
   * same accept-finite-number / clamp / snap grammar — but on the
   * title's own `<c:spPr>` block.
   */
  titleBorderWidth?: number | null;
  /**
   * Override `SheetChart.titleBorderDash`. `undefined` (or omitted)
   * inherits the source's parsed dash; `null` drops the inherited
   * dash (the writer renders solid); a {@link ChartBorderDash} value
   * replaces it. Unrecognized tokens (and the OOXML default `"solid"`)
   * collapse to `undefined`.
   *
   * Composes independently with `titleBorderColor` and
   * `titleBorderWidth` — all three knobs share the same `<a:ln>`
   * element. Silently dropped from the cloned `SheetChart` when the
   * resolved chart renders no title.
   */
  titleBorderDash?: ChartBorderDash | null;
  /**
   * Override `<c:autoTitleDeleted>` (the "user explicitly deleted the
   * auto-generated title" flag).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * `autoTitleDeleted`. `null` drops the inherited value so the
   * writer falls back to its derived default (the value pinned by
   * the title presence on the cloned chart — `true` when no literal
   * title is rendered, `false` when one is). A `boolean` replaces it.
   *
   * The override is independent of the resolved title — `<c:autoTitleDeleted>`
   * sits on `<c:chart>` directly (not nested inside `<c:title>`), so
   * a clone with no literal title can still pin `false` to let Excel
   * synthesise the auto-title from the series name, and a clone with
   * a literal title can pin `true` to suppress the synthesis even
   * though the literal renders.
   *
   * The grammar mirrors `titleOverlay` / `roundedCorners` /
   * `plotVisOnly` so the chart-level title flags compose the same way
   * at the call site.
   */
  autoTitleDeleted?: boolean | null;
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
   * Override `<c:showDLblsOverMax>` (the "show data labels for values
   * over maximum scale" toggle).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * `showDLblsOverMax`. `null` drops the inherited value so the writer
   * falls back to the OOXML `true` default (labels render for every
   * point regardless of the axis ceiling). A `boolean` replaces it —
   * useful for stripping labels off over-max points on a clone whose
   * value axis pins a tight `<c:max>` (`false`), or for restoring the
   * default behavior on a clone whose template overrode it (`true`).
   *
   * The grammar mirrors `plotVisOnly` / `dispBlanksAs` so the
   * chart-level toggles compose the same way at the call site.
   */
  showDLblsOverMax?: boolean | null;
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
   * Override `<c:upDownBars>` (the line-chart up / down bars toggle).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * `upDownBars`. `null` drops the inherited value so the writer
   * falls back to the OOXML default (no up / down bars). A `boolean`
   * replaces it — useful for adding the bars to a dashboard line clone
   * whose template did not carry them, or stripping them from a
   * template-supplied stock-style line chart.
   *
   * Only meaningful when the resolved chart type is `line` — the OOXML
   * schema places `<c:upDownBars>` on `CT_LineChart` /
   * `CT_Line3DChart` / `CT_StockChart`. The field is silently dropped
   * when the clone targets any other family (so a line-template
   * up/down-bars hint never leaks into a column / pie / doughnut /
   * area / scatter clone).
   */
  upDownBars?: boolean | null;
  /**
   * Override `<c:upDownBars><c:gapWidth val=".."/>` (the gap width
   * between up / down bars on a line chart, as a percentage of the
   * bar width).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * `upDownBarsGapWidth`. `null` drops the inherited value so the
   * writer falls back to the OOXML default `150` (Excel's reference
   * value for a freshly-toggled "Add Chart Element -> Up/Down Bars").
   * A `number` replaces it — pass any value in the inclusive `0..500`
   * band; out-of-range or non-finite values fall through to the
   * default at write time rather than emit a token Excel rejects.
   *
   * Only meaningful when the resolved chart type is `line` and the
   * resolved `upDownBars` is `true` — the writer drops the value
   * silently otherwise (the OOXML schema scopes `<c:gapWidth>`
   * exclusively to `<c:upDownBars>`, so there is no slot for it when
   * the parent element is not emitted).
   *
   * The grammar mirrors the other line-only overrides
   * ({@link upDownBars} / {@link showLineMarkers}) so the chart-level
   * line-bar knobs compose the same way at the call site.
   */
  upDownBarsGapWidth?: number | null;
  /**
   * Override `<c:lineChart><c:marker val=".."/>` (the chart-level
   * line-marker visibility toggle).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * `showLineMarkers`. `null` drops the inherited value so the writer
   * falls back to the Excel default (`<c:marker val="1"/>` — markers
   * shown). A `boolean` replaces it — `true` keeps markers on (matches
   * the default), `false` flips the chart-level gate off and emits
   * `<c:marker val="0"/>` so per-series marker definitions stop
   * rendering chart-wide.
   *
   * Only meaningful when the resolved chart type is `line` — the OOXML
   * schema places the chart-level `<c:marker>` (CT_Boolean) exclusively
   * on `CT_LineChart`. The field is silently dropped when the clone
   * targets any other family (so a line-template marker-off hint never
   * leaks into a column / pie / doughnut / area / scatter clone).
   *
   * Independent of any per-series marker overrides — this gate sits at
   * the chart level and decides whether markers paint at all; the
   * per-series block then picks the symbol / size / fill that paints
   * when the gate is open.
   */
  showLineMarkers?: boolean | null;
  /**
   * Override `<c:style>` (the built-in chart style preset, 1–48).
   *
   * `undefined` (or omitted) inherits the source's parsed `style`.
   * `null` drops the inherited value so the writer skips the element
   * entirely — Excel falls back to its application default look. A
   * number replaces the preset; out-of-range / non-integer values are
   * dropped at the writer side rather than emit a token Excel would
   * reject.
   *
   * Useful when restyling a cloned chart to a different gallery
   * preset, or stripping a template's pinned style so the clone picks
   * up the host workbook's default. The grammar mirrors
   * `roundedCorners` / `plotVisOnly` so the chart-frame toggles
   * compose the same way at the call site.
   */
  style?: number | null;
  /**
   * Override `<c:lang>` (the chart-space editing-locale hint).
   *
   * `undefined` (or omitted) inherits the source's parsed `lang`.
   * `null` drops the inherited value so the writer skips the element
   * entirely — Excel falls back to the host workbook's editing
   * language. A string replaces the locale; malformed culture names
   * are dropped at the writer side rather than emit a token Excel
   * would reject (`<c:lang>` is `xsd:language` per the OOXML schema,
   * the BCP-47 shape `[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*`, e.g.
   * `en-US`, `tr-TR`, `zh-Hant-TW`).
   *
   * Useful when restamping a templated chart for a different locale,
   * or stripping a template's pinned `en-US` so a translated
   * dashboard inherits the host workbook's locale. The grammar
   * mirrors `style` so the chart-space toggles compose the same way
   * at the call site.
   */
  lang?: string | null;
  /**
   * Override `<c:date1904>` (the chart-space date-system toggle).
   *
   * `undefined` (or omitted) inherits the source's parsed `date1904`.
   * `null` drops the inherited value so the writer skips the element
   * entirely — Excel falls back to the host workbook's date system.
   * `true` pins the chart to the 1904 base (Excel for Mac's legacy
   * epoch) and `false` collapses to absence on the writer side
   * because `<c:date1904 val="0"/>` is the OOXML default and the
   * writer follows the minimal-shape contract every other chart-space
   * toggle uses.
   *
   * Useful when restamping a chart from a 1904-based template into a
   * 1900-based workbook (or vice versa) — pinning the field keeps the
   * chart's date references anchored to the source's epoch even after
   * the host changes. The grammar mirrors `roundedCorners` /
   * `plotVisOnly` so the chart-space toggles compose the same way at
   * the call site.
   */
  date1904?: boolean | null;
  /**
   * Override `<c:plotArea><c:dTable>` (the data-table beneath the plot
   * area).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * {@link Chart.dataTable}. `null` drops the inherited block so the
   * writer skips the element entirely — Excel renders no data table.
   * `false` is equivalent to `null` (suppression). `true` pins every
   * border / outline / key flag to its OOXML default `true`. A
   * {@link ChartDataTable} object replaces the block wholesale (no
   * per-field merge; pass every flag you want preserved). Each
   * unspecified boolean flag inside the object falls back to `true` at
   * the writer side because every `<c:dTable>` boolean child is
   * required on `CT_DTable` and Excel emits all four. The optional
   * {@link ChartDataTable.fontSize} / {@link ChartDataTable.fontColor}
   * / {@link ChartDataTable.bold} typography pins survive the
   * wholesale-replace path along with the four boolean toggles, so a
   * clone that inherits a templated data-table carries the typography
   * forward unchanged.
   *
   * Only meaningful when the resolved chart type has axes — `bar`,
   * `column`, `line`, `area`, `scatter`. The field is silently dropped
   * when the clone targets `pie` / `doughnut` because the OOXML schema
   * places `<c:dTable>` inside `<c:plotArea>` alongside the axes; pie /
   * doughnut have no axes and no slot for the element.
   */
  dataTable?: ChartDataTable | boolean | null;
  /**
   * Override `<c:chartSpace><c:protection>` (the chart-space
   * protection block).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * {@link Chart.protection}. `null` drops the inherited block so the
   * writer skips the element entirely — Excel applies no chart-level
   * protection. `false` is equivalent to `null` (suppression). `true`
   * pins every flag to its OOXML default `false` so the writer emits
   * the bare `<c:protection>` shell. A {@link ChartProtection} object
   * replaces the block wholesale (no per-field merge; pass every flag
   * you want preserved). Each unspecified flag inside the object falls
   * back to `false` at the writer side — `<c:protection>` accepts
   * every child as optional and Excel treats a missing child as the
   * unlocked default.
   *
   * Applies to every chart family — `<c:protection>` lives on
   * `<c:chartSpace>` (not inside `<c:plotArea>`), so the element has
   * a slot on pie / doughnut charts too. The grammar mirrors
   * {@link CloneChartOptions.dataTable} so the chart-level block
   * toggles compose the same way at the call site.
   */
  protection?: ChartProtection | boolean | null;
  /**
   * Override `<c:chart><c:view3D>` (the 3-D rotation / perspective
   * block).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * {@link Chart.view3D}. `null` drops the inherited block so the
   * writer skips the element entirely — Excel falls back to its
   * per-family default rotation / perspective. A {@link ChartView3D}
   * object replaces the inherited block wholesale (no per-field merge;
   * pass every field you want preserved). Pass an empty object (`{}`)
   * to declare a bare `<c:view3D/>` shell — useful for round-trip
   * parity with templates that author the element with no children
   * pinned. Each unspecified field falls back to absence at the writer
   * side because every CT_View3D child is independently optional and
   * Excel treats a missing child as the per-family default.
   *
   * Applies to every chart family — `<c:view3D>` lives on `<c:chart>`
   * (between `<c:autoTitleDeleted>` and `<c:plotArea>`), so the OOXML
   * schema accepts the element on both 2D and 3D families. The toggle
   * is only meaningful on 3D families (`bar3D`, `line3D`, `pie3D`,
   * `area3D`, `surface3D`), but the writer carries a templated value
   * through every clone so a 3D template chart round-trips losslessly.
   * The grammar mirrors {@link CloneChartOptions.protection} so the
   * chart-level block toggles compose the same way at the call site.
   */
  view3D?: ChartView3D | null;
  /**
   * Override `<c:chart><c:floor><c:thickness val="N"/></c:floor>`
   * (the 3-D floor extrusion thickness on `CT_Surface`, ECMA-376
   * Part 1, §21.2.2.69).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * {@link Chart.floorThickness}. `null` drops the inherited value so
   * the writer skips `<c:floor>` entirely — Excel falls back to its
   * per-family floor default (no extrusion). A `number` replaces it —
   * pass any value in the inclusive `1..100` band Excel's "Format
   * Floor -> Floor -> Thickness" pane exposes; out-of-range, `0`, or
   * non-finite values fall through to absence at write time rather
   * than emit a token Excel rejects.
   *
   * Applies to every chart family — `<c:floor>` lives on `<c:chart>`
   * (between `<c:view3D>` and `<c:plotArea>`), so the OOXML schema
   * accepts the element on both 2D and 3D families. The toggle is
   * only meaningful on 3D families (`bar3D`, `line3D`, `pie3D`,
   * `area3D`, `surface3D`), but the writer carries a templated value
   * through every clone so a 3D template chart round-trips losslessly.
   * The grammar mirrors {@link CloneChartOptions.upDownBarsGapWidth}
   * so the chart-level numeric knobs compose the same way at the call
   * site.
   */
  floorThickness?: number | null;
  /**
   * Override `<c:chart><c:sideWall><c:thickness val="N"/></c:sideWall>`
   * (the 3-D side-wall extrusion thickness on `CT_Surface`, ECMA-376
   * Part 1, §21.2.2.187).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * {@link Chart.sideWallThickness}. `null` drops the inherited value
   * so the writer skips `<c:sideWall>` entirely — Excel falls back to
   * its per-family side-wall default (no extrusion). A `number`
   * replaces it — pass any value in the inclusive `1..100` band
   * Excel's "Format Side Wall -> Side Wall -> Thickness" pane
   * exposes; out-of-range, `0`, or non-finite values fall through to
   * absence at write time rather than emit a token Excel rejects.
   *
   * Applies to every chart family — `<c:sideWall>` lives on
   * `<c:chart>` (between `<c:floor>` and `<c:backWall>` /
   * `<c:plotArea>`), so the OOXML schema accepts the element on both
   * 2D and 3D families. The toggle is only meaningful on 3D families
   * (`bar3D`, `line3D`, `pie3D`, `area3D`, `surface3D`), but the
   * writer carries a templated value through every clone so a 3D
   * template chart round-trips losslessly. The grammar mirrors
   * {@link CloneChartOptions.upDownBarsGapWidth} so the chart-level
   * numeric knobs compose the same way at the call site.
   */
  sideWallThickness?: number | null;
  /**
   * Override `<c:chart><c:backWall><c:thickness val="N"/></c:backWall>`
   * (the 3-D back-wall extrusion thickness on `CT_Surface`, ECMA-376
   * Part 1, §21.2.2.31).
   *
   * `undefined` (or omitted) inherits the source's parsed
   * {@link Chart.backWallThickness}. `null` drops the inherited value
   * so the writer skips `<c:backWall>` entirely — Excel falls back to
   * its per-family back-wall default (no extrusion). A `number`
   * replaces it — pass any value in the inclusive `1..100` band
   * Excel's "Format Back Wall -> Back Wall -> Thickness" pane exposes;
   * out-of-range, `0`, or non-finite values fall through to absence at
   * write time rather than emit a token Excel rejects.
   *
   * Applies to every chart family — `<c:backWall>` lives on `<c:chart>`
   * (between `<c:sideWall>` and `<c:plotArea>`), so the OOXML schema
   * accepts the element on both 2D and 3D families. The toggle is only
   * meaningful on 3D families (`bar3D`, `line3D`, `pie3D`, `area3D`,
   * `surface3D`), but the writer carries a templated value through
   * every clone so a 3D template chart round-trips losslessly. The
   * grammar mirrors {@link CloneChartOptions.floorThickness} so the
   * chart-level numeric knobs compose the same way at the call site.
   */
  backWallThickness?: number | null;
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
      /**
       * Override `SheetChart.axes.x.axisTitleRotation`. `undefined` (or
       * omitted) inherits the source axis's parsed value; `null` drops
       * the inherited rotation (the writer falls back to the OOXML
       * default `0` — the title renders horizontally); a number in the
       * `-90..90` band replaces it (out-of-range and non-finite inputs
       * collapse to `undefined`).
       *
       * `<c:title>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts (no axes at all) and on any axis
       * whose `title` is unset (no `<c:title>` block to host the
       * rotation).
       */
      axisTitleRotation?: number | null;
      /**
       * Override `SheetChart.axes.x.axisTitleFontSize`. `undefined`
       * (or omitted) inherits the source axis's parsed value; `null`
       * drops the inherited size (the writer falls back to the
       * hardcoded `1000` (10pt) default Excel itself emits on a fresh
       * axis title); a number in the `1..400`pt band replaces it
       * (out-of-range, non-finite, and non-numeric inputs collapse
       * to `undefined`, and fractional inputs round to the nearest
       * 0.5pt).
       *
       * `<c:title>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently
       * dropped on `pie` / `doughnut` charts (no axes at all) and on
       * any axis whose `title` is unset (no `<c:title>` block to
       * host the size).
       */
      axisTitleFontSize?: number | null;
      /**
       * Override `SheetChart.axes.x.axisTitleBold`. `undefined` (or
       * omitted) inherits the source axis's parsed flag; `null` drops
       * the inherited flag (the writer falls back to the OOXML
       * default `b="0"` — the title renders non-bold); a `boolean`
       * replaces it.
       *
       * Non-boolean overrides (typed escapes from an untyped caller)
       * collapse to a drop so the cloned `SheetChart` always carries
       * a value the writer will accept.
       *
       * `<c:title>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts (no axes at all) and on any axis
       * whose `title` is unset (no `<c:title>` block to host the
       * flag).
       */
      axisTitleBold?: boolean | null;
      /**
       * Override `SheetChart.axes.x.axisTitleItalic`. `undefined` (or
       * omitted) inherits the source axis's parsed value; `null` drops
       * the inherited flag (the writer falls back to the OOXML default
       * — no `i` attribute, equivalent to non-italic); a literal
       * `boolean` replaces it. Non-boolean overrides (typed escape
       * from an untyped caller) collapse to `undefined`, so the
       * cloned `SheetChart` always carries a value the writer will
       * accept.
       *
       * `<c:title>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently
       * dropped on `pie` / `doughnut` charts (no axes at all) and on
       * any axis whose `title` is unset (no `<c:title>` block to
       * host the flag).
       */
      axisTitleItalic?: boolean | null;
      /**
       * Override `SheetChart.axes.x.axisTitleColor`. `undefined` (or
       * omitted) inherits the source axis's parsed value; `null` drops
       * the inherited fill so the writer falls back to the theme text
       * color (no `<a:solidFill>` element on the axis title's
       * default-paragraph properties); a 6-character hex string (with
       * or without a leading `#`, any case) replaces it.
       *
       * Malformed overrides (wrong length, non-hex characters,
       * alpha-channel forms, non-string escapes from an untyped
       * caller) collapse to a drop so the cloned `SheetChart` always
       * carries a value the writer will accept.
       *
       * `<c:title>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts (no axes at all) and on any axis
       * whose `title` is unset (no `<c:title>` block to host the
       * fill). The grammar mirrors `axisTitleRotation` /
       * `axisTitleFontSize` / `axisTitleBold` / `axisTitleItalic` so
       * the axis-title knobs compose the same way at the call site.
       */
      axisTitleColor?: string | null;
      /**
       * Override `SheetChart.axes.x.axisTitleStrike`. `undefined` (or
       * omitted) inherits the source axis's parsed value; `null` drops
       * the inherited flag so the writer falls back to the OOXML
       * default — no `strike` attribute, equivalent to no
       * strikethrough. A `boolean` replaces it: `true` emits
       * `strike="sngStrike"` (Excel's UI "Strikethrough" — single line);
       * `false` pins the non-default omission (functionally identical
       * to dropping).
       *
       * Non-boolean overrides (typed escapes from an untyped caller)
       * collapse to a drop so the cloned `SheetChart` always carries
       * a value the writer will accept.
       *
       * `<c:title>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts (no axes at all) and on any axis
       * whose `title` is unset (no `<c:title>` block to host the
       * flag). The grammar mirrors `axisTitleRotation` /
       * `axisTitleFontSize` / `axisTitleBold` / `axisTitleItalic` /
       * `axisTitleColor` so the axis-title knobs compose the same way
       * at the call site.
       */
      axisTitleStrike?: boolean | null;
      /**
       * Override `SheetChart.axes.x.axisTitleUnderline`. `undefined`
       * (or omitted) inherits the source axis's parsed value; `null`
       * drops the inherited flag so the writer falls back to the
       * OOXML default — no `u` attribute, equivalent to no underline.
       * A `boolean` replaces it: `true` emits `u="sng"` (Excel's UI
       * "Underline" — single line); `false` pins the non-default
       * omission (functionally identical to dropping).
       *
       * Non-boolean overrides (typed escapes from an untyped caller)
       * collapse to a drop so the cloned `SheetChart` always carries
       * a value the writer will accept.
       *
       * `<c:title>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts (no axes at all) and on any axis
       * whose `title` is unset (no `<c:title>` block to host the
       * flag). The grammar mirrors `axisTitleRotation` /
       * `axisTitleFontSize` / `axisTitleBold` / `axisTitleItalic` /
       * `axisTitleColor` / `axisTitleStrike` so the axis-title knobs
       * compose the same way at the call site.
       */
      axisTitleUnderline?: boolean | null;
      /**
       * Override `SheetChart.axes.x.axisTitleFontFamily`. `undefined`
       * (or omitted) inherits the source axis's parsed typeface;
       * `null` drops the inherited typeface so the writer falls back
       * to the OOXML default (no `<a:latin>` element, the title
       * inherits the theme typeface). A non-empty string replaces
       * it; the override is trimmed.
       *
       * Empty / whitespace-only strings and non-string overrides
       * (typed escapes from an untyped caller) collapse to a drop so
       * the cloned `SheetChart` always carries a value the writer
       * will accept.
       *
       * `<c:title>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts (no axes at all) and on any
       * axis whose `title` is unset (no `<c:title>` block to host
       * the typeface). The grammar mirrors `titleFontFamily` (the
       * chart-level analog) and `axisTitleColor` (the other string-
       * typed knob) / `axisTitleRotation` / `axisTitleFontSize` /
       * `axisTitleBold` / `axisTitleItalic` / `axisTitleStrike` /
       * `axisTitleUnderline` so the axis-title knobs compose the
       * same way at the call site.
       */
      axisTitleFontFamily?: string | null;
      /**
       * Override `SheetChart.axes.x.axisTitleOverlay`. `undefined`
       * (or omitted) inherits the source axis's parsed value; `null`
       * drops the inherited value (the writer falls back to the OOXML
       * `false` default — the title reserves its own slot adjacent to
       * the axis, no overlap with the plot area); a `boolean`
       * replaces it.
       *
       * The override is silently dropped from the cloned `SheetChart`
       * when the axis renders no title (the resolved `title` is
       * `undefined`) — there is no `<c:title>` block to host the
       * overlay flag in either case.
       *
       * The grammar mirrors `titleOverlay` (the chart-level analog)
       * and the other axis-title knobs (`axisTitleRotation` /
       * `axisTitleFontSize` / `axisTitleBold` / `axisTitleItalic` /
       * `axisTitleColor` / `axisTitleStrike` / `axisTitleUnderline` /
       * `axisTitleFontFamily`) so the axis-title knobs compose the
       * same way at the call site.
       */
      axisTitleOverlay?: boolean | null;
      /**
       * Override `SheetChart.axes.x.axisTitleLayout`. `undefined` (or
       * omitted) inherits the source axis's parsed `axisTitleLayout`;
       * `null` drops the inherited layout (the writer falls back to
       * Excel's auto-layout position — no `<c:layout>` block on the
       * axis title); a {@link ChartManualLayout} replaces it.
       *
       * Each of {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
       * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} runs
       * through the same `0..1` filter as the writer-side
       * normalization: out-of-range / non-finite / non-numeric tokens
       * collapse on the matching axis. An override whose every
       * coordinate dropped collapses to `undefined` so the cloned
       * `SheetChart` skips the entire `<c:layout>` block.
       *
       * The override is silently dropped from the cloned `SheetChart`
       * when the axis renders no title (the resolved `title` is
       * `undefined`) — there is no `<c:title>` block to host the
       * layout in either case.
       *
       * The grammar mirrors the chart-level `titleLayout` /
       * `legendLayout` / `plotAreaLayout` so the four manual-layout
       * knobs compose the same way at the call site.
       */
      axisTitleLayout?: ChartManualLayout | null;
      /**
       * Override `SheetChart.axes.x.axisTitleFillColor`. `undefined`
       * (or omitted) inherits the source axis's parsed value; `null`
       * drops the inherited fill so the writer falls back to the
       * theme default fill (no `<c:spPr>` block on the axis title,
       * typically a transparent title background); a 6-character hex
       * string (with or without a leading `#`, any case) replaces it.
       *
       * Malformed overrides (wrong length, non-hex characters,
       * alpha-channel forms, non-string escapes from an untyped
       * caller) collapse to a drop so the cloned `SheetChart` always
       * carries a value the writer will accept.
       *
       * `<c:title>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts (no axes at all) and on any axis
       * whose `title` is unset (no `<c:title>` block to host the
       * fill). Composes independently with `axisTitleColor` (the font
       * color) — the two knobs target different children of
       * `<c:title>`.
       *
       * Mirrors the chart-level `titleFillColor` so a single
       * configuration call can thread a fill through the chart title
       * and either axis title without bookkeeping the canonical OOXML
       * slots.
       */
      axisTitleFillColor?: string | null;
      /**
       * Override `SheetChart.axes.x.axisTitleBorderColor`. `undefined`
       * (or omitted) inherits the source axis's parsed value; `null`
       * drops the inherited stroke so the writer falls back to the
       * theme-default auto-stroke (no `<a:ln>` block on the axis
       * title, typically no visible border); a 6-character hex string
       * (with or without a leading `#`, any case) replaces it.
       *
       * Malformed overrides (wrong length, non-hex characters,
       * alpha-channel forms, non-string escapes from an untyped
       * caller) collapse to a drop so the cloned `SheetChart` always
       * carries a value the writer will accept.
       *
       * `<c:title>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts (no axes at all) and on any
       * axis whose `title` is unset (no `<c:title>` block to host
       * the stroke). Composes independently with `axisTitleFillColor`
       * (the background fill on `<c:spPr><a:solidFill>`) and
       * `axisTitleColor` (the font color on the inner
       * `<a:defRPr><a:solidFill>`) — the three knobs target different
       * children of `<c:title>` so a caller can pin all three without
       * conflict.
       *
       * Mirrors the chart-level `titleBorderColor` so a single
       * configuration call can thread a border color through the
       * chart title and either axis title without bookkeeping the
       * canonical OOXML slots.
       */
      axisTitleBorderColor?: string | null;
      /**
       * Override `SheetChart.axes.x.axisTitleBorderWidth`. Same
       * `undefined` / `null` / number grammar as the chart-level
       * `titleBorderWidth` knob — values are clamped to the
       * `0.25..13.5` pt band Excel's UI exposes and snapped to the
       * 0.25 pt grid; non-finite / non-numeric overrides collapse to
       * `undefined`. Silently dropped on `pie` / `doughnut` charts
       * (no axes) and on any axis whose `title` is unset.
       */
      axisTitleBorderWidth?: number | null;
      /**
       * Override `SheetChart.axes.x.axisTitleBorderDash`. Same
       * `undefined` / `null` / value grammar as the chart-level
       * `titleBorderDash` knob — `"solid"` and unrecognized tokens
       * collapse to `undefined`. Silently dropped on `pie` /
       * `doughnut` charts and on any axis whose `title` is unset.
       */
      axisTitleBorderDash?: ChartBorderDash | null;
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
       * Override `SheetChart.axes.x.labelRotation`. `undefined` (or
       * omitted) inherits the source axis's rotation; `null` drops the
       * inherited value (the writer falls back to the OOXML default `0`
       * — labels render flat); a number in the `-90..90` band replaces
       * it (out-of-range and non-finite inputs collapse to `undefined`).
       *
       * `<c:txPr>` lives on every axis flavour per the OOXML schema, so
       * the override carries through every chart family that has axes
       * (bar / column / line / area / scatter). Silently dropped on
       * `pie` / `doughnut` charts since neither has axes.
       */
      labelRotation?: number | null;
      /**
       * Override `SheetChart.axes.x.labelFontSize`. `undefined` (or
       * omitted) inherits the source axis's tick-label font size;
       * `null` drops the inherited value (the writer falls back to
       * Excel's reference 10pt); a number in the `1..400`pt band
       * replaces it (out-of-range, non-finite, and non-numeric inputs
       * collapse to `undefined`). Fractional inputs round to the
       * nearest 0.5pt, matching Excel's UI granularity.
       *
       * `<c:txPr>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts since neither has axes. Composes
       * independently with {@link labelRotation}: both fields land on
       * the same `<c:txPr>` body.
       */
      labelFontSize?: number | null;
      /**
       * Override `SheetChart.axes.x.labelBold`. `undefined` (or
       * omitted) inherits the source axis's tick-label bold flag;
       * `null` drops the inherited value (the writer falls back to
       * the OOXML default — the tick labels render at the theme's
       * default weight); a `boolean` replaces it.
       *
       * Non-boolean overrides (typed escapes from an untyped caller)
       * collapse to a drop so the cloned `SheetChart` always carries
       * a value the writer will accept.
       *
       * `<c:txPr>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts since neither has axes. Composes
       * independently with {@link labelRotation} / {@link labelFontSize}:
       * all three knobs land on the same `<c:txPr>` body.
       */
      labelBold?: boolean | null;
      /**
       * Override `SheetChart.axes.x.labelItalic`. `undefined` (or
       * omitted) inherits the source axis's tick-label italic flag;
       * `null` drops the inherited value (the writer falls back to
       * the OOXML default — the tick labels render at the theme's
       * default slant); a `boolean` replaces it.
       *
       * Non-boolean overrides (typed escapes from an untyped caller)
       * collapse to a drop so the cloned `SheetChart` always carries
       * a value the writer will accept.
       *
       * `<c:txPr>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts since neither has axes. Composes
       * independently with {@link labelRotation} / {@link labelFontSize} /
       * {@link labelBold}: all four knobs land on the same `<c:txPr>` body.
       */
      labelItalic?: boolean | null;
      /**
       * Override `SheetChart.axes.x.labelColor`. `undefined` (or
       * omitted) inherits the source axis's tick-label color;
       * `null` drops the inherited fill (the writer falls back to the
       * theme text color — no `<a:solidFill>` block on the axis tick-
       * label `<c:txPr>` default-paragraph properties); a 6-character
       * hex string (with or without a leading `#`, any case) replaces
       * it.
       *
       * Malformed overrides (wrong length, non-hex characters,
       * alpha-channel forms, non-string escapes from an untyped
       * caller) collapse to a drop so the cloned `SheetChart` always
       * carries a value the writer will accept.
       *
       * `<c:txPr>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts since neither has axes. Composes
       * independently with {@link labelRotation} / {@link labelFontSize} /
       * {@link labelBold} / {@link labelItalic}: all five knobs land
       * on the same `<c:txPr>` body.
       */
      labelColor?: string | null;
      /**
       * Override `SheetChart.axes.x.labelUnderline`. `undefined` (or
       * omitted) inherits the source axis's tick-label underline flag;
       * `null` drops the inherited value (the writer falls back to the
       * OOXML default — the tick labels render non-underlined); a
       * `boolean` replaces it.
       *
       * Non-boolean overrides (typed escapes from an untyped caller)
       * collapse to a drop so the cloned `SheetChart` always carries
       * a value the writer will accept.
       *
       * `<c:txPr>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts since neither has axes. Composes
       * independently with {@link labelRotation} / {@link labelFontSize} /
       * {@link labelBold} / {@link labelItalic} / {@link labelColor}:
       * all six knobs land on the same `<c:txPr>` body.
       */
      labelUnderline?: boolean | null;
      /**
       * Override `SheetChart.axes.x.labelStrike`. `undefined` (or
       * omitted) inherits the source axis's tick-label strikethrough
       * flag; `null` drops the inherited value (the writer falls back
       * to the OOXML default — the tick labels render non-
       * strikethrough); a `boolean` replaces it.
       *
       * Non-boolean overrides (typed escapes from an untyped caller)
       * collapse to a drop so the cloned `SheetChart` always carries
       * a value the writer will accept.
       *
       * `<c:txPr>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts since neither has axes. Composes
       * independently with {@link labelRotation} / {@link labelFontSize} /
       * {@link labelBold} / {@link labelItalic} / {@link labelColor} /
       * {@link labelUnderline}: all seven knobs land on the same
       * `<c:txPr>` body.
       */
      labelStrike?: boolean | null;
      /**
       * Override `SheetChart.axes.x.labelFontFamily`. `undefined` (or
       * omitted) inherits the source axis's tick-label typeface;
       * `null` drops the inherited typeface (the writer falls back to
       * the OOXML default — no `<a:latin>` element, the labels
       * inherit the theme typeface); a non-empty string replaces it;
       * the override is trimmed.
       *
       * Empty / whitespace-only strings and non-string overrides
       * (typed escapes from an untyped caller) collapse to a drop so
       * the cloned `SheetChart` always carries a value the writer
       * will accept.
       *
       * `<c:txPr>` lives on every axis flavour per the OOXML schema,
       * so the override carries through every chart family that has
       * axes (bar / column / line / area / scatter). Silently dropped
       * on `pie` / `doughnut` charts since neither has axes. Composes
       * independently with {@link labelRotation} / {@link labelFontSize} /
       * {@link labelBold} / {@link labelItalic} / {@link labelColor} /
       * {@link labelUnderline} / {@link labelStrike}: all eight knobs
       * land on the same `<c:txPr>` body.
       */
      labelFontFamily?: string | null;
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
       * Override `SheetChart.axes.x.auto`. `undefined` (or omitted)
       * inherits the source axis's flag; `null` drops the inherited
       * value (the writer falls back to the OOXML `true` default —
       * Excel auto-detects whether to render the axis as a date or
       * category axis); a `boolean` replaces it. Only `false` actually
       * surfaces on the cloned `SheetChart` (the writer treats `true`
       * and absence identically — both produce `<c:auto val="1"/>`),
       * so an override of `true` collapses to `undefined`.
       *
       * Only meaningful for resolved chart types whose X axis is
       * `<c:catAx>` (bar / column / line / area); silently dropped on
       * scatter and pie / doughnut.
       */
      auto?: boolean | null;
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
      /**
       * Override `SheetChart.axes.x.crosses`. `undefined` (or omitted)
       * inherits the source axis's semantic crossing pin; `null` drops
       * the inherited value (the writer falls back to the OOXML default
       * `"autoZero"`); a {@link ChartAxisCrosses} token replaces it.
       *
       * Mutually exclusive with {@link crossesAt} — when both are set
       * (here or on the source chart) the writer favours `crossesAt`,
       * mirroring how the OOXML schema places the two elements in an
       * XSD choice. Silently dropped on `pie` / `doughnut` charts since
       * neither has axes.
       */
      crosses?: ChartAxisCrosses | null;
      /**
       * Override `SheetChart.axes.x.crossesAt`. `undefined` (or omitted)
       * inherits the source axis's numeric crossing pin; `null` drops
       * the inherited value (the writer falls back to the semantic
       * crossing pin from {@link crosses}, or to the OOXML default
       * `"autoZero"`); a finite number replaces it. `0` is preserved —
       * it is a valid pin, distinct from the `"autoZero"` default.
       *
       * When set, takes precedence over {@link crosses} because the
       * OOXML schema places `<c:crosses>` and `<c:crossesAt>` in an XSD
       * choice — only one may legally appear at a time.
       */
      crossesAt?: number | null;
      /**
       * Override `SheetChart.axes.x.dispUnits`. `undefined` (or omitted)
       * inherits the source axis's parsed display-unit preset; `null`
       * drops the inherited value (the writer leaves Excel's default
       * "no display unit" state untouched); a {@link ChartAxisDispUnit}
       * shorthand or a {@link ChartAxisDispUnits} object replaces it.
       *
       * `<c:dispUnits>` lives exclusively on `<c:valAx>` per the OOXML
       * schema, so the override only takes effect when the resolved
       * chart type routes the X axis through `<c:valAx>` — that is the
       * scatter family. Bar / column / line / area route the X axis
       * through `<c:catAx>` (which rejects `<c:dispUnits>`); the
       * resolver collapses the field to `undefined` on those families
       * so a stale hint never leaks into the writer. Pie / doughnut
       * have no axes at all.
       */
      dispUnits?: ChartAxisDispUnits | ChartAxisDispUnit | null;
      /**
       * Override `SheetChart.axes.x.crossBetween`. `undefined` (or
       * omitted) inherits the source axis's parsed cross-between mode;
       * `null` drops the inherited value (the writer falls back to the
       * per-family default each axis builder pins today); a
       * {@link ChartAxisCrossBetween} token replaces it.
       *
       * `<c:crossBetween>` lives exclusively on `<c:valAx>` per the
       * OOXML schema, so the override only takes effect when the
       * resolved chart type routes the X axis through `<c:valAx>` —
       * that is the scatter family. Bar / column / line / area route
       * the X axis through `<c:catAx>` (which rejects
       * `<c:crossBetween>`); the resolver collapses the field to
       * `undefined` on those families so a stale hint never leaks into
       * the writer. Pie / doughnut have no axes at all.
       */
      crossBetween?: ChartAxisCrossBetween | null;
    };
    y?: {
      title?: string | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleRotation}. */
      axisTitleRotation?: number | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleFontSize}. */
      axisTitleFontSize?: number | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleBold}. */
      axisTitleBold?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleItalic}. */
      axisTitleItalic?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleColor}. */
      axisTitleColor?: string | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleStrike}. */
      axisTitleStrike?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleUnderline}. */
      axisTitleUnderline?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleFontFamily}. */
      axisTitleFontFamily?: string | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleOverlay}. */
      axisTitleOverlay?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleLayout}. */
      axisTitleLayout?: ChartManualLayout | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleFillColor}. */
      axisTitleFillColor?: string | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleBorderColor}. */
      axisTitleBorderColor?: string | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleBorderWidth}. */
      axisTitleBorderWidth?: number | null;
      /** See {@link CloneChartOptions.axes.x.axisTitleBorderDash}. */
      axisTitleBorderDash?: ChartBorderDash | null;
      gridlines?: ChartAxisGridlines | null;
      scale?: ChartAxisScale | null;
      numberFormat?: ChartAxisNumberFormat | null;
      /** See {@link CloneChartOptions.axes.x.majorTickMark}. */
      majorTickMark?: ChartAxisTickMark | null;
      /** See {@link CloneChartOptions.axes.x.minorTickMark}. */
      minorTickMark?: ChartAxisTickMark | null;
      /** See {@link CloneChartOptions.axes.x.tickLblPos}. */
      tickLblPos?: ChartAxisTickLabelPosition | null;
      /** See {@link CloneChartOptions.axes.x.labelRotation}. */
      labelRotation?: number | null;
      /** See {@link CloneChartOptions.axes.x.labelFontSize}. */
      labelFontSize?: number | null;
      /** See {@link CloneChartOptions.axes.x.labelBold}. */
      labelBold?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.labelItalic}. */
      labelItalic?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.labelColor}. */
      labelColor?: string | null;
      /** See {@link CloneChartOptions.axes.x.labelUnderline}. */
      labelUnderline?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.labelStrike}. */
      labelStrike?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.labelFontFamily}. */
      labelFontFamily?: string | null;
      /** See {@link CloneChartOptions.axes.x.hidden}. */
      hidden?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.reverse}. */
      reverse?: boolean | null;
      /** See {@link CloneChartOptions.axes.x.crosses}. */
      crosses?: ChartAxisCrosses | null;
      /** See {@link CloneChartOptions.axes.x.crossesAt}. */
      crossesAt?: number | null;
      /**
       * Override `SheetChart.axes.y.dispUnits`. Same `undefined` /
       * `null` / replace grammar as
       * {@link CloneChartOptions.axes.x.dispUnits}.
       *
       * The Y axis is a value axis on every chart family that has axes
       * — bar / column / line / area / scatter — so the override
       * always takes effect on those families. Pie / doughnut have no
       * axes at all and the resolver collapses the field to `undefined`
       * on those types.
       */
      dispUnits?: ChartAxisDispUnits | ChartAxisDispUnit | null;
      /**
       * Override `SheetChart.axes.y.crossBetween`. Same `undefined`
       * (inherit) / `null` (drop) / replace grammar as
       * {@link CloneChartOptions.axes.x.crossBetween}.
       *
       * The Y axis is a value axis on every chart family that has axes
       * — bar / column / line / area / scatter — so the override always
       * takes effect on those families. Pie / doughnut have no axes at
       * all and the resolver collapses the field to `undefined` on
       * those types.
       */
      crossBetween?: ChartAxisCrossBetween | null;
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

  // `<c:explosion>` lives exclusively on pie-family series (CT_PieSer,
  // shared across `<c:pieChart>` / `<c:doughnutChart>` via EG_PieSer);
  // drop the field from every other resolved type so a pie → bar
  // flatten (or any other coercion) does not leak the value into a
  // chart kind whose schema rejects it.
  if (type !== "pie" && type !== "doughnut") {
    for (const s of series) {
      if (s.explosion !== undefined) delete s.explosion;
    }
  }

  // `<c:trendline>` and `<c:errBars>` live on bar / column / line /
  // area / scatter / bubble series — never on pie / doughnut. Drop
  // the fields when the resolved family is pie / doughnut so a pie →
  // line clone (or any other coercion) does not leak the inherited
  // arrays into the cloned chart.
  if (type === "pie" || type === "doughnut") {
    for (const s of series) {
      if (s.trendlines !== undefined) delete s.trendlines;
      if (s.errorBars !== undefined) delete s.errorBars;
    }
  }

  // `<c:bubbleSize>` and `<c:shape>` are family-scoped: bubbleSize on
  // bubble series, shape on bar3D series. Hucre's writer authors
  // neither family today, but the fields live in the cloned `SheetChart`
  // so a templated chart's metadata round-trips. Drop bubbleSize on
  // every family except `bubble`-coerced clones (none exist on the
  // writer); the writer in series.ts ignores both fields when the
  // chart kind cannot host them.
  if (type !== "bar" && type !== "column") {
    for (const s of series) {
      if (s.shape3D !== undefined) delete s.shape3D;
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
    const resolvedLegendOverlay = resolveCloneLegendOverlay(
      source.legendOverlay,
      options.legendOverlay,
    );
    if (resolvedLegendOverlay !== undefined) out.legendOverlay = resolvedLegendOverlay;

    // `<c:legendEntry>` lives inside `<c:legend>`, so the same hidden /
    // missing-legend scoping that drops `legendOverlay` also drops the
    // inherited entry list. Mirrors the legendOverlay grammar:
    // `undefined` inherits the parsed value, `null` drops it (the writer
    // emits no `<c:legendEntry>` children), a `ChartLegendEntry[]`
    // replaces it outright.
    const resolvedLegendEntries = resolveCloneLegendEntries(
      source.legendEntries,
      options.legendEntries,
    );
    if (resolvedLegendEntries !== undefined) out.legendEntries = resolvedLegendEntries;

    // `<c:txPr>` only renders inside `<c:legend>`, so the same hidden /
    // missing-legend scoping that drops `legendOverlay` /
    // `legendEntries` also drops the inherited font size — the writer
    // has no slot to populate when the legend is hidden. Mirrors the
    // `titleFontSize` / `axisTitleFontSize` grammar: `undefined`
    // inherits the parsed value (after running it through the
    // half-step / range normalizer), `null` drops it (the writer
    // emits no `<c:txPr>` block), a number replaces.
    const resolvedLegendFontSize = resolveCloneLegendFontSize(
      source.legendFontSize,
      options.legendFontSize,
    );
    if (resolvedLegendFontSize !== undefined) out.legendFontSize = resolvedLegendFontSize;

    // Same hidden-legend scoping for the bold flag: the writer has no
    // `<c:txPr>` slot to populate when the legend is hidden. Mirrors the
    // `titleBold` / `axisTitleBold` grammar: `undefined` inherits
    // (after running through the boolean normalizer), `null` drops it
    // (the writer emits no `<c:txPr>` block), a `boolean` replaces.
    const resolvedLegendBold = resolveCloneLegendBold(source.legendBold, options.legendBold);
    if (resolvedLegendBold !== undefined) out.legendBold = resolvedLegendBold;

    // Same hidden-legend scoping for the italic flag: `undefined`
    // inherits (after the boolean normalizer), `null` drops, a
    // `boolean` replaces.
    const resolvedLegendItalic = resolveCloneLegendItalic(
      source.legendItalic,
      options.legendItalic,
    );
    if (resolvedLegendItalic !== undefined) out.legendItalic = resolvedLegendItalic;

    // Same hidden-legend scoping for the underline flag.
    const resolvedLegendUnderline = resolveCloneLegendUnderline(
      source.legendUnderline,
      options.legendUnderline,
    );
    if (resolvedLegendUnderline !== undefined) out.legendUnderline = resolvedLegendUnderline;

    // Same hidden-legend scoping for the strikethrough flag.
    const resolvedLegendStrikethrough = resolveCloneLegendStrikethrough(
      source.legendStrikethrough,
      options.legendStrikethrough,
    );
    if (resolvedLegendStrikethrough !== undefined) {
      out.legendStrikethrough = resolvedLegendStrikethrough;
    }

    // Same hidden-legend scoping for the font color — `<c:txPr>` is
    // the shared host element. `undefined` inherits, `null` drops, a
    // hex string replaces.
    const resolvedLegendFontColor = resolveCloneLegendFontColor(
      source.legendFontColor,
      options.legendFontColor,
    );
    if (resolvedLegendFontColor !== undefined) out.legendFontColor = resolvedLegendFontColor;

    // Same hidden-legend scoping for the font family — `<c:txPr>` is
    // the shared host element. `undefined` inherits, `null` drops, a
    // non-empty string replaces. Empty / whitespace-only / non-string
    // overrides collapse via the normalizer so the cloned
    // `SheetChart` always carries a value the writer will accept.
    const resolvedLegendFontFamily = resolveCloneLegendFontFamily(
      source.legendFontFamily,
      options.legendFontFamily,
    );
    if (resolvedLegendFontFamily !== undefined) out.legendFontFamily = resolvedLegendFontFamily;

    // Same hidden-legend scoping for the manual layout — `<c:layout>`
    // lives inside `<c:legend>` per CT_Legend, so a clone whose legend
    // is hidden has no slot for the placement. `undefined` inherits the
    // parsed value (after the writer-side normalizer drops out-of-range
    // axes), `null` drops it (the writer emits no `<c:layout>` block,
    // falling back to Excel's auto-layout position), a
    // `ChartManualLayout` replaces it. An override whose every
    // coordinate dropped on normalization collapses the entire layout
    // to `undefined` so the writer skips the `<c:layout>` block.
    const resolvedLegendLayout = resolveCloneLegendLayout(
      source.legendLayout,
      options.legendLayout,
    );
    if (resolvedLegendLayout !== undefined) out.legendLayout = resolvedLegendLayout;

    // Same hidden-legend scoping for the background fill — `<c:spPr>`
    // is a direct child of `<c:legend>` per CT_Legend, so a clone whose
    // legend is hidden has no slot for the fill. `undefined` inherits,
    // `null` drops, a hex string replaces. The override runs through
    // the same hex normalizer as the writer so the cloned `SheetChart`
    // always carries a value the writer will accept; malformed source
    // values likewise collapse on the resolver path. Composes
    // independently with `legendFontColor` — the two knobs target
    // different children of `<c:legend>` (`<c:spPr>` for the fill,
    // `<c:txPr>` for the font color).
    const resolvedLegendFillColor = resolveCloneLegendFillColor(
      source.legendFillColor,
      options.legendFillColor,
    );
    if (resolvedLegendFillColor !== undefined) out.legendFillColor = resolvedLegendFillColor;

    // Same hidden-legend scoping for the border (line) stroke —
    // `<a:ln>` lives inside `<c:legend><c:spPr>` per CT_ShapeProperties,
    // so a clone whose legend is hidden has no slot for the stroke.
    // `undefined` inherits the source's parsed `legendBorderColor`
    // (after the writer-side normalizer collapses any malformed token),
    // `null` drops the inherited stroke (the writer emits no `<a:ln>`
    // block, the legend inherits the auto-stroke Excel picks from the
    // chart's theme), a 6-digit hex string replaces it. Composes
    // independently with `legendFillColor` — the two knobs share the
    // `<c:spPr>` host but land on different children (`<a:solidFill>`
    // for fill, `<a:ln>` for stroke).
    const resolvedLegendBorderColor = resolveCloneLegendBorderColor(
      source.legendBorderColor,
      options.legendBorderColor,
    );
    if (resolvedLegendBorderColor !== undefined) {
      out.legendBorderColor = resolvedLegendBorderColor;
    }

    // Same hidden-legend scoping for the border (line) stroke width —
    // `<a:ln w=..>` lives inside `<c:legend><c:spPr>` per
    // CT_ShapeProperties, so a clone whose legend is hidden has no slot
    // for the width. `undefined` inherits the source's parsed
    // `legendBorderWidth` (after the writer-side clamp collapses any
    // malformed token), `null` drops the inherited width (the writer
    // emits `<a:ln>` without the `w` attribute, the line keeps Excel's
    // auto-thickness), a finite point value replaces it. Malformed
    // overrides collapse to `undefined` via the normalizer. Composes
    // independently with `legendBorderColor` — both knobs land on the
    // same `<a:ln>` element but on a different attribute (color is
    // `<a:solidFill><a:srgbClr>`, width is the `w` attribute on
    // `<a:ln>`).
    const resolvedLegendBorderWidth = resolveCloneLegendBorderWidth(
      source.legendBorderWidth,
      options.legendBorderWidth,
    );
    if (resolvedLegendBorderWidth !== undefined) {
      out.legendBorderWidth = resolvedLegendBorderWidth;
    }

    // Legend border preset dash pattern — same hidden-legend scoping
    // as the color / width knobs above.
    const resolvedLegendBorderDash = resolveBorderDash(
      source.legendBorderDash,
      options.legendBorderDash,
    );
    if (resolvedLegendBorderDash !== undefined) {
      out.legendBorderDash = resolvedLegendBorderDash;
    }
  }

  // Plot-area manual layout is independent of the legend visibility —
  // every chart has a `<c:plotArea>` element to host `<c:layout>`. The
  // resolution mirrors `resolveCloneLegendLayout` exactly: `undefined`
  // inherits the source's parsed `plotAreaLayout` (after normalization
  // drops any out-of-range axes), `null` drops the inherited layout
  // (the writer falls back to the bare `<c:layout/>` placeholder), a
  // `ChartManualLayout` replaces it. An override whose every coordinate
  // dropped on normalization collapses the entire layout to `undefined`
  // so the writer skips the `<c:manualLayout>` body.
  const resolvedPlotAreaLayout = resolveClonePlotAreaLayout(
    source.plotAreaLayout,
    options.plotAreaLayout,
  );
  if (resolvedPlotAreaLayout !== undefined) out.plotAreaLayout = resolvedPlotAreaLayout;

  // Plot-area solid fill color is independent of the legend / title
  // visibility — every chart has a `<c:plotArea>` element to host the
  // `<c:spPr>` slot. `undefined` inherits the source's parsed
  // `plotAreaFillColor` (after the writer-side normalizer collapses any
  // malformed token), `null` drops the inherited fill (the writer
  // emits no `<c:spPr>` block, the plot area inherits the auto-fill
  // Excel picks from the chart's theme), a 6-digit hex string replaces
  // it. Malformed overrides collapse to `undefined` via the normalizer
  // so the cloned `SheetChart` always carries a value the writer will
  // accept.
  const resolvedPlotAreaFillColor = resolveClonePlotAreaFillColor(
    source.plotAreaFillColor,
    options.plotAreaFillColor,
  );
  if (resolvedPlotAreaFillColor !== undefined) out.plotAreaFillColor = resolvedPlotAreaFillColor;

  // Plot-area border (stroke) color is independent of the fill — every
  // chart has a `<c:plotArea>` element to host the `<c:spPr><a:ln>`
  // slot. `undefined` inherits the source's parsed `plotAreaBorderColor`
  // (after the writer-side normalizer collapses any malformed token),
  // `null` drops the inherited stroke (the writer emits no `<a:ln>`
  // block, the plot area inherits the auto-stroke Excel picks from
  // the chart's theme), a 6-digit hex string replaces it. Composes
  // independently with `plotAreaFillColor` — the two knobs share the
  // `<c:spPr>` host but land on different children (`<a:solidFill>`
  // for fill, `<a:ln>` for stroke).
  const resolvedPlotAreaBorderColor = resolveClonePlotAreaBorderColor(
    source.plotAreaBorderColor,
    options.plotAreaBorderColor,
  );
  if (resolvedPlotAreaBorderColor !== undefined) {
    out.plotAreaBorderColor = resolvedPlotAreaBorderColor;
  }

  // Plot-area border thickness composes independently with the border
  // color — both lands on the same `<a:ln>` element but on a different
  // attribute (color is `<a:solidFill><a:srgbClr>`, width is the `w`
  // attribute on `<a:ln>`). `undefined` inherits the source's parsed
  // `plotAreaBorderWidth` (after the writer-side clamp collapses any
  // malformed token), `null` drops the inherited width (the writer
  // emits `<a:ln>` without the `w` attribute, the line keeps Excel's
  // auto-thickness), a finite point value replaces it. Malformed
  // overrides collapse to `undefined` via the normalizer.
  const resolvedPlotAreaBorderWidth = resolveClonePlotAreaBorderWidth(
    source.plotAreaBorderWidth,
    options.plotAreaBorderWidth,
  );
  if (resolvedPlotAreaBorderWidth !== undefined) {
    out.plotAreaBorderWidth = resolvedPlotAreaBorderWidth;
  }

  // Plot-area border preset dash pattern. Same `<a:ln>` host as the
  // color and width knobs above, but lands on the `<a:prstDash>` child.
  // `"solid"` collapses to `undefined` so absence and the OOXML default
  // round-trip identically.
  const resolvedPlotAreaBorderDash = resolveBorderDash(
    source.plotAreaBorderDash,
    options.plotAreaBorderDash,
  );
  if (resolvedPlotAreaBorderDash !== undefined) {
    out.plotAreaBorderDash = resolvedPlotAreaBorderDash;
  }

  // Chart-space solid fill color is independent of every visibility
  // flag — every chart has a `<c:chartSpace>` document root to host the
  // `<c:spPr>` slot. `undefined` inherits the source's parsed
  // `chartSpaceFillColor` (after the writer-side normalizer collapses
  // any malformed token), `null` drops the inherited fill (the writer
  // emits no `<c:spPr>` block on `<c:chartSpace>`, the chart inherits
  // the auto-fill Excel picks from the workbook theme — typically
  // opaque white), a 6-digit hex string replaces it. Malformed
  // overrides collapse to `undefined` via the normalizer so the cloned
  // `SheetChart` always carries a value the writer will accept.
  // Composes independently with `plotAreaFillColor` — the two knobs
  // land on different host elements (`<c:chartSpace>` for the entire
  // frame, `<c:plotArea>` for the inner band that hosts the series).
  const resolvedChartSpaceFillColor = resolveCloneChartSpaceFillColor(
    source.chartSpaceFillColor,
    options.chartSpaceFillColor,
  );
  if (resolvedChartSpaceFillColor !== undefined) {
    out.chartSpaceFillColor = resolvedChartSpaceFillColor;
  }

  // Chart-space border (stroke) color is independent of every visibility
  // flag — every chart has a `<c:chartSpace>` document root to host the
  // `<c:spPr>` slot. `undefined` inherits the source's parsed
  // `chartSpaceBorderColor` (after the writer-side normalizer collapses
  // any malformed token), `null` drops the inherited stroke (the writer
  // emits no `<a:ln>` block — the chart inherits the auto-stroke Excel
  // picks from the workbook theme), a 6-digit hex string replaces it.
  // Malformed overrides collapse to `undefined` via the normalizer so
  // the cloned `SheetChart` always carries a value the writer will
  // accept. Composes independently with `chartSpaceFillColor` — the two
  // knobs land on different children of the same `<c:spPr>` block.
  const resolvedChartSpaceBorderColor = resolveCloneChartSpaceBorderColor(
    source.chartSpaceBorderColor,
    options.chartSpaceBorderColor,
  );
  if (resolvedChartSpaceBorderColor !== undefined) {
    out.chartSpaceBorderColor = resolvedChartSpaceBorderColor;
  }

  // Chart-space border (stroke) thickness shares the same `<a:ln>`
  // host as the color knob above, but lands on the `w` attribute
  // rather than the `<a:solidFill>` child. Independent of every
  // visibility flag — every chart has a `<c:chartSpace>` document root
  // to host the slot. `undefined` inherits the source's parsed width
  // (after clamp / snap), `null` drops the inherited width (the writer
  // omits the `w` attribute, the line keeps Excel's auto-thickness),
  // a finite number replaces it. Composes independently with
  // `chartSpaceBorderColor` and `chartSpaceBorderDash`.
  const resolvedChartSpaceBorderWidth = resolveBorderWidthPt(
    source.chartSpaceBorderWidth,
    options.chartSpaceBorderWidth,
  );
  if (resolvedChartSpaceBorderWidth !== undefined) {
    out.chartSpaceBorderWidth = resolvedChartSpaceBorderWidth;
  }

  // Chart-space border preset dash pattern. Same `<a:ln>` host as the
  // color and width knobs above, but lands on the `<a:prstDash>` child.
  // `"solid"` collapses to `undefined` so absence and the OOXML default
  // round-trip identically.
  const resolvedChartSpaceBorderDash = resolveBorderDash(
    source.chartSpaceBorderDash,
    options.chartSpaceBorderDash,
  );
  if (resolvedChartSpaceBorderDash !== undefined) {
    out.chartSpaceBorderDash = resolvedChartSpaceBorderDash;
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

  // `<c:dropLines>` lives on `<c:lineChart>` / `<c:line3DChart>` /
  // `<c:areaChart>` / `<c:area3DChart>` per the OOXML schema. Hucre's
  // writer authors `<c:lineChart>` and `<c:areaChart>` only, so the
  // flag carries through line / area resolutions and is dropped on
  // every other family — coercing a line template into a column clone
  // therefore never leaks the connector lines into a chart kind whose
  // schema rejects the element.
  if (type === "line" || type === "area") {
    const dropLines = resolveCloneDropLines(source.dropLines, options.dropLines);
    if (dropLines !== undefined) out.dropLines = dropLines;
  }

  // `<c:hiLowLines>` lives on `<c:lineChart>` / `<c:line3DChart>` /
  // `<c:stockChart>` per the OOXML schema. Hucre's writer authors
  // `<c:lineChart>` only, so the flag carries through line resolutions
  // and is dropped on every other family — coercing a line template
  // into an area clone therefore never leaks the connector lines into
  // a chart kind whose schema rejects the element.
  if (type === "line") {
    const hiLowLines = resolveCloneHiLowLines(source.hiLowLines, options.hiLowLines);
    if (hiLowLines !== undefined) out.hiLowLines = hiLowLines;
  }

  // `<c:serLines>` lives on `<c:barChart>` / `<c:ofPieChart>` per the
  // OOXML schema. Hucre's writer authors `<c:barChart>` only (`bar` /
  // `column` from the public `SheetChart.type` enum both resolve to
  // `<c:barChart>`), so the flag carries through bar / column
  // resolutions and is dropped on every other family — coercing a
  // stacked-bar template into a line / pie / area clone therefore
  // never leaks the connector lines into a chart kind whose schema
  // rejects the element.
  if (type === "bar" || type === "column") {
    const serLines = resolveCloneSerLines(source.serLines, options.serLines);
    if (serLines !== undefined) out.serLines = serLines;
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

  // `titleOverlay` only renders inside `<c:title>`, so a clone that
  // omits the title (resolved title is undefined or `showTitle === false`)
  // drops the inherited overlay flag — there is no `<c:overlay>` slot on
  // a missing title for the writer to populate. The override wins over
  // the source's parsed value; absence inherits, `null` drops, a `boolean`
  // replaces. Mirrors the legendOverlay scoping rule.
  const titleRendered = (out.showTitle ?? Boolean(out.title)) && out.title !== undefined;
  if (titleRendered) {
    const resolvedTitleOverlay = resolveCloneTitleOverlay(
      source.titleOverlay,
      options.titleOverlay,
    );
    if (resolvedTitleOverlay !== undefined) out.titleOverlay = resolvedTitleOverlay;

    // `titleRotation` only renders inside `<c:title>` — a clone that
    // omits the title has no `<a:bodyPr rot="N"/>` slot for the writer
    // to populate. Same scope rule as `titleOverlay`: the override wins
    // over the source's parsed value; absence inherits, `null` drops,
    // a `number` replaces. Out-of-range / non-finite / non-numeric
    // overrides collapse via the writer's normalizer so the cloned
    // `SheetChart` always carries a value the writer will accept.
    const resolvedTitleRotation = resolveCloneTitleRotation(
      source.titleRotation,
      options.titleRotation,
    );
    if (resolvedTitleRotation !== undefined) out.titleRotation = resolvedTitleRotation;

    // `titleFontSize` only renders inside `<c:title>` — a clone that
    // omits the title has no `<a:defRPr sz="N"/>` slot for the writer
    // to populate. Same scope rule as `titleOverlay` / `titleRotation`:
    // the override wins over the source's parsed value; absence
    // inherits, `null` drops, a `number` replaces. Out-of-range
    // (outside the `1..400`pt band the OOXML `ST_TextFontSize` schema
    // exposes) / non-finite / non-numeric overrides collapse via the
    // normalizer so the cloned `SheetChart` always carries a value the
    // writer will accept.
    const resolvedTitleFontSize = resolveCloneTitleFontSize(
      source.titleFontSize,
      options.titleFontSize,
    );
    if (resolvedTitleFontSize !== undefined) out.titleFontSize = resolvedTitleFontSize;

    // `titleBold` only renders inside `<c:title>` — a clone that
    // omits the title has no `<a:defRPr b=".."/>` slot for the writer
    // to populate. Same scope rule as `titleOverlay` / `titleRotation`
    // / `titleFontSize`: the override wins over the source's parsed
    // value; absence inherits, `null` drops, a `boolean` replaces.
    // Non-boolean overrides collapse via the normalizer so the cloned
    // `SheetChart` always carries a value the writer will accept.
    const resolvedTitleBold = resolveCloneTitleBold(source.titleBold, options.titleBold);
    if (resolvedTitleBold !== undefined) out.titleBold = resolvedTitleBold;

    // `titleItalic` only renders inside `<c:title>` — a clone that
    // omits the title has no `<a:defRPr i=".."/>` slot for the writer
    // to populate. Same scope rule as `titleOverlay` / `titleRotation`
    // / `titleFontSize` / `titleBold`: the override wins over the
    // source's parsed value; absence inherits, `null` drops, a
    // `boolean` replaces. Non-boolean overrides collapse via the
    // normalizer so the cloned `SheetChart` always carries a value
    // the writer will accept.
    const resolvedTitleItalic = resolveCloneTitleItalic(source.titleItalic, options.titleItalic);
    if (resolvedTitleItalic !== undefined) out.titleItalic = resolvedTitleItalic;

    // `titleColor` only renders inside `<c:title>` — a clone that
    // omits the title has no `<a:defRPr><a:solidFill>` slot for the
    // writer to populate. Same scope rule as `titleOverlay` /
    // `titleRotation` / `titleFontSize` / `titleBold` / `titleItalic`:
    // the override wins over the source's parsed value; absence
    // inherits, `null` drops, a hex string replaces. Malformed
    // overrides collapse via the normalizer so the cloned
    // `SheetChart` always carries a value the writer will accept.
    const resolvedTitleColor = resolveCloneTitleColor(source.titleColor, options.titleColor);
    if (resolvedTitleColor !== undefined) out.titleColor = resolvedTitleColor;

    // `titleStrike` only renders inside `<c:title>` — a clone that
    // omits the title has no `<a:defRPr strike="..">` slot for the
    // writer to populate. Same scope rule as `titleOverlay` /
    // `titleRotation` / `titleFontSize` / `titleBold` / `titleItalic`
    // / `titleColor`: the override wins over the source's parsed
    // value; absence inherits, `null` drops, a `boolean` replaces.
    // Non-boolean overrides collapse via the normalizer so the cloned
    // `SheetChart` always carries a value the writer will accept.
    const resolvedTitleStrike = resolveCloneTitleStrike(source.titleStrike, options.titleStrike);
    if (resolvedTitleStrike !== undefined) out.titleStrike = resolvedTitleStrike;

    // `titleUnderline` only renders inside `<c:title>` — a clone that
    // omits the title has no `<a:defRPr u="..">` slot for the writer
    // to populate. Same scope rule as `titleStrike` / `titleOverlay`
    // / `titleRotation` / `titleFontSize` / `titleBold` /
    // `titleItalic` / `titleColor`: the override wins over the
    // source's parsed value; absence inherits, `null` drops, a
    // `boolean` replaces. Non-boolean overrides collapse via the
    // normalizer so the cloned `SheetChart` always carries a value
    // the writer will accept.
    const resolvedTitleUnderline = resolveCloneTitleUnderline(
      source.titleUnderline,
      options.titleUnderline,
    );
    if (resolvedTitleUnderline !== undefined) out.titleUnderline = resolvedTitleUnderline;

    // `titleFontFamily` only renders inside `<c:title>` — a clone that
    // omits the title has no `<a:defRPr><a:latin>` slot for the writer
    // to populate. Same scope rule as `titleColor` (the other string-
    // typed knob) / `titleStrike` / `titleUnderline` / `titleOverlay`
    // / `titleRotation` / `titleFontSize` / `titleBold` /
    // `titleItalic`: the override wins over the source's parsed value;
    // absence inherits, `null` drops, a non-empty string replaces.
    // Empty / whitespace-only strings and non-string overrides
    // collapse via the normalizer so the cloned `SheetChart` always
    // carries a value the writer will accept.
    const resolvedTitleFontFamily = resolveCloneTitleFontFamily(
      source.titleFontFamily,
      options.titleFontFamily,
    );
    if (resolvedTitleFontFamily !== undefined) out.titleFontFamily = resolvedTitleFontFamily;

    // `titleLayout` only renders inside `<c:title>` — a clone that
    // omits the title has no `<c:layout>` slot for the writer to
    // populate. Same scope rule as the other title knobs:
    // `undefined` inherits the parsed value (after the writer-side
    // normalizer drops out-of-range axes), `null` drops it (the writer
    // emits no `<c:layout>` block, falling back to Excel's auto-
    // layout position above the plot area), a `ChartManualLayout`
    // replaces it. An override whose every coordinate dropped on
    // normalization collapses the entire layout to `undefined` so the
    // writer skips the `<c:layout>` block. Mirrors the legendLayout
    // grammar — both manual-layout slots use the same
    // `ChartManualLayout` shape, so a caller can thread a single
    // layout value through both call sites.
    const resolvedTitleLayout = resolveCloneTitleLayout(source.titleLayout, options.titleLayout);
    if (resolvedTitleLayout !== undefined) out.titleLayout = resolvedTitleLayout;

    // `titleFillColor` only renders inside `<c:title>` — a clone that
    // omits the title has no `<c:spPr>` slot for the writer to
    // populate. Same scope rule as the typography knobs / title
    // layout: the override wins over the source's parsed value;
    // absence inherits, `null` drops, a hex string replaces. Malformed
    // overrides collapse via the normalizer so the cloned `SheetChart`
    // always carries a value the writer will accept; malformed source
    // values likewise collapse on the resolver path. Composes
    // independently with `titleColor` — the two knobs target different
    // children of `<c:title>` (`<c:spPr>` for the background fill,
    // `<c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>` for the
    // font color).
    const resolvedTitleFillColor = resolveCloneTitleFillColor(
      source.titleFillColor,
      options.titleFillColor,
    );
    if (resolvedTitleFillColor !== undefined) out.titleFillColor = resolvedTitleFillColor;

    // `titleBorderColor` only renders inside `<c:title>` — a clone
    // that omits the title has no `<c:spPr><a:ln>` slot for the writer
    // to populate. Same scope rule as the typography knobs / title
    // layout / title fill: the override wins over the source's parsed
    // value; absence inherits, `null` drops, a hex string replaces.
    // Malformed overrides collapse via the normalizer so the cloned
    // `SheetChart` always carries a value the writer will accept;
    // malformed source values likewise collapse on the resolver path.
    // Composes independently with `titleFillColor` — the two knobs
    // share the `<c:spPr>` host but land on different children
    // (`<a:solidFill>` for fill, `<a:ln>` for stroke).
    const resolvedTitleBorderColor = resolveCloneTitleBorderColor(
      source.titleBorderColor,
      options.titleBorderColor,
    );
    if (resolvedTitleBorderColor !== undefined) out.titleBorderColor = resolvedTitleBorderColor;

    // Same hidden-title scoping for the border (line) stroke width —
    // `<a:ln w=..>` lives inside `<c:title><c:spPr>` per
    // CT_ShapeProperties, so a clone whose title is hidden has no slot
    // for the width. `undefined` inherits the source's parsed
    // `titleBorderWidth` (after the writer-side clamp collapses any
    // malformed token), `null` drops the inherited width (the writer
    // emits `<a:ln>` without the `w` attribute, the line keeps Excel's
    // auto-thickness), a finite point value replaces it. Malformed
    // overrides collapse to `undefined` via the normalizer. Composes
    // independently with `titleBorderColor` — both knobs land on the
    // same `<a:ln>` element but on a different attribute (color is
    // `<a:solidFill><a:srgbClr>`, width is the `w` attribute on
    // `<a:ln>`).
    const resolvedTitleBorderWidth = resolveCloneTitleBorderWidth(
      source.titleBorderWidth,
      options.titleBorderWidth,
    );
    if (resolvedTitleBorderWidth !== undefined) out.titleBorderWidth = resolvedTitleBorderWidth;

    // Title border preset dash pattern. Same hidden-title scoping as
    // the border color / width knobs above — the writer drops the
    // dash when no title is rendered. `undefined` inherits the
    // source's parsed dash, `null` drops the inherited dash (the
    // writer renders solid), a {@link ChartBorderDash} value replaces
    // it. Unrecognized tokens (and the OOXML default `"solid"`)
    // collapse to `undefined` via the shared normalizer.
    const resolvedTitleBorderDash = resolveBorderDash(
      source.titleBorderDash,
      options.titleBorderDash,
    );
    if (resolvedTitleBorderDash !== undefined) out.titleBorderDash = resolvedTitleBorderDash;
  }

  // `<c:autoTitleDeleted>` sits on `<c:chart>` directly, not inside
  // `<c:title>`, so the override carries through every clone — independent
  // of whether the resolved chart renders a literal title. Pinning the
  // flag lets a titleless clone suppress (or keep) Excel's auto-generated
  // series-name title regardless of what the source declared. The
  // override wins over the source's parsed value; absence inherits,
  // `null` drops (writer falls back to its title-presence-derived
  // default), a `boolean` replaces.
  const resolvedAutoTitleDeleted = resolveCloneAutoTitleDeleted(
    source.autoTitleDeleted,
    options.autoTitleDeleted,
  );
  if (resolvedAutoTitleDeleted !== undefined) out.autoTitleDeleted = resolvedAutoTitleDeleted;

  const resolvedDataLabels = resolveChartDataLabels(source.dataLabels, options.dataLabels);
  if (resolvedDataLabels !== undefined) out.dataLabels = resolvedDataLabels;

  const resolvedDispBlanks = resolveCloneDispBlanksAs(source.dispBlanksAs, options.dispBlanksAs);
  if (resolvedDispBlanks !== undefined) out.dispBlanksAs = resolvedDispBlanks;

  const resolvedVaryColors = resolveCloneVaryColors(source.varyColors, options.varyColors);
  if (resolvedVaryColors !== undefined) out.varyColors = resolvedVaryColors;

  const resolvedPlotVisOnly = resolveClonePlotVisOnly(source.plotVisOnly, options.plotVisOnly);
  if (resolvedPlotVisOnly !== undefined) out.plotVisOnly = resolvedPlotVisOnly;

  const resolvedShowDLblsOverMax = resolveCloneShowDLblsOverMax(
    source.showDLblsOverMax,
    options.showDLblsOverMax,
  );
  if (resolvedShowDLblsOverMax !== undefined) out.showDLblsOverMax = resolvedShowDLblsOverMax;

  const resolvedRoundedCorners = resolveCloneRoundedCorners(
    source.roundedCorners,
    options.roundedCorners,
  );
  if (resolvedRoundedCorners !== undefined) out.roundedCorners = resolvedRoundedCorners;

  const resolvedStyle = resolveCloneStyle(source.style, options.style);
  if (resolvedStyle !== undefined) out.style = resolvedStyle;

  const resolvedLang = resolveCloneLang(source.lang, options.lang);
  if (resolvedLang !== undefined) out.lang = resolvedLang;

  const resolvedDate1904 = resolveCloneDate1904(source.date1904, options.date1904);
  if (resolvedDate1904 !== undefined) out.date1904 = resolvedDate1904;

  // `<c:dTable>` only renders inside `<c:plotArea>` alongside the axes
  // — pie / doughnut have no axes at all, so the OOXML schema places no
  // slot for the element on those families. Drop the field on those
  // resolved types so a templated bar / line / scatter chart with a
  // pinned data table does not leak the element into a doughnut clone
  // whose schema rejects it. Override wins over the source's parsed
  // value.
  if (type !== "pie" && type !== "doughnut") {
    const resolvedDataTable = resolveCloneDataTable(source.dataTable, options.dataTable);
    if (resolvedDataTable !== undefined) out.dataTable = resolvedDataTable;
  }

  // `<c:protection>` lives on `<c:chartSpace>` (not inside
  // `<c:plotArea>`), so every chart family — including pie / doughnut
  // — carries a slot for it. Override wins over the source's parsed
  // value, and the grammar follows the same `object | boolean | null`
  // shape as `dataTable` so the chart-level block toggles compose
  // identically at the call site.
  const resolvedProtection = resolveCloneProtection(source.protection, options.protection);
  if (resolvedProtection !== undefined) out.protection = resolvedProtection;

  // `<c:view3D>` lives on `<c:chart>` directly, so the OOXML schema
  // accepts it on every chart family — both 2D and 3D. The toggle is
  // only meaningful on 3D families, but the resolver applies to every
  // type so a 3D template chart round-trips losslessly through a clone
  // (and a 2D clone of a 3D template that happens to inherit the
  // value silently keeps the element — Excel ignores it on 2D).
  // Override wins over the source's parsed value, and the grammar
  // follows the standard `object | null` shape so the chart-level
  // block toggles compose the same way at the call site.
  const resolvedView3D = resolveView3D(source.view3D, options.view3D);
  if (resolvedView3D !== undefined) out.view3D = resolvedView3D;

  // `<c:floor>` lives on `<c:chart>` directly (between `<c:view3D>`
  // and `<c:plotArea>` per CT_Chart), so the OOXML schema accepts it
  // on every chart family — both 2D and 3D. The toggle is only
  // meaningful on 3D families, but the resolver applies to every type
  // so a 3D template chart round-trips losslessly through a clone
  // (and a 2D clone of a 3D template that happens to inherit the
  // value silently keeps the element — Excel ignores it on 2D).
  // Override wins over the source's parsed value, and the grammar
  // follows the standard `number | null` shape so the chart-level
  // numeric knobs compose the same way at the call site.
  const resolvedFloorThickness = resolveFloorThickness(
    source.floorThickness,
    options.floorThickness,
  );
  if (resolvedFloorThickness !== undefined) out.floorThickness = resolvedFloorThickness;

  // `<c:sideWall>` lives on `<c:chart>` directly (between `<c:floor>`
  // and `<c:backWall>` / `<c:plotArea>` per CT_Chart), so the OOXML
  // schema accepts it on every chart family — both 2D and 3D. The
  // toggle is only meaningful on 3D families, but the resolver applies
  // to every type so a 3D template chart round-trips losslessly
  // through a clone (and a 2D clone of a 3D template that happens to
  // inherit the value silently keeps the element — Excel ignores it
  // on 2D). Override wins over the source's parsed value, and the
  // grammar follows the standard `number | null` shape so the chart-
  // level numeric knobs compose the same way at the call site.
  const resolvedSideWallThickness = resolveSideWallThickness(
    source.sideWallThickness,
    options.sideWallThickness,
  );
  if (resolvedSideWallThickness !== undefined) out.sideWallThickness = resolvedSideWallThickness;

  // `<c:backWall>` lives on `<c:chart>` directly (between `<c:sideWall>`
  // and `<c:plotArea>` per CT_Chart), so the OOXML schema accepts it
  // on every chart family — both 2D and 3D. The toggle is only
  // meaningful on 3D families, but the resolver applies to every type
  // so a 3D template chart round-trips losslessly through a clone
  // (and a 2D clone of a 3D template that happens to inherit the
  // value silently keeps the element — Excel ignores it on 2D).
  // Override wins over the source's parsed value, and the grammar
  // follows the standard `number | null` shape so the chart-level
  // numeric knobs compose the same way at the call site.
  const resolvedBackWallThickness = resolveBackWallThickness(
    source.backWallThickness,
    options.backWallThickness,
  );
  if (resolvedBackWallThickness !== undefined) out.backWallThickness = resolvedBackWallThickness;

  // `<c:scatterStyle>` only renders inside `<c:scatterChart>`. Drop the
  // field on every other resolved type so a scatter template flattened
  // to line / column does not leak the preset into a chart kind whose
  // schema rejects it. Override wins over the source's parsed value.
  if (type === "scatter") {
    const resolvedScatterStyle = resolveCloneScatterStyle(
      source.scatterStyle,
      options.scatterStyle,
    );
    if (resolvedScatterStyle !== undefined) out.scatterStyle = resolvedScatterStyle;
  }

  // `<c:upDownBars>` only renders inside `<c:lineChart>` (the writer
  // never authors `<c:line3DChart>` or `<c:stockChart>`). Drop the
  // flag on every other resolved type so a line-template up/down-bars
  // hint never leaks into a column / pie / doughnut / area / scatter
  // clone — the OOXML schema places the element exclusively on the
  // line-flavored chart-type elements. Override wins over the source's
  // parsed value.
  if (type === "line") {
    const resolvedUpDownBars = resolveCloneUpDownBars(source.upDownBars, options.upDownBars);
    if (resolvedUpDownBars !== undefined) out.upDownBars = resolvedUpDownBars;

    // `<c:upDownBars><c:gapWidth>` only makes sense when the parent
    // toggle is on. Drop the gap-width value silently when the resolved
    // `upDownBars` is `false` / `undefined` — there is no `<c:upDownBars>`
    // element to host the value in either case, so leaking it through to
    // the cloned `SheetChart` would surface a setting the writer would
    // never emit anyway.
    if (resolvedUpDownBars === true) {
      const resolvedUpDownBarsGapWidth = resolveCloneUpDownBarsGapWidth(
        source.upDownBarsGapWidth,
        options.upDownBarsGapWidth,
      );
      if (resolvedUpDownBarsGapWidth !== undefined) {
        out.upDownBarsGapWidth = resolvedUpDownBarsGapWidth;
      }
    }
  }

  // `<c:marker>` (the chart-level CT_Boolean variant) lives exclusively
  // on `<c:lineChart>` per the OOXML schema. Drop the flag on every
  // other resolved type so a line-template marker-off hint never leaks
  // into a column / pie / doughnut / area / scatter clone. Override
  // wins over the source's parsed value.
  if (type === "line") {
    const resolvedShowLineMarkers = resolveShowLineMarkers(
      source.showLineMarkers,
      options.showLineMarkers,
    );
    if (resolvedShowLineMarkers !== undefined) out.showLineMarkers = resolvedShowLineMarkers;
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

// `normalizeBorderWidthPt` (aliased to the shared `clampStrokeWidthPt`)
// and `normalizeBorderDash` now live in `./chart/shape.ts`. Imported at
// the top of this module so every chart-frame slot the clone surface
// exposes shares one EMU encoding and one accept-or-drop dash grammar.

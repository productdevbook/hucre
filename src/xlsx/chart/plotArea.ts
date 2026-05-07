// ── Chart Plot Area ───────────────────────────────────────────────
// Per-host module for `<c:plotArea>` (CT_PlotArea, ECMA-376 Part 1,
// §21.2.2.145). Holds the reader / writer helpers that orchestrate the
// per-family chart-type elements (`<c:barChart>`, `<c:lineChart>`,
// `<c:areaChart>`, `<c:pieChart>`, `<c:doughnutChart>`,
// `<c:scatterChart>`) plus the plot-area-level `<c:layout>` and
// `<c:spPr>` (fill / border) slots.
//
// `buildPlotArea` (the writer-side orchestrator that assembles the
// chart-type element, axes, data-table, and plot-area styling) lives
// here too because it is plot-area-scope. It calls into chart/axis
// for `buildBarAxes` / `buildScatterAxes` and into chart/dataTable for
// `buildDataTable`.

import type {
  ChartAxisCrossBetween,
  ChartAxisInfo,
  ChartBarGrouping,
  ChartKind,
  ChartLineAreaGrouping,
  ChartManualLayout,
  ChartScatterStyle,
  SheetChart,
} from "../../_types";
import type { XmlElement } from "../../xml/parser";
import { xmlElement, xmlSelfClose } from "../../xml/writer";
import {
  EMU_PER_PT,
  clampStrokeWidthPt,
  normalizeBorderDash,
  normalizeRgbHex,
  parseBorderDashFromSpPr,
  parseBorderWidthFromSpPr,
  parseSpPrBorderColor,
  parseSpPrFill,
} from "./shape";
import {
  type ResolvedManualLayout,
  buildManualLayout,
  normalizeChartManualLayout,
  normalizeManualLayout,
  parseManualLayout,
} from "./layout";
import {
  childElements,
  findChild,
  parseBoolAttr,
  parseNumericChildVal,
  readBoolAttr,
} from "./util";
import {
  AXIS_ID_CAT,
  AXIS_ID_VAL,
  AXIS_ID_VAL_X,
  AXIS_ID_VAL_Y,
  type AxisRenderOptions,
  buildBarAxes,
  buildScatterAxes,
  normalizeAxisCrossBetween,
  normalizeAxisDispUnits,
  normalizeAxisGridlines,
  normalizeAxisHidden,
  normalizeAxisLabelBold,
  normalizeAxisLabelColor,
  normalizeAxisLabelFontFamily,
  normalizeAxisLabelFontSize,
  normalizeAxisLabelItalic,
  normalizeAxisLabelRotation,
  normalizeAxisLabelStrike,
  normalizeAxisLabelUnderline,
  normalizeAxisLblAlgn,
  normalizeAxisLblOffset,
  normalizeAxisNumberFormat,
  normalizeAxisScale,
  normalizeAxisSkip,
  normalizeAxisTitle,
  normalizeAxisTitleBold,
  normalizeAxisTitleColor,
  normalizeAxisTitleFontFamily,
  normalizeAxisTitleFontSize,
  normalizeAxisTitleItalic,
  normalizeAxisTitleRotation,
  normalizeAxisTitleStrike,
  normalizeAxisTitleUnderline,
  normalizeTickLblPos,
  normalizeTickMark,
  normalizeAxisCrosses,
  parseAxisInfo,
} from "./axis";
import { normalizeTitleColor } from "./title";
import { buildDataTable, resolveDataTable } from "./dataTable";
import { buildChartLevelDataLabels } from "./dataLabels";
import { buildSeries } from "./series";

// ── Plot-area-scope constants ─────────────────────────────────────

/**
 * Chart kinds that default `<c:varyColors>` to `1` in OOXML — every
 * data point in the (single) series carries a unique color. Excel's
 * pie / doughnut / ofPie templates emit `<c:varyColors val="1"/>` so
 * absence and `1` collapse to `undefined` here; only an explicit `0`
 * surfaces `false`.
 */
const VARY_COLORS_DEFAULT_TRUE: ReadonlySet<ChartKind> = new Set([
  "pie",
  "pie3D",
  "doughnut",
  "ofPie",
]);

/**
 * Recognized values of `<c:scatterStyle>` per the OOXML
 * `ST_ScatterStyle` enumeration. Tokens outside the set drop to
 * `undefined` so a corrupt template does not surface a string Excel
 * would not emit.
 */
const VALID_SCATTER_STYLES: ReadonlySet<ChartScatterStyle> = new Set([
  "none",
  "line",
  "lineMarker",
  "marker",
  "smooth",
  "smoothMarker",
]);

const DOUGHNUT_HOLE_DEFAULT = 50;
const DOUGHNUT_HOLE_MIN = 10;
const DOUGHNUT_HOLE_MAX = 90;

/**
 * Chart kinds that emit `<c:varyColors val="1"/>` by default at the
 * writer side. Mirrors the reader's {@link VARY_COLORS_DEFAULT_TRUE}
 * — kept under a distinct name because the two sides drive different
 * defaults (the writer cares about the writable-kind `WriteChartKind`
 * subset, the reader cares about every parsed `ChartKind`).
 */
const VARY_COLORS_DEFAULT_TRUE_TYPES: ReadonlySet<import("../../_types").WriteChartKind> = new Set([
  "pie",
  "doughnut",
]);

/**
 * Recognized values of `<c:scatterStyle>` per the OOXML
 * `ST_ScatterStyle` enumeration. Used to validate
 * `chart.scatterStyle` before it lands in the rendered XML.
 */
const SCATTER_STYLE_VALUES: ReadonlySet<ChartScatterStyle> = new Set([
  "none",
  "line",
  "lineMarker",
  "marker",
  "smooth",
  "smoothMarker",
]);

/**
 * Resolve `<c:plotArea><c:layout><c:manualLayout>...</c:manualLayout>
 * </c:layout></c:plotArea>` from {@link SheetChart.plotAreaLayout}.
 *
 * Returns the normalized coordinate set, or `undefined` when every axis
 * the caller pinned dropped to `undefined`. The caller emits the bare
 * `<c:layout/>` placeholder in that case so a fresh chart matches
 * Excel's reference shape byte-for-byte (Excel itself emits the empty
 * placeholder on every auto-layout chart — the element is the first
 * child of `<c:plotArea>` per `CT_PlotArea`, ECMA-376 Part 1,
 * §21.2.2.145).
 *
 * Coordinates outside the OOXML `0..1` band, `NaN`, `Infinity`, and
 * non-numeric inputs all collapse to `undefined` on the matching axis
 * so the writer drops the matching `<c:x>` / `<c:y>` / `<c:w>` /
 * `<c:h>` slot rather than emit a token Excel would reject — same
 * accept-or-drop grammar as {@link resolveLegendLayout}.
 */
function resolvePlotAreaLayout(chart: SheetChart): ResolvedManualLayout | undefined {
  return normalizeManualLayout(chart.plotAreaLayout);
}

// ── Reader ────────────────────────────────────────────────────────

/**
 * Pull per-axis metadata from the plot area's `<c:catAx>` / `<c:valAx>`
 * children.
 *
 * The mapping mirrors the writer side:
 *   - bar / column / line / area: `x` = `<c:catAx>`, `y` = first `<c:valAx>`.
 *   - scatter / bubble:           `x` = first `<c:valAx>`, `y` = second `<c:valAx>`.
 *
 * Returns `undefined` when neither axis surfaces a title — keeps the
 * default `Chart` shape lean.
 */
export function parseAxes(
  plotArea: XmlElement,
): { x?: ChartAxisInfo; y?: ChartAxisInfo } | undefined {
  let catAx: XmlElement | undefined;
  const valAxes: XmlElement[] = [];
  for (const child of childElements(plotArea)) {
    if (child.local === "catAx") {
      catAx ??= child;
    } else if (child.local === "valAx") {
      valAxes.push(child);
    }
  }

  let xAxis: XmlElement | undefined;
  let yAxis: XmlElement | undefined;
  if (catAx) {
    xAxis = catAx;
    yAxis = valAxes[0];
  } else {
    // Scatter / bubble: both axes are valAx. The first declared one is
    // the X axis (`axPos="b"`), the second is the Y axis (`axPos="l"`).
    xAxis = valAxes[0];
    yAxis = valAxes[1];
  }

  // `<c:crossBetween>` is required on every `<c:valAx>` — Excel always
  // emits the element — so the reader needs a per-family default to
  // collapse against. The catAx-anchored families (bar / column / line
  // / area) emit `"between"` on the value axis; scatter (catAx-less,
  // both axes are valAx) emits `"midCat"` on both axes. Pass the
  // expected default to `parseAxisInfo` so a chart that inherited the
  // default round-trips identically through {@link cloneChart} —
  // absence on the parsed shape and the writer-emitted default produce
  // the same `<c:crossBetween val=".."/>` byte-for-byte.
  const familyDefaultCrossBetween: ChartAxisCrossBetween = catAx ? "between" : "midCat";

  const x = xAxis ? parseAxisInfo(xAxis, familyDefaultCrossBetween) : undefined;
  const y = yAxis ? parseAxisInfo(yAxis, familyDefaultCrossBetween) : undefined;

  if (!x && !y) return undefined;
  const out: { x?: ChartAxisInfo; y?: ChartAxisInfo } = {};
  if (x) out.x = x;
  if (y) out.y = y;
  return out;
}

/**
 * Pull `<c:plotArea><c:layout><c:manualLayout>` off the chart. Reflects
 * Excel's "Format Plot Area -> Position -> Custom" placement — the
 * `(x, y)` anchor and `(w, h)` size of the plot area as fractions of
 * the chart frame in the `0..1` band.
 *
 * The OOXML schema (`CT_PlotArea`, ECMA-376 Part 1, §21.2.2.145) places
 * `<c:layout>` as the first child of `<c:plotArea>`, before any
 * chart-type element / axes / `<c:dTable>` / `<c:spPr>`. The
 * `<c:manualLayout>` block (`CT_ManualLayout`, §21.2.2.115) exposes
 * optional `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` children whose `val`
 * attributes carry an `xsd:double`. The reader admits the coordinate
 * only when `val` parses to a finite number in the `0..1` band;
 * out-of-range / non-finite / non-numeric tokens drop to `undefined`
 * on the matching axis so absence and a malformed token round-trip
 * identically through {@link cloneChart}.
 *
 * Both `<c:xMode val="edge"/>` (absolute fraction of the chart frame)
 * and `<c:xMode val="factor"/>` (delta from auto-layout) are accepted —
 * the reader surfaces the same `ChartManualLayout` shape regardless,
 * since the writer always normalizes to `"edge"` on emit (Excel itself
 * emits the absolute form when the user drags an element to a custom
 * position).
 *
 * Returns `undefined` whenever the plot area omits the `<c:layout>` /
 * `<c:manualLayout>` chain at any link (the bare `<c:layout/>`
 * placeholder Excel emits on auto-layout charts has no `<c:manualLayout>`
 * child, so it collapses cleanly here), or when every coordinate
 * dropped on normalization — the field is omitted entirely on a clean
 * parse so absence and an empty layout round-trip identically through
 * the writer.
 */
export function parsePlotAreaLayout(plotArea: XmlElement): ChartManualLayout | undefined {
  return parseManualLayout(plotArea);
}

/**
 * Pull `<c:plotArea><c:spPr><a:solidFill><a:srgbClr val=".."/></a:solidFill>
 * </c:spPr></c:plotArea>` off the plot area. Returns the plot-area solid
 * fill color as a 6-character uppercase hex string.
 *
 * The OOXML `<a:srgbClr>` element carries the literal sRGB color
 * (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32) inside the fill choice
 * of `<c:spPr>` (`CT_ShapeProperties`, §20.1.2.3.13). The `<c:spPr>` slot
 * sits at the tail of `<c:plotArea>` per `CT_PlotArea` (§21.2.2.145),
 * after every chart-type element / axes / `<c:dTable>`.
 *
 * The reader surfaces only the literal `<a:srgbClr>` form — absence,
 * non-solid fills (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>` /
 * `<a:blipFill>`), and theme-color references (`<a:schemeClr>`) all
 * collapse to `undefined` so a chart that pinned a fill the writer
 * cannot reproduce on emit drops the field rather than fabricate one
 * Excel would render differently. Malformed `val` tokens (wrong length,
 * non-hex characters, alpha-channel forms, non-string escapes) likewise
 * drop to `undefined`.
 *
 * Mirrors the writer-side {@link SheetChart.plotAreaFillColor} so a
 * parsed value slots straight into {@link cloneChart} without
 * conversion. The lookup is scoped to direct children of `<c:plotArea>`
 * so a stray `<c:spPr>` elsewhere (e.g. on a series, on an axis, on the
 * `<c:dTable>` block) cannot leak in.
 */
export function parsePlotAreaFillColor(plotArea: XmlElement): string | undefined {
  return parseSpPrFill(plotArea);
}

/**
 * Pull `<c:plotArea><c:spPr><a:ln><a:solidFill><a:srgbClr val=".."/>
 * </a:solidFill></a:ln></c:spPr></c:plotArea>` off the plot-area block.
 * Returns the line stroke color as a 6-character uppercase hex string.
 *
 * The OOXML `<a:srgbClr>` element carries the literal sRGB color
 * (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32) inside the line's
 * solid fill choice (`CT_LineProperties`, §20.1.2.3.24) which itself
 * sits inside `<c:spPr>` (`CT_ShapeProperties`, §20.1.2.3.13). The
 * `<c:spPr>` slot lives at the tail of `<c:plotArea>` per
 * `CT_PlotArea` (§21.2.2.145), after every chart-type element / axes /
 * `<c:dTable>`.
 *
 * The reader surfaces only the literal `<a:srgbClr>` form — absence,
 * non-solid line fills (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>`),
 * and theme-color references (`<a:schemeClr>`) all collapse to
 * `undefined` so a chart that pinned a stroke the writer cannot
 * reproduce on emit drops the field rather than fabricate one Excel
 * would render differently. Malformed `val` tokens (wrong length,
 * non-hex characters, alpha-channel forms, non-string escapes)
 * likewise drop to `undefined`.
 *
 * Mirrors the writer-side {@link SheetChart.plotAreaBorderColor} so a
 * parsed value slots straight into {@link cloneChart} without
 * conversion. The lookup is scoped to direct children of
 * `<c:plotArea>` so a stray `<c:spPr>` elsewhere (e.g. on a series,
 * on an axis, on the `<c:dTable>` block) cannot leak in. Mirrors
 * {@link parsePlotAreaFillColor} — same `<c:spPr>` host element on
 * the same `<c:plotArea>` parent — but lands on the line
 * (`<a:ln><a:solidFill>`) child rather than the fill
 * (`<a:solidFill>`) child.
 */
export function parsePlotAreaBorderColor(plotArea: XmlElement): string | undefined {
  return parseSpPrBorderColor(plotArea);
}

/**
 * Pull the `w` attribute off `<c:plotArea><c:spPr><a:ln w="EMU"/>` and
 * return the stroke width in points after clamping to the
 * `0.25..13.5` pt band Excel's UI exposes. The OOXML `w` attribute
 * carries the stroke width in English Metric Units (1 pt = 12 700 EMU)
 * per `CT_LineProperties` (ECMA-376 Part 1, §20.1.2.3.24); the reader
 * snaps the result to the 0.25 pt grid so a parsed-then-written width
 * does not drift across round-trips (Excel rounds in the UI anyway).
 *
 * Returns `undefined` when there is no `<c:spPr><a:ln w=..>` slot, when
 * the attribute is missing, when the value cannot be parsed as a finite
 * positive number, or when it parses to zero (Excel's "no border"
 * marker — the writer-side knob does not model that state). Mirrors
 * the writer-side {@link SheetChart.plotAreaBorderWidth} so a parsed
 * value slots straight into {@link cloneChart} without conversion.
 *
 * The lookup is scoped to direct children of `<c:plotArea>` so a stray
 * `<a:ln w=..>` elsewhere (on a series stroke, on an axis line) cannot
 * leak in. Mirrors {@link parsePlotAreaBorderColor} — same `<c:spPr>`
 * host on the same `<c:plotArea>` parent — but lands on the `w`
 * attribute rather than the `<a:solidFill><a:srgbClr>` color child.
 */
export function parsePlotAreaBorderWidth(plotArea: XmlElement): number | undefined {
  return parseBorderWidthFromSpPr(plotArea);
}

/**
 * Pull `<c:varyColors val=".."/>` off a chart-type element.
 *
 * Excel's per-family default flips the meaning: pie / doughnut /
 * pie3D / ofPie default to `true` (every slice unique) while every
 * other chart family defaults to `false` (one color per series).
 * Matching values collapse to `undefined` so a roundtrip of a stock
 * template stays minimal — only non-default values surface so
 * {@link cloneChart} can carry them through. Unknown values and
 * missing `val` attributes drop to `undefined`.
 */
export function parseVaryColors(chartTypeEl: XmlElement, kind: ChartKind): boolean | undefined {
  const el = findChild(chartTypeEl, "varyColors");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const familyDefaultsTrue = VARY_COLORS_DEFAULT_TRUE.has(kind);
  // Accept the OOXML truthy / falsy spellings. `1` / `true` map to true,
  // `0` / `false` map to false, anything else drops.
  let parsed: boolean;
  switch (raw) {
    case "1":
    case "true":
      parsed = true;
      break;
    case "0":
    case "false":
      parsed = false;
      break;
    default:
      return undefined;
  }
  // Collapse the per-family default so absence and the default
  // round-trip identically.
  if (parsed === familyDefaultsTrue) return undefined;
  return parsed;
}

/**
 * Pull `<c:scatterStyle val=".."/>` off a `<c:scatterChart>` element.
 *
 * The OOXML schema lists the element as required on `<c:scatterChart>`
 * but tolerates absence in practice — Excel falls back to `"marker"`
 * (the schema default per CT_ScatterStyle) when the file omits it.
 * The reader does not pin a default of its own: every literal value
 * in {@link VALID_SCATTER_STYLES} surfaces as-is so a clone preserves
 * the exact preset the template authored. Missing elements, missing
 * `val` attributes, and tokens outside the enum drop to `undefined`.
 *
 * Note that the writer's default is `"lineMarker"` (Excel's chart-
 * picker default and what every fresh hucre scatter chart emits today),
 * which differs from the OOXML schema default of `"marker"`. The
 * asymmetry is intentional — writing `"lineMarker"` matches Excel's
 * UI default; not collapsing it on read keeps the round-trip exact.
 */
export function parseScatterStyle(scatterChart: XmlElement): ChartScatterStyle | undefined {
  const el = findChild(scatterChart, "scatterStyle");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  if (!VALID_SCATTER_STYLES.has(raw as ChartScatterStyle)) return undefined;
  return raw as ChartScatterStyle;
}

/**
 * Pull `<c:dropLines/>` off a `<c:lineChart>` / `<c:line3DChart>` /
 * `<c:areaChart>` / `<c:area3DChart>` element. Returns `true` when
 * the element is present (its mere presence paints the connector
 * lines per OOXML CT_ChartLines), `undefined` otherwise so absence
 * collapses to the writer's default.
 *
 * `<c:dropLines>` is structurally a `CT_ChartLines` and may carry a
 * nested `<c:spPr>` for stroke styling, but hucre's reader only
 * surfaces the on/off bit — the shape properties are not modelled in
 * this phase. A template that pins custom drop-line colors / widths
 * therefore round-trips as a default-styled line; the on/off intent
 * still survives, which is what {@link cloneChart} needs.
 */
export function parseDropLines(chartTypeEl: XmlElement): boolean | undefined {
  return findChild(chartTypeEl, "dropLines") ? true : undefined;
}

/**
 * Pull `<c:hiLowLines/>` off a `<c:lineChart>` / `<c:line3DChart>` /
 * `<c:stockChart>` element. Same on/off shape as
 * {@link parseDropLines}; the element is bare so its mere presence
 * surfaces `true`, absence collapses to `undefined`.
 */
export function parseHiLowLines(chartTypeEl: XmlElement): boolean | undefined {
  return findChild(chartTypeEl, "hiLowLines") ? true : undefined;
}

/**
 * Pull `<c:serLines/>` off a `<c:barChart>` / `<c:ofPieChart>` element.
 * Same on/off shape as {@link parseDropLines} / {@link parseHiLowLines};
 * the element is bare so its mere presence surfaces `true`, absence
 * collapses to `undefined`.
 *
 * `<c:serLines>` is structurally a `CT_ChartLines` and may carry a
 * nested `<c:spPr>` for stroke styling, but hucre's reader only
 * surfaces the on/off bit at this layer (mirrors how `parseDropLines`
 * handles the same shape on its hosts). Even when the nested `<c:spPr>`
 * is the only child, the presence flag still survives, which is what
 * {@link cloneChart} needs.
 */
export function parseSerLines(chartTypeEl: XmlElement): boolean | undefined {
  return findChild(chartTypeEl, "serLines") ? true : undefined;
}

/**
 * Pull `<c:marker val=".."/>` off a `<c:lineChart>` element. This is
 * the chart-level CT_Boolean variant of `<c:marker>` — distinct from
 * the per-series `<c:marker>` (CT_Marker, with style / size / fill).
 * The element gates whether per-series markers paint at all on the
 * line chart.
 *
 * The OOXML / Excel default `val="1"` (markers shown) collapses to
 * `undefined` so absence and the default round-trip identically
 * through {@link cloneChart}; only an explicit `val="0"` surfaces
 * `false`. Accepts the OOXML truthy / falsy spellings (`"1"` /
 * `"true"` / `"0"` / `"false"`); unknown values, missing `val`
 * attributes, and a missing element all drop to `undefined`.
 *
 * The chart-level slot lives exclusively on `CT_LineChart` per the
 * OOXML schema — `CT_Line3DChart` and `CT_StockChart` have no
 * chart-level marker toggle. Caller is expected to gate the lookup
 * on the matching chart-type kind.
 */
export function parseShowLineMarkers(lineChart: XmlElement): boolean | undefined {
  const el = findChild(lineChart, "marker");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "0":
    case "false":
      return false;
    case "1":
    case "true":
      // OOXML / Excel default — collapse to undefined for symmetry
      // with the writer's `showLineMarkers` field, so a fresh chart
      // and a marker-on chart round-trip identically.
      return undefined;
    default:
      return undefined;
  }
}

/**
 * Pull `<c:grouping val=".."/>` off a `<c:barChart>` element. Returns
 * `undefined` when the grouping element is missing or carries the
 * default `"standard"` / `"clustered"` value — the writer's
 * {@link SheetChart.barGrouping} treats both as the unspecified
 * default, so omitting them keeps the parsed shape minimal.
 */
export function parseBarGrouping(barChart: XmlElement): ChartBarGrouping | undefined {
  const grouping = findChild(barChart, "grouping");
  if (!grouping) return undefined;
  const val = grouping.attrs.val;
  if (typeof val !== "string") return undefined;
  switch (val) {
    case "stacked":
      return "stacked";
    case "percentStacked":
      return "percentStacked";
    case "clustered":
      return "clustered";
    case "standard":
      // OOXML's `standard` for barChart is functionally equivalent to
      // `clustered` (Excel renders side-by-side). Surface neither so
      // the cloned chart inherits the writer's default.
      return undefined;
    default:
      return undefined;
  }
}

/**
 * Pull `<c:holeSize val=".."/>` off a `<c:doughnutChart>` element.
 * Returns `undefined` when the attribute is missing, malformed, or
 * outside the 1–99 range OOXML allows. Excel itself only writes values
 * in 10–90 (the UI clamps to that band) but the spec is wider, so we
 * accept the full schema range and let the writer re-clamp on the way
 * back out.
 */
export function parseHoleSize(doughnut: XmlElement): number | undefined {
  const el = findChild(doughnut, "holeSize");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < 1 || parsed > 99) return undefined;
  return parsed;
}

/**
 * Pull `<c:gapWidth val=".."/>` off a `<c:barChart>` / `<c:bar3DChart>`
 * element.
 *
 * The OOXML schema (`ST_GapAmount`) restricts the value to the
 * inclusive `0..500` band; out-of-range values are dropped rather than
 * clamped so a corrupt template does not silently rewrite as a
 * different gap. The OOXML default of `150` collapses to `undefined`
 * for symmetry with the writer's {@link SheetChart.gapWidth} default
 * — absence and `150` mean the same thing.
 */
export function parseGapWidth(barChart: XmlElement): number | undefined {
  const el = findChild(barChart, "gapWidth");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < 0 || parsed > 500) return undefined;
  if (parsed === 150) return undefined;
  return parsed;
}

/**
 * Pull `<c:gapWidth val=".."/>` off a `<c:upDownBars>` element. The
 * value controls the spacing between the up / down bars themselves on
 * a line chart (distinct from the bar-chart `<c:gapWidth>` which
 * controls spacing between category groups).
 *
 * The OOXML schema (`ST_GapAmount`) restricts the value to the
 * inclusive `0..500` band; out-of-range values are dropped rather than
 * clamped so a corrupt template does not silently rewrite as a
 * different gap. The OOXML default of `150` collapses to `undefined`
 * for symmetry with the writer's {@link SheetChart.upDownBarsGapWidth}
 * default — absence and `150` mean the same thing.
 */
export function parseUpDownBarsGapWidth(upDownBars: XmlElement): number | undefined {
  const el = findChild(upDownBars, "gapWidth");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < 0 || parsed > 500) return undefined;
  if (parsed === 150) return undefined;
  return parsed;
}

/**
 * Pull `<c:overlap val=".."/>` off a `<c:barChart>` / `<c:bar3DChart>`
 * element.
 *
 * The OOXML schema (`ST_Overlap`) restricts the value to the inclusive
 * `-100..100` band; out-of-range values are dropped rather than
 * clamped. The OOXML default of `0` collapses to `undefined` for
 * symmetry with the writer's {@link SheetChart.overlap} default. Note
 * that Excel's reference serialization emits `<c:overlap val="100"/>`
 * for stacked charts even though the schema default is `0`; we surface
 * the literal value carried by the file rather than try to invert
 * Excel's per-grouping default — `100` on a stacked chart therefore
 * round-trips as `100`.
 */
export function parseOverlap(barChart: XmlElement): number | undefined {
  const el = findChild(barChart, "overlap");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < -100 || parsed > 100) return undefined;
  if (parsed === 0) return undefined;
  return parsed;
}

/**
 * Pull `<c:firstSliceAng val=".."/>` off a `<c:pieChart>` /
 * `<c:doughnutChart>` element. Returns `undefined` when the attribute
 * is missing, malformed, or carries the OOXML default of `0` — the
 * writer's {@link SheetChart.firstSliceAng} treats absence and `0`
 * identically, so collapsing here keeps the round-trip stable.
 *
 * The OOXML schema (CT_FirstSliceAng) restricts the value to the
 * inclusive range `0..360`; out-of-range values are dropped rather
 * than clamped so a corrupt template does not silently rewrite as a
 * different angle.
 */
export function parseFirstSliceAng(chartType: XmlElement): number | undefined {
  const el = findChild(chartType, "firstSliceAng");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < 0 || parsed > 360) return undefined;
  // Collapse `0` and the schema-equivalent `360` to undefined — both
  // mean "first slice at 12 o'clock", which is the writer's default.
  if (parsed === 0 || parsed === 360) return undefined;
  return parsed;
}

/**
 * Pull `<c:grouping val=".."/>` off a `<c:lineChart>` or `<c:areaChart>`
 * element. Returns `undefined` when the grouping element is missing or
 * carries the default `"standard"` value — the writer's
 * {@link SheetChart.lineGrouping} / {@link SheetChart.areaGrouping}
 * treat that as the absence of the field.
 */
export function parseLineAreaGrouping(chartType: XmlElement): ChartLineAreaGrouping | undefined {
  const grouping = findChild(chartType, "grouping");
  if (!grouping) return undefined;
  const val = grouping.attrs.val;
  if (typeof val !== "string") return undefined;
  switch (val) {
    case "stacked":
      return "stacked";
    case "percentStacked":
      return "percentStacked";
    case "standard":
      return undefined;
    default:
      return undefined;
  }
}

// ── Writer ────────────────────────────────────────────────────────

export function buildPlotArea(chart: SheetChart, sheetName: string): string {
  // CT_PlotArea (ECMA-376 Part 1, §21.2.2.145) starts with `<c:layout>`
  // before any chart-type element / axes / `<c:dTable>` / `<c:spPr>`. The
  // writer always emits the element so the file's intent is explicit
  // even on roundtrip — Excel itself includes the (empty) auto-layout
  // placeholder in every reference serialization. When
  // `chart.plotAreaLayout` is pinned the placeholder upgrades to
  // `<c:layout><c:manualLayout>...</c:manualLayout></c:layout>` carrying
  // the caller's `(x, y, w, h)` coordinates per `CT_ManualLayout`
  // (§21.2.2.115). An empty layout (every coordinate dropped on
  // normalization) collapses back to the bare placeholder so a fresh
  // chart matches Excel's reference shape byte-for-byte.
  const plotAreaLayoutXml = buildManualLayout(resolvePlotAreaLayout(chart));
  const children: string[] = [plotAreaLayoutXml ?? xmlSelfClose("c:layout")];

  // Axis titles, gridlines, scaling, number format and tick rendering
  // surface for every chart family except pie/doughnut. Pull them once
  // so each branch can hand them off to the matching axis builder.
  const opts: AxisRenderOptions = {
    xAxisTitle: normalizeAxisTitle(chart.axes?.x?.title),
    yAxisTitle: normalizeAxisTitle(chart.axes?.y?.title),
    // `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title>`
    // sits on every axis flavour per the OOXML schema (CT_CatAx,
    // CT_ValAx, CT_DateAx, CT_SerAx all carry the same `<c:title>`
    // shape). Normalize the caller's degree input — clamp to the
    // `-90..90` band Excel's UI exposes; non-finite / non-numeric
    // inputs collapse to `undefined` so the writer emits the OOXML
    // default `rot="0"` byte-for-byte. The per-family axis builders
    // only honour the rotation when the axis actually renders a title.
    xAxisTitleRotation: normalizeAxisTitleRotation(chart.axes?.x?.axisTitleRotation),
    yAxisTitleRotation: normalizeAxisTitleRotation(chart.axes?.y?.axisTitleRotation),
    // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
    // <a:r><a:rPr sz="N"/></a:r></a:p></c:rich></c:tx></c:title>`
    // also sits on every axis flavour. Normalize the caller's point
    // input — drop out-of-range and non-finite / non-numeric inputs at
    // write time rather than emit a token Excel would reject; absence
    // collapses to `undefined` so the writer falls back to the
    // hardcoded 10pt default Excel itself emits on a fresh axis title.
    xAxisTitleFontSize: normalizeAxisTitleFontSize(chart.axes?.x?.axisTitleFontSize),
    yAxisTitleFontSize: normalizeAxisTitleFontSize(chart.axes?.y?.axisTitleFontSize),
    // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
    // <a:r><a:rPr b=".."/></a:r></a:p></c:rich></c:tx></c:title>` also
    // sits on every axis flavour. Normalize the caller's boolean
    // input — non-boolean tokens (typed escapes from an untyped
    // caller) collapse to `undefined` so the writer falls back to the
    // OOXML default `b="0"` (non-bold) Excel itself emits on a fresh
    // axis title. The per-family axis builders only honour the flag
    // when the axis actually renders a title.
    xAxisTitleBold: normalizeAxisTitleBold(chart.axes?.x?.axisTitleBold),
    yAxisTitleBold: normalizeAxisTitleBold(chart.axes?.y?.axisTitleBold),
    // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
    // <a:r><a:rPr i=".."/></a:r></a:p></c:rich></c:tx></c:title>` —
    // axis-title italic flag. The OOXML attribute is the `xsd:boolean`
    // `i` on `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7)
    // and the slot lives on every axis flavour. Normalize the caller's
    // boolean input — the writer keeps `true` / `false` literally so a
    // re-parse picks the value up off either canonical slot, while every
    // other token (typed escape from an untyped caller) collapses to
    // `undefined` and the writer omits the `i` attribute (Excel's
    // reference serialization for a non-italic axis title).
    xAxisTitleItalic: normalizeAxisTitleItalic(chart.axes?.x?.axisTitleItalic),
    yAxisTitleItalic: normalizeAxisTitleItalic(chart.axes?.y?.axisTitleItalic),
    // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
    // <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr>
    // <a:r><a:rPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
    // </a:rPr></a:r></a:p></c:rich></c:tx></c:title>` — axis-title
    // font color. The OOXML `<a:srgbClr val=".."/>` carries the
    // 6-character uppercase hex sRGB color (CT_SRgbColor inside
    // CT_TextCharacterProperties' fill choice — ECMA-376 Part 1,
    // §20.1.2.3.32 / §21.1.2.3.7) and the slot lives on every axis
    // flavour. Normalize the caller's hex input — the writer accepts
    // a leading `#` and any case, then collapses to the OOXML
    // canonical uppercase form. Malformed inputs (wrong length,
    // non-hex characters, alpha-channel forms, non-string escapes)
    // collapse to `undefined` and the writer omits the entire
    // `<a:solidFill>` block (Excel's reference serialization for an
    // axis title that inherits the theme text color).
    xAxisTitleColor: normalizeAxisTitleColor(chart.axes?.x?.axisTitleColor),
    yAxisTitleColor: normalizeAxisTitleColor(chart.axes?.y?.axisTitleColor),
    // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
    // <a:r><a:rPr strike=".."/></a:r></a:p></c:rich></c:tx></c:title>` —
    // axis-title strikethrough flag. The OOXML attribute is the
    // `ST_TextStrikeType` enum on `CT_TextCharacterProperties` (ECMA-376
    // Part 1, §21.1.2.3.7) and the slot lives on every axis flavour.
    // The writer emits only the UI variant `"sngStrike"`. Normalize the
    // caller's boolean input — `true` / `false` pass through literally,
    // every other token (typed escape from an untyped caller) collapses
    // to `undefined` and the writer omits the `strike` attribute (Excel's
    // reference serialization for a non-strikethrough axis title).
    xAxisTitleStrike: normalizeAxisTitleStrike(chart.axes?.x?.axisTitleStrike),
    yAxisTitleStrike: normalizeAxisTitleStrike(chart.axes?.y?.axisTitleStrike),
    // `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>`
    // also lives on every axis title `<c:rich>` body — same canonical
    // slot pair as the strike flag above. The writer emits only the UI
    // variant `"sng"`. Normalize the caller's boolean input — `true` /
    // `false` pass through literally, every other token (typed escape
    // from an untyped caller) collapses to `undefined` and the writer
    // omits the `u` attribute (Excel's reference serialization for a
    // non-underlined axis title).
    xAxisTitleUnderline: normalizeAxisTitleUnderline(chart.axes?.x?.axisTitleUnderline),
    yAxisTitleUnderline: normalizeAxisTitleUnderline(chart.axes?.y?.axisTitleUnderline),
    // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
    // typeface=".."/></a:defRPr></a:pPr><a:r><a:rPr><a:latin
    // typeface=".."/></a:rPr></a:r></a:p></c:rich></c:tx></c:title>` —
    // axis-title font family. The OOXML `<a:latin typeface=".."/>`
    // element carries the typeface name on `CT_TextFont` (ECMA-376
    // Part 1, §21.1.2.3.7) and the slot lives on every axis flavour.
    // Normalize the caller's string input — non-empty strings pass
    // through trimmed, every other token (empty / whitespace-only
    // strings, typed escapes from an untyped caller) collapses to
    // `undefined` and the writer skips the entire `<a:latin>` element
    // (Excel's reference serialization for an axis title that
    // inherits the theme typeface).
    xAxisTitleFontFamily: normalizeAxisTitleFontFamily(chart.axes?.x?.axisTitleFontFamily),
    yAxisTitleFontFamily: normalizeAxisTitleFontFamily(chart.axes?.y?.axisTitleFontFamily),
    // `<c:title><c:overlay val=".."/></c:title>` — axis-title overlay
    // flag. The element sits as a direct child of `<c:title>` per
    // CT_Title schema, and is always emitted by the writer (Excel's
    // reference serialization includes it on every visible axis
    // title) — only the `val` attribute flips when the caller pins
    // `axisTitleOverlay: true`. Anything other than literal `true`
    // collapses to `false` so a stray non-boolean leaking through the
    // type guard never produces `<c:overlay val="1"/>`.
    xAxisTitleOverlay: chart.axes?.x?.axisTitleOverlay === true,
    yAxisTitleOverlay: chart.axes?.y?.axisTitleOverlay === true,
    // `<c:title><c:layout><c:manualLayout>...</c:manualLayout></c:layout>
    // </c:title>` — axis-title manual placement. The OOXML
    // `CT_ManualLayout` block (ECMA-376 Part 1, §21.2.2.115) sits
    // inside `CT_Title` between `<c:tx>` and `<c:overlay>` and carries
    // the title's `(x, y)` anchor and `(w, h)` size as fractions of
    // the chart frame in the `0..1` band. Reuses the same
    // `normalizeManualLayout` helper as the chart-level legend /
    // plot-area layouts — out-of-range / non-finite / non-numeric
    // coordinates collapse to `undefined` axis-by-axis, and an empty
    // layout (every coordinate dropped) collapses to `undefined` so
    // the writer skips the entire `<c:layout>` block. Only meaningful
    // when the axis renders a title — the per-family axis builders
    // gate the value on the `xAxisTitle` / `yAxisTitle` field.
    xAxisTitleLayout: normalizeManualLayout(chart.axes?.x?.axisTitleLayout),
    yAxisTitleLayout: normalizeManualLayout(chart.axes?.y?.axisTitleLayout),
    // `<c:title><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
    // </a:solidFill></c:spPr></c:title>` — axis-title background fill.
    // The OOXML `<c:spPr>` block sits on `CT_Title` between
    // `<c:overlay>` and `<c:txPr>` / `<c:extLst>` (ECMA-376 Part 1,
    // §21.2.2.210). Mirrors the chart-level `titleFillColor` writer
    // path so a single hex string threads cleanly through both
    // call sites; reuses {@link normalizeTitleColor} so the
    // accept-with-or-without-`#` grammar matches the chart-title
    // fill / plot-area fill / legend fill resolvers exactly.
    // Malformed inputs (wrong length, non-hex characters,
    // alpha-channel forms, empty / whitespace-only strings,
    // non-string escapes from an untyped caller) collapse to
    // `undefined` and the writer omits the entire `<c:spPr>` block
    // (Excel's reference serialization for an axis title that
    // inherits the theme default fill — typically a transparent
    // title background with no `<c:spPr>` block).
    xAxisTitleFillColor: normalizeTitleColor(chart.axes?.x?.axisTitleFillColor),
    yAxisTitleFillColor: normalizeTitleColor(chart.axes?.y?.axisTitleFillColor),
    // `<c:title><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
    // </a:solidFill></a:ln></c:spPr></c:title>` — axis-title border
    // (line stroke) color. Same accept-with-or-without-`#` /
    // case-insensitive hex grammar as the chart-level
    // `titleBorderColor` knob. Malformed inputs (wrong length,
    // non-hex characters, alpha-channel forms, empty / whitespace-
    // only strings, non-string escapes from an untyped caller)
    // collapse to `undefined` and the writer omits the entire
    // `<a:ln>` block (Excel's reference serialization for an axis
    // title that inherits the auto-stroke — typically no visible
    // border).
    xAxisTitleBorderColor: normalizeTitleColor(chart.axes?.x?.axisTitleBorderColor),
    yAxisTitleBorderColor: normalizeTitleColor(chart.axes?.y?.axisTitleBorderColor),
    // `<c:title><c:spPr><a:ln w="EMU"/></c:spPr></c:title>` —
    // axis-title border (line stroke) thickness. Reuse the chart-level
    // {@link clampStrokeWidthPt} so the snap / clamp grammar matches
    // every other `<a:ln w=..>` slot the writer authors. Only
    // meaningful when the axis actually emits a title; the per-family
    // axis builder gates the value on the `xAxisTitle` / `yAxisTitle`
    // field.
    xAxisTitleBorderWidth: clampStrokeWidthPt(chart.axes?.x?.axisTitleBorderWidth),
    yAxisTitleBorderWidth: clampStrokeWidthPt(chart.axes?.y?.axisTitleBorderWidth),
    // `<c:title><c:spPr><a:ln><a:prstDash val=".."/></a:ln></c:spPr>
    // </c:title>` — axis-title border preset dash pattern. The
    // {@link normalizeBorderDash} helper drops `"solid"` and any
    // unrecognized value to `undefined` so a fresh axis title matches
    // Excel's reference shape byte-for-byte.
    xAxisTitleBorderDash: normalizeBorderDash(chart.axes?.x?.axisTitleBorderDash),
    yAxisTitleBorderDash: normalizeBorderDash(chart.axes?.y?.axisTitleBorderDash),
    xGridlines: normalizeAxisGridlines(chart.axes?.x?.gridlines),
    yGridlines: normalizeAxisGridlines(chart.axes?.y?.gridlines),
    xScale: normalizeAxisScale(chart.axes?.x?.scale),
    yScale: normalizeAxisScale(chart.axes?.y?.scale),
    xNumFmt: normalizeAxisNumberFormat(chart.axes?.x?.numberFormat),
    yNumFmt: normalizeAxisNumberFormat(chart.axes?.y?.numberFormat),
    xMajorTickMark: normalizeTickMark(chart.axes?.x?.majorTickMark),
    yMajorTickMark: normalizeTickMark(chart.axes?.y?.majorTickMark),
    xMinorTickMark: normalizeTickMark(chart.axes?.x?.minorTickMark),
    yMinorTickMark: normalizeTickMark(chart.axes?.y?.minorTickMark),
    xTickLblPos: normalizeTickLblPos(chart.axes?.x?.tickLblPos),
    yTickLblPos: normalizeTickLblPos(chart.axes?.y?.tickLblPos),
    // `<c:txPr><a:bodyPr rot="N"/></c:txPr>` lives on every axis
    // flavour per the OOXML schema (CT_CatAx, CT_ValAx, CT_DateAx,
    // CT_SerAx all carry the optional `<c:txPr>`). Normalize the
    // caller's degree input — clamp to the `-90..90` band Excel's UI
    // exposes; non-finite / non-numeric inputs and the OOXML default
    // `0` collapse to `undefined` so the writer can elide the entire
    // `<c:txPr>` block on a fresh chart.
    xLabelRotation: normalizeAxisLabelRotation(chart.axes?.x?.labelRotation),
    yLabelRotation: normalizeAxisLabelRotation(chart.axes?.y?.labelRotation),
    // `<c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr>`
    // shares the same `<c:txPr>` block as the rotation slot above. The
    // writer normalizes the points input — clamp to the `1..400`pt
    // band the OOXML `ST_TextFontSize` schema exposes; non-finite /
    // out-of-range / non-numeric inputs collapse to `undefined` so a
    // fresh chart inherits Excel's reference 10pt tick-label size.
    xLabelFontSize: normalizeAxisLabelFontSize(chart.axes?.x?.labelFontSize),
    yLabelFontSize: normalizeAxisLabelFontSize(chart.axes?.y?.labelFontSize),
    // `<c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr>`
    // shares the same `<c:txPr>` block as the rotation / size slots
    // above. `true` / `false` pass through literally; non-boolean
    // tokens (typed escapes from an untyped caller) collapse to
    // `undefined` so the writer omits the `b` attribute and a fresh
    // chart inherits the theme-default tick-label weight.
    xLabelBold: normalizeAxisLabelBold(chart.axes?.x?.labelBold),
    yLabelBold: normalizeAxisLabelBold(chart.axes?.y?.labelBold),
    // `<c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr>`
    // shares the same `<c:txPr>` block as the rotation / size / bold
    // slots above. `true` / `false` pass through literally; non-boolean
    // tokens (typed escapes from an untyped caller) collapse to
    // `undefined` so the writer omits the `i` attribute and a fresh
    // chart inherits the theme-default tick-label slant.
    xLabelItalic: normalizeAxisLabelItalic(chart.axes?.x?.labelItalic),
    yLabelItalic: normalizeAxisLabelItalic(chart.axes?.y?.labelItalic),
    // `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val=".."/>
    // </a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>` — tick-label
    // font color. Shares the same `<c:txPr>` block as the rotation /
    // size / bold / italic slots above. Normalize the caller's hex
    // input — the writer accepts a leading `#` and any case, then
    // collapses to the OOXML canonical uppercase form. Malformed
    // inputs (wrong length, non-hex characters, alpha-channel forms,
    // non-string escapes) collapse to `undefined` and the writer
    // omits the entire `<a:solidFill>` block (Excel's reference
    // serialization for tick labels that inherit the theme text color).
    xLabelColor: normalizeAxisLabelColor(chart.axes?.x?.labelColor),
    yLabelColor: normalizeAxisLabelColor(chart.axes?.y?.labelColor),
    // `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>`
    // shares the same `<c:txPr>` block as the rotation / size / bold
    // / italic / color slots above. The writer emits only the UI
    // variant `"sng"` when the input is `true`. `true` / `false` pass
    // through literally; non-boolean tokens (typed escapes from an
    // untyped caller) collapse to `undefined` so the writer omits the
    // `u` attribute and a fresh chart inherits Excel's reference
    // non-underlined tick labels.
    xLabelUnderline: normalizeAxisLabelUnderline(chart.axes?.x?.labelUnderline),
    yLabelUnderline: normalizeAxisLabelUnderline(chart.axes?.y?.labelUnderline),
    // `<c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr>`
    // shares the same `<c:txPr>` block as the rotation / size / bold
    // / italic / color / underline slots above. The writer emits only
    // the UI variant `"sngStrike"` when the input is `true`. `true` /
    // `false` pass through literally; non-boolean tokens (typed
    // escapes from an untyped caller) collapse to `undefined` so the
    // writer omits the `strike` attribute and a fresh chart inherits
    // Excel's reference non-strikethrough tick labels.
    xLabelStrike: normalizeAxisLabelStrike(chart.axes?.x?.labelStrike),
    yLabelStrike: normalizeAxisLabelStrike(chart.axes?.y?.labelStrike),
    // `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
    // </a:pPr></a:p></c:txPr>` — axis tick-label font family. The
    // element shares the same `<c:txPr>` block as the rotation / size
    // / bold / italic / color / underline / strike slots. The writer
    // trims surrounding whitespace and emits the trimmed typeface
    // verbatim. Empty / whitespace-only / non-string tokens collapse
    // to `undefined` so the writer skips the entire `<a:latin>`
    // element and a fresh chart inherits Excel's reference theme
    // typeface.
    xLabelFontFamily: normalizeAxisLabelFontFamily(chart.axes?.x?.labelFontFamily),
    yLabelFontFamily: normalizeAxisLabelFontFamily(chart.axes?.y?.labelFontFamily),
    xReverse: chart.axes?.x?.reverse === true,
    yReverse: chart.axes?.y?.reverse === true,
    // `tickLblSkip` / `tickMarkSkip` only round-trip on category axes
    // (`<c:catAx>` / `<c:dateAx>`). The scatter writer never emits
    // them — both axes are value axes — so the bar/column/line/area
    // catAx builder is the only consumer of these knobs.
    xTickLblSkip: normalizeAxisSkip(chart.axes?.x?.tickLblSkip),
    xTickMarkSkip: normalizeAxisSkip(chart.axes?.x?.tickMarkSkip),
    // `lblOffset` lives exclusively on `CT_CatAx` / `CT_DateAx` per
    // the OOXML schema. Same scope rule as the skip elements above —
    // scatter has no category axis, so the catAx builder is the only
    // consumer of this knob.
    xLblOffset: normalizeAxisLblOffset(chart.axes?.x?.lblOffset),
    // `lblAlgn` also lives exclusively on `CT_CatAx` / `CT_DateAx`
    // (`ST_LblAlgn`) — `<c:valAx>` and `<c:serAx>` reject it. Same
    // scope rule as `lblOffset`; the catAx builder is the sole
    // consumer.
    xLblAlgn: normalizeAxisLblAlgn(chart.axes?.x?.lblAlgn),
    // `noMultiLvlLbl` lives exclusively on `CT_CatAx` per ECMA-376
    // Part 1, §21.2.2 — even `<c:dateAx>` rejects the element. Same
    // catAx-only scope rule as the surrounding category-axis knobs;
    // the catAx builder is the sole consumer.
    xNoMultiLvlLbl: chart.axes?.x?.noMultiLvlLbl === true,
    // `<c:auto>` lives exclusively on `CT_CatAx` per ECMA-376 Part 1,
    // §21.2.2.7 — `<c:dateAx>`, `<c:valAx>`, and `<c:serAx>` reject the
    // element. Same catAx-only scope rule as `noMultiLvlLbl`. Only an
    // explicit `axes.x.auto === false` flips the toggle off; absence
    // (and any non-boolean) falls back to the OOXML default `true` so
    // the writer always emits Excel's reference `<c:auto val="1"/>`
    // shape on a stock chart.
    xAuto: chart.axes?.x?.auto !== false,
    // `<c:delete>` lives on every axis flavour (CT_CatAx / CT_ValAx /
    // CT_DateAx / CT_SerAx). The writer always emits the element —
    // Excel's reference serialization includes `<c:delete val="0"/>`
    // on every axis — so the axis builders read these flags directly
    // rather than skipping the element on `false`. Non-boolean inputs
    // collapse to `false` to keep the on-the-wire output stable.
    xHidden: normalizeAxisHidden(chart.axes?.x?.hidden),
    yHidden: normalizeAxisHidden(chart.axes?.y?.hidden),
    // `<c:crosses>` and `<c:crossesAt>` sit on every axis flavour
    // (CT_CatAx / CT_ValAx / CT_DateAx / CT_SerAx) but live in an XSD
    // choice — only one of them may appear at a time. The normalizer
    // resolves that choice once here so the per-family axis builders
    // can emit whichever element the caller pinned without duplicating
    // the precedence rule.
    xCrosses: normalizeAxisCrosses(chart.axes?.x?.crosses, chart.axes?.x?.crossesAt),
    yCrosses: normalizeAxisCrosses(chart.axes?.y?.crosses, chart.axes?.y?.crossesAt),
    // `<c:dispUnits>` lives exclusively on `<c:valAx>` per ECMA-376
    // §21.2.2.32 (CT_ValAx → CT_DispUnits). The category-axis builder
    // ignores `xDispUnits`; only the scatter X-axis (a value axis) and
    // every Y axis pick the field up. The normalizer collapses the
    // `ChartAxisDispUnit` shorthand to the full {@link ChartAxisDispUnits}
    // shape and rejects unknown tokens so the writer never emits a
    // `<c:builtInUnit>` value the OOXML `ST_BuiltInUnit` enum would
    // refuse.
    xDispUnits: normalizeAxisDispUnits(chart.axes?.x?.dispUnits),
    yDispUnits: normalizeAxisDispUnits(chart.axes?.y?.dispUnits),
    // `<c:crossBetween>` is value-axis-only per ECMA-376 §21.2.2.10
    // (CT_ValAx → CT_CrossBetween). The category-axis builder ignores
    // `xCrossBetween`; only the scatter X-axis (a value axis) and every
    // Y axis pick the field up. The normalizer rejects unknown tokens
    // so the writer never emits a value the OOXML `ST_CrossBetween`
    // enum would refuse — absence falls back to the per-family default
    // each axis builder pins today (`"between"` on bar / column / line
    // / area Y axes; `"midCat"` on both scatter axes).
    xCrossBetween: normalizeAxisCrossBetween(chart.axes?.x?.crossBetween),
    yCrossBetween: normalizeAxisCrossBetween(chart.axes?.y?.crossBetween),
  };

  switch (chart.type) {
    case "bar":
    case "column": {
      children.push(buildBarChart(chart, sheetName));
      children.push(...buildBarAxes(chart.type, opts));
      break;
    }
    case "line": {
      children.push(buildLineChart(chart, sheetName));
      children.push(...buildBarAxes("column", opts));
      break;
    }
    case "area": {
      children.push(buildAreaChart(chart, sheetName));
      children.push(...buildBarAxes("column", opts));
      break;
    }
    case "pie": {
      children.push(buildPieChart(chart, sheetName));
      break;
    }
    case "doughnut": {
      children.push(buildDoughnutChart(chart, sheetName));
      break;
    }
    case "scatter": {
      children.push(buildScatterChart(chart, sheetName));
      children.push(...buildScatterAxes(opts));
      break;
    }
    default: {
      // exhaustiveness guard
      const _exhaustive: never = chart.type;
      throw new Error(`Unsupported chart type: ${String(_exhaustive)}`);
    }
  }

  // `<c:dTable>` sits inside `<c:plotArea>` after the axes per
  // CT_PlotArea (ECMA-376 Part 1, §21.2.2.145) — between the last
  // `<c:valAx>` / `<c:catAx>` and the optional `<c:spPr>` that
  // `buildPlotAreaSpPr` below emits. Pie / doughnut have no axes at
  // all, so the OOXML schema places no slot for `<c:dTable>` on those
  // families; `resolveDataTable` short-circuits them by returning
  // `undefined`.
  const dTable = resolveDataTable(chart);
  if (dTable !== undefined) {
    children.push(buildDataTable(dTable));
  }

  // `<c:plotArea><c:spPr><a:solidFill><a:srgbClr val=".."/></a:solidFill>
  // </c:spPr></c:plotArea>` — Excel's "Format Plot Area -> Fill -> Solid
  // fill -> Color" pin. The slot sits at the tail of `<c:plotArea>` per
  // `CT_PlotArea` (ECMA-376 Part 1, §21.2.2.145), after every chart-type
  // element / axes / `<c:dTable>`. The writer emits the block only when
  // `chart.plotAreaFillColor` normalizes to a literal hex; absence and
  // every malformed token collapse to no `<c:spPr>` so a fresh chart
  // matches Excel's reference shape byte-for-byte.
  const plotAreaSpPr = buildPlotAreaSpPr(chart);
  if (plotAreaSpPr !== undefined) {
    children.push(plotAreaSpPr);
  }

  return xmlElement("c:plotArea", undefined, children);
}

/**
 * Build the optional `<c:spPr>` block at the tail of `<c:plotArea>`.
 * Surfaces the solid fill color knob
 * ({@link SheetChart.plotAreaFillColor}), the border (line) color
 * knob ({@link SheetChart.plotAreaBorderColor}) and the border width
 * knob ({@link SheetChart.plotAreaBorderWidth}) — every other `<c:spPr>`
 * child (`<a:effectLst>` effects, gradient / pattern / picture fills,
 * line dash / compound styles) is intentionally not modelled at this
 * layer.
 *
 * Returns `undefined` when every field is unset / malformed so the
 * writer skips the entire `<c:spPr>` block — an empty `<c:spPr/>`
 * collapses to the inherited theme fill / stroke Excel picks anyway,
 * and omitting it keeps untouched chart XML byte-clean. When at least
 * one knob lands on the wire, the children are emitted in
 * `CT_ShapeProperties` schema order: `<a:solidFill>` (fill) then
 * `<a:ln>` (line / stroke). The width knob lands on the `w` attribute
 * of `<a:ln>` (EMU; 1 pt = 12 700 EMU), authored together with the
 * border-color child so a stroke-only or color-only chart still emits a
 * single `<a:ln>` block.
 */
export function buildPlotAreaSpPr(chart: SheetChart): string | undefined {
  const fillHex = normalizePlotAreaFillColor(chart.plotAreaFillColor);
  const borderHex = normalizePlotAreaBorderColor(chart.plotAreaBorderColor);
  const borderWidthPt = clampStrokeWidthPt(chart.plotAreaBorderWidth);
  const borderDash = normalizeBorderDash(chart.plotAreaBorderDash);
  if (
    fillHex === undefined &&
    borderHex === undefined &&
    borderWidthPt === undefined &&
    borderDash === undefined
  ) {
    return undefined;
  }

  const children: string[] = [];
  if (fillHex !== undefined) {
    children.push(
      xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: fillHex })]),
    );
  }
  if (borderHex !== undefined || borderWidthPt !== undefined || borderDash !== undefined) {
    const lnAttrs: Record<string, string | number> = {};
    if (borderWidthPt !== undefined) {
      // OOXML stores stroke width in EMU (1 pt = 12 700 EMU). Round to
      // the nearest integer because the schema types `w` as `xsd:int`.
      lnAttrs.w = Math.round(borderWidthPt * EMU_PER_PT);
    }
    const lnChildren: string[] = [];
    if (borderHex !== undefined) {
      lnChildren.push(
        xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: borderHex })]),
      );
    }
    // `<a:prstDash>` follows `<a:solidFill>` per CT_LineProperties
    // (ECMA-376 Part 1, §20.1.2.3.24) — fill before dash before
    // headEnd / tailEnd. Skip emission for `"solid"` and unset values
    // so a fresh chart matches Excel's reference shape byte-for-byte.
    if (borderDash !== undefined) {
      lnChildren.push(xmlSelfClose("a:prstDash", { val: borderDash }));
    }
    children.push(
      lnChildren.length === 0
        ? xmlSelfClose("a:ln", lnAttrs)
        : xmlElement("a:ln", Object.keys(lnAttrs).length > 0 ? lnAttrs : undefined, lnChildren),
    );
  }
  return xmlElement("c:spPr", undefined, children);
}

/**
 * Normalize a {@link SheetChart.plotAreaFillColor} value for the
 * `<c:plotArea><c:spPr><a:solidFill><a:srgbClr val=".."/></a:solidFill>
 * </c:spPr></c:plotArea>` writer slot. Returns the 6-character uppercase
 * hex form when the input is a valid sRGB triple (with or without a
 * leading `#`), or `undefined` for any malformed token — wrong length,
 * non-hex characters, alpha-channel forms, or non-string escapes from an
 * untyped caller.
 *
 * Absence and malformed tokens both collapse to `undefined` so the
 * writer skips the entire `<c:spPr>` block and the plot area inherits
 * the auto-fill Excel picks from the chart's theme (Excel's reference
 * behavior for a fresh plot area without a custom color). Delegates to
 * the chart-level {@link normalizeTitleColor} so the two share the same
 * sRGB grammar.
 */
export function normalizePlotAreaFillColor(value: string | undefined): string | undefined {
  return normalizeTitleColor(value);
}

/**
 * Normalize a {@link SheetChart.plotAreaBorderColor} value for the
 * `<c:plotArea><c:spPr><a:ln><a:solidFill><a:srgbClr val=".."/>
 * </a:solidFill></a:ln></c:spPr></c:plotArea>` writer slot. Returns
 * the 6-character uppercase hex form when the input is a valid sRGB
 * triple (with or without a leading `#`), or `undefined` for any
 * malformed token — wrong length, non-hex characters, alpha-channel
 * forms, or non-string escapes from an untyped caller.
 *
 * Absence and malformed tokens both collapse to `undefined` so the
 * writer skips the `<a:ln>` block and the plot area inherits the
 * auto-stroke Excel picks from the chart's theme (Excel's reference
 * behavior for a fresh plot area without a custom border). Delegates
 * to the chart-level {@link normalizeTitleColor} so every `<a:srgbClr>`
 * fill / line slot shares the same sRGB grammar. Mirrors
 * {@link normalizePlotAreaFillColor} — same hex grammar, distinct
 * writer slot (`<a:ln>` rather than `<a:solidFill>`).
 */
export function normalizePlotAreaBorderColor(value: string | undefined): string | undefined {
  return normalizeTitleColor(value);
}

export function buildBarChart(chart: SheetChart, sheetName: string): string {
  const grouping = chart.barGrouping ?? "clustered";
  const barDir = chart.type === "bar" ? "bar" : "col";
  const isStacked = grouping === "percentStacked" || grouping === "stacked";

  const children: string[] = [
    xmlSelfClose("c:barDir", { val: barDir }),
    xmlSelfClose("c:grouping", { val: grouping }),
    xmlSelfClose("c:varyColors", { val: resolveVaryColors(chart) ? 1 : 0 }),
  ];

  for (let i = 0; i < chart.series.length; i++) {
    children.push(
      buildSeries(chart.series[i], i, sheetName, /* numericCategories */ false, {
        chartType: chart.type,
        dataLabels: chart.dataLabels,
        invertIfNegative: chart.series[i].invertIfNegative === true,
      }),
    );
  }

  const chartLevelDLbls = buildChartLevelDataLabels(chart);
  if (chartLevelDLbls) children.push(chartLevelDLbls);

  // OOXML CT_BarChart enforces a strict child order:
  // barDir → grouping → varyColors → ser* → dLbls? → gapWidth? →
  // overlap? → serLines* → axId+. `gapWidth` therefore lands before
  // `overlap` regardless of the chosen grouping.
  //
  // The defaults preserve Excel's reference serialization:
  //   - clustered                  → emit gapWidth=150, omit overlap
  //   - stacked / percentStacked   → emit overlap=100, omit gapWidth
  // An explicit `chart.gapWidth` / `chart.overlap` always emits the
  // matching element (even when the value happens to equal the default
  // for that grouping), so callers can pin both knobs on a stacked
  // chart or relax overlap on a clustered one.
  const explicitGapWidth = clampGapWidth(chart.gapWidth);
  const explicitOverlap = clampOverlap(chart.overlap);

  const emitGapWidth = explicitGapWidth ?? (isStacked ? undefined : 150);
  if (emitGapWidth !== undefined) {
    children.push(xmlSelfClose("c:gapWidth", { val: emitGapWidth }));
  }

  const emitOverlap = explicitOverlap ?? (isStacked ? 100 : undefined);
  if (emitOverlap !== undefined) {
    children.push(xmlSelfClose("c:overlap", { val: emitOverlap }));
  }

  // CT_BarChart sequence places `<c:serLines>` between `<c:overlap>`
  // and `<c:axId>`. The element is bare — its mere presence paints the
  // connectors between paired data points across consecutive series on
  // a stacked bar / column chart — so we only emit when the caller
  // explicitly opted in. Absence and an explicit `false` both collapse
  // to no element so untouched bar charts match Excel's reference
  // serialization. Excel only renders the connectors on stacked /
  // percentStacked groupings, but the writer still honours the toggle
  // on a clustered chart (matches Excel's own behavior — the element
  // pins, the renderer paints nothing).
  if (chart.serLines === true) {
    children.push(xmlElement("c:serLines", undefined, []));
  }

  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_CAT }));
  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_VAL }));

  return xmlElement("c:barChart", undefined, children);
}

/**
 * Normalize {@link SheetChart.gapWidth} to an integer in the inclusive
 * `0..500` band the OOXML schema (`ST_GapAmount`) allows.
 *
 * Returns `undefined` when the input is missing or non-finite so the
 * caller can fall through to the per-grouping default. Non-integer
 * values round to the nearest integer; out-of-range values clamp to
 * the schema bounds rather than wrap — `gapWidth` is a percentage of
 * the bar width with no natural wrap-around (a `600` group spacing is
 * not the same as `100`).
 */
export function clampGapWidth(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  if (rounded < 0) return 0;
  if (rounded > 500) return 500;
  return rounded;
}

/**
 * Normalize {@link SheetChart.overlap} to an integer in the inclusive
 * `-100..100` band the OOXML schema (`ST_Overlap`) allows.
 *
 * Returns `undefined` when the input is missing or non-finite so the
 * caller can fall through to the per-grouping default. Non-integer
 * values round to the nearest integer; out-of-range values clamp to
 * the schema bounds (`-100` and `100` are the geometric extremes —
 * series fully separated and series fully overlapped — wrapping makes
 * no physical sense).
 */
export function clampOverlap(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  if (rounded < -100) return -100;
  if (rounded > 100) return 100;
  return rounded;
}

export function buildLineChart(chart: SheetChart, sheetName: string): string {
  const grouping = chart.lineGrouping ?? "standard";
  const children: string[] = [
    xmlSelfClose("c:grouping", { val: grouping }),
    xmlSelfClose("c:varyColors", { val: resolveVaryColors(chart) ? 1 : 0 }),
  ];

  for (let i = 0; i < chart.series.length; i++) {
    // `<c:smooth>` is required on `CT_LineSer` per the OOXML schema, so
    // the line writer always emits the element — straight by default
    // (`val="0"`), curved when the caller pinned `smooth: true`.
    const seriesXml = buildSeries(chart.series[i], i, sheetName, /* numericCategories */ false, {
      chartType: chart.type,
      smooth: chart.series[i].smooth === true,
      dataLabels: chart.dataLabels,
      stroke: chart.series[i].stroke,
      marker: chart.series[i].marker,
    });
    children.push(seriesXml);
  }

  const chartLevelDLbls = buildChartLevelDataLabels(chart);
  if (chartLevelDLbls) children.push(chartLevelDLbls);

  // CT_LineChart child order: grouping, varyColors?, ser*, dLbls?,
  // dropLines?, hiLowLines?, upDownBars?, marker?, axId+. The
  // dropLines / hiLowLines / upDownBars blocks sit before `<c:marker>`
  // so the schema sequence is respected even on a chart that pins all
  // three flags. Each element is bare (or, for upDownBars, presence-
  // gated), so we only emit when the caller explicitly opted in
  // (`true`). Absence and an explicit `false` both collapse to no
  // element so untouched line charts match Excel's reference
  // serialization.
  if (chart.dropLines === true) {
    children.push(xmlElement("c:dropLines", undefined, []));
  }
  if (chart.hiLowLines === true) {
    children.push(xmlElement("c:hiLowLines", undefined, []));
  }
  if (chart.upDownBars === true) {
    children.push(buildUpDownBars(chart.upDownBarsGapWidth));
  }

  // `<c:marker>` (the chart-level CT_Boolean variant) gates per-series
  // marker rendering across the entire line chart. Excel's reference
  // serialization always emits the element on every authored line chart
  // — `val="1"` for the default "Line with Markers" look, `val="0"`
  // for the bare "Line" preset. The writer mirrors that always-emit
  // contract so a roundtrip preserves Excel's reference shape; only an
  // explicit `showLineMarkers: false` flips the value to `0` to suppress
  // the per-point dots chart-wide. `undefined` and `true` both emit
  // `val="1"` so a fresh chart matches Excel's default render and a
  // back-compat caller that never set the flag keeps the same output.
  children.push(xmlSelfClose("c:marker", { val: chart.showLineMarkers === false ? 0 : 1 }));
  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_CAT }));
  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_VAL }));

  return xmlElement("c:lineChart", undefined, children);
}

/**
 * Build a `<c:upDownBars>` block for {@link buildLineChart}.
 *
 * The OOXML schema (`CT_UpDownBars`) allows three optional children —
 * `<c:gapWidth>`, `<c:upBars>`, and `<c:downBars>` — but the up / down
 * bars themselves are painted by the mere presence of the parent
 * element. The writer emits a `<c:gapWidth val="N"/>` child to mirror
 * Excel's reference serialization for a freshly-toggled "Add Chart
 * Element -> Up/Down Bars" — `150` is the OOXML default for
 * `CT_UpDownBars/gapWidth` and the value Excel itself emits, so the
 * writer falls back to it when the caller leaves
 * {@link SheetChart.upDownBarsGapWidth} unset or pins an out-of-range
 * value. An explicit value in the inclusive `0..500` band is rounded
 * to the nearest integer and emitted literally.
 *
 * `<c:upBars>` / `<c:downBars>` are intentionally omitted: each is a
 * `CT_UpDownBar` (only `<c:spPr>` inside) and their absence makes
 * Excel paint the default white-up / black-down bars Excel uses on a
 * fresh toggle. A richer model — per-bar styling — can layer on top
 * in a follow-up if needed.
 */
export function buildUpDownBars(gapWidth: number | undefined): string {
  const resolved = clampUpDownBarsGapWidth(gapWidth) ?? 150;
  return xmlElement("c:upDownBars", undefined, [xmlSelfClose("c:gapWidth", { val: resolved })]);
}

/**
 * Normalize {@link SheetChart.upDownBarsGapWidth} to an integer in the
 * inclusive `0..500` band the OOXML schema (`ST_GapAmount`) allows.
 *
 * Returns `undefined` when the input is missing or non-finite so the
 * caller can fall through to the OOXML default `150`. Non-integer
 * values round to the nearest integer; out-of-range values drop to
 * `undefined` rather than clamp — a templated chart whose gap width
 * fell outside the schema bounds is treated as a fresh chart and
 * collapses to the default. Mirrors {@link clampGapWidth} but uses a
 * stricter "drop on out-of-range" policy because the up/down-bars gap
 * width has no per-grouping default to fall through to (every line
 * chart with the parent toggle on emits the same `150` default), so
 * silently rewriting an `800` to `500` would mislead the caller about
 * what Excel ends up rendering.
 */
export function clampUpDownBarsGapWidth(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  if (rounded < 0 || rounded > 500) return undefined;
  return rounded;
}

export function buildAreaChart(chart: SheetChart, sheetName: string): string {
  const grouping = chart.areaGrouping ?? "standard";
  const children: string[] = [
    xmlSelfClose("c:grouping", { val: grouping }),
    xmlSelfClose("c:varyColors", { val: resolveVaryColors(chart) ? 1 : 0 }),
  ];

  for (let i = 0; i < chart.series.length; i++) {
    children.push(
      buildSeries(chart.series[i], i, sheetName, /* numericCategories */ false, {
        chartType: chart.type,
        dataLabels: chart.dataLabels,
      }),
    );
  }

  const chartLevelDLbls = buildChartLevelDataLabels(chart);
  if (chartLevelDLbls) children.push(chartLevelDLbls);

  // CT_AreaChart sequence places `<c:dropLines>` between `<c:dLbls>`
  // and `<c:axId>`. The element is bare — its mere presence paints
  // the connectors — so we only emit when the caller explicitly opted
  // in. `<c:hiLowLines>` has no slot on `<c:areaChart>` per the OOXML
  // schema, so the area writer ignores `chart.hiLowLines` entirely.
  if (chart.dropLines === true) {
    children.push(xmlElement("c:dropLines", undefined, []));
  }

  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_CAT }));
  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_VAL }));

  return xmlElement("c:areaChart", undefined, children);
}

export function buildPieChart(chart: SheetChart, sheetName: string): string {
  const children: string[] = [
    xmlSelfClose("c:varyColors", { val: resolveVaryColors(chart) ? 1 : 0 }),
  ];

  // A pie chart only paints the first series; additional ones are
  // valid OOXML but Excel ignores them.
  if (chart.series.length > 0) {
    children.push(
      buildSeries(chart.series[0], 0, sheetName, /* numericCategories */ false, {
        chartType: chart.type,
        dataLabels: chart.dataLabels,
        explosion: chart.series[0].explosion,
      }),
    );
  }

  const chartLevelDLbls = buildChartLevelDataLabels(chart);
  if (chartLevelDLbls) children.push(chartLevelDLbls);

  // `<c:firstSliceAng>` is optional on `<c:pieChart>` (CT_PieChart);
  // omit it when the angle is the default `0` (12 o'clock start) so
  // we do not bloat untouched chart XML.
  const sliceAng = clampFirstSliceAng(chart.firstSliceAng);
  if (sliceAng !== undefined) {
    children.push(xmlSelfClose("c:firstSliceAng", { val: sliceAng }));
  }

  return xmlElement("c:pieChart", undefined, children);
}

export function buildDoughnutChart(chart: SheetChart, sheetName: string): string {
  const children: string[] = [
    xmlSelfClose("c:varyColors", { val: resolveVaryColors(chart) ? 1 : 0 }),
  ];

  // Like pie, doughnut paints every declared series — Excel renders
  // each as a concentric ring (rare in practice; most templates have
  // one). Carry every series through so multi-ring templates round-trip.
  for (let i = 0; i < chart.series.length; i++) {
    children.push(
      buildSeries(chart.series[i], i, sheetName, /* numericCategories */ false, {
        chartType: chart.type,
        dataLabels: chart.dataLabels,
        explosion: chart.series[i].explosion,
      }),
    );
  }

  const chartLevelDLbls = buildChartLevelDataLabels(chart);
  if (chartLevelDLbls) children.push(chartLevelDLbls);

  // `<c:firstSliceAng>` and `<c:holeSize>` are the two doughnut-only
  // knobs. firstSliceAng defaults to 0 (12 o'clock start); holeSize is
  // required by OOXML — the schema rejects a `<c:doughnutChart>` without
  // it. Clamp to the 10–90 band Excel's UI enforces; values outside
  // this range render but trigger Excel's repair dialog.
  //
  // The doughnut writer always emits `<c:firstSliceAng>`, falling back
  // to the default `0` when the caller did not request a rotation —
  // that mirrors the spec's reference serialization Excel produces.
  children.push(
    xmlSelfClose("c:firstSliceAng", { val: clampFirstSliceAng(chart.firstSliceAng) ?? 0 }),
  );
  children.push(xmlSelfClose("c:holeSize", { val: clampHoleSize(chart.holeSize) }));

  return xmlElement("c:doughnutChart", undefined, children);
}

/**
 * Normalize {@link SheetChart.firstSliceAng} to an integer in the
 * inclusive 0..360 band the OOXML schema (CT_FirstSliceAng) allows.
 *
 * Returns `undefined` for the default `0` so the pie writer can elide
 * the element entirely (Excel treats absence and `0` identically). The
 * doughnut writer must always emit the element, so it explicitly
 * substitutes `0` when the helper returns `undefined`.
 *
 * Out-of-range values are wrapped modulo 360 — `380` becomes `20`,
 * `-90` becomes `270` — which matches how Excel itself renders an
 * out-of-band value the user types into the chart-formatting pane.
 */
export function clampFirstSliceAng(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  // Wrap into 0..360 (inclusive). The OOXML schema actually allows
  // 360 as a value, so we keep it distinct from 0.
  let normalized = rounded % 360;
  if (normalized < 0) normalized += 360;
  if (normalized === 0) return undefined;
  return normalized;
}

export function clampHoleSize(value: number | undefined): number {
  if (value === undefined || !Number.isFinite(value)) return DOUGHNUT_HOLE_DEFAULT;
  const rounded = Math.round(value);
  if (rounded < DOUGHNUT_HOLE_MIN) return DOUGHNUT_HOLE_MIN;
  if (rounded > DOUGHNUT_HOLE_MAX) return DOUGHNUT_HOLE_MAX;
  return rounded;
}

export function buildScatterChart(chart: SheetChart, sheetName: string): string {
  const children: string[] = [
    xmlSelfClose("c:scatterStyle", { val: resolveScatterStyle(chart) }),
    xmlSelfClose("c:varyColors", { val: resolveVaryColors(chart) ? 1 : 0 }),
  ];

  for (let i = 0; i < chart.series.length; i++) {
    // `<c:smooth>` is optional on `CT_ScatterSer`; emit only when the
    // caller pinned `smooth: true`, falling back to the omit-by-default
    // shape Excel writes for straight scatter series.
    children.push(
      buildSeries(chart.series[i], i, sheetName, /* numericCategories */ true, {
        chartType: chart.type,
        smooth: chart.series[i].smooth === true ? true : undefined,
        dataLabels: chart.dataLabels,
        stroke: chart.series[i].stroke,
        marker: chart.series[i].marker,
      }),
    );
  }

  const chartLevelDLbls = buildChartLevelDataLabels(chart);
  if (chartLevelDLbls) children.push(chartLevelDLbls);

  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_VAL_X }));
  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_VAL_Y }));

  return xmlElement("c:scatterChart", undefined, children);
}

/**
 * Resolve the `<c:varyColors>` value emitted on the chart-type element.
 *
 * Falls back to the per-family default when the chart does not pin the
 * field, matching Excel's reference serialization (`true` for pie /
 * doughnut, `false` everywhere else). An explicit `chart.varyColors`
 * always wins, so a pie chart can collapse to a single color and a
 * column chart can paint each bar a different color.
 *
 * The writer always emits the element — the OOXML schema lists it as
 * required on every chart-type element except `surface` / `surface3D` /
 * `stock`, none of which hucre's writer authors. Emitting the explicit
 * value (matching Excel's reference output) keeps the rendered intent
 * unambiguous on roundtrip.
 */
export function resolveVaryColors(chart: SheetChart): boolean {
  if (typeof chart.varyColors === "boolean") return chart.varyColors;
  return VARY_COLORS_DEFAULT_TRUE_TYPES.has(chart.type);
}

/**
 * Resolve the `<c:scatterStyle>` value emitted on `<c:scatterChart>`.
 *
 * Defaults to `"lineMarker"` — Excel's chart-picker default and the
 * shape every existing scatter chart hucre writes uses. An explicit
 * `chart.scatterStyle` always wins; values outside the OOXML enum drop
 * back to the default rather than emit a token Excel would reject.
 *
 * The element is always emitted on `<c:scatterChart>` because the
 * OOXML schema lists it as required there — omitting it would produce
 * an invalid chart document Excel refuses to open.
 */
export function resolveScatterStyle(chart: SheetChart): ChartScatterStyle {
  const raw = chart.scatterStyle;
  if (raw && SCATTER_STYLE_VALUES.has(raw)) return raw;
  return "lineMarker";
}

// ── Clone-side plot-area constants ────────────────────────────────

const PLOT_AREA_BORDER_WIDTH_MIN_PT = 0.25;
const PLOT_AREA_BORDER_WIDTH_MAX_PT = 13.5;

// ── Clone resolvers (3-arg source/override) ───────────────────────

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
export function resolveCloneVaryColors(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve an `upDownBars` override.
 *
 * `undefined` → inherit the source's parsed `upDownBars`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML default — no `<c:upDownBars>` element emitted).
 * `boolean`   → replace.
 *
 * The grammar mirrors `roundedCorners` / `plotVisOnly` so the chart-
 * level line-only toggle composes the same way at the call site.
 * `false` collapses to absence on the writer side because the writer
 * only emits `<c:upDownBars>` when the flag is literally `true`; the
 * `false` value still surfaces in the cloned `SheetChart` for
 * symmetry with other resolve helpers, leaving the renderer to drop
 * it during emit.
 */
export function resolveCloneUpDownBars(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve an `upDownBarsGapWidth` override.
 *
 * `undefined` → inherit the source's parsed `upDownBarsGapWidth`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML default `150` Excel itself emits on a fresh
 *               toggle).
 * `number`    → replace. Out-of-range or non-finite values still
 *               surface in the cloned `SheetChart` for symmetry with
 *               the other override helpers; the writer's
 *               `clampUpDownBarsGapWidth` then drops them at emit
 *               time so a fresh chart matches Excel's reference
 *               serialization.
 *
 * The grammar mirrors `gapWidth` / `holeSize` / `firstSliceAng` so the
 * numeric chart-level knobs compose the same way at the call site.
 */
export function resolveCloneUpDownBarsGapWidth(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `plotAreaLayout` override.
 *
 * `undefined` → inherit the source's parsed `plotAreaLayout` (after
 *               running it through {@link normalizeLegendLayout} so a
 *               malformed source value drops cleanly — the normalizer
 *               is purely shape-based, no host-element awareness, so it
 *               applies identically to legend / plot-area layouts).
 * `null`      → drop the inherited layout (the writer falls back to the
 *               bare `<c:layout/>` placeholder Excel itself emits on
 *               every auto-layout chart).
 * `ChartManualLayout` → replace, after running through
 *               {@link normalizeLegendLayout}. Coordinates outside the
 *               `0..1` band collapse on the matching axis so the
 *               cloned `SheetChart` always carries a value the writer
 *               will accept; an override whose every axis dropped
 *               collapses to `undefined` so the writer skips the
 *               `<c:manualLayout>` body.
 *
 * The grammar mirrors `resolveCloneLegendLayout` so the manual-layout knobs
 * compose the same way at the call site. Unlike the legend variant, the
 * caller does not need to gate the result on any visibility flag —
 * every chart has a `<c:plotArea>` element to host `<c:layout>`.
 */
export function resolveClonePlotAreaLayout(
  sourceValue: ChartManualLayout | undefined,
  override: ChartManualLayout | null | undefined,
): ChartManualLayout | undefined {
  if (override === undefined) return normalizeChartManualLayout(sourceValue);
  if (override === null) return undefined;
  return normalizeChartManualLayout(override);
}

/**
 * Resolve a `plotAreaFillColor` override.
 *
 * `undefined` → inherit the source's parsed `plotAreaFillColor` (after
 *               running it through {@link normalizePlotAreaFillColor}
 *               so a malformed source value drops cleanly).
 * `null`      → drop the inherited fill (the writer emits no `<c:spPr>`
 *               block, the plot area inherits the auto-fill Excel
 *               picks from the chart's theme).
 * `string`    → replace with the normalized 6-character uppercase hex
 *               form. Malformed overrides collapse to `undefined` via
 *               the normalizer so the cloned `SheetChart` always
 *               carries a value the writer will accept.
 *
 * The grammar mirrors `titleColor` / `axes.x.axisTitleColor` /
 * `axes.x.labelColor` so the chart `<a:srgbClr>` knobs compose the
 * same way at the call site. Unlike those text-color knobs, the
 * plot-area fill is never gated on a visibility flag — every chart has
 * a `<c:plotArea>` element to host the fill.
 */
export function resolveClonePlotAreaFillColor(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizePlotAreaFillColor(sourceValue);
  if (override === null) return undefined;
  return normalizePlotAreaFillColor(override);
}

/**
 * Resolve a `plotAreaBorderColor` override.
 *
 * `undefined` → inherit the source's parsed `plotAreaBorderColor`
 *               (after running it through
 *               {@link normalizePlotAreaBorderColor} so a malformed
 *               source value drops cleanly).
 * `null`      → drop the inherited stroke (the writer emits no
 *               `<a:ln>` block on `<c:plotArea><c:spPr>`, the plot
 *               area inherits the auto-stroke Excel picks from the
 *               chart's theme).
 * `string`    → replace with the normalized 6-character uppercase hex
 *               form. Malformed overrides collapse to `undefined` via
 *               the normalizer so the cloned `SheetChart` always
 *               carries a value the writer will accept.
 *
 * The grammar mirrors `plotAreaFillColor` so the chart `<c:spPr>`
 * knobs compose the same way at the call site. Like the fill knob,
 * the border is never gated on a visibility flag — every chart has a
 * `<c:plotArea>` element to host the stroke.
 */
export function resolveClonePlotAreaBorderColor(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizePlotAreaBorderColor(sourceValue);
  if (override === null) return undefined;
  return normalizePlotAreaBorderColor(override);
}

/**
 * Normalize a `plotAreaBorderWidth` value for the cloned `SheetChart`.
 * Mirrors the writer's `clampStrokeWidthPt` — values are clamped to the
 * `0.25..13.5` pt band Excel's UI exposes and snapped to the 0.25 pt
 * grid so a parsed-then-cloned-then-written width does not drift across
 * round-trips (Excel rounds in the UI anyway). Non-finite / non-numeric
 * tokens (`NaN`, `Infinity`, strings, `null` from an untyped caller)
 * collapse to `undefined` so the cloned chart drops the field rather
 * than carry a value the writer would silently elide back to absence.
 */
export function normalizeClonePlotAreaBorderWidth(value: number | undefined): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  // Snap to the 0.25 pt grid Excel's UI exposes (Math.round(x * 4) / 4).
  const snapped = Math.round(value * 4) / 4;
  if (snapped < PLOT_AREA_BORDER_WIDTH_MIN_PT) return PLOT_AREA_BORDER_WIDTH_MIN_PT;
  if (snapped > PLOT_AREA_BORDER_WIDTH_MAX_PT) return PLOT_AREA_BORDER_WIDTH_MAX_PT;
  return snapped;
}

/**
 * Resolve a `plotAreaBorderWidth` override.
 *
 * `undefined` → inherit the source's parsed `plotAreaBorderWidth`
 *               (after running it through
 *               {@link normalizeClonePlotAreaBorderWidth} so a malformed
 *               source value drops cleanly).
 * `null`      → drop the inherited width (the writer emits `<a:ln>`
 *               without a `w` attribute, the line keeps Excel's
 *               auto-thickness).
 * `number`    → replace with the clamped / snapped point value.
 *               Non-finite / non-numeric overrides collapse to
 *               `undefined` via the normalizer so the cloned
 *               `SheetChart` always carries a value the writer will
 *               accept.
 *
 * The grammar mirrors the series-line stroke width so the chart
 * `<a:ln w=..>` knobs compose the same way at the call site. Like the
 * border-color knob, the width is never gated on a visibility flag —
 * every chart has a `<c:plotArea>` element to host the stroke.
 */
export function resolveClonePlotAreaBorderWidth(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return normalizeClonePlotAreaBorderWidth(sourceValue);
  if (override === null) return undefined;
  return normalizeClonePlotAreaBorderWidth(override);
}

/**
 * Resolve a `dropLines` override.
 *
 * `undefined` → inherit the source's parsed `dropLines`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML default — no `<c:dropLines>` element).
 * `boolean`   → replace. Only `true` round-trips into the cloned
 *               `SheetChart`; `false` collapses to `undefined` because
 *               the writer treats absence and `false` identically (no
 *               element emitted).
 *
 * The grammar mirrors `plotVisOnly` / `roundedCorners` so the chart-
 * level toggles compose the same way at the call site. Callers should
 * gate the result on the resolved chart family — `<c:dropLines>` has
 * no slot on `<c:barChart>` / `<c:pieChart>` / `<c:scatterChart>`.
 */
export function resolveCloneDropLines(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return sourceValue === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve a `hiLowLines` override. Mirrors {@link resolveCloneDropLines};
 * the only difference is the per-family scope — `<c:hiLowLines>` has
 * no slot on `<c:areaChart>`.
 */
export function resolveCloneHiLowLines(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return sourceValue === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve a `serLines` override. Mirrors {@link resolveCloneDropLines} /
 * {@link resolveCloneHiLowLines}; the only difference is the per-family
 * scope — `<c:serLines>` has no slot on `<c:lineChart>` /
 * `<c:areaChart>` / `<c:pieChart>` / `<c:doughnutChart>` /
 * `<c:scatterChart>`.
 */
export function resolveCloneSerLines(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return sourceValue === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
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
export function resolveCloneScatterStyle(
  sourceValue: ChartScatterStyle | undefined,
  override: ChartScatterStyle | null | undefined,
): ChartScatterStyle | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

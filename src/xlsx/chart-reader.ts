// ── Chart Reader ──────────────────────────────────────────────────
// Parses xl/charts/chartN.xml into a minimal structured record.
//
// Charts in OOXML live under the `c` (DrawingML chart) namespace. The
// root element is `<c:chartSpace>` and the visible chart sits inside
// `<c:chart>`. The `<c:plotArea>` child contains one or more chart-type
// elements (`<c:barChart>`, `<c:lineChart>`, `<c:pieChart>`, ...) — each
// chart-type element holds the series and axis bindings.
//
// This reader only extracts metadata that's cheap to surface: the chart
// kind(s), title, and series count. It does not decode series bindings.
//
// OOXML reference: ECMA-376 Part 1, §21.2 (DrawingML — Charts).

import type {
  Chart,
  ChartAxisCrossBetween,
  ChartAxisCrosses,
  ChartAxisDispUnit,
  ChartAxisDispUnits,
  ChartAxisGridlines,
  ChartAxisInfo,
  ChartAxisLabelAlign,
  ChartAxisNumberFormat,
  ChartAxisScale,
  ChartAxisTickLabelPosition,
  ChartAxisTickMark,
  ChartBarGrouping,
  ChartBorderDash,
  ChartDataLabelPosition,
  ChartDataLabelsInfo,
  ChartDataTable,
  ChartDisplayBlanksAs,
  ChartKind,
  ChartLegendEntry,
  ChartLegendPosition,
  ChartLineAreaGrouping,
  ChartLineDashStyle,
  ChartLineStroke,
  ChartManualLayout,
  ChartMarker,
  ChartMarkerSymbol,
  ChartProtection,
  ChartScatterStyle,
  ChartSeriesInfo,
  ChartView3D,
} from "../_types";
import { parseXml } from "../xml/parser";
import type { XmlElement } from "../xml/parser";
import {
  EMU_PER_PT,
  STROKE_WIDTH_MAX_PT,
  STROKE_WIDTH_MIN_PT,
  VALID_DASH_STYLES,
  normalizeRgbHex,
  parseBorderDashFromSpPr,
  parseBorderWidthFromSpPr,
  parseSpPrBorderColor,
  parseSpPrFill,
} from "./chart/shape";
import { parseManualLayout } from "./chart/layout";
import {
  parseBackWallThickness,
  parseFloorThickness,
  parseSideWallThickness,
  parseView3D,
} from "./chart/walls";
import {
  parseTitle,
  parseTitleBold,
  parseTitleBorderColor,
  parseTitleBorderDash,
  parseTitleBorderWidth,
  parseTitleColor,
  parseTitleFillColor,
  parseTitleFontFamily,
  parseTitleFontSize,
  parseTitleItalic,
  parseTitleLayout,
  parseTitleOverlay,
  parseTitleRotation,
  parseTitleStrike,
  parseTitleUnderline,
} from "./chart/title";
import {
  parseLegend,
  parseLegendBold,
  parseLegendBorderColor,
  parseLegendBorderDash,
  parseLegendBorderWidth,
  parseLegendEntries,
  parseLegendFillColor,
  parseLegendFontColor,
  parseLegendFontFamily,
  parseLegendFontSize,
  parseLegendItalic,
  parseLegendLayout,
  parseLegendOverlay,
  parseLegendStrikethrough,
  parseLegendUnderline,
} from "./chart/legend";
import { parseDataTable } from "./chart/dataTable";
import { parseDataLabels } from "./chart/dataLabels";
import { parseSeries } from "./chart/series";
import { parseAutoTitleDeleted, parseAxisInfo } from "./chart/axis";
import {
  parseAxes,
  parseBarGrouping,
  parseDropLines,
  parseFirstSliceAng,
  parseGapWidth,
  parseHiLowLines,
  parseHoleSize,
  parseLineAreaGrouping,
  parseOverlap,
  parsePlotAreaBorderColor,
  parsePlotAreaBorderWidth,
  parsePlotAreaFillColor,
  parsePlotAreaLayout,
  parseScatterStyle,
  parseSerLines,
  parseShowLineMarkers,
  parseUpDownBarsGapWidth,
  parseVaryColors,
} from "./chart/plotArea";

/** All chart-type element local names recognized by Excel. */
const CHART_KIND_TAGS: ReadonlyMap<string, ChartKind> = new Map([
  ["barChart", "bar"],
  ["bar3DChart", "bar3D"],
  ["lineChart", "line"],
  ["line3DChart", "line3D"],
  ["pieChart", "pie"],
  ["pie3DChart", "pie3D"],
  ["doughnutChart", "doughnut"],
  ["areaChart", "area"],
  ["area3DChart", "area3D"],
  ["scatterChart", "scatter"],
  ["bubbleChart", "bubble"],
  ["radarChart", "radar"],
  ["surfaceChart", "surface"],
  ["surface3DChart", "surface3D"],
  ["stockChart", "stock"],
  ["ofPieChart", "ofPie"],
]);

/**
 * Parse a chart file (`xl/charts/chartN.xml`) into a {@link Chart}.
 *
 * Returns `undefined` when the document is not recognizable as a
 * `c:chartSpace`. Returns a record with `kinds: []` when the chart has
 * no chart-type element (extremely rare, but possible for empty charts).
 */
export function parseChart(xml: string): Chart | undefined {
  const root = parseXml(xml);
  // chartSpace can be the root, or wrapped if the file has been
  // pre-processed; tolerate both shapes.
  const chartSpace = root.local === "chartSpace" ? root : findDescendant(root, "chartSpace");
  if (!chartSpace) return undefined;

  const chartEl = findChild(chartSpace, "chart");
  if (!chartEl) return { kinds: [], seriesCount: 0 };

  const out: Chart = { kinds: [], seriesCount: 0 };

  const title = parseTitle(chartEl);
  if (title !== undefined) out.title = title;

  // `<c:overlay>` is a child of `<c:title>`, so a chart that omits the
  // title element has no overlay flag to surface — pulling the value
  // off a `<c:title>` that is not part of the chart's render would leak
  // a toggle that has no effect. Only attempt the parse when the chart
  // declares a title element.
  const titleOverlay = parseTitleOverlay(chartEl);
  if (titleOverlay !== undefined) out.titleOverlay = titleOverlay;

  // `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title>`
  // mirrors Excel's "Format Chart Title -> Size & Properties ->
  // Alignment -> Custom angle" knob. Same scope rule as `<c:overlay>` —
  // a chart that omits the `<c:title>` element has no rotation to
  // surface, so the helper short-circuits to `undefined` when the title
  // is absent. The value comes back in whole degrees (range `-90..90`)
  // for symmetry with the writer-side
  // {@link SheetChart.titleRotation} field.
  const titleRotation = parseTitleRotation(chartEl);
  if (titleRotation !== undefined) out.titleRotation = titleRotation;

  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` mirrors Excel's "Format Chart Title ->
  // Font -> Size" knob. Same scope rule as `<c:overlay>` /
  // `<a:bodyPr rot>` — a chart that omits `<c:title>` (or whose title
  // is a `<c:strRef>` formula reference with no `<c:rich>` body) has
  // no `<a:p>` slot to surface the size from, so the helper
  // short-circuits to `undefined`. The value comes back in points
  // (range `1..400`) for symmetry with the writer-side
  // {@link SheetChart.titleFontSize} field.
  const titleFontSize = parseTitleFontSize(chartEl);
  if (titleFontSize !== undefined) out.titleFontSize = titleFontSize;

  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` mirrors Excel's "Format Chart Title ->
  // Font -> Bold" toggle. Same scope rule as the size — a chart that
  // omits `<c:title>` (or whose title is a `<c:strRef>` formula
  // reference with no `<c:rich>` body) has no `<a:p>` slot to surface
  // the flag from, so the helper short-circuits to `undefined`. The
  // OOXML default `false` collapses to `undefined` so absence and
  // `b="0"` round-trip identically through {@link cloneChart} — only
  // an explicit `b="1"` surfaces `true`. The value threads straight
  // back into the writer-side {@link SheetChart.titleBold} field.
  const titleBold = parseTitleBold(chartEl);
  if (titleBold !== undefined) out.titleBold = titleBold;

  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` mirrors Excel's "Format Chart Title ->
  // Font -> Italic" toggle. Same scope rule as the bold flag — a chart
  // that omits `<c:title>` (or whose title is a `<c:strRef>` formula
  // reference with no `<c:rich>` body) has no `<a:p>` slot to surface
  // the flag from, so the helper short-circuits to `undefined`. The
  // OOXML default `false` collapses to `undefined` so absence and
  // `i="0"` round-trip identically through {@link cloneChart} — only
  // an explicit `i="1"` surfaces `true`. The value threads straight
  // back into the writer-side {@link SheetChart.titleItalic} field.
  const titleItalic = parseTitleItalic(chartEl);
  if (titleItalic !== undefined) out.titleItalic = titleItalic;

  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
  // <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` mirrors Excel's "Format Chart Title ->
  // Font -> Font Color" picker. Same scope rule as the bold/italic
  // flags — a chart that omits `<c:title>` (or whose title is a
  // `<c:strRef>` formula reference with no `<c:rich>` body) has no
  // `<a:p>` slot to surface the fill from, so the helper short-circuits
  // to `undefined`. Non-sRGB color picks (`<a:schemeClr>` theme refs,
  // `<a:hslClr>`, etc.) collapse to `undefined` since only the literal
  // RGB triple round-trips losslessly through the writer. The value
  // threads straight back into the writer-side {@link SheetChart.titleColor}.
  const titleColor = parseTitleColor(chartEl);
  if (titleColor !== undefined) out.titleColor = titleColor;

  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
  // </a:p></c:rich></c:tx></c:title>` mirrors Excel's "Format Chart
  // Title -> Font -> Strikethrough" toggle. Same scope rule as the
  // bold/italic flags — a chart that omits `<c:title>` (or whose
  // title is a `<c:strRef>` formula reference with no `<c:rich>`
  // body) has no `<a:p>` slot to surface the flag from, so the
  // helper short-circuits to `undefined`. Only the UI-default
  // `"sngStrike"` surfaces as `true`; `"noStrike"` (the OOXML
  // application default) and the non-UI `"dblStrike"` both collapse
  // to `undefined` so absence and the OOXML default round-trip
  // identically through {@link cloneChart}. The value threads
  // straight back into the writer-side {@link SheetChart.titleStrike}.
  const titleStrike = parseTitleStrike(chartEl);
  if (titleStrike !== undefined) out.titleStrike = titleStrike;

  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
  // </a:p></c:rich></c:tx></c:title>` mirrors Excel's "Format Chart
  // Title -> Font -> Underline" picker. Same scope rule as the
  // bold / italic / strike flags — a chart that omits `<c:title>`
  // (or whose title is a `<c:strRef>` formula reference with no
  // `<c:rich>` body) has no `<a:p>` slot to surface the flag from,
  // so the helper short-circuits to `undefined`. Only the UI-default
  // `"sng"` surfaces as `true`; `"none"` (the OOXML application
  // default), the non-UI `"dbl"` variant, and the sixteen exotic
  // tokens (`"words"`, `"heavy"`, `"dotted"`, etc.) all collapse to
  // `undefined` so absence and the OOXML default round-trip
  // identically through {@link cloneChart}. The value threads
  // straight back into the writer-side {@link SheetChart.titleUnderline}.
  const titleUnderline = parseTitleUnderline(chartEl);
  if (titleUnderline !== undefined) out.titleUnderline = titleUnderline;

  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
  // typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx></c:title>`
  // mirrors Excel's "Format Chart Title -> Font -> Font" picker. Same
  // scope rule as the size — a chart that omits `<c:title>` (or whose
  // title is a `<c:strRef>` formula reference with no `<c:rich>` body)
  // has no `<a:p>` slot to surface the typeface from, so the helper
  // short-circuits to `undefined`. Empty / whitespace-only `typeface`
  // attributes collapse to `undefined` so absence and the empty form
  // round-trip identically through {@link cloneChart}. The value
  // threads straight back into the writer-side
  // {@link SheetChart.titleFontFamily} field.
  const titleFontFamily = parseTitleFontFamily(chartEl);
  if (titleFontFamily !== undefined) out.titleFontFamily = titleFontFamily;

  // `<c:title><c:layout><c:manualLayout>` carries Excel's "Format Chart
  // Title -> Title Options -> Position -> Custom" placement. CT_Title
  // (ECMA-376 Part 1, §21.2.2.210) places the block between `<c:tx>`
  // and `<c:overlay>`. The reader surfaces the `<c:x>` / `<c:y>` /
  // `<c:w>` / `<c:h>` coordinates off the canonical slot; absence of
  // any meaningful coordinate collapses the field to `undefined` so a
  // fresh chart and a chart that pinned an out-of-range layout both
  // round-trip lossless. Same accept-or-drop grammar as
  // {@link parseLegendLayout}.
  const titleLayout = parseTitleLayout(chartEl);
  if (titleLayout !== undefined) out.titleLayout = titleLayout;

  // `<c:title><c:spPr><a:solidFill>` carries Excel's "Format Chart
  // Title -> Fill -> Solid fill -> Color" picker. CT_Title places the
  // `<c:spPr>` block between `<c:overlay>` and `<c:txPr>` per the
  // schema sequence (ECMA-376 Part 1, §21.2.2.210). The reader
  // surfaces only literal `<a:srgbClr val="RRGGBB"/>` fills; theme
  // references and non-solid fills (`<a:noFill>` / `<a:gradFill>` /
  // `<a:pattFill>` / `<a:blipFill>` / `<a:schemeClr>`) drop to
  // `undefined` so a round-trip never fabricates a fill the writer
  // cannot reproduce on emit. The lookup is on `<c:title>` directly
  // rather than gated on `<c:rich>` so a title authored as a
  // `<c:strRef>` formula reference can still surface its background
  // fill — Excel's "Format Title -> Fill" dialog is independent of
  // whether the text body is rich or a formula.
  const titleFillColor = parseTitleFillColor(chartEl);
  if (titleFillColor !== undefined) out.titleFillColor = titleFillColor;

  // `<c:title><c:spPr><a:ln><a:solidFill>` carries Excel's "Format
  // Chart Title -> Border -> Solid line -> Color" picker. The
  // `<a:ln>` block lives inside the same `<c:spPr>` slot as the fill
  // (`<a:solidFill>`), per CT_ShapeProperties — the reader scopes the
  // lookup to direct children of `<c:title>` so a stray `<c:spPr>`
  // elsewhere (on the plot area, a series, on the legend) cannot
  // leak into this field. Theme references (`<a:schemeClr>`) and
  // non-solid line fills (`<a:noFill>` / `<a:gradFill>` /
  // `<a:pattFill>`) all collapse to `undefined` so a round-trip
  // never fabricates a stroke the writer cannot reproduce on emit.
  const titleBorderColor = parseTitleBorderColor(chartEl);
  if (titleBorderColor !== undefined) out.titleBorderColor = titleBorderColor;

  // `<c:title><c:spPr><a:ln w="EMU">` carries Excel's "Format Chart
  // Title -> Border -> Width" pin. The OOXML `w` attribute stores the
  // stroke width in English Metric Units (1 pt = 12 700 EMU) per
  // CT_LineProperties (ECMA-376 Part 1, §20.1.2.3.24); the reader
  // converts back to points and clamps to the same 0.25..13.5 pt band
  // Excel's UI exposes so a template carrying an exotic width still
  // round-trips through the writer's clamp. Scoped to the title's
  // `<c:spPr>` so a stray `<a:ln w=..>` elsewhere (series stroke, axis
  // line, plot-area / legend border) cannot leak into this field.
  const titleBorderWidth = parseTitleBorderWidth(chartEl);
  if (titleBorderWidth !== undefined) out.titleBorderWidth = titleBorderWidth;

  // `<c:title><c:spPr><a:ln><a:prstDash val=".."/></a:ln></c:spPr>
  // </c:title>` carries Excel's "Format Chart Title -> Border -> Dash
  // type" pin. Same accept-or-drop grammar as every other chart-frame
  // border-dash slot — `"solid"` collapses to `undefined` so absence
  // and the OOXML default round-trip identically. Scoped to the
  // title's `<c:spPr>` so a stray `<a:prstDash>` elsewhere cannot leak
  // in.
  const titleBorderDash = parseTitleBorderDash(chartEl);
  if (titleBorderDash !== undefined) out.titleBorderDash = titleBorderDash;

  // `<c:autoTitleDeleted>` records whether the user explicitly deleted
  // the auto-generated title — independent of whether a literal
  // `<c:title>` is present. The element sits on `<c:chart>` directly
  // (between `<c:title>` and `<c:plotArea>` per CT_Chart, ECMA-376
  // Part 1, §21.2.2.4), not nested inside `<c:title>`, so a chart with
  // no `<c:title>` may still pin the flag. The OOXML default `false`
  // collapses to `undefined` so absence and the default round-trip
  // identically through cloneChart.
  const autoTitleDeleted = parseAutoTitleDeleted(chartEl);
  if (autoTitleDeleted !== undefined) out.autoTitleDeleted = autoTitleDeleted;

  const plotArea = findChild(chartEl, "plotArea");
  if (plotArea) {
    let seriesCount = 0;
    const series: ChartSeriesInfo[] = [];
    let barGrouping: ChartBarGrouping | undefined;
    let lineGrouping: ChartLineAreaGrouping | undefined;
    let areaGrouping: ChartLineAreaGrouping | undefined;
    let chartLevelLabels: ChartDataLabelsInfo | undefined;
    let holeSize: number | undefined;
    let gapWidth: number | undefined;
    let overlap: number | undefined;
    let firstSliceAng: number | undefined;
    let varyColors: boolean | undefined;
    let scatterStyle: ChartScatterStyle | undefined;
    let dropLines: boolean | undefined;
    let hiLowLines: boolean | undefined;
    let serLines: boolean | undefined;
    let upDownBars: boolean | undefined;
    let upDownBarsGapWidth: number | undefined;
    let showLineMarkers: boolean | undefined;
    for (const child of childElements(plotArea)) {
      const kind = CHART_KIND_TAGS.get(child.local);
      if (!kind) continue;
      if (!out.kinds.includes(kind)) out.kinds.push(kind);
      // Pull `<c:varyColors>` off the first chart-type element that
      // carries one. The OOXML schema places `<c:varyColors>` on every
      // chart-type element except `surface`, `surface3D`, and `stock`,
      // so most templates surface a value here. The per-family default
      // collapse (true on pie / doughnut / ofPie, false elsewhere)
      // happens inside `parseVaryColors`.
      if (varyColors === undefined) {
        varyColors = parseVaryColors(child, kind);
      }
      // Pull grouping off the first bar/column-flavored chart-type
      // element. Combo charts that mix bar with line/area would
      // otherwise need a per-series field; for the common case of a
      // single `<c:barChart>` body this is the value Excel applies.
      if (barGrouping === undefined && (kind === "bar" || kind === "bar3D")) {
        barGrouping = parseBarGrouping(child);
      }
      // Pull `<c:gapWidth>` / `<c:overlap>` off the first bar/column
      // chart-type element. Both are CT_BarChart-only knobs — they sit
      // alongside `<c:grouping>` inside `<c:barChart>` / `<c:bar3DChart>`
      // and are ignored elsewhere by the OOXML schema. The OOXML default
      // of `150` (gapWidth) and `0` (overlap) collapse to `undefined`
      // here so absence and the default round-trip identically through
      // {@link cloneChart}.
      if (gapWidth === undefined && (kind === "bar" || kind === "bar3D")) {
        gapWidth = parseGapWidth(child);
      }
      if (overlap === undefined && (kind === "bar" || kind === "bar3D")) {
        overlap = parseOverlap(child);
      }
      // Same shape for line/area: surface the first stacked variant
      // we encounter. `"standard"` collapses to undefined for symmetry
      // with the writer's default.
      if (lineGrouping === undefined && (kind === "line" || kind === "line3D")) {
        lineGrouping = parseLineAreaGrouping(child);
      }
      if (areaGrouping === undefined && (kind === "area" || kind === "area3D")) {
        areaGrouping = parseLineAreaGrouping(child);
      }
      // Pull `<c:holeSize>` off a doughnut chart so a parsed template
      // can round-trip its hole back through {@link cloneChart}.
      if (holeSize === undefined && kind === "doughnut") {
        holeSize = parseHoleSize(child);
      }
      // `<c:firstSliceAng>` lives on `<c:pieChart>` and
      // `<c:doughnutChart>` (also pie3D / ofPie which we lump in here
      // for symmetry — the writer never emits those, but a parsed
      // template carrying one round-trips cleanly into a pie/doughnut
      // clone). `0` collapses to undefined because it is the OOXML
      // default that the writer also treats as absence of the field.
      if (
        firstSliceAng === undefined &&
        (kind === "pie" || kind === "pie3D" || kind === "doughnut" || kind === "ofPie")
      ) {
        firstSliceAng = parseFirstSliceAng(child);
      }
      // `<c:scatterStyle>` lives exclusively on `<c:scatterChart>` per
      // the OOXML schema, so the lookup is gated on the matching kind.
      // The element is required there, but a corrupt template may omit
      // it or carry a token outside the enum — `parseScatterStyle`
      // returns `undefined` in both cases.
      if (scatterStyle === undefined && kind === "scatter") {
        scatterStyle = parseScatterStyle(child);
      }
      // `<c:dropLines>` lives on `<c:lineChart>` / `<c:line3DChart>` /
      // `<c:areaChart>` / `<c:area3DChart>`. The element is bare — its
      // mere presence paints the connectors — so absence collapses to
      // `undefined`.
      if (
        dropLines === undefined &&
        (kind === "line" || kind === "line3D" || kind === "area" || kind === "area3D")
      ) {
        dropLines = parseDropLines(child);
      }
      // `<c:hiLowLines>` lives on `<c:lineChart>` / `<c:line3DChart>` /
      // `<c:stockChart>`. Hucre's writer authors `<c:lineChart>` only,
      // but a stock-chart template that round-trips through hucre will
      // surface the flag here too. Same bare-element shape as
      // `<c:dropLines>`.
      if (hiLowLines === undefined && (kind === "line" || kind === "line3D" || kind === "stock")) {
        hiLowLines = parseHiLowLines(child);
      }
      // `<c:serLines>` lives on `<c:barChart>` / `<c:ofPieChart>` per
      // the OOXML schema. Hucre's writer authors `<c:barChart>` only,
      // but a parsed of-pie template carrying the element should
      // round-trip the flag too. Same bare-element shape as
      // `<c:dropLines>` / `<c:hiLowLines>`.
      if (serLines === undefined && (kind === "bar" || kind === "ofPie")) {
        serLines = parseSerLines(child);
      }
      // `<c:upDownBars>` lives on `CT_LineChart`, `CT_Line3DChart`, and
      // `CT_StockChart` per the OOXML schema. Surface the flag from the
      // first line-flavored chart-type element that carries one — the
      // schema places the element on the chart-type element itself, not
      // the per-series body, so this is a chart-level toggle. Per-bar
      // styling can layer on later.
      if (upDownBars === undefined && (kind === "line" || kind === "line3D" || kind === "stock")) {
        const udb = findChild(child, "upDownBars");
        if (udb !== undefined) {
          upDownBars = true;
          // `<c:gapWidth val="N"/>` (CT_GapAmount, ST_GapAmount) lives
          // inside `<c:upDownBars>` and controls the spacing between the
          // up / down bars themselves. The OOXML default of `150`
          // collapses to `undefined` for symmetry with the writer's
          // {@link SheetChart.upDownBarsGapWidth} default — absence and
          // `150` mean the same thing on roundtrip. Out-of-range values
          // are dropped rather than clamped so a corrupt template does
          // not silently rewrite as a different gap.
          upDownBarsGapWidth = parseUpDownBarsGapWidth(udb);
        }
      }
      // `<c:marker>` (the chart-level CT_Boolean variant) lives on
      // `CT_LineChart` only — `CT_Line3DChart` and `CT_StockChart` have
      // no slot for it per the OOXML schema. Surface the value from the
      // first `<c:lineChart>` element so a combo chart that mixes line
      // with another family still carries the line side's flag. The
      // OOXML / Excel default `val="1"` collapses to `undefined` so
      // absence and the default round-trip identically through
      // {@link cloneChart} — only an explicit `val="0"` surfaces
      // `false`.
      if (showLineMarkers === undefined && kind === "line") {
        showLineMarkers = parseShowLineMarkers(child);
      }
      let localIndex = 0;
      for (const ser of childElements(child)) {
        if (ser.local !== "ser") continue;
        seriesCount++;
        series.push(parseSeries(ser, kind, localIndex));
        localIndex++;
      }
      // Chart-type-level <c:dLbls> sits as a sibling of <c:ser> inside
      // the chart-type element. Surface the first one we find — combo
      // charts can carry one per kind, but the common case is a single
      // chart-type element so we keep the model flat.
      if (chartLevelLabels === undefined) {
        const dLbls = findChild(child, "dLbls");
        if (dLbls) {
          const parsed = parseDataLabels(dLbls);
          if (parsed) chartLevelLabels = parsed;
        }
      }
    }
    out.seriesCount = seriesCount;
    if (series.length > 0) out.series = series;
    if (barGrouping !== undefined) out.barGrouping = barGrouping;
    if (lineGrouping !== undefined) out.lineGrouping = lineGrouping;
    if (areaGrouping !== undefined) out.areaGrouping = areaGrouping;
    if (chartLevelLabels) out.dataLabels = chartLevelLabels;
    if (holeSize !== undefined) out.holeSize = holeSize;
    if (gapWidth !== undefined) out.gapWidth = gapWidth;
    if (overlap !== undefined) out.overlap = overlap;
    if (firstSliceAng !== undefined) out.firstSliceAng = firstSliceAng;
    if (varyColors !== undefined) out.varyColors = varyColors;
    if (scatterStyle !== undefined) out.scatterStyle = scatterStyle;
    if (dropLines !== undefined) out.dropLines = dropLines;
    if (hiLowLines !== undefined) out.hiLowLines = hiLowLines;
    if (serLines !== undefined) out.serLines = serLines;
    if (upDownBars !== undefined) out.upDownBars = upDownBars;
    if (upDownBarsGapWidth !== undefined) out.upDownBarsGapWidth = upDownBarsGapWidth;
    if (showLineMarkers !== undefined) out.showLineMarkers = showLineMarkers;

    const axes = parseAxes(plotArea);
    if (axes !== undefined) out.axes = axes;

    // `<c:dTable>` lives inside `<c:plotArea>` after the axes per
    // CT_PlotArea — the data table renders the underlying series values
    // as a small grid beneath the plot. Only chart families with axes
    // (bar / column / line / area / scatter / surface / stock) carry
    // a slot for it; pie / doughnut have no axes at all.
    const dataTable = parseDataTable(plotArea);
    if (dataTable !== undefined) out.dataTable = dataTable;

    // `<c:plotArea><c:layout><c:manualLayout>` carries Excel's "Format
    // Plot Area -> Position -> Custom" placement. CT_PlotArea places the
    // `<c:layout>` element first, before any chart-type element / axes /
    // `<c:dTable>` / `<c:spPr>`. The reader surfaces the `<c:x>` /
    // `<c:y>` / `<c:w>` / `<c:h>` coordinates off the canonical slot;
    // absence of any meaningful coordinate (or the bare `<c:layout/>`
    // placeholder Excel itself emits on auto-layout charts) collapses
    // the field to `undefined` so a fresh chart and a chart that pinned
    // an out-of-range layout both round-trip lossless.
    const plotAreaLayout = parsePlotAreaLayout(plotArea);
    if (plotAreaLayout !== undefined) out.plotAreaLayout = plotAreaLayout;

    // `<c:plotArea><c:spPr><a:solidFill><a:srgbClr val=".."/></a:solidFill>
    // </c:spPr></c:plotArea>` carries Excel's "Format Plot Area -> Fill ->
    // Solid fill -> Color" pin. CT_PlotArea places the `<c:spPr>` slot at
    // the tail of `<c:plotArea>` after every chart-type element / axes /
    // `<c:dTable>`. The reader surfaces only the literal `<a:srgbClr>`
    // form so absence, malformed hex tokens, non-solid fills (`<a:noFill>`
    // / `<a:gradFill>` / `<a:pattFill>` / `<a:blipFill>`), and theme-color
    // references (`<a:schemeClr>`) all collapse to `undefined` so a
    // round-trip never fabricates a color Excel cannot render.
    const plotAreaFillColor = parsePlotAreaFillColor(plotArea);
    if (plotAreaFillColor !== undefined) out.plotAreaFillColor = plotAreaFillColor;

    // `<c:plotArea><c:spPr><a:ln><a:solidFill><a:srgbClr val=".."/>
    // </a:solidFill></a:ln></c:spPr></c:plotArea>` carries Excel's
    // "Format Plot Area -> Border -> Solid line -> Color" pin. The
    // `<a:ln>` block lives inside the same `<c:spPr>` slot as the fill
    // (`<a:solidFill>`), per CT_ShapeProperties — the reader scopes the
    // lookup so a stray `<a:ln>` elsewhere (on a series, on an axis)
    // cannot leak into this field.
    const plotAreaBorderColor = parsePlotAreaBorderColor(plotArea);
    if (plotAreaBorderColor !== undefined) out.plotAreaBorderColor = plotAreaBorderColor;

    // `<c:plotArea><c:spPr><a:ln w="EMU">` carries Excel's "Format Plot
    // Area -> Border -> Width" pin. The OOXML `w` attribute stores the
    // stroke width in English Metric Units (1 pt = 12 700 EMU) per
    // CT_LineProperties (ECMA-376 Part 1, §20.1.2.3.24); the reader
    // converts back to points and clamps to the same 0.25..13.5 pt band
    // Excel's UI exposes so a template carrying an exotic width still
    // round-trips through the writer's clamp. Scoped to the plot-area's
    // `<c:spPr>` so a stray `<a:ln w=..>` elsewhere (series stroke,
    // axis line) cannot leak into this field.
    const plotAreaBorderWidth = parsePlotAreaBorderWidth(plotArea);
    if (plotAreaBorderWidth !== undefined) out.plotAreaBorderWidth = plotAreaBorderWidth;

    // `<c:plotArea><c:spPr><a:ln><a:prstDash val=".."/></a:ln></c:spPr>
    // </c:plotArea>` carries Excel's "Format Plot Area -> Border ->
    // Dash type" pin. Same accept-or-drop grammar as every other
    // chart-frame border-dash slot.
    const plotAreaBorderDash = parseBorderDashFromSpPr(plotArea);
    if (plotAreaBorderDash !== undefined) out.plotAreaBorderDash = plotAreaBorderDash;
  }

  const legend = parseLegend(chartEl);
  if (legend !== undefined) out.legend = legend;

  // `<c:overlay>` is a child of `<c:legend>`, so a chart that hides the
  // legend (legend === false) or omits the element entirely (legend ===
  // undefined) has no overlay flag to surface — pulling the value off a
  // `<c:legend>` that is not part of the chart's render would leak a
  // toggle that has no effect. Only attempt the parse when the chart
  // declares a visible legend.
  if (legend !== undefined && legend !== false) {
    const legendOverlay = parseLegendOverlay(chartEl);
    if (legendOverlay !== undefined) out.legendOverlay = legendOverlay;

    // `<c:legendEntry>` lives inside `<c:legend>` per CT_Legend
    // (ECMA-376 Part 1, §21.2.2.114) — the element block sits between
    // `<c:legendPos>` and `<c:layout>` / `<c:overlay>`. A hidden or
    // missing legend has no slot for entry overrides, so the parser
    // only inspects the children when the chart actually declares a
    // visible legend. Same scoping rule as `legendOverlay`.
    const legendEntries = parseLegendEntries(chartEl);
    if (legendEntries !== undefined) out.legendEntries = legendEntries;

    // `<c:legend><c:txPr>` carries the legend's typography pins —
    // tick-label / axis-title / chart-title style. The CT_Legend schema
    // places `<c:txPr>` after `<c:overlay>` (and before `<c:extLst>`).
    // A hidden or missing legend has no slot for the block, so the
    // parser only inspects it when the chart actually declares a
    // visible legend. Same scoping rule as `legendOverlay` /
    // `legendEntries`.
    const legendFontSize = parseLegendFontSize(chartEl);
    if (legendFontSize !== undefined) out.legendFontSize = legendFontSize;

    // Same scoping for the bold flag — `<c:txPr>` is the shared host
    // element, and the OOXML default `false` collapses to `undefined`
    // so absence and `b="0"` round-trip identically.
    const legendBold = parseLegendBold(chartEl);
    if (legendBold !== undefined) out.legendBold = legendBold;

    // Same scoping for the italic flag — only an explicit `i="1"`
    // surfaces `true`; the OOXML default and absence both collapse to
    // `undefined`.
    const legendItalic = parseLegendItalic(chartEl);
    if (legendItalic !== undefined) out.legendItalic = legendItalic;

    // Same scoping for the underline flag — only an explicit
    // `u="sng"` (Excel's UI variant) surfaces `true`; the OOXML default
    // `"none"` (and every non-`"sng"` variant the schema allows)
    // collapse to `undefined`.
    const legendUnderline = parseLegendUnderline(chartEl);
    if (legendUnderline !== undefined) out.legendUnderline = legendUnderline;

    // Same scoping for the strikethrough flag — only an explicit
    // `strike="sngStrike"` (Excel's UI variant — single line) surfaces
    // `true`; the OOXML default `"noStrike"` and the non-UI variant
    // `"dblStrike"` (double line) both collapse to `undefined` so a
    // templated chart with the double-line variant round-trips
    // lossless rather than silently downgrading on re-emit.
    const legendStrikethrough = parseLegendStrikethrough(chartEl);
    if (legendStrikethrough !== undefined) out.legendStrikethrough = legendStrikethrough;

    // Same scoping for the font color — `<c:txPr>` is the shared host
    // element. Only the literal `<a:srgbClr val="RRGGBB"/>` round-trips
    // losslessly; theme references / preset colors / malformed val
    // tokens all collapse to `undefined`.
    const legendFontColor = parseLegendFontColor(chartEl);
    if (legendFontColor !== undefined) out.legendFontColor = legendFontColor;

    // Same scoping for the font family — `<c:txPr>` is the shared host
    // element. Empty / whitespace-only `typeface` attributes collapse
    // to `undefined` so absence and the empty form round-trip
    // identically through the writer.
    const legendFontFamily = parseLegendFontFamily(chartEl);
    if (legendFontFamily !== undefined) out.legendFontFamily = legendFontFamily;

    // `<c:legend><c:layout><c:manualLayout>` carries Excel's "Format
    // Legend -> Position -> Custom" placement. CT_Legend places the
    // block between `<c:legendEntry>` and `<c:overlay>`. The reader
    // surfaces the `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` coordinates
    // off the canonical slot; absence of any meaningful coordinate
    // collapses the field to `undefined` so a fresh chart and a chart
    // that pinned an out-of-range layout both round-trip lossless.
    const legendLayout = parseLegendLayout(chartEl);
    if (legendLayout !== undefined) out.legendLayout = legendLayout;

    // `<c:legend><c:spPr><a:solidFill>` carries Excel's "Format Legend
    // -> Fill -> Solid fill -> Color" picker. CT_Legend places the
    // `<c:spPr>` block between `<c:overlay>` and `<c:txPr>`. The
    // reader surfaces only literal `<a:srgbClr val="RRGGBB"/>` fills;
    // theme references and non-solid fills (`<a:noFill>` /
    // `<a:gradFill>` / `<a:pattFill>` / `<a:blipFill>` /
    // `<a:schemeClr>`) drop to `undefined` so a round-trip never
    // fabricates a fill the writer cannot reproduce on emit. Same
    // hidden-legend scoping as the typography knobs.
    const legendFillColor = parseLegendFillColor(chartEl);
    if (legendFillColor !== undefined) out.legendFillColor = legendFillColor;

    // `<c:legend><c:spPr><a:ln><a:solidFill>` carries Excel's "Format
    // Legend -> Border -> Solid line -> Color" picker. The `<a:ln>`
    // block lives inside the same `<c:spPr>` slot as the fill
    // (`<a:solidFill>`), per CT_ShapeProperties — the reader scopes
    // the lookup so a stray `<a:ln>` elsewhere (on a series, on an
    // axis, on the legend's `<c:txPr>` block) cannot leak into this
    // field. Same hidden-legend scoping as the fill / typography
    // knobs.
    const legendBorderColor = parseLegendBorderColor(chartEl);
    if (legendBorderColor !== undefined) out.legendBorderColor = legendBorderColor;

    // `<c:legend><c:spPr><a:ln w="EMU">` carries Excel's "Format Legend
    // -> Border -> Width" pin. The OOXML `w` attribute stores the
    // stroke width in English Metric Units (1 pt = 12 700 EMU) per
    // CT_LineProperties (ECMA-376 Part 1, §20.1.2.3.24); the reader
    // converts back to points and clamps to the same 0.25..13.5 pt band
    // Excel's UI exposes so a template carrying an exotic width still
    // round-trips through the writer's clamp. Scoped to the legend's
    // `<c:spPr>` so a stray `<a:ln w=..>` elsewhere (series stroke,
    // axis line, plot-area border) cannot leak into this field. Same
    // hidden-legend scoping as the fill / border-color / typography
    // knobs.
    const legendBorderWidth = parseLegendBorderWidth(chartEl);
    if (legendBorderWidth !== undefined) out.legendBorderWidth = legendBorderWidth;

    // `<c:legend><c:spPr><a:ln><a:prstDash val=".."/></a:ln></c:spPr>
    // </c:legend>` carries Excel's "Format Legend -> Border -> Dash
    // type" pin. Same accept-or-drop grammar as every other chart-
    // frame border-dash slot.
    const legendBorderDash = parseLegendBorderDash(chartEl);
    if (legendBorderDash !== undefined) out.legendBorderDash = legendBorderDash;
  }

  const dispBlanksAs = parseDispBlanksAs(chartEl);
  if (dispBlanksAs !== undefined) out.dispBlanksAs = dispBlanksAs;

  const plotVisOnly = parsePlotVisOnly(chartEl);
  if (plotVisOnly !== undefined) out.plotVisOnly = plotVisOnly;

  // `<c:showDLblsOverMax>` sits at the tail of CT_Chart (after
  // `<c:dispBlanksAs>` and before `<c:extLst>`). Mirrors the writer
  // side, which always emits the element — only the non-default
  // `val="0"` surfaces here (`true` collapses to `undefined` for the
  // standard minimal-shape contract).
  const showDLblsOverMax = parseShowDLblsOverMax(chartEl);
  if (showDLblsOverMax !== undefined) out.showDLblsOverMax = showDLblsOverMax;

  // `<c:roundedCorners>` lives on `<c:chartSpace>` (the chart's outer
  // wrapper), not inside `<c:chart>` — the toggle styles the chart
  // frame's outer border rather than the plot area.
  const roundedCorners = parseRoundedCorners(chartSpace);
  if (roundedCorners !== undefined) out.roundedCorners = roundedCorners;

  // `<c:style>` also sits on `<c:chartSpace>` — it picks one of the 48
  // built-in chart-style presets that style the entire chart space
  // (frame fill, plot area look, default text font), not just the
  // plot area.
  const style = parseStyle(chartSpace);
  if (style !== undefined) out.style = style;

  // `<c:lang>` records the editing locale Excel used to author the
  // chart. It also sits on `<c:chartSpace>` (per CT_ChartSpace, between
  // `<c:date1904>` and `<c:roundedCorners>`), not inside `<c:chart>` —
  // the value drives locale-sensitive defaults across the entire chart
  // document.
  const lang = parseLang(chartSpace);
  if (lang !== undefined) out.lang = lang;

  // `<c:date1904>` mirrors the host workbook's date-system toggle for
  // chart date-axis interpretation. It sits at the head of
  // `<c:chartSpace>` (per CT_ChartSpace, before `<c:lang>` and
  // `<c:roundedCorners>`), not inside `<c:chart>` — the toggle governs
  // date interpretation across the whole chart document.
  const date1904 = parseDate1904(chartSpace);
  if (date1904 !== undefined) out.date1904 = date1904;

  // `<c:protection>` (CT_Protection, ECMA-376 Part 1, §21.2.2.142)
  // sits on `<c:chartSpace>` between `<c:style>` / `<c:pivotSource>`
  // and `<c:chart>`. The element holds five optional `<xsd:boolean>`
  // children (`<c:chartObject>`, `<c:data>`, `<c:formatting>`,
  // `<c:selection>`, `<c:userInterface>`). Unlike `<c:dTable>` (whose
  // children are required) every protection flag is independently
  // optional, so the reader only surfaces the ones the file actually
  // pinned.
  const protection = parseProtection(chartSpace);
  if (protection !== undefined) out.protection = protection;

  // `<c:view3D>` (CT_View3D, ECMA-376 Part 1, §21.2.2.228) sits on
  // `<c:chart>` between `<c:autoTitleDeleted>` / `<c:pivotFmts>` and
  // `<c:floor>` / `<c:plotArea>`. The element holds six independently
  // optional children (`<c:rotX>`, `<c:hPercent>`, `<c:rotY>`,
  // `<c:depthPercent>`, `<c:rAngAx>`, `<c:perspective>`); the reader
  // surfaces only the fields the file actually pinned. The element is
  // only meaningful on 3D chart families but the OOXML schema accepts
  // it on every CT_Chart, so the reader looks for it on every chart —
  // a stray element on a 2D chart still surfaces here so the round-
  // trip through cloneChart stays lossless.
  const view3D = parseView3D(chartEl);
  if (view3D !== undefined) out.view3D = view3D;

  // `<c:floor>` (CT_Surface, ECMA-376 Part 1, §21.2.2.69) sits on
  // `<c:chart>` between `<c:view3D>` and `<c:sideWall>` /
  // `<c:backWall>` / `<c:plotArea>` per CT_Chart. The reader surfaces
  // only the `<c:thickness>` child here — `<c:spPr>` / `<c:pictureOptions>`
  // / `<c:extLst>` styling on the floor block is not modelled at this
  // layer. Like `<c:view3D>`, the schema accepts `<c:floor>` on every
  // CT_Chart even though it is only meaningful on 3D families, so a
  // stray element on a 2D chart still surfaces the value for round-
  // trip parity.
  const floorThickness = parseFloorThickness(chartEl);
  if (floorThickness !== undefined) out.floorThickness = floorThickness;

  // `<c:sideWall>` (CT_Surface, ECMA-376 Part 1, §21.2.2.187) sits on
  // `<c:chart>` between `<c:floor>` and `<c:backWall>` /
  // `<c:plotArea>` per CT_Chart. The reader surfaces only the
  // `<c:thickness>` child here — `<c:spPr>` / `<c:pictureOptions>` /
  // `<c:extLst>` styling on the side-wall block is not modelled at
  // this layer. Like `<c:view3D>` / `<c:floor>`, the schema accepts
  // `<c:sideWall>` on every CT_Chart even though it is only
  // meaningful on 3D families, so a stray element on a 2D chart still
  // surfaces the value for round-trip parity.
  const sideWallThickness = parseSideWallThickness(chartEl);
  if (sideWallThickness !== undefined) out.sideWallThickness = sideWallThickness;

  // `<c:backWall>` (CT_Surface, ECMA-376 Part 1, §21.2.2.31) sits on
  // `<c:chart>` between `<c:sideWall>` and `<c:plotArea>` per CT_Chart.
  // Like `<c:floor>`, the reader surfaces only the `<c:thickness>`
  // child here — `<c:spPr>` / `<c:pictureOptions>` / `<c:extLst>`
  // styling on the back-wall block is not modelled at this layer. The
  // schema accepts `<c:backWall>` on every CT_Chart even though it is
  // only meaningful on 3D families, so a stray element on a 2D chart
  // still surfaces the value for round-trip parity.
  const backWallThickness = parseBackWallThickness(chartEl);
  if (backWallThickness !== undefined) out.backWallThickness = backWallThickness;

  // `<c:chartSpace><c:spPr><a:solidFill><a:srgbClr val=".."/></a:solidFill>
  // </c:spPr></c:chartSpace>` carries Excel's "Format Chart Area -> Fill
  // -> Solid fill -> Color" pin (the entire chart background, distinct
  // from the inner `<c:plotArea>` slot). CT_ChartSpace places the
  // `<c:spPr>` slot at the tail of the document root, after `<c:chart>`
  // / `<c:externalData>` / `<c:printSettings>` / `<c:userShapes>`. The
  // reader surfaces only the literal `<a:srgbClr>` form so absence,
  // malformed hex tokens, non-solid fills (`<a:noFill>` /
  // `<a:gradFill>` / `<a:pattFill>` / `<a:blipFill>`), and theme-color
  // references (`<a:schemeClr>`) all collapse to `undefined` so a
  // round-trip never fabricates a color Excel cannot render.
  const chartSpaceFillColor = parseChartSpaceFillColor(chartSpace);
  if (chartSpaceFillColor !== undefined) out.chartSpaceFillColor = chartSpaceFillColor;

  // `<c:chartSpace><c:spPr><a:ln><a:solidFill><a:srgbClr val=".."/>
  // </a:solidFill></a:ln></c:spPr></c:chartSpace>` carries Excel's
  // "Format Chart Area -> Border -> Solid line -> Color" pin (the outer
  // border around the entire chart frame, distinct from the inner
  // `<c:plotArea>` stroke). The reader surfaces only the literal
  // `<a:srgbClr>` form so absence, malformed hex tokens, non-solid line
  // fills, and theme-color references all collapse to `undefined` so a
  // round-trip never fabricates a stroke Excel cannot render.
  const chartSpaceBorderColor = parseChartSpaceBorderColor(chartSpace);
  if (chartSpaceBorderColor !== undefined) out.chartSpaceBorderColor = chartSpaceBorderColor;

  // `<c:chartSpace><c:spPr><a:ln w="EMU">` carries Excel's "Format
  // Chart Area -> Border -> Width" pin. Same EMU encoding and clamp /
  // snap grammar as every other chart-frame border-width slot. Scoped
  // to direct children of `<c:chartSpace>` so a stray `<a:ln w=..>`
  // elsewhere cannot leak in.
  const chartSpaceBorderWidth = parseBorderWidthFromSpPr(chartSpace);
  if (chartSpaceBorderWidth !== undefined) out.chartSpaceBorderWidth = chartSpaceBorderWidth;

  // `<c:chartSpace><c:spPr><a:ln><a:prstDash val=".."/>` carries
  // Excel's "Format Chart Area -> Border -> Dash type" pin. Same
  // accept-or-drop grammar as every other chart-frame border-dash
  // slot — `"solid"` collapses to `undefined` so absence and the OOXML
  // default round-trip identically.
  const chartSpaceBorderDash = parseBorderDashFromSpPr(chartSpace);
  if (chartSpaceBorderDash !== undefined) out.chartSpaceBorderDash = chartSpaceBorderDash;

  return out;
}

// ── Axes ──────────────────────────────────────────────────────────

/**
 * Pull a finite numeric `val=".."` attribute off a named child of
 * `parent`. Tolerates whitespace and trailing zeros; returns
 * `undefined` for missing children, missing attributes, and values
 * that fail `Number.isFinite`.
 */
function parseNumericChildVal(parent: XmlElement, localName: string): number | undefined {
  const child = findChild(parent, localName);
  if (!child) return undefined;
  const raw = child.attrs.val;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : undefined;
}

/** Coerce an XML boolean attribute (`"0"`, `"1"`, `"true"`, `"false"`). */
function parseBoolAttr(value: unknown): boolean | undefined {
  if (typeof value !== "string") return undefined;
  const v = value.trim().toLowerCase();
  if (v === "1" || v === "true") return true;
  if (v === "0" || v === "false") return false;
  return undefined;
}

// ── Series ────────────────────────────────────────────────────────

// ── Marker ────────────────────────────────────────────────────────

const VALID_MARKER_SYMBOLS: ReadonlySet<ChartMarkerSymbol> = new Set([
  "none",
  "auto",
  "circle",
  "square",
  "diamond",
  "triangle",
  "x",
  "star",
  "dot",
  "dash",
  "plus",
]);

// ── Data Labels ───────────────────────────────────────────────────

/**
 * Read a boolean-style `val` attribute. Excel emits `"1"` / `"0"` but
 * the OOXML spec also blesses `"true"` / `"false"`. Returns `undefined`
 * when the attribute is missing.
 */
function readBoolAttr(el: XmlElement): boolean | undefined {
  const v = el.attrs.val;
  if (typeof v !== "string") return undefined;
  return v === "1" || v.toLowerCase() === "true";
}

/** Read the `<c:tx>` series-name element (literal or strRef cache). */

/**
 * Walk `<c:val>` / `<c:cat>` / `<c:xVal>` / `<c:yVal>` to its inner
 * `<c:f>` formula text. Returns `undefined` for embedded `<c:numLit>`
 * literal data (no formula) or when the element is absent.
 */
function formulaText(wrapper: XmlElement | undefined): string | undefined {
  if (!wrapper) return undefined;
  const numRef = findChild(wrapper, "numRef") ?? findChild(wrapper, "strRef");
  if (numRef) {
    const f = findChild(numRef, "f");
    if (f) {
      const text = elementText(f).trim();
      if (text.length > 0) return text;
    }
  }
  // Some writers put <c:f> directly under <c:strRef> (already handled
  // above via numRef fallback) or under the wrapper itself.
  const direct = findChild(wrapper, "f");
  if (direct) {
    const text = elementText(direct).trim();
    if (text.length > 0) return text;
  }
  return undefined;
}

/** Pull the first `<a:srgbClr val="RRGGBB">` under `<c:spPr>`. */

// ── Stroke ────────────────────────────────────────────────────────
//
// `STROKE_WIDTH_MIN_PT`, `STROKE_WIDTH_MAX_PT`, `EMU_PER_PT`,
// `VALID_DASH_STYLES`, `VALID_BORDER_DASHES`, `parseBorderWidthFromSpPr`
// and `parseBorderDashFromSpPr` now live in `./chart/shape.ts`. Imported
// at the top of this module so every host-specific helper can keep
// using the same generic primitives without duplicating the OOXML
// schema knowledge across reader / writer / clone.

// ── Legend ────────────────────────────────────────────────────────

/**
 * Pull `<c:chartSpace><c:spPr><a:solidFill><a:srgbClr val=".."/>
 * </a:solidFill></c:spPr></c:chartSpace>` off the chart-space document
 * root. Returns the entire chart background's solid fill color as a
 * 6-character uppercase hex string.
 *
 * The OOXML `<a:srgbClr>` element carries the literal sRGB color
 * (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32) inside the fill
 * choice of `<c:spPr>` (`CT_ShapeProperties`, §20.1.2.3.13). The
 * `<c:spPr>` slot sits at the tail of `<c:chartSpace>` per
 * CT_ChartSpace (§21.2.2.29), after `<c:chart>` / `<c:externalData>` /
 * `<c:printSettings>` / `<c:userShapes>` and before the optional
 * `<c:txPr>` / `<c:extLst>`.
 *
 * The reader surfaces only the literal `<a:srgbClr>` form — absence,
 * non-solid fills (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>` /
 * `<a:blipFill>`), and theme-color references (`<a:schemeClr>`) all
 * collapse to `undefined` so a chart that pinned a fill the writer
 * cannot reproduce on emit drops the field rather than fabricate one
 * Excel would render differently. Malformed `val` tokens (wrong
 * length, non-hex characters, alpha-channel forms, non-string escapes)
 * likewise drop to `undefined`.
 *
 * Mirrors the writer-side {@link SheetChart.chartSpaceFillColor} so a
 * parsed value slots straight into {@link cloneChart} without
 * conversion. The lookup is scoped to direct children of
 * `<c:chartSpace>` so a stray `<c:spPr>` elsewhere (e.g. on
 * `<c:plotArea>` / `<c:legend>` / `<c:title>` / a series) cannot leak
 * into this field. Mirrors {@link parsePlotAreaFillColor} /
 * {@link parseLegendFillColor} — same `<c:spPr><a:solidFill>
 * <a:srgbClr>` chain on a different host element.
 */
function parseChartSpaceFillColor(chartSpace: XmlElement): string | undefined {
  return parseSpPrFill(chartSpace);
}

/**
 * Pull `<c:chartSpace><c:spPr><a:ln><a:solidFill><a:srgbClr val=".."/>
 * </a:solidFill></a:ln></c:spPr></c:chartSpace>` off the chart-space
 * document root. Returns the entire chart frame's border (stroke)
 * color as a 6-character uppercase hex string.
 *
 * The OOXML `<a:srgbClr>` element carries the literal sRGB color
 * (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32) inside the line's
 * solid fill choice (`CT_LineProperties`' `<a:solidFill>` child —
 * §20.1.2.3.24). The `<a:ln>` slot follows the optional
 * `<a:solidFill>` (fill) child inside `<c:spPr>` per
 * `CT_ShapeProperties` (§20.1.2.3.13). The `<c:spPr>` slot itself
 * sits at the tail of `<c:chartSpace>` per CT_ChartSpace (§21.2.2.29).
 *
 * The reader surfaces only the literal `<a:srgbClr>` form — absence,
 * non-solid line fills (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>` /
 * `<a:blipFill>`), and theme-color references (`<a:schemeClr>`) all
 * collapse to `undefined` so a chart that pinned a stroke the writer
 * cannot reproduce on emit drops the field rather than fabricate one
 * Excel would render differently. Malformed `val` tokens (wrong
 * length, non-hex characters, alpha-channel forms, non-string escapes)
 * likewise drop to `undefined`.
 *
 * Mirrors the writer-side {@link SheetChart.chartSpaceBorderColor} so
 * a parsed value slots straight into {@link cloneChart} without
 * conversion. The lookup is scoped to direct children of
 * `<c:chartSpace>` so a stray `<c:spPr>` elsewhere (e.g. on
 * `<c:plotArea>` / `<c:legend>` / `<c:title>` / a series) cannot leak
 * into this field. Mirrors {@link parsePlotAreaBorderColor} /
 * {@link parseLegendBorderColor} / {@link parseTitleBorderColor} —
 * same `<c:spPr><a:ln><a:solidFill><a:srgbClr>` chain on a different
 * host element.
 */
function parseChartSpaceBorderColor(chartSpace: XmlElement): string | undefined {
  return parseSpPrBorderColor(chartSpace);
}

// `readLayoutCoordinate` and the `<c:layout><c:manualLayout>` walk now
// live in `./chart/layout.ts`. Imported at the top of this module so
// every per-host `parseXxxLayout` wrapper shares one accept-or-drop
// grammar.

/**
 * Conversion factor between OOXML's `rot` attribute (60000ths of a
 * degree, the integer Excel writes inside `<a:bodyPr rot="N"/>`) and
 * whole degrees. Excel's UI exposes the -90..90 degree band — the
 * reader clamps anything outside that band so a corrupt template
 * cannot surface a value the writer would never emit.
 */
const TITLE_ROT_PER_DEGREE = 60000;
const TITLE_ROTATION_MIN_DEG = -90;
const TITLE_ROTATION_MAX_DEG = 90;

/**
 * Conversion factor between OOXML's `sz` attribute (100ths of a point,
 * the integer Excel writes inside `<a:defRPr sz="N"/>` /
 * `<a:rPr sz="N"/>`) and whole / half points. The OOXML
 * `ST_TextFontSize` schema restricts `sz` to the inclusive
 * `100..400000` band — the writer's clamp uses the same range
 * converted to points (`1..400`), so any out-of-range value collapses
 * to `undefined` here rather than surface a token Excel would never
 * emit.
 */
const TITLE_FONT_SZ_PER_POINT = 100;
const TITLE_FONT_SIZE_MIN_PT = 1;
const TITLE_FONT_SIZE_MAX_PT = 400;

// ── Title Italic ─────────────────────────────────────────────────────

// ── Title Color ─────────────────────────────────────────────────────

// `normalizeRgbHex` now lives in `./chart/shape.ts` and is imported at
// the top of this module so every host-specific helper (title, axis
// title, legend, plot area, chart space, data labels, data table)
// shares the same accept-or-drop hex grammar.

// ── Title Strike ────────────────────────────────────────────────────

// ── Title Underline ─────────────────────────────────────────────────

// ── Title Font Family ───────────────────────────────────────────────

// ── Auto Title Deleted ────────────────────────────────────────────

// ── Display Blanks As ─────────────────────────────────────────────

/**
 * Pull `<c:dispBlanksAs val=".."/>` off `<c:chart>`. The OOXML default
 * is `"gap"`, which collapses to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart}.
 *
 * Only the three values OOXML defines (`"gap"`, `"zero"`, `"span"`)
 * surface; unknown or malformed values drop to `undefined` rather than
 * fabricate a token Excel rejects.
 */
function parseDispBlanksAs(chartEl: XmlElement): ChartDisplayBlanksAs | undefined {
  const el = findChild(chartEl, "dispBlanksAs");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "zero":
      return "zero";
    case "span":
      return "span";
    case "gap":
      // OOXML default — collapse to undefined for symmetry with the
      // writer's `dispBlanksAs` field.
      return undefined;
    default:
      return undefined;
  }
}

// ── Plot Visible Only ─────────────────────────────────────────────

/**
 * Pull `<c:plotVisOnly val=".."/>` off `<c:chart>`. The OOXML default
 * is `true` (hidden cells drop out of the chart), which collapses to
 * `undefined` so absence and the default round-trip identically
 * through {@link cloneChart} — only an explicit `<c:plotVisOnly val="0"/>`
 * surfaces `false`.
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` /
 * `"0"` / `"false"`); unknown values and missing `val` attributes drop
 * to `undefined` rather than fabricate a flag Excel would not emit.
 */
function parsePlotVisOnly(chartEl: XmlElement): boolean | undefined {
  const el = findChild(chartEl, "plotVisOnly");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "0":
    case "false":
      return false;
    case "1":
    case "true":
      // OOXML default — collapse to undefined for symmetry with the
      // writer's `plotVisOnly` field.
      return undefined;
    default:
      return undefined;
  }
}

// ── Show Data Labels Over Max ─────────────────────────────────────

/**
 * Pull `<c:showDLblsOverMax val=".."/>` off `<c:chart>`. The OOXML
 * default is `true` (data labels render for every point regardless of
 * whether the value exceeds the pinned axis ceiling), which collapses
 * to `undefined` so absence and the default round-trip identically
 * through {@link cloneChart} — only an explicit `<c:showDLblsOverMax val="0"/>`
 * surfaces `false`.
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` /
 * `"0"` / `"false"`); unknown values and missing `val` attributes drop
 * to `undefined` rather than fabricate a flag Excel would not emit.
 *
 * `<c:showDLblsOverMax>` sits at the tail of CT_Chart (after
 * `<c:dispBlanksAs>` and before `<c:extLst>`); the parser pulls it off
 * `<c:chart>` directly, so the toggle's order relative to its sibling
 * elements does not matter.
 */
function parseShowDLblsOverMax(chartEl: XmlElement): boolean | undefined {
  const el = findChild(chartEl, "showDLblsOverMax");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "0":
    case "false":
      return false;
    case "1":
    case "true":
      // OOXML default — collapse to undefined for symmetry with the
      // writer's `showDLblsOverMax` field.
      return undefined;
    default:
      return undefined;
  }
}

// ── Rounded Corners ───────────────────────────────────────────────

/**
 * Pull `<c:roundedCorners val=".."/>` off `<c:chartSpace>`. The OOXML
 * default is `false` (square chart frame), which collapses to
 * `undefined` so absence and the default round-trip identically through
 * {@link cloneChart} — only an explicit `<c:roundedCorners val="1"/>`
 * surfaces `true`.
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` / `"0"`
 * / `"false"`); unknown values and missing `val` attributes drop to
 * `undefined` rather than fabricate a flag Excel would not emit.
 *
 * Note: `<c:roundedCorners>` sits on `<c:chartSpace>`, not inside
 * `<c:chart>` — the toggle styles the chart frame's outer border, not
 * the plot area, and the OOXML schema reflects that with the placement.
 */
function parseRoundedCorners(chartSpace: XmlElement): boolean | undefined {
  const el = findChild(chartSpace, "roundedCorners");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "1":
    case "true":
      return true;
    case "0":
    case "false":
      // OOXML default — collapse to undefined for symmetry with the
      // writer's `roundedCorners` field.
      return undefined;
    default:
      return undefined;
  }
}

// ── Chart Style Preset ────────────────────────────────────────────

/**
 * Pull `<c:style val=".."/>` off `<c:chartSpace>`. Surfaces the
 * integer value verbatim when `val` parses as an integer in the OOXML
 * range (1–48); absence and out-of-range / non-integer values drop to
 * `undefined`.
 *
 * The reader does not pin a default — Excel's reference serialization
 * for a fresh chart emits `<c:style val="2"/>`, but a chart that omits
 * the element renders identically (Excel falls back to its application
 * default). Surfacing only the values that round-trip preserves the
 * minimal-shape contract the rest of {@link Chart} follows.
 *
 * Note: `<c:style>` lives on `<c:chartSpace>`, not inside `<c:chart>`
 * — the preset styles the outer chart space (frame fill, plot area
 * look, default text font), not just the plot area. Per the
 * CT_ChartSpace sequence the element sits after `<c:roundedCorners>`
 * and before `<c:chart>`.
 */
function parseStyle(chartSpace: XmlElement): number | undefined {
  const el = findChild(chartSpace, "style");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  // Strict integer parse — `parseInt` would accept `"3px"` / `"3.5"`,
  // either of which is outside the `xsd:unsignedByte` shape `<c:style>`
  // expects per CT_Style.
  if (!/^\d+$/.test(raw)) return undefined;
  const n = Number(raw);
  if (!Number.isInteger(n)) return undefined;
  if (n < 1 || n > 48) return undefined;
  return n;
}

// ── Editing Locale ────────────────────────────────────────────────

/**
 * Pull `<c:lang val=".."/>` off `<c:chartSpace>`. Surfaces the
 * culture-name verbatim when `val` matches the IETF BCP-47 subset
 * Excel emits (`[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*`, e.g. `en-US`,
 * `tr-TR`, `zh-Hant-TW`); absence and malformed tokens drop to
 * `undefined`.
 *
 * The reader does not pin a default — Excel's reference serialization
 * for a fresh chart authored on an English locale emits `<c:lang
 * val="en-US"/>`, but a chart that omits the element renders
 * identically (Excel falls back to the workbook's editing language).
 * Surfacing only the values that round-trip preserves the minimal-
 * shape contract the rest of {@link Chart} follows.
 *
 * Note: `<c:lang>` lives on `<c:chartSpace>` (per the CT_ChartSpace
 * sequence the element sits between `<c:date1904>` and
 * `<c:roundedCorners>`), not inside `<c:chart>` — the locale governs
 * the entire chart document, not just the plot area.
 */
function parseLang(chartSpace: XmlElement): string | undefined {
  const el = findChild(chartSpace, "lang");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  // Strict shape check — Excel's `<c:lang>` is `xsd:language`
  // (RFC-1766 / BCP-47 culture name). The pattern matches a primary
  // 2- / 3-letter language tag plus zero or more `-`-separated 2–8
  // alphanumeric subtags, which covers everything Excel emits
  // (`en-US`, `tr-TR`, `pt-BR`, `zh-Hans-CN`, …) without admitting
  // raw garbage like `"english"` or `"en US"`.
  if (!/^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$/.test(raw)) return undefined;
  return raw;
}

// ── Date System ────────────────────────────────────────────────────

/**
 * Pull `<c:date1904 val=".."/>` off `<c:chartSpace>`. Surfaces `true`
 * only when the chart pinned `<c:date1904 val="1"/>` (the non-default
 * state — date-axis values inside the chart use the 1904 base, Excel
 * for Mac's legacy epoch where day 0 falls on 1904-01-01). The OOXML
 * default `val="0"` and absence both collapse to `undefined` so
 * absence and the default round-trip identically through
 * {@link cloneChart}.
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` /
 * `"0"` / `"false"`); unknown values and missing `val` attributes drop
 * to `undefined` rather than fabricate a flag Excel would not emit.
 *
 * Note: `<c:date1904>` lives on `<c:chartSpace>` (per CT_ChartSpace
 * the element sits at the head of the sequence, before `<c:lang>`
 * and `<c:roundedCorners>`), not inside `<c:chart>` — the toggle
 * governs date interpretation across the whole chart document, not
 * just the plot area.
 */
function parseDate1904(chartSpace: XmlElement): boolean | undefined {
  const el = findChild(chartSpace, "date1904");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "1":
    case "true":
      return true;
    case "0":
    case "false":
      // OOXML default — collapse to undefined for symmetry with the
      // writer's `date1904` field.
      return undefined;
    default:
      return undefined;
  }
}

// ── Data Table ─────────────────────────────────────────────────────

// ── Protection ────────────────────────────────────────────────────

/**
 * Pull `<c:protection>...</c:protection>` off `<c:chartSpace>`.
 * Surfaces a {@link ChartProtection} object whenever the source chart
 * declares the element; absence collapses to `undefined`.
 *
 * Each of the five boolean children (`<c:chartObject>`, `<c:data>`,
 * `<c:formatting>`, `<c:selection>`, `<c:userInterface>`) is optional
 * on `CT_Protection`, so the reader only surfaces the flags the file
 * actually pinned. Children that are missing or carry an unknown
 * `val` attribute drop to `undefined` rather than fabricate a flag
 * the file did not pin; the writer falls back to the OOXML default
 * `false` for any field the object omits, mirroring how Excel's
 * reader treats a missing child.
 *
 * The element itself is the gating signal — a `<c:protection>` block
 * with no resolvable children surfaces as an empty `{}` rather than
 * `undefined`, so a chart that authors the bare element (Excel's
 * "Protect Chart" preset with every flag at the default) round-trips
 * literally instead of silently disappearing through the parse loop.
 */
function parseProtection(chartSpace: XmlElement): ChartProtection | undefined {
  const el = findChild(chartSpace, "protection");
  if (!el) return undefined;
  const out: ChartProtection = {};
  const chartObject = parseProtectionFlag(el, "chartObject");
  if (chartObject !== undefined) out.chartObject = chartObject;
  const data = parseProtectionFlag(el, "data");
  if (data !== undefined) out.data = data;
  const formatting = parseProtectionFlag(el, "formatting");
  if (formatting !== undefined) out.formatting = formatting;
  const selection = parseProtectionFlag(el, "selection");
  if (selection !== undefined) out.selection = selection;
  const userInterface = parseProtectionFlag(el, "userInterface");
  if (userInterface !== undefined) out.userInterface = userInterface;
  return out;
}

/**
 * Pull a single boolean child off `<c:protection>`. Accepts the OOXML
 * truthy / falsy spellings (`"1"` / `"true"` / `"0"` / `"false"`);
 * unknown tokens, missing `val` attributes, and missing elements all
 * collapse to `undefined` rather than fabricate a flag the file did
 * not pin. Mirrors {@link parseDataTableFlag} — the same OOXML
 * `<xsd:boolean>` lexical-space rule.
 */
function parseProtectionFlag(protection: XmlElement, local: string): boolean | undefined {
  const el = findChild(protection, local);
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "1":
    case "true":
      return true;
    case "0":
    case "false":
      return false;
    default:
      return undefined;
  }
}

// ── 3-D View ──────────────────────────────────────────────────────

// ── Floor Thickness ───────────────────────────────────────────────

// ── Side Wall Thickness ───────────────────────────────────────────

// ── Back Wall Thickness ───────────────────────────────────────────

// ── Vary Colors ────────────────────────────────────────────────────

// ── Bar Grouping ──────────────────────────────────────────────────

// ── Doughnut Hole ─────────────────────────────────────────────────

// ── Bar / Column gap width & overlap ──────────────────────────────

// ── First Slice Angle ─────────────────────────────────────────────

/**
 * Parse an OOXML boolean attribute. The spec allows `"1"` / `"0"` /
 * `"true"` / `"false"`.
 */
function readBoolVal(raw: string | undefined): boolean | undefined {
  if (raw === undefined) return undefined;
  if (raw === "1" || raw === "true") return true;
  if (raw === "0" || raw === "false") return false;
  return undefined;
}

// ── Internals ─────────────────────────────────────────────────────

function collectTextRuns(el: XmlElement, out: string[]): void {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    if (child.local === "t") {
      out.push(elementText(child));
    } else {
      collectTextRuns(child, out);
    }
  }
}

function elementText(el: XmlElement): string {
  let buf = "";
  for (const child of el.children) {
    if (typeof child === "string") buf += child;
    else buf += elementText(child);
  }
  return buf;
}

function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c;
  }
  return undefined;
}

function findDescendant(el: XmlElement, localName: string): XmlElement | undefined {
  if (el.local === localName) return el;
  for (const c of el.children) {
    if (typeof c === "string") continue;
    const hit = findDescendant(c, localName);
    if (hit) return hit;
  }
  return undefined;
}

function childElements(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of el.children) {
    if (typeof c !== "string") out.push(c);
  }
  return out;
}

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

  return out;
}

// ── Axes ──────────────────────────────────────────────────────────

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
function parseAxes(plotArea: XmlElement): { x?: ChartAxisInfo; y?: ChartAxisInfo } | undefined {
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

function parseAxisInfo(
  axis: XmlElement,
  familyDefaultCrossBetween: ChartAxisCrossBetween,
): ChartAxisInfo | undefined {
  const title = parseAxisTitle(axis);
  // `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title>` —
  // axis-title rotation in 60000ths of a degree. Sits on every axis
  // flavour per the OOXML schema (CT_CatAx, CT_ValAx, CT_DateAx,
  // CT_SerAx all share the same `<c:title>` shape). The lookup is
  // scoped to the `<c:title>` body so a stray `<a:bodyPr>` elsewhere
  // on the axis (e.g. on the tick-label `<c:txPr>`) cannot leak in.
  // Out-of-range values clamp to the `-90..90` band Excel's UI
  // exposes; the OOXML default `0` and absence both collapse to
  // `undefined`. Returns `undefined` when the axis omits `<c:title>`
  // entirely or when the title is a `<c:strRef>` (formula reference)
  // with no `<c:rich>` body.
  const axisTitleRotation = parseAxisTitleRotation(axis);
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis-title font size in 100ths of a
  // point. Same `<c:title>` body scope as `axisTitleRotation` so a
  // stray `<a:defRPr>` elsewhere on the axis (e.g. on the tick-label
  // `<c:txPr>`) cannot leak in. The value comes back in points (range
  // `1..400`); out-of-range / non-numeric inputs drop to `undefined`,
  // and absence of `<c:title>` / `<c:rich>` / `<a:p>` / `<a:pPr>` /
  // `<a:defRPr>` / the `sz` attribute likewise collapses to
  // `undefined` for symmetry with the writer-side
  // {@link SheetChart.axes.x.axisTitleFontSize}.
  const axisTitleFontSize = parseAxisTitleFontSize(axis);
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis-title bold flag. Same `<c:title>`
  // body scope as `axisTitleRotation` so a stray `<a:defRPr>` elsewhere
  // on the axis (e.g. on the tick-label `<c:txPr>`) cannot leak in.
  // The OOXML default `false` collapses to `undefined` so absence and
  // `b="0"` round-trip identically through {@link cloneChart} — only
  // an explicit `b="1"` surfaces `true`. Returns `undefined` when the
  // axis omits `<c:title>` entirely or when the title is a `<c:strRef>`
  // (formula reference) with no `<c:rich>` body.
  const axisTitleBold = parseAxisTitleBold(axis);
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis-title italic flag. Same
  // `<c:title>` body scope as `axisTitleFontSize`, so a stray
  // `<a:defRPr>` elsewhere on the axis (e.g. on the tick-label
  // `<c:txPr>`) cannot leak in. The OOXML default `false` collapses to
  // `undefined` so absence and `i="0"` round-trip identically through
  // the writer — only an explicit `i="1"` surfaces `true`. Returns
  // `undefined` when the axis omits `<c:title>` entirely or when the
  // title is a `<c:strRef>` (formula reference) with no `<c:rich>`
  // body.
  const axisTitleItalic = parseAxisTitleItalic(axis);
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
  // <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis-title font color. Same
  // `<c:title>` body scope as `axisTitleItalic`, so a stray
  // `<a:solidFill>` elsewhere on the axis (e.g. on a tick-label
  // `<c:txPr>` block or a `<c:spPr>` series fill) cannot leak in.
  // Theme references (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`,
  // `<a:prstClr>`, and malformed `val` tokens all collapse to
  // `undefined` since only the literal RGB triple round-trips
  // losslessly through the writer. Returns `undefined` when the axis
  // omits `<c:title>` entirely or when the title is a `<c:strRef>`
  // (formula reference) with no `<c:rich>` body.
  const axisTitleColor = parseAxisTitleColor(axis);
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
  // </a:p></c:rich></c:tx></c:title>` — axis-title strikethrough flag.
  // Same `<c:title>` body scope as `axisTitleItalic`, so a stray
  // `<a:defRPr>` elsewhere on the axis (e.g. on the tick-label
  // `<c:txPr>`) cannot leak in. Only the UI-default `"sngStrike"`
  // surfaces as `true`; `"noStrike"` (the OOXML application default)
  // and the non-UI `"dblStrike"` both collapse to `undefined` so absence
  // and the OOXML default round-trip identically through the writer
  // (which emits only `"sngStrike"`). Returns `undefined` when the axis
  // omits `<c:title>` entirely or when the title is a `<c:strRef>`
  // (formula reference) with no `<c:rich>` body.
  const axisTitleStrike = parseAxisTitleStrike(axis);
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
  // </a:p></c:rich></c:tx></c:title>` — axis-title underline flag.
  // Same `<c:title>` body scope as `axisTitleStrike`, so a stray
  // `<a:defRPr>` elsewhere on the axis (e.g. on the tick-label
  // `<c:txPr>`) cannot leak in. Only the UI-default `"sng"` surfaces
  // as `true`; `"none"` (the OOXML application default), the non-UI
  // `"dbl"` variant, and the sixteen exotic tokens all collapse to
  // `undefined` so absence and the OOXML default round-trip
  // identically through the writer (which emits only `"sng"`). Returns
  // `undefined` when the axis omits `<c:title>` entirely or when the
  // title is a `<c:strRef>` (formula reference) with no `<c:rich>`
  // body.
  const axisTitleUnderline = parseAxisTitleUnderline(axis);
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
  // typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx></c:title>` —
  // axis-title font family. Same `<c:title>` body scope as
  // `axisTitleColor` / `axisTitleStrike` / `axisTitleUnderline`, so a
  // stray `<a:latin>` elsewhere on the axis (e.g. on the tick-label
  // `<c:txPr>`) cannot leak in. Empty / whitespace-only `typeface`
  // attributes and missing `<a:latin>` elements both collapse to
  // `undefined` so absence and the empty form round-trip identically
  // through the writer. Returns `undefined` when the axis omits
  // `<c:title>` entirely or when the title is a `<c:strRef>` (formula
  // reference) with no `<c:rich>` body.
  const axisTitleFontFamily = parseAxisTitleFontFamily(axis);
  // `<c:title><c:overlay val=".."/></c:title>` — axis-title overlay
  // flag. Sits as a direct child of `<c:title>` per CT_Title schema,
  // so the lookup is scoped to direct title children. The OOXML
  // default `false` collapses to `undefined` so absence and
  // `<c:overlay val="0"/>` round-trip identically through
  // {@link cloneChart}. Returns `undefined` when the axis omits
  // `<c:title>` entirely.
  const axisTitleOverlay = parseAxisTitleOverlay(axis);
  const gridlines = parseAxisGridlines(axis);
  const scale = parseAxisScale(axis);
  const numberFormat = parseAxisNumberFormat(axis);
  // Tick-mark and tick-label-position children sit alongside the
  // gridlines / numFmt on every CT_CatAx / CT_ValAx / CT_DateAx /
  // CT_SerAx — see CT_TickMark, ST_TickMark, ST_TickLblPos in
  // ECMA-376 Part 1, §21.2.2. The reader collapses each value to
  // `undefined` when it matches the OOXML default so absence and the
  // default round-trip identically through {@link cloneChart}.
  const majorTickMark = parseAxisTickMark(axis, "majorTickMark", "out");
  const minorTickMark = parseAxisTickMark(axis, "minorTickMark", "none");
  const tickLblPos = parseAxisTickLblPos(axis);
  // `<c:txPr><a:bodyPr rot="N"/></c:txPr>` — tick-label rotation in
  // 60000ths of a degree. The element sits on every axis flavour per
  // the OOXML schema (CT_CatAx, CT_ValAx, CT_DateAx, CT_SerAx all
  // carry an optional `<c:txPr>`), so the reader runs on every axis
  // flavour. Out-of-range values clamp to the `-90..90` band Excel's
  // UI exposes; the OOXML default `0` and absence both collapse to
  // `undefined`.
  const labelRotation = parseAxisLabelRotation(axis);
  // `<c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr>` —
  // tick-label font size in 100ths of a point. Same `<c:txPr>` slot
  // as `labelRotation` above. Out-of-range / non-numeric values drop
  // to `undefined` so a corrupt template cannot surface a value the
  // writer would never emit. Surfaced on every axis flavour for
  // symmetry with the writer.
  const labelFontSize = parseAxisLabelFontSize(axis);
  // `<c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr>` —
  // tick-label bold flag. Same `<c:txPr>` slot scope as the rotation /
  // size readers above. The OOXML default `false` collapses to
  // `undefined` so absence and `b="0"` round-trip identically; only
  // an explicit `b="1"` surfaces `true`. Surfaced on every axis
  // flavour for symmetry with the writer.
  const labelBold = parseAxisLabelBold(axis);
  // `<c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr>` —
  // tick-label italic flag. Same `<c:txPr>` slot scope as the bold
  // reader above. The OOXML default `false` collapses to `undefined`
  // so absence and `i="0"` round-trip identically; only an explicit
  // `i="1"` surfaces `true`. Surfaced on every axis flavour for
  // symmetry with the writer.
  const labelItalic = parseAxisLabelItalic(axis);
  // `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val=".."/>
  // </a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>` — tick-label
  // font color. Same `<c:txPr>` slot scope as the rotation / size /
  // bold / italic readers above. Theme references (`<a:schemeClr>`),
  // `<a:hslClr>`, `<a:sysClr>`, `<a:prstClr>`, and malformed `val`
  // tokens all collapse to `undefined` since only the literal RGB
  // triple round-trips losslessly through the writer. Surfaced on
  // every axis flavour for symmetry with the writer.
  const labelColor = parseAxisLabelColor(axis);
  // `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>` —
  // tick-label underline flag. Same `<c:txPr>` slot scope as the
  // rotation / size / bold / italic / color readers above. Only the
  // UI-default `"sng"` surfaces as `true`; the OOXML default `"none"`,
  // absence, the non-UI `"dbl"` variant, and the sixteen exotic
  // tokens all collapse to `undefined` so absence and the OOXML
  // default round-trip identically through the writer (which emits
  // only `"sng"`). Surfaced on every axis flavour for symmetry with
  // the writer.
  const labelUnderline = parseAxisLabelUnderline(axis);
  // `<c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr>` —
  // tick-label strikethrough flag. Same `<c:txPr>` slot scope as the
  // rotation / size / bold / italic / color / underline readers above.
  // Only the UI-default `"sngStrike"` surfaces as `true`; the OOXML
  // default `"noStrike"`, absence, and the non-UI `"dblStrike"` variant
  // all collapse to `undefined` so absence and the OOXML default round-
  // trip identically through the writer (which emits only
  // `"sngStrike"`). Surfaced on every axis flavour for symmetry with
  // the writer.
  const labelStrike = parseAxisLabelStrike(axis);
  // `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
  // </a:pPr></a:p></c:txPr>` — axis tick-label font family. Same
  // axis-level `<c:txPr>` body scope as the rotation / size / bold /
  // italic / color / underline / strike readers above. Empty /
  // whitespace-only `typeface` attributes and missing `<a:latin>`
  // elements both collapse to `undefined` so absence and the empty
  // form round-trip identically through the writer. The reader is
  // scoped to the axis-level `<c:txPr>` so a stray `<a:latin>` inside
  // `<c:title>` (surfaced by `axisTitleFontFamily`) cannot leak in.
  // Surfaced on every axis flavour for symmetry with the writer.
  const labelFontFamily = parseAxisLabelFontFamily(axis);
  // <c:scaling><c:orientation val=".."/></c:scaling> — ST_Orientation
  // accepts "minMax" (default, low → high) and "maxMin" (reversed).
  // The default collapses to undefined so a fresh chart and a chart
  // that explicitly pins "minMax" round-trip identically.
  const reverse = parseAxisReverse(axis);
  // `<c:tickLblSkip>` / `<c:tickMarkSkip>` live exclusively on
  // `CT_CatAx` / `CT_DateAx` per ECMA-376 Part 1, §21.2.2 — the
  // `<c:valAx>` schema rejects them entirely. Skip the parse on
  // value axes so a corrupt template carrying a stray skip element
  // on a value axis does not surface a field the writer would never
  // emit anyway.
  const isCategoryAxis = axis.local === "catAx" || axis.local === "dateAx";
  const tickLblSkip = isCategoryAxis ? parseAxisSkip(axis, "tickLblSkip") : undefined;
  const tickMarkSkip = isCategoryAxis ? parseAxisSkip(axis, "tickMarkSkip") : undefined;
  // `<c:lblOffset>` lives exclusively on `CT_CatAx` / `CT_DateAx` per
  // ECMA-376 Part 1, §21.2.2 — the `<c:valAx>` and `<c:serAx>` schemas
  // reject it. Skip the parse on value axes for the same reason as
  // the skip elements above.
  const lblOffset = isCategoryAxis ? parseAxisLblOffset(axis) : undefined;
  // `<c:lblAlgn>` is also category-axis-only per ECMA-376 Part 1,
  // §21.2.2 — the OOXML `ST_LblAlgn` schema places the element on
  // `CT_CatAx` / `CT_DateAx` only. Same scope rule as `lblOffset`.
  const lblAlgn = isCategoryAxis ? parseAxisLblAlgn(axis) : undefined;
  // `<c:noMultiLvlLbl>` lives exclusively on `CT_CatAx` per ECMA-376
  // Part 1, §21.2.2 — even `<c:dateAx>`, `<c:valAx>`, and `<c:serAx>`
  // reject the element. Skip the parse on every other axis flavour so
  // a corrupt template carrying a stray flag does not surface a value
  // the writer would never emit anyway.
  const noMultiLvlLbl = axis.local === "catAx" ? parseAxisNoMultiLvlLbl(axis) : undefined;
  // `<c:auto>` lives exclusively on `CT_CatAx` per ECMA-376 Part 1,
  // §21.2.2.7 — `<c:dateAx>`, `<c:valAx>`, and `<c:serAx>` reject the
  // element. Skip the parse on every other axis flavour for symmetry
  // with the writer's catAx-only emit path. Only `false` surfaces; the
  // OOXML default `true` (Excel inspects the data and decides whether
  // to treat the axis as a date axis) collapses to `undefined`.
  const auto = axis.local === "catAx" ? parseAxisAuto(axis) : undefined;
  // `<c:delete>` sits on every axis flavour (CT_CatAx / CT_ValAx /
  // CT_DateAx / CT_SerAx) per ECMA-376 Part 1, §21.2.2. The OOXML
  // default `val="0"` (axis visible) collapses to `undefined` so
  // absence and the default round-trip identically.
  const hidden = parseAxisHidden(axis);
  // `<c:crosses>` and `<c:crossesAt>` sit on every axis flavour and live
  // in an XSD choice (CT_Crosses ⊕ CT_Double) — only one may legally
  // appear at a time per ECMA-376 Part 1, §21.2.2. The reader honours
  // the schema by preferring `crossesAt` when both elements show up
  // together (a malformed template); the writer mirrors that order so a
  // round-trip surfaces the numeric pin and drops the redundant
  // semantic toggle.
  const crossesPair = parseAxisCrosses(axis);
  const crosses = crossesPair.crosses;
  const crossesAt = crossesPair.crossesAt;
  // `<c:dispUnits>` lives exclusively on `<c:valAx>` per ECMA-376 Part 1,
  // §21.2.2.32 (CT_ValAx → CT_DispUnits). Skip the parse on every other
  // axis flavour so a corrupt template carrying a stray element does
  // not surface a value the writer would never emit anyway.
  const dispUnits = axis.local === "valAx" ? parseAxisDispUnits(axis) : undefined;
  // `<c:crossBetween>` is also value-axis-only per ECMA-376 Part 1,
  // §21.2.2.10 (CT_ValAx → CT_CrossBetween). The OOXML schema rejects
  // the element on `<c:catAx>` / `<c:dateAx>` / `<c:serAx>`, so the
  // reader skips the parse on every other axis flavour to mirror the
  // writer's scope rule. The element is required on every `<c:valAx>`
  // and Excel always emits the family default — collapse the parsed
  // value when it matches the family default so absence and the
  // default round-trip identically through {@link cloneChart}.
  const parsedCrossBetween = axis.local === "valAx" ? parseAxisCrossBetween(axis) : undefined;
  const crossBetween =
    parsedCrossBetween === familyDefaultCrossBetween ? undefined : parsedCrossBetween;
  if (
    title === undefined &&
    axisTitleRotation === undefined &&
    axisTitleFontSize === undefined &&
    axisTitleBold === undefined &&
    axisTitleItalic === undefined &&
    axisTitleColor === undefined &&
    axisTitleStrike === undefined &&
    axisTitleUnderline === undefined &&
    axisTitleFontFamily === undefined &&
    axisTitleOverlay === undefined &&
    gridlines === undefined &&
    scale === undefined &&
    numberFormat === undefined &&
    majorTickMark === undefined &&
    minorTickMark === undefined &&
    tickLblPos === undefined &&
    labelRotation === undefined &&
    labelFontSize === undefined &&
    labelBold === undefined &&
    labelItalic === undefined &&
    labelColor === undefined &&
    labelFontFamily === undefined &&
    labelUnderline === undefined &&
    labelStrike === undefined &&
    reverse === undefined &&
    tickLblSkip === undefined &&
    tickMarkSkip === undefined &&
    lblOffset === undefined &&
    lblAlgn === undefined &&
    noMultiLvlLbl === undefined &&
    auto === undefined &&
    hidden === undefined &&
    crosses === undefined &&
    crossesAt === undefined &&
    dispUnits === undefined &&
    crossBetween === undefined
  ) {
    return undefined;
  }
  const out: ChartAxisInfo = {};
  if (title !== undefined) out.title = title;
  if (axisTitleRotation !== undefined) out.axisTitleRotation = axisTitleRotation;
  if (axisTitleFontSize !== undefined) out.axisTitleFontSize = axisTitleFontSize;
  if (axisTitleBold !== undefined) out.axisTitleBold = axisTitleBold;
  if (axisTitleItalic !== undefined) out.axisTitleItalic = axisTitleItalic;
  if (axisTitleColor !== undefined) out.axisTitleColor = axisTitleColor;
  if (axisTitleStrike !== undefined) out.axisTitleStrike = axisTitleStrike;
  if (axisTitleUnderline !== undefined) out.axisTitleUnderline = axisTitleUnderline;
  if (axisTitleFontFamily !== undefined) out.axisTitleFontFamily = axisTitleFontFamily;
  if (axisTitleOverlay !== undefined) out.axisTitleOverlay = axisTitleOverlay;
  if (gridlines !== undefined) out.gridlines = gridlines;
  if (scale !== undefined) out.scale = scale;
  if (numberFormat !== undefined) out.numberFormat = numberFormat;
  if (majorTickMark !== undefined) out.majorTickMark = majorTickMark;
  if (minorTickMark !== undefined) out.minorTickMark = minorTickMark;
  if (tickLblPos !== undefined) out.tickLblPos = tickLblPos;
  if (labelRotation !== undefined) out.labelRotation = labelRotation;
  if (labelFontSize !== undefined) out.labelFontSize = labelFontSize;
  if (labelBold !== undefined) out.labelBold = labelBold;
  if (labelItalic !== undefined) out.labelItalic = labelItalic;
  if (labelColor !== undefined) out.labelColor = labelColor;
  if (labelFontFamily !== undefined) out.labelFontFamily = labelFontFamily;
  if (labelUnderline !== undefined) out.labelUnderline = labelUnderline;
  if (labelStrike !== undefined) out.labelStrike = labelStrike;
  if (reverse !== undefined) out.reverse = reverse;
  if (tickLblSkip !== undefined) out.tickLblSkip = tickLblSkip;
  if (tickMarkSkip !== undefined) out.tickMarkSkip = tickMarkSkip;
  if (lblOffset !== undefined) out.lblOffset = lblOffset;
  if (lblAlgn !== undefined) out.lblAlgn = lblAlgn;
  if (noMultiLvlLbl !== undefined) out.noMultiLvlLbl = noMultiLvlLbl;
  if (auto !== undefined) out.auto = auto;
  if (hidden !== undefined) out.hidden = hidden;
  if (crosses !== undefined) out.crosses = crosses;
  if (crossesAt !== undefined) out.crossesAt = crossesAt;
  if (dispUnits !== undefined) out.dispUnits = dispUnits;
  if (crossBetween !== undefined) out.crossBetween = crossBetween;
  return out;
}

/**
 * Recognized values of `<c:majorTickMark>` / `<c:minorTickMark>` per
 * the OOXML `ST_TickMark` enumeration.
 */
const VALID_TICK_MARKS: ReadonlySet<ChartAxisTickMark> = new Set(["none", "in", "out", "cross"]);

/**
 * Pull `<c:majorTickMark val=".."/>` (or `<c:minorTickMark>`) off an
 * axis element. Returns `undefined` when the element is absent, the
 * `val` attribute is missing, the value is not in
 * {@link VALID_TICK_MARKS}, or the value matches the per-element
 * OOXML default — `"out"` for major, `"none"` for minor — so absence
 * and the default round-trip identically.
 */
function parseAxisTickMark(
  axis: XmlElement,
  localName: "majorTickMark" | "minorTickMark",
  defaultValue: ChartAxisTickMark,
): ChartAxisTickMark | undefined {
  const el = findChild(axis, localName);
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const value = raw.trim() as ChartAxisTickMark;
  if (!VALID_TICK_MARKS.has(value)) return undefined;
  return value === defaultValue ? undefined : value;
}

/**
 * Recognized values of `<c:tickLblPos>` per the OOXML
 * `ST_TickLblPos` enumeration.
 */
const VALID_TICK_LBL_POSITIONS: ReadonlySet<ChartAxisTickLabelPosition> = new Set([
  "nextTo",
  "low",
  "high",
  "none",
]);

/**
 * Pull `<c:tickLblPos val=".."/>` off an axis element. Returns
 * `undefined` when the element is absent, the `val` attribute is
 * missing or unrecognized, or the value matches the OOXML default
 * `"nextTo"` so absence and the default round-trip identically.
 */
function parseAxisTickLblPos(axis: XmlElement): ChartAxisTickLabelPosition | undefined {
  const el = findChild(axis, "tickLblPos");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const value = raw.trim() as ChartAxisTickLabelPosition;
  if (!VALID_TICK_LBL_POSITIONS.has(value)) return undefined;
  return value === "nextTo" ? undefined : value;
}

/**
 * Pull the `ST_Orientation` value off `<c:scaling><c:orientation/></c:scaling>`.
 * Returns `true` only when the axis pinned `"maxMin"` (Excel's
 * "Categories / Values in reverse order" toggle); the OOXML default
 * `"minMax"` collapses to `undefined` so absence and the default
 * round-trip identically. Unknown tokens (e.g. typo'd templates) drop
 * to `undefined` rather than fabricate a flag.
 */
function parseAxisReverse(axis: XmlElement): boolean | undefined {
  const scaling = findChild(axis, "scaling");
  if (!scaling) return undefined;
  const orientation = findChild(scaling, "orientation");
  if (!orientation) return undefined;
  const raw = orientation.attrs.val;
  if (typeof raw !== "string") return undefined;
  const value = raw.trim();
  if (value === "maxMin") return true;
  // "minMax" and unknown tokens both fall through to undefined — only
  // an explicit reversed orientation surfaces.
  return undefined;
}

/**
 * Pull `<c:tickLblSkip val=".."/>` or `<c:tickMarkSkip val=".."/>`
 * off a category axis element. Returns `undefined` when:
 *   - the element is absent,
 *   - the `val` attribute is missing or non-numeric,
 *   - the parsed value is `1` (the OOXML default — show every label /
 *     mark),
 *   - the parsed value falls outside the OOXML `ST_SkipIntervals`
 *     range (`1..32767`).
 *
 * Negative / zero / out-of-range inputs are dropped rather than
 * clamped so a corrupt template cannot leak a skip count Excel would
 * reject.
 */
function parseAxisSkip(
  axis: XmlElement,
  localName: "tickLblSkip" | "tickMarkSkip",
): number | undefined {
  const el = findChild(axis, localName);
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < 1 || parsed > 32767) return undefined;
  if (parsed === 1) return undefined;
  return parsed;
}

/**
 * Pull `<c:lblOffset val=".."/>` off a category axis element. Returns
 * `undefined` when:
 *   - the element is absent,
 *   - the `val` attribute is missing or non-numeric,
 *   - the parsed value is `100` (the OOXML default — Excel's
 *     reference label spacing),
 *   - the parsed value falls outside the OOXML
 *     `ST_LblOffsetPercent` range (`0..1000`).
 *
 * Out-of-range / non-numeric inputs are dropped rather than clamped
 * so a corrupt template cannot leak an offset Excel would reject.
 */
function parseAxisLblOffset(axis: XmlElement): number | undefined {
  const el = findChild(axis, "lblOffset");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < 0 || parsed > 1000) return undefined;
  if (parsed === 100) return undefined;
  return parsed;
}

/**
 * Recognized values of `<c:lblAlgn>` per the OOXML `ST_LblAlgn`
 * enumeration.
 */
const VALID_LBL_ALIGNS: ReadonlySet<ChartAxisLabelAlign> = new Set(["ctr", "l", "r"]);

/**
 * Pull `<c:lblAlgn val=".."/>` off a category axis element. Returns
 * `undefined` when:
 *   - the element is absent,
 *   - the `val` attribute is missing or blank,
 *   - the value is not in {@link VALID_LBL_ALIGNS},
 *   - the value is `"ctr"` (the OOXML default — Excel's reference
 *     centered alignment).
 *
 * Unknown tokens drop rather than fall through to the default so a
 * corrupt template cannot leak an alignment Excel would reject.
 */
function parseAxisLblAlgn(axis: XmlElement): ChartAxisLabelAlign | undefined {
  const el = findChild(axis, "lblAlgn");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const value = raw.trim() as ChartAxisLabelAlign;
  if (!VALID_LBL_ALIGNS.has(value)) return undefined;
  return value === "ctr" ? undefined : value;
}

/**
 * Pull `<c:noMultiLvlLbl val=".."/>` off a category axis element.
 * Returns `true` only when the axis pinned `val="1"` / `val="true"`
 * (Excel's "Multi-level Category Labels" checkbox unchecked, i.e.
 * tiered category labels collapsed onto a single line). The OOXML
 * default `val="0"` / `val="false"`, absence, missing `val`, and
 * unknown tokens all collapse to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart}.
 *
 * Mirrors the truthy / falsy parsing in {@link parseAxisHidden} —
 * the OOXML schema (`xsd:boolean`) accepts `0` / `1` / `false` /
 * `true` for `<c:noMultiLvlLbl>` just as it does for every other
 * Boolean-valued chart attribute.
 */
function parseAxisNoMultiLvlLbl(axis: XmlElement): boolean | undefined {
  const el = findChild(axis, "noMultiLvlLbl");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw.trim()) {
    case "1":
    case "true":
      return true;
    case "0":
    case "false":
      return undefined;
    default:
      return undefined;
  }
}

/**
 * Pull `<c:auto val=".."/>` off a category axis element. Returns
 * `false` only when the axis pinned `val="0"` / `val="false"` (Excel's
 * "Text axis" radio button under "Format Axis -> Axis Options -> Axis
 * Type" — Excel keeps every label as-is regardless of whether the
 * cells parse as dates / numerics). The OOXML default `val="1"` /
 * `val="true"` (Excel inspects the data and decides whether to treat
 * the axis as a discrete category axis or a chronological date axis),
 * absence, missing `val`, and unknown tokens all collapse to
 * `undefined` so absence and the default round-trip identically
 * through {@link cloneChart}.
 *
 * Mirrors the truthy / falsy parsing in {@link parseAxisNoMultiLvlLbl}
 * — the OOXML schema (`xsd:boolean`) accepts `0` / `1` / `false` /
 * `true` for `<c:auto>` just as it does for every other Boolean-valued
 * chart attribute. The element's default is the OOXML inverse of
 * `noMultiLvlLbl` (auto defaults to `true`, noMultiLvlLbl defaults to
 * `false`), so this parser collapses `true` rather than `false`.
 */
function parseAxisAuto(axis: XmlElement): boolean | undefined {
  const el = findChild(axis, "auto");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw.trim()) {
    case "0":
    case "false":
      return false;
    case "1":
    case "true":
      // OOXML default — collapse to undefined so absence and the
      // default round-trip identically.
      return undefined;
    default:
      return undefined;
  }
}

/**
 * Pull `<c:delete val=".."/>` off an axis element. Returns `true`
 * only when the axis pinned `val="1"` / `val="true"` (Excel's "hide
 * axis" toggle). The OOXML default `val="0"` / `val="false"`,
 * absence, missing `val`, and unknown tokens all collapse to
 * `undefined` so absence and the default round-trip identically.
 *
 * Mirrors the truthy / falsy parsing in {@link parsePlotVisOnly} —
 * the OOXML schema (`xsd:boolean`) accepts `0` / `1` / `false` /
 * `true` for `<c:delete>` just as it does for every other Boolean-
 * valued chart attribute.
 */
function parseAxisHidden(axis: XmlElement): boolean | undefined {
  const el = findChild(axis, "delete");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw.trim()) {
    case "1":
    case "true":
      return true;
    case "0":
    case "false":
      // OOXML default — collapse to undefined so absence and the
      // default round-trip identically.
      return undefined;
    default:
      return undefined;
  }
}

/**
 * Conversion factor between OOXML's `rot` attribute (60000ths of a
 * degree, the integer Excel writes inside `<a:bodyPr rot="N"/>`) and
 * whole degrees. Excel's UI exposes the -90..90 degree band — the
 * reader clamps anything outside that band so a corrupt template
 * cannot surface a value the writer would never emit.
 */
const TXPR_ROT_PER_DEGREE = 60000;
const LABEL_ROTATION_MIN_DEG = -90;
const LABEL_ROTATION_MAX_DEG = 90;

/**
 * Pull `<c:txPr><a:bodyPr rot="N"/></c:txPr>` off an axis element.
 * Returns the rotation in whole degrees (range `-90..90`).
 *
 * The OOXML default `0` (and absence of the element / attribute) all
 * collapse to `undefined` so absence and the default round-trip
 * identically through {@link cloneChart}. Non-integer / non-numeric /
 * out-of-range values clamp to the nearest endpoint of the
 * `-90..90` band Excel's UI exposes; non-finite (`NaN`, `Infinity`)
 * inputs drop to `undefined`.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the rotation
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelRotation.
 */
function parseAxisLabelRotation(axis: XmlElement): number | undefined {
  const txPr = findChild(axis, "txPr");
  if (!txPr) return undefined;
  const bodyPr = findChild(txPr, "bodyPr");
  if (!bodyPr) return undefined;
  const raw = bodyPr.attrs.rot;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  // Convert from 60000ths of a degree to whole degrees.
  const degrees = Math.round(parsed / TXPR_ROT_PER_DEGREE);
  if (degrees === 0) return undefined;
  if (degrees < LABEL_ROTATION_MIN_DEG) return LABEL_ROTATION_MIN_DEG;
  if (degrees > LABEL_ROTATION_MAX_DEG) return LABEL_ROTATION_MAX_DEG;
  return degrees;
}

/**
 * Pull `<c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr>`
 * off an axis element. Returns the tick-label font size in points
 * (range `1..400`).
 *
 * Mirrors {@link parseAxisTitleFontSize} for tick labels — same
 * 0.5pt half-step granularity, same `1..400`pt band, same
 * drop-on-out-of-range / non-numeric / absence semantics. The lookup
 * is scoped to the axis-level `<c:txPr>` so a stray `<a:defRPr>`
 * inside `<c:title><c:tx><c:rich>` (surfaced by
 * {@link parseAxisTitleFontSize}) cannot leak in.
 *
 * The OOXML default — `<a:defRPr>` with no `sz` attribute — collapses
 * to `undefined` so absence and the default round-trip identically
 * through {@link cloneChart}. Out-of-range / non-numeric `sz` values
 * drop rather than fabricate a value the writer would never emit.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the size
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelFontSize.
 */
function parseAxisLabelFontSize(axis: XmlElement): number | undefined {
  const txPr = findChild(axis, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.sz;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  // Convert from 100ths of a point to points, rounding to the nearest
  // 0.5pt to match the granularity Excel's UI exposes. Mirrors the
  // chart-level / axis-title `parseTitleFontSize` half-step
  // normalisation.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2);
  const points = halfSteps / 2;
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined;
  return points;
}

/**
 * Pull `<c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr>`
 * off an axis element. Returns the tick-label bold flag.
 *
 * Mirrors {@link parseAxisTitleBold} for tick labels — same
 * canonical-slot pair, same drop-on-default-`false` semantics. The
 * OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * default `false` collapses to `undefined` so absence and `b="0"`
 * round-trip identically — only an explicit `b="1"` surfaces `true`.
 * Unknown / malformed `b` tokens drop to `undefined` rather than
 * fabricate a value the writer would never emit.
 *
 * The lookup is scoped to the axis-level `<c:txPr>` so a stray
 * `<a:defRPr b=".."/>` inside `<c:title><c:tx><c:rich>` (surfaced by
 * {@link parseAxisTitleBold}) cannot leak in. Returns `undefined`
 * whenever the axis omits `<c:txPr>` entirely or the canonical
 * `<a:p><a:pPr><a:defRPr>` chain is malformed.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the flag
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelBold.
 */
function parseAxisLabelBold(axis: XmlElement): boolean | undefined {
  const txPr = findChild(axis, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const parsed = parseBoolAttr(defRPr.attrs.b);
  // The OOXML default `false` collapses to `undefined` so absence and
  // `b="0"` round-trip identically through the writer — only an
  // explicit `b="1"` surfaces `true`.
  if (parsed === true) return true;
  return undefined;
}

/**
 * Pull the axis tick-label italic flag off the canonical
 * `<c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr>` chain
 * Excel writes when the user pins italic on the axis tick labels.
 *
 * Returns `true` only when the parser walks the full chain and lands on
 * an `<a:defRPr i="1"/>` (or the OOXML truthy spelling `i="true"`); the
 * OOXML default `i="0"` collapses to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart}. Returns
 * `undefined` whenever the axis omits `<c:txPr>` entirely or the
 * canonical `<a:p><a:pPr><a:defRPr>` chain is malformed.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the flag
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelItalic.
 */
function parseAxisLabelItalic(axis: XmlElement): boolean | undefined {
  const txPr = findChild(axis, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const parsed = parseBoolAttr(defRPr.attrs.i);
  // The OOXML default `false` collapses to `undefined` so absence and
  // `i="0"` round-trip identically through the writer — only an
  // explicit `i="1"` surfaces `true`.
  if (parsed === true) return true;
  return undefined;
}

/**
 * Pull the axis tick-label font color off the canonical
 * `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>` chain Excel writes
 * when the user pins a custom font color on the axis tick labels.
 *
 * Returns the 6-character uppercase hex string when the parser walks
 * the full chain and lands on an `<a:srgbClr val="RRGGBB"/>`. Theme
 * references (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`, and
 * `<a:prstClr>` all collapse to `undefined` — only the literal RGB
 * triple round-trips losslessly through {@link writeChart}. Malformed
 * `val` tokens (wrong length, non-hex characters) likewise drop to
 * `undefined` rather than fabricate a value the writer would round-
 * trip into a malformed `<a:srgbClr>`.
 *
 * Returns `undefined` whenever the axis omits `<c:txPr>` entirely or
 * the canonical `<a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr>` chain
 * is malformed at any link.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the value
 * regardless of axis flavour so a parsed chart preserves the color
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelColor. The lookup is scoped to the
 * axis-level `<c:txPr>` so a stray `<a:solidFill>` inside `<c:title>`
 * (surfaced by {@link parseAxisTitleColor}) or on a `<c:spPr>` series
 * fill cannot leak in.
 */
function parseAxisLabelColor(axis: XmlElement): string | undefined {
  const txPr = findChild(axis, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const solidFill = findChild(defRPr, "solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "srgbClr");
  if (!srgbClr) return undefined;
  return normalizeRgbHex(srgbClr.attrs.val);
}

/**
 * Pull the axis tick-label underline flag off the canonical
 * `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>` chain
 * Excel writes when the user pins an underline on the axis tick labels.
 *
 * The OOXML `u` attribute is the `ST_TextUnderlineType` enum on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
 * eighteen values; Excel's UI exposes only `"sng"` (single line — the
 * default underline checkbox) and `"dbl"` (double line). The reader
 * surfaces only the UI-default `"sng"` as `true`; `"none"` (the OOXML
 * application default), absence, the non-UI `"dbl"` variant, and the
 * sixteen exotic tokens (`"words"`, `"heavy"`, `"dotted"`,
 * `"dottedHeavy"`, `"dash"`, `"dashHeavy"`, `"dashLong"`,
 * `"dashLongHeavy"`, `"dotDash"`, `"dotDashHeavy"`, `"dotDotDash"`,
 * `"dotDotDashHeavy"`, `"wavy"`, `"wavyHeavy"`, `"wavyDbl"`) all
 * collapse to `undefined` — the writer emits only `"sng"`, so
 * reporting any non-single underline as `true` would silently
 * downgrade the choice to a single line on round-trip. Unknown /
 * malformed `u` tokens likewise drop to `undefined`.
 *
 * Mirrors {@link parseAxisTitleUnderline} for tick labels — same
 * canonical-slot semantics but scoped to the axis-level `<c:txPr>` so
 * a stray `<a:defRPr u=".."/>` inside `<c:title><c:tx><c:rich>`
 * (surfaced by {@link parseAxisTitleUnderline}) cannot leak in.
 * Returns `undefined` whenever the axis omits `<c:txPr>` entirely or
 * the canonical `<a:p><a:pPr><a:defRPr>` chain is malformed.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the flag
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelUnderline.
 */
function parseAxisLabelUnderline(axis: XmlElement): boolean | undefined {
  const txPr = findChild(axis, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.u;
  // Only the UI-default `"sng"` surfaces as `true`. The OOXML
  // application default `"none"`, the non-UI `"dbl"` variant, and
  // every exotic token (`"words"`, `"heavy"`, `"dotted"`, etc.) all
  // collapse to `undefined` so absence and the OOXML default
  // round-trip identically through the writer; the writer emits only
  // `"sng"`, so reporting a non-single underline here would silently
  // downgrade the choice on round-trip.
  if (raw === "sng") return true;
  return undefined;
}

/**
 * Pull the axis tick-label strikethrough flag off the canonical
 * `<c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr>`
 * chain Excel writes when the user pins a strikethrough on the axis
 * tick labels.
 *
 * The OOXML `strike` attribute is the `ST_TextStrikeType` enum on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
 * three values: `"noStrike"` (the OOXML application default),
 * `"sngStrike"` (single line — the value Excel's UI checkbox
 * emits), and `"dblStrike"` (double line — a non-UI variant). The
 * reader surfaces only the UI-default `"sngStrike"` as `true`;
 * `"noStrike"` (the OOXML application default), absence, and the
 * non-UI `"dblStrike"` variant all collapse to `undefined` — the
 * writer emits only `"sngStrike"`, so reporting `"dblStrike"` as
 * `true` would silently downgrade the choice to a single line on
 * round-trip. Unknown / malformed `strike` tokens likewise drop to
 * `undefined`.
 *
 * Mirrors {@link parseAxisTitleStrike} for tick labels — same
 * canonical-slot semantics but scoped to the axis-level `<c:txPr>` so
 * a stray `<a:defRPr strike=".."/>` inside `<c:title><c:tx><c:rich>`
 * (surfaced by {@link parseAxisTitleStrike}) cannot leak in. Returns
 * `undefined` whenever the axis omits `<c:txPr>` entirely or the
 * canonical `<a:p><a:pPr><a:defRPr>` chain is malformed.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the flag
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelStrike.
 */
function parseAxisLabelStrike(axis: XmlElement): boolean | undefined {
  const txPr = findChild(axis, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.strike;
  // Only the UI-default `"sngStrike"` surfaces as `true`. The OOXML
  // application default `"noStrike"`, the non-UI `"dblStrike"` variant,
  // and unknown / malformed tokens all collapse to `undefined` so
  // absence and the OOXML default round-trip identically through the
  // writer; the writer emits only `"sngStrike"`, so reporting
  // `"dblStrike"` here would silently downgrade the choice on round-
  // trip.
  if (raw === "sngStrike") return true;
  return undefined;
}

/**
 * Pull the axis tick-label font family off the canonical
 * `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
 * </a:pPr></a:p></c:txPr>` chain Excel writes when the user pins a
 * typeface on the axis tick labels.
 *
 * The OOXML `<a:latin>` element carries the typeface name on
 * `CT_TextFont` (ECMA-376 Part 1, §21.1.2.3.7). The reader trims
 * surrounding whitespace and reports the trimmed typeface; empty /
 * whitespace-only `typeface` attributes and missing `<a:latin>`
 * elements both collapse to `undefined` so absence and the empty
 * form round-trip identically through the writer. Non-string
 * `typeface` tokens (defensive — the XML parser only ever surfaces
 * strings) likewise drop to `undefined`.
 *
 * Returns `undefined` whenever the axis omits `<c:txPr>` entirely or
 * the canonical `<a:p><a:pPr><a:defRPr><a:latin>` chain is malformed
 * at any link.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the value
 * regardless of axis flavour so a parsed chart preserves the
 * typeface for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelFontFamily. The lookup is scoped to
 * the axis-level `<c:txPr>` so a stray `<a:latin>` inside the axis
 * `<c:title><c:tx><c:rich>` body cannot leak in.
 */
function parseAxisLabelFontFamily(axis: XmlElement): string | undefined {
  const txPr = findChild(axis, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const latin = findChild(defRPr, "latin");
  if (!latin) return undefined;
  const raw = latin.attrs.typeface;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  return trimmed;
}

/** Recognized values of `<c:crosses>` per the OOXML `ST_Crosses` enum. */
const VALID_CROSSES: ReadonlySet<ChartAxisCrosses> = new Set(["autoZero", "min", "max"]);

/**
 * Pull the axis crossing pin off `<c:crosses>` / `<c:crossesAt>`. The
 * OOXML schema (`CT_CatAx`, `CT_ValAx`, `CT_DateAx`, `CT_SerAx`) places
 * the two elements in an XSD choice — only one may legally appear at a
 * time per ECMA-376 Part 1, §21.2.2. The reader still handles both
 * appearing on the same axis (a malformed template) by preferring
 * `crossesAt` and dropping the redundant `crosses` value, mirroring the
 * writer's emit order.
 *
 * Returns:
 *   - `crosses`   — set when only `<c:crosses>` is present and the value
 *                   is a non-default token. The OOXML default `"autoZero"`
 *                   collapses to `undefined` so absence and the default
 *                   round-trip identically. Unknown tokens drop rather
 *                   than fabricate a value the writer would never emit.
 *   - `crossesAt` — set when `<c:crossesAt>` is present with a
 *                   parseable numeric `val`. Non-numeric / missing
 *                   `val` attributes drop to `undefined`. `0` is
 *                   preserved (it is a valid pin, distinct from the
 *                   `"autoZero"` default).
 */
function parseAxisCrosses(axis: XmlElement): {
  crosses?: ChartAxisCrosses;
  crossesAt?: number;
} {
  const crossesAtEl = findChild(axis, "crossesAt");
  if (crossesAtEl) {
    const raw = crossesAtEl.attrs.val;
    if (typeof raw === "string") {
      const trimmed = raw.trim();
      if (trimmed.length > 0) {
        const parsed = Number.parseFloat(trimmed);
        if (Number.isFinite(parsed)) {
          return { crossesAt: parsed };
        }
      }
    }
  }

  const crossesEl = findChild(axis, "crosses");
  if (!crossesEl) return {};
  const raw = crossesEl.attrs.val;
  if (typeof raw !== "string") return {};
  const value = raw.trim() as ChartAxisCrosses;
  if (!VALID_CROSSES.has(value)) return {};
  if (value === "autoZero") return {};
  return { crosses: value };
}

/** Recognized values of `<c:builtInUnit>` per the OOXML `ST_BuiltInUnit` enum. */
const VALID_DISP_UNITS: ReadonlySet<ChartAxisDispUnit> = new Set([
  "hundreds",
  "thousands",
  "tenThousands",
  "hundredThousands",
  "millions",
  "tenMillions",
  "hundredMillions",
  "billions",
  "trillions",
]);

/**
 * Read a value axis's `<c:dispUnits>` block. The element holds an
 * `xsd:choice` between `<c:builtInUnit val=".."/>` and
 * `<c:custUnit val=".."/>`, optionally followed by `<c:dispUnitsLbl>`.
 * The reader surfaces both: a recognized `<c:builtInUnit>` token lands
 * in `unit`, and a finite positive `<c:custUnit>` value lands in
 * `custUnit`. When both children are present (a malformed template,
 * since the schema's choice forbids it), `custUnit` wins and `unit`
 * drops — the writer applies the same precedence on emit, so the parsed
 * shape round-trips identically through {@link cloneChart}.
 *
 * Returns `undefined` when:
 *   - the axis declares no `<c:dispUnits>` at all,
 *   - `<c:dispUnits>` is present but neither child resolves to a
 *     valid value (missing children, malformed `val`, unknown
 *     `<c:builtInUnit>` token, non-positive / non-finite `<c:custUnit>`).
 *
 * `showLabel` is set `true` only when `<c:dispUnitsLbl>` is present
 * inside `<c:dispUnits>` (Excel paints its automatic annotation in
 * that case). Absence collapses to absence on the surfaced object so
 * a round-trip stays minimal.
 */
function parseAxisDispUnits(axis: XmlElement): ChartAxisDispUnits | undefined {
  const dispUnits = findChild(axis, "dispUnits");
  if (!dispUnits) return undefined;
  const out: ChartAxisDispUnits = {};
  // `<c:custUnit>` wins when both children are pinned — the OOXML
  // schema's `xsd:choice` forbids both, but a corrupt template may
  // declare them simultaneously. The writer mirrors this preference so
  // the round-trip stays consistent.
  const custUnit = findChild(dispUnits, "custUnit");
  if (custUnit) {
    const raw = custUnit.attrs.val;
    if (typeof raw === "string") {
      const parsed = Number.parseFloat(raw.trim());
      if (Number.isFinite(parsed) && parsed > 0) {
        out.custUnit = parsed;
      }
    }
  }
  if (out.custUnit === undefined) {
    const builtInUnit = findChild(dispUnits, "builtInUnit");
    if (builtInUnit) {
      const raw = builtInUnit.attrs.val;
      if (typeof raw === "string") {
        const trimmed = raw.trim() as ChartAxisDispUnit;
        if (VALID_DISP_UNITS.has(trimmed)) {
          out.unit = trimmed;
        }
      }
    }
  }
  if (out.unit === undefined && out.custUnit === undefined) return undefined;
  if (findChild(dispUnits, "dispUnitsLbl")) {
    out.showLabel = true;
  }
  return out;
}

/** Recognized values of `<c:crossBetween>` per the OOXML `ST_CrossBetween` enum. */
const VALID_CROSS_BETWEEN: ReadonlySet<ChartAxisCrossBetween> = new Set(["between", "midCat"]);

/**
 * Read a value axis's `<c:crossBetween val=".."/>`. The OOXML schema
 * places the element exclusively on `CT_ValAx` per ECMA-376 Part 1,
 * §21.2.2.10 — `<c:catAx>`, `<c:dateAx>`, and `<c:serAx>` reject it —
 * so the caller is expected to gate the parse on `axis.local === "valAx"`
 * before calling this helper.
 *
 * Returns `undefined` when:
 *   - the axis declares no `<c:crossBetween>` at all,
 *   - the `val` attribute is missing, empty, or not a string,
 *   - the `val` attribute is not one of the OOXML `ST_CrossBetween`
 *     tokens (`"between"` / `"midCat"`).
 *
 * Unknown tokens drop rather than fabricate a value the writer would
 * never emit — the caller cannot tell absence from a corrupt template
 * without the parser's help.
 */
function parseAxisCrossBetween(axis: XmlElement): ChartAxisCrossBetween | undefined {
  const el = findChild(axis, "crossBetween");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim() as ChartAxisCrossBetween;
  if (!VALID_CROSS_BETWEEN.has(trimmed)) return undefined;
  return trimmed;
}

/**
 * Read an axis's numeric scale block. The scale lives inside
 * `<c:scaling>`, with one optional child per pinned bound:
 *
 *   <c:scaling>
 *     <c:orientation val="minMax"/>
 *     <c:logBase val="10"/>
 *     <c:min val="0"/>
 *     <c:max val="100"/>
 *     <c:majorUnit val="20"/>
 *     <c:minorUnit val="5"/>
 *   </c:scaling>
 *
 * Returns `undefined` when none of the numeric children declare a
 * usable value — the orientation child alone (Excel's autoscale
 * baseline) does not surface a scale.
 */
function parseAxisScale(axis: XmlElement): ChartAxisScale | undefined {
  const out: ChartAxisScale = {};

  // <c:min>, <c:max>, and <c:logBase> live inside <c:scaling>; the
  // tick-spacing children <c:majorUnit> / <c:minorUnit> sit directly
  // under <c:catAx>/<c:valAx> per CT_CatAx / CT_ValAx in ECMA-376.
  const scaling = findChild(axis, "scaling");
  if (scaling) {
    const min = parseNumericChildVal(scaling, "min");
    if (min !== undefined) out.min = min;

    const max = parseNumericChildVal(scaling, "max");
    if (max !== undefined) out.max = max;

    const logBase = parseNumericChildVal(scaling, "logBase");
    if (logBase !== undefined) out.logBase = logBase;
  }

  const majorUnit = parseNumericChildVal(axis, "majorUnit");
  if (majorUnit !== undefined && majorUnit > 0) out.majorUnit = majorUnit;

  const minorUnit = parseNumericChildVal(axis, "minorUnit");
  if (minorUnit !== undefined && minorUnit > 0) out.minorUnit = minorUnit;

  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Read an axis's `<c:numFmt formatCode=".." sourceLinked=".."/>`.
 * Returns `undefined` when the element is absent or carries an empty
 * `formatCode`. `sourceLinked` is normalized to a boolean — `0`/`1`
 * and `"true"`/`"false"` are both accepted.
 */
function parseAxisNumberFormat(axis: XmlElement): ChartAxisNumberFormat | undefined {
  const numFmt = findChild(axis, "numFmt");
  if (!numFmt) return undefined;
  const formatCode = numFmt.attrs.formatCode;
  if (typeof formatCode !== "string" || formatCode.length === 0) return undefined;
  const out: ChartAxisNumberFormat = { formatCode };
  const sourceLinked = numFmt.attrs.sourceLinked;
  if (sourceLinked !== undefined && parseBoolAttr(sourceLinked) === true) {
    out.sourceLinked = true;
  }
  return out;
}

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

/**
 * Detect `<c:majorGridlines>` / `<c:minorGridlines>` children on an
 * axis element. The mere presence of either child element flips the
 * corresponding flag on — Excel allows but does not require nested
 * `<c:spPr>` styling, and the toggle survives even when the body is
 * empty.
 *
 * Returns `undefined` when neither element is present so the consumer
 * never sees a "{ major: false, minor: false }" record that
 * round-trips into a redundant write.
 */
function parseAxisGridlines(axis: XmlElement): ChartAxisGridlines | undefined {
  const major = findChild(axis, "majorGridlines") !== undefined;
  const minor = findChild(axis, "minorGridlines") !== undefined;
  if (!major && !minor) return undefined;
  const out: ChartAxisGridlines = {};
  if (major) out.major = true;
  if (minor) out.minor = true;
  return out;
}

/**
 * Read an axis's `<c:title>` text. Mirrors {@link parseTitle} but
 * scoped to a single axis element rather than the chart root.
 */
function parseAxisTitle(axis: XmlElement): string | undefined {
  const title = findChild(axis, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (rich) {
    const parts: string[] = [];
    collectTextRuns(rich, parts);
    const joined = parts.join("").trim();
    return joined.length > 0 ? joined : undefined;
  }
  const strRef = findChild(tx, "strRef");
  if (strRef) {
    const cache = findChild(strRef, "strCache");
    if (cache) {
      for (const pt of childElements(cache)) {
        if (pt.local !== "pt") continue;
        const v = findChild(pt, "v");
        if (v) {
          const text = elementText(v).trim();
          if (text.length > 0) return text;
        }
      }
    }
  }
  return undefined;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title>`
 * off an axis element. Returns the rotation in whole degrees (range
 * `-90..90`).
 *
 * The OOXML default `0` (and absence of the element / attribute) all
 * collapse to `undefined` so absence and the default round-trip
 * identically through {@link cloneChart}. Out-of-range values clamp to
 * the nearest endpoint of the `-90..90` band Excel's UI exposes;
 * non-finite (`NaN`, `Infinity`) inputs drop to `undefined`.
 *
 * The lookup is scoped to the title's `<c:rich>` body so a stray
 * `<a:bodyPr>` elsewhere on the axis (e.g. the tick-label `<c:txPr>`
 * surfaced by {@link parseAxisLabelRotation}) cannot leak in. Returns
 * `undefined` when the axis omits `<c:title>` entirely or when the
 * title is a `<c:strRef>` (formula reference) with no `<c:rich>` body.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape per
 * the OOXML schema. Mirrors the chart-level title rotation
 * {@link parseTitleRotation} so a parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleRotation.
 */
function parseAxisTitleRotation(axis: XmlElement): number | undefined {
  const title = findChild(axis, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  const bodyPr = findChild(rich, "bodyPr");
  if (!bodyPr) return undefined;
  const raw = bodyPr.attrs.rot;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  // Convert from 60000ths of a degree to whole degrees.
  const degrees = Math.round(parsed / TXPR_ROT_PER_DEGREE);
  if (degrees === 0) return undefined;
  if (degrees < LABEL_ROTATION_MIN_DEG) return LABEL_ROTATION_MIN_DEG;
  if (degrees > LABEL_ROTATION_MAX_DEG) return LABEL_ROTATION_MAX_DEG;
  return degrees;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off an axis element. Returns the
 * font size in points (range `1..400`).
 *
 * Mirrors {@link parseTitleFontSize} for axis titles — same 0.5pt
 * half-step granularity, same `1..400`pt band, same drop-on-out-of-
 * range / non-numeric / absence semantics. The lookup is scoped to
 * the axis title's `<c:rich>` body so a stray `<a:defRPr>` elsewhere
 * on the axis (e.g. on the tick-label `<c:txPr>` surfaced by
 * {@link parseAxisLabelRotation}) cannot leak in.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the size
 * from in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape
 * per the OOXML schema. Mirrors the chart-level title size
 * {@link parseTitleFontSize} so a parsed value slots straight into
 * the writer-side {@link SheetChart.axes}.x.axisTitleFontSize.
 */
function parseAxisTitleFontSize(axis: XmlElement): number | undefined {
  const title = findChild(axis, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph font size. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.sz;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  // Convert from 100ths of a point to points, rounding to the nearest
  // 0.5pt to match the granularity Excel's UI exposes. Mirrors the
  // chart-level `parseTitleFontSize` half-step normalisation.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2);
  const points = halfSteps / 2;
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined;
  return points;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off an axis element. Returns the
 * bold flag.
 *
 * Mirrors {@link parseTitleBold} for axis titles — same
 * canonical-slot pair, same drop-on-default-`false` semantics. The
 * OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * default `false` collapses to `undefined` so absence and `b="0"`
 * round-trip identically — only an explicit `b="1"` surfaces `true`.
 * Unknown / malformed `b` tokens drop to `undefined` rather than
 * fabricate a value the writer would never emit.
 *
 * The lookup is scoped to the axis title's `<c:rich>` body so a stray
 * `<a:defRPr>` elsewhere on the axis (e.g. on the tick-label
 * `<c:txPr>` surfaced by {@link parseAxisLabelRotation}) cannot leak
 * in. Returns `undefined` when the axis omits `<c:title>` entirely or
 * when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to host the flag in
 * either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape per
 * the OOXML schema. Mirrors the chart-level title bold
 * {@link parseTitleBold} so a parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleBold.
 */
function parseAxisTitleBold(axis: XmlElement): boolean | undefined {
  const title = findChild(axis, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph bold flag. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const parsed = parseBoolAttr(defRPr.attrs.b);
  // The OOXML default `false` collapses to `undefined` so absence and
  // `b="0"` round-trip identically through the writer — only an
  // explicit `b="1"` surfaces `true`.
  if (parsed === true) return true;
  return undefined;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off an axis element. Returns
 * `true` only when the OOXML attribute is the literal truthy spelling
 * (`"1"` / `"true"`); the OOXML default `false` collapses to
 * `undefined` so absence and `i="0"` round-trip identically through
 * the writer.
 *
 * Mirrors {@link parseTitleItalic} for axis titles — same canonical-
 * slot pair (`<a:defRPr>` carries the default-paragraph italic flag,
 * which the writer keeps in sync with the literal run's `<a:rPr>` so
 * the reader only needs to consult one of the two slots), same
 * drop-on-default semantics. The lookup is scoped to the axis title's
 * `<c:rich>` body so a stray `<a:defRPr>` elsewhere on the axis (e.g.
 * on the tick-label `<c:txPr>`) cannot leak in.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the flag
 * from in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape
 * per the OOXML schema. Mirrors the chart-level title italic
 * {@link parseTitleItalic} so a parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleItalic.
 */
function parseAxisTitleItalic(axis: XmlElement): boolean | undefined {
  const title = findChild(axis, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph italic flag. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const parsed = parseBoolAttr(defRPr.attrs.i);
  // The OOXML default `false` collapses to `undefined` so absence and
  // `i="0"` round-trip identically through the writer — only an
  // explicit `i="1"` surfaces `true`.
  if (parsed === true) return true;
  return undefined;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:rich></c:tx></c:title>` off an axis element. Returns the axis
 * title's sRGB font color as a 6-character uppercase hex string.
 *
 * Mirrors {@link parseTitleColor} for axis titles — same canonical-
 * slot pair (`<a:defRPr>` carries the default-paragraph fill, which
 * the writer keeps in sync with the literal run's `<a:rPr>` so the
 * reader only needs to consult one of the two slots), same drop-on-
 * non-sRGB semantics. The lookup is scoped to the axis title's
 * `<c:rich>` body so a stray `<a:solidFill>` elsewhere on the axis
 * (e.g. on a tick-label `<c:txPr>` block or a `<c:spPr>` series
 * fill) cannot leak in.
 *
 * The OOXML `<a:srgbClr val=".."/>` is the literal sRGB triple Excel
 * lands on the axis title's default-paragraph properties when the
 * user picks a custom font color. Theme references (`<a:schemeClr>`),
 * `<a:hslClr>`, `<a:sysClr>`, and `<a:prstClr>` all collapse to
 * `undefined` — only the literal RGB triple round-trips losslessly
 * through {@link writeChart}. Malformed `val` tokens (wrong length,
 * non-hex characters) likewise drop to `undefined` rather than
 * fabricate a value the writer would round-trip into a malformed
 * `<a:srgbClr>`.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the fill
 * from in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape
 * per the OOXML schema. Mirrors the chart-level title color
 * {@link parseTitleColor} so a parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleColor.
 */
function parseAxisTitleColor(axis: XmlElement): string | undefined {
  const title = findChild(axis, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr>` is the OOXML path
  // Excel writes for the default-paragraph font color. The reader walks
  // the canonical chain and bails on the first missing link so a
  // malformed `<c:rich>` surfaces as absence rather than a fabricated
  // value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const solidFill = findChild(defRPr, "solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "srgbClr");
  if (!srgbClr) return undefined;
  const raw = srgbClr.attrs.val;
  return normalizeRgbHex(raw);
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` off an axis element.
 * Returns the axis-title strikethrough flag.
 *
 * The OOXML `strike` attribute is the `ST_TextStrikeType` enum on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
 * three values: `"noStrike"` (the OOXML application default),
 * `"sngStrike"` (single line, the value Excel's UI checkbox emits),
 * and `"dblStrike"` (double line, a non-UI variant). The reader
 * surfaces only the UI-default `"sngStrike"` as `true`; `"noStrike"`,
 * absence, and the non-UI `"dblStrike"` all collapse to `undefined` —
 * the writer emits only `"sngStrike"`, so reporting `"dblStrike"` as
 * `true` would silently downgrade the choice to a single line on
 * round-trip. Unknown / malformed `strike` tokens likewise drop to
 * `undefined`.
 *
 * Mirrors {@link parseTitleStrike} for axis titles — same canonical-
 * slot pair (`<a:defRPr>` carries the default-paragraph strike flag,
 * which the writer keeps in sync with the literal run's `<a:rPr>` so
 * the reader only needs to consult one of the two slots), same
 * drop-on-default semantics. The lookup is scoped to the axis title's
 * `<c:rich>` body so a stray `<a:defRPr>` elsewhere on the axis (e.g.
 * on the tick-label `<c:txPr>`) cannot leak in.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely or
 * when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the flag from
 * in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape per
 * the OOXML schema. Mirrors the chart-level title strike
 * {@link parseTitleStrike} so a parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleStrike.
 */
function parseAxisTitleStrike(axis: XmlElement): boolean | undefined {
  const title = findChild(axis, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph strikethrough flag. The reader walks the
  // canonical chain and bails on the first missing link so a malformed
  // `<c:rich>` surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.strike;
  // Only the UI-default `"sngStrike"` surfaces as `true`. The OOXML
  // application default `"noStrike"` and the non-UI `"dblStrike"` both
  // collapse to `undefined` so absence and the OOXML default round-trip
  // identically through the writer; the writer emits only `"sngStrike"`,
  // so reporting `"dblStrike"` here would silently downgrade the choice
  // on round-trip.
  if (raw === "sngStrike") return true;
  return undefined;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` off an axis element.
 * Returns the axis-title underline flag.
 *
 * The OOXML `u` attribute is the `ST_TextUnderlineType` enum on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
 * eighteen values; Excel's UI exposes only `"sng"` (single line — the
 * default underline checkbox) and `"dbl"` (double line). The reader
 * surfaces only the UI-default `"sng"` as `true`; `"none"` (the OOXML
 * application default), absence, the non-UI `"dbl"` variant, and the
 * sixteen exotic tokens (`"words"`, `"heavy"`, `"dotted"`,
 * `"dottedHeavy"`, `"dash"`, `"dashHeavy"`, `"dashLong"`,
 * `"dashLongHeavy"`, `"dotDash"`, `"dotDashHeavy"`, `"dotDotDash"`,
 * `"dotDotDashHeavy"`, `"wavy"`, `"wavyHeavy"`, `"wavyDbl"`) all
 * collapse to `undefined` — the writer emits only `"sng"`, so
 * reporting any non-single underline as `true` would silently
 * downgrade the choice to a single line on round-trip. Unknown /
 * malformed `u` tokens likewise drop to `undefined`.
 *
 * Mirrors {@link parseTitleUnderline} for axis titles — same
 * canonical-slot pair (`<a:defRPr>` carries the default-paragraph
 * underline flag, which the writer keeps in sync with the literal
 * run's `<a:rPr>` so the reader only needs to consult one of the two
 * slots), same drop-on-default semantics. The lookup is scoped to the
 * axis title's `<c:rich>` body so a stray `<a:defRPr>` elsewhere on
 * the axis (e.g. on the tick-label `<c:txPr>`) cannot leak in.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely or
 * when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the flag from
 * in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape per
 * the OOXML schema. Mirrors the chart-level title underline
 * {@link parseTitleUnderline} so a parsed value slots straight into
 * the writer-side {@link SheetChart.axes}.x.axisTitleUnderline.
 */
function parseAxisTitleUnderline(axis: XmlElement): boolean | undefined {
  const title = findChild(axis, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph underline flag. The reader walks the canonical
  // chain and bails on the first missing link so a malformed
  // `<c:rich>` surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.u;
  // Only the UI-default `"sng"` surfaces as `true`. The OOXML
  // application default `"none"`, the non-UI `"dbl"` variant, and
  // every exotic token (`"words"`, `"heavy"`, `"dotted"`, etc.) all
  // collapse to `undefined` so absence and the OOXML default
  // round-trip identically through the writer; the writer emits only
  // `"sng"`, so reporting a non-single underline here would silently
  // downgrade the choice on round-trip.
  if (raw === "sng") return true;
  return undefined;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx></c:title>`
 * off the axis. Returns the typeface string the title was authored
 * with.
 *
 * The OOXML `<a:latin>` element carries the typeface name on
 * `CT_TextFont` (ECMA-376 Part 1, §21.1.2.3.7). The reader trims
 * surrounding whitespace and reports the trimmed typeface; empty /
 * whitespace-only `typeface` attributes and missing `<a:latin>`
 * elements both collapse to `undefined` so absence and the empty
 * form round-trip identically through the writer. Non-string
 * `typeface` tokens (defensive — the XML parser only ever surfaces
 * strings) likewise drop to `undefined`.
 *
 * Mirrors the chart-level title typeface {@link parseTitleFontFamily} —
 * same canonical-slot pair (`<a:defRPr>` carries the default-paragraph
 * typeface, which the writer keeps in sync with the literal run's
 * `<a:rPr>` so the reader only needs to consult one of the two
 * slots). The lookup is scoped to the axis title's `<c:rich>` body
 * so a stray `<a:latin>` elsewhere on the axis (e.g. on the
 * tick-label `<c:txPr>`) cannot leak in.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the
 * typeface from in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape
 * per the OOXML schema. The parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleFontFamily.
 */
function parseAxisTitleFontFamily(axis: XmlElement): string | undefined {
  const title = findChild(axis, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr><a:latin>` is the OOXML path Excel writes
  // for the default-paragraph typeface. The reader walks the
  // canonical chain and bails on the first missing link so a
  // malformed `<c:rich>` surfaces as absence rather than a fabricated
  // value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const latin = findChild(defRPr, "latin");
  if (!latin) return undefined;
  const raw = latin.attrs.typeface;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  return trimmed;
}

/**
 * Pull `<c:title><c:overlay val=".."/></c:title>` off an axis
 * element. Returns the axis-title overlay flag.
 *
 * The OOXML default `false` (the title reserves its own slot adjacent
 * to the axis, no overlap with the plot area) collapses to
 * `undefined` so absence and `<c:overlay val="0"/>` round-trip
 * identically through {@link cloneChart} — only an explicit
 * `<c:overlay val="1"/>` surfaces `true`.
 *
 * Returns `undefined` whenever the axis omits the `<c:title>`
 * element — there is no overlay slot to surface in that case. The
 * element is a sibling of `<c:tx>` inside `<c:title>` per the
 * CT_Title schema, so the lookup is scoped to direct title children
 * (a stray `<c:overlay>` elsewhere on the axis or chart cannot leak
 * in). Mirrors the chart-level title-overlay reader so a parsed
 * value flows straight back into the writer-side
 * {@link SheetChart.axes}.x.axisTitleOverlay.
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` /
 * `"0"` / `"false"`); unknown values and missing `val` attributes
 * drop to `undefined` rather than fabricate a flag Excel would not
 * emit.
 */
function parseAxisTitleOverlay(axis: XmlElement): boolean | undefined {
  const title = findChild(axis, "title");
  if (!title) return undefined;
  const overlay = findChild(title, "overlay");
  if (!overlay) return undefined;
  const raw = overlay.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "1":
    case "true":
      return true;
    case "0":
    case "false":
      // OOXML default — collapse to undefined for symmetry with the
      // writer's `axisTitleOverlay` field.
      return undefined;
    default:
      return undefined;
  }
}

// ── Series ────────────────────────────────────────────────────────

/**
 * Pull the metadata fields {@link ChartSeriesInfo} surfaces out of a
 * single `<c:ser>` element. Missing pieces (no name, no categories,
 * literal numbers instead of a range) are simply omitted.
 */
function parseSeries(ser: XmlElement, kind: ChartKind, index: number): ChartSeriesInfo {
  const out: ChartSeriesInfo = { kind, index };

  const name = parseSeriesName(ser);
  if (name !== undefined) out.name = name;

  // Numeric values land in <c:val> for most chart types; scatter and
  // bubble use <c:yVal> instead.
  const valuesRef = formulaText(findChild(ser, "val")) ?? formulaText(findChild(ser, "yVal"));
  if (valuesRef !== undefined) out.valuesRef = valuesRef;

  // Categories live in <c:cat> for category-axis charts and in
  // <c:xVal> for scatter/bubble.
  const catRef = formulaText(findChild(ser, "cat")) ?? formulaText(findChild(ser, "xVal"));
  if (catRef !== undefined) out.categoriesRef = catRef;

  const color = parseSeriesColor(ser);
  if (color !== undefined) out.color = color;

  const dLbls = findChild(ser, "dLbls");
  if (dLbls) {
    const parsed = parseDataLabels(dLbls);
    if (parsed) out.dataLabels = parsed;
  }

  // `<c:smooth>` lives on `CT_LineSer` and `CT_ScatterSer` only — every
  // other chart family rejects the element. Surface it just for those
  // two kinds so a corrupt template carrying `<c:smooth>` on a bar/pie
  // series does not silently flip a flag that the writer would never
  // emit anyway.
  if (kind === "line" || kind === "line3D" || kind === "scatter") {
    const smooth = parseSmooth(ser);
    if (smooth !== undefined) out.smooth = smooth;

    // Stroke (dash + width) lives in `<c:spPr><a:ln>`. The same
    // schema-only-on-line/scatter rule applies — bar / pie / area
    // never paint a connecting line, so surfacing a stroke field
    // there would mislead a clone consumer about what the chart
    // actually renders.
    const stroke = parseSeriesStroke(ser);
    if (stroke !== undefined) out.stroke = stroke;

    // `<c:marker>` mirrors the same scope — CT_LineSer / CT_ScatterSer
    // only. Skip the element on every other family so a stray
    // `<c:marker>` on a bar / pie / area template does not surface a
    // setting that the writer would never emit anyway.
    const marker = parseMarker(ser);
    if (marker !== undefined) out.marker = marker;
  }

  // `<c:invertIfNegative>` lives on `CT_BarSer` / `CT_Bar3DSer` only —
  // every other chart family rejects the element. Surface the flag
  // just for those two kinds so a corrupt template carrying
  // `<c:invertIfNegative>` on a line/pie/area/scatter series does not
  // silently flip a flag that the writer would never emit anyway.
  if (kind === "bar" || kind === "bar3D") {
    const invertIfNegative = parseInvertIfNegative(ser);
    if (invertIfNegative !== undefined) out.invertIfNegative = invertIfNegative;
  }

  // `<c:explosion>` lives on `CT_PieSer` only — the OOXML schema
  // shares the type across every pie-family chart (`<c:pieChart>`,
  // `<c:pie3DChart>`, `<c:doughnutChart>`, `<c:ofPieChart>`) so
  // surface the value for any of those kinds. A stray element on a
  // bar / line / area / scatter template is dropped rather than
  // surfaced — the writer would never emit it back anyway.
  if (kind === "pie" || kind === "pie3D" || kind === "doughnut" || kind === "ofPie") {
    const explosion = parseExplosion(ser);
    if (explosion !== undefined) out.explosion = explosion;
  }

  return out;
}

/**
 * Pull `<c:smooth val=".."/>` off a series element. Returns `undefined`
 * when the attribute is absent, malformed, or carries the OOXML default
 * `false` — absence and `false` round-trip identically through the
 * writer's elision logic, so collapsing them keeps the parsed shape
 * minimal.
 */
function parseSmooth(ser: XmlElement): boolean | undefined {
  const el = findChild(ser, "smooth");
  if (!el) return undefined;
  const v = readBoolAttr(el);
  if (v !== true) return undefined;
  return true;
}

/**
 * Pull `<c:invertIfNegative val=".."/>` off a bar/column series
 * element. Returns `undefined` when the attribute is absent,
 * malformed, or carries the OOXML default `false` — absence and
 * `false` round-trip identically through the writer's elision logic,
 * so collapsing them keeps the parsed shape minimal.
 */
function parseInvertIfNegative(ser: XmlElement): boolean | undefined {
  const el = findChild(ser, "invertIfNegative");
  if (!el) return undefined;
  const v = readBoolAttr(el);
  if (v !== true) return undefined;
  return true;
}

/**
 * Pull `<c:explosion val=".."/>` off a pie / doughnut series element.
 * The element's `val` attribute is `xsd:unsignedInt` per the OOXML
 * schema (CT_UnsignedInt) — the slice is pulled away from the chart
 * center by `val` percent of the radius. Returns `undefined` when the
 * attribute is absent, malformed, negative, or carries the OOXML
 * default `0` — absence and `0` round-trip identically through the
 * writer's elision logic, so collapsing them keeps the parsed shape
 * minimal. Non-integer input rounds to the nearest integer for parity
 * with the writer (Excel's UI accepts integer percentages only).
 */
function parseExplosion(ser: XmlElement): number | undefined {
  const el = findChild(ser, "explosion");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const n = Number.parseFloat(raw);
  if (!Number.isFinite(n) || n < 0) return undefined;
  const rounded = Math.round(n);
  if (rounded === 0) return undefined;
  return rounded;
}

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

/**
 * Pull `<c:marker>` off a line / scatter series. Returns `undefined`
 * when the marker block is absent or carries no meaningful settings —
 * an empty `<c:marker/>` element collapses identically to absence
 * through the writer's elision logic, so omitting it keeps the parsed
 * shape minimal.
 *
 * Field semantics mirror {@link ChartMarker}: an unknown `<c:symbol>`
 * value is dropped (rather than surfaced), `<c:size>` outside the
 * 2..72 band is clamped, and the fill / outline colors come from
 * `<c:spPr><a:solidFill>` and `<c:spPr><a:ln><a:solidFill>`
 * respectively.
 */
function parseMarker(ser: XmlElement): ChartMarker | undefined {
  const el = findChild(ser, "marker");
  if (!el) return undefined;

  const out: ChartMarker = {};

  const sym = findChild(el, "symbol");
  if (sym) {
    const v = sym.attrs.val;
    if (typeof v === "string" && VALID_MARKER_SYMBOLS.has(v as ChartMarkerSymbol)) {
      out.symbol = v as ChartMarkerSymbol;
    }
  }

  const sizeEl = findChild(el, "size");
  if (sizeEl) {
    const v = sizeEl.attrs.val;
    if (typeof v === "string") {
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n)) {
        // OOXML ST_MarkerSize is `xsd:unsignedByte` constrained to
        // 2..72; clamp anything outside that band on the way in so a
        // template with an out-of-range value still round-trips.
        if (n < 2) out.size = 2;
        else if (n > 72) out.size = 72;
        else out.size = n;
      }
    }
  }

  const spPr = findChild(el, "spPr");
  if (spPr) {
    const fill = findChild(spPr, "solidFill");
    if (fill) {
      const srgb = findChild(fill, "srgbClr");
      const v = srgb?.attrs.val;
      if (typeof v === "string") {
        const hex = v.replace(/^#/, "").toUpperCase();
        if (/^[0-9A-F]{6}$/.test(hex)) out.fill = hex;
      }
    }
    const ln = findChild(spPr, "ln");
    if (ln) {
      const lnFill = findChild(ln, "solidFill");
      if (lnFill) {
        const srgb = findChild(lnFill, "srgbClr");
        const v = srgb?.attrs.val;
        if (typeof v === "string") {
          const hex = v.replace(/^#/, "").toUpperCase();
          if (/^[0-9A-F]{6}$/.test(hex)) out.line = hex;
        }
      }
    }
  }

  if (
    out.symbol === undefined &&
    out.size === undefined &&
    out.fill === undefined &&
    out.line === undefined
  ) {
    return undefined;
  }
  return out;
}

// ── Data Labels ───────────────────────────────────────────────────

const VALID_DLBL_POSITIONS: ReadonlySet<ChartDataLabelPosition> = new Set([
  "t",
  "b",
  "l",
  "r",
  "ctr",
  "inEnd",
  "inBase",
  "outEnd",
  "bestFit",
]);

/**
 * Read a `<c:dLbls>` block. Returns `undefined` when the block is
 * empty or only contains a `<c:delete val="1">` (which suppresses
 * labels rather than describing them). All toggles default to `false`
 * when the matching `<c:show*>` element is absent.
 */
function parseDataLabels(el: XmlElement): ChartDataLabelsInfo | undefined {
  // <c:delete val="1"> at the root of <c:dLbls> means "suppress for
  // this scope". We don't surface a dataLabels record for that case —
  // it's the absence of labels, not a configuration.
  const deleteEl = findChild(el, "delete");
  if (deleteEl && readBoolAttr(deleteEl) === true) return undefined;

  const out: ChartDataLabelsInfo = {};

  const pos = findChild(el, "dLblPos");
  if (pos) {
    const val = pos.attrs.val;
    if (typeof val === "string" && VALID_DLBL_POSITIONS.has(val as ChartDataLabelPosition)) {
      out.position = val as ChartDataLabelPosition;
    }
  }

  // `<c:showLegendKey val=".."/>` mirrors Excel's "Format Data Labels
  // -> Legend Key" checkbox. The OOXML default is `false`, so absence
  // and an explicit `val="0"` collapse to `undefined` — only an
  // explicit `val="1"` (or `"true"`) surfaces `true`. Same shape as the
  // other `show*` toggles so the parsed record can be fed straight back
  // into {@link cloneChart}.
  const showLeg = findChild(el, "showLegendKey");
  if (showLeg && readBoolAttr(showLeg) === true) out.showLegendKey = true;

  const showVal = findChild(el, "showVal");
  if (showVal && readBoolAttr(showVal) === true) out.showValue = true;

  const showCat = findChild(el, "showCatName");
  if (showCat && readBoolAttr(showCat) === true) out.showCategoryName = true;

  const showSer = findChild(el, "showSerName");
  if (showSer && readBoolAttr(showSer) === true) out.showSeriesName = true;

  const showPct = findChild(el, "showPercent");
  if (showPct && readBoolAttr(showPct) === true) out.showPercent = true;

  const sep = findChild(el, "separator");
  if (sep) {
    const text = elementText(sep);
    if (text.length > 0) out.separator = text;
  }

  // `<c:numFmt formatCode=".." sourceLinked=".."/>` mirrors Excel's
  // "Format Data Labels -> Number" panel — pinning a custom number
  // format on the rendered label values. Same shape as the axis-side
  // `<c:numFmt>` so the parsed value can be fed straight back into
  // {@link cloneChart}. The element sits early in the CT_DLbls
  // sequence (after the optional `<c:dLbl>` instances), so the lookup
  // is scoped to direct `<c:dLbls>` children — a `<c:numFmt>` nested
  // inside a per-point `<c:dLbl>` does not leak into the block-level
  // record.
  const numFmt = parseDataLabelsNumberFormat(el);
  if (numFmt) out.numberFormat = numFmt;

  // `<c:showLeaderLines val=".."/>` mirrors Excel's "Format Data
  // Labels -> Show Leader Lines" checkbox. The OOXML default is
  // `true` (Excel paints leader lines on every label that gets pushed
  // outside its slice), so absence and `val="1"` collapse to
  // `undefined` — only an explicit `val="0"` (or `"false"`) surfaces
  // `false`. Mirrors how the writer-side scope guard treats the
  // element: only meaningful on pie / doughnut, but the parser is
  // permissive (the OOXML schema scopes the element to `EG_DLbls` for
  // `CT_PieChart` / `CT_DoughnutChart`, but a templated chart whose
  // type element ends up coerced should still surface the source's
  // intent so the cloned model stays accurate).
  // Only the literal OOXML falsy spellings (`"0"` / `"false"`) flip the
  // toggle — unknown / missing val tokens collapse to `undefined` rather
  // than surface a `false` the writer would round-trip into a non-default
  // `<c:showLeaderLines val="0"/>` Excel never authored.
  const showLeader = findChild(el, "showLeaderLines");
  if (showLeader) {
    const v = showLeader.attrs.val;
    if (typeof v === "string" && (v === "0" || v.toLowerCase() === "false")) {
      out.showLeaderLines = false;
    }
  }

  // `<c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr>`
  // — data-label font size pinned via Excel's "Format Data Labels ->
  // Font -> Size" knob. The OOXML attribute is in 100ths of a point on
  // `CT_TextCharacterProperties`' `sz` slot; the reader divides by
  // 100 at parse time so the surfaced value matches what the user
  // sees in the UI. Out-of-range / malformed tokens collapse to
  // `undefined` so absence and a malformed source value round-trip
  // identically through `cloneChart`.
  const fontSize = parseDataLabelsFontSize(el);
  if (fontSize !== undefined) out.fontSize = fontSize;

  // `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val=".."/>
  // </a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>` — data-label
  // font color pinned via Excel's "Format Data Labels -> Font -> Font
  // color" picker. Theme references (`<a:schemeClr>`) and malformed
  // `val` tokens collapse to `undefined` since only the literal RGB
  // triple round-trips losslessly through `writeChart`.
  const fontColor = parseDataLabelsFontColor(el);
  if (fontColor !== undefined) out.fontColor = fontColor;

  // `<c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr>` —
  // data-label bold flag pinned via Excel's "Format Data Labels ->
  // Font -> Bold" toggle. Only an explicit `b="1"` surfaces `true`;
  // the OOXML default `b="0"` collapses to `undefined` so absence
  // and `b="0"` round-trip identically through `cloneChart`.
  const bold = parseDataLabelsBold(el);
  if (bold !== undefined) out.bold = bold;

  // `<c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr>` —
  // data-label italic flag pinned via Excel's "Format Data Labels ->
  // Font -> Italic" toggle. Only an explicit `i="1"` surfaces `true`;
  // the OOXML default `i="0"` collapses to `undefined` so absence
  // and `i="0"` round-trip identically through `cloneChart`.
  const italic = parseDataLabelsItalic(el);
  if (italic !== undefined) out.italic = italic;

  // `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>` —
  // data-label underline flag pinned via Excel's "Format Data Labels
  // -> Font -> Underline" toggle. Only `u="sng"` (Excel's UI variant
  // — single underline) surfaces `true`; the OOXML default `"none"`
  // (and every other ST_TextUnderlineType variant) collapse to
  // `undefined` so absence and `u="none"` round-trip identically
  // through `cloneChart`.
  const underline = parseDataLabelsUnderline(el);
  if (underline !== undefined) out.underline = underline;

  // `<c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p>
  // </c:txPr>` — data-label strikethrough flag pinned via Excel's
  // "Format Data Labels -> Font -> Strikethrough" toggle. Only
  // `strike="sngStrike"` (Excel's UI variant — single line) surfaces
  // `true`; the OOXML default `"noStrike"` and the non-UI variant
  // `"dblStrike"` (and any malformed token) collapse to `undefined`
  // so absence and `"noStrike"` round-trip identically through
  // `cloneChart`.
  const strikethrough = parseDataLabelsStrikethrough(el);
  if (strikethrough !== undefined) out.strikethrough = strikethrough;

  // `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
  // </a:pPr></a:p></c:txPr>` — data-label font family pinned via
  // Excel's "Format Data Labels -> Font -> Font" picker. Empty /
  // whitespace-only `typeface` attributes and missing `<a:latin>`
  // elements both collapse to `undefined` so absence and the empty
  // form round-trip identically through the writer.
  const fontFamily = parseDataLabelsFontFamily(el);
  if (fontFamily !== undefined) out.fontFamily = fontFamily;

  // Empty record is meaningless to a consumer — collapse to undefined.
  if (
    out.position === undefined &&
    !out.showValue &&
    !out.showCategoryName &&
    !out.showSeriesName &&
    !out.showPercent &&
    !out.showLegendKey &&
    out.separator === undefined &&
    out.numberFormat === undefined &&
    out.showLeaderLines === undefined &&
    out.fontSize === undefined &&
    out.fontColor === undefined &&
    out.bold === undefined &&
    out.italic === undefined &&
    out.underline === undefined &&
    out.strikethrough === undefined &&
    out.fontFamily === undefined
  ) {
    return undefined;
  }
  return out;
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
 * </c:txPr></c:dLbls>` off a data-labels block. Returns the font size
 * in points (`1..400`), or `undefined` when the element is absent /
 * the chain is malformed at any link / the surfaced value is out of
 * the supported range.
 *
 * The OOXML attribute is in 100ths of a point on
 * `CT_TextCharacterProperties`' `sz` slot (ECMA-376 Part 1,
 * §21.1.2.3.7); the reader divides by 100 at parse time so the
 * surfaced value matches what the user sees in Excel's UI. Mirrors
 * the chart-title / axis-title / axis tick-label / legend font-size
 * readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
function parseDataLabelsFontSize(dLbls: XmlElement): number | undefined {
  const txPr = findChild(dLbls, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.sz;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  // Convert from 100ths of a point to points, rounding to the nearest
  // 0.5pt to match the granularity Excel's UI exposes. Mirrors the
  // chart-title / axis-title / tick-label / legend sibling parsers
  // exactly so a parsed value flows through every typography slot
  // without bookkeeping the units.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2);
  const points = halfSteps / 2;
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined;
  return points;
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:txPr></c:dLbls>` off a data-labels block. Returns the
 * 6-character uppercase hex string when the parser walks the full
 * chain and lands on an `<a:srgbClr val="RRGGBB"/>`. Theme references
 * (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`, and `<a:prstClr>`
 * all collapse to `undefined` — only the literal RGB triple
 * round-trips losslessly through {@link writeChart}. Malformed `val`
 * tokens (wrong length, non-hex characters) likewise drop to
 * `undefined` rather than fabricate a value the writer would
 * round-trip into a malformed `<a:srgbClr>`.
 *
 * Returns `undefined` whenever the data-labels block omits `<c:txPr>`
 * entirely or the canonical chain is malformed at any link. Mirrors
 * the chart-title / axis-title / axis tick-label / legend color
 * readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
function parseDataLabelsFontColor(dLbls: XmlElement): string | undefined {
  const txPr = findChild(dLbls, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const solidFill = findChild(defRPr, "solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "srgbClr");
  if (!srgbClr) return undefined;
  return normalizeRgbHex(srgbClr.attrs.val);
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
 * </c:txPr></c:dLbls>` off a data-labels block. Returns the bold
 * flag.
 *
 * The OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties`. Only an explicit `b="1"` (or
 * `"true"`) surfaces `true`; the OOXML default `0` (and absence /
 * malformed tokens) collapses to `undefined` so absence and the
 * default round-trip identically through `cloneChart`. Mirrors the
 * chart-title / axis-title / axis tick-label / legend bold readers
 * exactly so a parsed value slots straight back into the writer's
 * emit path.
 */
function parseDataLabelsBold(dLbls: XmlElement): boolean | undefined {
  const txPr = findChild(dLbls, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.b;
  if (raw === "1" || raw === "true") return true;
  return undefined;
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
 * </c:txPr></c:dLbls>` off a data-labels block. Returns the italic
 * flag.
 *
 * The OOXML `i` attribute is the `xsd:boolean` italic flag on
 * `CT_TextCharacterProperties`. Only an explicit `i="1"` (or
 * `"true"`) surfaces `true`; the OOXML default `0` (and absence /
 * malformed tokens) collapses to `undefined` so absence and the
 * default round-trip identically through `cloneChart`. Mirrors the
 * chart-title / axis-title / axis tick-label / legend italic readers
 * exactly so a parsed value slots straight back into the writer's
 * emit path.
 */
function parseDataLabelsItalic(dLbls: XmlElement): boolean | undefined {
  const txPr = findChild(dLbls, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.i;
  if (raw === "1" || raw === "true") return true;
  return undefined;
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p>
 * </c:txPr></c:dLbls>` off a data-labels block. Returns the underline
 * flag.
 *
 * The OOXML `u` attribute is the `ST_TextUnderlineType` enumeration
 * on `CT_TextCharacterProperties`. Only `u="sng"` (Excel's UI variant
 * — single underline) surfaces `true`; the OOXML default `"none"`
 * (and every other variant the schema allows — `"dbl"`, `"heavy"`,
 * `"dotted"`, `"dotDash"`, `"wavy"`, etc.) collapse to `undefined` so
 * absence and `u="none"` round-trip identically through `cloneChart`.
 * Reporting any non-`"sng"` underline as `true` would silently
 * downgrade the choice to a single line on round-trip; the writer
 * emits only `u="sng"` / `u="none"`, matching the boolean shape the
 * UI exposes. Mirrors the chart-title / axis-title / axis tick-label
 * / legend underline readers exactly so a parsed value slots straight
 * back into the writer's emit path.
 */
function parseDataLabelsUnderline(dLbls: XmlElement): boolean | undefined {
  const txPr = findChild(dLbls, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.u;
  if (raw === "sng") return true;
  return undefined;
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
 * </a:p></c:txPr></c:dLbls>` off a data-labels block. Returns the
 * strikethrough flag.
 *
 * The OOXML `strike` attribute is the `ST_TextStrikeType`
 * enumeration on `CT_TextCharacterProperties`. Only
 * `strike="sngStrike"` (Excel's UI variant — single line) surfaces
 * `true`; the OOXML default `"noStrike"` and the non-UI variant
 * `"dblStrike"` (and any malformed token) collapse to `undefined` so
 * absence and `"noStrike"` round-trip identically through
 * `cloneChart`. Reporting `"dblStrike"` as `true` would silently
 * downgrade the choice to a single line on round-trip; the writer
 * emits only `"sngStrike"`, matching the boolean shape the UI
 * exposes. Mirrors the chart-title / axis-title / axis tick-label /
 * legend strikethrough readers exactly so a parsed value slots
 * straight back into the writer's emit path.
 */
function parseDataLabelsStrikethrough(dLbls: XmlElement): boolean | undefined {
  const txPr = findChild(dLbls, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.strike;
  if (raw === "sngStrike") return true;
  return undefined;
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbls>` off a
 * data-labels block. Returns the typeface string the labels were
 * authored with.
 *
 * The OOXML `<a:latin>` element carries the typeface name on
 * `CT_TextFont` (ECMA-376 Part 1, §21.1.2.3.7). The reader trims
 * surrounding whitespace and reports the trimmed typeface; empty /
 * whitespace-only `typeface` attributes and missing `<a:latin>`
 * elements both collapse to `undefined` so absence and the empty
 * form round-trip identically through the writer. Non-string
 * `typeface` tokens (defensive — the XML parser only ever surfaces
 * strings) likewise drop to `undefined`.
 *
 * Returns `undefined` whenever the data-labels block omits `<c:txPr>`
 * entirely or the canonical `<a:p><a:pPr><a:defRPr><a:latin>` chain
 * is malformed at any link. Mirrors the chart-title / axis-title /
 * axis tick-label / legend font family readers exactly so a parsed
 * value slots straight back into the writer's emit path.
 */
function parseDataLabelsFontFamily(dLbls: XmlElement): string | undefined {
  const txPr = findChild(dLbls, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const latin = findChild(defRPr, "latin");
  if (!latin) return undefined;
  const raw = latin.attrs.typeface;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  return trimmed;
}

/**
 * Pull `<c:numFmt formatCode=".." sourceLinked=".."/>` off a
 * `<c:dLbls>` block. Returns `undefined` when the element is absent or
 * when `formatCode` is missing / empty (the OOXML schema requires the
 * attribute on every emitted `<c:numFmt>` so a fabricated empty record
 * cannot round-trip cleanly).
 *
 * `sourceLinked` accepts the same OOXML truthy / falsy spellings the
 * other boolean attributes do (`"1"` / `"true"` / `"0"` / `"false"`);
 * absence and the OOXML default `"0"` collapse to `undefined` so the
 * parsed shape stays minimal — only an explicit `"1"` / `"true"`
 * surfaces `true`. Mirrors how the axis-side numFmt parser shapes its
 * output.
 */
function parseDataLabelsNumberFormat(el: XmlElement): ChartAxisNumberFormat | undefined {
  const numFmt = findChild(el, "numFmt");
  if (!numFmt) return undefined;
  const formatCode = numFmt.attrs.formatCode;
  if (typeof formatCode !== "string" || formatCode.length === 0) return undefined;
  const out: ChartAxisNumberFormat = { formatCode };
  const sourceLinked = numFmt.attrs.sourceLinked;
  if (typeof sourceLinked === "string") {
    if (sourceLinked === "1" || sourceLinked.toLowerCase() === "true") {
      out.sourceLinked = true;
    }
  }
  return out;
}

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
function parseSeriesName(ser: XmlElement): string | undefined {
  const tx = findChild(ser, "tx");
  if (!tx) return undefined;
  const literal = findChild(tx, "v");
  if (literal) {
    const text = elementText(literal).trim();
    if (text.length > 0) return text;
  }
  const strRef = findChild(tx, "strRef");
  if (strRef) {
    const cache = findChild(strRef, "strCache");
    if (cache) {
      for (const pt of childElements(cache)) {
        if (pt.local !== "pt") continue;
        const v = findChild(pt, "v");
        if (v) {
          const text = elementText(v).trim();
          if (text.length > 0) return text;
        }
      }
    }
    // Fall back to the formula reference itself when no cached value.
    const f = formulaText(strRef);
    if (f) return f;
  }
  return undefined;
}

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
function parseSeriesColor(ser: XmlElement): string | undefined {
  const spPr = findChild(ser, "spPr");
  if (!spPr) return undefined;
  const fill = findChild(spPr, "solidFill");
  if (!fill) return undefined;
  const srgb = findChild(fill, "srgbClr");
  if (!srgb) return undefined;
  const val = srgb.attrs.val;
  if (typeof val !== "string") return undefined;
  const normalized = val.replace(/^#/, "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(normalized) ? normalized : undefined;
}

// ── Stroke ────────────────────────────────────────────────────────

const VALID_DASH_STYLES: ReadonlySet<ChartLineDashStyle> = new Set([
  "solid",
  "dot",
  "dash",
  "lgDash",
  "dashDot",
  "lgDashDot",
  "lgDashDotDot",
  "sysDash",
  "sysDot",
  "sysDashDot",
  "sysDashDotDot",
]);

const STROKE_WIDTH_MIN_PT = 0.25;
const STROKE_WIDTH_MAX_PT = 13.5;
const EMU_PER_PT = 12700;

/**
 * Pull `<c:spPr><a:ln>` off a series and surface its dash + width as
 * a {@link ChartLineStroke}. Returns `undefined` when the block is
 * absent or carries no meaningful settings — an empty `<a:ln/>`
 * collapses identically to absence through the writer's elision
 * logic, so omitting it keeps the parsed shape minimal.
 *
 * `<a:ln>` also nests the line color (`<a:solidFill>`) which mirrors
 * the series fill — parseSeriesColor already surfaces that as
 * {@link ChartSeriesInfo.color}, so the stroke object intentionally
 * does not duplicate the field.
 */
function parseSeriesStroke(ser: XmlElement): ChartLineStroke | undefined {
  const spPr = findChild(ser, "spPr");
  if (!spPr) return undefined;
  const ln = findChild(spPr, "ln");
  if (!ln) return undefined;

  const out: ChartLineStroke = {};

  // Stroke width is on the `w` attribute of `<a:ln>` (EMU). Convert
  // back to points and clamp to the band Excel's UI exposes so a
  // template carrying an exotic width still round-trips through the
  // writer's clamp.
  const wAttr = ln.attrs.w;
  if (typeof wAttr === "string") {
    const emu = Number.parseFloat(wAttr);
    if (Number.isFinite(emu) && emu > 0) {
      // Snap to the 0.25 pt grid Excel's UI exposes (Math.round(x * 4) / 4).
      const pt = Math.round((emu / EMU_PER_PT) * 4) / 4;
      if (pt < STROKE_WIDTH_MIN_PT) out.width = STROKE_WIDTH_MIN_PT;
      else if (pt > STROKE_WIDTH_MAX_PT) out.width = STROKE_WIDTH_MAX_PT;
      else out.width = pt;
    }
  }

  // Dash style is `<a:prstDash val="..."/>` inside `<a:ln>`.
  const dashEl = findChild(ln, "prstDash");
  if (dashEl) {
    const v = dashEl.attrs.val;
    if (typeof v === "string" && VALID_DASH_STYLES.has(v as ChartLineDashStyle)) {
      out.dash = v as ChartLineDashStyle;
    }
  }

  if (out.dash === undefined && out.width === undefined) return undefined;
  return out;
}

// ── Legend ────────────────────────────────────────────────────────

/**
 * Map `<c:legend><c:legendPos val=".."/></c:legend>` to the writer-side
 * {@link ChartLegendPosition}. Returns `false` when `<c:delete val="1"/>`
 * is present (Excel's "no legend" state); returns `undefined` when the
 * chart has no `<c:legend>` element at all.
 */
function parseLegend(chartEl: XmlElement): false | ChartLegendPosition | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;

  // <c:delete val="1"/> means the chart explicitly suppresses the
  // legend. Some Excel versions emit just an empty `<c:legend/>`
  // followed by `<c:overlay/>` even when the legend is hidden, but
  // `<c:delete val="1">` is the canonical "no legend" marker.
  const del = findChild(legend, "delete");
  if (del && readBoolVal(del.attrs.val) === true) return false;

  const pos = findChild(legend, "legendPos");
  if (!pos) {
    // A legend element without legendPos is valid OOXML (Excel falls
    // back to "right"). Surface "right" so the cloned chart preserves
    // the visible-legend state.
    return "right";
  }
  const val = pos.attrs.val;
  if (typeof val !== "string") return "right";
  switch (val) {
    case "t":
      return "top";
    case "b":
      return "bottom";
    case "l":
      return "left";
    case "r":
      return "right";
    case "tr":
      return "topRight";
    default:
      // Unknown legendPos values are dropped rather than fabricated.
      return undefined;
  }
}

/**
 * Pull `<c:legend><c:overlay val=".."/></c:legend>` off the chart. The
 * OOXML default `false` (the legend reserves its own slot, no overlap
 * with the plot area) collapses to `undefined` so absence and
 * `<c:overlay val="0"/>` round-trip identically through
 * {@link cloneChart} — only an explicit `<c:overlay val="1"/>` surfaces
 * `true`.
 *
 * The caller is expected to confirm a visible legend exists before
 * invoking this — `<c:overlay>` only renders when the legend is part of
 * the chart, so reading it from a chart that hides or omits the legend
 * would surface a flag with no on-screen effect.
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` / `"0"`
 * / `"false"`); unknown values and missing `val` attributes drop to
 * `undefined` rather than fabricate a flag Excel would not emit.
 */
function parseLegendOverlay(chartEl: XmlElement): boolean | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const overlay = findChild(legend, "overlay");
  if (!overlay) return undefined;
  const raw = overlay.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "1":
    case "true":
      return true;
    case "0":
    case "false":
      // OOXML default — collapse to undefined for symmetry with the
      // writer's `legendOverlay` field.
      return undefined;
    default:
      return undefined;
  }
}

/**
 * Pull `<c:legend><c:legendEntry>` overrides off the chart. Returns
 * `undefined` when the chart declares no entries so the field is
 * elided entirely on a clean parse — absence and an empty array
 * round-trip identically through {@link cloneChart} (the writer skips
 * emission when the resolved list is empty).
 *
 * Each entry is admitted only when its `<c:idx val=".."/>` selector
 * parses to a non-negative integer (matches the OOXML
 * `xsd:unsignedInt` schema). Entries without an `<c:idx>` child or with
 * a malformed `val` attribute are dropped rather than surface a
 * fabricated index. The `<c:delete>` flag accepts the OOXML truthy /
 * falsy spellings (`"1"` / `"true"` / `"0"` / `"false"`); absence
 * collapses to `false` (the OOXML default — the entry renders).
 *
 * The caller is expected to confirm a visible legend exists before
 * invoking this — `<c:legendEntry>` only renders inside `<c:legend>`,
 * so reading from a chart that hides or omits the legend would surface
 * overrides with no on-screen effect.
 *
 * Duplicate `idx` values keep the first occurrence — Excel's renderer
 * treats later duplicates as overrides on the same series, but the
 * writer's `resolveLegendEntries` deduplicates with last-wins semantics
 * to give clone-through callers a way to override without manually
 * pruning. Reading "first wins" pairs naturally with that behaviour:
 * a parsed list re-emits cleanly, and an explicit clone override that
 * appends an entry still beats the parsed value.
 */
function parseLegendEntries(chartEl: XmlElement): ChartLegendEntry[] | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;

  const seen = new Set<number>();
  const out: ChartLegendEntry[] = [];
  for (const child of childElements(legend)) {
    if (child.local !== "legendEntry") continue;
    const idxEl = findChild(child, "idx");
    if (!idxEl) continue;
    const raw = idxEl.attrs.val;
    if (typeof raw !== "string") continue;
    const idx = Number.parseInt(raw, 10);
    if (!Number.isFinite(idx) || idx < 0) continue;
    if (seen.has(idx)) continue;
    seen.add(idx);

    const deleteEl = findChild(child, "delete");
    const deleteFlag = deleteEl !== undefined ? readBoolVal(deleteEl.attrs.val) === true : false;
    out.push({ idx, delete: deleteFlag });
  }

  return out.length > 0 ? out : undefined;
}

/**
 * Pull `<c:legend><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
 * </c:txPr></c:legend>` off the chart. Returns the font size in points
 * (range `1..400`).
 *
 * The OOXML `sz` attribute is in 100ths of a point — the reader
 * converts to points and rounds to the nearest 0.5pt (Excel's UI
 * exposes the same 0.5pt granularity). Absence of the element /
 * attribute and out-of-range / non-numeric / non-finite values all
 * collapse to `undefined` so a fresh chart and a chart that pinned an
 * out-of-range size both round-trip to the writer's "skip the size
 * attribute" path.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>` element
 * — there is no `<c:txPr>` slot to surface the size from in that case.
 * The `<a:defRPr>` lives inside `<c:txPr><a:p><a:pPr>` per the
 * CT_TextBody schema (the default-paragraph properties on the
 * legend's text-body's first paragraph); the lookup is scoped to that
 * path so a stray `<a:defRPr>` elsewhere in the chart (e.g. on the
 * chart title or an axis) cannot leak in. Mirrors the chart-title /
 * axis-title / axis tick-label readers exactly so a parsed value slots
 * straight back into the writer's emit path.
 */
function parseLegendFontSize(chartEl: XmlElement): number | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.sz;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  // Convert from 100ths of a point to points, rounding to the nearest
  // 0.5pt to match the granularity Excel's UI exposes. `Math.round`
  // on `2 * (parsed / 100)` and dividing by 2 gives a clean half-step
  // band that mirrors the writer's emit-time normalization. The
  // chart-title / axis-title / tick-label sibling parsers use the
  // identical conversion so a parsed value flows through every
  // typography slot without bookkeeping the units.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2);
  const points = halfSteps / 2;
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined;
  return points;
}

/**
 * Pull `<c:legend><c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
 * </c:txPr></c:legend>` off the chart. Returns the bold flag.
 *
 * The OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * OOXML default `false` collapses to `undefined` so absence and
 * `b="0"` round-trip identically — only an explicit `b="1"` surfaces
 * `true`. Unknown / malformed `b` tokens drop to `undefined` rather
 * than fabricate a value the writer would never emit.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element — there is no `<c:txPr>` slot to surface the flag from in
 * that case. The `<a:defRPr>` lives inside `<c:txPr><a:p><a:pPr>` per
 * the CT_TextBody schema (the default-paragraph properties on the
 * legend's text-body's first paragraph); the lookup is scoped to that
 * path so a stray `<a:defRPr>` elsewhere in the chart (e.g. on the
 * chart title or an axis) cannot leak in. Mirrors the chart-title /
 * axis-title / axis tick-label readers exactly so a parsed value
 * slots straight back into the writer's emit path.
 */
function parseLegendBold(chartEl: XmlElement): boolean | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.b;
  // OOXML `xsd:boolean` accepts `"1"` / `"true"` (truthy) and `"0"` /
  // `"false"` (falsy). Truthy spellings surface `true`; falsy
  // spellings collapse to `undefined` so the OOXML default and an
  // explicit `b="0"` round-trip identically through `cloneChart`.
  // Unknown / missing `b` tokens drop to `undefined` for the same
  // reason — never fabricate a flag Excel would not emit.
  if (raw === "1" || raw === "true") return true;
  return undefined;
}

/**
 * Pull `<c:legend><c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
 * </c:txPr></c:legend>` off the chart. Returns the italic flag.
 *
 * The OOXML `i` attribute is the `xsd:boolean` italic flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * OOXML default `false` collapses to `undefined` so absence and
 * `i="0"` round-trip identically — only an explicit `i="1"` surfaces
 * `true`. Unknown / malformed `i` tokens drop to `undefined` rather
 * than fabricate a value the writer would never emit.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element — there is no `<c:txPr>` slot to surface the flag from in
 * that case. Mirrors the chart-title / axis-title / axis tick-label
 * readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
function parseLegendItalic(chartEl: XmlElement): boolean | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.i;
  if (raw === "1" || raw === "true") return true;
  return undefined;
}

/**
 * Pull `<c:legend><c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p>
 * </c:txPr></c:legend>` off the chart. Returns the underline flag.
 *
 * The OOXML `u` attribute is the `ST_TextUnderlineType` enumeration
 * on `CT_TextCharacterProperties`. Only `u="sng"` (Excel's UI variant
 * — single underline) surfaces `true`; the OOXML default `"none"`
 * (and every other variant the schema allows — `"dbl"`, `"heavy"`,
 * `"dotted"`, `"dotDash"`, `"wavy"`, etc.) collapse to `undefined` so
 * absence and `u="none"` round-trip identically through `cloneChart`.
 * Reporting any non-`"sng"` underline as `true` would silently
 * downgrade the choice to a single line on round-trip; the writer
 * emits only `u="sng"` / `u="none"`, matching the boolean shape the
 * UI exposes.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element. Mirrors the chart-title / axis-title / axis tick-label
 * underline readers exactly so a parsed value slots straight back
 * into the writer's emit path.
 */
function parseLegendUnderline(chartEl: XmlElement): boolean | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.u;
  if (raw === "sng") return true;
  return undefined;
}

/**
 * Pull `<c:legend><c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
 * </a:p></c:txPr></c:legend>` off the chart. Returns the strikethrough
 * flag.
 *
 * The OOXML `strike` attribute is the `ST_TextStrikeType` enumeration
 * on `CT_TextCharacterProperties` — `"noStrike"` (default),
 * `"sngStrike"` (single line, Excel's UI variant), `"dblStrike"`
 * (double line, non-UI). Only `strike="sngStrike"` surfaces `true`;
 * the OOXML default `"noStrike"` and the double-line variant both
 * collapse to `undefined` so absence and `strike="noStrike"`
 * round-trip identically through `cloneChart`. Reporting `"dblStrike"`
 * as `true` would silently downgrade the choice to a single line on
 * re-emit; the writer emits only `"sngStrike"`, matching the boolean
 * shape the UI exposes.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element. Mirrors the chart-title / axis-title / axis tick-label
 * strikethrough readers exactly so a parsed value slots straight back
 * into the writer's emit path.
 */
function parseLegendStrikethrough(chartEl: XmlElement): boolean | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.strike;
  if (raw === "sngStrike") return true;
  return undefined;
}

/**
 * Pull the legend font color off the canonical
 * `<c:legend><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr
 * val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
 * </c:legend>` chain Excel writes when the user pins a custom font
 * color on the legend.
 *
 * Returns the 6-character uppercase hex string when the parser walks
 * the full chain and lands on an `<a:srgbClr val="RRGGBB"/>`. Theme
 * references (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`, and
 * `<a:prstClr>` all collapse to `undefined` — only the literal RGB
 * triple round-trips losslessly through {@link writeChart}. Malformed
 * `val` tokens (wrong length, non-hex characters) likewise drop to
 * `undefined` rather than fabricate a value the writer would round-
 * trip into a malformed `<a:srgbClr>`.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element or the canonical `<c:txPr><a:p><a:pPr><a:defRPr>
 * <a:solidFill><a:srgbClr>` chain is malformed at any link. Mirrors
 * the chart-title / axis-title / axis tick-label color readers
 * exactly so a parsed value slots straight back into the writer's
 * emit path.
 */
function parseLegendFontColor(chartEl: XmlElement): string | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const solidFill = findChild(defRPr, "solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "srgbClr");
  if (!srgbClr) return undefined;
  return normalizeRgbHex(srgbClr.attrs.val);
}

/**
 * Pull the legend font family off the canonical
 * `<c:legend><c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/>
 * </a:defRPr></a:pPr></a:p></c:txPr></c:legend>` chain Excel writes
 * when the user pins a typeface on the legend.
 *
 * The OOXML `<a:latin>` element carries the typeface name on
 * `CT_TextFont` (ECMA-376 Part 1, §21.1.2.3.7). The reader trims
 * surrounding whitespace and reports the trimmed typeface; empty /
 * whitespace-only `typeface` attributes and missing `<a:latin>`
 * elements both collapse to `undefined` so absence and the empty
 * form round-trip identically through the writer. Non-string
 * `typeface` tokens (defensive — the XML parser only ever surfaces
 * strings) likewise drop to `undefined`.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element or the canonical `<c:txPr><a:p><a:pPr><a:defRPr><a:latin>`
 * chain is malformed at any link. Mirrors the chart-title /
 * axis-title / axis tick-label font family readers exactly so a
 * parsed value slots straight back into the writer's emit path.
 */
function parseLegendFontFamily(chartEl: XmlElement): string | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const latin = findChild(defRPr, "latin");
  if (!latin) return undefined;
  const raw = latin.attrs.typeface;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  return trimmed;
}

/**
 * Pull `<c:legend><c:layout><c:manualLayout>` off the chart. Reflects
 * Excel's "Format Legend -> Position -> Custom" knob — the `(x, y)`
 * anchor and `(w, h)` size of the legend block as fractions of the
 * chart frame in the `0..1` band.
 *
 * The OOXML schema (`CT_ManualLayout`, ECMA-376 Part 1, §21.2.2.115)
 * exposes optional `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` children whose
 * `val` attributes carry an `xsd:double`. The reader admits the
 * coordinate only when `val` parses to a finite number in the `0..1`
 * band; out-of-range / non-finite / non-numeric tokens drop to
 * `undefined` on the matching axis so absence and a malformed token
 * round-trip identically through {@link cloneChart}.
 *
 * Both `<c:xMode val="edge"/>` (absolute fraction of the chart frame)
 * and `<c:xMode val="factor"/>` (delta from auto-layout) are accepted
 * — the reader surfaces the same `ChartManualLayout` shape regardless,
 * since the writer always normalizes to `"edge"` on emit (Excel itself
 * emits the absolute form when the user drags an element to a custom
 * position).
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>` /
 * `<c:layout>` / `<c:manualLayout>` chain at any link, or when every
 * coordinate dropped on normalization — the field is omitted entirely
 * on a clean parse so absence and an empty layout round-trip identically
 * through the writer.
 */
function parseLegendLayout(chartEl: XmlElement): ChartManualLayout | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const layout = findChild(legend, "layout");
  if (!layout) return undefined;
  const manual = findChild(layout, "manualLayout");
  if (!manual) return undefined;

  const out: ChartManualLayout = {};
  const x = readLayoutCoordinate(findChild(manual, "x"));
  if (x !== undefined) out.x = x;
  const y = readLayoutCoordinate(findChild(manual, "y"));
  if (y !== undefined) out.y = y;
  const w = readLayoutCoordinate(findChild(manual, "w"));
  if (w !== undefined) out.w = w;
  const h = readLayoutCoordinate(findChild(manual, "h"));
  if (h !== undefined) out.h = h;

  if (out.x === undefined && out.y === undefined && out.w === undefined && out.h === undefined) {
    return undefined;
  }
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
function parsePlotAreaLayout(plotArea: XmlElement): ChartManualLayout | undefined {
  const layout = findChild(plotArea, "layout");
  if (!layout) return undefined;
  const manual = findChild(layout, "manualLayout");
  if (!manual) return undefined;

  const out: ChartManualLayout = {};
  const x = readLayoutCoordinate(findChild(manual, "x"));
  if (x !== undefined) out.x = x;
  const y = readLayoutCoordinate(findChild(manual, "y"));
  if (y !== undefined) out.y = y;
  const w = readLayoutCoordinate(findChild(manual, "w"));
  if (w !== undefined) out.w = w;
  const h = readLayoutCoordinate(findChild(manual, "h"));
  if (h !== undefined) out.h = h;

  if (out.x === undefined && out.y === undefined && out.w === undefined && out.h === undefined) {
    return undefined;
  }
  return out;
}

/**
 * Parse a single `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` element off a
 * `<c:manualLayout>` block. Returns the `val` attribute as a finite
 * number in the `0..1` band; everything else (missing element, missing
 * attribute, non-numeric / non-finite / out-of-range token) collapses
 * to `undefined` so the matching axis on {@link parseLegendLayout}'s
 * return value is omitted.
 */
function readLayoutCoordinate(el: XmlElement | undefined): number | undefined {
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number(trimmed);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < 0 || parsed > 1) return undefined;
  return parsed;
}

/**
 * Pull `<c:title><c:overlay val=".."/></c:title>` off the chart. The
 * OOXML default `false` (the title reserves its own slot above the plot
 * area, no overlap) collapses to `undefined` so absence and
 * `<c:overlay val="0"/>` round-trip identically through
 * {@link cloneChart} — only an explicit `<c:overlay val="1"/>` surfaces
 * `true`.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no overlay slot to surface in that case. The element is a
 * sibling of `<c:tx>` inside `<c:title>` per the CT_Title schema, so the
 * lookup is scoped to direct title children.
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` / `"0"`
 * / `"false"`); unknown values and missing `val` attributes drop to
 * `undefined` rather than fabricate a flag Excel would not emit.
 */
function parseTitleOverlay(chartEl: XmlElement): boolean | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const overlay = findChild(title, "overlay");
  if (!overlay) return undefined;
  const raw = overlay.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "1":
    case "true":
      return true;
    case "0":
    case "false":
      // OOXML default — collapse to undefined for symmetry with the
      // writer's `titleOverlay` field.
      return undefined;
    default:
      return undefined;
  }
}

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
 * Pull `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx>
 * </c:title>` off the chart. Returns the rotation in whole degrees
 * (range `-90..90`).
 *
 * The OOXML default `0` (and absence of the `<a:bodyPr>` element /
 * `rot` attribute) all collapse to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart}.
 * Non-integer / non-numeric / out-of-range values clamp to the nearest
 * endpoint of the `-90..90` band Excel's UI exposes; non-finite
 * (`NaN`, `Infinity`) inputs drop to `undefined`.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no rotation slot to surface in that case. The
 * `<a:bodyPr>` lives inside `<c:tx><c:rich>` per the CT_Title schema
 * (the rich-text body's body-properties); the lookup is scoped to that
 * path so a stray `<a:bodyPr>` elsewhere in the chart cannot leak in.
 */
function parseTitleRotation(chartEl: XmlElement): number | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  const bodyPr = findChild(rich, "bodyPr");
  if (!bodyPr) return undefined;
  const raw = bodyPr.attrs.rot;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  // Convert from 60000ths of a degree to whole degrees.
  const degrees = Math.round(parsed / TITLE_ROT_PER_DEGREE);
  if (degrees === 0) return undefined;
  if (degrees < TITLE_ROTATION_MIN_DEG) return TITLE_ROTATION_MIN_DEG;
  if (degrees > TITLE_ROTATION_MAX_DEG) return TITLE_ROTATION_MAX_DEG;
  return degrees;
}

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

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off the chart. Returns the font
 * size in points (range `1..400`).
 *
 * The OOXML `sz` attribute is in 100ths of a point — the reader
 * converts to points and rounds to the nearest 0.5pt (Excel's UI
 * exposes the same 0.5pt granularity). Absence of the element /
 * attribute and out-of-range / non-numeric / non-finite values all
 * collapse to `undefined` so a fresh chart and a chart that pinned an
 * out-of-range size both round-trip to the writer's "skip the size
 * attribute" path.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no `<a:p>` slot to surface the size from in that case —
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the
 * default-paragraph properties on the rich-text body's first
 * paragraph); the lookup is scoped to that path so a stray
 * `<a:defRPr>` elsewhere in the chart (e.g. on an axis title or a
 * data-labels block) cannot leak in.
 */
function parseTitleFontSize(chartEl: XmlElement): number | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph font size. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.sz;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  // Convert from 100ths of a point to points, rounding to the nearest
  // 0.5pt to match the granularity Excel's UI exposes. `Math.round`
  // on `2 * (parsed / 100)` and dividing by 2 gives a clean half-step
  // band that mirrors the writer's emit-time normalization.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2);
  const points = halfSteps / 2;
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined;
  return points;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off the chart. Returns the bold
 * flag.
 *
 * The OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * OOXML default `false` collapses to `undefined` so absence and
 * `b="0"` round-trip identically — only an explicit `b="1"` surfaces
 * `true`. Unknown / malformed `b` tokens drop to `undefined` rather
 * than fabricate a value the writer would never emit.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no `<a:p>` slot to surface the flag from in that case —
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the
 * default-paragraph properties on the rich-text body's first
 * paragraph); the lookup is scoped to that path so a stray
 * `<a:defRPr>` elsewhere in the chart (e.g. on an axis title or a
 * data-labels block) cannot leak in.
 */
function parseTitleBold(chartEl: XmlElement): boolean | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph bold flag. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const parsed = parseBoolAttr(defRPr.attrs.b);
  // The OOXML default `false` collapses to `undefined` so absence and
  // `b="0"` round-trip identically through the writer — only an
  // explicit `b="1"` surfaces `true`.
  if (parsed === true) return true;
  return undefined;
}

// ── Title Italic ─────────────────────────────────────────────────────

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off the chart. Returns the italic
 * flag.
 *
 * The OOXML `i` attribute is the `xsd:boolean` italic flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * OOXML default `false` collapses to `undefined` so absence and
 * `i="0"` round-trip identically — only an explicit `i="1"` surfaces
 * `true`. Unknown / malformed `i` tokens drop to `undefined` rather
 * than fabricate a value the writer would never emit.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no `<a:p>` slot to surface the flag from in that case —
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the
 * default-paragraph properties on the rich-text body's first
 * paragraph); the lookup is scoped to that path so a stray
 * `<a:defRPr>` elsewhere in the chart (e.g. on an axis title or a
 * data-labels block) cannot leak in.
 */
function parseTitleItalic(chartEl: XmlElement): boolean | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph italic flag. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const parsed = parseBoolAttr(defRPr.attrs.i);
  // The OOXML default `false` collapses to `undefined` so absence and
  // `i="0"` round-trip identically through the writer — only an
  // explicit `i="1"` surfaces `true`.
  if (parsed === true) return true;
  return undefined;
}

// ── Title Color ─────────────────────────────────────────────────────

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:rich></c:tx></c:title>` off the chart. Returns the title's
 * sRGB font color as a 6-character uppercase hex string.
 *
 * The OOXML `<a:srgbClr val=".."/>` is the literal sRGB triple Excel
 * lands on the title's default-paragraph properties when the user
 * picks a custom font color. Theme references (`<a:schemeClr>`),
 * `<a:hslClr>`, `<a:sysClr>`, and `<a:prstClr>` all collapse to
 * `undefined` — only the literal RGB triple round-trips losslessly
 * through {@link writeChart}. Malformed `val` tokens (wrong length,
 * non-hex characters) likewise drop to `undefined` rather than
 * fabricate a value the writer would round-trip into a malformed
 * `<a:srgbClr>`.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>`
 * element — there is no `<a:p>` slot to surface the fill from in
 * that case — or when the title is a `<c:strRef>` (formula
 * reference) with no `<c:rich>` body. The `<a:solidFill>` lives
 * inside `<c:tx><c:rich><a:p><a:pPr><a:defRPr>` per the CT_Title
 * schema (the default-paragraph properties on the rich-text body's
 * first paragraph); the lookup is scoped to that path so a stray
 * `<a:solidFill>` elsewhere in the chart (e.g. on an axis title or
 * a data-labels block) cannot leak in.
 */
function parseTitleColor(chartEl: XmlElement): string | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr>` is the OOXML path
  // Excel writes for the default-paragraph font color. The reader
  // walks the canonical chain and bails on the first missing link so
  // a malformed `<c:rich>` surfaces as absence rather than a
  // fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const solidFill = findChild(defRPr, "solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "srgbClr");
  if (!srgbClr) return undefined;
  const raw = srgbClr.attrs.val;
  return normalizeRgbHex(raw);
}

/**
 * Normalize an sRGB hex token. Returns the 6-character uppercase form
 * when the input is a valid 6-character hex string (with or without a
 * leading `#`), or `undefined` for any malformed input — wrong length,
 * non-hex characters, alpha-channel forms, or non-string tokens.
 */
function normalizeRgbHex(raw: unknown): string | undefined {
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const hex = trimmed.startsWith("#") ? trimmed.slice(1) : trimmed;
  if (hex.length !== 6) return undefined;
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return undefined;
  return hex.toUpperCase();
}

// ── Title Strike ────────────────────────────────────────────────────

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` off the chart. Returns
 * the strikethrough flag.
 *
 * The OOXML `strike` attribute is the `ST_TextStrikeType` enum on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
 * three values: `"noStrike"` (the OOXML application default),
 * `"sngStrike"` (single line, the value Excel's UI checkbox emits),
 * and `"dblStrike"` (double line, a non-UI variant). The reader
 * surfaces only the UI-default `"sngStrike"` as `true`; `"noStrike"`,
 * absence, and the non-UI `"dblStrike"` all collapse to `undefined` —
 * the writer emits only `"sngStrike"`, so reporting `"dblStrike"` as
 * `true` would silently downgrade the choice to a single line on
 * round-trip. Unknown / malformed `strike` tokens likewise drop to
 * `undefined`.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>`
 * element — there is no `<a:p>` slot to surface the flag from in
 * that case — or when the title is a `<c:strRef>` (formula
 * reference) with no `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the default-
 * paragraph properties on the rich-text body's first paragraph); the
 * lookup is scoped to that path so a stray `<a:defRPr>` elsewhere in
 * the chart (e.g. on an axis title or a data-labels block) cannot
 * leak in.
 */
function parseTitleStrike(chartEl: XmlElement): boolean | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph strikethrough flag. The reader walks the
  // canonical chain and bails on the first missing link so a
  // malformed `<c:rich>` surfaces as absence rather than a fabricated
  // value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.strike;
  // Only the UI-default `"sngStrike"` surfaces as `true`. The OOXML
  // application default `"noStrike"` and the non-UI `"dblStrike"` both
  // collapse to `undefined` so absence and the OOXML default round-trip
  // identically through the writer; the writer emits only `"sngStrike"`,
  // so reporting `"dblStrike"` here would silently downgrade the choice
  // on round-trip.
  if (raw === "sngStrike") return true;
  return undefined;
}

// ── Title Underline ─────────────────────────────────────────────────

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` off the chart. Returns
 * the underline flag.
 *
 * The OOXML `u` attribute is the `ST_TextUnderlineType` enum on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
 * eighteen values; Excel's UI exposes only `"sng"` (single line —
 * the default underline checkbox) and `"dbl"` (double line). The
 * reader surfaces only the UI-default `"sng"` as `true`; `"none"`
 * (the OOXML application default), absence, the non-UI `"dbl"`
 * variant, and the sixteen exotic tokens (`"words"`, `"heavy"`,
 * `"dotted"`, `"dottedHeavy"`, `"dash"`, `"dashHeavy"`, `"dashLong"`,
 * `"dashLongHeavy"`, `"dotDash"`, `"dotDashHeavy"`, `"dotDotDash"`,
 * `"dotDotDashHeavy"`, `"wavy"`, `"wavyHeavy"`, `"wavyDbl"`) all
 * collapse to `undefined` — the writer emits only `"sng"`, so
 * reporting any non-single underline as `true` would silently
 * downgrade the choice to a single line on round-trip. Unknown /
 * malformed `u` tokens likewise drop to `undefined`.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>`
 * element — there is no `<a:p>` slot to surface the flag from in
 * that case — or when the title is a `<c:strRef>` (formula
 * reference) with no `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the default-
 * paragraph properties on the rich-text body's first paragraph); the
 * lookup is scoped to that path so a stray `<a:defRPr>` elsewhere in
 * the chart (e.g. on an axis title or a data-labels block) cannot
 * leak in.
 */
function parseTitleUnderline(chartEl: XmlElement): boolean | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph underline flag. The reader walks the canonical
  // chain and bails on the first missing link so a malformed
  // `<c:rich>` surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.u;
  // Only the UI-default `"sng"` surfaces as `true`. The OOXML
  // application default `"none"`, the non-UI `"dbl"` variant, and
  // every exotic token (`"words"`, `"heavy"`, `"dotted"`, etc.) all
  // collapse to `undefined` so absence and the OOXML default
  // round-trip identically through the writer; the writer emits only
  // `"sng"`, so reporting a non-single underline here would silently
  // downgrade the choice on round-trip.
  if (raw === "sng") return true;
  return undefined;
}

// ── Title Font Family ───────────────────────────────────────────────

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx></c:title>`
 * off the chart. Returns the typeface string the title was authored
 * with.
 *
 * The OOXML `<a:latin>` element carries the typeface name on
 * `CT_TextFont` (ECMA-376 Part 1, §21.1.2.3.7). The reader trims
 * surrounding whitespace and reports the trimmed typeface; empty /
 * whitespace-only `typeface` attributes and missing `<a:latin>`
 * elements both collapse to `undefined` so absence and the empty
 * form round-trip identically through the writer. Non-string
 * `typeface` tokens (defensive — the XML parser only ever surfaces
 * strings) likewise drop to `undefined`.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no `<a:p>` slot to surface the typeface from in that case
 * — or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the default-
 * paragraph properties on the rich-text body's first paragraph); the
 * lookup is scoped to that path so a stray `<a:latin>` elsewhere in
 * the chart (e.g. on an axis title or a data-labels block) cannot
 * leak in.
 */
function parseTitleFontFamily(chartEl: XmlElement): string | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr><a:latin>` is the OOXML path Excel writes
  // for the default-paragraph typeface. The reader walks the
  // canonical chain and bails on the first missing link so a
  // malformed `<c:rich>` surfaces as absence rather than a fabricated
  // value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const latin = findChild(defRPr, "latin");
  if (!latin) return undefined;
  const raw = latin.attrs.typeface;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  return trimmed;
}

// ── Auto Title Deleted ────────────────────────────────────────────

/**
 * Pull `<c:autoTitleDeleted val=".."/>` off `<c:chart>`. Surfaces
 * `true` only when the chart pinned `<c:autoTitleDeleted val="1"/>`
 * (the non-default state — the user explicitly deleted the
 * auto-generated title that single-series charts synthesise from the
 * series name). The OOXML default `val="0"` and absence both collapse
 * to `undefined` so absence and the default round-trip identically
 * through {@link cloneChart}.
 *
 * The element is independent of `<c:title>` — it sits on `<c:chart>`
 * directly (between `<c:title>` and `<c:plotArea>` per CT_Chart,
 * ECMA-376 Part 1, §21.2.2.4), not nested inside `<c:title>`. A chart
 * with a literal `<c:title>` typically pins `val="0"` because the user
 * has not deleted the auto-title (they overrode it with a literal
 * one); a chart with no `<c:title>` may pin `val="1"` to suppress
 * Excel's auto-title synthesis or omit the element entirely (Excel
 * may still synthesise an auto-title in that case for a single-series
 * chart).
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` / `"0"`
 * / `"false"`); unknown values and missing `val` attributes drop to
 * `undefined` rather than fabricate a flag Excel would not emit.
 */
function parseAutoTitleDeleted(chartEl: XmlElement): boolean | undefined {
  const el = findChild(chartEl, "autoTitleDeleted");
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
      // writer's `autoTitleDeleted` field.
      return undefined;
    default:
      return undefined;
  }
}

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

/**
 * Pull `<c:dTable>...</c:dTable>` off `<c:plotArea>`. Surfaces a
 * {@link ChartDataTable} whenever the source chart declares the
 * element; absence collapses to `undefined`.
 *
 * Each of the four boolean children (`<c:showHorzBorder>`,
 * `<c:showVertBorder>`, `<c:showOutline>`, `<c:showKeys>`) round-trips
 * literally — the reader does not collapse any per-field default
 * because all four are required on `CT_DTable` and Excel always emits
 * every one. Children that are missing or carry an unknown `val`
 * attribute drop to `undefined` rather than fabricate a flag the file
 * did not pin; the writer falls back to the OOXML reference defaults
 * (`true` for every child) on round-trip.
 */
function parseDataTable(plotArea: XmlElement): ChartDataTable | undefined {
  const el = findChild(plotArea, "dTable");
  if (!el) return undefined;
  const out: ChartDataTable = {};
  const showHorzBorder = parseDataTableFlag(el, "showHorzBorder");
  if (showHorzBorder !== undefined) out.showHorzBorder = showHorzBorder;
  const showVertBorder = parseDataTableFlag(el, "showVertBorder");
  if (showVertBorder !== undefined) out.showVertBorder = showVertBorder;
  const showOutline = parseDataTableFlag(el, "showOutline");
  if (showOutline !== undefined) out.showOutline = showOutline;
  const showKeys = parseDataTableFlag(el, "showKeys");
  if (showKeys !== undefined) out.showKeys = showKeys;
  const fontSize = parseDataTableFontSize(el);
  if (fontSize !== undefined) out.fontSize = fontSize;
  const fontColor = parseDataTableFontColor(el);
  if (fontColor !== undefined) out.fontColor = fontColor;
  const bold = parseDataTableBold(el);
  if (bold !== undefined) out.bold = bold;
  const italic = parseDataTableItalic(el);
  if (italic !== undefined) out.italic = italic;
  const underline = parseDataTableUnderline(el);
  if (underline !== undefined) out.underline = underline;
  const strikethrough = parseDataTableStrikethrough(el);
  if (strikethrough !== undefined) out.strikethrough = strikethrough;
  const fontFamily = parseDataTableFontFamily(el);
  if (fontFamily !== undefined) out.fontFamily = fontFamily;
  return out;
}

/**
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
 * </c:txPr></c:dTable>` off a data-table block. Returns the bold flag.
 *
 * The OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties`. Only an explicit `b="1"` (or `"true"`)
 * surfaces `true`; the OOXML default `0` (and absence / malformed
 * tokens) collapses to `undefined` so absence and the default
 * round-trip identically through `cloneChart`. Mirrors the
 * chart-title / axis-title / axis tick-label / legend / data-label
 * bold readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
function parseDataTableBold(dTable: XmlElement): boolean | undefined {
  const txPr = findChild(dTable, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.b;
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

/**
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
 * </c:txPr></c:dTable>` off a data-table block. Returns the italic
 * flag.
 *
 * The OOXML `i` attribute is the `xsd:boolean` italic flag on
 * `CT_TextCharacterProperties`. An explicit `i="1"` / `"true"`
 * surfaces `true`; an explicit `i="0"` / `"false"` surfaces `false`
 * so a templated chart that pinned `i="0"` round-trips literally and
 * a clone target can override an upstream `i="1"` cleanly. Absence /
 * malformed tokens collapse to `undefined` so a fresh data table that
 * never pinned the flag round-trips identically through `cloneChart`.
 * Mirrors the data-table bold reader (and the chart-title /
 * axis-title / axis tick-label / legend / data-label italic readers
 * for the OOXML attribute layout) so a parsed value slots straight
 * back into the writer's emit path.
 */
function parseDataTableItalic(dTable: XmlElement): boolean | undefined {
  const txPr = findChild(dTable, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.i;
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

/**
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p>
 * </c:txPr></c:dTable>` off a data-table block. Returns the underline
 * flag.
 *
 * The OOXML `u` attribute is the `ST_TextUnderlineType` enumeration
 * on `CT_TextCharacterProperties`. `u="sng"` (Excel's UI variant —
 * single underline) surfaces `true`; `u="none"` (the OOXML default)
 * surfaces `false` so a templated chart that pinned `u="none"`
 * round-trips literally and a clone target can override an upstream
 * `u="sng"` cleanly. The schema's other variants (`"dbl"`, `"heavy"`,
 * `"dotted"`, `"dotDash"`, `"wavy"`, etc.) and absence / malformed
 * tokens collapse to `undefined` so a fresh data table that never
 * pinned the flag round-trips identically through `cloneChart`. The
 * reader never silently downgrades a non-`"sng"` underline to a
 * single line on round-trip; only the boolean shape Excel's UI
 * exposes survives the parse.
 *
 * Mirrors the data-table bold reader pattern — `true` and `false`
 * round-trip explicitly so a clone can override either direction —
 * while the OOXML attribute layout matches the chart-title /
 * axis-title / axis tick-label / legend / data-label underline
 * readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
function parseDataTableUnderline(dTable: XmlElement): boolean | undefined {
  const txPr = findChild(dTable, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.u;
  if (raw === "sng") return true;
  if (raw === "none") return false;
  return undefined;
}

/**
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
 * </a:p></c:txPr></c:dTable>` off a data-table block. Returns the
 * strikethrough flag.
 *
 * The OOXML `strike` attribute is the `ST_TextStrikeType`
 * enumeration on `CT_TextCharacterProperties`. Only
 * `strike="sngStrike"` (Excel's UI variant — single line) surfaces
 * `true`; the OOXML default `"noStrike"` and the non-UI variant
 * `"dblStrike"` (and any malformed token) collapse to `undefined` so
 * absence and `"noStrike"` round-trip identically through
 * `cloneChart`. Reporting `"dblStrike"` as `true` would silently
 * downgrade the choice to a single line on round-trip; the writer
 * emits only `"sngStrike"`, matching the boolean shape the UI
 * exposes. Mirrors the chart-title / axis-title / axis tick-label /
 * legend / data-label strikethrough readers exactly so a parsed value
 * slots straight back into the writer's emit path.
 */
function parseDataTableStrikethrough(dTable: XmlElement): boolean | undefined {
  const txPr = findChild(dTable, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.strike;
  if (raw === "sngStrike") return true;
  return undefined;
}

/**
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:txPr></c:dTable>` off
 * a data-table block. Returns the typeface string the table was
 * authored with.
 *
 * The OOXML `<a:latin>` element carries the typeface name on
 * `CT_TextFont` (ECMA-376 Part 1, §21.1.2.3.7). The reader trims
 * surrounding whitespace and reports the trimmed typeface; empty /
 * whitespace-only `typeface` attributes and missing `<a:latin>`
 * elements both collapse to `undefined` so absence and the empty
 * form round-trip identically through the writer. Non-string
 * `typeface` tokens (defensive — the XML parser only ever surfaces
 * strings) likewise drop to `undefined`.
 *
 * Returns `undefined` whenever the data-table block omits `<c:txPr>`
 * entirely or the canonical `<a:p><a:pPr><a:defRPr><a:latin>` chain
 * is malformed at any link. Mirrors the chart-title / axis-title /
 * axis tick-label / legend / data-label font family readers exactly
 * so a parsed value slots straight back into the writer's emit path.
 */
function parseDataTableFontFamily(dTable: XmlElement): string | undefined {
  const txPr = findChild(dTable, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const latin = findChild(defRPr, "latin");
  if (!latin) return undefined;
  const raw = latin.attrs.typeface;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  return trimmed;
}

/**
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:txPr></c:dTable>` off a data-table block. Returns the
 * 6-character uppercase hex string when the parser walks the full
 * chain and lands on an `<a:srgbClr val="RRGGBB"/>`. Theme references
 * (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`, and `<a:prstClr>`
 * all collapse to `undefined` — only the literal RGB triple round-
 * trips losslessly through {@link writeChart}. Malformed `val` tokens
 * (wrong length, non-hex characters) likewise drop to `undefined`
 * rather than fabricate a value the writer would round-trip into a
 * malformed `<a:srgbClr>`.
 *
 * Returns `undefined` whenever the data-table block omits `<c:txPr>`
 * entirely or the canonical chain is malformed at any link. Mirrors
 * the chart-title / axis-title / axis tick-label / legend / data-label
 * color readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
function parseDataTableFontColor(dTable: XmlElement): string | undefined {
  const txPr = findChild(dTable, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const solidFill = findChild(defRPr, "solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "srgbClr");
  if (!srgbClr) return undefined;
  return normalizeRgbHex(srgbClr.attrs.val);
}

/**
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
 * </c:txPr></c:dTable>` off a data-table block. Returns the font size
 * in points (`1..400`), or `undefined` when the element is absent /
 * the chain is malformed at any link / the surfaced value is out of
 * the supported range.
 *
 * The OOXML attribute is in 100ths of a point on
 * `CT_TextCharacterProperties`' `sz` slot (ECMA-376 Part 1,
 * §21.1.2.3.7); the reader divides by 100 at parse time so the
 * surfaced value matches what the user sees in Excel's UI. Mirrors
 * the chart-title / axis-title / axis tick-label / legend / data-label
 * font-size readers exactly so a parsed value slots straight back into
 * the writer's emit path.
 */
function parseDataTableFontSize(dTable: XmlElement): number | undefined {
  const txPr = findChild(dTable, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.sz;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  // Convert from 100ths of a point to points, rounding to the nearest
  // 0.5pt to match the granularity Excel's UI exposes. Mirrors the
  // chart-title / axis-title / tick-label / legend / data-label
  // sibling parsers exactly so a parsed value flows through every
  // typography slot without bookkeeping the units.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2);
  const points = halfSteps / 2;
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined;
  return points;
}

/**
 * Pull a single boolean child off `<c:dTable>`. Accepts the OOXML
 * truthy / falsy spellings (`"1"` / `"true"` / `"0"` / `"false"`);
 * unknown tokens, missing `val` attributes, and missing elements all
 * collapse to `undefined` rather than fabricate a flag the file did
 * not pin.
 */
function parseDataTableFlag(dTable: XmlElement, local: string): boolean | undefined {
  const el = findChild(dTable, local);
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

/**
 * Pull `<c:view3D>` (CT_View3D) off `<c:chart>`. Surfaces a
 * {@link ChartView3D} object whenever the source chart declares the
 * element. Each of the six children (`<c:rotX>`, `<c:hPercent>`,
 * `<c:rotY>`, `<c:depthPercent>`, `<c:rAngAx>`, `<c:perspective>`)
 * is independently optional on CT_View3D, so the reader only surfaces
 * the fields the file actually pinned. A child that is missing or
 * carries an out-of-range / unparseable `val` attribute drops to
 * `undefined` for that field rather than fabricate a value the file
 * did not declare.
 *
 * The element itself is the gating signal — a `<c:view3D>` block with
 * no resolvable children surfaces as an empty `{}`, mirroring how
 * `dataTable` / `protection` handle a malformed inner block. This
 * keeps a chart that authors the bare element (Excel's "default 3D
 * view" preset) from silently disappearing through the parse loop.
 *
 * Note: `<c:view3D>` lives on `<c:chart>` (between `<c:autoTitleDeleted>`
 * / `<c:pivotFmts>` and `<c:floor>` / `<c:plotArea>` per CT_Chart
 * §21.2.2.4), not on `<c:chartSpace>` — the toggle governs the 3D
 * projection of the rendered chart, not the outer chart frame.
 */
function parseView3D(chartEl: XmlElement): ChartView3D | undefined {
  const el = findChild(chartEl, "view3D");
  if (!el) return undefined;
  const out: ChartView3D = {};
  // `<c:rotX>` (CT_RotX, ST_RotX) is a signed byte in the range
  // -90..90. Out-of-range values drop rather than emit a token Excel
  // would clamp at parse time.
  const rotX = parseView3DInt(el, "rotX", -90, 90);
  if (rotX !== undefined) out.rotX = rotX;
  // `<c:hPercent>` (CT_HPercent, ST_HPercent) is a percent value in
  // the range 5..500. Same drop-on-out-of-range rule.
  const hPercent = parseView3DInt(el, "hPercent", 5, 500);
  if (hPercent !== undefined) out.hPercent = hPercent;
  // `<c:rotY>` (CT_RotY, ST_RotY) is an unsigned short in the range
  // 0..360.
  const rotY = parseView3DInt(el, "rotY", 0, 360);
  if (rotY !== undefined) out.rotY = rotY;
  // `<c:depthPercent>` (CT_DepthPercent, ST_DepthPercent) is a percent
  // value in the range 20..2000.
  const depthPercent = parseView3DInt(el, "depthPercent", 20, 2000);
  if (depthPercent !== undefined) out.depthPercent = depthPercent;
  // `<c:rAngAx>` (CT_Boolean) — accepts the OOXML truthy / falsy
  // spellings; unknown values and missing `val` attributes drop to
  // `undefined`. Mirrors the parsing semantics of the chartSpace-level
  // `<c:protection>` boolean children.
  const rAngAx = parseView3DBoolean(el, "rAngAx");
  if (rAngAx !== undefined) out.rAngAx = rAngAx;
  // `<c:perspective>` (CT_Perspective, ST_Perspective) is a percent
  // value in the range 0..240.
  const perspective = parseView3DInt(el, "perspective", 0, 240);
  if (perspective !== undefined) out.perspective = perspective;
  return out;
}

/**
 * Pull a single integer child off `<c:view3D>`. Surfaces the value
 * only when `val` parses as an integer inside the matching OOXML
 * simple-type range; absence and out-of-range / non-integer values
 * collapse to `undefined`.
 *
 * Accepts an optional leading `-` so signed types (`<c:rotX>`) round-
 * trip cleanly. The strict integer regex rejects fractional values
 * (`"15.5"`) and non-numeric tokens (`"15px"`) — `parseInt` would
 * coerce both into a number Excel never emits.
 */
function parseView3DInt(
  view3D: XmlElement,
  local: string,
  min: number,
  max: number,
): number | undefined {
  const el = findChild(view3D, local);
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  if (!/^-?\d+$/.test(raw)) return undefined;
  const n = Number(raw);
  if (!Number.isInteger(n)) return undefined;
  if (n < min || n > max) return undefined;
  return n;
}

/**
 * Pull a single boolean child off `<c:view3D>`. Accepts the OOXML
 * truthy / falsy spellings (`"1"` / `"true"` / `"0"` / `"false"`);
 * unknown tokens, missing `val` attributes, and missing elements all
 * collapse to `undefined` rather than fabricate a flag the file did
 * not pin. Mirrors {@link parseProtectionFlag} — the same OOXML
 * `<xsd:boolean>` lexical-space rule.
 */
function parseView3DBoolean(view3D: XmlElement, local: string): boolean | undefined {
  const el = findChild(view3D, local);
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

// ── Floor Thickness ───────────────────────────────────────────────

/**
 * Pull `<c:chart><c:floor><c:thickness val=".."/></c:floor>` off
 * `<c:chart>`. The `<c:floor>` element (CT_Surface, ECMA-376 Part 1,
 * §21.2.2.69) sits on `<c:chart>` between `<c:view3D>` and
 * `<c:sideWall>` / `<c:backWall>` / `<c:plotArea>` per CT_Chart and
 * carries an optional `<c:thickness>` child whose `val` attribute is
 * an `xsd:unsignedInt` — Excel's "Format Floor -> Floor -> Thickness"
 * pin on 3D chart families.
 *
 * Returns the integer pinned by the source chart. The OOXML default
 * `0` (and absence of the `<c:thickness>` child or the parent
 * `<c:floor>` element) collapses to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart} — only an
 * explicit positive thickness surfaces here. Out-of-range or
 * unparseable values also drop to `undefined` rather than fabricate a
 * value the file did not declare.
 *
 * The `<c:thickness>` element only carries the `val` attribute on
 * `CT_Thickness` — other floor styling (`<c:spPr>`, `<c:pictureOptions>`,
 * `<c:extLst>`) is not modelled at this layer, so a stray styling
 * block on the floor passes through the parse loop without surfacing.
 */
function parseFloorThickness(chartEl: XmlElement): number | undefined {
  const floor = findChild(chartEl, "floor");
  if (!floor) return undefined;
  const thickness = findChild(floor, "thickness");
  if (!thickness) return undefined;
  const raw = thickness.attrs.val;
  if (typeof raw !== "string") return undefined;
  // ST_Thickness is `xsd:unsignedInt` — strict integer regex rejects
  // fractional / negative / non-numeric tokens.
  if (!/^\d+$/.test(raw)) return undefined;
  const n = Number(raw);
  if (!Number.isInteger(n)) return undefined;
  // Collapse the OOXML default `0` to undefined so absence and the
  // default round-trip identically through cloneChart — only an
  // explicit positive thickness surfaces here. Mirrors how the writer
  // skips emission entirely for `0` / undefined.
  if (n === 0) return undefined;
  // Cap at Excel's UI band ceiling (`100`) — the OOXML schema accepts
  // the full `xsd:unsignedInt` range but Excel's "Format Floor"
  // dialogue rejects values above 100 with a repair warning. Anything
  // larger drops here so a corrupt template does not silently rewrite
  // as an absurd thickness; absence keeps the round-trip stable.
  if (n > 100) return undefined;
  return n;
}

// ── Side Wall Thickness ───────────────────────────────────────────

/**
 * Pull `<c:chart><c:sideWall><c:thickness val=".."/></c:sideWall>` off
 * `<c:chart>`. The `<c:sideWall>` element (CT_Surface, ECMA-376 Part 1,
 * §21.2.2.187) sits on `<c:chart>` between `<c:floor>` and
 * `<c:backWall>` / `<c:plotArea>` per CT_Chart and carries an optional
 * `<c:thickness>` child whose `val` attribute is an `xsd:unsignedInt` —
 * Excel's "Format Side Wall -> Side Wall -> Thickness" pin on 3D chart
 * families.
 *
 * Returns the integer pinned by the source chart. The OOXML default
 * `0` (and absence of the `<c:thickness>` child or the parent
 * `<c:sideWall>` element) collapses to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart} — only an
 * explicit positive thickness surfaces here. Out-of-range or
 * unparseable values also drop to `undefined` rather than fabricate a
 * value the file did not declare.
 *
 * The `<c:thickness>` element only carries the `val` attribute on
 * `CT_Thickness` — other side-wall styling (`<c:spPr>`,
 * `<c:pictureOptions>`, `<c:extLst>`) is not modelled at this layer,
 * so a stray styling block on the wall passes through the parse loop
 * without surfacing.
 */
function parseSideWallThickness(chartEl: XmlElement): number | undefined {
  const sideWall = findChild(chartEl, "sideWall");
  if (!sideWall) return undefined;
  const thickness = findChild(sideWall, "thickness");
  if (!thickness) return undefined;
  const raw = thickness.attrs.val;
  if (typeof raw !== "string") return undefined;
  // ST_Thickness is `xsd:unsignedInt` — strict integer regex rejects
  // fractional / negative / non-numeric tokens.
  if (!/^\d+$/.test(raw)) return undefined;
  const n = Number(raw);
  if (!Number.isInteger(n)) return undefined;
  // Collapse the OOXML default `0` to undefined so absence and the
  // default round-trip identically through cloneChart — only an
  // explicit positive thickness surfaces here. Mirrors how the writer
  // skips emission entirely for `0` / undefined.
  if (n === 0) return undefined;
  // Cap at Excel's UI band ceiling (`100`) — the OOXML schema accepts
  // the full `xsd:unsignedInt` range but Excel's "Format Side Wall"
  // dialogue rejects values above 100 with a repair warning. Anything
  // larger drops here so a corrupt template does not silently rewrite
  // as an absurd thickness; absence keeps the round-trip stable.
  if (n > 100) return undefined;
  return n;
}

// ── Back Wall Thickness ───────────────────────────────────────────

/**
 * Pull `<c:chart><c:backWall><c:thickness val=".."/></c:backWall>` off
 * `<c:chart>`. The `<c:backWall>` element (CT_Surface, ECMA-376 Part 1,
 * §21.2.2.31) sits on `<c:chart>` between `<c:sideWall>` and
 * `<c:plotArea>` per CT_Chart and carries an optional `<c:thickness>`
 * child whose `val` attribute is an `xsd:unsignedInt` — Excel's
 * "Format Back Wall -> Back Wall -> Thickness" pin on 3D chart families.
 *
 * Returns the integer pinned by the source chart. The OOXML default
 * `0` (and absence of the `<c:thickness>` child or the parent
 * `<c:backWall>` element) collapses to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart} — only an
 * explicit positive thickness surfaces here. Out-of-range or
 * unparseable values also drop to `undefined` rather than fabricate a
 * value the file did not declare.
 *
 * The `<c:thickness>` element only carries the `val` attribute on
 * `CT_Thickness` — other back-wall styling (`<c:spPr>`,
 * `<c:pictureOptions>`, `<c:extLst>`) is not modelled at this layer,
 * so a stray styling block on the back wall passes through the parse
 * loop without surfacing.
 */
function parseBackWallThickness(chartEl: XmlElement): number | undefined {
  const backWall = findChild(chartEl, "backWall");
  if (!backWall) return undefined;
  const thickness = findChild(backWall, "thickness");
  if (!thickness) return undefined;
  const raw = thickness.attrs.val;
  if (typeof raw !== "string") return undefined;
  // ST_Thickness is `xsd:unsignedInt` — strict integer regex rejects
  // fractional / negative / non-numeric tokens.
  if (!/^\d+$/.test(raw)) return undefined;
  const n = Number(raw);
  if (!Number.isInteger(n)) return undefined;
  // Collapse the OOXML default `0` to undefined so absence and the
  // default round-trip identically through cloneChart — only an
  // explicit positive thickness surfaces here. Mirrors how the writer
  // skips emission entirely for `0` / undefined.
  if (n === 0) return undefined;
  // Cap at Excel's UI band ceiling (`100`) — the OOXML schema accepts
  // the full `xsd:unsignedInt` range but Excel's "Format Back Wall"
  // dialogue rejects values above 100 with a repair warning. Anything
  // larger drops here so a corrupt template does not silently rewrite
  // as an absurd thickness; absence keeps the round-trip stable.
  if (n > 100) return undefined;
  return n;
}

// ── Vary Colors ────────────────────────────────────────────────────

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
function parseVaryColors(chartTypeEl: XmlElement, kind: ChartKind): boolean | undefined {
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

// ── Scatter Style ─────────────────────────────────────────────────

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
function parseScatterStyle(scatterChart: XmlElement): ChartScatterStyle | undefined {
  const el = findChild(scatterChart, "scatterStyle");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  if (!VALID_SCATTER_STYLES.has(raw as ChartScatterStyle)) return undefined;
  return raw as ChartScatterStyle;
}

// ── Drop Lines / Hi-Low Lines ─────────────────────────────────────

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
function parseDropLines(chartTypeEl: XmlElement): boolean | undefined {
  return findChild(chartTypeEl, "dropLines") ? true : undefined;
}

/**
 * Pull `<c:hiLowLines/>` off a `<c:lineChart>` / `<c:line3DChart>` /
 * `<c:stockChart>` element. Same on/off shape as
 * {@link parseDropLines}; the element is bare so its mere presence
 * surfaces `true`, absence collapses to `undefined`.
 */
function parseHiLowLines(chartTypeEl: XmlElement): boolean | undefined {
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
function parseSerLines(chartTypeEl: XmlElement): boolean | undefined {
  return findChild(chartTypeEl, "serLines") ? true : undefined;
}

// ── Chart-level Marker Visibility ─────────────────────────────────

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
function parseShowLineMarkers(lineChart: XmlElement): boolean | undefined {
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

// ── Bar Grouping ──────────────────────────────────────────────────

/**
 * Pull `<c:grouping val=".."/>` off a `<c:barChart>` element. Returns
 * `undefined` when the grouping element is missing or carries the
 * default `"standard"` / `"clustered"` value — the writer's
 * {@link SheetChart.barGrouping} treats both as the unspecified
 * default, so omitting them keeps the parsed shape minimal.
 */
function parseBarGrouping(barChart: XmlElement): ChartBarGrouping | undefined {
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

// ── Doughnut Hole ─────────────────────────────────────────────────

/**
 * Pull `<c:holeSize val=".."/>` off a `<c:doughnutChart>` element.
 * Returns `undefined` when the attribute is missing, malformed, or
 * outside the 1–99 range OOXML allows. Excel itself only writes values
 * in 10–90 (the UI clamps to that band) but the spec is wider, so we
 * accept the full schema range and let the writer re-clamp on the way
 * back out.
 */
function parseHoleSize(doughnut: XmlElement): number | undefined {
  const el = findChild(doughnut, "holeSize");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < 1 || parsed > 99) return undefined;
  return parsed;
}

// ── Bar / Column gap width & overlap ──────────────────────────────

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
function parseGapWidth(barChart: XmlElement): number | undefined {
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
function parseUpDownBarsGapWidth(upDownBars: XmlElement): number | undefined {
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
function parseOverlap(barChart: XmlElement): number | undefined {
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

// ── First Slice Angle ─────────────────────────────────────────────

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
function parseFirstSliceAng(chartType: XmlElement): number | undefined {
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
function parseLineAreaGrouping(chartType: XmlElement): ChartLineAreaGrouping | undefined {
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

/**
 * Read `<c:title>` text. The title may be a rich-text run tree or a
 * formula reference; we only surface plain text runs joined together.
 */
function parseTitle(chartEl: XmlElement): string | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  // tx can hold either <c:rich> (literal text) or <c:strRef> (formula).
  const rich = findChild(tx, "rich");
  if (rich) {
    const parts: string[] = [];
    collectTextRuns(rich, parts);
    const joined = parts.join("").trim();
    return joined.length > 0 ? joined : undefined;
  }
  const strRef = findChild(tx, "strRef");
  if (strRef) {
    const cache = findChild(strRef, "strCache");
    if (cache) {
      for (const pt of childElements(cache)) {
        if (pt.local !== "pt") continue;
        const v = findChild(pt, "v");
        if (v) {
          const text = elementText(v).trim();
          if (text.length > 0) return text;
        }
      }
    }
  }
  return undefined;
}

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

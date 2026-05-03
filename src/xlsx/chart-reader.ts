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
  ChartAxisCrosses,
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
  ChartDisplayBlanksAs,
  ChartKind,
  ChartLegendPosition,
  ChartLineAreaGrouping,
  ChartLineDashStyle,
  ChartLineStroke,
  ChartMarker,
  ChartMarkerSymbol,
  ChartScatterStyle,
  ChartSeriesInfo,
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

    const axes = parseAxes(plotArea);
    if (axes !== undefined) out.axes = axes;
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
  }

  const dispBlanksAs = parseDispBlanksAs(chartEl);
  if (dispBlanksAs !== undefined) out.dispBlanksAs = dispBlanksAs;

  const plotVisOnly = parsePlotVisOnly(chartEl);
  if (plotVisOnly !== undefined) out.plotVisOnly = plotVisOnly;

  // `<c:roundedCorners>` lives on `<c:chartSpace>` (the chart's outer
  // wrapper), not inside `<c:chart>` — the toggle styles the chart
  // frame's outer border rather than the plot area.
  const roundedCorners = parseRoundedCorners(chartSpace);
  if (roundedCorners !== undefined) out.roundedCorners = roundedCorners;

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

  const x = xAxis ? parseAxisInfo(xAxis) : undefined;
  const y = yAxis ? parseAxisInfo(yAxis) : undefined;

  if (!x && !y) return undefined;
  const out: { x?: ChartAxisInfo; y?: ChartAxisInfo } = {};
  if (x) out.x = x;
  if (y) out.y = y;
  return out;
}

function parseAxisInfo(axis: XmlElement): ChartAxisInfo | undefined {
  const title = parseAxisTitle(axis);
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
  if (
    title === undefined &&
    gridlines === undefined &&
    scale === undefined &&
    numberFormat === undefined &&
    majorTickMark === undefined &&
    minorTickMark === undefined &&
    tickLblPos === undefined &&
    reverse === undefined &&
    tickLblSkip === undefined &&
    tickMarkSkip === undefined &&
    lblOffset === undefined &&
    lblAlgn === undefined &&
    noMultiLvlLbl === undefined &&
    hidden === undefined &&
    crosses === undefined &&
    crossesAt === undefined
  ) {
    return undefined;
  }
  const out: ChartAxisInfo = {};
  if (title !== undefined) out.title = title;
  if (gridlines !== undefined) out.gridlines = gridlines;
  if (scale !== undefined) out.scale = scale;
  if (numberFormat !== undefined) out.numberFormat = numberFormat;
  if (majorTickMark !== undefined) out.majorTickMark = majorTickMark;
  if (minorTickMark !== undefined) out.minorTickMark = minorTickMark;
  if (tickLblPos !== undefined) out.tickLblPos = tickLblPos;
  if (reverse !== undefined) out.reverse = reverse;
  if (tickLblSkip !== undefined) out.tickLblSkip = tickLblSkip;
  if (tickMarkSkip !== undefined) out.tickMarkSkip = tickMarkSkip;
  if (lblOffset !== undefined) out.lblOffset = lblOffset;
  if (lblAlgn !== undefined) out.lblAlgn = lblAlgn;
  if (noMultiLvlLbl !== undefined) out.noMultiLvlLbl = noMultiLvlLbl;
  if (hidden !== undefined) out.hidden = hidden;
  if (crosses !== undefined) out.crosses = crosses;
  if (crossesAt !== undefined) out.crossesAt = crossesAt;
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

  // Empty record is meaningless to a consumer — collapse to undefined.
  if (
    out.position === undefined &&
    !out.showValue &&
    !out.showCategoryName &&
    !out.showSeriesName &&
    !out.showPercent &&
    !out.showLegendKey &&
    out.separator === undefined
  ) {
    return undefined;
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

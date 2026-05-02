// ── Chart Writer ─────────────────────────────────────────────────────
// Generates xl/charts/chartN.xml for native Excel chart creation.
//
// Phase 1 of issue #152: bar / column / line / pie / scatter / area.
// The chart XML follows the DrawingML chart spec (ECMA-376 Part 1,
// Chapter 21). Each chart is a self-contained <c:chartSpace> document
// referenced from a drawing part via a `chart` relationship.

import type {
  ChartAxisGridlines,
  ChartAxisNumberFormat,
  ChartAxisScale,
  ChartDataLabels,
  ChartSeries,
  SheetChart,
  WriteChartKind,
} from "../_types";
import { xmlDocument, xmlElement, xmlEscape, xmlSelfClose } from "../xml/writer";

// ── Namespaces ───────────────────────────────────────────────────────

const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";

// ── Public API ───────────────────────────────────────────────────────

export interface ChartWriteResult {
  /** Body of `xl/charts/chartN.xml`. */
  chartXml: string;
  /**
   * Body of `xl/charts/_rels/chartN.xml.rels`. Always present so the
   * package validator stays happy even though Phase 1 charts have no
   * outgoing relationships.
   */
  chartRels: string;
}

/**
 * Generate the OOXML chart document for a single chart.
 *
 * @param chart - High-level chart definition from the user.
 * @param sheetName - Sheet that owns the chart. Used to qualify bare
 *                    cell references such as `"B2:B4"`.
 */
export function writeChart(chart: SheetChart, sheetName: string): ChartWriteResult {
  const showTitle = chart.showTitle ?? Boolean(chart.title);
  const legendPos = resolveLegendPosition(chart);

  const chartChildren: string[] = [];

  // ── Title ──
  if (showTitle && chart.title) {
    chartChildren.push(buildTitle(chart.title));
    chartChildren.push(xmlSelfClose("c:autoTitleDeleted", { val: 0 }));
  } else {
    chartChildren.push(xmlSelfClose("c:autoTitleDeleted", { val: 1 }));
  }

  // ── Plot Area ──
  chartChildren.push(buildPlotArea(chart, sheetName));

  // ── Legend ──
  if (legendPos) {
    chartChildren.push(buildLegend(legendPos));
  }

  chartChildren.push(xmlSelfClose("c:plotVisOnly", { val: 1 }));
  chartChildren.push(xmlSelfClose("c:dispBlanksAs", { val: "gap" }));

  const chartElement = xmlElement("c:chart", undefined, chartChildren);

  const chartXml = xmlDocument(
    "c:chartSpace",
    {
      "xmlns:c": NS_C,
      "xmlns:a": NS_A,
      "xmlns:r": NS_R,
    },
    [xmlSelfClose("c:roundedCorners", { val: 0 }), chartElement],
  );

  // Always emit an empty rels file. Phase 1 charts do not depend on
  // any other parts (no themeOverride, no userShapes, no embedded
  // spreadsheets), but Excel and several validators expect the file
  // to exist whenever a `chartN.xml` is declared.
  const chartRels = xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, []);

  return { chartXml, chartRels };
}

// ── Title ────────────────────────────────────────────────────────────

function buildTitle(title: string): string {
  return xmlElement("c:title", undefined, [
    xmlElement("c:tx", undefined, [
      xmlElement("c:rich", undefined, [
        xmlElement(
          "a:bodyPr",
          {
            rot: 0,
            spcFirstLastPara: 1,
            vertOverflow: "ellipsis",
            wrap: "square",
            anchor: "ctr",
            anchorCtr: 1,
          },
          [],
        ),
        xmlSelfClose("a:lstStyle"),
        xmlElement("a:p", undefined, [
          xmlElement("a:pPr", undefined, [xmlSelfClose("a:defRPr", { sz: 1400, b: 0 })]),
          xmlElement("a:r", undefined, [
            xmlSelfClose("a:rPr", { lang: "en-US", sz: 1400, b: 0 }),
            xmlElement("a:t", undefined, xmlEscape(title)),
          ]),
        ]),
      ]),
    ]),
    xmlSelfClose("c:overlay", { val: 0 }),
  ]);
}

// ── Plot Area ────────────────────────────────────────────────────────

function buildPlotArea(chart: SheetChart, sheetName: string): string {
  const children: string[] = [xmlSelfClose("c:layout")];

  // Axis titles, gridlines, scaling and number format surface for
  // every chart family except pie/doughnut. Pull them once so each
  // branch can hand them off to the matching axis builder.
  const opts: AxisRenderOptions = {
    xAxisTitle: normalizeAxisTitle(chart.axes?.x?.title),
    yAxisTitle: normalizeAxisTitle(chart.axes?.y?.title),
    xGridlines: normalizeAxisGridlines(chart.axes?.x?.gridlines),
    yGridlines: normalizeAxisGridlines(chart.axes?.y?.gridlines),
    xScale: normalizeAxisScale(chart.axes?.x?.scale),
    yScale: normalizeAxisScale(chart.axes?.y?.scale),
    xNumFmt: normalizeAxisNumberFormat(chart.axes?.x?.numberFormat),
    yNumFmt: normalizeAxisNumberFormat(chart.axes?.y?.numberFormat),
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

  return xmlElement("c:plotArea", undefined, children);
}

interface AxisRenderOptions {
  xAxisTitle: string | undefined;
  yAxisTitle: string | undefined;
  xGridlines: { major: boolean; minor: boolean } | undefined;
  yGridlines: { major: boolean; minor: boolean } | undefined;
  xScale: ChartAxisScale | undefined;
  yScale: ChartAxisScale | undefined;
  xNumFmt: ChartAxisNumberFormat | undefined;
  yNumFmt: ChartAxisNumberFormat | undefined;
}

/**
 * Normalize an axis title input to either a non-empty trimmed string
 * or `undefined`. Empty strings are dropped so the writer never emits
 * an empty `<c:title>` element (Excel renders that as an unintended
 * blank label).
 */
function normalizeAxisTitle(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

/**
 * Resolve the gridline toggles to a stable record (or `undefined` when
 * neither is on). Mirrors {@link normalizeAxisTitle} so the per-branch
 * code in `buildPlotArea` only needs a single null check.
 */
function normalizeAxisGridlines(
  value: ChartAxisGridlines | undefined,
): { major: boolean; minor: boolean } | undefined {
  if (!value) return undefined;
  const major = value.major === true;
  const minor = value.minor === true;
  if (!major && !minor) return undefined;
  return { major, minor };
}

/**
 * Build the `<c:majorGridlines>` / `<c:minorGridlines>` block for an
 * axis. The returned XML fragments must be appended in spec order
 * (major before minor) and slot in immediately after `<c:axPos>`,
 * before the optional `<c:title>`. Excel's strict-validator rejects
 * any other position.
 */
function buildAxisGridlines(gridlines: { major: boolean; minor: boolean } | undefined): string[] {
  if (!gridlines) return [];
  const out: string[] = [];
  if (gridlines.major) out.push(xmlElement("c:majorGridlines", undefined, []));
  if (gridlines.minor) out.push(xmlElement("c:minorGridlines", undefined, []));
  return out;
}

/**
 * Drop fields that won't survive Excel's strict validator. Non-finite
 * numbers, `min >= max`, and zero/negative tick spacings all collapse
 * the corresponding entry to `undefined` so the writer never emits a
 * `<c:min>`/`<c:max>`/`<c:majorUnit>`/`<c:minorUnit>` Excel would
 * reject.
 *
 * Returns `undefined` when nothing usable remains so the writer can
 * skip the entire `<c:scaling>` augmentation.
 */
function normalizeAxisScale(value: ChartAxisScale | undefined): ChartAxisScale | undefined {
  if (!value) return undefined;
  const out: ChartAxisScale = {};
  if (typeof value.min === "number" && Number.isFinite(value.min)) out.min = value.min;
  if (typeof value.max === "number" && Number.isFinite(value.max)) out.max = value.max;
  if (out.min !== undefined && out.max !== undefined && out.min >= out.max) {
    // min >= max is meaningless; preserve the user-supplied min only
    // so validators don't choke on a flipped/empty axis range.
    delete out.max;
  }
  if (
    typeof value.majorUnit === "number" &&
    Number.isFinite(value.majorUnit) &&
    value.majorUnit > 0
  ) {
    out.majorUnit = value.majorUnit;
  }
  if (
    typeof value.minorUnit === "number" &&
    Number.isFinite(value.minorUnit) &&
    value.minorUnit > 0
  ) {
    out.minorUnit = value.minorUnit;
  }
  if (
    typeof value.logBase === "number" &&
    Number.isFinite(value.logBase) &&
    value.logBase >= 2 &&
    value.logBase <= 1000
  ) {
    out.logBase = value.logBase;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Normalize a tick-label number format to a value the writer can emit.
 * An empty `formatCode` collapses the whole record — Excel rejects
 * `<c:numFmt formatCode=""/>`.
 */
function normalizeAxisNumberFormat(
  value: ChartAxisNumberFormat | undefined,
): ChartAxisNumberFormat | undefined {
  if (!value) return undefined;
  const formatCode = typeof value.formatCode === "string" ? value.formatCode : "";
  if (formatCode.length === 0) return undefined;
  const out: ChartAxisNumberFormat = { formatCode };
  if (value.sourceLinked === true) out.sourceLinked = true;
  return out;
}

/**
 * Build the children that augment a `<c:scaling>` element. Order is
 * spec-enforced: `<c:logBase>` → `<c:orientation>` → `<c:max>` →
 * `<c:min>`. The orientation child is always emitted by the caller
 * (every axis declares `minMax`); this helper handles the rest.
 *
 * Returns the children to splice in after `<c:orientation>`.
 */
function buildAxisScalingExtras(scale: ChartAxisScale | undefined): {
  before: string[];
  after: string[];
} {
  if (!scale) return { before: [], after: [] };
  const before: string[] = [];
  const after: string[] = [];
  // logBase comes before orientation per CT_Scaling.
  if (scale.logBase !== undefined) {
    before.push(xmlSelfClose("c:logBase", { val: scale.logBase }));
  }
  // max and min come after orientation, with max first (CT_Scaling).
  if (scale.max !== undefined) after.push(xmlSelfClose("c:max", { val: scale.max }));
  if (scale.min !== undefined) after.push(xmlSelfClose("c:min", { val: scale.min }));
  return { before, after };
}

/**
 * Build the `<c:scaling>` element. Always emits `<c:orientation>` so
 * the axis renders correctly even when no extra scale fields are set.
 */
function buildAxisScaling(scale: ChartAxisScale | undefined): string {
  const { before, after } = buildAxisScalingExtras(scale);
  const children: string[] = [
    ...before,
    xmlSelfClose("c:orientation", { val: "minMax" }),
    ...after,
  ];
  return xmlElement("c:scaling", undefined, children);
}

/**
 * Build the optional `<c:majorUnit>` / `<c:minorUnit>` siblings that
 * sit later in the axis-element child sequence (after `<c:numFmt>`,
 * before `<c:crossAx>` per CT_CatAx / CT_ValAx).
 */
function buildAxisTickUnits(scale: ChartAxisScale | undefined): string[] {
  if (!scale) return [];
  const out: string[] = [];
  if (scale.majorUnit !== undefined) {
    out.push(xmlSelfClose("c:majorUnit", { val: scale.majorUnit }));
  }
  if (scale.minorUnit !== undefined) {
    out.push(xmlSelfClose("c:minorUnit", { val: scale.minorUnit }));
  }
  return out;
}

/**
 * Build the axis tick-label `<c:numFmt formatCode=".." sourceLinked=".."/>`.
 * Returns an empty array when the axis declares no number format — the
 * writer then leaves Excel's default linked behaviour untouched.
 */
function buildAxisNumFmt(numFmt: ChartAxisNumberFormat | undefined): string[] {
  if (!numFmt) return [];
  const sourceLinked = numFmt.sourceLinked === true ? 1 : 0;
  return [xmlSelfClose("c:numFmt", { formatCode: numFmt.formatCode, sourceLinked })];
}

// ── Bar / Column ─────────────────────────────────────────────────────

const AXIS_ID_CAT = 111111111;
const AXIS_ID_VAL = 222222222;
const AXIS_ID_VAL_X = 333333333;
const AXIS_ID_VAL_Y = 444444444;

function buildBarChart(chart: SheetChart, sheetName: string): string {
  const grouping = chart.barGrouping ?? "clustered";
  const barDir = chart.type === "bar" ? "bar" : "col";
  const isStacked = grouping === "percentStacked" || grouping === "stacked";

  const children: string[] = [
    xmlSelfClose("c:barDir", { val: barDir }),
    xmlSelfClose("c:grouping", { val: grouping }),
    xmlSelfClose("c:varyColors", { val: 0 }),
  ];

  for (let i = 0; i < chart.series.length; i++) {
    children.push(
      buildSeries(chart.series[i], i, sheetName, /* numericCategories */ false, {
        dataLabels: chart.dataLabels,
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
function clampGapWidth(value: number | undefined): number | undefined {
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
function clampOverlap(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  if (rounded < -100) return -100;
  if (rounded > 100) return 100;
  return rounded;
}

function buildBarAxes(orientation: "bar" | "column", opts: AxisRenderOptions): string[] {
  // For a vertical column chart, categories sit on the bottom (catAx)
  // and values run vertically (valAx). For a horizontal bar chart the
  // axes swap orientation.
  const catPos = orientation === "column" ? "b" : "l";
  const valPos = orientation === "column" ? "l" : "b";

  // OOXML enforces a strict child order inside <c:catAx>/<c:valAx>:
  // axId → scaling → delete → axPos → majorGridlines → minorGridlines
  // → title → numFmt → ... → crossAx → crosses → ... → majorUnit →
  // minorUnit. Each block below mirrors that order.
  // The category axis on bar/column rarely uses scaling, but Excel
  // tolerates the augmentation either way; surface it whenever the
  // caller pinned a value so write-side templates round-trip.
  const catAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_CAT }),
    buildAxisScaling(opts.xScale),
    xmlSelfClose("c:delete", { val: 0 }),
    xmlSelfClose("c:axPos", { val: catPos }),
    ...buildAxisGridlines(opts.xGridlines),
  ];
  if (opts.xAxisTitle) catAxChildren.push(buildAxisTitle(opts.xAxisTitle));
  catAxChildren.push(
    ...buildAxisNumFmt(opts.xNumFmt),
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL }),
    xmlSelfClose("c:crosses", { val: "autoZero" }),
    xmlSelfClose("c:auto", { val: 1 }),
    xmlSelfClose("c:lblAlgn", { val: "ctr" }),
    xmlSelfClose("c:lblOffset", { val: 100 }),
    xmlSelfClose("c:noMultiLvlLbl", { val: 0 }),
  );

  const valAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL }),
    buildAxisScaling(opts.yScale),
    xmlSelfClose("c:delete", { val: 0 }),
    xmlSelfClose("c:axPos", { val: valPos }),
    ...buildAxisGridlines(opts.yGridlines),
  ];
  if (opts.yAxisTitle) valAxChildren.push(buildAxisTitle(opts.yAxisTitle));
  valAxChildren.push(
    ...buildAxisNumFmt(opts.yNumFmt),
    xmlSelfClose("c:crossAx", { val: AXIS_ID_CAT }),
    xmlSelfClose("c:crosses", { val: "autoZero" }),
    xmlSelfClose("c:crossBetween", { val: "between" }),
    ...buildAxisTickUnits(opts.yScale),
  );

  return [
    xmlElement("c:catAx", undefined, catAxChildren),
    xmlElement("c:valAx", undefined, valAxChildren),
  ];
}

// ── Line ─────────────────────────────────────────────────────────────

function buildLineChart(chart: SheetChart, sheetName: string): string {
  const grouping = chart.lineGrouping ?? "standard";
  const children: string[] = [
    xmlSelfClose("c:grouping", { val: grouping }),
    xmlSelfClose("c:varyColors", { val: 0 }),
  ];

  for (let i = 0; i < chart.series.length; i++) {
    const seriesXml = buildSeries(chart.series[i], i, sheetName, /* numericCategories */ false, {
      smooth: false,
      dataLabels: chart.dataLabels,
    });
    children.push(seriesXml);
  }

  const chartLevelDLbls = buildChartLevelDataLabels(chart);
  if (chartLevelDLbls) children.push(chartLevelDLbls);

  children.push(xmlSelfClose("c:marker", { val: 1 }));
  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_CAT }));
  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_VAL }));

  return xmlElement("c:lineChart", undefined, children);
}

// ── Area ─────────────────────────────────────────────────────────────

function buildAreaChart(chart: SheetChart, sheetName: string): string {
  const grouping = chart.areaGrouping ?? "standard";
  const children: string[] = [
    xmlSelfClose("c:grouping", { val: grouping }),
    xmlSelfClose("c:varyColors", { val: 0 }),
  ];

  for (let i = 0; i < chart.series.length; i++) {
    children.push(
      buildSeries(chart.series[i], i, sheetName, /* numericCategories */ false, {
        dataLabels: chart.dataLabels,
      }),
    );
  }

  const chartLevelDLbls = buildChartLevelDataLabels(chart);
  if (chartLevelDLbls) children.push(chartLevelDLbls);

  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_CAT }));
  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_VAL }));

  return xmlElement("c:areaChart", undefined, children);
}

// ── Pie ──────────────────────────────────────────────────────────────

function buildPieChart(chart: SheetChart, sheetName: string): string {
  const children: string[] = [xmlSelfClose("c:varyColors", { val: 1 })];

  // A pie chart only paints the first series; additional ones are
  // valid OOXML but Excel ignores them.
  if (chart.series.length > 0) {
    children.push(
      buildSeries(chart.series[0], 0, sheetName, /* numericCategories */ false, {
        dataLabels: chart.dataLabels,
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

// ── Doughnut ─────────────────────────────────────────────────────────

const DOUGHNUT_HOLE_DEFAULT = 50;
const DOUGHNUT_HOLE_MIN = 10;
const DOUGHNUT_HOLE_MAX = 90;

function buildDoughnutChart(chart: SheetChart, sheetName: string): string {
  const children: string[] = [xmlSelfClose("c:varyColors", { val: 1 })];

  // Like pie, doughnut paints every declared series — Excel renders
  // each as a concentric ring (rare in practice; most templates have
  // one). Carry every series through so multi-ring templates round-trip.
  for (let i = 0; i < chart.series.length; i++) {
    children.push(
      buildSeries(chart.series[i], i, sheetName, /* numericCategories */ false, {
        dataLabels: chart.dataLabels,
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
function clampFirstSliceAng(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  // Wrap into 0..360 (inclusive). The OOXML schema actually allows
  // 360 as a value, so we keep it distinct from 0.
  let normalized = rounded % 360;
  if (normalized < 0) normalized += 360;
  if (normalized === 0) return undefined;
  return normalized;
}

function clampHoleSize(value: number | undefined): number {
  if (value === undefined || !Number.isFinite(value)) return DOUGHNUT_HOLE_DEFAULT;
  const rounded = Math.round(value);
  if (rounded < DOUGHNUT_HOLE_MIN) return DOUGHNUT_HOLE_MIN;
  if (rounded > DOUGHNUT_HOLE_MAX) return DOUGHNUT_HOLE_MAX;
  return rounded;
}

// ── Scatter ──────────────────────────────────────────────────────────

function buildScatterChart(chart: SheetChart, sheetName: string): string {
  const children: string[] = [
    xmlSelfClose("c:scatterStyle", { val: "lineMarker" }),
    xmlSelfClose("c:varyColors", { val: 0 }),
  ];

  for (let i = 0; i < chart.series.length; i++) {
    children.push(
      buildSeries(chart.series[i], i, sheetName, /* numericCategories */ true, {
        dataLabels: chart.dataLabels,
      }),
    );
  }

  const chartLevelDLbls = buildChartLevelDataLabels(chart);
  if (chartLevelDLbls) children.push(chartLevelDLbls);

  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_VAL_X }));
  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_VAL_Y }));

  return xmlElement("c:scatterChart", undefined, children);
}

function buildScatterAxes(opts: AxisRenderOptions): string[] {
  const xAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL_X }),
    buildAxisScaling(opts.xScale),
    xmlSelfClose("c:delete", { val: 0 }),
    xmlSelfClose("c:axPos", { val: "b" }),
    ...buildAxisGridlines(opts.xGridlines),
  ];
  if (opts.xAxisTitle) xAxChildren.push(buildAxisTitle(opts.xAxisTitle));
  xAxChildren.push(
    ...buildAxisNumFmt(opts.xNumFmt),
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL_Y }),
    xmlSelfClose("c:crosses", { val: "autoZero" }),
    xmlSelfClose("c:crossBetween", { val: "midCat" }),
    ...buildAxisTickUnits(opts.xScale),
  );

  const yAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL_Y }),
    buildAxisScaling(opts.yScale),
    xmlSelfClose("c:delete", { val: 0 }),
    xmlSelfClose("c:axPos", { val: "l" }),
    ...buildAxisGridlines(opts.yGridlines),
  ];
  if (opts.yAxisTitle) yAxChildren.push(buildAxisTitle(opts.yAxisTitle));
  yAxChildren.push(
    ...buildAxisNumFmt(opts.yNumFmt),
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL_X }),
    xmlSelfClose("c:crosses", { val: "autoZero" }),
    xmlSelfClose("c:crossBetween", { val: "midCat" }),
    ...buildAxisTickUnits(opts.yScale),
  );

  return [
    xmlElement("c:valAx", undefined, xAxChildren),
    xmlElement("c:valAx", undefined, yAxChildren),
  ];
}

/**
 * Build a `<c:title>` for an axis. The structure mirrors the chart-
 * level title but renders the label at a smaller default font (10pt vs
 * 14pt) to match Excel's axis-title style.
 */
function buildAxisTitle(label: string): string {
  return xmlElement("c:title", undefined, [
    xmlElement("c:tx", undefined, [
      xmlElement("c:rich", undefined, [
        xmlElement(
          "a:bodyPr",
          {
            rot: 0,
            spcFirstLastPara: 1,
            vertOverflow: "ellipsis",
            wrap: "square",
            anchor: "ctr",
            anchorCtr: 1,
          },
          [],
        ),
        xmlSelfClose("a:lstStyle"),
        xmlElement("a:p", undefined, [
          xmlElement("a:pPr", undefined, [xmlSelfClose("a:defRPr", { sz: 1000, b: 0 })]),
          xmlElement("a:r", undefined, [
            xmlSelfClose("a:rPr", { lang: "en-US", sz: 1000, b: 0 }),
            xmlElement("a:t", undefined, xmlEscape(label)),
          ]),
        ]),
      ]),
    ]),
    xmlSelfClose("c:overlay", { val: 0 }),
  ]);
}

// ── Series ───────────────────────────────────────────────────────────

interface SeriesOptions {
  smooth?: boolean;
  /**
   * Chart-level data label defaults from {@link SheetChart.dataLabels}.
   * Used when the series itself does not specify `dataLabels`. Series
   * passing `dataLabels: false` always wins over this default.
   */
  dataLabels?: ChartDataLabels;
}

function buildSeries(
  series: ChartSeries,
  index: number,
  sheetName: string,
  numericCategories: boolean,
  options?: SeriesOptions,
): string {
  const children: string[] = [
    xmlSelfClose("c:idx", { val: index }),
    xmlSelfClose("c:order", { val: index }),
  ];

  if (series.name) {
    // Literal series names go inside <c:tx><c:v>…</c:v></c:tx>. Excel
    // also accepts <c:strRef> for cell-bound names; literals are the
    // simpler shape and round-trip just as well.
    children.push(
      xmlElement("c:tx", undefined, [xmlElement("c:v", undefined, xmlEscape(series.name))]),
    );
  }

  // Optional fill color
  if (series.color) {
    children.push(buildSpPr(series.color));
  }

  // Data labels — series-level override always wins over the chart-level
  // default. `<c:dLbls>` sits between <c:spPr> and <c:cat>/<c:val> per
  // the OOXML series schema (CT_BarSer, CT_LineSer, ...).
  const seriesDLblsXml = buildSeriesDataLabels(series.dataLabels, options?.dataLabels);
  if (seriesDLblsXml) children.push(seriesDLblsXml);

  // Categories (skipped for pie when omitted; allowed for all)
  if (series.categories) {
    const ref = qualifyRef(series.categories, sheetName);
    if (numericCategories) {
      children.push(
        xmlElement("c:xVal", undefined, [
          xmlElement("c:numRef", undefined, [xmlElement("c:f", undefined, xmlEscape(ref))]),
        ]),
      );
    } else {
      children.push(
        xmlElement("c:cat", undefined, [
          xmlElement("c:strRef", undefined, [xmlElement("c:f", undefined, xmlEscape(ref))]),
        ]),
      );
    }
  }

  // Values
  const valuesRef = qualifyRef(series.values, sheetName);
  if (numericCategories) {
    children.push(
      xmlElement("c:yVal", undefined, [
        xmlElement("c:numRef", undefined, [xmlElement("c:f", undefined, xmlEscape(valuesRef))]),
      ]),
    );
  } else {
    children.push(
      xmlElement("c:val", undefined, [
        xmlElement("c:numRef", undefined, [xmlElement("c:f", undefined, xmlEscape(valuesRef))]),
      ]),
    );
  }

  if (options?.smooth !== undefined) {
    children.push(xmlSelfClose("c:smooth", { val: options.smooth ? 1 : 0 }));
  }

  return xmlElement("c:ser", undefined, children);
}

function buildSpPr(rgbHex: string): string {
  const normalized = rgbHex.replace(/^#/, "").toUpperCase();
  return xmlElement("c:spPr", undefined, [
    xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: normalized })]),
    xmlElement("a:ln", undefined, [
      xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: normalized })]),
    ]),
  ]);
}

// ── Data Labels ──────────────────────────────────────────────────────

/**
 * Resolve and emit the `<c:dLbls>` element for a single series.
 *
 * Series override semantics:
 *
 * - Series sets `dataLabels: false`  → emit a `delete=1` block to
 *   suppress this series even when the chart-level default enables labels.
 * - Series sets `dataLabels: <obj>`  → emit `<obj>`. Chart-level config is ignored.
 * - Series omits `dataLabels`        → no per-series `<c:dLbls>`. Excel
 *   inherits the chart-type-level `<c:dLbls>` block emitted by
 *   `buildChartLevelDataLabels` instead.
 *
 * Returns `undefined` when nothing should be emitted at the series level.
 */
function buildSeriesDataLabels(
  seriesDLbls: ChartDataLabels | false | undefined,
  chartDLbls: ChartDataLabels | undefined,
): string | undefined {
  if (seriesDLbls === false) {
    // Suppress this series even when chart-level labels are on.
    return xmlElement("c:dLbls", undefined, [
      xmlElement("c:dLbl", undefined, [
        xmlSelfClose("c:idx", { val: 0 }),
        xmlSelfClose("c:delete", { val: 1 }),
      ]),
      xmlSelfClose("c:delete", { val: 1 }),
    ]);
  }
  if (seriesDLbls) {
    return buildDataLabelsBody(seriesDLbls);
  }
  // Series doesn't override → fall through to chart-level. Returning
  // undefined here keeps the chart-level <c:dLbls> as the single source
  // of truth so we don't duplicate the same toggles N times.
  void chartDLbls;
  return undefined;
}

/**
 * Build the chart-type-level `<c:dLbls>` block from
 * {@link SheetChart.dataLabels}. Returns `undefined` when no chart-level
 * labels are configured.
 */
function buildChartLevelDataLabels(chart: SheetChart): string | undefined {
  if (!chart.dataLabels) return undefined;
  return buildDataLabelsBody(chart.dataLabels);
}

/**
 * Render the OOXML `<c:dLbls>` body. Element order follows CT_DLbls:
 * delete? before numFmt? before spPr? before txPr? before dLblPos? before
 * showLegendKey, showVal, showCatName, showSerName, showPercent,
 * showBubbleSize, separator?, showLeaderLines? — toggles must appear
 * in that exact order or Excel ignores the block.
 */
function buildDataLabelsBody(dl: ChartDataLabels): string {
  const children: string[] = [];

  if (dl.position) {
    children.push(xmlSelfClose("c:dLblPos", { val: dl.position }));
  }

  // OOXML requires showLegendKey to appear first when any toggle is set.
  // Always emit it explicitly so the rendered XML is deterministic.
  children.push(xmlSelfClose("c:showLegendKey", { val: 0 }));
  children.push(xmlSelfClose("c:showVal", { val: dl.showValue ? 1 : 0 }));
  children.push(xmlSelfClose("c:showCatName", { val: dl.showCategoryName ? 1 : 0 }));
  children.push(xmlSelfClose("c:showSerName", { val: dl.showSeriesName ? 1 : 0 }));
  children.push(xmlSelfClose("c:showPercent", { val: dl.showPercent ? 1 : 0 }));
  children.push(xmlSelfClose("c:showBubbleSize", { val: 0 }));

  if (dl.separator !== undefined) {
    children.push(xmlElement("c:separator", undefined, xmlEscape(dl.separator)));
  }

  return xmlElement("c:dLbls", undefined, children);
}

// ── Legend ───────────────────────────────────────────────────────────

type LegendPos = "t" | "b" | "l" | "r" | "tr";

function resolveLegendPosition(chart: SheetChart): LegendPos | null {
  if (chart.legend === false) return null;
  if (chart.legend === undefined) {
    // Sensible defaults that match Excel's behaviour.
    return chart.type === "scatter" ? "b" : "r";
  }
  switch (chart.legend) {
    case "top":
      return "t";
    case "bottom":
      return "b";
    case "left":
      return "l";
    case "right":
      return "r";
    case "topRight":
      return "tr";
  }
}

function buildLegend(pos: LegendPos): string {
  return xmlElement("c:legend", undefined, [
    xmlSelfClose("c:legendPos", { val: pos }),
    xmlSelfClose("c:overlay", { val: 0 }),
  ]);
}

// ── Reference qualification ──────────────────────────────────────────

/**
 * Ensure a range reference is sheet-qualified. Excel chart `<c:f>`
 * elements accept either `Sheet1!$A$2:$A$10` or the unquoted form
 * `Sheet1!A2:A10`; the input is preserved when a sheet is already
 * present. Bare ranges like `B2:B10` are auto-qualified with the
 * owning sheet's name.
 */
function qualifyRef(ref: string, sheetName: string): string {
  if (ref.includes("!")) return ref;
  return `${quoteSheetName(sheetName)}!${ref}`;
}

/**
 * Quote a sheet name when it contains characters Excel considers
 * unsafe in a 3D reference (whitespace, punctuation, etc.). Single
 * quotes inside the name are doubled per the OOXML spec.
 */
function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

// ── Helpers exposed for the drawing layer ────────────────────────────

/**
 * Return the chart-kind labels in declaration order. Useful for
 * tests that need to assert the rendered XML carries the expected
 * `<c:barChart>` / `<c:lineChart>` element.
 */
export function chartKindElement(kind: WriteChartKind): string {
  switch (kind) {
    case "bar":
    case "column":
      return "c:barChart";
    case "line":
      return "c:lineChart";
    case "pie":
      return "c:pieChart";
    case "doughnut":
      return "c:doughnutChart";
    case "scatter":
      return "c:scatterChart";
    case "area":
      return "c:areaChart";
  }
}

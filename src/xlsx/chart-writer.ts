// ── Chart Writer ─────────────────────────────────────────────────────
// Generates xl/charts/chartN.xml for native Excel chart creation.
//
// Phase 1 of issue #152: bar / column / line / pie / scatter / area.
// The chart XML follows the DrawingML chart spec (ECMA-376 Part 1,
// Chapter 21). Each chart is a self-contained <c:chartSpace> document
// referenced from a drawing part via a `chart` relationship.

import type {
  ChartAxisGridlines,
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

  // Axis titles and gridlines surface for every chart family except
  // pie. Pull them once so each branch can hand them off to the
  // matching axis builder.
  const xAxisTitle = normalizeAxisTitle(chart.axes?.x?.title);
  const yAxisTitle = normalizeAxisTitle(chart.axes?.y?.title);
  const xGridlines = normalizeAxisGridlines(chart.axes?.x?.gridlines);
  const yGridlines = normalizeAxisGridlines(chart.axes?.y?.gridlines);

  switch (chart.type) {
    case "bar":
    case "column": {
      children.push(buildBarChart(chart, sheetName));
      children.push(...buildBarAxes(chart.type, xAxisTitle, yAxisTitle, xGridlines, yGridlines));
      break;
    }
    case "line": {
      children.push(buildLineChart(chart, sheetName));
      children.push(...buildBarAxes("column", xAxisTitle, yAxisTitle, xGridlines, yGridlines));
      break;
    }
    case "area": {
      children.push(buildAreaChart(chart, sheetName));
      children.push(...buildBarAxes("column", xAxisTitle, yAxisTitle, xGridlines, yGridlines));
      break;
    }
    case "pie": {
      children.push(buildPieChart(chart, sheetName));
      break;
    }
    case "scatter": {
      children.push(buildScatterChart(chart, sheetName));
      children.push(...buildScatterAxes(xAxisTitle, yAxisTitle, xGridlines, yGridlines));
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

// ── Bar / Column ─────────────────────────────────────────────────────

const AXIS_ID_CAT = 111111111;
const AXIS_ID_VAL = 222222222;
const AXIS_ID_VAL_X = 333333333;
const AXIS_ID_VAL_Y = 444444444;

function buildBarChart(chart: SheetChart, sheetName: string): string {
  const grouping = chart.barGrouping ?? "clustered";
  const barDir = chart.type === "bar" ? "bar" : "col";

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

  if (grouping === "percentStacked" || grouping === "stacked") {
    children.push(xmlSelfClose("c:overlap", { val: 100 }));
  } else {
    children.push(xmlSelfClose("c:gapWidth", { val: 150 }));
  }

  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_CAT }));
  children.push(xmlSelfClose("c:axId", { val: AXIS_ID_VAL }));

  return xmlElement("c:barChart", undefined, children);
}

function buildBarAxes(
  orientation: "bar" | "column",
  xAxisTitle: string | undefined,
  yAxisTitle: string | undefined,
  xGridlines: { major: boolean; minor: boolean } | undefined,
  yGridlines: { major: boolean; minor: boolean } | undefined,
): string[] {
  // For a vertical column chart, categories sit on the bottom (catAx)
  // and values run vertically (valAx). For a horizontal bar chart the
  // axes swap orientation.
  const catPos = orientation === "column" ? "b" : "l";
  const valPos = orientation === "column" ? "l" : "b";

  // OOXML enforces a strict child order inside <c:catAx>/<c:valAx>:
  // axId → scaling → delete → axPos → majorGridlines → minorGridlines
  // → title → numFmt → crossAx → ...
  // Gridlines and title must therefore land before crossAx/crosses,
  // and gridlines must come before the title or Excel ignores them.
  const catAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_CAT }),
    xmlElement("c:scaling", undefined, [xmlSelfClose("c:orientation", { val: "minMax" })]),
    xmlSelfClose("c:delete", { val: 0 }),
    xmlSelfClose("c:axPos", { val: catPos }),
    ...buildAxisGridlines(xGridlines),
  ];
  if (xAxisTitle) catAxChildren.push(buildAxisTitle(xAxisTitle));
  catAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL }),
    xmlSelfClose("c:crosses", { val: "autoZero" }),
    xmlSelfClose("c:auto", { val: 1 }),
    xmlSelfClose("c:lblAlgn", { val: "ctr" }),
    xmlSelfClose("c:lblOffset", { val: 100 }),
    xmlSelfClose("c:noMultiLvlLbl", { val: 0 }),
  );

  const valAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL }),
    xmlElement("c:scaling", undefined, [xmlSelfClose("c:orientation", { val: "minMax" })]),
    xmlSelfClose("c:delete", { val: 0 }),
    xmlSelfClose("c:axPos", { val: valPos }),
    ...buildAxisGridlines(yGridlines),
  ];
  if (yAxisTitle) valAxChildren.push(buildAxisTitle(yAxisTitle));
  valAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_CAT }),
    xmlSelfClose("c:crosses", { val: "autoZero" }),
    xmlSelfClose("c:crossBetween", { val: "between" }),
  );

  return [
    xmlElement("c:catAx", undefined, catAxChildren),
    xmlElement("c:valAx", undefined, valAxChildren),
  ];
}

// ── Line ─────────────────────────────────────────────────────────────

function buildLineChart(chart: SheetChart, sheetName: string): string {
  const children: string[] = [
    xmlSelfClose("c:grouping", { val: "standard" }),
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
  const children: string[] = [
    xmlSelfClose("c:grouping", { val: "standard" }),
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

  return xmlElement("c:pieChart", undefined, children);
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

function buildScatterAxes(
  xAxisTitle: string | undefined,
  yAxisTitle: string | undefined,
  xGridlines: { major: boolean; minor: boolean } | undefined,
  yGridlines: { major: boolean; minor: boolean } | undefined,
): string[] {
  const xAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL_X }),
    xmlElement("c:scaling", undefined, [xmlSelfClose("c:orientation", { val: "minMax" })]),
    xmlSelfClose("c:delete", { val: 0 }),
    xmlSelfClose("c:axPos", { val: "b" }),
    ...buildAxisGridlines(xGridlines),
  ];
  if (xAxisTitle) xAxChildren.push(buildAxisTitle(xAxisTitle));
  xAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL_Y }),
    xmlSelfClose("c:crosses", { val: "autoZero" }),
    xmlSelfClose("c:crossBetween", { val: "midCat" }),
  );

  const yAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL_Y }),
    xmlElement("c:scaling", undefined, [xmlSelfClose("c:orientation", { val: "minMax" })]),
    xmlSelfClose("c:delete", { val: 0 }),
    xmlSelfClose("c:axPos", { val: "l" }),
    ...buildAxisGridlines(yGridlines),
  ];
  if (yAxisTitle) yAxChildren.push(buildAxisTitle(yAxisTitle));
  yAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL_X }),
    xmlSelfClose("c:crosses", { val: "autoZero" }),
    xmlSelfClose("c:crossBetween", { val: "midCat" }),
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
    case "scatter":
      return "c:scatterChart";
    case "area":
      return "c:areaChart";
  }
}

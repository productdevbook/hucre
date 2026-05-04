// ── Chart Writer ─────────────────────────────────────────────────────
// Generates xl/charts/chartN.xml for native Excel chart creation.
//
// Phase 1 of issue #152: bar / column / line / pie / scatter / area.
// The chart XML follows the DrawingML chart spec (ECMA-376 Part 1,
// Chapter 21). Each chart is a self-contained <c:chartSpace> document
// referenced from a drawing part via a `chart` relationship.

import type {
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
  ChartDataLabels,
  ChartDisplayBlanksAs,
  ChartLineDashStyle,
  ChartLineStroke,
  ChartMarker,
  ChartMarkerSymbol,
  ChartProtection,
  ChartScatterStyle,
  ChartSeries,
  ChartView3D,
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
    chartChildren.push(
      buildTitle(chart.title, resolveTitleOverlay(chart), resolveTitleRotation(chart)),
    );
  }
  // `<c:autoTitleDeleted>` records whether the user explicitly deleted
  // Excel's auto-generated title (the synthesised series-name title
  // single-series charts grow). The element sits on `<c:chart>`
  // directly (between `<c:title>` and `<c:plotArea>` per CT_Chart,
  // ECMA-376 Part 1, §21.2.2.4) and is independent of whether a
  // literal `<c:title>` is emitted — a chart with no title may pin
  // `val="1"` to suppress the auto-title or `val="0"` to let Excel
  // synthesise one.
  //
  // Defaults derive from the title presence so back-compat holds: a
  // chart with a literal title emits `val="0"` (Excel keeps the
  // literal visible) and a chart with no literal title emits
  // `val="1"` (Excel does not silently grow an auto-title from the
  // series name). The caller can override the derivation via
  // `autoTitleDeleted` — pin `false` on a titleless single-series
  // column chart to let Excel synthesise the series-name title, or
  // `true` on a charted dashboard tile that should stay anonymous
  // even if a literal title is emitted.
  chartChildren.push(
    xmlSelfClose("c:autoTitleDeleted", { val: resolveAutoTitleDeleted(chart) ? 1 : 0 }),
  );

  // `<c:view3D>` (CT_View3D, ECMA-376 Part 1, §21.2.2.228) sits on
  // `<c:chart>` between `<c:autoTitleDeleted>` / `<c:pivotFmts>` and
  // `<c:floor>` / `<c:plotArea>`. The element is only meaningful on
  // 3D chart families but the OOXML schema accepts it on every
  // CT_Chart, so the writer emits it whenever the caller pins a
  // non-empty configuration — Excel silently ignores it on 2D
  // families. Useful primarily for round-tripping a 3D template chart
  // through cloneChart. The writer skips emission entirely when the
  // caller leaves `view3D` unset so a fresh chart matches Excel's
  // reference serialization byte-for-byte.
  const view3DXml = buildView3D(chart.view3D);
  if (view3DXml !== undefined) {
    chartChildren.push(view3DXml);
  }

  // ── Plot Area ──
  chartChildren.push(buildPlotArea(chart, sheetName));

  // ── Legend ──
  if (legendPos) {
    chartChildren.push(
      buildLegend(legendPos, resolveLegendOverlay(chart), resolveLegendEntries(chart)),
    );
  }

  chartChildren.push(xmlSelfClose("c:plotVisOnly", { val: resolvePlotVisOnly(chart) ? 1 : 0 }));
  chartChildren.push(xmlSelfClose("c:dispBlanksAs", { val: resolveDispBlanksAs(chart) }));
  // `<c:showDLblsOverMax>` sits at the tail of CT_Chart per ECMA-376
  // Part 1, §21.2.2.29 (after `<c:dispBlanksAs>` and before
  // `<c:extLst>`). The writer always emits the element so the rendered
  // intent is explicit on roundtrip — Excel itself includes it in every
  // reference serialization. Mirrors the always-emit contract `<c:plotVisOnly>`
  // and `<c:dispBlanksAs>` follow.
  chartChildren.push(
    xmlSelfClose("c:showDLblsOverMax", { val: resolveShowDLblsOverMax(chart) ? 1 : 0 }),
  );

  const chartElement = xmlElement("c:chart", undefined, chartChildren);

  // `<c:chartSpace>` element ordering per CT_ChartSpace
  // (ECMA-376 Part 1, §21.2.2.29): date1904?, lang?, roundedCorners?,
  // AlternateContent?, clrMapOvr?, style?, ... chart, ...
  // — `<c:date1904>` sits at the head of the sequence, `<c:lang>` next
  // (between `<c:date1904>` and `<c:roundedCorners>`), and `<c:style>`
  // after `<c:roundedCorners>` and before `<c:chart>`. The writer
  // skips emission for any element the chart leaves unset so a fresh
  // chart stays minimal; Excel itself falls back to the workbook's
  // date system / editing language / application default look
  // respectively.
  const chartSpaceChildren: string[] = [];
  if (resolveDate1904(chart)) {
    // `<c:date1904 val="0"/>` is the OOXML default — skip emission so
    // the rendered shape matches absence (every other chart-space
    // toggle follows the same minimal-emission contract). Only the
    // non-default `val="1"` surfaces so a re-parse of the writer's
    // output collapses back to the same `undefined` an unmarked
    // chart parses to.
    chartSpaceChildren.push(xmlSelfClose("c:date1904", { val: 1 }));
  }
  const langVal = resolveLang(chart);
  if (langVal !== undefined) {
    chartSpaceChildren.push(xmlSelfClose("c:lang", { val: langVal }));
  }
  chartSpaceChildren.push(
    xmlSelfClose("c:roundedCorners", { val: resolveRoundedCorners(chart) ? 1 : 0 }),
  );
  const styleVal = resolveStyle(chart);
  if (styleVal !== undefined) {
    chartSpaceChildren.push(xmlSelfClose("c:style", { val: styleVal }));
  }
  // `<c:protection>` (CT_Protection, ECMA-376 Part 1, §21.2.2.142)
  // sits on `<c:chartSpace>` between `<c:style>` / `<c:clrMapOvr>` /
  // `<c:pivotSource>` and `<c:chart>`. The writer skips the element
  // when the caller did not opt in (`undefined` / `false`) and emits
  // it whenever the chart pins `true` or an object — the bare element
  // round-trips when the override is `true` / `{}` because every
  // child is `<xsd:boolean>`-typed and absence of a child is itself
  // valid OOXML (CT_Protection lists every flag as optional).
  const protection = resolveProtection(chart);
  if (protection !== undefined) {
    chartSpaceChildren.push(buildProtection(protection));
  }
  chartSpaceChildren.push(chartElement);

  const chartXml = xmlDocument(
    "c:chartSpace",
    {
      "xmlns:c": NS_C,
      "xmlns:a": NS_A,
      "xmlns:r": NS_R,
    },
    chartSpaceChildren,
  );

  // Always emit an empty rels file. Phase 1 charts do not depend on
  // any other parts (no themeOverride, no userShapes, no embedded
  // spreadsheets), but Excel and several validators expect the file
  // to exist whenever a `chartN.xml` is declared.
  const chartRels = xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, []);

  return { chartXml, chartRels };
}

// ── Title ────────────────────────────────────────────────────────────

function buildTitle(title: string, overlay: boolean, rotationDeg: number | undefined): string {
  // OOXML's `<a:bodyPr rot="N"/>` attribute is in 60000ths of a degree.
  // The writer holds `titleRotation` in whole degrees and converts at
  // emit time. Absence (`undefined`) collapses to the OOXML default
  // `0` so a fresh chart matches Excel's reference serialization
  // byte-for-byte.
  const rot = rotationDeg === undefined ? 0 : rotationDeg * TITLE_ROT_PER_DEGREE;
  return xmlElement("c:title", undefined, [
    xmlElement("c:tx", undefined, [
      xmlElement("c:rich", undefined, [
        xmlElement(
          "a:bodyPr",
          {
            rot,
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
    xmlSelfClose("c:overlay", { val: overlay ? 1 : 0 }),
  ]);
}

/**
 * OOXML's `<a:bodyPr rot="N"/>` attribute is in 60000ths of a degree —
 * the writer holds `titleRotation` in whole degrees and converts at
 * emit time. Excel's UI exposes the `-90..90` band; out-of-band values
 * clamp to the nearest endpoint so a corrupt template cannot leak
 * through to the writer either.
 */
const TITLE_ROT_PER_DEGREE = 60000;
const TITLE_ROTATION_MIN_DEG = -90;
const TITLE_ROTATION_MAX_DEG = 90;

/**
 * Normalize a {@link SheetChart.titleRotation} value (whole degrees)
 * for the `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx>
 * </c:title>` writer slot. Returns `undefined` when the input is unset,
 * non-finite, non-numeric, or resolves to `0` after rounding — every
 * absence path collapses to the same omit-the-attribute shape so
 * absence and the OOXML default `0` round-trip identically through
 * {@link cloneChart}. Out-of-range inputs clamp to the `-90..90` band
 * Excel's UI exposes; non-integer inputs round to the nearest whole
 * degree (the OOXML attribute is an integer in 60000ths of a degree,
 * so a fractional whole-degree value has no meaningful refinement at
 * emit time).
 */
function normalizeTitleRotation(value: number | undefined): number | undefined {
  if (value === undefined || typeof value !== "number" || !Number.isFinite(value)) return undefined;
  let degrees = Math.round(value);
  if (degrees < TITLE_ROTATION_MIN_DEG) degrees = TITLE_ROTATION_MIN_DEG;
  else if (degrees > TITLE_ROTATION_MAX_DEG) degrees = TITLE_ROTATION_MAX_DEG;
  if (degrees === 0) return undefined;
  return degrees;
}

/**
 * Resolve `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx>
 * </c:title>` from {@link SheetChart.titleRotation}.
 *
 * Returns the rotation in whole degrees, or `undefined` when the chart
 * leaves the field unset / pinned the OOXML default `0` / passed a
 * non-numeric or non-finite token. The flag is only meaningful when
 * the chart actually emits a title — the caller is expected to gate
 * the call on `showTitle && chart.title`. A chart whose title is
 * suppressed has no `<c:title>` block to host the rotation in either
 * case.
 */
function resolveTitleRotation(chart: SheetChart): number | undefined {
  return normalizeTitleRotation(chart.titleRotation);
}

/**
 * Resolve `<c:title><c:overlay val=".."/></c:title>` from
 * {@link SheetChart.titleOverlay}.
 *
 * Defaults to `false` (the OOXML default Excel itself emits — the title
 * reserves its own slot above the plot area and the plot area shrinks
 * to make room). Anything other than literal `true` collapses to `false`
 * so a stray non-boolean leaking through the type guard (e.g. `0` / `1` /
 * `"true"` / `null`) never produces `<c:overlay val="1"/>`. This matches
 * how `legendOverlay` / `roundedCorners` / `plotVisOnly` / axis `hidden`
 * treat their inputs: a literal boolean is the only path to a non-default
 * value.
 *
 * The writer always emits `<c:overlay>` inside `<c:title>` because Excel's
 * reference serialization includes the element on every visible title;
 * only the `val` flips when the caller pins `titleOverlay: true`.
 *
 * The flag is only meaningful when the chart actually emits a title — the
 * caller is expected to gate the call on `showTitle && chart.title`. A
 * chart whose title is suppressed has no `<c:title>` block to host the
 * overlay element.
 */
function resolveTitleOverlay(chart: SheetChart): boolean {
  return chart.titleOverlay === true;
}

/**
 * Resolve `<c:autoTitleDeleted val=".."/>` from
 * {@link SheetChart.autoTitleDeleted}.
 *
 * The element records whether the user explicitly deleted the
 * auto-generated title that single-series charts synthesise from the
 * series name. The flag is independent of whether a literal
 * `<c:title>` is emitted — a chart with no title may still pin
 * `val="0"` to let Excel synthesise the auto-title, or `val="1"` to
 * suppress it.
 *
 * When the caller pins {@link SheetChart.autoTitleDeleted} explicitly
 * the literal boolean value wins. When the field is omitted the writer
 * derives the value from the title presence so back-compat holds: a
 * chart with a literal title (and `showTitle !== false`) emits
 * `val="0"` so Excel keeps the literal visible; a chart with no
 * literal title emits `val="1"` so Excel does not silently grow an
 * auto-title from the series name.
 *
 * Anything other than literal `true` / `false` collapses to the
 * derived default so a stray non-boolean leaking through the type
 * guard (e.g. `0` / `1` / `"true"` / `null`) never inverts the
 * derivation. This matches how `titleOverlay` / `roundedCorners` /
 * `plotVisOnly` treat their inputs.
 */
function resolveAutoTitleDeleted(chart: SheetChart): boolean {
  if (chart.autoTitleDeleted === true) return true;
  if (chart.autoTitleDeleted === false) return false;
  // Derive from title presence — preserves back-compat for callers
  // that never set the field.
  const showTitle = chart.showTitle ?? Boolean(chart.title);
  return !(showTitle && chart.title);
}

// ── Plot Area ────────────────────────────────────────────────────────

function buildPlotArea(chart: SheetChart, sheetName: string): string {
  const children: string[] = [xmlSelfClose("c:layout")];

  // Axis titles, gridlines, scaling, number format and tick rendering
  // surface for every chart family except pie/doughnut. Pull them once
  // so each branch can hand them off to the matching axis builder.
  const opts: AxisRenderOptions = {
    xAxisTitle: normalizeAxisTitle(chart.axes?.x?.title),
    yAxisTitle: normalizeAxisTitle(chart.axes?.y?.title),
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
  // `<c:valAx>` / `<c:catAx>` and the optional `<c:spPr>` that the
  // writer never emits. Pie / doughnut have no axes at all, so the
  // OOXML schema places no slot for `<c:dTable>` on those families;
  // `resolveDataTable` short-circuits them by returning `undefined`.
  const dTable = resolveDataTable(chart);
  if (dTable !== undefined) {
    children.push(buildDataTable(dTable));
  }

  return xmlElement("c:plotArea", undefined, children);
}

// ── Data Table ───────────────────────────────────────────────────────

/**
 * Resolve the {@link SheetChart.dataTable} field into the four boolean
 * children `<c:dTable>` requires, or `undefined` to signal that the
 * writer should skip emission of the element.
 *
 * Returns `undefined` when:
 *  - The chart's family has no axes (pie / doughnut). The OOXML schema
 *    places `<c:dTable>` inside `<c:plotArea>` alongside the axes, so
 *    no axes means no slot to host the element.
 *  - The caller did not opt in (`dataTable` is `undefined` or `false`).
 *
 * Returns the four resolved booleans when the caller passed `true`
 * (every default `true`) or an object (per-field overrides on top of the
 * `true` defaults). Stray non-boolean inputs collapse to the matching
 * default rather than emit a token Excel rejects, mirroring how every
 * other chart-level boolean writer treats its input.
 */
function resolveDataTable(chart: SheetChart):
  | {
      showHorzBorder: boolean;
      showVertBorder: boolean;
      showOutline: boolean;
      showKeys: boolean;
    }
  | undefined {
  // Pie / doughnut have no axes — the OOXML schema places `<c:dTable>`
  // alongside `<c:catAx>` / `<c:valAx>`, so there is no slot for it on
  // those families. Drop the field silently rather than emit an element
  // Excel's strict validator would reject.
  if (chart.type === "pie" || chart.type === "doughnut") return undefined;

  const raw = chart.dataTable;
  if (raw === undefined || raw === false) return undefined;

  if (raw === true) {
    return {
      showHorzBorder: true,
      showVertBorder: true,
      showOutline: true,
      showKeys: true,
    };
  }

  // Per-field overrides on top of the `true` defaults. Only literal
  // `false` flips a flag — anything else (including stray `undefined`,
  // `null`, or a non-boolean) falls back to the default `true` so the
  // writer never emits a value the OOXML schema would refuse.
  return {
    showHorzBorder: raw.showHorzBorder !== false,
    showVertBorder: raw.showVertBorder !== false,
    showOutline: raw.showOutline !== false,
    showKeys: raw.showKeys !== false,
  };
}

/**
 * Serialize a resolved data-table into `<c:dTable>` with its four
 * required boolean children, in the order CT_DTable mandates:
 * `showHorzBorder`, `showVertBorder`, `showOutline`, `showKeys`.
 *
 * The writer always emits all four children — the OOXML schema marks
 * them required on `CT_DTable`, and Excel's reference serialization
 * includes every one even when the caller leaves it at the default. The
 * optional `<c:spPr>` / `<c:txPr>` / `<c:extLst>` children are skipped
 * because hucre's data-table model does not surface fill / text styling
 * yet.
 */
function buildDataTable(table: {
  showHorzBorder: boolean;
  showVertBorder: boolean;
  showOutline: boolean;
  showKeys: boolean;
}): string {
  return xmlElement("c:dTable", undefined, [
    xmlSelfClose("c:showHorzBorder", { val: table.showHorzBorder ? 1 : 0 }),
    xmlSelfClose("c:showVertBorder", { val: table.showVertBorder ? 1 : 0 }),
    xmlSelfClose("c:showOutline", { val: table.showOutline ? 1 : 0 }),
    xmlSelfClose("c:showKeys", { val: table.showKeys ? 1 : 0 }),
  ]);
}

// ── Protection ───────────────────────────────────────────────────────

/**
 * Resolve the {@link SheetChart.protection} field into the per-flag
 * shape `<c:protection>` emits, or `undefined` to signal that the
 * writer should skip the element entirely.
 *
 * Returns `undefined` when the caller did not opt in (`protection` is
 * `undefined` or `false`).
 *
 * Returns the resolved per-flag block when the caller passed `true`
 * (every flag at the OOXML default `false` — equivalent to a bare
 * `<c:protection/>` shell) or an object (per-field overrides). Stray
 * non-boolean inputs collapse to `false` (the OOXML default) rather
 * than emit a token Excel rejects, mirroring how every other
 * chart-level boolean writer treats its input.
 *
 * Unlike {@link resolveDataTable}, this resolver applies to every
 * chart family — `<c:protection>` lives on `<c:chartSpace>`, not
 * inside `<c:plotArea>`, so the element has a slot on pie / doughnut
 * charts too.
 */
function resolveProtection(chart: SheetChart):
  | {
      chartObject: boolean;
      data: boolean;
      formatting: boolean;
      selection: boolean;
      userInterface: boolean;
    }
  | undefined {
  const raw = chart.protection;
  if (raw === undefined || raw === false) return undefined;

  if (raw === true) {
    return {
      chartObject: false,
      data: false,
      formatting: false,
      selection: false,
      userInterface: false,
    };
  }

  // Per-field overrides on top of the `false` defaults. Only literal
  // `true` flips a flag — anything else (including stray `undefined`,
  // `null`, or a non-boolean) falls back to the default `false` so the
  // writer never emits a token the OOXML schema would refuse. The
  // empty-object case (`{}`) collapses to a bare `<c:protection/>` with
  // every flag at its default, so Excel still records the chart-level
  // protection block on roundtrip.
  return {
    chartObject: raw.chartObject === true,
    data: raw.data === true,
    formatting: raw.formatting === true,
    selection: raw.selection === true,
    userInterface: raw.userInterface === true,
  };
}

/**
 * Serialize a resolved protection block into `<c:protection>` with its
 * five optional boolean children, in the order CT_Protection mandates:
 * `chartObject`, `data`, `formatting`, `selection`, `userInterface`.
 *
 * Unlike `<c:dTable>` (whose four children are required on
 * CT_DTable), every CT_Protection child is optional — but the writer
 * always emits all five so the rendered intent is explicit on
 * roundtrip. Default-valued (`false`) children still surface as
 * `<c:chartObject val="0"/>` to match the always-emit contract every
 * other chart-level boolean writer follows (compare `<c:plotVisOnly>`
 * and `<c:dispBlanksAs>`). Excel's reader treats a missing child as
 * `false` either way.
 */
function buildProtection(protection: {
  chartObject: boolean;
  data: boolean;
  formatting: boolean;
  selection: boolean;
  userInterface: boolean;
}): string {
  return xmlElement("c:protection", undefined, [
    xmlSelfClose("c:chartObject", { val: protection.chartObject ? 1 : 0 }),
    xmlSelfClose("c:data", { val: protection.data ? 1 : 0 }),
    xmlSelfClose("c:formatting", { val: protection.formatting ? 1 : 0 }),
    xmlSelfClose("c:selection", { val: protection.selection ? 1 : 0 }),
    xmlSelfClose("c:userInterface", { val: protection.userInterface ? 1 : 0 }),
  ]);
}

// ── 3-D View ─────────────────────────────────────────────────────────

/**
 * Serialize a {@link ChartView3D} into `<c:view3D>` with one self-
 * closing child per pinned field, in the order CT_View3D mandates:
 * `<c:rotX>`, `<c:hPercent>`, `<c:rotY>`, `<c:depthPercent>`,
 * `<c:rAngAx>`, `<c:perspective>`. Returns `undefined` when the
 * caller did not opt in (`view3D` is `undefined`) so the writer can
 * skip the element entirely; returns the bare `<c:view3D/>` shell
 * when an empty object is passed (round-trips a template that
 * authored the element with no children pinned).
 *
 * Each numeric field is clamped against the matching OOXML simple-
 * type range — out-of-range and non-finite inputs drop silently
 * rather than emit a token Excel's strict validator would reject.
 * The boolean field surfaces only as a literal `0` / `1` `val`
 * attribute; non-boolean inputs collapse to `false` (the OOXML
 * default), mirroring how every other chart-level boolean writer
 * treats its input.
 */
function buildView3D(view3D: ChartView3D | undefined): string | undefined {
  if (view3D === undefined) return undefined;
  const children: string[] = [];
  // CT_View3D children sequence per ECMA-376 §21.2.2.228:
  // rotX?, hPercent?, rotY?, depthPercent?, rAngAx?, perspective?,
  // extLst?
  const rotX = clampView3DInt(view3D.rotX, -90, 90);
  if (rotX !== undefined) children.push(xmlSelfClose("c:rotX", { val: rotX }));
  const hPercent = clampView3DInt(view3D.hPercent, 5, 500);
  if (hPercent !== undefined) {
    // `<c:hPercent>` accepts the bare integer per ST_HPercent — Excel
    // emits a plain percent value with no `%` suffix.
    children.push(xmlSelfClose("c:hPercent", { val: hPercent }));
  }
  const rotY = clampView3DInt(view3D.rotY, 0, 360);
  if (rotY !== undefined) children.push(xmlSelfClose("c:rotY", { val: rotY }));
  const depthPercent = clampView3DInt(view3D.depthPercent, 20, 2000);
  if (depthPercent !== undefined) {
    children.push(xmlSelfClose("c:depthPercent", { val: depthPercent }));
  }
  if (view3D.rAngAx === true) {
    children.push(xmlSelfClose("c:rAngAx", { val: 1 }));
  } else if (view3D.rAngAx === false) {
    // Explicit `false` round-trips as `<c:rAngAx val="0"/>` so the
    // caller can pin the OOXML default literally — useful for parity
    // with templates that author the explicit value.
    children.push(xmlSelfClose("c:rAngAx", { val: 0 }));
  }
  const perspective = clampView3DInt(view3D.perspective, 0, 240);
  if (perspective !== undefined) {
    children.push(xmlSelfClose("c:perspective", { val: perspective }));
  }
  // Empty object (`{}`) collapses to a bare `<c:view3D/>` shell —
  // `xmlElement` with an empty child array emits the self-closing form.
  return xmlElement("c:view3D", undefined, children);
}

/**
 * Clamp a `<c:view3D>` numeric field against the matching OOXML
 * simple-type range. Returns `undefined` when the input is non-finite,
 * non-integer, or out-of-range — the writer drops the matching child
 * rather than emit a token Excel's strict validator would reject.
 *
 * The strict integer check rejects fractional inputs (`15.5`) so the
 * round-trip stays lossless — `parseView3DInt` on the reader side
 * also rejects fractional `val` attributes, and a fractional input
 * here would silently mismatch on the next parse.
 */
function clampView3DInt(value: number | undefined, min: number, max: number): number | undefined {
  if (typeof value !== "number") return undefined;
  if (!Number.isFinite(value)) return undefined;
  if (!Number.isInteger(value)) return undefined;
  if (value < min || value > max) return undefined;
  return value;
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
  xMajorTickMark: ChartAxisTickMark | undefined;
  yMajorTickMark: ChartAxisTickMark | undefined;
  xMinorTickMark: ChartAxisTickMark | undefined;
  yMinorTickMark: ChartAxisTickMark | undefined;
  xTickLblPos: ChartAxisTickLabelPosition | undefined;
  yTickLblPos: ChartAxisTickLabelPosition | undefined;
  /**
   * Tick-label rotation in whole degrees emitted on the X axis via
   * `<c:txPr><a:bodyPr rot="N"/></c:txPr>`. The OOXML `rot` attribute
   * is in 60000ths of a degree; the writer converts at emit time.
   * Range: `-90..90` (Excel's UI band). `undefined` skips the
   * `<c:txPr>` block entirely so a fresh chart matches Excel's
   * minimal serialization. Surfaces on every axis flavour because the
   * OOXML schema places `<c:txPr>` on `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` alike.
   */
  xLabelRotation: number | undefined;
  /**
   * Tick-label rotation in whole degrees emitted on the Y axis. Same
   * shape and conversion semantics as {@link xLabelRotation}.
   */
  yLabelRotation: number | undefined;
  xReverse: boolean;
  yReverse: boolean;
  /**
   * Tick-label skip interval emitted on the X axis only when the axis
   * is `<c:catAx>` (i.e. bar / column / line / area). Scatter charts
   * have no category axis, so the skip is dropped silently.
   */
  xTickLblSkip: number | undefined;
  /**
   * Tick-mark skip interval emitted on the X axis only when the axis
   * is `<c:catAx>`. Same scope rule as {@link xTickLblSkip}.
   */
  xTickMarkSkip: number | undefined;
  /**
   * Label offset percentage emitted on the X axis only when the axis
   * is `<c:catAx>` (i.e. bar / column / line / area). Scatter charts
   * have no category axis, so the value is dropped silently.
   */
  xLblOffset: number | undefined;
  /**
   * Tick-label horizontal alignment emitted on the X axis only when
   * the axis is `<c:catAx>`. Scatter charts have no category axis, so
   * the value is dropped silently. `undefined` means absent (the
   * writer falls back to the OOXML default `"ctr"`).
   */
  xLblAlgn: ChartAxisLabelAlign | undefined;
  /**
   * Whether the X axis should pin `<c:noMultiLvlLbl val="1"/>`
   * (multi-level category labels suppressed). Always defined — `false`
   * keeps Excel's reference `val="0"` while `true` collapses multi-tier
   * category labels onto a single line. Only meaningful for the catAx
   * builder; scatter has no category axis, so the value is silently
   * dropped at the per-chart-type branch.
   */
  xNoMultiLvlLbl: boolean;
  /**
   * Whether the X axis should render its `<c:auto>` element with
   * `val="1"` (Excel's default — auto-detect whether the axis is a
   * date axis or category axis). Always defined — `true` keeps Excel's
   * reference `val="1"` while `false` pins the axis as a literal
   * category axis (Excel's "Text axis" radio under "Format Axis -> Axis
   * Options"). Only meaningful for the catAx builder; scatter has no
   * category axis, so the value is silently dropped at the per-chart-
   * type branch.
   */
  xAuto: boolean;
  /**
   * Whether the X axis should render its `<c:delete>` element with
   * `val="1"` (axis hidden). Always defined — `false` keeps Excel's
   * reference `val="0"` while `true` collapses the axis line, ticks,
   * and labels off the rendered chart.
   */
  xHidden: boolean;
  /** Whether the Y axis should render hidden. Same shape as {@link xHidden}. */
  yHidden: boolean;
  /**
   * Resolved axis-crosses pin for the X axis. The XSD choice between
   * `<c:crosses>` and `<c:crossesAt>` is collapsed to a single tagged
   * union: `kind: "default"` emits the OOXML default `<c:crosses
   * val="autoZero"/>`, `kind: "semantic"` emits the resolved
   * {@link ChartAxisCrosses} token, and `kind: "numeric"` emits
   * `<c:crossesAt>` with the literal value the caller pinned.
   */
  xCrosses: ResolvedAxisCrosses;
  /** Resolved axis-crosses pin for the Y axis. Same shape as {@link xCrosses}. */
  yCrosses: ResolvedAxisCrosses;
  /**
   * Display-unit preset emitted on the X axis only when the axis is
   * `<c:valAx>` (i.e. scatter charts). Bar / column / line / area route
   * the X axis through `<c:catAx>` which rejects `<c:dispUnits>`, so
   * the catAx builder ignores this field.
   */
  xDispUnits: ChartAxisDispUnits | undefined;
  /**
   * Display-unit preset emitted on the value axis. The catAx builder
   * (bar / column / line / area) routes the Y axis through `<c:valAx>`,
   * and the scatter builder routes both axes through `<c:valAx>` — so
   * this field surfaces on every chart family that has a value axis.
   * Pie / doughnut have no axes at all and the caller already
   * short-circuits those branches.
   */
  yDispUnits: ChartAxisDispUnits | undefined;
  /**
   * Cross-between override for the X axis. Only honoured on scatter
   * (the X axis is a value axis there); the catAx builder ignores it
   * because `<c:crossBetween>` is value-axis-only per ECMA-376
   * §21.2.2.10. `undefined` falls back to the per-family default each
   * axis builder pins today.
   */
  xCrossBetween: ChartAxisCrossBetween | undefined;
  /**
   * Cross-between override for the value axis. The catAx builder (bar
   * / column / line / area) routes the Y axis through `<c:valAx>`, and
   * the scatter builder routes both axes through `<c:valAx>` — so this
   * field surfaces on every chart family that has a value axis. Pie /
   * doughnut have no axes at all and the caller already short-circuits
   * those branches.
   */
  yCrossBetween: ChartAxisCrossBetween | undefined;
}

/**
 * Resolved per-axis crossing pin. The OOXML schema places `<c:crosses>`
 * and `<c:crossesAt>` in an XSD choice — only one may appear at a time.
 * `normalizeAxisCrosses` collapses the writer's two input fields
 * (`crosses` and `crossesAt`) into this tagged union so the per-family
 * axis builders can emit the right element without re-implementing the
 * precedence rule.
 */
type ResolvedAxisCrosses =
  | { kind: "default" }
  | { kind: "semantic"; value: ChartAxisCrosses }
  | { kind: "numeric"; value: number };

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
 * Normalize a `tickLblSkip` / `tickMarkSkip` value to a positive
 * integer in the OOXML `ST_SkipIntervals` band (`1..32767`).
 *
 * Returns `undefined` when:
 *   - the input is missing or non-finite,
 *   - the rounded value is `1` (the OOXML default — show every label /
 *     mark — and what absence already means),
 *   - the rounded value falls outside the `1..32767` range.
 *
 * Out-of-range values drop rather than clamp because a skip count of
 * `100` and `32767` mean structurally different things to Excel — a
 * silent clamp would mask the configuration error rather than reveal
 * it.
 */
function normalizeAxisSkip(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  if (rounded < 1 || rounded > 32767) return undefined;
  if (rounded === 1) return undefined;
  return rounded;
}

/**
 * Normalize a category-axis `lblOffset` percentage to an integer in
 * the OOXML `ST_LblOffsetPercent` band (`0..1000`).
 *
 * Returns `undefined` when:
 *   - the input is missing or non-finite,
 *   - the rounded value is `100` (the OOXML default — Excel's
 *     reference label spacing — and what absence already means),
 *   - the rounded value falls outside the `0..1000` range.
 *
 * Out-of-range values drop rather than clamp so a malformed override
 * surfaces as "no offset emitted" instead of silently snapping to the
 * extreme — a clamp from `9999` to `1000` would mask a programming
 * error in the caller.
 */
function normalizeAxisLblOffset(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  if (rounded < 0 || rounded > 1000) return undefined;
  if (rounded === 100) return undefined;
  return rounded;
}

/**
 * Normalize a category-axis `lblAlgn` value to one of the three OOXML
 * `ST_LblAlgn` tokens (`"ctr"` / `"l"` / `"r"`).
 *
 * Returns `undefined` when:
 *   - the input is missing,
 *   - the value is not in the `ST_LblAlgn` enumeration,
 *   - the value is `"ctr"` (the OOXML default — Excel's reference
 *     centered alignment — and what absence already means).
 *
 * Unknown tokens drop rather than fall back to the default so a
 * malformed override surfaces as "no alignment emitted" instead of
 * silently snapping to `"ctr"` (which would mask the configuration
 * error in the caller).
 */
function normalizeAxisLblAlgn(
  value: ChartAxisLabelAlign | undefined,
): ChartAxisLabelAlign | undefined {
  if (value === undefined) return undefined;
  if (value !== "ctr" && value !== "l" && value !== "r") return undefined;
  if (value === "ctr") return undefined;
  return value;
}

/**
 * Normalize an axis `hidden` flag to a strict boolean. Anything other
 * than literal `true` collapses to `false` so the writer never emits
 * `<c:delete val="1"/>` from a stray non-boolean leaking through the
 * type guard (e.g. `0` / `1` / `"true"` / `null`). This matches how
 * `roundedCorners` / `plotVisOnly` / `varyColors` treat their inputs:
 * a literal boolean is the only path to a non-default value.
 */
function normalizeAxisHidden(value: boolean | undefined): boolean {
  return value === true;
}

/**
 * OOXML's `<a:bodyPr rot="N"/>` attribute is in 60000ths of a degree —
 * the writer holds the value in whole degrees and converts at emit
 * time. Excel's UI exposes the `-90..90` band; out-of-band values clamp
 * to the nearest endpoint so a corrupt template cannot leak through to
 * the writer either.
 */
const TXPR_ROT_PER_DEGREE = 60000;
const LABEL_ROTATION_MIN_DEG = -90;
const LABEL_ROTATION_MAX_DEG = 90;

/**
 * Normalize an axis `labelRotation` value (whole degrees) for the
 * `<c:txPr><a:bodyPr rot="N"/></c:txPr>` writer slot. Returns
 * `undefined` when the input is unset, non-finite, non-numeric, or
 * resolves to `0` after rounding — every absence path collapses to the
 * same omit-the-element shape so absence and the OOXML default `0`
 * round-trip identically through {@link cloneChart}. Out-of-range
 * inputs clamp to the `-90..90` band Excel's UI exposes; non-integer
 * inputs round to the nearest whole degree (the OOXML attribute is an
 * integer in 60000ths of a degree, so a fractional whole-degree value
 * has no meaningful refinement at emit time).
 */
function normalizeAxisLabelRotation(value: number | undefined): number | undefined {
  if (value === undefined || typeof value !== "number" || !Number.isFinite(value)) return undefined;
  let degrees = Math.round(value);
  if (degrees < LABEL_ROTATION_MIN_DEG) degrees = LABEL_ROTATION_MIN_DEG;
  else if (degrees > LABEL_ROTATION_MAX_DEG) degrees = LABEL_ROTATION_MAX_DEG;
  if (degrees === 0) return undefined;
  return degrees;
}

/**
 * Build the `<c:txPr>` block that carries an axis tick-label rotation.
 * Returns `undefined` when the resolved degree value is unset so the
 * caller can elide the element entirely (Excel's reference
 * serialization on a fresh axis omits `<c:txPr>` when the labels
 * render at the default `rot="0"`).
 *
 * The emitted block mirrors the minimal `<c:txPr>` shape Excel writes
 * when the user pins a custom angle — `<a:bodyPr rot="N"/>` carries
 * the rotation, `<a:lstStyle/>` is the empty list-style placeholder
 * the schema requires, and `<a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr/></a:p>`
 * is the empty paragraph stub Excel always emits even when no per-run
 * styling is pinned. Additional `<a:bodyPr>` attributes Excel writes
 * in its full reference (`spcFirstLastPara` / `vertOverflow` / `wrap`
 * / `anchor` / `anchorCtr`) are intentionally omitted — the OOXML
 * schema marks them all optional, and dropping them keeps the writer's
 * footprint minimal while preserving the rotation intent.
 */
function buildAxisTxPr(rotationDeg: number | undefined): string | undefined {
  if (rotationDeg === undefined) return undefined;
  const rot = rotationDeg * TXPR_ROT_PER_DEGREE;
  return xmlElement("c:txPr", undefined, [
    xmlSelfClose("a:bodyPr", { rot }),
    xmlSelfClose("a:lstStyle"),
    xmlElement("a:p", undefined, [
      xmlElement("a:pPr", undefined, [xmlSelfClose("a:defRPr")]),
      xmlSelfClose("a:endParaRPr", { lang: "en-US" }),
    ]),
  ]);
}

/** Recognized values of `<c:crosses>` per the OOXML `ST_Crosses` enum. */
const VALID_AXIS_CROSSES: ReadonlySet<ChartAxisCrosses> = new Set(["autoZero", "min", "max"]);

/**
 * Resolve the writer's `axes.x.crosses` / `axes.x.crossesAt` pair into
 * the {@link ResolvedAxisCrosses} tagged union the per-family axis
 * builders emit. The OOXML schema places `<c:crosses>` and
 * `<c:crossesAt>` in an XSD choice — only one may legally appear at a
 * time per ECMA-376 Part 1, §21.2.2 — so the normalizer collapses the
 * caller's two fields to a single resolved shape:
 *
 *   - A finite numeric `crossesAt` always wins, mirroring how Excel
 *     treats the choice (the explicit numeric pin overrides the
 *     semantic default). Non-finite inputs (NaN / Infinity) drop so the
 *     writer never emits an attribute Excel would reject.
 *   - When only `crosses` is set, the resolved kind is `"semantic"` for
 *     `"min"` / `"max"`. The OOXML default `"autoZero"` collapses to
 *     `kind: "default"` so absence and the default emit the same
 *     `<c:crosses val="autoZero"/>` byte-for-byte. Unknown tokens drop
 *     to `kind: "default"` for the same reason.
 *   - When neither is set, the resolved kind is `"default"` (the writer
 *     still emits `<c:crosses val="autoZero"/>` to match Excel's
 *     reference serialization on every freshly-drawn axis).
 */
function normalizeAxisCrosses(
  semantic: ChartAxisCrosses | undefined,
  numeric: number | undefined,
): ResolvedAxisCrosses {
  if (typeof numeric === "number" && Number.isFinite(numeric)) {
    return { kind: "numeric", value: numeric };
  }
  if (semantic !== undefined && VALID_AXIS_CROSSES.has(semantic) && semantic !== "autoZero") {
    return { kind: "semantic", value: semantic };
  }
  return { kind: "default" };
}

/**
 * Render the resolved axis crossing pin as the matching child element.
 * `kind: "numeric"` emits `<c:crossesAt val=".."/>`; every other kind
 * emits `<c:crosses val=".."/>` so Excel's reference serialization
 * (which always pins `<c:crosses val="autoZero"/>` on every axis) is
 * preserved on freshly-drawn charts.
 */
function buildAxisCrosses(resolved: ResolvedAxisCrosses): string {
  switch (resolved.kind) {
    case "numeric":
      return xmlSelfClose("c:crossesAt", { val: resolved.value });
    case "semantic":
      return xmlSelfClose("c:crosses", { val: resolved.value });
    case "default":
      return xmlSelfClose("c:crosses", { val: "autoZero" });
  }
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
 * the axis renders correctly even when no extra scale fields are set —
 * `"minMax"` (the OOXML default) for a forward axis, `"maxMin"` when
 * the caller pinned `reverse: true` to flip the plotting order.
 */
function buildAxisScaling(scale: ChartAxisScale | undefined, reverse: boolean = false): string {
  const { before, after } = buildAxisScalingExtras(scale);
  const children: string[] = [
    ...before,
    xmlSelfClose("c:orientation", { val: reverse ? "maxMin" : "minMax" }),
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
 * Normalize the {@link SheetChart.axes.x.dispUnits} /
 * {@link SheetChart.axes.y.dispUnits} input — accept either the
 * `ChartAxisDispUnit` shorthand (e.g. `"millions"`) or the full
 * {@link ChartAxisDispUnits} object — into a single canonical shape the
 * writer can hand off to {@link buildAxisDispUnits}. Unknown / typo'd
 * tokens collapse to `undefined` so the writer never emits a value the
 * OOXML `ST_BuiltInUnit` enum rejects. Non-object / non-string inputs
 * (e.g. `null`, numbers, arrays) also collapse to `undefined`.
 *
 * The OOXML schema places `<c:builtInUnit>` and `<c:custUnit>` in an
 * `xsd:choice` — at most one of the two may appear inside `<c:dispUnits>`.
 * The normalizer keeps both fields when the input declares both, leaving
 * the precedence pick to {@link buildAxisDispUnits} (where `custUnit`
 * wins because it is the more specific OOXML element). When neither
 * field is valid, the normalizer returns `undefined` so the writer skips
 * the `<c:dispUnits>` element entirely — a `<c:dispUnits>` shell with
 * no `<c:builtInUnit>` / `<c:custUnit>` child would fail Excel's strict
 * validator.
 */
function normalizeAxisDispUnits(
  value: ChartAxisDispUnits | ChartAxisDispUnit | undefined,
): ChartAxisDispUnits | undefined {
  if (value === undefined) return undefined;
  if (typeof value === "string") {
    return VALID_DISP_UNITS.has(value as ChartAxisDispUnit)
      ? { unit: value as ChartAxisDispUnit }
      : undefined;
  }
  if (typeof value !== "object" || value === null) return undefined;
  const out: ChartAxisDispUnits = {};
  const unit = value.unit;
  if (typeof unit === "string" && VALID_DISP_UNITS.has(unit as ChartAxisDispUnit)) {
    out.unit = unit as ChartAxisDispUnit;
  }
  const custUnit = value.custUnit;
  if (typeof custUnit === "number" && Number.isFinite(custUnit) && custUnit > 0) {
    out.custUnit = custUnit;
  }
  // Drop the entire object when neither child resolves — a bare
  // `<c:dispUnits/>` shell would fail Excel's strict validator (the
  // CT_DispUnits choice has `minOccurs="0"` on the choice itself, but
  // an empty element with the parent's `<c:extLst>` slot also empty
  // is rejected by Excel's reference renderer).
  if (out.unit === undefined && out.custUnit === undefined) return undefined;
  if (value.showLabel === true) out.showLabel = true;
  return out;
}

/**
 * Build the optional `<c:dispUnits>` block that sits at the very end of
 * `<c:valAx>` per CT_ValAx (after `<c:minorUnit>`). The element itself
 * holds the choice between `<c:builtInUnit>` and `<c:custUnit>` — the
 * writer emits exactly one per the OOXML `xsd:choice`, preferring
 * `<c:custUnit>` when both fields are pinned because the more specific
 * numeric divisor takes precedence (a caller appending a custom unit to
 * a cloned source need not manually prune the inherited preset). When
 * `showLabel` is `true` the writer emits a bare `<c:dispUnitsLbl/>` so
 * Excel paints its default automatic annotation; the rich-text label
 * customization is intentionally not surfaced.
 *
 * Returns an empty array when the caller did not pin either child so
 * the writer leaves Excel's default "no display unit" state untouched.
 */
function buildAxisDispUnits(dispUnits: ChartAxisDispUnits | undefined): string[] {
  if (!dispUnits) return [];
  const children: string[] = [];
  if (dispUnits.custUnit !== undefined) {
    children.push(xmlSelfClose("c:custUnit", { val: dispUnits.custUnit }));
  } else if (dispUnits.unit !== undefined) {
    children.push(xmlSelfClose("c:builtInUnit", { val: dispUnits.unit }));
  } else {
    // Neither child resolved — skip emission rather than ship a bare
    // `<c:dispUnits/>` Excel rejects. The normalizer should have
    // pre-filtered this case, but the guard here keeps the writer
    // robust against a stray runtime object slipping past the type
    // boundary.
    return [];
  }
  if (dispUnits.showLabel === true) {
    children.push(xmlSelfClose("c:dispUnitsLbl"));
  }
  return [xmlElement("c:dispUnits", undefined, children)];
}

/** Recognized values of `<c:crossBetween>` per the OOXML `ST_CrossBetween` enum. */
const VALID_CROSS_BETWEEN: ReadonlySet<ChartAxisCrossBetween> = new Set(["between", "midCat"]);

/**
 * Normalize the {@link SheetChart.axes.x.crossBetween} /
 * {@link SheetChart.axes.y.crossBetween} input. Unknown / typo'd tokens
 * collapse to `undefined` so the writer never emits a value the OOXML
 * `ST_CrossBetween` enum rejects — the caller's per-family default
 * (`"between"` on bar / column / line / area Y axes; `"midCat"` on
 * scatter axes) takes over instead.
 *
 * Non-string inputs (e.g. `null`, numbers, arrays) likewise collapse to
 * `undefined` so a stray runtime value leaking through the type guard
 * cannot poison the output.
 */
function normalizeAxisCrossBetween(
  value: ChartAxisCrossBetween | undefined,
): ChartAxisCrossBetween | undefined {
  if (typeof value !== "string") return undefined;
  return VALID_CROSS_BETWEEN.has(value as ChartAxisCrossBetween)
    ? (value as ChartAxisCrossBetween)
    : undefined;
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

/** Recognized values of `<c:majorTickMark>` / `<c:minorTickMark>`. */
const TICK_MARK_VALUES: ReadonlySet<ChartAxisTickMark> = new Set(["none", "in", "out", "cross"]);

/**
 * Normalize a tick-mark value to a token Excel accepts. Unknown / typo'd
 * inputs collapse to `undefined` so the writer never emits a value the
 * OOXML `ST_TickMark` enum rejects.
 */
function normalizeTickMark(value: ChartAxisTickMark | undefined): ChartAxisTickMark | undefined {
  if (value === undefined) return undefined;
  return TICK_MARK_VALUES.has(value) ? value : undefined;
}

/** Recognized values of `<c:tickLblPos>`. */
const TICK_LBL_POS_VALUES: ReadonlySet<ChartAxisTickLabelPosition> = new Set([
  "nextTo",
  "low",
  "high",
  "none",
]);

/**
 * Normalize a tick-label-position value to a token Excel accepts.
 * Unknown / typo'd inputs collapse to `undefined` so the writer never
 * emits a value the OOXML `ST_TickLblPos` enum rejects.
 */
function normalizeTickLblPos(
  value: ChartAxisTickLabelPosition | undefined,
): ChartAxisTickLabelPosition | undefined {
  if (value === undefined) return undefined;
  return TICK_LBL_POS_VALUES.has(value) ? value : undefined;
}

/**
 * Build the `<c:majorTickMark>` / `<c:minorTickMark>` / `<c:tickLblPos>`
 * children for an axis. The OOXML schema (CT_CatAx / CT_ValAx /
 * CT_DateAx / CT_SerAx) places the three elements together right after
 * `<c:numFmt>` and before `<c:crossAx>`. Excel's strict validator
 * rejects any other ordering — keep the tuple together.
 *
 * Each value is omitted when the caller did not pin it; the OOXML
 * defaults (`majorTickMark="out"`, `minorTickMark="none"`,
 * `tickLblPos="nextTo"`) match Excel's reference serialization, so
 * absence and the default round-trip identically through the reader.
 */
function buildAxisTickRendering(
  majorTickMark: ChartAxisTickMark | undefined,
  minorTickMark: ChartAxisTickMark | undefined,
  tickLblPos: ChartAxisTickLabelPosition | undefined,
): string[] {
  const out: string[] = [];
  if (majorTickMark !== undefined) {
    out.push(xmlSelfClose("c:majorTickMark", { val: majorTickMark }));
  }
  if (minorTickMark !== undefined) {
    out.push(xmlSelfClose("c:minorTickMark", { val: minorTickMark }));
  }
  if (tickLblPos !== undefined) {
    out.push(xmlSelfClose("c:tickLblPos", { val: tickLblPos }));
  }
  return out;
}

/**
 * Build the `<c:tickLblSkip>` / `<c:tickMarkSkip>` siblings that sit
 * between `<c:lblOffset>` and `<c:noMultiLvlLbl>` inside `<c:catAx>`
 * (CT_CatAx). Order is `tickLblSkip` first, then `tickMarkSkip` per
 * the OOXML schema. Each element is emitted only when the caller
 * pinned a non-default value (the helper relies on
 * {@link normalizeAxisSkip} having already collapsed `1` and out-of-
 * range inputs to `undefined`).
 */
function buildAxisSkips(
  tickLblSkip: number | undefined,
  tickMarkSkip: number | undefined,
): string[] {
  const out: string[] = [];
  if (tickLblSkip !== undefined) {
    out.push(xmlSelfClose("c:tickLblSkip", { val: tickLblSkip }));
  }
  if (tickMarkSkip !== undefined) {
    out.push(xmlSelfClose("c:tickMarkSkip", { val: tickMarkSkip }));
  }
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
  // → title → numFmt → majorTickMark → minorTickMark → tickLblPos →
  // crossAx → crosses → ... → majorUnit → minorUnit. Each block below
  // mirrors that order.
  // The category axis on bar/column rarely uses scaling, but Excel
  // tolerates the augmentation either way; surface it whenever the
  // caller pinned a value so write-side templates round-trip.
  const catAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_CAT }),
    buildAxisScaling(opts.xScale, opts.xReverse),
    xmlSelfClose("c:delete", { val: opts.xHidden ? 1 : 0 }),
    xmlSelfClose("c:axPos", { val: catPos }),
    ...buildAxisGridlines(opts.xGridlines),
  ];
  if (opts.xAxisTitle) catAxChildren.push(buildAxisTitle(opts.xAxisTitle));
  catAxChildren.push(
    ...buildAxisNumFmt(opts.xNumFmt),
    ...buildAxisTickRendering(opts.xMajorTickMark, opts.xMinorTickMark, opts.xTickLblPos),
  );
  // `<c:txPr>` sits between `<c:tickLblPos>` (the last child of
  // `buildAxisTickRendering`) and `<c:crossAx>` per CT_CatAx (ECMA-376
  // Part 1, §21.2.2.7). Skip the entire block when the caller did not
  // pin a rotation so a fresh chart matches Excel's minimal serialization.
  const xCatAxTxPr = buildAxisTxPr(opts.xLabelRotation);
  if (xCatAxTxPr) catAxChildren.push(xCatAxTxPr);
  catAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL }),
    buildAxisCrosses(opts.xCrosses),
    // `<c:auto>` is always emitted because Excel's reference
    // serialization includes it on every category axis. The writer
    // pins the caller's override when `false`; absence (and any non-
    // boolean) collapses to the OOXML default `true` so untouched
    // charts match Excel's output byte-for-byte.
    xmlSelfClose("c:auto", { val: opts.xAuto ? 1 : 0 }),
    // `<c:lblAlgn>` is always emitted because Excel's reference
    // serialization includes it on every category axis. The writer
    // pins the caller's override when set; absence (or the OOXML
    // default `"ctr"` collapsed by `normalizeAxisLblAlgn`) emits the
    // default so untouched charts match Excel's output byte-for-byte.
    xmlSelfClose("c:lblAlgn", { val: opts.xLblAlgn ?? "ctr" }),
    // `<c:lblOffset>` is always emitted because Excel's reference
    // serialization includes it on every category axis. The writer
    // pins the caller's override when set; absence (or the OOXML
    // default `100` collapsed by `normalizeAxisLblOffset`) emits the
    // default so untouched charts match Excel's output byte-for-byte.
    xmlSelfClose("c:lblOffset", { val: opts.xLblOffset ?? 100 }),
    // OOXML CT_CatAx places `<c:tickLblSkip>` / `<c:tickMarkSkip>`
    // after `<c:lblOffset>` and before `<c:noMultiLvlLbl>`. Only
    // emit each element when the caller pinned a non-default value
    // so a fresh chart matches Excel's reference serialization (the
    // default `1` is omitted and Excel renders every tick).
    ...buildAxisSkips(opts.xTickLblSkip, opts.xTickMarkSkip),
    // `<c:noMultiLvlLbl>` is always emitted because Excel's reference
    // serialization includes it on every category axis. The writer
    // pins the caller's override when `true`; absence and an explicit
    // `false` both produce `val="0"` so untouched charts match Excel's
    // output byte-for-byte.
    xmlSelfClose("c:noMultiLvlLbl", { val: opts.xNoMultiLvlLbl ? 1 : 0 }),
  );

  const valAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL }),
    buildAxisScaling(opts.yScale, opts.yReverse),
    xmlSelfClose("c:delete", { val: opts.yHidden ? 1 : 0 }),
    xmlSelfClose("c:axPos", { val: valPos }),
    ...buildAxisGridlines(opts.yGridlines),
  ];
  if (opts.yAxisTitle) valAxChildren.push(buildAxisTitle(opts.yAxisTitle));
  valAxChildren.push(
    ...buildAxisNumFmt(opts.yNumFmt),
    ...buildAxisTickRendering(opts.yMajorTickMark, opts.yMinorTickMark, opts.yTickLblPos),
  );
  // `<c:txPr>` sits between `<c:tickLblPos>` and `<c:crossAx>` per
  // CT_ValAx (ECMA-376 Part 1, §21.2.2.32). Same omit-by-default
  // contract as the catAx slot above — emit nothing when the caller
  // did not pin a rotation so the writer matches Excel's reference
  // serialization on a fresh value axis.
  const yValAxTxPr = buildAxisTxPr(opts.yLabelRotation);
  if (yValAxTxPr) valAxChildren.push(yValAxTxPr);
  valAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_CAT }),
    buildAxisCrosses(opts.yCrosses),
    // `<c:crossBetween>` sits between `<c:crosses>`/`<c:crossesAt>` and
    // `<c:majorUnit>` per CT_ValAx (ECMA-376 §21.2.2.32). The default
    // for bar / column / line / area is `"between"` — the writer
    // honours an override when the caller pinned `"midCat"` and falls
    // back to the family default otherwise.
    xmlSelfClose("c:crossBetween", { val: opts.yCrossBetween ?? "between" }),
    ...buildAxisTickUnits(opts.yScale),
    // `<c:dispUnits>` is the last child slot on `<c:valAx>` per
    // CT_ValAx (after `<c:minorUnit>`). Bar / column / line / area
    // charts route the X axis through `<c:catAx>` (which rejects the
    // element), so only the Y axis picks up the writer-side input.
    ...buildAxisDispUnits(opts.yDispUnits),
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
function buildUpDownBars(gapWidth: number | undefined): string {
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
function clampUpDownBarsGapWidth(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  if (rounded < 0 || rounded > 500) return undefined;
  return rounded;
}

// ── Area ─────────────────────────────────────────────────────────────

function buildAreaChart(chart: SheetChart, sheetName: string): string {
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

// ── Pie ──────────────────────────────────────────────────────────────

function buildPieChart(chart: SheetChart, sheetName: string): string {
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

// ── Doughnut ─────────────────────────────────────────────────────────

const DOUGHNUT_HOLE_DEFAULT = 50;
const DOUGHNUT_HOLE_MIN = 10;
const DOUGHNUT_HOLE_MAX = 90;

function buildDoughnutChart(chart: SheetChart, sheetName: string): string {
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

const EXPLOSION_MAX = 400;

/**
 * Normalize {@link ChartSeries.explosion} for emission inside
 * `<c:explosion val=".."/>` on a pie / doughnut series.
 *
 * The OOXML schema (`CT_UnsignedInt`) accepts any non-negative integer,
 * but Excel's UI only exposes 0..400% — values outside that band render
 * but trigger Excel's repair dialog. Clamp to the UI band on the way
 * out so a round-trip stays inside the range Excel will render.
 *
 * Returns `undefined` for the default `0` (and any negative / non-finite
 * input) so the writer can elide the element entirely; absence and `0`
 * round-trip identically through the parser's collapse logic.
 */
function clampExplosion(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  if (rounded <= 0) return undefined;
  if (rounded > EXPLOSION_MAX) return EXPLOSION_MAX;
  return rounded;
}

// ── Scatter ──────────────────────────────────────────────────────────

function buildScatterChart(chart: SheetChart, sheetName: string): string {
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

function buildScatterAxes(opts: AxisRenderOptions): string[] {
  const xAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL_X }),
    buildAxisScaling(opts.xScale, opts.xReverse),
    xmlSelfClose("c:delete", { val: opts.xHidden ? 1 : 0 }),
    xmlSelfClose("c:axPos", { val: "b" }),
    ...buildAxisGridlines(opts.xGridlines),
  ];
  if (opts.xAxisTitle) xAxChildren.push(buildAxisTitle(opts.xAxisTitle));
  xAxChildren.push(
    ...buildAxisNumFmt(opts.xNumFmt),
    ...buildAxisTickRendering(opts.xMajorTickMark, opts.xMinorTickMark, opts.xTickLblPos),
  );
  // `<c:txPr>` slot — same CT_ValAx position as the bar / column
  // builder above. Scatter X is a value axis, so the rotation pins on
  // the X-axis just as it does on the Y-axis.
  const xValAxTxPr = buildAxisTxPr(opts.xLabelRotation);
  if (xValAxTxPr) xAxChildren.push(xValAxTxPr);
  xAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL_Y }),
    buildAxisCrosses(opts.xCrosses),
    // Scatter charts default to `"midCat"` (data points sit ON the
    // perpendicular-axis ticks rather than between them). The writer
    // honours an override when the caller pinned `"between"` and falls
    // back to the family default otherwise.
    xmlSelfClose("c:crossBetween", { val: opts.xCrossBetween ?? "midCat" }),
    ...buildAxisTickUnits(opts.xScale),
    // `<c:dispUnits>` slots onto `<c:valAx>` per CT_ValAx (after
    // `<c:minorUnit>`). Scatter charts route both axes through
    // `<c:valAx>`, so the X-axis builder picks up `xDispUnits` here.
    ...buildAxisDispUnits(opts.xDispUnits),
  );

  const yAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL_Y }),
    buildAxisScaling(opts.yScale, opts.yReverse),
    xmlSelfClose("c:delete", { val: opts.yHidden ? 1 : 0 }),
    xmlSelfClose("c:axPos", { val: "l" }),
    ...buildAxisGridlines(opts.yGridlines),
  ];
  if (opts.yAxisTitle) yAxChildren.push(buildAxisTitle(opts.yAxisTitle));
  yAxChildren.push(
    ...buildAxisNumFmt(opts.yNumFmt),
    ...buildAxisTickRendering(opts.yMajorTickMark, opts.yMinorTickMark, opts.yTickLblPos),
  );
  // `<c:txPr>` slot for the scatter Y axis — same CT_ValAx position
  // and omit-by-default contract as the catAx / valAx builders above.
  const yScatterTxPr = buildAxisTxPr(opts.yLabelRotation);
  if (yScatterTxPr) yAxChildren.push(yScatterTxPr);
  yAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL_X }),
    buildAxisCrosses(opts.yCrosses),
    // Scatter Y axis defaults to `"midCat"`. Same override grammar as
    // the X axis above.
    xmlSelfClose("c:crossBetween", { val: opts.yCrossBetween ?? "midCat" }),
    ...buildAxisTickUnits(opts.yScale),
    // `<c:dispUnits>` on the Y axis. Scatter Y is also a value axis,
    // so the same builder applies. See `buildBarAxes` for the broader
    // scope notes.
    ...buildAxisDispUnits(opts.yDispUnits),
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
   * Owning chart's family. Used to scope-guard schema-restricted
   * data-label flags such as `<c:showLeaderLines>` (pie / doughnut
   * only) so a templated chart's pin never leaks onto a chart family
   * Excel's strict validator would reject. Required so the per-family
   * caller cannot forget to thread the type through; every chart
   * builder passes `chart.type` directly.
   */
  chartType: WriteChartKind;
  /**
   * Chart-level data label defaults from {@link SheetChart.dataLabels}.
   * Used when the series itself does not specify `dataLabels`. Series
   * passing `dataLabels: false` always wins over this default.
   */
  dataLabels?: ChartDataLabels;
  /**
   * Per-series line stroke (dash pattern + width). Only meaningful for
   * line / scatter series — every other family ignores the field. The
   * OOXML schema places stroke styling inside `<c:spPr><a:ln>` which is
   * shared with the series fill color, so the writer threads the
   * stroke into the same `<c:spPr>` block whether or not a fill color
   * is set.
   */
  stroke?: ChartLineStroke;
  /**
   * Per-series marker styling. Only meaningful for line / scatter
   * series — every other family ignores the field. The OOXML schema
   * places `<c:marker>` between `<c:spPr>` and `<c:dLbls>` on
   * `CT_LineSer` / `CT_ScatterSer`, so the writer slots it there
   * regardless of which fields are populated.
   */
  marker?: ChartMarker;
  /**
   * Per-series invert-if-negative flag. Only meaningful for bar /
   * column series — every other family ignores the field. The OOXML
   * schema places `<c:invertIfNegative>` between `<c:spPr>` and
   * `<c:dLbls>` on `CT_BarSer` / `CT_Bar3DSer`, so the writer slots
   * it there. The element is only emitted when the field resolves to
   * `true` — `false` is the OOXML default and absence round-trips
   * identically.
   */
  invertIfNegative?: boolean;
  /**
   * Per-series slice explosion (percentage of the radius). Only
   * meaningful for pie / doughnut series — every other family ignores
   * the field. The OOXML schema places `<c:explosion>` between
   * `<c:spPr>` and `<c:dPt>` / `<c:dLbls>` on `CT_PieSer`. The element
   * is only emitted when the resolved value is `> 0` — `0` is the OOXML
   * default and absence round-trips identically.
   */
  explosion?: number;
}

function buildSeries(
  series: ChartSeries,
  index: number,
  sheetName: string,
  numericCategories: boolean,
  options: SeriesOptions,
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

  // Optional fill color and / or line stroke (line / scatter only emit
  // `<a:ln>` width / dash from `options.stroke`; non-line callers leave
  // `options.stroke` undefined so the field is silently dropped on
  // every other chart family).
  const spPr = buildSeriesSpPr(series.color, options?.stroke);
  if (spPr) children.push(spPr);

  // Marker — only line/scatter series honor `<c:marker>` per the OOXML
  // schema (CT_LineSer / CT_ScatterSer). The element sits between
  // `<c:spPr>` and `<c:dLbls>`; non-line/non-scatter callers leave
  // `options.marker` undefined so the field is silently dropped on
  // every other chart family.
  const markerXml = buildSeriesMarker(options?.marker);
  if (markerXml) children.push(markerXml);

  // `<c:invertIfNegative>` — only bar / column (CT_BarSer /
  // CT_Bar3DSer) series carry the element per the OOXML schema. It
  // sits between `<c:spPr>` (and the bar-irrelevant `<c:marker>`
  // slot, which is never populated for bar/column callers anyway)
  // and `<c:dLbls>`. Non-bar callers leave `options.invertIfNegative`
  // undefined so the field is silently dropped on every other chart
  // family. Emit only when the resolved value is `true` — `false`
  // matches the OOXML default and absence round-trips identically.
  if (options?.invertIfNegative === true) {
    children.push(xmlSelfClose("c:invertIfNegative", { val: 1 }));
  }

  // `<c:explosion>` — only pie / doughnut (CT_PieSer, shared across
  // the pie family via `EG_PieSer`) series carry the element per the
  // OOXML schema. It sits between `<c:spPr>` and `<c:dPt>` / `<c:dLbls>`.
  // Non-pie callers leave `options.explosion` undefined so the field
  // is silently dropped on every other chart family. Emit only when
  // the resolved value is non-zero — `0` matches the OOXML default and
  // absence round-trips identically.
  const explosion = clampExplosion(options?.explosion);
  if (explosion !== undefined) {
    children.push(xmlSelfClose("c:explosion", { val: explosion }));
  }

  // Data labels — series-level override always wins over the chart-level
  // default. `<c:dLbls>` sits between <c:spPr> and <c:cat>/<c:val> per
  // the OOXML series schema (CT_BarSer, CT_LineSer, ...). The chart
  // type threads through so the dLbls body can scope-guard the pie /
  // doughnut-only `<c:showLeaderLines>` flag.
  const seriesDLblsXml = buildSeriesDataLabels(
    series.dataLabels,
    options.dataLabels,
    options.chartType,
  );
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

// ── Stroke ───────────────────────────────────────────────────────────

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
 * Validate a dash style against `ST_PresetLineDashVal`. Returns
 * `undefined` for unrecognized values so the writer can elide
 * `<a:prstDash>` rather than emit a token Excel will reject.
 */
function normalizeDashStyle(value: ChartLineDashStyle | undefined): ChartLineDashStyle | undefined {
  if (value === undefined) return undefined;
  return VALID_DASH_STYLES.has(value) ? value : undefined;
}

/**
 * Convert a stroke width in points to the integer EMU value the OOXML
 * `w` attribute requires. Excel's UI exposes 0.25..13.5 pt — values
 * outside that band are clamped to keep round-trips inside the range
 * Excel will render. Non-finite values collapse to `undefined` so the
 * writer can omit the attribute entirely. The point value is also
 * snapped to the nearest quarter-point so a parsed-then-written stroke
 * does not drift across round-trips (Excel rounds in the UI anyway).
 */
function clampStrokeWidthPt(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  // Snap to the 0.25 pt grid Excel's UI exposes (Math.round(x * 4) / 4).
  const snapped = Math.round(value * 4) / 4;
  if (snapped < STROKE_WIDTH_MIN_PT) return STROKE_WIDTH_MIN_PT;
  if (snapped > STROKE_WIDTH_MAX_PT) return STROKE_WIDTH_MAX_PT;
  return snapped;
}

/**
 * Build the `<c:spPr>` element shared by series fill color and series
 * line stroke. Returns `undefined` when neither field carries any
 * meaningful settings — an empty `<c:spPr/>` collapses to the
 * inherited series-rotation default Excel picks anyway, so omitting it
 * keeps untouched chart XML byte-clean.
 *
 * The OOXML `<a:ln>` element accepts both a `w` attribute (stroke
 * width in EMU) and child elements `<a:solidFill>` / `<a:prstDash>` in
 * a fixed order. When a fill color is set, the stroke also renders the
 * same color (matching Excel's "Format Data Series → Fill" default
 * which paints the line in the fill color). Stroke metadata (dash and
 * width) layers on top without overriding the line color so a `color +
 * stroke` combo behaves like Excel's UI: the line picks up the fill
 * color and the dash / width override visibility-only attributes.
 */
function buildSeriesSpPr(
  rgbHex: string | undefined,
  stroke: ChartLineStroke | undefined,
): string | undefined {
  const fillHex = rgbHex ? rgbHex.replace(/^#/, "").toUpperCase() : undefined;
  const dash = normalizeDashStyle(stroke?.dash);
  const widthPt = clampStrokeWidthPt(stroke?.width);

  if (!fillHex && dash === undefined && widthPt === undefined) {
    return undefined;
  }

  const spPrChildren: string[] = [];
  if (fillHex) {
    spPrChildren.push(
      xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: fillHex })]),
    );
  }

  // `<a:ln>` carries stroke metadata. Emit it whenever a fill color is
  // set (so the connecting line picks up the same color, matching the
  // legacy behavior) or whenever stroke width / dash is configured.
  if (fillHex || dash !== undefined || widthPt !== undefined) {
    const lnAttrs: Record<string, string | number> = {};
    if (widthPt !== undefined) {
      // OOXML stores stroke width in EMU (1 pt = 12 700 EMU). Round to
      // the nearest integer because the schema types `w` as `xsd:int`.
      lnAttrs.w = Math.round(widthPt * EMU_PER_PT);
    }
    const lnChildren: string[] = [];
    if (fillHex) {
      lnChildren.push(
        xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: fillHex })]),
      );
    }
    if (dash !== undefined) {
      lnChildren.push(xmlSelfClose("a:prstDash", { val: dash }));
    }
    spPrChildren.push(
      lnChildren.length === 0
        ? xmlSelfClose("a:ln", lnAttrs)
        : xmlElement("a:ln", Object.keys(lnAttrs).length > 0 ? lnAttrs : undefined, lnChildren),
    );
  }

  return xmlElement("c:spPr", undefined, spPrChildren);
}

// ── Marker ───────────────────────────────────────────────────────────

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

const MARKER_SIZE_MIN = 2;
const MARKER_SIZE_MAX = 72;

/**
 * Normalize a marker size to the OOXML 2..72 band (`ST_MarkerSize`).
 * Excel's UI clamps anything outside this range; we mirror that on the
 * write side so an out-of-range hint never reaches the chart XML.
 *
 * Returns `undefined` for non-finite values so the writer can elide
 * `<c:size>` (Excel falls back to its series-rotation default).
 */
function clampMarkerSize(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  if (rounded < MARKER_SIZE_MIN) return MARKER_SIZE_MIN;
  if (rounded > MARKER_SIZE_MAX) return MARKER_SIZE_MAX;
  return rounded;
}

/**
 * Validate a marker symbol against the OOXML `ST_MarkerStyle` enum.
 * Returns `undefined` for unrecognized values so the writer can elide
 * `<c:symbol>` rather than emit a token Excel will reject.
 */
function normalizeMarkerSymbol(
  value: ChartMarkerSymbol | undefined,
): ChartMarkerSymbol | undefined {
  if (value === undefined) return undefined;
  return VALID_MARKER_SYMBOLS.has(value) ? value : undefined;
}

/**
 * Build a `<c:marker>` element for a series. Returns `undefined` when
 * the marker block carries no meaningful settings — an empty marker
 * element collapses to the inherited series-rotation default Excel
 * picks anyway, so omitting it keeps untouched XML byte-clean.
 */
function buildSeriesMarker(marker: ChartMarker | undefined): string | undefined {
  if (!marker) return undefined;
  const symbol = normalizeMarkerSymbol(marker.symbol);
  const size = clampMarkerSize(marker.size);
  const fill = normalizeRgbHex(marker.fill);
  const line = normalizeRgbHex(marker.line);

  if (symbol === undefined && size === undefined && !fill && !line) return undefined;

  const children: string[] = [];
  if (symbol !== undefined) children.push(xmlSelfClose("c:symbol", { val: symbol }));
  if (size !== undefined) children.push(xmlSelfClose("c:size", { val: size }));

  if (fill || line) {
    const spPrChildren: string[] = [];
    if (fill) {
      spPrChildren.push(
        xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: fill })]),
      );
    }
    if (line) {
      spPrChildren.push(
        xmlElement("a:ln", undefined, [
          xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: line })]),
        ]),
      );
    }
    children.push(xmlElement("c:spPr", undefined, spPrChildren));
  }

  return xmlElement("c:marker", undefined, children);
}

/**
 * Validate a 6-digit RGB hex string and return the uppercase form.
 * Strips a leading `#`. Returns `undefined` for missing or malformed
 * values so the writer can elide colored sub-elements that would
 * otherwise carry a token Excel rejects.
 */
function normalizeRgbHex(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined;
  const stripped = value.replace(/^#/, "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(stripped) ? stripped : undefined;
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
  chartType: WriteChartKind,
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
    return buildDataLabelsBody(seriesDLbls, chartType);
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
  return buildDataLabelsBody(chart.dataLabels, chart.type);
}

/**
 * Render the OOXML `<c:dLbls>` body. Element order follows CT_DLbls:
 * delete? before numFmt? before spPr? before txPr? before dLblPos? before
 * showLegendKey, showVal, showCatName, showSerName, showPercent,
 * showBubbleSize, separator?, showLeaderLines? — toggles must appear
 * in that exact order or Excel ignores the block.
 *
 * `chartType` lets the builder gate the pie / doughnut-only
 * `<c:showLeaderLines>` element — the OOXML schema scopes that flag to
 * `EG_DLbls` for `CT_PieChart` / `CT_DoughnutChart` only, so the writer
 * silently drops `dl.showLeaderLines` on every other family rather than
 * emit a child Excel's strict validator would reject.
 */
function buildDataLabelsBody(dl: ChartDataLabels, chartType: WriteChartKind): string {
  const children: string[] = [];

  // `<c:numFmt>` sits at the head of the CT_DLbls sequence (before
  // `<c:spPr>` / `<c:txPr>` / `<c:dLblPos>` / the show* toggles). The
  // writer skips emission entirely when the caller leaves `numberFormat`
  // unset so a fresh chart matches Excel's reference shape (no number
  // override means Excel inherits from the source cells).
  const numFmt = resolveDataLabelsNumberFormat(dl.numberFormat);
  if (numFmt) {
    children.push(
      xmlSelfClose("c:numFmt", {
        formatCode: numFmt.formatCode,
        sourceLinked: numFmt.sourceLinked === true ? 1 : 0,
      }),
    );
  }

  if (dl.position) {
    children.push(xmlSelfClose("c:dLblPos", { val: dl.position }));
  }

  // OOXML requires showLegendKey to appear first when any toggle is set.
  // Always emit it explicitly so the rendered XML is deterministic.
  // Non-boolean inputs collapse to `false` to keep the on-the-wire output
  // stable, mirroring how the other `show*` toggles treat their inputs.
  children.push(xmlSelfClose("c:showLegendKey", { val: dl.showLegendKey === true ? 1 : 0 }));
  children.push(xmlSelfClose("c:showVal", { val: dl.showValue ? 1 : 0 }));
  children.push(xmlSelfClose("c:showCatName", { val: dl.showCategoryName ? 1 : 0 }));
  children.push(xmlSelfClose("c:showSerName", { val: dl.showSeriesName ? 1 : 0 }));
  children.push(xmlSelfClose("c:showPercent", { val: dl.showPercent ? 1 : 0 }));
  children.push(xmlSelfClose("c:showBubbleSize", { val: 0 }));

  if (dl.separator !== undefined) {
    children.push(xmlElement("c:separator", undefined, xmlEscape(dl.separator)));
  }

  // `<c:showLeaderLines>` sits at the tail of the `EG_DLbls` group
  // (after `<c:separator>`, before `<c:extLst>`). The OOXML schema
  // scopes the element to pie / doughnut chart families exclusively
  // (`EG_DLbls` for `CT_PieChart` / `CT_DoughnutChart` only — bar /
  // column / line / area / scatter route through `EG_DLblsShared` which
  // omits it). The writer drops the field silently on every non-pie /
  // non-doughnut family to mirror Excel's reference serialization.
  //
  // The OOXML default is `true` (Excel paints leader lines on every
  // label that gets pushed outside its slice). Only an explicit
  // `false` flips the toggle, so absence and the default round-trip
  // identically through {@link parseChart}. Non-boolean inputs collapse
  // to the default, mirroring how the other `show*` toggles treat
  // their inputs.
  if ((chartType === "pie" || chartType === "doughnut") && dl.showLeaderLines === false) {
    children.push(xmlSelfClose("c:showLeaderLines", { val: 0 }));
  }

  return xmlElement("c:dLbls", undefined, children);
}

/**
 * Resolve the `<c:numFmt>` value emitted inside `<c:dLbls>`.
 *
 * Returns `undefined` when the caller leaves `numberFormat` unset or
 * when `formatCode` is missing / non-string / empty — the OOXML schema
 * requires `formatCode` on every emitted `<c:numFmt>` so a malformed
 * record is dropped rather than fabricate a placeholder Excel rejects.
 * Mirrors how the axis-side number-format pipeline shapes its output
 * so a parsed value can flow straight from the read side back into the
 * write side without transformation.
 *
 * `sourceLinked` is normalized to a literal boolean — only `true`
 * survives, every other shape (`undefined` / `false` / non-boolean)
 * collapses so the writer's `val` attribute defaults to `0`.
 */
function resolveDataLabelsNumberFormat(
  value: ChartAxisNumberFormat | undefined,
): ChartAxisNumberFormat | undefined {
  if (!value) return undefined;
  const formatCode = value.formatCode;
  if (typeof formatCode !== "string" || formatCode.length === 0) return undefined;
  const out: ChartAxisNumberFormat = { formatCode };
  if (value.sourceLinked === true) out.sourceLinked = true;
  return out;
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

function buildLegend(
  pos: LegendPos,
  overlay: boolean,
  entries: readonly ResolvedLegendEntry[],
): string {
  const children: string[] = [xmlSelfClose("c:legendPos", { val: pos })];

  // CT_Legend sequence places `<c:legendEntry>` after `<c:legendPos>`
  // and before `<c:layout>` / `<c:overlay>` (ECMA-376 Part 1,
  // §21.2.2.114). The writer never emits `<c:layout>` here, so the
  // entries collapse straight onto `<c:overlay>`. Each entry is emitted
  // with both `<c:idx>` and `<c:delete>` so a re-parse sees the
  // canonical shape — Excel itself emits `<c:delete val="1"/>` whenever
  // the action is "Hide legend entry", and the writer mirrors that even
  // for the OOXML default `false` value (an explicit `<c:delete val="0"/>`
  // round-trips cleanly through `parseChart`).
  for (const entry of entries) {
    children.push(
      xmlElement("c:legendEntry", undefined, [
        xmlSelfClose("c:idx", { val: entry.idx }),
        xmlSelfClose("c:delete", { val: entry.delete ? 1 : 0 }),
      ]),
    );
  }

  children.push(xmlSelfClose("c:overlay", { val: overlay ? 1 : 0 }));
  return xmlElement("c:legend", undefined, children);
}

interface ResolvedLegendEntry {
  idx: number;
  delete: boolean;
}

/**
 * Normalize {@link SheetChart.legendEntries} into an emit-ready list.
 *
 * The OOXML schema (`CT_LegendEntry`) places `<c:idx val="N"/>` as the
 * required selector and `<c:delete val=".."/>` as the hide flag. Hucre
 * accepts a free-form `ChartLegendEntry[]` from callers; this helper
 * strips entries whose `idx` cannot land on a real series and
 * deduplicates duplicate `idx` values so the writer never emits the
 * same selector twice (the last entry wins so a clone-through that
 * appends an override naturally beats the source's parsed value).
 *
 * Validation rules:
 *   - `idx` must be a non-negative integer (matches `xsd:unsignedInt`
 *     on `<c:idx val=".."/>`); non-finite, negative, or non-integer
 *     values drop entirely rather than emit a token Excel rejects.
 *   - `delete` collapses to a strict boolean — anything other than
 *     literal `true` resolves to `false`. Mirrors how `legendOverlay`
 *     / `roundedCorners` / `plotVisOnly` treat their inputs.
 *
 * Entries are emitted in ascending `idx` order so the rendered chart
 * matches Excel's reference serialization (Excel sorts by `<c:idx>` on
 * write even when the in-memory list arrived unsorted). Returns an
 * empty array when the chart has no entries to emit so the caller can
 * avoid touching the legend block.
 */
function resolveLegendEntries(chart: SheetChart): ResolvedLegendEntry[] {
  const raw = chart.legendEntries;
  if (!Array.isArray(raw) || raw.length === 0) return [];

  const byIdx = new Map<number, ResolvedLegendEntry>();
  for (const entry of raw) {
    if (!entry || typeof entry !== "object") continue;
    const idx = entry.idx;
    if (typeof idx !== "number" || !Number.isFinite(idx)) continue;
    if (!Number.isInteger(idx) || idx < 0) continue;
    byIdx.set(idx, { idx, delete: entry.delete === true });
  }

  return Array.from(byIdx.values()).sort((a, b) => a.idx - b.idx);
}

/**
 * Resolve `<c:legend><c:overlay val=".."/></c:legend>` from
 * {@link SheetChart.legendOverlay}.
 *
 * Defaults to `false` (the OOXML default Excel itself emits — the
 * legend reserves its own slot and the plot area shrinks to make room).
 * Anything other than literal `true` collapses to `false` so a stray
 * non-boolean leaking through the type guard (e.g. `0` / `1` / `"true"`
 * / `null`) never produces `<c:overlay val="1"/>`. This matches how
 * `roundedCorners` / `plotVisOnly` / axis `hidden` treat their inputs:
 * a literal boolean is the only path to a non-default value.
 *
 * The writer always emits `<c:overlay>` because Excel's reference
 * serialization includes the element on every visible legend; only the
 * `val` flips when the caller pins `legendOverlay: true`.
 */
function resolveLegendOverlay(chart: SheetChart): boolean {
  return chart.legendOverlay === true;
}

// ── Display Blanks As ────────────────────────────────────────────────

const DISP_BLANKS_AS_VALUES: ReadonlySet<ChartDisplayBlanksAs> = new Set(["gap", "zero", "span"]);

/**
 * Resolve the `<c:dispBlanksAs>` value emitted on `<c:chart>`.
 *
 * Defaults to `"gap"` (the OOXML default) when the chart does not set
 * the field. Unknown / unsupported tokens collapse to `"gap"` rather
 * than emit an attribute Excel ignores. The writer always emits the
 * element so the file's intent is explicit even on roundtrip — Excel
 * itself includes it in every reference serialization.
 */
function resolveDispBlanksAs(chart: SheetChart): ChartDisplayBlanksAs {
  const raw = chart.dispBlanksAs;
  if (raw && DISP_BLANKS_AS_VALUES.has(raw)) return raw;
  return "gap";
}

// ── Plot Visible Only ────────────────────────────────────────────────

/**
 * Resolve the `<c:plotVisOnly>` value emitted on `<c:chart>`.
 *
 * Defaults to `true` (the OOXML schema default — hidden rows/columns
 * drop out of the chart). An explicit `chart.plotVisOnly === false`
 * flips the toggle to mirror Excel's "Show data in hidden rows and
 * columns" preference. The writer always emits the element so the
 * file's intent is explicit even on roundtrip — Excel itself includes
 * it in every reference serialization.
 */
function resolvePlotVisOnly(chart: SheetChart): boolean {
  if (typeof chart.plotVisOnly === "boolean") return chart.plotVisOnly;
  return true;
}

// ── Show Data Labels Over Max ────────────────────────────────────────

/**
 * Resolve the `<c:showDLblsOverMax>` value emitted on `<c:chart>`.
 *
 * Defaults to `true` (the OOXML schema default — labels render for
 * every data point regardless of whether the value exceeds the pinned
 * axis maximum). An explicit `chart.showDLblsOverMax === false` flips
 * the toggle to mirror Excel's "Format Axis → Labels → Show data labels
 * for values over maximum scale" checkbox unchecked. The writer always
 * emits the element so the file's intent is explicit even on roundtrip
 * — Excel itself includes it in every reference serialization.
 *
 * `<c:showDLblsOverMax>` sits at the tail of CT_Chart per ECMA-376
 * Part 1, §21.2.2.29 (after `<c:dispBlanksAs>` and before `<c:extLst>`).
 * Mirrors the always-emit contract of {@link resolvePlotVisOnly} and
 * {@link resolveDispBlanksAs}.
 */
function resolveShowDLblsOverMax(chart: SheetChart): boolean {
  if (typeof chart.showDLblsOverMax === "boolean") return chart.showDLblsOverMax;
  return true;
}

// ── Rounded Corners ──────────────────────────────────────────────────

/**
 * Resolve the `<c:roundedCorners>` value emitted on `<c:chartSpace>`.
 *
 * Defaults to `false` (the OOXML schema default — square chart frame).
 * An explicit `chart.roundedCorners === true` flips the toggle to mirror
 * Excel's "Format Chart Area → Border → Rounded corners" preference.
 * The writer always emits the element so the file's intent is explicit
 * even on roundtrip — Excel itself includes it in every reference
 * serialization.
 *
 * `<c:roundedCorners>` is the first child of `<c:chartSpace>` per the
 * `CT_ChartSpace` sequence, sitting before `<c:chart>` rather than
 * inside it (the toggle styles the outer frame, not the plot area).
 */
function resolveRoundedCorners(chart: SheetChart): boolean {
  if (typeof chart.roundedCorners === "boolean") return chart.roundedCorners;
  return false;
}

// ── Chart Style Preset ──────────────────────────────────────────────

/**
 * Resolve the `<c:style val=".."/>` value emitted on `<c:chartSpace>`.
 *
 * Returns `undefined` when the chart leaves `style` unset (the writer
 * skips the element entirely so a fresh chart matches Excel's implicit
 * default rather than pinning the application's `2` preset). Out-of-
 * range and non-integer values also collapse to `undefined` rather
 * than emit a token Excel would reject — `<c:style>` is `xsd:unsigned
 * Byte` in the OOXML schema with the gallery range of 1–48.
 *
 * `<c:style>` sits on `<c:chartSpace>` (a sibling of `<c:chart>`, not
 * a child) per CT_ChartSpace. The element follows `<c:roundedCorners>`
 * and precedes `<c:chart>` in the schema sequence.
 */
function resolveStyle(chart: SheetChart): number | undefined {
  const raw = chart.style;
  if (typeof raw !== "number") return undefined;
  if (!Number.isInteger(raw)) return undefined;
  if (raw < 1 || raw > 48) return undefined;
  return raw;
}

// ── Date System ──────────────────────────────────────────────────────

/**
 * Resolve the `<c:date1904 val=".."/>` value emitted on
 * `<c:chartSpace>`.
 *
 * Returns `true` when the chart pins `date1904: true` (the
 * non-default state), `false` otherwise. The caller decides whether
 * to emit the element — the writer skips it whenever the resolved
 * value is `false` so absence and the OOXML default `val="0"`
 * round-trip identically through {@link parseChart}. Non-boolean
 * values collapse to `false` so a stray runtime value never reaches
 * the rendered XML.
 *
 * `<c:date1904>` mirrors the host workbook's
 * `<workbookPr date1904="1"/>` toggle — `true` interprets date-axis
 * values under the 1904 base (Excel for Mac's legacy epoch where day
 * 0 falls on 1904-01-01) and `false` under the 1900 base. The
 * element governs the whole chart document, not just the plot area.
 *
 * `<c:date1904>` sits at the head of `<c:chartSpace>` per
 * CT_ChartSpace — before `<c:lang>` and `<c:roundedCorners>` — so
 * the writer threads it first when the chart pins it.
 */
function resolveDate1904(chart: SheetChart): boolean {
  return chart.date1904 === true;
}

// ── Editing Locale ──────────────────────────────────────────────────

/**
 * Resolve the `<c:lang val=".."/>` value emitted on `<c:chartSpace>`.
 *
 * Returns `undefined` when the chart leaves `lang` unset (the writer
 * skips the element entirely so a fresh chart falls back to Excel's
 * workbook-level editing language rather than fabricating a token
 * neither the caller nor a re-parse would carry). Malformed and
 * non-string values also collapse to `undefined` — `<c:lang>` is
 * `xsd:language` in the OOXML schema, the IETF BCP-47 culture-name
 * shape `[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*` (e.g. `en-US`, `tr-TR`,
 * `zh-Hant-TW`).
 *
 * `<c:lang>` sits on `<c:chartSpace>` (a sibling of `<c:chart>`, not
 * a child) per CT_ChartSpace. The element follows `<c:date1904>` and
 * precedes `<c:roundedCorners>` in the schema sequence — the locale
 * governs the entire chart document (locale-sensitive separators on
 * unformatted axis ticks, default text font fallback, the locale
 * recorded for in-chart text runs), not just the plot area.
 */
function resolveLang(chart: SheetChart): string | undefined {
  const raw = chart.lang;
  if (typeof raw !== "string") return undefined;
  if (!/^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$/.test(raw)) return undefined;
  return raw;
}

// ── Vary Colors ──────────────────────────────────────────────────────

/**
 * Chart families whose Excel-default `<c:varyColors>` value is `true`
 * (each data point in the lone series renders in a unique color). Pie
 * and doughnut both ship that way out of Excel's chart UI; every other
 * authored family defaults to `false`.
 */
const VARY_COLORS_DEFAULT_TRUE_TYPES: ReadonlySet<WriteChartKind> = new Set(["pie", "doughnut"]);

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
function resolveVaryColors(chart: SheetChart): boolean {
  if (typeof chart.varyColors === "boolean") return chart.varyColors;
  return VARY_COLORS_DEFAULT_TRUE_TYPES.has(chart.type);
}

// ── Scatter Style ────────────────────────────────────────────────────

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
function resolveScatterStyle(chart: SheetChart): ChartScatterStyle {
  const raw = chart.scatterStyle;
  if (raw && SCATTER_STYLE_VALUES.has(raw)) return raw;
  return "lineMarker";
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

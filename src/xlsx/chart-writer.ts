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
  ChartBorderDash,
  ChartDataLabels,
  ChartDisplayBlanksAs,
  ChartLineDashStyle,
  ChartLineStroke,
  ChartManualLayout,
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
import {
  EMU_PER_PT,
  VALID_DASH_STYLES,
  clampStrokeWidthPt,
  normalizeBorderDash,
  normalizeRgbHex as normalizeRgbHexShared,
} from "./chart/shape";
import {
  buildBackWallThickness,
  buildFloorThickness,
  buildSideWallThickness,
  buildView3D,
} from "./chart/walls";
import {
  type ResolvedManualLayout,
  buildManualLayout,
  normalizeManualLayout,
} from "./chart/layout";
import {
  FONT_SIZE_MAX_PT,
  FONT_SIZE_MIN_PT,
  FONT_SZ_PER_POINT,
  ROTATION_MAX_DEG,
  ROTATION_MIN_DEG,
  TXPR_ROT_PER_DEGREE,
} from "./chart/text";
import {
  type LegendPos,
  type ResolvedLegendEntry,
  buildLegend,
  buildLegendSpPr,
  buildLegendTxPr,
  normalizeLegendFontFamily,
  resolveLegendBold,
  resolveLegendBorderColor,
  resolveLegendBorderDash,
  resolveLegendBorderWidth,
  resolveLegendEntries,
  resolveLegendFillColor,
  resolveLegendFontColor,
  resolveLegendFontFamily,
  resolveLegendFontSize,
  resolveLegendItalic,
  resolveLegendLayout,
  resolveLegendOverlay,
  resolveLegendPosition,
  resolveLegendStrikethrough,
  resolveLegendUnderline,
} from "./chart/legend";
import {
  buildTitle,
  buildTitleSpPr,
  normalizeTitleBold,
  normalizeTitleColor,
  normalizeTitleFontFamily,
  normalizeTitleFontSize,
  normalizeTitleItalic,
  normalizeTitleRotation,
  normalizeTitleStrike,
  normalizeTitleUnderline,
  resolveTitleBold,
  resolveTitleBorderColor,
  resolveTitleBorderDash,
  resolveTitleBorderWidth,
  resolveTitleColor,
  resolveTitleFillColor,
  resolveTitleFontFamily,
  resolveTitleFontSize,
  resolveTitleItalic,
  resolveTitleLayout,
  resolveTitleOverlay,
  resolveTitleRotation,
  resolveTitleStrike,
  resolveTitleUnderline,
} from "./chart/title";
import { buildDataTable, resolveDataTable } from "./chart/dataTable";
import {
  buildChartLevelDataLabels,
  buildSeriesDataLabels,
} from "./chart/dataLabels";
import { buildSeries } from "./chart/series";
import {
  AXIS_ID_CAT,
  AXIS_ID_VAL,
  AXIS_ID_VAL_X,
  AXIS_ID_VAL_Y,
  type AxisRenderOptions,
  buildBarAxes,
  buildScatterAxes,
  normalizeAxisCrossBetween,
  normalizeAxisCrosses,
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
  resolveAutoTitleDeleted,
} from "./chart/axis";

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
      buildTitle(
        chart.title,
        resolveTitleOverlay(chart),
        resolveTitleRotation(chart),
        resolveTitleFontSize(chart),
        resolveTitleBold(chart),
        resolveTitleItalic(chart),
        resolveTitleColor(chart),
        resolveTitleStrike(chart),
        resolveTitleUnderline(chart),
        resolveTitleFontFamily(chart),
        resolveTitleLayout(chart),
        resolveTitleFillColor(chart),
        resolveTitleBorderColor(chart),
        resolveTitleBorderWidth(chart),
        resolveTitleBorderDash(chart),
      ),
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

  // `<c:floor>` (CT_Surface, ECMA-376 Part 1, §21.2.2.69) sits on
  // `<c:chart>` between `<c:view3D>` and `<c:sideWall>` /
  // `<c:backWall>` / `<c:plotArea>` per CT_Chart. The writer pins only
  // the `<c:thickness>` child here — `<c:spPr>` / `<c:pictureOptions>`
  // / `<c:extLst>` styling on the floor block is not modelled at this
  // layer. Like `<c:view3D>`, the schema accepts `<c:floor>` on every
  // CT_Chart even though it is only meaningful on 3D families
  // (`bar3D`, `line3D`, `pie3D`, `area3D`, `surface3D`); Excel
  // silently ignores it on 2D families. The writer skips emission
  // entirely when the caller leaves `floorThickness` unset (or pins
  // `0`) so a fresh chart matches Excel's reference serialization
  // byte-for-byte.
  const floorXml = buildFloorThickness(chart.floorThickness);
  if (floorXml !== undefined) {
    chartChildren.push(floorXml);
  }

  // `<c:sideWall>` (CT_Surface, ECMA-376 Part 1, §21.2.2.187) sits on
  // `<c:chart>` between `<c:floor>` and `<c:backWall>` /
  // `<c:plotArea>` per CT_Chart. The writer pins only the
  // `<c:thickness>` child here — `<c:spPr>` / `<c:pictureOptions>` /
  // `<c:extLst>` styling on the side-wall block is not modelled at
  // this layer. Like `<c:view3D>`, the schema accepts `<c:sideWall>`
  // on every CT_Chart even though it is only meaningful on 3D
  // families (`bar3D`, `line3D`, `pie3D`, `area3D`, `surface3D`);
  // Excel silently ignores it on 2D families. The writer skips
  // emission entirely when the caller leaves `sideWallThickness`
  // unset (or pins `0`) so a fresh chart matches Excel's reference
  // serialization byte-for-byte.
  const sideWallXml = buildSideWallThickness(chart.sideWallThickness);
  if (sideWallXml !== undefined) {
    chartChildren.push(sideWallXml);
  }

  // `<c:backWall>` (CT_Surface, ECMA-376 Part 1, §21.2.2.31) sits on
  // `<c:chart>` between `<c:sideWall>` and `<c:plotArea>` per CT_Chart.
  // The writer pins only the `<c:thickness>` child here — `<c:spPr>`
  // / `<c:pictureOptions>` / `<c:extLst>` styling on the back-wall
  // block is not modelled at this layer. Like `<c:floor>`, the schema
  // accepts `<c:backWall>` on every CT_Chart even though it is only
  // meaningful on 3D families (`bar3D`, `line3D`, `pie3D`, `area3D`,
  // `surface3D`); Excel silently ignores it on 2D families. The writer
  // skips emission entirely when the caller leaves `backWallThickness`
  // unset (or pins `0`) so a fresh chart matches Excel's reference
  // serialization byte-for-byte.
  const backWallXml = buildBackWallThickness(chart.backWallThickness);
  if (backWallXml !== undefined) {
    chartChildren.push(backWallXml);
  }

  // ── Plot Area ──
  chartChildren.push(buildPlotArea(chart, sheetName));

  // ── Legend ──
  if (legendPos) {
    chartChildren.push(
      buildLegend(
        legendPos,
        resolveLegendOverlay(chart),
        resolveLegendEntries(chart),
        resolveLegendFontSize(chart),
        resolveLegendBold(chart),
        resolveLegendItalic(chart),
        resolveLegendUnderline(chart),
        resolveLegendStrikethrough(chart),
        resolveLegendFontColor(chart),
        resolveLegendFontFamily(chart),
        resolveLegendLayout(chart),
        resolveLegendFillColor(chart),
        resolveLegendBorderColor(chart),
        resolveLegendBorderWidth(chart),
        resolveLegendBorderDash(chart),
      ),
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

  // `<c:chartSpace><c:spPr><a:solidFill><a:srgbClr val=".."/></a:solidFill>
  // </c:spPr></c:chartSpace>` — Excel's "Format Chart Area -> Fill ->
  // Solid fill -> Color" pin (the same dialog the user reaches by
  // right-clicking the chart's outer frame). The slot sits at the tail
  // of `<c:chartSpace>` per CT_ChartSpace (ECMA-376 Part 1, §21.2.2.29),
  // after `<c:chart>` / `<c:externalData>` / `<c:printSettings>` /
  // `<c:userShapes>` and before the optional `<c:txPr>` / `<c:extLst>`.
  // The writer emits the block only when `chart.chartSpaceFillColor`
  // normalizes to a literal hex; absence and every malformed token
  // collapse to no `<c:spPr>` so a fresh chart matches Excel's
  // reference shape byte-for-byte.
  const chartSpaceSpPrXml = buildChartSpaceSpPr(chart);
  if (chartSpaceSpPrXml !== undefined) {
    chartSpaceChildren.push(chartSpaceSpPrXml);
  }

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


/**
 * OOXML's `<a:bodyPr rot="N"/>` attribute is in 60000ths of a degree —
 * the writer holds `titleRotation` in whole degrees and converts at
 * emit time. Excel's UI exposes the `-90..90` band; out-of-band values
 * clamp to the nearest endpoint so a corrupt template cannot leak
 * through to the writer either.
 *
 * Aliased onto the shared {@link TXPR_ROT_PER_DEGREE} /
 * {@link ROTATION_MIN_DEG} / {@link ROTATION_MAX_DEG} constants in
 * `chart/text` so every typography host (chart-title, axis-title,
 * tick-label, legend, data-label, data-table) shares the same conversion
 * factor.
 */
const TITLE_ROT_PER_DEGREE = TXPR_ROT_PER_DEGREE;
const TITLE_ROTATION_MIN_DEG = ROTATION_MIN_DEG;
const TITLE_ROTATION_MAX_DEG = ROTATION_MAX_DEG;


/**
 * OOXML's `<a:defRPr sz="N"/>` / `<a:rPr sz="N"/>` attribute is in
 * 100ths of a point — the writer holds {@link SheetChart.titleFontSize}
 * in points and converts at emit time. The OOXML `ST_TextFontSize`
 * schema restricts `sz` to the inclusive `100..400000` band; the
 * writer's clamp uses the same range converted to points (`1..400`pt),
 * so any out-of-range value drops at emit time rather than surface a
 * token Excel would reject.
 *
 * Aliased onto the shared {@link FONT_SZ_PER_POINT} /
 * {@link FONT_SIZE_MIN_PT} / {@link FONT_SIZE_MAX_PT} constants in
 * `chart/text` so every typography host shares the same range.
 */
const TITLE_FONT_SZ_PER_POINT = FONT_SZ_PER_POINT;
const TITLE_FONT_SIZE_MIN_PT = FONT_SIZE_MIN_PT;
const TITLE_FONT_SIZE_MAX_PT = FONT_SIZE_MAX_PT;

/**
 * Application-default `sz` value for the chart title's `<a:defRPr>` /
 * `<a:rPr>` slots — Excel renders the title at 14pt (`sz="1400"`)
 * unless the user pins a custom size. Absence of
 * {@link SheetChart.titleFontSize} resolves to this default so a fresh
 * chart matches Excel's reference serialization byte-for-byte.
 */
const TITLE_DEFAULT_FONT_SIZE_SZ = 1400;


// ── Plot Area ────────────────────────────────────────────────────────

function buildPlotArea(chart: SheetChart, sheetName: string): string {
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
function buildPlotAreaSpPr(chart: SheetChart): string | undefined {
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
function normalizePlotAreaFillColor(value: string | undefined): string | undefined {
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
function normalizePlotAreaBorderColor(value: string | undefined): string | undefined {
  return normalizeTitleColor(value);
}

/**
 * Build the optional `<c:spPr>` block at the tail of `<c:chartSpace>`
 * (the document root). Surfaces the solid fill color knob
 * ({@link SheetChart.chartSpaceFillColor}) and the border (line) color
 * knob ({@link SheetChart.chartSpaceBorderColor}) — every other
 * `<c:spPr>` child (`<a:effectLst>` effects, gradient / pattern /
 * picture fills, line dash / width / compound styles) is intentionally
 * not modelled at this layer.
 *
 * Returns `undefined` when both fields are unset / malformed so the
 * writer skips the entire `<c:spPr>` block — an empty `<c:spPr/>`
 * collapses to the inherited theme fill / stroke Excel picks anyway,
 * and omitting it keeps untouched chart XML byte-clean. When at least
 * one knob lands on the wire, the children are emitted in
 * `CT_ShapeProperties` schema order: `<a:solidFill>` (fill) then
 * `<a:ln>` (line / stroke).
 *
 * Mirrors {@link buildPlotAreaSpPr} but on a distinct host element —
 * the chart-space fill / stroke paints the entire chart frame (title
 * slot, legend slot, axis label margins, plot area together), while
 * the plot-area knobs paint only the inner band that hosts the series.
 */
function buildChartSpaceSpPr(chart: SheetChart): string | undefined {
  const fillHex = normalizeChartSpaceFillColor(chart.chartSpaceFillColor);
  const borderHex = normalizeChartSpaceBorderColor(chart.chartSpaceBorderColor);
  const borderWidthPt = clampStrokeWidthPt(chart.chartSpaceBorderWidth);
  const borderDash = normalizeBorderDash(chart.chartSpaceBorderDash);
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
    // schema sequence (ECMA-376 Part 1, §20.1.2.3.24).
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
 * Normalize a {@link SheetChart.chartSpaceFillColor} value for the
 * `<c:chartSpace><c:spPr><a:solidFill><a:srgbClr val=".."/></a:solidFill>
 * </c:spPr></c:chartSpace>` writer slot. Returns the 6-character
 * uppercase hex form when the input is a valid sRGB triple (with or
 * without a leading `#`), or `undefined` for any malformed token —
 * wrong length, non-hex characters, alpha-channel forms, or non-string
 * escapes from an untyped caller.
 *
 * Absence and malformed tokens both collapse to `undefined` so the
 * writer skips the entire `<c:spPr>` block and the chart inherits
 * the auto-fill Excel picks from the workbook theme (Excel's reference
 * behavior for a fresh chart without a custom frame color). Delegates
 * to the chart-level {@link normalizeTitleColor} so every `<a:srgbClr>`
 * fill slot shares the same sRGB grammar.
 */
function normalizeChartSpaceFillColor(value: string | undefined): string | undefined {
  return normalizeTitleColor(value);
}

/**
 * Normalize a {@link SheetChart.chartSpaceBorderColor} value for the
 * `<c:chartSpace><c:spPr><a:ln><a:solidFill><a:srgbClr val=".."/>
 * </a:solidFill></a:ln></c:spPr></c:chartSpace>` writer slot. Returns
 * the 6-character uppercase hex form when the input is a valid sRGB
 * triple (with or without a leading `#`), or `undefined` for any
 * malformed token — wrong length, non-hex characters, alpha-channel
 * forms, or non-string escapes from an untyped caller.
 *
 * Absence and malformed tokens both collapse to `undefined` so the
 * writer skips the `<a:ln>` block and the chart inherits the auto-
 * stroke Excel picks from the workbook theme (Excel's reference
 * behavior for a fresh chart without a custom border). Delegates to
 * the chart-level {@link normalizeTitleColor} so every `<a:srgbClr>`
 * fill / line slot shares the same sRGB grammar. Mirrors
 * {@link normalizeChartSpaceFillColor} — same hex grammar, distinct
 * writer slot (`<a:ln>` rather than `<a:solidFill>`).
 */
function normalizeChartSpaceBorderColor(value: string | undefined): string | undefined {
  return normalizeTitleColor(value);
}

// ── Data Table ───────────────────────────────────────────────────────


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




// ── Bar / Column ─────────────────────────────────────────────────────

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


/**
 * Application-default `sz` value for an axis title's `<a:defRPr>` /
 * `<a:rPr>` slots — Excel renders axis titles at 10pt (`sz="1000"`)
 * unless the user pins a custom size. Absence of
 * {@link SheetChart.axes.x.axisTitleFontSize} resolves to this default
 * so a fresh chart matches Excel's reference serialization byte-for-
 * byte, and round-trips of templates that never pinned the field stay
 * stable across the parse -> clone -> write loop.
 */
const AXIS_TITLE_DEFAULT_FONT_SIZE_SZ = 1000;


// ── Data Labels ──────────────────────────────────────────────────────


// ── Manual Layout ────────────────────────────────────────────────────


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

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
import { buildChartLevelDataLabels, buildSeriesDataLabels } from "./chart/dataLabels";
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
import { buildPlotArea } from "./chart/plotArea";

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

// ── Line ─────────────────────────────────────────────────────────────

// ── Area ─────────────────────────────────────────────────────────────

// ── Pie ──────────────────────────────────────────────────────────────

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

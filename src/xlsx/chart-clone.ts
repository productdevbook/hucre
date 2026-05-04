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
  ChartDataLabels,
  ChartDataLabelsInfo,
  ChartDataTable,
  ChartDisplayBlanksAs,
  ChartKind,
  ChartLegendEntry,
  ChartLineStroke,
  ChartMarker,
  ChartProtection,
  ChartScatterStyle,
  ChartSeries,
  ChartSeriesInfo,
  ChartView3D,
  SheetChart,
  WriteChartKind,
} from "../_types";

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
   * unspecified flag inside the object falls back to `true` at the
   * writer side because every `<c:dTable>` boolean child is required
   * on `CT_DTable` and Excel emits all four.
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
    const resolvedLegendOverlay = resolveLegendOverlay(source.legendOverlay, options.legendOverlay);
    if (resolvedLegendOverlay !== undefined) out.legendOverlay = resolvedLegendOverlay;

    // `<c:legendEntry>` lives inside `<c:legend>`, so the same hidden /
    // missing-legend scoping that drops `legendOverlay` also drops the
    // inherited entry list. Mirrors the legendOverlay grammar:
    // `undefined` inherits the parsed value, `null` drops it (the writer
    // emits no `<c:legendEntry>` children), a `ChartLegendEntry[]`
    // replaces it outright.
    const resolvedLegendEntries = resolveLegendEntries(source.legendEntries, options.legendEntries);
    if (resolvedLegendEntries !== undefined) out.legendEntries = resolvedLegendEntries;
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
    const dropLines = resolveDropLines(source.dropLines, options.dropLines);
    if (dropLines !== undefined) out.dropLines = dropLines;
  }

  // `<c:hiLowLines>` lives on `<c:lineChart>` / `<c:line3DChart>` /
  // `<c:stockChart>` per the OOXML schema. Hucre's writer authors
  // `<c:lineChart>` only, so the flag carries through line resolutions
  // and is dropped on every other family — coercing a line template
  // into an area clone therefore never leaks the connector lines into
  // a chart kind whose schema rejects the element.
  if (type === "line") {
    const hiLowLines = resolveHiLowLines(source.hiLowLines, options.hiLowLines);
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
    const serLines = resolveSerLines(source.serLines, options.serLines);
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
    const resolvedTitleOverlay = resolveTitleOverlay(source.titleOverlay, options.titleOverlay);
    if (resolvedTitleOverlay !== undefined) out.titleOverlay = resolvedTitleOverlay;

    // `titleRotation` only renders inside `<c:title>` — a clone that
    // omits the title has no `<a:bodyPr rot="N"/>` slot for the writer
    // to populate. Same scope rule as `titleOverlay`: the override wins
    // over the source's parsed value; absence inherits, `null` drops,
    // a `number` replaces. Out-of-range / non-finite / non-numeric
    // overrides collapse via the writer's normalizer so the cloned
    // `SheetChart` always carries a value the writer will accept.
    const resolvedTitleRotation = resolveTitleRotation(source.titleRotation, options.titleRotation);
    if (resolvedTitleRotation !== undefined) out.titleRotation = resolvedTitleRotation;
  }

  // `<c:autoTitleDeleted>` sits on `<c:chart>` directly, not inside
  // `<c:title>`, so the override carries through every clone — independent
  // of whether the resolved chart renders a literal title. Pinning the
  // flag lets a titleless clone suppress (or keep) Excel's auto-generated
  // series-name title regardless of what the source declared. The
  // override wins over the source's parsed value; absence inherits,
  // `null` drops (writer falls back to its title-presence-derived
  // default), a `boolean` replaces.
  const resolvedAutoTitleDeleted = resolveAutoTitleDeleted(
    source.autoTitleDeleted,
    options.autoTitleDeleted,
  );
  if (resolvedAutoTitleDeleted !== undefined) out.autoTitleDeleted = resolvedAutoTitleDeleted;

  const resolvedDataLabels = resolveChartDataLabels(source.dataLabels, options.dataLabels);
  if (resolvedDataLabels !== undefined) out.dataLabels = resolvedDataLabels;

  const resolvedDispBlanks = resolveDispBlanksAs(source.dispBlanksAs, options.dispBlanksAs);
  if (resolvedDispBlanks !== undefined) out.dispBlanksAs = resolvedDispBlanks;

  const resolvedVaryColors = resolveVaryColors(source.varyColors, options.varyColors);
  if (resolvedVaryColors !== undefined) out.varyColors = resolvedVaryColors;

  const resolvedPlotVisOnly = resolvePlotVisOnly(source.plotVisOnly, options.plotVisOnly);
  if (resolvedPlotVisOnly !== undefined) out.plotVisOnly = resolvedPlotVisOnly;

  const resolvedShowDLblsOverMax = resolveShowDLblsOverMax(
    source.showDLblsOverMax,
    options.showDLblsOverMax,
  );
  if (resolvedShowDLblsOverMax !== undefined) out.showDLblsOverMax = resolvedShowDLblsOverMax;

  const resolvedRoundedCorners = resolveRoundedCorners(
    source.roundedCorners,
    options.roundedCorners,
  );
  if (resolvedRoundedCorners !== undefined) out.roundedCorners = resolvedRoundedCorners;

  const resolvedStyle = resolveStyle(source.style, options.style);
  if (resolvedStyle !== undefined) out.style = resolvedStyle;

  const resolvedLang = resolveLang(source.lang, options.lang);
  if (resolvedLang !== undefined) out.lang = resolvedLang;

  const resolvedDate1904 = resolveDate1904(source.date1904, options.date1904);
  if (resolvedDate1904 !== undefined) out.date1904 = resolvedDate1904;

  // `<c:dTable>` only renders inside `<c:plotArea>` alongside the axes
  // — pie / doughnut have no axes at all, so the OOXML schema places no
  // slot for the element on those families. Drop the field on those
  // resolved types so a templated bar / line / scatter chart with a
  // pinned data table does not leak the element into a doughnut clone
  // whose schema rejects it. Override wins over the source's parsed
  // value.
  if (type !== "pie" && type !== "doughnut") {
    const resolvedDataTable = resolveDataTable(source.dataTable, options.dataTable);
    if (resolvedDataTable !== undefined) out.dataTable = resolvedDataTable;
  }

  // `<c:protection>` lives on `<c:chartSpace>` (not inside
  // `<c:plotArea>`), so every chart family — including pie / doughnut
  // — carries a slot for it. Override wins over the source's parsed
  // value, and the grammar follows the same `object | boolean | null`
  // shape as `dataTable` so the chart-level block toggles compose
  // identically at the call site.
  const resolvedProtection = resolveProtection(source.protection, options.protection);
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

  // `<c:scatterStyle>` only renders inside `<c:scatterChart>`. Drop the
  // field on every other resolved type so a scatter template flattened
  // to line / column does not leak the preset into a chart kind whose
  // schema rejects it. Override wins over the source's parsed value.
  if (type === "scatter") {
    const resolvedScatterStyle = resolveScatterStyle(source.scatterStyle, options.scatterStyle);
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
    const resolvedUpDownBars = resolveUpDownBars(source.upDownBars, options.upDownBars);
    if (resolvedUpDownBars !== undefined) out.upDownBars = resolvedUpDownBars;
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

function buildSeriesFromSource(
  source: Chart,
  overrides: ReadonlyArray<CloneChartSeriesOverride | undefined> | undefined,
): ChartSeries[] {
  const sourceSeries = source.series ?? [];
  // The override array can be longer than the source (caller wants to
  // append a fully-specified series). Walk the union of both lengths.
  const length = Math.max(sourceSeries.length, overrides?.length ?? 0);
  const out: ChartSeries[] = [];

  for (let i = 0; i < length; i++) {
    const src: ChartSeriesInfo | undefined = sourceSeries[i];
    const ov = overrides?.[i];
    const merged = mergeSeries(src, ov, i);
    out.push(merged);
  }

  return out;
}

function mergeSeries(
  src: ChartSeriesInfo | undefined,
  ov: CloneChartSeriesOverride | undefined,
  index: number,
): ChartSeries {
  // Resolve `values` first — it's the only mandatory field.
  const values = ov?.values ?? src?.valuesRef;
  if (!values) {
    throw new Error(
      `cloneChart: series #${index} has no values reference; provide \`seriesOverrides[${index}].values\``,
    );
  }

  const out: ChartSeries = { values };

  const name = applyOverride(src?.name, ov?.name);
  if (name !== undefined) out.name = name;

  const categories = applyOverride(src?.categoriesRef, ov?.categories);
  if (categories !== undefined) out.categories = categories;

  const color = applyOverride(src?.color, ov?.color);
  if (color !== undefined) out.color = color;

  const dataLabels = resolveSeriesDataLabels(src?.dataLabels, ov?.dataLabels);
  if (dataLabels !== undefined) out.dataLabels = dataLabels;

  const smooth = resolveSmooth(src?.smooth, ov?.smooth);
  if (smooth !== undefined) out.smooth = smooth;

  const stroke = resolveStroke(src?.stroke, ov?.stroke);
  if (stroke !== undefined) out.stroke = stroke;

  const marker = resolveMarker(src?.marker, ov?.marker);
  if (marker !== undefined) out.marker = marker;

  const invertIfNegative = resolveInvertIfNegative(src?.invertIfNegative, ov?.invertIfNegative);
  if (invertIfNegative !== undefined) out.invertIfNegative = invertIfNegative;

  const explosion = resolveExplosion(src?.explosion, ov?.explosion);
  if (explosion !== undefined) out.explosion = explosion;

  return out;
}

/**
 * Resolve a per-series line-stroke override.
 *
 * `undefined` → inherit the source series' `stroke` (a fresh shallow
 *               copy so the caller cannot mutate the parsed source).
 * `null`      → drop the inherited block.
 * object      → replace the inherited block wholesale (no per-field
 *               merge; pass the full shape you want).
 *
 * An empty stroke block (no dash, no width) collapses to `undefined`
 * so the writer can elide the element rather than emit a bare
 * `<a:ln/>` that Excel paints with the inherited default.
 */
function resolveStroke(
  sourceStroke: ChartLineStroke | undefined,
  override: ChartLineStroke | null | undefined,
): ChartLineStroke | undefined {
  if (override === undefined) {
    if (!sourceStroke) return undefined;
    return cloneStroke(sourceStroke);
  }
  if (override === null) return undefined;
  return cloneStroke(override);
}

function cloneStroke(source: ChartLineStroke): ChartLineStroke | undefined {
  const out: ChartLineStroke = {};
  if (source.dash !== undefined) out.dash = source.dash;
  if (typeof source.width === "number" && Number.isFinite(source.width)) out.width = source.width;
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Resolve a per-series smooth-line override.
 *
 * `undefined` → inherit the source series' `smooth`.
 * `null`      → drop the inherited flag (the cloned series renders straight).
 * `boolean`   → replace.
 *
 * Only the `true` outcome materializes on the result — `false` collapses
 * to `undefined` so absence and the OOXML default round-trip identically
 * (the writer emits straight segments either way).
 */
function resolveSmooth(
  sourceSmooth: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return sourceSmooth === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve a per-series invert-if-negative override.
 *
 * `undefined` → inherit the source series' `invertIfNegative`.
 * `null`      → drop the inherited flag (the cloned series renders
 *               negatives in the series fill color).
 * `boolean`   → replace.
 *
 * Only the `true` outcome materializes on the result — `false` collapses
 * to `undefined` so absence and the OOXML default round-trip identically
 * (the writer omits `<c:invertIfNegative>` either way).
 */
function resolveInvertIfNegative(
  sourceFlag: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return sourceFlag === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve a per-series slice-explosion override.
 *
 * `undefined` → inherit the source series' `explosion`.
 * `null`      → drop the inherited value (the cloned series renders
 *               flush against its neighbors).
 * `number`    → replace.
 *
 * Non-finite or non-positive numbers (and the OOXML default `0`)
 * collapse to `undefined` so absence and the default round-trip
 * identically through the writer's elision logic. Out-of-band values
 * (the writer also clamps) are passed through here — the writer
 * applies the final `0..400` clamp at emit time so a parsed-then-cloned
 * value remains visible on the resulting `SheetChart` object.
 */
function resolveExplosion(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) {
    if (sourceValue === undefined || !Number.isFinite(sourceValue) || sourceValue <= 0) {
      return undefined;
    }
    return sourceValue;
  }
  if (override === null) return undefined;
  if (!Number.isFinite(override) || override <= 0) return undefined;
  return override;
}

/**
 * Resolve a per-series marker override.
 *
 * `undefined` → inherit the source series' `marker` (a fresh shallow
 * copy so the caller cannot mutate the parsed source).
 * `null`      → drop the inherited block (the cloned series falls back
 *               to Excel's series-rotation default).
 * object      → replace the inherited block wholesale.
 *
 * An empty marker block (no symbol, size, or color) collapses to
 * `undefined` so the writer can elide the element rather than emit a
 * bare `<c:marker/>` that Excel paints with the inherited default.
 */
function resolveMarker(
  sourceMarker: ChartMarker | undefined,
  override: ChartMarker | null | undefined,
): ChartMarker | undefined {
  if (override === undefined) {
    if (!sourceMarker) return undefined;
    return cloneMarker(sourceMarker);
  }
  if (override === null) return undefined;
  return cloneMarker(override);
}

function cloneMarker(source: ChartMarker): ChartMarker | undefined {
  const out: ChartMarker = {};
  if (source.symbol !== undefined) out.symbol = source.symbol;
  if (typeof source.size === "number" && Number.isFinite(source.size)) out.size = source.size;
  if (typeof source.fill === "string" && source.fill.length > 0) out.fill = source.fill;
  if (typeof source.line === "string" && source.line.length > 0) out.line = source.line;
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Resolve a `dispBlanksAs` override.
 *
 * `undefined` → inherit the source's parsed `dispBlanksAs`.
 * `null`      → drop the inherited value (the writer falls back to
 *               the OOXML `"gap"` default).
 * value       → replace.
 *
 * Unknown strings are ignored (treated as `undefined`); only the three
 * OOXML-defined tokens propagate through to the writer.
 */
function resolveDispBlanksAs(
  sourceValue: ChartDisplayBlanksAs | undefined,
  override: ChartDisplayBlanksAs | null | undefined,
): ChartDisplayBlanksAs | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

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
function resolveVaryColors(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `plotVisOnly` override.
 *
 * `undefined` → inherit the source's parsed `plotVisOnly`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `true` default — hidden cells drop out of the chart).
 * `boolean`   → replace.
 *
 * The grammar mirrors `dispBlanksAs` / `varyColors` so the chart-level
 * toggles compose the same way at the call site.
 */
function resolvePlotVisOnly(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `showDLblsOverMax` override.
 *
 * `undefined` → inherit the source's parsed `showDLblsOverMax`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `true` default — labels render for every point
 *               regardless of the pinned axis ceiling).
 * `boolean`   → replace.
 *
 * The grammar mirrors `plotVisOnly` / `dispBlanksAs` so the chart-level
 * toggles compose the same way at the call site.
 */
function resolveShowDLblsOverMax(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `roundedCorners` override.
 *
 * `undefined` → inherit the source's parsed `roundedCorners`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `false` default — square chart frame).
 * `boolean`   → replace.
 *
 * The grammar mirrors `plotVisOnly` / `varyColors` so the chart-frame
 * toggles compose the same way at the call site.
 */
function resolveRoundedCorners(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `style` (built-in chart preset) override.
 *
 * `undefined` → inherit the source's parsed `style`.
 * `null`      → drop the inherited value (the writer skips `<c:style>`
 *               so Excel falls back to its application default look).
 * `number`    → replace. Out-of-range / non-integer values are not
 *               filtered here — the writer's `resolveStyle` performs
 *               the same shape check on emit, so a stray value never
 *               reaches the rendered XML regardless of the path it
 *               took through clone.
 *
 * The grammar mirrors `roundedCorners` / `plotVisOnly` so the chart-
 * frame toggles compose the same way at the call site.
 */
function resolveStyle(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `lang` (chart-space editing-locale hint) override.
 *
 * `undefined` → inherit the source's parsed `lang`.
 * `null`      → drop the inherited value (the writer skips `<c:lang>`
 *               so Excel falls back to the host workbook's editing
 *               language).
 * `string`    → replace. Malformed culture names are not filtered
 *               here — the writer's `resolveLang` performs the same
 *               BCP-47 shape check on emit, so a stray value never
 *               reaches the rendered XML regardless of the path it
 *               took through clone.
 *
 * The grammar mirrors `style` / `roundedCorners` / `plotVisOnly` so
 * the chart-space toggles compose the same way at the call site.
 */
function resolveLang(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `date1904` (chart-space date-system) override.
 *
 * `undefined` → inherit the source's parsed `date1904`.
 * `null`      → drop the inherited value (the writer skips
 *               `<c:date1904>` so Excel falls back to the host
 *               workbook's date system).
 * `boolean`   → replace. `false` collapses to absence on the writer
 *               side because `<c:date1904 val="0"/>` is the OOXML
 *               default and the writer follows the minimal-shape
 *               contract every other chart-space toggle uses.
 *
 * The grammar mirrors `roundedCorners` / `plotVisOnly` so the
 * chart-space toggles compose the same way at the call site. `false`
 * here means "explicitly pin the 1900 base" — but because absence
 * and `val="0"` round-trip identically the resolved value still
 * collapses to `undefined` (the writer would emit nothing either
 * way).
 */
function resolveDate1904(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  const merged = override === undefined ? sourceValue : override === null ? undefined : override;
  if (merged === true) return true;
  // `false` and `undefined` both collapse to `undefined` — absence
  // and the OOXML default `<c:date1904 val="0"/>` round-trip the
  // same way through parseChart -> cloneChart -> writeChart, so the
  // resolved chart drops the field rather than carry a value the
  // writer would skip on emit anyway.
  return undefined;
}

/**
 * Resolve a `dataTable` (plot-area data-table) override.
 *
 * `undefined` → inherit the source's parsed {@link Chart.dataTable}.
 * `null`      → drop the inherited block so the writer skips
 *               `<c:dTable>` entirely (no data table rendered).
 * `false`     → equivalent to `null` (suppression); kept distinct in
 *               the API surface so callers can write `dataTable: false`
 *               for symmetry with the writer's `boolean | object` shape.
 * `true`      → enable with the OOXML reference defaults (every flag
 *               `true`).
 * `object`    → replace the inherited block wholesale (no per-field
 *               merge with the source — pass every flag the cloned
 *               table should render). Each unspecified field falls back
 *               to `true` at the writer side because every `<c:dTable>`
 *               boolean child is required on `CT_DTable` and Excel
 *               always emits all four.
 *
 * The grammar mirrors {@link CloneChartSeriesOverride.marker} (and the
 * other `object | null` / wholesale-replace patterns) so the
 * chart-level block toggles compose the same way at the call site.
 *
 * The caller already short-circuits this for pie / doughnut clones
 * because the OOXML schema places `<c:dTable>` inside `<c:plotArea>`
 * alongside the axes, and pie / doughnut have no axes at all.
 */
function resolveDataTable(
  sourceValue: ChartDataTable | undefined,
  override: ChartDataTable | boolean | null | undefined,
): ChartDataTable | boolean | undefined {
  if (override === undefined) {
    // Inherit — pass the source through verbatim. The writer accepts
    // both the boolean and object shapes, so a parsed `ChartDataTable`
    // round-trips directly.
    return sourceValue;
  }
  if (override === null) {
    // Drop the inherited block. The writer treats `undefined` as
    // suppression and skips `<c:dTable>` entirely.
    return undefined;
  }
  if (override === false) {
    // Symmetric with `null` — kept distinct in the API surface for
    // ergonomic alignment with the writer's `boolean | object` shape,
    // but emits the same on-the-wire result (no `<c:dTable>`).
    return undefined;
  }
  // `true` or a {@link ChartDataTable} object — replace the inherited
  // block wholesale. The writer accepts both forms and falls back to
  // the OOXML reference defaults for any field the object leaves unset.
  return override;
}

/**
 * Resolve a `protection` (chart-space protection) override.
 *
 * `undefined` → inherit the source's parsed {@link Chart.protection}.
 * `null`      → drop the inherited block so the writer skips
 *               `<c:protection>` entirely (no chart-level lock).
 * `false`     → equivalent to `null` (suppression); kept distinct in
 *               the API surface so callers can write `protection:
 *               false` for symmetry with the writer's `boolean |
 *               object` shape.
 * `true`      → enable with the OOXML reference defaults (every flag
 *               `false` — the bare `<c:protection>` shell).
 * `object`    → replace the inherited block wholesale (no per-field
 *               merge with the source — pass every flag the cloned
 *               protection should pin). Each unspecified field falls
 *               back to `false` at the writer side because every
 *               `<c:protection>` boolean child is independently
 *               optional and Excel treats a missing child as
 *               `false`.
 *
 * The grammar mirrors {@link resolveDataTable} so the chart-level
 * block toggles compose the same way at the call site. Unlike
 * `dataTable`, `<c:protection>` lives on `<c:chartSpace>` (not inside
 * `<c:plotArea>`) so the resolver applies to every chart family —
 * pie / doughnut included.
 */
function resolveProtection(
  sourceValue: ChartProtection | undefined,
  override: ChartProtection | boolean | null | undefined,
): ChartProtection | boolean | undefined {
  if (override === undefined) {
    // Inherit — pass the source through verbatim. The writer accepts
    // both the boolean and object shapes, so a parsed
    // {@link ChartProtection} round-trips directly.
    return sourceValue;
  }
  if (override === null) {
    // Drop the inherited block. The writer treats `undefined` as
    // suppression and skips `<c:protection>` entirely.
    return undefined;
  }
  if (override === false) {
    // Symmetric with `null` — kept distinct in the API surface for
    // ergonomic alignment with the writer's `boolean | object` shape,
    // but emits the same on-the-wire result (no `<c:protection>`).
    return undefined;
  }
  // `true` or a {@link ChartProtection} object — replace the inherited
  // block wholesale. The writer accepts both forms and falls back to
  // the OOXML reference default `false` for any field the object
  // leaves unset.
  return override;
}

/**
 * Resolve a `view3D` override.
 *
 * `undefined` → inherit the source's parsed {@link Chart.view3D}. The
 *               source's parsed object is defensively shallow-copied
 *               so a downstream mutation to the cloned SheetChart
 *               never leaks back into the parsed Chart.
 * `null`      → drop the inherited block so the writer skips
 *               `<c:view3D>` entirely (no chart-level 3D pin).
 * `object`    → replace the inherited block wholesale (no per-field
 *               merge with the source — pass every field the cloned
 *               view3D should pin). An empty object emits a bare
 *               `<c:view3D/>` shell at the writer side.
 *
 * The grammar mirrors {@link resolveProtection} / {@link resolveDataTable}
 * so the chart-level block toggles compose the same way at the call
 * site. Unlike `dataTable`, `<c:view3D>` lives on `<c:chart>` (not
 * inside `<c:plotArea>`) so the resolver applies to every chart family
 * — pie / doughnut included.
 */
function resolveView3D(
  sourceValue: ChartView3D | undefined,
  override: ChartView3D | null | undefined,
): ChartView3D | undefined {
  if (override === undefined) {
    // Inherit — defensively shallow-copy the source so a downstream
    // mutation to the cloned SheetChart never leaks back into the
    // parsed Chart. The CT_View3D children are all scalars (numbers
    // and a boolean), so a single-level spread is enough.
    if (sourceValue === undefined) return undefined;
    return { ...sourceValue };
  }
  if (override === null) {
    // Drop the inherited block. The writer treats `undefined` as
    // suppression and skips `<c:view3D>` entirely.
    return undefined;
  }
  // Replace the inherited block wholesale. The writer accepts the
  // empty-object shape and emits a bare `<c:view3D/>` shell, mirroring
  // how `resolveProtection` handles the `true` / `{}` forms.
  return { ...override };
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
function resolveUpDownBars(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `showLineMarkers` override.
 *
 * `undefined` → inherit the source's parsed `showLineMarkers`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               Excel default — `<c:marker val="1"/>`, markers shown).
 * `boolean`   → replace.
 *
 * The grammar mirrors `upDownBars` / `dropLines` / `hiLowLines` so the
 * chart-level line-only toggles compose the same way at the call site.
 * `true` collapses to `undefined` on the writer side because the writer
 * already emits `val="1"` by default; the `true` value still surfaces
 * in the cloned `SheetChart` for symmetry with other resolve helpers,
 * leaving the renderer to fold it back into the default during emit.
 */
function resolveShowLineMarkers(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `legendOverlay` override.
 *
 * `undefined` → inherit the source's parsed `legendOverlay`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `false` default — the legend reserves its own
 *               slot, no overlap with the plot area).
 * `boolean`   → replace.
 *
 * The grammar mirrors `plotVisOnly` / `roundedCorners` so the chart-
 * level toggles compose the same way at the call site. Callers should
 * gate the result on the resolved legend visibility — when no legend
 * is emitted, the overlay flag has no slot in the rendered chart.
 */
function resolveLegendOverlay(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `legendEntries` override.
 *
 * `undefined` → inherit the source's parsed `legendEntries`.
 * `null`      → drop the inherited list (the writer emits no
 *               `<c:legendEntry>` children).
 * `array`     → replace the inherited list outright. Empty arrays
 *               collapse to `undefined` so the writer never emits an
 *               empty selector block — Excel's reference serialization
 *               omits the children entirely when no entry is hidden.
 *
 * Callers should gate the result on the resolved legend visibility —
 * when no legend is emitted, the entry list has no slot in the rendered
 * chart. Mirrors the `legendOverlay` grammar so the legend-scoped
 * fields compose the same way at the call site.
 *
 * The returned array is always a fresh copy of the source / override
 * (never a shared reference) so a downstream mutation to the cloned
 * `SheetChart` never leaks back into the parsed `Chart` the caller
 * passed in. Each entry is also copied to keep the writer's resolution
 * pass free to dedupe / sort without touching the inputs.
 */
function resolveLegendEntries(
  sourceValue: ChartLegendEntry[] | undefined,
  override: ChartLegendEntry[] | null | undefined,
): ChartLegendEntry[] | undefined {
  if (override === undefined) {
    if (!sourceValue || sourceValue.length === 0) return undefined;
    return sourceValue.map((entry) => ({ ...entry }));
  }
  if (override === null) return undefined;
  if (!Array.isArray(override) || override.length === 0) return undefined;
  return override.map((entry) => ({ ...entry }));
}

/**
 * Resolve a `titleOverlay` override.
 *
 * `undefined` → inherit the source's parsed `titleOverlay`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `false` default — the title reserves its own slot
 *               above the plot area, no overlap with it).
 * `boolean`   → replace.
 *
 * The grammar mirrors `legendOverlay` / `roundedCorners` so the chart-
 * level overlay toggles compose the same way at the call site. Callers
 * should gate the result on the resolved title visibility — when no
 * title is emitted, the overlay flag has no slot in the rendered chart.
 */
function resolveTitleOverlay(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Conversion bookkeeping for the chart-title rotation override. Same
 * `-90..90` band the writer enforces so the cloned `SheetChart` always
 * carries a value the writer will accept.
 */
const TITLE_ROTATION_MIN_DEG = -90;
const TITLE_ROTATION_MAX_DEG = 90;

/**
 * Normalize a `titleRotation` value (whole degrees) for the cloned
 * `SheetChart`. Mirrors the writer's `normalizeTitleRotation` — the
 * cloned shape is guaranteed to round-trip through the writer without
 * surprise: out-of-range inputs clamp to the `-90..90` band; non-integer
 * inputs round to the nearest whole degree; `0`, `NaN`, `Infinity`, and
 * non-numeric inputs collapse to `undefined` so the cloned chart drops
 * the field rather than carry a value the writer would silently elide.
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
 * Resolve a `titleRotation` override.
 *
 * `undefined` → inherit the source's parsed `titleRotation`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `0` default — the title renders horizontally).
 * `number`    → replace, after clamping / rounding through
 *               {@link normalizeTitleRotation}.
 *
 * The grammar mirrors `titleOverlay` / `legendOverlay` so the chart-
 * level title knobs compose the same way at the call site. Callers
 * should gate the result on the resolved title visibility — when no
 * title is emitted, the rotation has no slot in the rendered chart.
 */
function resolveTitleRotation(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return normalizeTitleRotation(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleRotation(override);
}

/**
 * Resolve an `autoTitleDeleted` override.
 *
 * `undefined` → inherit the source's parsed `autoTitleDeleted`.
 * `null`      → drop the inherited value (the writer falls back to its
 *               title-presence-derived default — `val="0"` when the
 *               cloned chart has a literal title, `val="1"` when it
 *               does not).
 * `boolean`   → replace.
 *
 * The grammar mirrors `titleOverlay` / `roundedCorners` /
 * `plotVisOnly` so the chart-level title flags compose the same way
 * at the call site. Independent of the resolved title presence —
 * `<c:autoTitleDeleted>` sits on `<c:chart>` directly (between
 * `<c:title>` and `<c:plotArea>` per CT_Chart, ECMA-376 Part 1,
 * §21.2.2.4), not nested inside `<c:title>`, so a clone with no
 * literal title can still pin `false` (let Excel synthesise the
 * series-name auto-title) and a clone with a literal title can pin
 * `true` (suppress the synthesis even though the literal renders).
 */
function resolveAutoTitleDeleted(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
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
function resolveDropLines(
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
 * Resolve a `hiLowLines` override. Mirrors {@link resolveDropLines};
 * the only difference is the per-family scope — `<c:hiLowLines>` has
 * no slot on `<c:areaChart>`.
 */
function resolveHiLowLines(
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
 * Resolve a `serLines` override. Mirrors {@link resolveDropLines} /
 * {@link resolveHiLowLines}; the only difference is the per-family
 * scope — `<c:serLines>` has no slot on `<c:lineChart>` /
 * `<c:areaChart>` / `<c:pieChart>` / `<c:doughnutChart>` /
 * `<c:scatterChart>`.
 */
function resolveSerLines(
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
function resolveScatterStyle(
  sourceValue: ChartScatterStyle | undefined,
  override: ChartScatterStyle | null | undefined,
): ChartScatterStyle | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a chart-level data-labels override.
 *
 * `undefined` → inherit the source's parsed `dataLabels` (downcast from
 * the read-side {@link ChartDataLabelsInfo} to the write-side
 * {@link ChartDataLabels} shape — they share field semantics).
 * `null`      → drop the inherited block.
 * object      → replace.
 */
function resolveChartDataLabels(
  sourceLabels: ChartDataLabelsInfo | undefined,
  override: ChartDataLabels | null | undefined,
): ChartDataLabels | undefined {
  if (override === undefined) {
    return sourceLabels ? { ...sourceLabels } : undefined;
  }
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a per-series data-labels override.
 *
 * `undefined` → inherit the source series' `dataLabels`.
 * `null`      → drop the inherited block (series will fall back to
 *               whatever the chart-level default is at write time).
 * `false`     → suppress labels on this series alone.
 * object      → replace.
 */
function resolveSeriesDataLabels(
  sourceLabels: ChartDataLabelsInfo | undefined,
  override: ChartDataLabels | false | null | undefined,
): ChartDataLabels | false | undefined {
  if (override === undefined) {
    return sourceLabels ? { ...sourceLabels } : undefined;
  }
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a "source value + optional override" pair where the override
 * may be `undefined` (no override), `null` (drop the source value), or
 * a string (replace).
 */
function applyOverride(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Merge the source chart's `axes` block with per-axis overrides. The
 * result mirrors the writer's {@link SheetChart.axes} shape — missing
 * fields are dropped so the writer doesn't emit empty `<c:title>`
 * elements or redundant gridline blocks.
 */
function resolveAxes(
  sourceAxes: Chart["axes"],
  overrides: CloneChartOptions["axes"],
  type: WriteChartKind,
): SheetChart["axes"] | undefined {
  const xTitle = applyOverride(sourceAxes?.x?.title, overrides?.x?.title);
  const yTitle = applyOverride(sourceAxes?.y?.title, overrides?.y?.title);
  const xGridlines = applyGridlinesOverride(sourceAxes?.x?.gridlines, overrides?.x?.gridlines);
  const yGridlines = applyGridlinesOverride(sourceAxes?.y?.gridlines, overrides?.y?.gridlines);
  const xScale = applyScaleOverride(sourceAxes?.x?.scale, overrides?.x?.scale);
  const yScale = applyScaleOverride(sourceAxes?.y?.scale, overrides?.y?.scale);
  const xNumFmt = applyNumberFormatOverride(
    sourceAxes?.x?.numberFormat,
    overrides?.x?.numberFormat,
  );
  const yNumFmt = applyNumberFormatOverride(
    sourceAxes?.y?.numberFormat,
    overrides?.y?.numberFormat,
  );
  const xMajorTickMark = applyTickMarkOverride(
    sourceAxes?.x?.majorTickMark,
    overrides?.x?.majorTickMark,
  );
  const yMajorTickMark = applyTickMarkOverride(
    sourceAxes?.y?.majorTickMark,
    overrides?.y?.majorTickMark,
  );
  const xMinorTickMark = applyTickMarkOverride(
    sourceAxes?.x?.minorTickMark,
    overrides?.x?.minorTickMark,
  );
  const yMinorTickMark = applyTickMarkOverride(
    sourceAxes?.y?.minorTickMark,
    overrides?.y?.minorTickMark,
  );
  const xTickLblPos = applyTickLblPosOverride(sourceAxes?.x?.tickLblPos, overrides?.x?.tickLblPos);
  const yTickLblPos = applyTickLblPosOverride(sourceAxes?.y?.tickLblPos, overrides?.y?.tickLblPos);
  // `<c:txPr><a:bodyPr rot="N"/></c:txPr>` lives on every axis flavour
  // per the OOXML schema (CT_CatAx, CT_ValAx, CT_DateAx, CT_SerAx all
  // carry an optional `<c:txPr>`), so the resolver applies on every
  // chart family that has axes (pie / doughnut were short-circuited
  // upstream). Out-of-range / non-numeric values clamp to the
  // `-90..90` band the writer accepts; the OOXML default `0` collapses
  // to `undefined` so absence and the default round-trip identically.
  const xLabelRotation = applyLabelRotationOverride(
    sourceAxes?.x?.labelRotation,
    overrides?.x?.labelRotation,
  );
  const yLabelRotation = applyLabelRotationOverride(
    sourceAxes?.y?.labelRotation,
    overrides?.y?.labelRotation,
  );
  const xReverse = applyReverseOverride(sourceAxes?.x?.reverse, overrides?.x?.reverse);
  const yReverse = applyReverseOverride(sourceAxes?.y?.reverse, overrides?.y?.reverse);
  // `tickLblSkip` / `tickMarkSkip` only render on category axes
  // (`<c:catAx>`). Scatter charts use two value axes, so the X axis
  // skip would be silently dropped by the writer anyway — collapse it
  // to undefined here so the cloned `SheetChart` accurately reflects
  // what the chart will paint.
  const isCatAxisX = type !== "scatter";
  const xTickLblSkip = isCatAxisX
    ? applySkipOverride(sourceAxes?.x?.tickLblSkip, overrides?.x?.tickLblSkip)
    : undefined;
  const xTickMarkSkip = isCatAxisX
    ? applySkipOverride(sourceAxes?.x?.tickMarkSkip, overrides?.x?.tickMarkSkip)
    : undefined;
  // `lblOffset` is also category-axis-only (CT_CatAx / CT_DateAx) per
  // the OOXML schema. Same scope rule as the skip elements above.
  const xLblOffset = isCatAxisX
    ? applyLblOffsetOverride(sourceAxes?.x?.lblOffset, overrides?.x?.lblOffset)
    : undefined;
  // `lblAlgn` is category-axis-only as well (CT_CatAx / CT_DateAx
  // per ECMA-376 §21.2.2). Same scope as `lblOffset`.
  const xLblAlgn = isCatAxisX
    ? applyLblAlgnOverride(sourceAxes?.x?.lblAlgn, overrides?.x?.lblAlgn)
    : undefined;
  // `noMultiLvlLbl` is even tighter — `CT_CatAx` only (no `<c:dateAx>`
  // slot per ECMA-376 §21.2.2). Reuse the catAx scope rule above; the
  // resolved chart type still funnels through `<c:catAx>` for every
  // bar / column / line / area family the writer supports.
  const xNoMultiLvlLbl = isCatAxisX
    ? applyNoMultiLvlLblOverride(sourceAxes?.x?.noMultiLvlLbl, overrides?.x?.noMultiLvlLbl)
    : undefined;
  // `<c:auto>` is also `CT_CatAx`-only per ECMA-376 §21.2.2.7 — same
  // scope rule as `noMultiLvlLbl`. The flag defaults to `true` in the
  // OOXML schema (Excel auto-detects whether to render the axis as a
  // date or category axis), so the resolver collapses `true` to
  // `undefined` and only surfaces an explicit `false`.
  const xAuto = isCatAxisX ? applyAutoOverride(sourceAxes?.x?.auto, overrides?.x?.auto) : undefined;
  // `<c:delete>` lives on every axis flavour — both `<c:catAx>` and
  // `<c:valAx>` accept it — so the hidden flag carries through every
  // chart family that has axes. Pie / doughnut have no axes at all
  // and the caller already short-circuited those above.
  const xHidden = applyHiddenOverride(sourceAxes?.x?.hidden, overrides?.x?.hidden);
  const yHidden = applyHiddenOverride(sourceAxes?.y?.hidden, overrides?.y?.hidden);
  // `<c:crosses>` and `<c:crossesAt>` live in an XSD choice on every
  // axis flavour. Resolve the pair together so the precedence rule
  // (numeric pin wins over semantic token) survives the inherit / null
  // / replace grammar — a `crossesAt` override of `null` falls through
  // to the (possibly inherited) semantic `crosses`, and vice versa.
  const xCrossesPair = applyCrossesOverride(
    { crosses: sourceAxes?.x?.crosses, crossesAt: sourceAxes?.x?.crossesAt },
    { crosses: overrides?.x?.crosses, crossesAt: overrides?.x?.crossesAt },
  );
  const yCrossesPair = applyCrossesOverride(
    { crosses: sourceAxes?.y?.crosses, crossesAt: sourceAxes?.y?.crossesAt },
    { crosses: overrides?.y?.crosses, crossesAt: overrides?.y?.crossesAt },
  );
  // `<c:dispUnits>` lives exclusively on `<c:valAx>` per ECMA-376
  // §21.2.2.32 (CT_ValAx → CT_DispUnits). Bar / column / line / area
  // route the X axis through `<c:catAx>`, so the X-axis override is
  // only honoured when the resolved chart type is `scatter` (both axes
  // are value axes). Pie / doughnut were already short-circuited
  // upstream — they have no axes at all. The Y axis is a value axis on
  // every remaining family, so the Y override always carries through.
  const xDispUnits =
    type === "scatter"
      ? applyDispUnitsOverride(sourceAxes?.x?.dispUnits, overrides?.x?.dispUnits)
      : undefined;
  const yDispUnits = applyDispUnitsOverride(sourceAxes?.y?.dispUnits, overrides?.y?.dispUnits);
  // `<c:crossBetween>` is also value-axis-only per ECMA-376 §21.2.2.10
  // (CT_ValAx → CT_CrossBetween). Same scope rule as `dispUnits` — the
  // X-axis override is only honoured on scatter (both axes are value
  // axes); bar / column / line / area route X through `<c:catAx>` which
  // rejects `<c:crossBetween>`. The Y axis is a value axis on every
  // family that has axes, so the Y override always carries through.
  const xCrossBetween =
    type === "scatter"
      ? applyCrossBetweenOverride(sourceAxes?.x?.crossBetween, overrides?.x?.crossBetween)
      : undefined;
  const yCrossBetween = applyCrossBetweenOverride(
    sourceAxes?.y?.crossBetween,
    overrides?.y?.crossBetween,
  );

  const out: NonNullable<SheetChart["axes"]> = {};
  if (
    xTitle !== undefined ||
    xGridlines !== undefined ||
    xScale !== undefined ||
    xNumFmt !== undefined ||
    xMajorTickMark !== undefined ||
    xMinorTickMark !== undefined ||
    xTickLblPos !== undefined ||
    xLabelRotation !== undefined ||
    xReverse !== undefined ||
    xTickLblSkip !== undefined ||
    xTickMarkSkip !== undefined ||
    xLblOffset !== undefined ||
    xLblAlgn !== undefined ||
    xNoMultiLvlLbl !== undefined ||
    xAuto !== undefined ||
    xHidden !== undefined ||
    xCrossesPair.crosses !== undefined ||
    xCrossesPair.crossesAt !== undefined ||
    xDispUnits !== undefined ||
    xCrossBetween !== undefined
  ) {
    out.x = {};
    if (xTitle !== undefined) out.x.title = xTitle;
    if (xGridlines !== undefined) out.x.gridlines = xGridlines;
    if (xScale !== undefined) out.x.scale = xScale;
    if (xNumFmt !== undefined) out.x.numberFormat = xNumFmt;
    if (xMajorTickMark !== undefined) out.x.majorTickMark = xMajorTickMark;
    if (xMinorTickMark !== undefined) out.x.minorTickMark = xMinorTickMark;
    if (xTickLblPos !== undefined) out.x.tickLblPos = xTickLblPos;
    if (xLabelRotation !== undefined) out.x.labelRotation = xLabelRotation;
    if (xReverse !== undefined) out.x.reverse = xReverse;
    if (xTickLblSkip !== undefined) out.x.tickLblSkip = xTickLblSkip;
    if (xTickMarkSkip !== undefined) out.x.tickMarkSkip = xTickMarkSkip;
    if (xLblOffset !== undefined) out.x.lblOffset = xLblOffset;
    if (xLblAlgn !== undefined) out.x.lblAlgn = xLblAlgn;
    if (xNoMultiLvlLbl !== undefined) out.x.noMultiLvlLbl = xNoMultiLvlLbl;
    if (xAuto !== undefined) out.x.auto = xAuto;
    if (xHidden !== undefined) out.x.hidden = xHidden;
    if (xCrossesPair.crosses !== undefined) out.x.crosses = xCrossesPair.crosses;
    if (xCrossesPair.crossesAt !== undefined) out.x.crossesAt = xCrossesPair.crossesAt;
    if (xDispUnits !== undefined) out.x.dispUnits = xDispUnits;
    if (xCrossBetween !== undefined) out.x.crossBetween = xCrossBetween;
  }
  if (
    yTitle !== undefined ||
    yGridlines !== undefined ||
    yScale !== undefined ||
    yNumFmt !== undefined ||
    yMajorTickMark !== undefined ||
    yMinorTickMark !== undefined ||
    yTickLblPos !== undefined ||
    yLabelRotation !== undefined ||
    yHidden !== undefined ||
    yReverse !== undefined ||
    yCrossesPair.crosses !== undefined ||
    yCrossesPair.crossesAt !== undefined ||
    yDispUnits !== undefined ||
    yCrossBetween !== undefined
  ) {
    out.y = {};
    if (yTitle !== undefined) out.y.title = yTitle;
    if (yGridlines !== undefined) out.y.gridlines = yGridlines;
    if (yScale !== undefined) out.y.scale = yScale;
    if (yNumFmt !== undefined) out.y.numberFormat = yNumFmt;
    if (yMajorTickMark !== undefined) out.y.majorTickMark = yMajorTickMark;
    if (yMinorTickMark !== undefined) out.y.minorTickMark = yMinorTickMark;
    if (yTickLblPos !== undefined) out.y.tickLblPos = yTickLblPos;
    if (yLabelRotation !== undefined) out.y.labelRotation = yLabelRotation;
    if (yHidden !== undefined) out.y.hidden = yHidden;
    if (yReverse !== undefined) out.y.reverse = yReverse;
    if (yCrossesPair.crosses !== undefined) out.y.crosses = yCrossesPair.crosses;
    if (yCrossesPair.crossesAt !== undefined) out.y.crossesAt = yCrossesPair.crossesAt;
    if (yDispUnits !== undefined) out.y.dispUnits = yDispUnits;
    if (yCrossBetween !== undefined) out.y.crossBetween = yCrossBetween;
  }

  return out.x || out.y ? out : undefined;
}

/**
 * Resolve a `tickLblSkip` / `tickMarkSkip` override using the same
 * `undefined` (inherit) / `null` (drop) / value (replace) grammar as
 * the other axis helpers. Out-of-range / non-positive values collapse
 * to `undefined` so they cannot leak into the writer (which would
 * silently drop them anyway via {@link normalizeAxisSkip}).
 */
function applySkipOverride(
  source: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) {
    if (typeof source !== "number" || !Number.isFinite(source)) return undefined;
    const rounded = Math.round(source);
    if (rounded < 1 || rounded > 32767 || rounded === 1) return undefined;
    return rounded;
  }
  if (override === null) return undefined;
  if (typeof override !== "number" || !Number.isFinite(override)) return undefined;
  const rounded = Math.round(override);
  if (rounded < 1 || rounded > 32767 || rounded === 1) return undefined;
  return rounded;
}

/**
 * Resolve an `lblOffset` override using the same `undefined` (inherit)
 * / `null` (drop) / value (replace) grammar as the other axis helpers.
 * Out-of-range / non-numeric values collapse to `undefined` so they
 * cannot leak into the writer (which would silently drop them anyway
 * via {@link normalizeAxisLblOffset}). The OOXML default `100` also
 * collapses to `undefined` so absence and the default round-trip
 * identically — symmetric with the parser-side default-collapse.
 */
function applyLblOffsetOverride(
  source: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) {
    if (typeof source !== "number" || !Number.isFinite(source)) return undefined;
    const rounded = Math.round(source);
    if (rounded < 0 || rounded > 1000 || rounded === 100) return undefined;
    return rounded;
  }
  if (override === null) return undefined;
  if (typeof override !== "number" || !Number.isFinite(override)) return undefined;
  const rounded = Math.round(override);
  if (rounded < 0 || rounded > 1000 || rounded === 100) return undefined;
  return rounded;
}

/**
 * Resolve an `lblAlgn` override using the same `undefined` (inherit)
 * / `null` (drop) / value (replace) grammar as the other axis helpers.
 * Unknown tokens collapse to `undefined` so they cannot leak into the
 * writer (which would silently drop them anyway via
 * {@link normalizeAxisLblAlgn}). The OOXML default `"ctr"` also
 * collapses to `undefined` so absence and the default round-trip
 * identically — symmetric with the parser-side default-collapse.
 */
function applyLblAlgnOverride(
  source: ChartAxisLabelAlign | undefined,
  override: ChartAxisLabelAlign | null | undefined,
): ChartAxisLabelAlign | undefined {
  if (override === undefined) {
    if (source !== "l" && source !== "r" && source !== "ctr") return undefined;
    return source === "ctr" ? undefined : source;
  }
  if (override === null) return undefined;
  if (override !== "l" && override !== "r" && override !== "ctr") return undefined;
  return override === "ctr" ? undefined : override;
}

/**
 * Resolve a `noMultiLvlLbl` override using the same `undefined`
 * (inherit) / `null` (drop) / `boolean` (replace) grammar as the
 * other axis helpers. Only `true` surfaces (the writer treats `false`
 * and absence identically — both produce `<c:noMultiLvlLbl val="0"/>`),
 * so an override of `false` collapses to `undefined` to keep the
 * cloned `SheetChart` shape minimal. Non-boolean inputs fall through
 * the type guard to `undefined`.
 */
function applyNoMultiLvlLblOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return source === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve an `auto` override using the same `undefined` (inherit) /
 * `null` (drop) / `boolean` (replace) grammar as the other axis
 * helpers. Only `false` surfaces (the writer treats `true` and absence
 * identically — both produce `<c:auto val="1"/>`), so an override of
 * `true` collapses to `undefined` to keep the cloned `SheetChart`
 * shape minimal. Non-boolean inputs fall through the type guard to
 * `undefined`.
 *
 * Inverse of {@link applyNoMultiLvlLblOverride}: `<c:auto>` defaults to
 * `true` in the OOXML schema, so the helper collapses `true` rather
 * than `false` — symmetric with the parser-side default-collapse on
 * the read layer.
 */
function applyAutoOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return source === false ? false : undefined;
  }
  if (override === null) return undefined;
  return override === false ? false : undefined;
}

/**
 * Resolve an axis `hidden` override using the same `undefined`
 * (inherit) / `null` (drop) / `boolean` (replace) grammar as the
 * other axis helpers. Only `true` surfaces (the writer treats `false`
 * and absence identically — both produce `<c:delete val="0"/>`), so
 * an override of `false` collapses to `undefined` to keep the cloned
 * `SheetChart` shape minimal. Non-boolean inputs fall through the
 * type guard to `undefined`.
 */
function applyHiddenOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return source === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve gridlines using the same `undefined` (inherit) / `null`
 * (drop) / object (replace) grammar as the other axis overrides.
 * Returns `undefined` when neither source nor override declares a
 * non-empty gridline configuration.
 */
function applyGridlinesOverride(
  source: ChartAxisGridlines | undefined,
  override: ChartAxisGridlines | null | undefined,
): ChartAxisGridlines | undefined {
  if (override === undefined) {
    if (!source) return undefined;
    const out: ChartAxisGridlines = {};
    if (source.major) out.major = true;
    if (source.minor) out.minor = true;
    return out.major || out.minor ? out : undefined;
  }
  if (override === null) return undefined;
  const out: ChartAxisGridlines = {};
  if (override.major === true) out.major = true;
  if (override.minor === true) out.minor = true;
  return out.major || out.minor ? out : undefined;
}

/**
 * Resolve a scale override using the same `undefined` / `null` /
 * object grammar as {@link applyGridlinesOverride}. The override
 * replaces the source wholesale rather than merging field-by-field —
 * a partial template scale `{ min: 0 }` plus an override
 * `{ max: 100 }` yields `{ max: 100 }`, not `{ min: 0, max: 100 }`.
 * Per-field merges proved confusing in the dashboard composition flow
 * (callers expected the override to fully describe the target scale),
 * so wholesale replacement is the simpler contract.
 */
function applyScaleOverride(
  source: ChartAxisScale | undefined,
  override: ChartAxisScale | null | undefined,
): ChartAxisScale | undefined {
  if (override === undefined) {
    if (!source) return undefined;
    return cloneScale(source);
  }
  if (override === null) return undefined;
  return cloneScale(override);
}

function cloneScale(source: ChartAxisScale): ChartAxisScale | undefined {
  const out: ChartAxisScale = {};
  if (typeof source.min === "number" && Number.isFinite(source.min)) out.min = source.min;
  if (typeof source.max === "number" && Number.isFinite(source.max)) out.max = source.max;
  if (
    typeof source.majorUnit === "number" &&
    Number.isFinite(source.majorUnit) &&
    source.majorUnit > 0
  ) {
    out.majorUnit = source.majorUnit;
  }
  if (
    typeof source.minorUnit === "number" &&
    Number.isFinite(source.minorUnit) &&
    source.minorUnit > 0
  ) {
    out.minorUnit = source.minorUnit;
  }
  if (typeof source.logBase === "number" && Number.isFinite(source.logBase)) {
    out.logBase = source.logBase;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Resolve a number-format override. Same grammar as the other
 * per-axis helpers: `undefined` inherits, `null` drops, an object
 * replaces.
 */
function applyNumberFormatOverride(
  source: ChartAxisNumberFormat | undefined,
  override: ChartAxisNumberFormat | null | undefined,
): ChartAxisNumberFormat | undefined {
  if (override === undefined) {
    if (!source) return undefined;
    if (typeof source.formatCode !== "string" || source.formatCode.length === 0) return undefined;
    const out: ChartAxisNumberFormat = { formatCode: source.formatCode };
    if (source.sourceLinked === true) out.sourceLinked = true;
    return out;
  }
  if (override === null) return undefined;
  if (typeof override.formatCode !== "string" || override.formatCode.length === 0) return undefined;
  const out: ChartAxisNumberFormat = { formatCode: override.formatCode };
  if (override.sourceLinked === true) out.sourceLinked = true;
  return out;
}

/** Recognized values of `<c:majorTickMark>` / `<c:minorTickMark>`. */
const VALID_TICK_MARK_VALUES: ReadonlySet<ChartAxisTickMark> = new Set([
  "none",
  "in",
  "out",
  "cross",
]);

/**
 * Resolve a tick-mark override using the same `undefined` (inherit) /
 * `null` (drop) / value (replace) grammar as the other axis helpers.
 * Unknown / typo'd inputs collapse to `undefined` so the writer never
 * emits a value the OOXML `ST_TickMark` enum rejects.
 */
function applyTickMarkOverride(
  source: ChartAxisTickMark | undefined,
  override: ChartAxisTickMark | null | undefined,
): ChartAxisTickMark | undefined {
  if (override === undefined) {
    if (source === undefined) return undefined;
    return VALID_TICK_MARK_VALUES.has(source) ? source : undefined;
  }
  if (override === null) return undefined;
  return VALID_TICK_MARK_VALUES.has(override) ? override : undefined;
}

/** Recognized values of `<c:tickLblPos>`. */
const VALID_TICK_LBL_POS_VALUES: ReadonlySet<ChartAxisTickLabelPosition> = new Set([
  "nextTo",
  "low",
  "high",
  "none",
]);

/**
 * Resolve a tick-label-position override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Unknown / typo'd inputs collapse to `undefined` so
 * the writer never emits a value the OOXML `ST_TickLblPos` enum
 * rejects.
 */
function applyTickLblPosOverride(
  source: ChartAxisTickLabelPosition | undefined,
  override: ChartAxisTickLabelPosition | null | undefined,
): ChartAxisTickLabelPosition | undefined {
  if (override === undefined) {
    if (source === undefined) return undefined;
    return VALID_TICK_LBL_POS_VALUES.has(source) ? source : undefined;
  }
  if (override === null) return undefined;
  return VALID_TICK_LBL_POS_VALUES.has(override) ? override : undefined;
}

/**
 * Resolve a `labelRotation` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Out-of-range / non-numeric values clamp to the
 * `-90..90` band the writer accepts and the OOXML default `0`
 * collapses to `undefined` so absence and the default round-trip
 * identically — symmetric with the parser-side default-collapse.
 */
function applyLabelRotationOverride(
  source: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) {
    if (typeof source !== "number" || !Number.isFinite(source)) return undefined;
    return clampLabelRotationDeg(source);
  }
  if (override === null) return undefined;
  if (typeof override !== "number" || !Number.isFinite(override)) return undefined;
  return clampLabelRotationDeg(override);
}

/**
 * Snap a `labelRotation` value (whole degrees) into the `-90..90` band
 * the writer accepts. Returns `undefined` for `0` so the OOXML default
 * collapses to absence — symmetric with the writer-side
 * {@link normalizeAxisLabelRotation} contract.
 */
function clampLabelRotationDeg(value: number): number | undefined {
  let degrees = Math.round(value);
  if (degrees < -90) degrees = -90;
  else if (degrees > 90) degrees = 90;
  if (degrees === 0) return undefined;
  return degrees;
}

/**
 * Resolve a reverse-axis override using the same `undefined` (inherit) /
 * `null` (drop) / value (replace) grammar as the other axis helpers.
 *
 * Only `true` round-trips meaningfully — `false` is the OOXML default
 * (`orientation="minMax"`) so it collapses to `undefined` to keep the
 * cloned shape minimal. A source carrying `false` (e.g. an over-eager
 * parser that surfaced the default) collapses to `undefined` on
 * inherit; an explicit `false` override likewise drops the field. The
 * writer's per-axis `reverse: false` default already produces a forward
 * orientation, so the dropped state is indistinguishable from a literal
 * `false`.
 */
function applyReverseOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return source === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/** Recognized values of `<c:crosses>` per the OOXML `ST_Crosses` enum. */
const VALID_CROSSES_VALUES: ReadonlySet<ChartAxisCrosses> = new Set(["autoZero", "min", "max"]);

interface CrossesPairSource {
  crosses?: ChartAxisCrosses;
  crossesAt?: number;
}

interface CrossesPairOverride {
  crosses?: ChartAxisCrosses | null;
  crossesAt?: number | null;
}

interface CrossesPair {
  crosses?: ChartAxisCrosses;
  crossesAt?: number;
}

/**
 * Resolve the `crosses` / `crossesAt` pair using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers — but applied to the XSD choice between `<c:crosses>`
 * and `<c:crossesAt>`. The two fields are resolved independently
 * (each follows the standard inherit / null / replace contract); the
 * writer's normalizer enforces the choice rule downstream by
 * preferring the numeric pin when both are set.
 *
 * The OOXML default `crosses: "autoZero"` collapses to `undefined` so
 * the cloned shape stays minimal. `crossesAt: 0` is preserved (it is
 * a valid pin, distinct from the `"autoZero"` default). Non-finite
 * inputs and unknown semantic tokens drop to `undefined` so they
 * cannot leak into the writer.
 */
function applyCrossesOverride(
  source: CrossesPairSource,
  override: CrossesPairOverride,
): CrossesPair {
  const out: CrossesPair = {};

  if (override.crosses !== undefined) {
    if (override.crosses !== null) {
      const value = override.crosses;
      if (VALID_CROSSES_VALUES.has(value) && value !== "autoZero") {
        out.crosses = value;
      }
    }
    // override.crosses === null drops the field entirely.
  } else if (source.crosses !== undefined) {
    if (VALID_CROSSES_VALUES.has(source.crosses) && source.crosses !== "autoZero") {
      out.crosses = source.crosses;
    }
  }

  if (override.crossesAt !== undefined) {
    if (
      override.crossesAt !== null &&
      typeof override.crossesAt === "number" &&
      Number.isFinite(override.crossesAt)
    ) {
      out.crossesAt = override.crossesAt;
    }
    // override.crossesAt === null drops the field entirely.
  } else if (typeof source.crossesAt === "number" && Number.isFinite(source.crossesAt)) {
    out.crossesAt = source.crossesAt;
  }

  return out;
}

/** Recognized values of `<c:builtInUnit>` per the OOXML `ST_BuiltInUnit` enum. */
const VALID_DISP_UNIT_VALUES: ReadonlySet<ChartAxisDispUnit> = new Set([
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
 * Normalize a {@link ChartAxisDispUnit} shorthand or full
 * {@link ChartAxisDispUnits} object into a stable shape the resolver
 * can hand back to the writer-side `SheetChart.axes.{x,y}.dispUnits`
 * field. Unknown / typo'd tokens collapse to `undefined` so they cannot
 * leak past the clone layer.
 *
 * Both `unit` (built-in preset) and `custUnit` (custom numeric divisor)
 * pass through when valid. The OOXML schema's `xsd:choice` between
 * `<c:builtInUnit>` and `<c:custUnit>` is enforced at emit time by the
 * writer (which prefers `custUnit` when both are pinned); the
 * normalizer keeps both fields so the clone layer can append a
 * `custUnit` override to a source whose parsed value pinned `unit`
 * without manually pruning the inherited preset. Returns `undefined`
 * when neither field resolves to a valid value — a stray
 * `ChartAxisDispUnits` with no usable child has nothing to emit.
 */
function normalizeDispUnits(
  value: ChartAxisDispUnits | ChartAxisDispUnit | undefined,
): ChartAxisDispUnits | undefined {
  if (value === undefined) return undefined;
  if (typeof value === "string") {
    return VALID_DISP_UNIT_VALUES.has(value as ChartAxisDispUnit)
      ? { unit: value as ChartAxisDispUnit }
      : undefined;
  }
  if (typeof value !== "object" || value === null) return undefined;
  const out: ChartAxisDispUnits = {};
  const unit = value.unit;
  if (typeof unit === "string" && VALID_DISP_UNIT_VALUES.has(unit as ChartAxisDispUnit)) {
    out.unit = unit as ChartAxisDispUnit;
  }
  const custUnit = value.custUnit;
  if (typeof custUnit === "number" && Number.isFinite(custUnit) && custUnit > 0) {
    out.custUnit = custUnit;
  }
  if (out.unit === undefined && out.custUnit === undefined) return undefined;
  if (value.showLabel === true) out.showLabel = true;
  return out;
}

/**
 * Resolve a `dispUnits` override using the standard `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar. Both inputs go
 * through {@link normalizeDispUnits} so unknown tokens collapse to
 * `undefined` rather than fabricate a value the writer would never
 * emit. The reader and writer mirror this normalizer so a parsed
 * source value slots straight back into a clone target without
 * transformation.
 */
function applyDispUnitsOverride(
  source: ChartAxisDispUnits | undefined,
  override: ChartAxisDispUnits | ChartAxisDispUnit | null | undefined,
): ChartAxisDispUnits | undefined {
  if (override === undefined) return normalizeDispUnits(source);
  if (override === null) return undefined;
  return normalizeDispUnits(override);
}

/** Recognized values of `<c:crossBetween>` per the OOXML `ST_CrossBetween` enum. */
const VALID_CROSS_BETWEEN_VALUES: ReadonlySet<ChartAxisCrossBetween> = new Set([
  "between",
  "midCat",
]);

/**
 * Resolve a `crossBetween` override using the standard `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar. Unknown / typo'd
 * tokens collapse to `undefined` rather than fabricate a value the
 * writer would never emit — the writer's per-family default
 * (`"between"` on bar / column / line / area Y axes; `"midCat"` on
 * scatter axes) takes over instead. The reader and writer mirror this
 * normalizer so a parsed source value slots straight back into a clone
 * target without transformation.
 */
function applyCrossBetweenOverride(
  source: ChartAxisCrossBetween | undefined,
  override: ChartAxisCrossBetween | null | undefined,
): ChartAxisCrossBetween | undefined {
  if (override === undefined) {
    if (source === undefined) return undefined;
    return VALID_CROSS_BETWEEN_VALUES.has(source) ? source : undefined;
  }
  if (override === null) return undefined;
  return VALID_CROSS_BETWEEN_VALUES.has(override) ? override : undefined;
}

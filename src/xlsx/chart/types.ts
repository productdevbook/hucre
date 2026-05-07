// ── Chart Types ────────────────────────────────────────────────────
// Every chart-related interface, type alias, and enum that defines
// the chart read/write/clone surface. Extracted verbatim from
// `src/_types.ts` so that the chart submodules under `src/xlsx/chart/`
// have a single, focused place to import their public types from
// without dragging the rest of the workbook type universe along.
//
// `src/_types.ts` re-exports every type defined here so existing
// consumers (`import { Chart, SheetChart, ... } from "hucre"`) keep
// working without change.

// ── Charts ─────────────────────────────────────────────────────────

/**
 * Chart kinds supported by {@link writeXlsx} when authoring charts via
 * {@link WriteSheet.charts}. Covers the most common chart families —
 * bar/column, line, pie, doughnut, scatter, and area.
 *
 * Distinct from the read-side {@link ChartKind} (which mirrors the
 * full set of OOXML chart-type element local names) — the write side
 * exposes only the kinds the chart author can emit today.
 */
export type WriteChartKind = "bar" | "column" | "line" | "pie" | "doughnut" | "scatter" | "area";

/**
 * Where a data label is placed relative to its data point.
 *
 * Mirrors the OOXML `c:dLblPos` value space. Not every chart kind
 * accepts every position — Excel will silently fall back to a sensible
 * default when an invalid combination is requested:
 *
 * - **Bar / column**: `"ctr"`, `"inEnd"`, `"inBase"`, `"outEnd"` (clustered) or `"ctr"`, `"inEnd"`, `"inBase"` (stacked).
 * - **Line / area / scatter**: `"t"`, `"b"`, `"l"`, `"r"`, `"ctr"`.
 * - **Pie / doughnut**: `"ctr"`, `"inEnd"`, `"outEnd"`, `"bestFit"`.
 */
export type ChartDataLabelPosition =
  | "t"
  | "b"
  | "l"
  | "r"
  | "ctr"
  | "inEnd"
  | "inBase"
  | "outEnd"
  | "bestFit";

/**
 * Configuration for the small text annotations Excel paints next to
 * each data point. Maps to the OOXML `<c:dLbls>` element.
 *
 * Apply at the chart level via {@link SheetChart.dataLabels} to label
 * every series, or at the series level via
 * {@link ChartSeries.dataLabels} to override a single series. A
 * series-level `dataLabels` always wins over the chart-level default,
 * including when the value is `false` (which suppresses the labels for
 * that series alone).
 *
 * At least one of `showValue`, `showCategoryName`, `showSeriesName`,
 * or `showPercent` should be `true` for the labels to render anything
 * meaningful — Excel hides the label box when no toggle is on.
 */
export interface ChartDataLabels {
  /** Show the numeric value of each data point. */
  showValue?: boolean;
  /** Show the category (X-axis) label. */
  showCategoryName?: boolean;
  /** Show the series name. Useful with multi-series legends collapsed. */
  showSeriesName?: boolean;
  /** Show the value as a percent of total. Pie / doughnut only. */
  showPercent?: boolean;
  /**
   * Render the legend's color swatch (the small marker / bar Excel
   * paints in the chart legend) inline with each data label. Mirrors
   * Excel's "Format Data Labels -> Legend Key" checkbox.
   *
   * Maps to `<c:showLegendKey val=".."/>` inside `<c:dLbls>`. The OOXML
   * default is `false` (no legend key); set to `true` to repeat the
   * legend swatch alongside every label.
   */
  showLegendKey?: boolean;
  /**
   * Where the label sits relative to its point. See
   * {@link ChartDataLabelPosition} for the valid set per chart kind.
   * Omit to let Excel pick a default (`outEnd` for bar/column,
   * `r` for line/scatter, `bestFit` for pie).
   */
  position?: ChartDataLabelPosition;
  /**
   * Separator between concatenated label parts when more than one
   * `show*` toggle is on. Defaults to `", "`. Common alternatives:
   * `" "`, `"; "`, `"\n"` (newline).
   */
  separator?: string;
  /**
   * Number format applied to the value rendered inside each data label.
   * Mirrors Excel's "Format Data Labels -> Number" panel — pinning a
   * `formatCode` such as `"0.00%"`, `"$#,##0.00"`, or `"#,##0"` lets a
   * dashboard chart show currency / percent labels without the source
   * cells carrying the format. Same shape as
   * {@link ChartAxisNumberFormat} so a single number-format helper can
   * thread through both the axis and the data-labels code paths.
   *
   * Maps to `<c:numFmt formatCode=".." sourceLinked=".."/>` inside
   * `<c:dLbls>`. The element sits right after `<c:dLbl>` (per the
   * `CT_DLbls` schema sequence — `dLbl* -> numFmt? -> spPr? -> txPr? ->
   * dLblPos? -> show*`). Omit to fall back to whatever Excel inherits
   * from the source cell formatting.
   */
  numberFormat?: ChartAxisNumberFormat;
  /**
   * Render the leader lines that connect each data label back to its
   * pie / doughnut slice when Excel pushes a label outside the slice it
   * belongs to. Mirrors Excel's "Format Data Labels -> Label Options ->
   * Show Leader Lines" checkbox.
   *
   * Maps to `<c:showLeaderLines val=".."/>` inside `<c:dLbls>`, which
   * sits at the tail of the `EG_DLbls` group (after `<c:separator>`,
   * before `<c:extLst>`). The OOXML schema scopes the element to
   * pie / doughnut chart families exclusively (`EG_DLbls` for
   * `CT_PieChart` / `CT_DoughnutChart` only — bar / column / line /
   * area / scatter route through `EG_DLblsShared` which omits it). The
   * writer drops the field silently on every non-pie / non-doughnut
   * family to mirror Excel's reference serialization.
   *
   * The OOXML default is `true` (Excel paints leader lines on every
   * label that gets pushed outside its slice). Set to `false` to opt
   * out and render labels without the connecting lines.
   */
  showLeaderLines?: boolean;
  /**
   * Data-label font size in points (range `1..400`), pinned via
   * `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
   * </c:txPr></c:dLbls>`. Mirrors Excel's "Format Data Labels -> Font
   * -> Size" knob — the OOXML attribute is in 100ths of a point on
   * `CT_TextCharacterProperties`' `sz` slot (ECMA-376 Part 1,
   * §21.1.2.3.7); the writer holds the value in points and converts
   * at emit time so a caller can pass the value directly without
   * doing the multiplication.
   *
   * Default: omitted — the data labels render at the theme-default
   * size (Excel's reference behavior for fresh data labels whose
   * typography has not been customized; the writer skips the entire
   * `<c:txPr>` block when no font knob is pinned).
   *
   * Out-of-range values (`< 1` or `> 400`) and non-numeric tokens
   * (typed escapes from an untyped caller — strings, `null`, `NaN`,
   * `Infinity`) collapse to `undefined` so the writer never emits a
   * malformed `sz` attribute. Fractional points round to the nearest
   * half-point (Excel's UI step) — `12.3` → `12.5`, `12.24` → `12`.
   *
   * The `<c:txPr>` block lands between `<c:spPr>` and `<c:dLblPos>`,
   * matching the CT_DLbls schema sequence (ECMA-376 Part 1,
   * §21.2.2.50). Composes independently with the other dLbls knobs —
   * {@link position} / {@link separator} / {@link numberFormat} /
   * {@link showLeaderLines} / the `show*` toggles — so a caller can
   * pin the font size without touching the rest of the configuration.
   * Mirrors {@link SheetChart.titleFontSize} /
   * {@link SheetChart.legendFontSize} — same range, same conversion
   * factor — so a caller can thread a single point value through every
   * typography-pinning slot.
   */
  fontSize?: number;
  /**
   * Data-label font color. Maps to `<c:dLbls><c:txPr><a:p><a:pPr>
   * <a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
   * </a:defRPr></a:pPr></a:p></c:txPr></c:dLbls>` — Excel's "Format
   * Data Labels -> Font -> Font color" picker. The OOXML
   * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
   * sRGB color (`CT_SRgbColor` inside `CT_TextCharacterProperties`'
   * fill choice — ECMA-376 Part 1, §20.1.2.3.32 / §21.1.2.3.7); the
   * writer lands the value on the default-paragraph `<a:defRPr>` slot
   * inside the `<c:dLbls><c:txPr>` block so a re-parse picks the
   * color up off the canonical slot the OOXML schema exposes.
   *
   * Accepts the color either with or without a leading `#` and in any
   * case — `"FF0000"`, `"#FF0000"`, and `"ff0000"` all collapse to
   * the OOXML uppercase canonical form `"FF0000"`. Malformed inputs
   * (wrong length, non-hex characters, alpha-channel forms like
   * `"#FF0000FF"`, non-string escapes from an untyped caller)
   * collapse to `undefined` so the writer skips the entire
   * `<a:solidFill>` block and the data labels inherit the theme text
   * color (Excel's reference behavior for fresh data labels that
   * have not had a custom color picked).
   *
   * Default: omitted — the data labels render at the theme text color
   * (no `<a:solidFill>` block, matching Excel's reference
   * serialization for fresh data labels whose typography has not been
   * customized).
   *
   * The `<c:txPr>` block lands between `<c:numFmt>` and `<c:dLblPos>`
   * (CT_DLbls schema, ECMA-376 Part 1, §21.2.2.50). Mirrors the
   * chart-title `titleColor` / axis-title `axisTitleColor` / axis
   * tick-label `labelColor` / legend `legendFontColor` knobs — same
   * accept-with-or-without-`#` hex grammar, same OOXML
   * `<a:solidFill><a:srgbClr val=".."/>` mapping — so a caller can
   * thread a single hex string through every typography-pinning slot.
   */
  fontColor?: string;
  /**
   * Data-label bold flag. Maps to `<c:dLbls><c:txPr><a:p><a:pPr>
   * <a:defRPr b=".."/></a:pPr></a:p></c:txPr></c:dLbls>` — Excel's
   * "Format Data Labels -> Font -> Bold" toggle. The OOXML attribute
   * is the `xsd:boolean` bold flag on `CT_TextCharacterProperties`
   * (ECMA-376 Part 1, §21.1.2.3.7); the writer lands `b="1"` (bold)
   * or `b="0"` (the OOXML default — non-bold) on the
   * default-paragraph `<a:defRPr>` slot inside the data-label's
   * `<c:txPr>` block so a re-parse picks the flag up off the
   * canonical slot the OOXML schema exposes.
   *
   * Default: omitted — the data labels render non-bold (no `b`
   * attribute, matching Excel's reference serialization for fresh
   * data labels whose typography has not been customized; the OOXML
   * default `0` collapses to absence). Set `true` to emit `b="1"` so
   * the labels render bold; set `false` explicitly to pin `b="0"`
   * (functionally identical to omission, but useful when overriding
   * a templated chart that had bold pinned upstream).
   *
   * The `<c:txPr>` block lands between `<c:numFmt>` and `<c:dLblPos>`
   * (CT_DLbls schema, ECMA-376 Part 1, §21.2.2.50). Composes
   * independently with the other dLbls knobs — {@link position} /
   * {@link separator} / {@link numberFormat} / {@link showLeaderLines}
   * / the `show*` toggles. Mirrors the chart-title `titleBold` /
   * axis-title `axisTitleBold` / axis tick-label `labelBold` / legend
   * `legendBold` knobs — same boolean shape, same OOXML
   * `<a:defRPr b=".."/>` mapping — so a caller can thread a single
   * bold value through every typography-pinning slot.
   */
  bold?: boolean;
  /**
   * Data-label italic flag. Maps to `<c:dLbls><c:txPr><a:p><a:pPr>
   * <a:defRPr i=".."/></a:pPr></a:p></c:txPr></c:dLbls>` — Excel's
   * "Format Data Labels -> Font -> Italic" toggle. The OOXML attribute
   * is the `xsd:boolean` italic flag on `CT_TextCharacterProperties`
   * (ECMA-376 Part 1, §21.1.2.3.7); the writer lands `i="1"` (italic)
   * or `i="0"` (the OOXML default — upright) on the default-paragraph
   * `<a:defRPr>` slot inside the data-label's `<c:txPr>` block so a
   * re-parse picks the flag up off the canonical slot the OOXML
   * schema exposes.
   *
   * Default: omitted — the data labels render upright (no `i`
   * attribute, matching Excel's reference serialization for fresh
   * data labels whose typography has not been customized; the OOXML
   * default `0` collapses to absence). Set `true` to emit `i="1"` so
   * the labels render italic; set `false` explicitly to pin `i="0"`
   * (functionally identical to omission, but useful when overriding
   * a templated chart that had italic pinned upstream).
   *
   * The `<c:txPr>` block lands between `<c:numFmt>` and `<c:dLblPos>`
   * (CT_DLbls schema, ECMA-376 Part 1, §21.2.2.50). Composes
   * independently with the other dLbls knobs — {@link position} /
   * {@link separator} / {@link numberFormat} / {@link showLeaderLines}
   * / the `show*` toggles / {@link bold}. Mirrors the chart-title
   * `titleItalic` / axis-title `axisTitleItalic` / axis tick-label
   * `labelItalic` / legend `legendItalic` knobs — same boolean shape,
   * same OOXML `<a:defRPr i=".."/>` mapping — so a caller can thread
   * a single italic value through every typography-pinning slot.
   */
  italic?: boolean;
  /**
   * Data-label underline flag. Maps to `<c:dLbls><c:txPr><a:p><a:pPr>
   * <a:defRPr u=".."/></a:pPr></a:p></c:txPr></c:dLbls>` — Excel's
   * "Format Data Labels -> Font -> Underline" toggle. The OOXML
   * attribute is the `ST_TextUnderlineType` enumeration on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7); the
   * writer lands `u="sng"` (single underline — Excel's UI checkbox)
   * or `u="none"` (the OOXML default — no underline) on the
   * default-paragraph `<a:defRPr>` slot inside the data-label's
   * `<c:txPr>` block so a re-parse picks the flag up off the
   * canonical slot the OOXML schema exposes.
   *
   * Modeled as a boolean for symmetry with {@link bold} / {@link italic}
   * / the other `*Underline` knobs across the chart-title, axis-title,
   * axis tick-label, and legend slots: `true` emits `u="sng"`; `false`
   * emits `u="none"` explicitly so a clone target can override an
   * upstream `u="sng"` from a templated chart; absence collapses to
   * omitting the attribute entirely. The non-UI variant `"dbl"`
   * (double line) and the sixteen exotic `ST_TextUnderlineType`
   * tokens (`"words"`, `"heavy"`, `"dotted"`, etc.) are read-only —
   * the writer emits only `"sng"` / `"none"` to keep the surfaced
   * shape consistent with what Excel's reference UI authors.
   *
   * The `<c:txPr>` block lands between `<c:numFmt>` and `<c:dLblPos>`
   * (CT_DLbls schema, ECMA-376 Part 1, §21.2.2.50). Composes
   * independently with the other dLbls knobs — {@link position} /
   * {@link separator} / {@link numberFormat} / {@link showLeaderLines}
   * / the `show*` toggles / {@link bold} / {@link italic}. Mirrors the
   * chart-title `titleUnderline` / axis-title `axisTitleUnderline` /
   * axis tick-label `labelUnderline` / legend `legendUnderline` knobs
   * — same boolean shape, same OOXML `<a:defRPr u=".."/>` mapping —
   * so a caller can thread a single underline value through every
   * typography-pinning slot.
   */
  underline?: boolean;
  /**
   * Data-label strikethrough flag. Maps to `<c:dLbls><c:txPr><a:p>
   * <a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr></c:dLbls>`
   * — Excel's "Format Data Labels -> Font -> Strikethrough" toggle.
   * The OOXML attribute is the `ST_TextStrikeType` enumeration on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7); the
   * writer lands `strike="sngStrike"` (single line — Excel's UI
   * checkbox) on the default-paragraph `<a:defRPr>` slot inside the
   * data-label's `<c:txPr>` block so a re-parse picks the flag up
   * off the canonical slot the OOXML schema exposes.
   *
   * Modeled as a boolean for symmetry with the data-labels {@link bold}
   * / {@link italic} / {@link underline} and the chart-title
   * `titleStrikethrough` / axis-title `axisTitleStrike` / axis
   * tick-label `labelStrikethrough` / legend `legendStrikethrough`
   * knobs: `true` emits `strike="sngStrike"`; absence and explicit
   * `false` both collapse to omitting the attribute (the OOXML default
   * `"noStrike"` is functionally identical to absence — the writer
   * never emits `"noStrike"` or `"dblStrike"` to keep the surfaced
   * shape consistent with what Excel's UI authors). The non-UI variant
   * `"dblStrike"` (double line) is read-only — the reader collapses
   * every non-`"sngStrike"` token to `undefined` so a templated chart
   * that pinned the double-line variant in raw OOXML round-trips
   * lossless rather than silently downgrading on re-emit.
   *
   * The `<c:txPr>` block lands between `<c:numFmt>` and `<c:dLblPos>`
   * (CT_DLbls schema, ECMA-376 Part 1, §21.2.2.50). Composes
   * independently with the other dLbls knobs — {@link position} /
   * {@link separator} / {@link numberFormat} / {@link showLeaderLines}
   * / the `show*` toggles / {@link bold} / {@link italic} /
   * {@link underline}.
   */
  strikethrough?: boolean;
  /**
   * Data-label font family / typeface. Maps to `<c:dLbls><c:txPr>
   * <a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
   * </a:pPr></a:p></c:txPr></c:dLbls>` — Excel's "Format Data Labels
   * -> Font -> Font" picker. The OOXML `<a:latin typeface=".."/>`
   * element carries the typeface name (`CT_TextFont`, ECMA-376 Part 1,
   * §21.1.2.3.7); the writer lands the element on the default-
   * paragraph `<a:defRPr>` slot inside the data-label's `<c:txPr>`
   * block so a re-parse picks the typeface up off the canonical slot
   * the OOXML schema exposes.
   *
   * Accepts any non-empty string typeface name (e.g. `"Calibri"`,
   * `"Arial"`, `"Times New Roman"`); the writer trims surrounding
   * whitespace and emits the trimmed value verbatim (XML-escaped) so
   * Excel can resolve the named font from the workbook's font scheme
   * or the host system's installed fonts. Empty / whitespace-only
   * strings and non-string tokens collapse to `undefined` so the
   * writer skips the entire `<a:latin>` element and the data labels
   * inherit Excel's reference theme typeface.
   *
   * Default: omitted — the data labels render in Excel's reference
   * theme typeface (no `<a:latin>` element, the writer skips the
   * element entirely). Pin a typeface name to render the labels in
   * that font.
   *
   * The `<c:txPr>` block lands between `<c:numFmt>` and `<c:dLblPos>`
   * (CT_DLbls schema, ECMA-376 Part 1, §21.2.2.50). Mirrors
   * {@link SheetChart.titleFontFamily} /
   * {@link SheetChart.legendFontFamily} /
   * {@link SheetChart.axes.x.axisTitleFontFamily} /
   * {@link SheetChart.axes.x.labelFontFamily} — same accept-and-trim
   * grammar, same OOXML `<a:latin typeface=".."/>` mapping — so a
   * caller can thread a single typeface string through every
   * typography-pinning slot. Composes independently with the other
   * dLbls knobs — {@link position} / {@link separator} /
   * {@link numberFormat} / {@link showLeaderLines} / the `show*`
   * toggles / {@link bold} / {@link italic} / {@link underline} /
   * {@link strikethrough} / {@link fontSize} / {@link fontColor}.
   */
  fontFamily?: string;
  /**
   * Data-labels background fill (solid). Maps to `<c:dLbls><c:spPr>
   * <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></c:spPr>
   * </c:dLbls>` — Excel's "Format Data Labels -> Fill -> Solid fill ->
   * Color" picker. The OOXML `<c:spPr>` block sits on `CT_DLbls` between
   * `<c:numFmt>` and `<c:txPr>` per the schema sequence (ECMA-376
   * Part 1, §21.2.2.50); the `<a:srgbClr val=".."/>` carries the
   * 6-character uppercase hex sRGB color (`CT_SRgbColor` inside
   * `CT_ShapeProperties`' fill choice — ECMA-376 Part 1, §20.1.2.3.32 /
   * §20.1.8.54).
   *
   * Accepts a 6-character hex string with or without a leading `#`,
   * any case (`"FF0000"`, `"#1070ca"`, `"abcdef"`); the writer
   * normalizes to the OOXML canonical 6-character uppercase form
   * (`"FF0000"`, `"1070CA"`, `"ABCDEF"`) so a re-parse round-trips
   * losslessly. Malformed inputs (wrong length, non-hex characters,
   * alpha-channel forms like `"FFAA0080"`, empty / whitespace-only
   * strings, non-string escapes from an untyped caller) collapse to
   * `undefined` and the writer skips the entire `<c:spPr>` block —
   * the data labels render with no background fill (Excel's reference
   * shape for fresh data labels whose fill has not been pinned).
   *
   * Default: omitted — the data labels render with no `<c:spPr>` block
   * (Excel's reference serialization for fresh data labels whose fill
   * has not been customized — typically a transparent label background).
   *
   * Distinct from {@link fontColor} (the typography color living on
   * `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>`); the two knobs target
   * different children of `<c:dLbls>` so a caller can pin both without
   * conflict — {@link fillColor} paints the label box background while
   * {@link fontColor} paints the text glyphs. Mirrors the chart-title
   * `titleFillColor` / axis-title `axisTitleFillColor` / plot-area
   * `plotAreaFillColor` / legend `legendFillColor` knobs — same
   * accept-with-or-without-`#` hex grammar, same OOXML
   * `<c:spPr><a:solidFill><a:srgbClr val=".."/>` mapping — so a caller
   * can thread a single hex string through every fill-pinning slot.
   * Composes independently with the other dLbls knobs — {@link position}
   * / {@link separator} / {@link numberFormat} / {@link showLeaderLines}
   * / the `show*` toggles / {@link bold} / {@link italic} /
   * {@link underline} / {@link strikethrough} / {@link fontSize} /
   * {@link fontColor} / {@link fontFamily}.
   */
  fillColor?: string;
  /**
   * Data-labels border (line) color as a 6-digit RGB hex string (e.g.
   * `"1F77B4"`). Maps to `<c:dLbls><c:spPr><a:ln><a:solidFill>
   * <a:srgbClr val="RRGGBB"/></a:solidFill></a:ln></c:spPr></c:dLbls>`
   * (CT_DLbls, ECMA-376 Part 1, §21.2.2.50) — Excel's "Format Data
   * Labels -> Border -> Solid line -> Color" picker. The OOXML
   * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
   * sRGB color (`CT_SRgbColor` inside `<a:ln>`'s solid fill choice —
   * ECMA-376 Part 1, §20.1.2.3.32 / §20.1.2.3.24). The `<a:ln>` child
   * sits inside the `<c:spPr>` block alongside the optional
   * `<a:solidFill>` fill child, in `CT_ShapeProperties` schema order
   * (fill before stroke).
   *
   * Distinct from {@link fillColor} — the border color paints the
   * outline around each data label box, while the fill color paints
   * the label box background. The two knobs share the `<c:spPr>` host
   * but land on different children (`<a:solidFill>` for the fill,
   * `<a:ln>` for the stroke), and the writer authors a single
   * `<c:spPr>` whenever either knob is set. A caller can pin one
   * without the other; pinning both produces a filled label box with
   * a colored border.
   *
   * Distinct from {@link fontColor} (the typography color living on
   * `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>`); the three knobs
   * target different children of `<c:dLbls>` so a caller can pin all
   * of them without conflict — {@link fillColor} paints the label box
   * background, {@link borderColor} paints its outline, and
   * {@link fontColor} paints the text glyphs.
   *
   * Accepts a leading `#` and any case; the writer collapses to the
   * OOXML canonical uppercase form. Malformed inputs (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` and the writer omits
   * the `<a:ln>` block (Excel's reference serialization for data
   * labels that inherit the auto-stroke — typically no border).
   *
   * Default: omitted — the data labels render with no `<c:spPr>`
   * border (Excel's reference serialization for fresh data labels
   * whose border has not been pinned). Pin a hex color to mirror
   * Excel's "Format Data Labels -> Border -> Solid line" knob and
   * paint a flat outline around each label — useful for dashboard
   * tiles where the data labels need to stand out against a busy
   * plot-area background.
   *
   * Patterned / gradient strokes are not modelled — only the solid
   * sRGB form lands on the wire. Theme-color references
   * (`<a:schemeClr>`) likewise drop to `undefined` so a parsed value
   * always carries a literal hex Excel will render byte-for-byte.
   * Mirrors `plotAreaBorderColor` / `legendBorderColor` /
   * `titleBorderColor` / `axisTitleBorderColor` /
   * `dataTableBorderColor` — same accept-with-or-without-`#` hex
   * grammar, same OOXML `<a:ln><a:solidFill><a:srgbClr val=".."/>
   * </a:solidFill></a:ln>` mapping — so a caller can thread a single
   * hex string through every `<a:ln>`-based stroke slot. Composes
   * independently with the other dLbls knobs — {@link position} /
   * {@link separator} / {@link numberFormat} / {@link showLeaderLines}
   * / the `show*` toggles / {@link bold} / {@link italic} /
   * {@link underline} / {@link strikethrough} / {@link fontSize} /
   * {@link fontColor} / {@link fontFamily} / {@link fillColor}.
   */
  borderColor?: string;
  /**
   * Data-labels border (stroke) thickness in points (e.g. `1.5`). Maps
   * to the `w` attribute on `<c:dLbls><c:spPr><a:ln w="EMU">` — Excel's
   * "Format Data Labels -> Border -> Width" spinner. The OOXML `w`
   * attribute carries the stroke width in English Metric Units
   * (1 pt = 12 700 EMU) per `CT_LineProperties` (ECMA-376 Part 1,
   * §20.1.2.3.24). The writer multiplies by 12 700 and rounds to the
   * nearest integer because the schema types `w` as `xsd:int`.
   *
   * Accepts any finite number; values are clamped to the `0.25..13.5`
   * pt band Excel's UI exposes (the same band used by every other
   * chart-frame border-width knob) and snapped to the 0.25 pt grid so
   * a parsed-then-written width does not drift across round-trips.
   * Non-finite / non-numeric / `NaN` values collapse to `undefined`
   * and the writer omits the `w` attribute (the line keeps Excel's
   * auto-thickness — typically 0.75 pt).
   *
   * Default: omitted — the data-label border inherits Excel's
   * auto-thickness.
   *
   * Composes independently with {@link borderColor} — the width
   * attribute lands on the same `<a:ln>` element as the color's
   * `<a:solidFill>` child, but the writer authors `<a:ln>` whenever
   * either knob is set.
   */
  borderWidth?: number;
  /**
   * Data-labels border (stroke) preset dash pattern. Maps to the `val`
   * attribute on `<c:dLbls><c:spPr><a:ln><a:prstDash val=".."/>`.
   * Mirrors Excel's "Format Data Labels -> Border -> Dash type"
   * picker. Same {@link ChartBorderDash} accept-or-drop grammar as the
   * chart-level {@link SheetChart.plotAreaBorderDash} — `"solid"`
   * collapses to `undefined` so absence and the OOXML default
   * round-trip identically.
   *
   * Composes independently with {@link borderColor} and
   * {@link borderWidth} — all three knobs share the same `<a:ln>`
   * element.
   */
  borderDash?: ChartBorderDash;
}

/**
 * Preset dash pattern for a chart series line stroke.
 *
 * Mirrors the OOXML `ST_PresetLineDashVal` enum exactly. Each value
 * names a stock pattern Excel paints without needing a custom dash
 * array. The Excel "Format Data Series → Line → Dash type" UI exposes
 * these stock patterns; Excel ignores any unrecognized value.
 */
export type ChartLineDashStyle =
  | "solid"
  | "dot"
  | "dash"
  | "lgDash"
  | "dashDot"
  | "lgDashDot"
  | "lgDashDotDot"
  | "sysDash"
  | "sysDot"
  | "sysDashDot"
  | "sysDashDotDot";

/**
 * Preset dash pattern for a chart-frame border stroke (plot-area /
 * legend / title / chart-space / axis-title / data-table / data-label
 * borders). Mirrors the OOXML `ST_PresetLineDashVal` enum exactly —
 * the same set of preset patterns the per-series
 * {@link ChartLineDashStyle} carries — but lands on `<c:spPr><a:ln>
 * <a:prstDash val="..">` for chart-frame borders.
 *
 * The OOXML default is `"solid"` (no `<a:prstDash>` child); the writer
 * omits the element when `"solid"` (or `undefined`) is passed so a
 * fresh chart matches Excel's reference shape byte-for-byte. Excel
 * ignores any unrecognized value.
 */
export type ChartBorderDash =
  | "solid"
  | "dash"
  | "dashDot"
  | "dot"
  | "lgDash"
  | "lgDashDot"
  | "lgDashDotDot"
  | "sysDash"
  | "sysDashDot"
  | "sysDashDotDot"
  | "sysDot";

/**
 * Per-series line stroke styling for line / scatter charts.
 *
 * Maps to the `<a:ln>` element nested inside `<c:ser><c:spPr>` — the
 * same wrapper that already carries the series fill color. Only
 * meaningful on `line` and `scatter` series; the field is silently
 * dropped on every other chart family at all three layers (read,
 * write, clone), since dashing and stroke width have no visible effect
 * on bar / pie / doughnut / area renderings.
 *
 * Every field is optional — a bare `{}` collapses to no stroke
 * configuration and leaves Excel's per-series default in place. Set
 * `dash: "solid"` to explicitly reset a template's dashed stroke back
 * to a continuous line.
 */
export interface ChartLineStroke {
  /**
   * Preset dash pattern. See {@link ChartLineDashStyle} for the
   * accepted set.
   */
  dash?: ChartLineDashStyle;
  /**
   * Stroke width in points. Excel's UI exposes the 0.25 – 13.5 pt band;
   * the writer clamps anything outside that range and rounds to the
   * nearest quarter-point so a round-trip cannot drift. The OOXML
   * attribute is in EMU (1 pt = 12 700 EMU); the writer performs the
   * conversion and the reader inverts it. Non-finite values are
   * dropped so the writer can elide the attribute entirely.
   */
  width?: number;
}

/**
 * Marker symbol shape rendered at each data point on a line / scatter
 * series.
 *
 * Mirrors the OOXML `ST_MarkerStyle` enum exactly. `"none"` suppresses
 * the marker (the Excel default for line charts beyond the first
 * series); `"auto"` defers to Excel's series-rotation default; every
 * other value pins a specific shape. `"picture"` is intentionally
 * omitted — it requires a separately-embedded picture part that Phase 1
 * native chart authoring does not support.
 */
export type ChartMarkerSymbol =
  | "none"
  | "auto"
  | "circle"
  | "square"
  | "diamond"
  | "triangle"
  | "x"
  | "star"
  | "dot"
  | "dash"
  | "plus";

/**
 * Per-series marker styling for line / scatter charts.
 *
 * Maps to `<c:marker>` inside `<c:ser>`. Only meaningful on `line` and
 * `scatter` series — the OOXML schema places `<c:marker>` exclusively
 * on `CT_LineSer` and `CT_ScatterSer`, so the field is silently
 * dropped on every other chart family at all three layers (read,
 * write, clone).
 *
 * Every field is optional — a bare `{}` collapses to no marker
 * configuration and leaves Excel's per-series default in place. Set
 * `symbol: "none"` to explicitly hide the marker (useful for a
 * scatter clone whose template uses markers but the dashboard wants
 * a clean line).
 */
export interface ChartMarker {
  /** Shape of the marker glyph. See {@link ChartMarkerSymbol}. */
  symbol?: ChartMarkerSymbol;
  /**
   * Marker glyph size in points, in the OOXML range `2..72`. Excel's
   * UI clamps values outside this band. Default (when omitted): Excel
   * picks a series-rotation default (typically `5`).
   */
  size?: number;
  /**
   * Marker fill color as a 6-digit RGB hex string (e.g. `"1F77B4"`).
   * Maps to `<c:marker><c:spPr><a:solidFill><a:srgbClr val="..">`.
   */
  fill?: string;
  /**
   * Marker outline color as a 6-digit RGB hex string. Maps to
   * `<c:marker><c:spPr><a:ln><a:solidFill><a:srgbClr val="..">`.
   */
  line?: string;
}

/**
 * How Excel paints a series across cells whose value is missing or
 * blank. Mirrors the OOXML `ST_DispBlanksAs` enum exactly and matches
 * the three options Excel exposes under "Select Data Source → Hidden
 * and Empty Cells":
 *
 * - `"gap"` — leave a gap at the missing point (the OOXML default and
 *   what Excel selects in fresh chart UI). A line chart shows a break,
 *   a bar chart simply skips the bar.
 * - `"zero"` — substitute `0` for the missing value, so a line chart
 *   drops to the X axis and bar charts render a flush-zero bar.
 * - `"span"` — connect adjacent points across the gap (line / scatter
 *   only; Excel falls back to `"gap"` for bar / pie / area).
 */
export type ChartDisplayBlanksAs = "gap" | "zero" | "span";

/**
 * Granular data-table configuration for {@link SheetChart.dataTable}.
 * Maps to the four boolean children of `<c:plotArea><c:dTable>`:
 * `<c:showHorzBorder>`, `<c:showVertBorder>`, `<c:showOutline>`, and
 * `<c:showKeys>`. Each toggle flips one of Excel's "Format Data Table"
 * checkboxes:
 *
 * - {@link showHorzBorder} — paint the horizontal lines between table
 *   rows. Default: `true` (Excel's reference serialization).
 * - {@link showVertBorder} — paint the vertical lines between category
 *   columns. Default: `true`.
 * - {@link showOutline}    — paint the outer border around the table.
 *   Default: `true`.
 * - {@link showKeys}       — render the legend swatch next to each
 *   series row so the table doubles as the chart legend. Default:
 *   `true`.
 *
 * The writer always emits all four children — the OOXML schema marks
 * them required on `CT_DTable` — falling back to the per-field defaults
 * for any field the caller leaves unset. Pass an empty object (`{}`) to
 * accept every default, equivalent to `dataTable: true`.
 */
export interface ChartDataTable {
  /** Paint horizontal lines between rows. Default: `true`. */
  showHorzBorder?: boolean;
  /** Paint vertical lines between category columns. Default: `true`. */
  showVertBorder?: boolean;
  /** Paint the outer border around the table. Default: `true`. */
  showOutline?: boolean;
  /** Render the legend swatch next to each series row. Default: `true`. */
  showKeys?: boolean;
  /**
   * Data-table font size in points (range `1..400`), pinned via
   * `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
   * </c:txPr></c:dTable>`. Mirrors Excel's "Format Data Table -> Font
   * -> Size" knob — the OOXML attribute is in 100ths of a point on
   * `CT_TextCharacterProperties`' `sz` slot (ECMA-376 Part 1,
   * §21.1.2.3.7); the writer holds the value in points and converts at
   * emit time so a caller can pass the value directly without doing the
   * multiplication.
   *
   * Default: omitted — the data table renders at the theme-default
   * size (Excel's reference behavior for fresh data tables whose
   * typography has not been customized; the writer skips the entire
   * `<c:txPr>` block when no font knob is pinned).
   *
   * Out-of-range values (`< 1` or `> 400`) and non-numeric tokens
   * (typed escapes from an untyped caller — strings, `null`, `NaN`,
   * `Infinity`) collapse to `undefined` so the writer never emits a
   * malformed `sz` attribute. Fractional points round to the nearest
   * half-point (Excel's UI step) — `12.3` → `12.5`, `12.24` → `12`.
   *
   * The `<c:txPr>` block lands after the four required boolean children
   * (`<c:showHorzBorder>`, `<c:showVertBorder>`, `<c:showOutline>`,
   * `<c:showKeys>`) per the CT_DTable schema sequence (ECMA-376 Part 1,
   * §21.2.2.54). Composes independently with the four boolean toggles —
   * so a caller can pin the font size without touching the rest of the
   * configuration. Mirrors {@link SheetChart.titleFontSize} /
   * {@link SheetChart.legendFontSize} / {@link ChartDataLabels.fontSize}
   * — same range, same fractional rounding, same OOXML conversion
   * factor — so a caller can thread a single point value through every
   * typography-pinning slot.
   *
   * Only meaningful for chart families with axes (`bar`, `column`,
   * `line`, `area`, `scatter`) — the OOXML schema places `<c:dTable>`
   * inside `<c:plotArea>` after the axes, so pie / doughnut have no
   * slot for it (they have no axes at all). The writer silently drops
   * the field on those families along with the rest of the data-table
   * configuration.
   */
  fontSize?: number;
  /**
   * Data-table font color. Maps to `<c:dTable><c:txPr><a:p><a:pPr>
   * <a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
   * </a:defRPr></a:pPr></a:p></c:txPr></c:dTable>` — Excel's "Format
   * Data Table -> Font -> Font color" picker. The OOXML
   * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
   * sRGB color (`CT_SRgbColor` inside `CT_TextCharacterProperties`'
   * fill choice — ECMA-376 Part 1, §20.1.2.3.32 / §21.1.2.3.7); the
   * writer lands the value on the default-paragraph `<a:defRPr>` slot
   * inside the `<c:dTable><c:txPr>` block so a re-parse picks the
   * color up off the canonical slot the OOXML schema exposes.
   *
   * Accepts the color either with or without a leading `#` and in any
   * case — `"FF0000"`, `"#FF0000"`, and `"ff0000"` all collapse to the
   * OOXML uppercase canonical form `"FF0000"`. Malformed inputs (wrong
   * length, non-hex characters, alpha-channel forms like `"#FF0000FF"`,
   * non-string escapes from an untyped caller) collapse to `undefined`
   * so the writer skips the entire `<a:solidFill>` block and the data
   * table inherits the theme text color (Excel's reference behavior
   * for fresh data tables that have not had a custom color picked).
   *
   * Default: omitted — the data table renders at the theme text color
   * (no `<a:solidFill>` block, matching Excel's reference
   * serialization for fresh data tables whose typography has not been
   * customized).
   *
   * The `<c:txPr>` block lands after the four required boolean children
   * per the CT_DTable schema sequence (ECMA-376 Part 1, §21.2.2.54).
   * Mirrors the chart-title `titleColor` / axis-title `axisTitleColor`
   * / axis tick-label `labelColor` / legend `legendFontColor` /
   * data-label `fontColor` knobs — same accept-with-or-without-`#` hex
   * grammar, same OOXML `<a:solidFill><a:srgbClr val=".."/>` mapping —
   * so a caller can thread a single hex string through every
   * typography-pinning slot. Composes independently with
   * {@link fontSize} and the four boolean toggles — so a caller can
   * pin the color without touching the rest of the configuration.
   *
   * Only meaningful for chart families with axes; silently dropped on
   * pie / doughnut along with the rest of the data-table configuration.
   */
  fontColor?: string;
  /**
   * Data-table bold flag. Maps to `<c:dTable><c:txPr><a:p><a:pPr>
   * <a:defRPr b=".."/></a:pPr></a:p></c:txPr></c:dTable>` — Excel's
   * "Format Data Table -> Font -> Bold" toggle. The OOXML attribute is
   * the `xsd:boolean` bold flag on `CT_TextCharacterProperties`
   * (ECMA-376 Part 1, §21.1.2.3.7); the writer lands `b="1"` (bold) or
   * `b="0"` (the OOXML default — non-bold) on the default-paragraph
   * `<a:defRPr>` slot inside the `<c:dTable><c:txPr>` block so a
   * re-parse picks the flag up off the canonical slot the OOXML schema
   * exposes.
   *
   * Default: omitted — the data table renders non-bold (no `b`
   * attribute, matching Excel's reference serialization for fresh data
   * tables whose typography has not been customized; the OOXML default
   * `0` collapses to absence). Set `true` to emit `b="1"` so the table
   * renders bold; set `false` explicitly to pin `b="0"` (functionally
   * identical to omission, but useful when overriding a templated
   * chart that had bold pinned upstream).
   *
   * Composes independently with the other dataTable typography knobs —
   * {@link fontSize} / {@link fontColor} — and the four boolean
   * toggles. Mirrors the chart-title `titleBold`, axis-title
   * `axisTitleBold`, axis tick-label `labelBold`, legend `legendBold`,
   * and data-label `dataLabels.bold` knobs — same boolean shape, same
   * OOXML `<a:defRPr b=".."/>` mapping — so a caller can thread a
   * single bold value through every typography-pinning slot.
   *
   * Only meaningful for chart families with axes; silently dropped on
   * pie / doughnut along with the rest of the data-table configuration.
   */
  bold?: boolean;
  /**
   * Data-table italic flag. Maps to `<c:dTable><c:txPr><a:p><a:pPr>
   * <a:defRPr i=".."/></a:pPr></a:p></c:txPr></c:dTable>` — Excel's
   * "Format Data Table -> Font -> Italic" toggle. The OOXML attribute
   * is the `xsd:boolean` italic flag on `CT_TextCharacterProperties`
   * (ECMA-376 Part 1, §21.1.2.3.7); the writer lands `i="1"` (italic)
   * or `i="0"` (the OOXML default — non-italic) on the default-paragraph
   * `<a:defRPr>` slot inside the `<c:dTable><c:txPr>` block so a
   * re-parse picks the flag up off the canonical slot the OOXML schema
   * exposes.
   *
   * Default: omitted — the data table renders non-italic (no `i`
   * attribute, matching Excel's reference serialization for fresh data
   * tables whose typography has not been customized; the OOXML default
   * `0` collapses to absence). Set `true` to emit `i="1"` so the table
   * renders italic; set `false` explicitly to pin `i="0"` (functionally
   * identical to omission, but useful when overriding a templated
   * chart that had italic pinned upstream).
   *
   * Composes independently with the other dataTable typography knobs —
   * {@link fontSize} / {@link fontColor} / {@link bold} — and the four
   * boolean toggles. Mirrors the chart-title `titleItalic`, axis-title
   * `axisTitleItalic`, axis tick-label `labelItalic`, legend
   * `legendItalic`, and data-label `dataLabels.italic` knobs — same
   * boolean shape, same OOXML `<a:defRPr i=".."/>` mapping — so a
   * caller can thread a single italic value through every typography-
   * pinning slot.
   *
   * Only meaningful for chart families with axes; silently dropped on
   * pie / doughnut along with the rest of the data-table configuration.
   */
  italic?: boolean;
  /**
   * Data-table underline flag. Maps to `<c:dTable><c:txPr><a:p><a:pPr>
   * <a:defRPr u=".."/></a:pPr></a:p></c:txPr></c:dTable>` — Excel's
   * "Format Data Table -> Font -> Underline" toggle. The OOXML
   * attribute is the `ST_TextUnderlineType` enumeration on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7); the
   * writer lands `u="sng"` (single underline — Excel's UI variant) or
   * `u="none"` (the OOXML default — non-underlined) on the
   * default-paragraph `<a:defRPr>` slot inside the `<c:dTable><c:txPr>`
   * block so a re-parse picks the flag up off the canonical slot the
   * OOXML schema exposes.
   *
   * Default: omitted — the data table renders non-underlined (no `u`
   * attribute, matching Excel's reference serialization for fresh data
   * tables whose typography has not been customized; the OOXML default
   * `"none"` collapses to absence). Set `true` to emit `u="sng"` so
   * the table renders single-underlined; set `false` explicitly to pin
   * `u="none"` (functionally identical to omission, but useful when
   * overriding a templated chart that had underline pinned upstream).
   *
   * The reader surfaces only the boolean shape Excel's UI exposes —
   * `u="sng"` collapses to `true`, `u="none"` collapses to `false`,
   * and the schema's other variants (`"dbl"`, `"heavy"`, `"dotted"`,
   * `"dotDash"`, `"wavy"`, etc.) collapse to `undefined` rather than
   * silently downgrade the choice to a single line on round-trip.
   *
   * Composes independently with the other dataTable typography knobs —
   * {@link fontSize} / {@link fontColor} / {@link bold} / {@link italic}
   * — and the four boolean toggles. Mirrors the chart-title
   * `titleUnderline`, axis-title `axisTitleUnderline`, axis tick-label
   * `labelUnderline`, legend `legendUnderline`, and data-label
   * `dataLabels.underline` knobs — same boolean shape, same OOXML
   * `<a:defRPr u=".."/>` mapping — so a caller can thread a single
   * underline value through every typography-pinning slot.
   *
   * Only meaningful for chart families with axes; silently dropped on
   * pie / doughnut along with the rest of the data-table configuration.
   */
  underline?: boolean;
  /**
   * Data-table strikethrough flag. Maps to `<c:dTable><c:txPr><a:p>
   * <a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr></c:dTable>` —
   * Excel's "Format Data Table -> Font -> Strikethrough" toggle. The
   * OOXML attribute is the `ST_TextStrikeType` enumeration on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7); the
   * writer lands `strike="sngStrike"` (single line — Excel's UI
   * variant) on the default-paragraph `<a:defRPr>` slot inside the
   * `<c:dTable><c:txPr>` block so a re-parse picks the flag up off the
   * canonical slot the OOXML schema exposes.
   *
   * Default: omitted — the data table renders without a strikethrough
   * (no `strike` attribute, matching Excel's reference serialization
   * for fresh data tables whose typography has not been customized).
   * Set `true` to emit `strike="sngStrike"` so the table renders with
   * a single-line strikethrough; explicit `false` collapses to
   * absence (functionally identical, since the OOXML default
   * `"noStrike"` is what Excel's UI exposes as "off"; the writer
   * keeps the surfaced shape consistent with what the UI authors —
   * `"sngStrike"` only, never `"noStrike"` or `"dblStrike"`).
   *
   * The reader surfaces only the boolean shape Excel's UI exposes —
   * `strike="sngStrike"` collapses to `true`; the OOXML default
   * `"noStrike"` and the schema's other variant `"dblStrike"` (and
   * malformed tokens) collapse to `undefined` so absence and
   * `"noStrike"` round-trip identically through `cloneChart`. The
   * reader never silently downgrades `"dblStrike"` to a single line on
   * round-trip; only the UI-default `"sngStrike"` survives the parse.
   *
   * Composes independently with the other dataTable typography knobs —
   * {@link fontSize} / {@link fontColor} / {@link bold} / {@link italic}
   * / {@link underline} — and the four boolean toggles. Mirrors the
   * chart-title `titleStrikethrough`, axis-title
   * `axisTitleStrikethrough`, axis tick-label `labelStrikethrough`,
   * legend `legendStrikethrough`, and data-label
   * `dataLabels.strikethrough` knobs — same boolean shape, same OOXML
   * `<a:defRPr strike=".."/>` mapping — so a caller can thread a
   * single strikethrough value through every typography-pinning slot.
   *
   * Only meaningful for chart families with axes; silently dropped on
   * pie / doughnut along with the rest of the data-table configuration.
   */
  strikethrough?: boolean;
  /**
   * Data-table font family / typeface. Maps to `<c:dTable><c:txPr>
   * <a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr></a:pPr>
   * </a:p></c:txPr></c:dTable>` — Excel's "Format Data Table -> Font
   * -> Font" picker. The OOXML `<a:latin typeface=".."/>` element
   * carries the typeface name (`CT_TextFont` on
   * `CT_TextCharacterProperties`' latin slot — ECMA-376 Part 1,
   * §21.1.2.3.7); the writer lands the element on the default-
   * paragraph `<a:defRPr>` slot inside the `<c:dTable><c:txPr>` block
   * so a re-parse picks the typeface up off the canonical slot the
   * OOXML schema exposes.
   *
   * Accepts any non-empty string typeface name (e.g. `"Calibri"`,
   * `"Arial"`, `"Times New Roman"`); the writer trims surrounding
   * whitespace and emits the trimmed value verbatim (XML-escaped) so
   * Excel can resolve the named font from the workbook's font scheme
   * or the host system's installed fonts. Empty / whitespace-only
   * strings and non-string tokens collapse to `undefined` so the
   * writer skips the entire `<a:latin>` element and the data table
   * inherits Excel's reference theme typeface.
   *
   * Default: omitted — the data table renders in Excel's reference
   * theme typeface (no `<a:latin>` element, the writer skips the
   * element entirely). Pin a typeface name to render the table in
   * that font.
   *
   * The `<a:latin>` element follows `<a:solidFill>` per the
   * CT_TextCharacterProperties child sequence so a data-table block
   * with both `fontColor` and `fontFamily` set lands the children in
   * canonical schema order — a fresh chart matches Excel's reference
   * serialization byte-for-byte. Composes independently with the
   * other dataTable typography knobs — {@link fontSize} /
   * {@link fontColor} / {@link bold} / {@link italic} /
   * {@link underline} / {@link strikethrough} — and the four boolean
   * toggles. Mirrors {@link SheetChart.titleFontFamily} /
   * {@link SheetChart.legendFontFamily} /
   * {@link SheetChart.axes.x.axisTitleFontFamily} /
   * {@link SheetChart.axes.x.labelFontFamily} /
   * {@link ChartDataLabels.fontFamily} — same accept-and-trim
   * grammar, same OOXML `<a:latin typeface=".."/>` mapping — so a
   * caller can thread a single typeface string through every
   * typography-pinning slot.
   *
   * Only meaningful for chart families with axes; silently dropped on
   * pie / doughnut along with the rest of the data-table configuration.
   */
  fontFamily?: string;
  /**
   * Data-table background fill color as a 6-digit RGB hex string (e.g.
   * `"F2F2F2"`). Maps to `<c:dTable><c:spPr><a:solidFill><a:srgbClr
   * val="RRGGBB"/></a:solidFill></c:spPr></c:dTable>` (CT_DTable,
   * ECMA-376 Part 1, §21.2.2.54) — Excel's "Format Data Table -> Fill
   * -> Solid fill -> Color" picker (the same dialog the user reaches
   * by right-clicking the data table grid). The OOXML `<a:srgbClr
   * val=".."/>` carries the 6-character uppercase hex sRGB color
   * (`CT_SRgbColor` inside `CT_ShapeProperties`' fill choice — ECMA-376
   * Part 1, §20.1.2.3.32 / §20.1.2.3.13). The `<c:spPr>` slot lives
   * after the four required boolean children (`<c:showHorzBorder>`,
   * `<c:showVertBorder>`, `<c:showOutline>`, `<c:showKeys>`) and before
   * the optional `<c:txPr>` per CT_DTable — distinct from the typography
   * `<c:txPr>` block which carries the font color knobs.
   *
   * Distinct from {@link fontColor} — the fill color paints the cell
   * backgrounds of the data table, while the font color tints the
   * series-name / value text drawn inside those cells. A caller can
   * pin both knobs (e.g. a brand-color background with white text)
   * since they land on different host elements (`<c:spPr>` for the
   * fill, `<c:txPr>` for the typography).
   *
   * Accepts a leading `#` and any case; the writer collapses to the
   * OOXML canonical uppercase form. Malformed inputs (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` and the writer omits
   * the entire `<c:spPr>` block (Excel's reference serialization for
   * a data table that inherits the auto-fill — typically transparent
   * on top of the plot area).
   *
   * Default: omitted — the data table inherits the auto-fill Excel
   * picks from the workbook theme (typically transparent so the plot
   * area's fill shows through). Pin a hex color to mirror Excel's
   * "Format Data Table -> Fill -> Solid fill" knob and paint a flat
   * background behind the table grid.
   *
   * Patterned / gradient / picture fills are not modelled — only the
   * solid sRGB form lands on the wire. Theme-color references
   * (`<a:schemeClr>`) likewise drop to `undefined` so a parsed value
   * always carries a literal hex Excel will render byte-for-byte.
   * Mirrors `plotAreaFillColor` / `legendFillColor` / `titleFillColor`
   * / `chartSpaceFillColor` / `axisTitleFillColor` — same accept-with-
   * or-without-`#` hex grammar, same OOXML `<c:spPr><a:solidFill>
   * <a:srgbClr val=".."/></a:solidFill></c:spPr>` mapping — so a
   * caller can thread a single hex string through every `<c:spPr>`-
   * based fill slot.
   *
   * Only meaningful for chart families with axes; silently dropped on
   * pie / doughnut along with the rest of the data-table configuration.
   */
  fillColor?: string;
  /**
   * Data-table border (line) color as a 6-digit RGB hex string (e.g.
   * `"1F77B4"`). Maps to `<c:dTable><c:spPr><a:ln><a:solidFill>
   * <a:srgbClr val="RRGGBB"/></a:solidFill></a:ln></c:spPr></c:dTable>`
   * (CT_DTable, ECMA-376 Part 1, §21.2.2.54) — Excel's "Format Data
   * Table -> Border -> Solid line -> Color" picker (the same dialog
   * the user reaches by right-clicking the data table grid). The
   * OOXML `<a:srgbClr val=".."/>` carries the 6-character uppercase
   * hex sRGB color (`CT_SRgbColor` inside `<a:ln>`'s solid fill choice
   * — ECMA-376 Part 1, §20.1.2.3.32 / §20.1.2.3.24). The `<a:ln>`
   * child sits inside the `<c:spPr>` block alongside the optional
   * `<a:solidFill>` fill child, in `CT_ShapeProperties` schema order
   * (fill before stroke).
   *
   * Distinct from {@link fillColor} — the border color paints the
   * outline around the data-table block, while the fill color paints
   * the cell backgrounds inside. The two knobs share the `<c:spPr>`
   * host but land on different children (`<a:solidFill>` for the
   * fill, `<a:ln>` for the stroke), and the writer authors a single
   * `<c:spPr>` whenever either knob is set. A caller can pin one
   * without the other; pinning both produces a filled data table
   * with a colored border.
   *
   * Distinct from the four boolean toggles ({@link showHorzBorder} /
   * {@link showVertBorder} / {@link showOutline}) which govern the
   * inner grid lines and outer outline visibility (rendered with the
   * theme's automatic stroke). The {@link borderColor} knob colors
   * the entire `<c:spPr>` outline — Excel applies the color to the
   * outer border and the grid lines together when the matching
   * toggles are on.
   *
   * Accepts a leading `#` and any case; the writer collapses to the
   * OOXML canonical uppercase form. Malformed inputs (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` and the writer omits
   * the `<a:ln>` block (Excel's reference serialization for a data
   * table that inherits the auto-stroke — typically the theme's
   * default line color).
   *
   * Default: omitted — the data table inherits the auto-stroke Excel
   * picks from the workbook theme. Pin a hex color to mirror Excel's
   * "Format Data Table -> Border -> Solid line" knob and paint a
   * flat outline around the table grid — useful for dashboard tiles
   * where the data table should be visually framed against the
   * surrounding chart frame.
   *
   * Patterned / gradient strokes are not modelled — only the solid
   * sRGB form lands on the wire. Theme-color references
   * (`<a:schemeClr>`) likewise drop to `undefined` so a parsed value
   * always carries a literal hex Excel will render byte-for-byte.
   * Mirrors `plotAreaBorderColor` / `legendBorderColor` /
   * `titleBorderColor` / `axisTitleBorderColor` — same accept-with-
   * or-without-`#` hex grammar, same OOXML `<a:ln><a:solidFill>
   * <a:srgbClr val=".."/></a:solidFill></a:ln>` mapping — so a
   * caller can thread a single hex string through every `<a:ln>`-
   * based stroke slot.
   *
   * Only meaningful for chart families with axes; silently dropped on
   * pie / doughnut along with the rest of the data-table configuration.
   */
  borderColor?: string;
  /**
   * Data-table border (stroke) thickness in points (e.g. `1.5`). Maps
   * to the `w` attribute on `<c:dTable><c:spPr><a:ln w="EMU">` — Excel's
   * "Format Data Table -> Border -> Width" spinner. The OOXML `w`
   * attribute carries the stroke width in English Metric Units
   * (1 pt = 12 700 EMU) per `CT_LineProperties` (ECMA-376 Part 1,
   * §20.1.2.3.24). The writer multiplies by 12 700 and rounds to the
   * nearest integer because the schema types `w` as `xsd:int`.
   *
   * Accepts any finite number; values are clamped to the `0.25..13.5`
   * pt band Excel's UI exposes (the same band used by every other
   * chart-frame border-width knob) and snapped to the 0.25 pt grid so
   * a parsed-then-written width does not drift across round-trips.
   * Non-finite / non-numeric / `NaN` values collapse to `undefined`
   * and the writer omits the `w` attribute (the line keeps Excel's
   * auto-thickness — typically 0.75 pt).
   *
   * Default: omitted — the data-table border inherits Excel's
   * auto-thickness.
   *
   * Composes independently with {@link borderColor} — the width
   * attribute lands on the same `<a:ln>` element as the color's
   * `<a:solidFill>` child, but the writer authors `<a:ln>` whenever
   * either knob is set.
   *
   * Only meaningful for chart families with axes; silently dropped on
   * pie / doughnut along with the rest of the data-table configuration.
   */
  borderWidth?: number;
  /**
   * Data-table border (stroke) preset dash pattern. Maps to the `val`
   * attribute on `<c:dTable><c:spPr><a:ln><a:prstDash val=".."/>`.
   * Mirrors Excel's "Format Data Table -> Border -> Dash type" picker.
   * Same {@link ChartBorderDash} accept-or-drop grammar as the
   * chart-level {@link SheetChart.plotAreaBorderDash} — `"solid"`
   * collapses to `undefined` so absence and the OOXML default
   * round-trip identically.
   *
   * Composes independently with {@link borderColor} and
   * {@link borderWidth} — all three knobs share the same `<a:ln>`
   * element. Only meaningful for chart families with axes; silently
   * dropped on pie / doughnut along with the rest of the data-table
   * configuration.
   */
  borderDash?: ChartBorderDash;
}

/**
 * Granular chart-protection configuration for {@link SheetChart.protection}.
 *
 * Each flag mirrors a `CT_Boolean` child of `<c:chartSpace><c:protection>`
 * (CT_Protection, ECMA-376 Part 1, §21.2.2.142) — the chart-space-level
 * lock that pairs with the host worksheet's `<sheetProtection>` to
 * decide which interactions Excel still allows when the parent sheet is
 * protected. The element only takes effect when the worksheet that
 * embeds the chart is itself protected; on an unprotected sheet the
 * flags are silently ignored by Excel.
 *
 * Every field defaults to `false` (the OOXML default Excel itself
 * emits) — `false` means the action is permitted, `true` means the
 * action is locked. Each field is independent; pass any combination of
 * the five flags to lock individual interactions.
 *
 * Pass an empty object (`{}`) to declare an empty `<c:protection>`
 * shell with every flag at the OOXML default — equivalent to passing
 * `protection: true` and useful for round-trip parity with templates
 * that author the bare element without pinning any flags.
 */
export interface ChartProtection {
  /**
   * Lock the chart object (its frame, anchor, and contents) against
   * direct manipulation. Maps to `<c:chartObject val=".."/>`.
   * Default: `false` — the chart is movable / resizable on a protected
   * worksheet. Set `true` to freeze the chart's position and contents.
   */
  chartObject?: boolean;
  /**
   * Lock the underlying data references the chart points at. Maps to
   * `<c:data val=".."/>`. Default: `false` — Excel allows the user to
   * re-pick data ranges via "Select Data Source" even when the parent
   * sheet is protected. Set `true` to keep the series ranges pinned.
   */
  data?: boolean;
  /**
   * Lock chart formatting (colors, fills, fonts, layout). Maps to
   * `<c:formatting val=".."/>`. Default: `false` — the user may still
   * tweak Format-Chart-Element panes on a protected sheet. Set `true`
   * to lock down the rendered look.
   */
  formatting?: boolean;
  /**
   * Lock click-to-select on chart elements. Maps to
   * `<c:selection val=".."/>`. Default: `false` — the user can still
   * click into series, legend entries, or labels to inspect them. Set
   * `true` to disable element selection entirely.
   */
  selection?: boolean;
  /**
   * Lock chart-level UI affordances such as the floating "+ Chart
   * Elements" button and the right-click context menu. Maps to
   * `<c:userInterface val=".."/>`. Default: `false` — Excel still
   * surfaces the on-canvas affordances. Set `true` to suppress them.
   */
  userInterface?: boolean;
}

/**
 * Granular 3D-view configuration for {@link SheetChart.view3D}.
 *
 * Maps to the six optional `CT_View3D` children of `<c:chart><c:view3D>`
 * (ECMA-376 Part 1, §21.2.2.228) — Excel's "3-D Rotation" pane on
 * 3D chart families (`bar3D`, `line3D`, `pie3D`, `area3D`, `surface3D`).
 * The element itself is optional; each child is independently optional
 * and Excel falls back to the per-family default for any field the
 * template leaves unset.
 *
 * Every field is independent — pass any subset to override the matching
 * per-family default. Pin a field to a value at the matching boundary
 * of the OOXML simple type and the writer round-trips it as the literal
 * `val` attribute; out-of-range and non-finite numeric inputs drop at
 * write time rather than emit a token Excel would reject.
 *
 * Although `<c:view3D>` is only meaningful on 3D chart families, the
 * OOXML schema accepts it on every CT_Chart, so the writer pins the
 * element whenever the caller provides a non-empty configuration —
 * Excel silently ignores it on 2D families. Useful primarily for
 * round-tripping a 3D template chart through {@link cloneChart}.
 */
export interface ChartView3D {
  /**
   * Rotation around the X axis in degrees. Maps to
   * `<c:rotX val=".."/>` (`ST_RotX`, signed byte). Accepted range:
   * `-90` – `90`. Excel's default is `15` for most 3D families and
   * `0` for `pie3D`. Values outside the range drop at write time.
   */
  rotX?: number;
  /**
   * Height as a percent of the chart width. Maps to
   * `<c:hPercent val=".."/>` (`ST_HPercent`, percent). Accepted range:
   * `5` – `500`. Excel's default is `100`. Values outside the range
   * drop at write time.
   */
  hPercent?: number;
  /**
   * Rotation around the Y axis in degrees. Maps to
   * `<c:rotY val=".."/>` (`ST_RotY`, unsigned short). Accepted range:
   * `0` – `360`. Excel's default is `20` for most 3D families and
   * `0` for `pie3D` / `bar3D` (column orientation). Values outside the
   * range drop at write time.
   */
  rotY?: number;
  /**
   * Depth as a percent of the chart width. Maps to
   * `<c:depthPercent val=".."/>` (`ST_DepthPercent`, percent). Accepted
   * range: `20` – `2000`. Excel's default is `100`. Values outside the
   * range drop at write time.
   */
  depthPercent?: number;
  /**
   * Right-angle-axes flag. Maps to `<c:rAngAx val=".."/>` (CT_Boolean).
   * Default: `false` — perspective foreshortening is applied to the
   * 3D box. Set `true` to draw the chart with axes at right angles
   * (Excel's "Right angle axes" checkbox), which suppresses perspective
   * even when {@link perspective} is non-zero.
   */
  rAngAx?: boolean;
  /**
   * Perspective factor. Maps to `<c:perspective val=".."/>`
   * (`ST_Perspective`, percent). Accepted range: `0` – `240`. Excel's
   * default is `30`. Higher values exaggerate the foreshortening; `0`
   * is the orthographic projection. Silently ignored by Excel when
   * {@link rAngAx} is `true`. Values outside the range drop at write
   * time.
   */
  perspective?: number;
}

/**
 * Scatter sub-style applied at the chart level. Maps to the OOXML
 * `ST_ScatterStyle` enum which sits inside `<c:scatterChart>` as
 * `<c:scatterStyle val=".."/>`. Excel exposes the same six presets
 * under "Change Chart Type → XY (Scatter)":
 *
 * - `"none"`         — markers only, no connecting line and no curves.
 *                      Equivalent to `"marker"` in modern Excel UI.
 * - `"line"`         — straight-line segments between points, no markers.
 * - `"lineMarker"`   — straight-line segments with markers (Excel's
 *                      reference default and the writer's fallback).
 * - `"marker"`       — markers only, no line. Same render as `"none"`;
 *                      OOXML lists both for legacy compatibility.
 * - `"smooth"`       — smoothed (Catmull-Rom-style) curves between
 *                      points, no markers.
 * - `"smoothMarker"` — smoothed curves with markers.
 *
 * Distinct from the per-series {@link ChartSeries.smooth} flag — the
 * series-level toggle paints individual points, while `scatterStyle`
 * is the chart-wide preset Excel selects in the chart-type picker.
 * When both are set, the OOXML schema lets Excel render the union
 * (smooth chart with the series-level smooth still emitted), but
 * Excel's UI normally pairs them: `scatterStyle: "smooth"` implies
 * smoothed series, `scatterStyle: "lineMarker"` implies straight ones.
 */
export type ChartScatterStyle =
  | "none"
  | "line"
  | "lineMarker"
  | "marker"
  | "smooth"
  | "smoothMarker";

/**
 * A single data series inside a chart.
 *
 * `values` and `categories` are A1-style cell range references.
 * Provide either a sheet-qualified reference (e.g. `"Sheet1!$B$2:$B$4"`)
 * or a bare range (`"B2:B4"`). Bare ranges are auto-qualified with the
 * sheet that owns the chart.
 */
export interface ChartSeries {
  /** Series name shown in the legend (e.g. "Revenue"). */
  name?: string;
  /** A1-style range with the series numeric values (e.g. "B2:B10"). */
  values: string;
  /** A1-style range with the category labels (e.g. "A2:A10"). */
  categories?: string;
  /** Optional fill color as a 6-digit RGB hex string (e.g. "1F77B4"). */
  color?: string;
  /**
   * Per-series data label override. Pass `false` to suppress labels
   * for this series even when the chart-level
   * {@link SheetChart.dataLabels} enables them.
   */
  dataLabels?: ChartDataLabels | false;
  /**
   * Smooth the line connecting data points using a Catmull-Rom-style
   * spline. Maps to `<c:smooth val="..">` inside the `<c:ser>` element.
   * Only meaningful for `line` and `scatter` charts — ignored for every
   * other chart kind (the OOXML schema does not allow `<c:smooth>` on
   * bar / column / pie / doughnut / area series).
   *
   * Default: `false` (straight-line segments). Set `true` to render the
   * curved variant Excel offers under "Format Data Series → Line →
   * Smoothed line".
   */
  smooth?: boolean;
  /**
   * Per-series line stroke (dash pattern + width) for line / scatter
   * charts. Maps to `<a:ln>` inside `<c:ser><c:spPr>`. Ignored on every
   * other chart family — bar / column / pie / doughnut / area never
   * render a connecting line, so dashing and stroke width have no
   * visible effect there. See {@link ChartLineStroke}.
   */
  stroke?: ChartLineStroke;
  /**
   * Per-series marker styling. Only meaningful for `line` and
   * `scatter` charts — the OOXML schema places `<c:marker>` on
   * `CT_LineSer` / `CT_ScatterSer` only. Ignored on every other
   * chart family at write time.
   */
  marker?: ChartMarker;
  /**
   * Invert the fill color when the value is negative. Maps to
   * `<c:invertIfNegative val=".."/>` inside the `<c:ser>` element.
   * Only meaningful for `bar` and `column` charts — the OOXML schema
   * places `<c:invertIfNegative>` exclusively on `CT_BarSer` and
   * `CT_Bar3DSer`, so the field is silently dropped on every other
   * chart family at write time.
   *
   * Default: `false` (negative bars share the series fill color).
   * Set `true` to mirror Excel's "Format Data Series → Fill → Invert
   * if negative" toggle, which paints negative bars with white (or
   * the inverted color when the spreadsheet supplies one).
   */
  invertIfNegative?: boolean;
  /**
   * Pie / doughnut slice explosion as a percentage of the radius —
   * the distance the slice is pulled away from the center. Maps to
   * `<c:explosion val=".."/>` inside the `<c:ser>` element. Only
   * meaningful for `pie` and `doughnut` charts — the OOXML schema
   * places `<c:explosion>` exclusively on `CT_PieSer`, so the field
   * is silently dropped on every other chart family at write time.
   *
   * Default: `0` (slices flush against each other). Excel's UI
   * exposes 0–400% under "Format Data Point → Series Options → Pie
   * Explosion"; values outside that band are clamped on write so a
   * round-trip stays inside the range Excel will render. Per-data-point
   * explosion (one slice pulled away while the rest stay flush) is not
   * yet supported — the field applies to every slice in the series.
   */
  explosion?: number;
}

/**
 * Per-series legend-entry override. Maps to `<c:legendEntry>` inside
 * `<c:legend>` — Excel's "Format Legend Entries -> Hide" action surfaces
 * the same element. The OOXML schema (CT_LegendEntry) carries an
 * `<c:idx val="N"/>` selector and an optional `<c:delete val=".."/>`
 * flag; hucre exposes only the hide bit because per-entry text styling
 * (`<c:txPr>`) is not modelled at this layer.
 *
 * The {@link idx} value is the **0-based series index** as it appears
 * in the chart's flattened series list (the same index the writer emits
 * on `<c:ser><c:idx val="N"/></c:ser>`). Pinning `delete: true` hides
 * that entry from the legend; the chart still renders the underlying
 * series, only its swatch / name disappears from the legend block.
 *
 * Useful when a templated dashboard chart carries a helper series whose
 * data should plot but whose name should not crowd the legend (e.g. a
 * trend baseline rendered as a faint area behind the real data).
 *
 * @see {@link SheetChart.legendEntries}
 * @see {@link Chart.legendEntries}
 */
export interface ChartLegendEntry {
  /**
   * 0-based series index the entry refers to. Must be a non-negative
   * integer (matches the OOXML `xsd:unsignedInt` schema on
   * `<c:idx val=".."/>`); non-finite, negative, or non-integer values
   * are dropped at write time rather than emit a token Excel rejects.
   */
  idx: number;
  /**
   * Whether the legend entry is hidden. Maps to
   * `<c:legendEntry><c:delete val=".."/></c:legendEntry>`. Defaults to
   * `false` (the OOXML default — the entry renders); set `true` to
   * mirror Excel's "Format Legend Entries -> Hide" action.
   *
   * The writer emits an entry only when at least one of {@link idx} /
   * {@link delete} carries a meaningful value, and it always emits the
   * `<c:delete>` child explicitly so a re-parse sees the same flag.
   */
  delete?: boolean;
}

/**
 * Manual placement of a chart sub-element (legend / plot area / title /
 * data table) inside the chart frame. Maps to OOXML's `<c:manualLayout>`
 * (`CT_ManualLayout`, ECMA-376 Part 1, §21.2.2.115) — the element block
 * that backs Excel's "Format <element> -> Position -> Custom" knob and
 * pins where the element draws inside the chart's drawing area.
 *
 * All four coordinates are fractions of the chart frame in the range
 * `0..1` — `(0, 0)` is the upper-left of the chart frame, `(1, 1)` is
 * the lower-right; widths / heights are sized as fractions of the chart
 * frame width / height. Each axis is independently optional so a caller
 * can pin only the position ({@link x} / {@link y}) and let the element
 * keep its automatic size, only the size ({@link w} / {@link h}) and
 * let the element keep its automatic anchor, or any combination.
 *
 * The writer always emits the matching `<c:xMode>` / `<c:yMode>` /
 * `<c:wMode>` / `<c:hMode>` children with `val="edge"` (Excel's "Format
 * Legend -> Position" reference shape — the coordinates are absolute
 * fractions of the chart frame, not deltas from the auto-layout
 * baseline). The reader collapses `val="factor"` (delta from
 * auto-layout) onto the same shape so a templated chart that pinned
 * the alternate mode still round-trips through {@link cloneChart}; the
 * writer normalizes to `"edge"` on emit since Excel itself emits the
 * absolute form when the user drags an element to a custom position.
 *
 * @see {@link SheetChart.legendLayout}
 * @see {@link Chart.legendLayout}
 */
export interface ChartManualLayout {
  /**
   * Horizontal anchor as a fraction of the chart frame width. Maps to
   * `<c:manualLayout><c:x val=".."/></c:manualLayout>`. Range: `0..1`
   * (`0` is the chart frame's left edge, `1` is the right edge).
   * Out-of-range / non-finite / non-numeric inputs collapse to
   * `undefined` so the writer skips the `<c:x>` slot rather than emit a
   * token Excel would reject.
   */
  x?: number;
  /**
   * Vertical anchor as a fraction of the chart frame height. Maps to
   * `<c:manualLayout><c:y val=".."/></c:manualLayout>`. Range: `0..1`
   * (`0` is the chart frame's top edge, `1` is the bottom edge).
   * Out-of-range / non-finite / non-numeric inputs collapse to
   * `undefined` so the writer skips the `<c:y>` slot.
   */
  y?: number;
  /**
   * Width as a fraction of the chart frame width. Maps to
   * `<c:manualLayout><c:w val=".."/></c:manualLayout>`. Range: `0..1`
   * (`0` collapses the element to a hairline, `1` spans the full chart
   * frame width). Out-of-range / non-finite / non-numeric inputs
   * collapse to `undefined` so the writer skips the `<c:w>` slot.
   */
  w?: number;
  /**
   * Height as a fraction of the chart frame height. Maps to
   * `<c:manualLayout><c:h val=".."/></c:manualLayout>`. Range: `0..1`
   * (`0` collapses the element to a hairline, `1` spans the full chart
   * frame height). Out-of-range / non-finite / non-numeric inputs
   * collapse to `undefined` so the writer skips the `<c:h>` slot.
   */
  h?: number;
}

/**
 * A chart embedded into a worksheet via the drawing layer.
 *
 * Excel anchors charts to cells using the same `xdr:twoCellAnchor`
 * mechanism it uses for images. The chart is stored in
 * `xl/charts/chartN.xml` and wired into the worksheet through a
 * drawing part.
 */
export interface SheetChart {
  /**
   * Chart family. `"bar"` is horizontal, `"column"` is vertical (the
   * Excel default). Both map to `<c:barChart>` with different
   * `<c:barDir>` values.
   */
  type: WriteChartKind;
  /** Optional chart title rendered above the plot area. */
  title?: string;
  /** One or more data series. */
  series: ChartSeries[];
  /** Cell anchor — `to` defaults to a 6×15 area below `from`. */
  anchor: {
    from: { row: number; col: number };
    to?: { row: number; col: number };
  };
  /**
   * Bar/column subtype. Default: `"clustered"`. `"stacked"` and
   * `"percentStacked"` group series end-to-end. Ignored for non-bar
   * chart kinds.
   */
  barGrouping?: "clustered" | "stacked" | "percentStacked";
  /**
   * Bar/column gap width as a percentage of the bar width — the empty
   * space between adjacent category groups. Accepted range: `0` – `500`
   * (the OOXML `ST_GapAmount` schema). Excel's default is `150` (each
   * group's gap equals 1.5× the bar width). Smaller values pack groups
   * tighter; `0` removes the gap entirely. Maps to
   * `<c:barChart><c:gapWidth val=".."/></c:barChart>`. Ignored for
   * non-bar / non-column chart kinds.
   */
  gapWidth?: number;
  /**
   * Bar/column series overlap as a percentage of the bar width.
   * Accepted range: `-100` – `100` (the OOXML `ST_Overlap` schema).
   * Negative values open a gap between series within a group, positive
   * values stack them on top of each other. Excel's default is `0` for
   * `clustered` (side-by-side) and `100` for `stacked` /
   * `percentStacked` (fully overlapped). Maps to
   * `<c:barChart><c:overlap val=".."/></c:barChart>`. Ignored for
   * non-bar / non-column chart kinds.
   */
  overlap?: number;
  /**
   * Line subtype. Default: `"standard"`. `"stacked"` accumulates
   * series end-to-end, `"percentStacked"` normalizes each category to
   * 100%. Ignored for non-line chart kinds. Maps to
   * `<c:lineChart><c:grouping val="..."/></c:lineChart>`.
   */
  lineGrouping?: "standard" | "stacked" | "percentStacked";
  /**
   * Area subtype. Default: `"standard"`. `"stacked"` paints series on
   * top of each other, `"percentStacked"` normalizes each category to
   * 100%. Ignored for non-area chart kinds. Maps to
   * `<c:areaChart><c:grouping val="..."/></c:areaChart>`.
   */
  areaGrouping?: "standard" | "stacked" | "percentStacked";
  /**
   * Whether the chart paints `<c:dropLines>` — vertical reference lines
   * that drop from each data point down to the category axis. Mirrors
   * Excel's "Add Chart Element -> Lines -> Drop Lines" toggle.
   *
   * Default: `false` (no drop lines, Excel's reference serialization).
   * Set `true` to emit `<c:dropLines/>` on the chart-type element so
   * Excel paints the connector lines.
   *
   * The OOXML schema places `<c:dropLines>` exclusively on
   * `<c:lineChart>`, `<c:line3DChart>`, `<c:areaChart>`, and
   * `<c:area3DChart>`. Hucre's writer authors `<c:lineChart>` and
   * `<c:areaChart>` only, so the flag is silently ignored on every
   * other chart kind (`bar` / `column` / `pie` / `doughnut` /
   * `scatter`).
   */
  dropLines?: boolean;
  /**
   * Whether the chart paints `<c:hiLowLines>` — vertical reference
   * lines that connect the highest and lowest series values at each
   * category position. Mirrors Excel's "Add Chart Element -> Lines ->
   * High-Low Lines" toggle (the same connector painted on stock
   * charts).
   *
   * Default: `false` (no high-low lines, Excel's reference
   * serialization). Set `true` to emit `<c:hiLowLines/>` on the
   * chart-type element so Excel paints the connectors.
   *
   * The OOXML schema places `<c:hiLowLines>` exclusively on
   * `<c:lineChart>`, `<c:line3DChart>`, and `<c:stockChart>`. Hucre's
   * writer authors `<c:lineChart>` only, so the flag is silently
   * ignored on every other chart kind (`bar` / `column` / `pie` /
   * `doughnut` / `area` / `scatter`).
   */
  hiLowLines?: boolean;
  /**
   * Whether the chart paints `<c:serLines>` — connector lines drawn
   * between paired data points across consecutive series in a stacked
   * bar / column chart. Mirrors Excel's "Add Chart Element -> Lines ->
   * Series Lines" toggle (visible only on stacked / 100% stacked bar
   * and column charts).
   *
   * Default: `false` (no series lines, Excel's reference serialization).
   * Set `true` to emit `<c:serLines/>` on the chart-type element so
   * Excel paints the connectors.
   *
   * The OOXML schema places `<c:serLines>` exclusively on
   * `<c:barChart>` and `<c:ofPieChart>`. Hucre's writer authors
   * `<c:barChart>` only, so the flag is silently ignored on every
   * other chart kind (`line` / `pie` / `doughnut` / `area` /
   * `scatter`). Excel only renders the connectors on the
   * `stacked` / `percentStacked` groupings — a clustered chart with
   * `serLines: true` still pins the element but Excel paints nothing
   * (matches Excel's own behavior when the toggle is flipped on a
   * clustered chart).
   */
  serLines?: boolean;
  /**
   * Doughnut hole size as a percentage of the outer radius. Accepted
   * range: 10 – 90 (Excel's UI clamps values outside this band).
   * Default: `50` — the Excel default. Ignored for non-doughnut chart
   * kinds.
   */
  holeSize?: number;
  /**
   * Pie / doughnut starting angle in degrees, measured clockwise from
   * the 12 o'clock position. Accepted range: 0 – 360 (the OOXML schema
   * range). Default: `0` — the Excel default (first slice begins at
   * 12 o'clock). Maps to `<c:firstSliceAng val=".."/>`. Ignored for
   * non-pie / non-doughnut chart kinds.
   *
   * Useful for rotating the first wedge into a specific quadrant when
   * composing a dashboard whose pie / doughnut charts should align
   * visually (e.g. `90` to start at 3 o'clock).
   */
  firstSliceAng?: number;
  /**
   * Whether the legend is shown and where. Default: `"right"` for
   * pie/doughnut/bar/line/area, `"bottom"` for scatter. Pass `false`
   * to hide the legend.
   */
  legend?: false | "top" | "bottom" | "left" | "right" | "topRight";
  /**
   * Whether the legend overlaps the plot area. Maps to
   * `<c:legend><c:overlay val=".."/></c:legend>` — Excel's "Format
   * Legend -> Show the legend without overlapping the chart" toggle
   * (the checkbox is the inverse of this flag — checked means `false`,
   * unchecked means `true`). Default: `false` (the OOXML default Excel
   * itself emits) — the legend reserves its own slot and the plot area
   * shrinks to accommodate it. Set `true` to draw the legend on top of
   * the plot area so the chart series get the full frame.
   *
   * Silently ignored when `legend === false` (no legend element is
   * emitted) — there is no overlay flag to set on a hidden legend.
   */
  legendOverlay?: boolean;
  /**
   * Per-series legend-entry overrides. Maps to
   * `<c:legend><c:legendEntry>...</c:legendEntry></c:legend>` — Excel's
   * "Format Legend Entries -> Hide" action surfaces the same element.
   * Useful for hiding individual series from the legend without removing
   * them from the plot (a templated dashboard's helper series whose
   * data should plot but whose name should not crowd the legend).
   *
   * Each entry references a series by 0-based {@link ChartLegendEntry.idx}
   * and pins {@link ChartLegendEntry.delete} `true` to hide it. Entries
   * whose `idx` is non-integer / non-finite / negative are silently
   * dropped at write time rather than emit a token Excel would reject.
   * Duplicate `idx` values are deduplicated — the last entry wins so a
   * caller can override an inherited override without manually pruning
   * the list.
   *
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no slot to host the entries on a hidden
   * legend. The OOXML schema (CT_Legend) places `<c:legendEntry>` after
   * `<c:legendPos>` and before `<c:layout>` / `<c:overlay>`; the writer
   * emits in that order so a re-parse sees the canonical sequence.
   */
  legendEntries?: ChartLegendEntry[];
  /**
   * Legend font size in whole or half points. Maps to
   * `<c:legend><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
   * </c:txPr></c:legend>` — Excel's "Format Legend -> Font -> Size"
   * knob. The OOXML attribute is in 100ths of a point, so 12pt
   * serializes as `sz="1200"` and 9pt (Excel's reference default for
   * the legend) as `sz="900"`; the writer performs the conversion at
   * emit time and lands the value on the default-paragraph
   * `<a:defRPr>` slot inside the legend's `<c:txPr>` block so a
   * re-parse picks the size up off the canonical slot the OOXML schema
   * exposes.
   *
   * Accepted range: `1..400`pt (the band the OOXML `ST_TextFontSize`
   * schema exposes — `100..400000` in 100ths of a point). Fractional
   * inputs round to the nearest 0.5pt (Excel's UI granularity); inputs
   * outside the band, `NaN`, `Infinity`, and non-numeric inputs all
   * collapse to `undefined` so the writer drops the `<c:txPr>` block
   * rather than emit a token Excel would reject — the rendered legend
   * falls back to the theme-default 9pt.
   *
   * Default: omitted — the legend renders at Excel's reference 9pt
   * (the OOXML schema's application-default for chart legend text). Set
   * an explicit value to scale the legend up for a hero dashboard tile
   * (e.g. `14`) or down to fit a tight sidebar slot (e.g. `7`).
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no slot to host the size in that case.
   *
   * Mirrors {@link titleFontSize} / {@link axes.x.axisTitleFontSize} /
   * {@link axes.x.labelFontSize} — same range, same normalization,
   * same OOXML conversion factor — so a caller can thread a single
   * size value through every typography-pinning slot without bookkeeping
   * the units. Composes independently with {@link legend} /
   * {@link legendOverlay} / {@link legendEntries}: all four fields land
   * on the same `<c:legend>` element so a single configuration call
   * threads cleanly through every legend knob Excel exposes.
   */
  legendFontSize?: number;
  /**
   * Legend bold flag. Maps to `<c:legend><c:txPr><a:p><a:pPr>
   * <a:defRPr b=".."/></a:pPr></a:p></c:txPr></c:legend>` — Excel's
   * "Format Legend -> Font -> Bold" toggle. The OOXML attribute is the
   * `xsd:boolean` `b` on `CT_TextCharacterProperties` (ECMA-376 Part 1,
   * §21.1.2.3.7); the writer lands the value on the default-paragraph
   * `<a:defRPr>` slot inside the legend's `<c:txPr>` block so a
   * re-parse picks the flag up off the canonical slot the OOXML schema
   * exposes.
   *
   * Default: omitted — the legend renders non-bold (no `b` attribute,
   * matching Excel's reference serialization for a fresh chart legend
   * whose typography has not been customized; the OOXML default `false`
   * collapses to absence). Set `true` to emit `b="1"` so the legend
   * renders bold; set `false` explicitly to pin the non-default `b="0"`
   * (functionally identical to omission, but useful when overriding a
   * templated legend that had bold pinned upstream).
   *
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no `<c:txPr>` slot to host the flag in that
   * case. Mirrors {@link titleBold} / {@link axes.x.axisTitleBold} /
   * {@link axes.x.labelBold} — same boolean-with-explicit-default
   * shape, same OOXML `<a:defRPr b=".."/>` mapping — so a caller can
   * thread a single bold value through every typography-pinning slot.
   * Composes independently with {@link legend} / {@link legendOverlay}
   * / {@link legendEntries}: all four fields land on the same
   * `<c:legend>` element so a single configuration call threads
   * cleanly through every legend knob Excel exposes.
   */
  legendBold?: boolean;
  /**
   * Legend italic flag. Maps to `<c:legend><c:txPr><a:p><a:pPr>
   * <a:defRPr i=".."/></a:pPr></a:p></c:txPr></c:legend>` — Excel's
   * "Format Legend -> Font -> Italic" toggle. The OOXML attribute is
   * the `xsd:boolean` `i` on `CT_TextCharacterProperties` (ECMA-376
   * Part 1, §21.1.2.3.7); the writer lands the value on the
   * default-paragraph `<a:defRPr>` slot inside the legend's
   * `<c:txPr>` block so a re-parse picks the flag up off the canonical
   * slot the OOXML schema exposes.
   *
   * Default: omitted — the legend renders non-italic (no `i`
   * attribute, matching Excel's reference serialization for a fresh
   * chart legend whose typography has not been customized; the OOXML
   * default `false` collapses to absence). Set `true` to emit `i="1"`
   * so the legend renders italic; set `false` explicitly to pin the
   * non-default `i="0"` (functionally identical to omission, but
   * useful when overriding a templated legend that had italic pinned
   * upstream).
   *
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no `<c:txPr>` slot to host the flag in
   * that case. Mirrors {@link titleItalic} /
   * {@link axes.x.axisTitleItalic} / {@link axes.x.labelItalic} —
   * same boolean-with-explicit-default shape, same OOXML
   * `<a:defRPr i=".."/>` mapping — so a caller can thread a single
   * italic value through every typography-pinning slot. Composes
   * independently with {@link legend} / {@link legendOverlay} /
   * {@link legendEntries}: all four fields land on the same
   * `<c:legend>` element so a single configuration call threads
   * cleanly through every legend knob Excel exposes.
   */
  legendItalic?: boolean;
  /**
   * Legend underline flag. Maps to `<c:legend><c:txPr><a:p><a:pPr>
   * <a:defRPr u=".."/></a:pPr></a:p></c:txPr></c:legend>` — Excel's
   * "Format Legend -> Font -> Underline" toggle. The OOXML attribute is
   * the `ST_TextUnderlineType` enumeration on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7); the
   * writer lands `u="sng"` (Excel's UI variant — single underline) or
   * `u="none"` on the default-paragraph `<a:defRPr>` slot inside the
   * legend's `<c:txPr>` block so a re-parse picks the flag up off the
   * canonical slot the OOXML schema exposes.
   *
   * Default: omitted — the legend renders non-underlined (no `u`
   * attribute, matching Excel's reference serialization for a fresh
   * chart legend whose typography has not been customized; the OOXML
   * default `"none"` collapses to absence). Set `true` to emit
   * `u="sng"` so the legend renders single-underlined; set `false`
   * explicitly to pin the OOXML default `u="none"` (functionally
   * identical to omission, but useful when overriding a templated
   * legend that had underline pinned upstream).
   *
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no `<c:txPr>` slot to host the flag in that
   * case. Mirrors {@link titleUnderline} /
   * {@link axes.x.axisTitleUnderline} / {@link axes.x.labelUnderline}
   * — same boolean-with-explicit-default shape, same OOXML
   * `<a:defRPr u=".."/>` mapping — so a caller can thread a single
   * underline value through every typography-pinning slot. Composes
   * independently with {@link legend} / {@link legendOverlay} /
   * {@link legendEntries} / {@link legendFontSize}: all five fields
   * land on the same `<c:legend>` element so a single configuration
   * call threads cleanly through every legend knob Excel exposes.
   */
  legendUnderline?: boolean;
  /**
   * Legend strikethrough flag. Maps to `<c:legend><c:txPr><a:p><a:pPr>
   * <a:defRPr strike=".."/></a:pPr></a:p></c:txPr></c:legend>` — Excel's
   * "Format Legend -> Font -> Strikethrough" toggle. The OOXML attribute
   * is the `ST_TextStrikeType` enumeration on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7); the
   * writer lands `strike="sngStrike"` (Excel's UI variant — single line)
   * on the default-paragraph `<a:defRPr>` slot inside the legend's
   * `<c:txPr>` block so a re-parse picks the flag up off the canonical
   * slot the OOXML schema exposes.
   *
   * Default: omitted — the legend renders non-strikethrough (no
   * `strike` attribute, matching Excel's reference serialization for a
   * fresh chart legend whose typography has not been customized; the
   * OOXML default `"noStrike"` collapses to absence). Set `true` to
   * emit `strike="sngStrike"` so the legend renders with a single
   * strikethrough line; absence and explicit `false` both collapse to
   * omitting the attribute entirely (mirrors `titleStrikethrough` /
   * `axisTitleStrike` / `labelStrikethrough` — the writer emits only
   * the UI variant `"sngStrike"`, never `"noStrike"` or `"dblStrike"`).
   *
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no `<c:txPr>` slot to host the flag in that
   * case. Mirrors {@link titleStrikethrough} /
   * {@link axes.x.axisTitleStrike} / {@link axes.x.labelStrikethrough}
   * — same boolean shape, same OOXML `<a:defRPr strike=".."/>` mapping
   * — so a caller can thread a single strikethrough value through every
   * typography-pinning slot. Composes independently with {@link legend}
   * / {@link legendOverlay} / {@link legendEntries} /
   * {@link legendFontSize} / {@link legendBold} / {@link legendItalic}
   * / {@link legendUnderline}: all eight fields land on the same
   * `<c:legend>` element so a single configuration call threads
   * cleanly through every legend knob Excel exposes.
   */
  legendStrikethrough?: boolean;
  /**
   * Legend font color. Maps to `<c:legend><c:txPr><a:p><a:pPr>
   * <a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
   * </a:defRPr></a:pPr></a:p></c:txPr></c:legend>` — Excel's "Format
   * Legend -> Font -> Font color" picker. The OOXML `<a:srgbClr val=".."/>`
   * carries the 6-character uppercase hex sRGB color (`CT_SRgbColor`
   * inside `CT_TextCharacterProperties`' fill choice — ECMA-376 Part 1,
   * §20.1.2.3.32 / §21.1.2.3.7); the writer lands the value on the
   * default-paragraph `<a:defRPr>` slot inside the legend's `<c:txPr>`
   * block so a re-parse picks the color up off the canonical slot the
   * OOXML schema exposes.
   *
   * Accepts the color either with or without a leading `#` and in any
   * case — `"FF0000"`, `"#FF0000"`, and `"ff0000"` all collapse to the
   * OOXML uppercase canonical form `"FF0000"`. Malformed inputs (wrong
   * length, non-hex characters, alpha-channel forms like `"#FF0000FF"`,
   * non-string escapes from an untyped caller) collapse to `undefined`
   * so the writer skips the entire `<a:solidFill>` block and the legend
   * inherits the theme text color (Excel's reference behavior for a
   * fresh legend that has not had a custom color picked).
   *
   * Default: omitted — the legend renders at the theme text color (no
   * `<a:solidFill>` block, matching Excel's reference serialization for
   * a fresh chart legend whose typography has not been customized).
   *
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no `<c:txPr>` slot to host the fill in that
   * case. Mirrors {@link titleColor} / {@link axes.x.axisTitleColor} /
   * {@link axes.x.labelColor} — same accept-with-or-without-`#` hex
   * grammar, same OOXML `<a:solidFill><a:srgbClr val=".."/>` mapping —
   * so a caller can thread a single hex string through every
   * typography-pinning slot. Composes independently with
   * {@link legend} / {@link legendOverlay} / {@link legendEntries} /
   * {@link legendFontSize} / {@link legendBold} / {@link legendItalic}
   * / {@link legendUnderline} / {@link legendStrikethrough}: all nine
   * fields land on the same `<c:legend>` element so a single
   * configuration call threads cleanly through every legend knob Excel
   * exposes.
   */
  legendFontColor?: string;
  /**
   * Legend font family / typeface. Maps to `<c:legend><c:txPr><a:p>
   * <a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr></a:pPr>
   * </a:p></c:txPr></c:legend>` — Excel's "Format Legend -> Font ->
   * Font" picker. The OOXML `<a:latin typeface=".."/>` element carries
   * the typeface name (`CT_TextFont`, ECMA-376 Part 1, §21.1.2.3.7);
   * the writer lands the element on the default-paragraph
   * `<a:defRPr>` slot inside the legend's `<c:txPr>` block so a re-
   * parse picks the typeface up off the canonical slot the OOXML
   * schema exposes.
   *
   * Accepts any non-empty string typeface name (e.g. `"Calibri"`,
   * `"Arial"`, `"Times New Roman"`); the writer trims surrounding
   * whitespace and emits the trimmed value verbatim (XML-escaped) so
   * Excel can resolve the named font from the workbook's font scheme
   * or the host system's installed fonts. Empty / whitespace-only
   * strings and non-string tokens collapse to `undefined` so the
   * writer skips the entire `<a:latin>` element and the legend
   * inherits Excel's reference theme typeface.
   *
   * Default: omitted — the legend renders in Excel's reference theme
   * typeface (no `<a:latin>` element, the writer skips the element
   * entirely). Pin a typeface name to render the legend in that font
   * (e.g. `"Arial"` for a corporate dashboard standard).
   *
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no `<c:txPr>` slot to host the typeface in
   * that case. Mirrors {@link titleFontFamily} /
   * {@link axes.x.axisTitleFontFamily} /
   * {@link axes.x.labelFontFamily} — same accept-and-trim grammar,
   * same OOXML `<a:latin typeface=".."/>` mapping — so a caller can
   * thread a single typeface string through every typography-pinning
   * slot. Composes independently with {@link legend} /
   * {@link legendOverlay} / {@link legendEntries} /
   * {@link legendFontSize} / {@link legendBold} /
   * {@link legendItalic} / {@link legendUnderline} /
   * {@link legendStrikethrough} / {@link legendFontColor}: all ten
   * fields land on the same `<c:legend>` element so a single
   * configuration call threads cleanly through every legend knob
   * Excel exposes.
   */
  legendFontFamily?: string;
  /**
   * Custom legend placement inside the chart frame. Maps to
   * `<c:legend><c:layout><c:manualLayout>...</c:manualLayout></c:layout>
   * </c:legend>` — Excel's "Format Legend -> Position -> Custom" knob.
   * The block sits between `<c:legendEntry>` and `<c:overlay>` per
   * `CT_Legend` (ECMA-376 Part 1, §21.2.2.114).
   *
   * Each of {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
   * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} is a
   * fraction of the chart frame in the range `0..1` — `(0, 0)` is the
   * upper-left of the chart frame, `(1, 1)` is the lower-right. The
   * coordinates compose independently with {@link legend} (the legend
   * still picks up its `<c:legendPos>` orientation hint, the manual
   * layout merely overrides where the legend block draws). Out-of-range
   * / non-finite / non-numeric coordinates collapse to omitting the
   * matching `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` slot so a caller can
   * pin only the position ({@link ChartManualLayout.x} /
   * {@link ChartManualLayout.y}) and let the legend keep its automatic
   * size, only the size ({@link ChartManualLayout.w} /
   * {@link ChartManualLayout.h}) and let it keep its automatic anchor,
   * or any combination.
   *
   * The writer always emits the matching `<c:xMode>` / `<c:yMode>` /
   * `<c:wMode>` / `<c:hMode>` children with `val="edge"` (Excel's
   * reference shape when the user drags the legend to a custom
   * position — the coordinates are absolute fractions of the chart
   * frame, not deltas from the auto-layout baseline).
   *
   * Default: omitted — the legend renders at the auto-layout position
   * Excel computes from the chart's dimensions and the resolved
   * `<c:legendPos>` orientation. Pin a {@link ChartManualLayout} to
   * place the legend in a specific quadrant of the chart frame —
   * useful for composing a templated dashboard whose legend needs to
   * align with neighbouring tiles, or a hero chart whose legend should
   * sit in a corner the auto-layout would not pick.
   *
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no slot to host the layout in that case. An
   * empty layout (every coordinate undefined) collapses to omitting the
   * entire `<c:layout>` block so a fresh chart matches Excel's
   * reference serialization byte-for-byte.
   */
  legendLayout?: ChartManualLayout;
  /**
   * Legend background fill color. Maps to `<c:legend><c:spPr>
   * <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></c:spPr>
   * </c:legend>` (CT_Legend, ECMA-376 Part 1, §21.2.2.114) — Excel's
   * "Format Legend -> Fill -> Solid fill -> Color" picker (the same
   * dialog the user reaches by right-clicking the legend background).
   * The element sits between `<c:overlay>` and `<c:txPr>` per the
   * CT_Legend schema sequence — distinct from `legendFontColor`,
   * which lands on the legend's `<c:txPr>` text-character-properties
   * slot.
   *
   * Accepts a 6-character hex sRGB triple with or without a leading
   * `#` (e.g. `"FFFF00"`, `"#FFFF00"`, `"ffff00"`); the writer
   * normalizes to OOXML's canonical 6-character uppercase form before
   * emit. Empty / whitespace-only strings, alpha-channel forms,
   * non-hex characters, and non-string escapes from an untyped caller
   * all collapse to `undefined` so the writer skips the entire
   * `<c:spPr>` block and the legend renders at the theme default fill
   * (Excel's reference behavior for a fresh legend that has not had a
   * custom fill picked — typically a transparent background with no
   * `<c:spPr>` block).
   *
   * Default: omitted — the legend renders at the theme default fill
   * (no `<c:spPr>` block, matching Excel's reference serialization
   * for a fresh chart legend whose background has not been
   * customized). Pin a hex color to render the legend background in
   * that color (e.g. `"FFFF00"` for a highlighted legend on a
   * dashboard tile, or `"FFFFFF"` for an opaque white legend on a
   * non-white plot area).
   *
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no slot to host the fill in that case.
   * Mirrors `plotAreaFillColor` — same accept-with-or-without-`#` hex
   * grammar, same OOXML `<c:spPr><a:solidFill><a:srgbClr val=".."/>
   * </a:solidFill></c:spPr>` mapping — so a caller can thread a
   * single hex string through every `<c:spPr>`-based fill slot.
   * Composes independently with `legend` / `legendOverlay` /
   * `legendEntries` / `legendFontSize` / `legendBold` /
   * `legendItalic` / `legendUnderline` / `legendStrikethrough` /
   * `legendFontColor` / `legendFontFamily` / `legendLayout`: the fill
   * lands on the legend's `<c:spPr>` block, the typography knobs on
   * the legend's `<c:txPr>` block, the layout on the legend's
   * `<c:layout>` block, and the visibility / position toggles on
   * `<c:legendPos>` / `<c:overlay>` — every knob targets a different
   * child of `<c:legend>`.
   */
  legendFillColor?: string;
  /**
   * Legend border (stroke) solid color as a 6-digit RGB hex string
   * (e.g. `"1F77B4"`). Maps to `<c:legend><c:spPr><a:ln><a:solidFill>
   * <a:srgbClr val="RRGGBB"/></a:solidFill></a:ln></c:spPr></c:legend>`
   * (CT_Legend, ECMA-376 Part 1, §21.2.2.114) — Excel's "Format Legend
   * -> Border -> Solid line -> Color" picker. The OOXML `<a:srgbClr
   * val=".."/>` carries the 6-character uppercase hex sRGB color
   * (`CT_SRgbColor` inside the line's solid fill choice — ECMA-376
   * Part 1, §20.1.2.3.32 / §20.1.2.3.24). The `<c:spPr>` slot sits
   * between `<c:overlay>` and `<c:txPr>` per the CT_Legend schema
   * sequence; `<a:ln>` follows the optional `<a:solidFill>` (fill)
   * child inside `<c:spPr>` per `CT_ShapeProperties` (ECMA-376 Part 1,
   * §20.1.2.3.13).
   *
   * Accepts a leading `#` and any case; the writer collapses to the
   * OOXML canonical uppercase form. Malformed inputs (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` and the writer omits
   * the `<a:ln>` block (Excel's reference serialization for a legend
   * that inherits the auto-stroke — typically a translucent gray
   * border or no border depending on the theme).
   *
   * Default: omitted — the legend inherits the auto-stroke Excel picks
   * from the chart's theme. Pin a hex color to mirror Excel's "Format
   * Legend -> Border -> Solid line" knob and paint a flat border around
   * the legend block — useful for dashboard tiles where the legend
   * should be visually framed against the surrounding chart frame, or
   * to highlight a legend against a busy theme.
   *
   * Composes independently with {@link legendFillColor} — the two knobs
   * land on the same `<c:spPr>` block but on different children
   * (`<a:solidFill>` for the fill, `<a:ln>` for the stroke), and the
   * writer authors a `<c:spPr>` whenever either knob is set. A caller
   * can pin one without the other; pinning both produces a filled
   * legend with a colored border.
   *
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no slot to host the stroke in that case.
   * Mirrors `plotAreaBorderColor` — same accept-with-or-without-`#`
   * hex grammar, same `<c:spPr>` host element on a different parent
   * (`<c:legend>` rather than `<c:plotArea>`), but lands on the line
   * (`<a:ln>`) child rather than the fill (`<a:solidFill>`) child.
   *
   * Patterned / gradient strokes are not modelled — only the solid
   * sRGB form lands on the wire. Theme-color references
   * (`<a:schemeClr>`) likewise drop to `undefined` so a parsed value
   * always carries a literal hex Excel will render byte-for-byte.
   */
  legendBorderColor?: string;
  /**
   * Legend border (stroke) thickness in points (e.g. `1.5`). Maps to
   * the `w` attribute on `<c:legend><c:spPr><a:ln w="EMU">` — Excel's
   * "Format Legend -> Border -> Width" spinner. The OOXML `w`
   * attribute carries the stroke width in English Metric Units
   * (`1 pt = 12 700 EMU`) per `CT_LineProperties` (ECMA-376 Part 1,
   * §20.1.2.3.24). The writer multiplies the input by 12 700 and rounds
   * to the nearest integer because the schema types `w` as `xsd:int`.
   *
   * Accepts any finite number; values are clamped to the `0.25..13.5`
   * pt band Excel's UI exposes (the same band used by the series
   * stroke knob `series[i].stroke.width` and the plot-area border
   * width knob {@link plotAreaBorderWidth}) and snapped to the
   * 0.25 pt grid so a parsed-then-written width does not drift across
   * round-trips. Non-finite / non-numeric / `NaN` values collapse to
   * `undefined` and the writer omits the `w` attribute (the line keeps
   * Excel's auto-thickness — typically 0.75 pt).
   *
   * Default: omitted — the legend border inherits Excel's
   * auto-thickness. Pin a value to thicken the border around the legend
   * block (a hairline at 0.25 pt, a heavy frame at several points) or
   * to match the stroke width of neighboring chart tiles.
   *
   * Composes independently with {@link legendBorderColor} — the width
   * attribute lands on the same `<a:ln>` element as the color's
   * `<a:solidFill>` child, but the writer authors `<a:ln>` whenever
   * either knob is set. A caller can pin a width without a color (the
   * border picks Excel's auto-color), pin a color without a width (the
   * border picks Excel's auto-thickness), or pin both. Setting only the
   * width is valid Excel UI — the user can drag the "Width" spinner
   * without picking a custom color.
   *
   * Silently ignored when `legend === false` (no `<c:legend>` element
   * is emitted) — there is no slot to host the stroke in that case.
   * Mirrors {@link plotAreaBorderWidth} — same accept-finite-number
   * grammar, same `<a:ln w="EMU">` host attribute on a different parent
   * (`<c:legend>` rather than `<c:plotArea>`), and shares the same
   * 0.25..13.5 pt clamp + 0.25 pt snap so width values compose the same
   * way at the call site.
   */
  legendBorderWidth?: number;
  /**
   * Chart legend border (stroke) preset dash pattern. Maps to the
   * `val` attribute on `<c:legend><c:spPr><a:ln><a:prstDash val=".."/>`.
   * Mirrors Excel's "Format Legend -> Border -> Dash type" picker. Same
   * {@link ChartBorderDash} accept-or-drop grammar as
   * {@link plotAreaBorderDash} — `"solid"` collapses to `undefined`
   * so absence and the OOXML default round-trip identically.
   *
   * Composes independently with {@link legendBorderColor} and
   * {@link legendBorderWidth} — all three knobs share the same `<a:ln>`
   * element. Silently ignored when `legend === false` (no `<c:legend>`
   * element is emitted).
   */
  legendBorderDash?: ChartBorderDash;
  /**
   * Custom plot-area placement inside the chart frame. Maps to
   * `<c:plotArea><c:layout><c:manualLayout>...</c:manualLayout></c:layout>
   * </c:plotArea>` — Excel's "Format Plot Area -> Position -> Custom"
   * placement (the same dialog the user gets by dragging the plot area's
   * border). The block is the first child of `<c:plotArea>` per
   * `CT_PlotArea` (ECMA-376 Part 1, §21.2.2.145), preceding every chart-type
   * element and the axes.
   *
   * Each of {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
   * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} is a
   * fraction of the chart frame in the range `0..1` — `(0, 0)` is the
   * upper-left of the chart frame, `(1, 1)` is the lower-right. The
   * coordinates compose independently — out-of-range / non-finite /
   * non-numeric coordinates collapse to omitting the matching `<c:x>` /
   * `<c:y>` / `<c:w>` / `<c:h>` slot so a caller can pin only the
   * position ({@link ChartManualLayout.x} / {@link ChartManualLayout.y})
   * and let the plot area keep its automatic size, only the size
   * ({@link ChartManualLayout.w} / {@link ChartManualLayout.h}) and let
   * it keep its automatic anchor, or any combination.
   *
   * The writer always emits the matching `<c:xMode>` / `<c:yMode>` /
   * `<c:wMode>` / `<c:hMode>` children with `val="edge"` (Excel's
   * reference shape when the user drags the plot area to a custom
   * position — the coordinates are absolute fractions of the chart
   * frame, not deltas from the auto-layout baseline).
   *
   * Default: omitted — the plot area renders at the auto-layout
   * position Excel computes from the chart's dimensions, the title /
   * legend slots, and the axis label widths. Pin a {@link ChartManualLayout}
   * to align the plot area with neighbouring chart tiles, reserve a
   * fixed margin for a wrapped axis label, or hand off vertical space
   * to a tall legend.
   *
   * The writer always emits a `<c:layout>` element on `<c:plotArea>` —
   * even when {@link plotAreaLayout} is omitted — because Excel's
   * reference serialization always includes the (empty) auto-layout
   * placeholder. When the field is omitted (or every coordinate dropped
   * on normalization), the writer emits the bare `<c:layout/>`
   * placeholder so a fresh chart matches Excel's reference shape
   * byte-for-byte; pinning at least one coordinate replaces the
   * placeholder with `<c:layout><c:manualLayout>...</c:manualLayout>
   * </c:layout>`.
   */
  plotAreaLayout?: ChartManualLayout;
  /**
   * Plot-area solid fill color as a 6-digit RGB hex string (e.g.
   * `"F2F2F2"`). Maps to `<c:plotArea><c:spPr><a:solidFill>
   * <a:srgbClr val="RRGGBB"/></a:solidFill></c:spPr></c:plotArea>` —
   * Excel's "Format Plot Area -> Fill -> Solid fill -> Color" picker.
   * The OOXML `<a:srgbClr val=".."/>` carries the 6-character uppercase
   * hex sRGB color (`CT_SRgbColor` inside `CT_ShapeProperties`' fill
   * choice — ECMA-376 Part 1, §20.1.2.3.32 / §20.1.2.3.13). The
   * `<c:spPr>` slot lives at the tail of `<c:plotArea>` per
   * `CT_PlotArea` (ECMA-376 Part 1, §21.2.2.145), after every chart-type
   * element / axes / `<c:dTable>`.
   *
   * Accepts a leading `#` and any case; the writer collapses to the
   * OOXML canonical uppercase form. Malformed inputs (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` and the writer omits
   * the entire `<c:spPr>` block (Excel's reference serialization for
   * a plot area that inherits the auto-fill — typically the chart-frame
   * background or a translucent white depending on the theme).
   *
   * Default: omitted — the plot area inherits the auto-fill Excel
   * picks from the chart's theme. Pin a hex color to mirror Excel's
   * "Format Plot Area -> Fill -> Solid fill" knob and paint a flat
   * background behind the series — useful for dashboard tiles where
   * the plot area should pick up a brand color or a contrast band
   * against the surrounding chart frame.
   *
   * Patterned / gradient / picture fills are not modelled — only the
   * solid sRGB form lands on the wire. Theme-color references
   * (`<a:schemeClr>`) likewise drop to `undefined` so a parsed value
   * always carries a literal hex Excel will render byte-for-byte.
   */
  plotAreaFillColor?: string;
  /**
   * Plot-area border (stroke) solid color as a 6-digit RGB hex string
   * (e.g. `"1F77B4"`). Maps to `<c:plotArea><c:spPr><a:ln><a:solidFill>
   * <a:srgbClr val="RRGGBB"/></a:solidFill></a:ln></c:spPr></c:plotArea>`
   * — Excel's "Format Plot Area -> Border -> Solid line -> Color"
   * picker. The OOXML `<a:srgbClr val=".."/>` carries the 6-character
   * uppercase hex sRGB color (`CT_SRgbColor` inside the line's solid
   * fill choice — ECMA-376 Part 1, §20.1.2.3.32 / §20.1.2.3.24). The
   * `<c:spPr>` slot lives at the tail of `<c:plotArea>` per
   * `CT_PlotArea` (ECMA-376 Part 1, §21.2.2.145), after every chart-type
   * element / axes / `<c:dTable>`; `<a:ln>` follows the optional
   * `<a:solidFill>` (fill) child inside `<c:spPr>` per
   * `CT_ShapeProperties` (ECMA-376 Part 1, §20.1.2.3.13).
   *
   * Accepts a leading `#` and any case; the writer collapses to the
   * OOXML canonical uppercase form. Malformed inputs (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` and the writer omits
   * the `<a:ln>` block (Excel's reference serialization for a plot
   * area that inherits the auto-stroke — typically a translucent gray
   * border or no border depending on the theme).
   *
   * Default: omitted — the plot area inherits the auto-stroke Excel
   * picks from the chart's theme. Pin a hex color to mirror Excel's
   * "Format Plot Area -> Border -> Solid line" knob and paint a flat
   * border around the series band — useful for dashboard tiles where
   * the plot area should be visually framed against the surrounding
   * chart frame, or to highlight a series band against a busy theme.
   *
   * Composes independently with {@link plotAreaFillColor} — the two
   * knobs land on the same `<c:spPr>` block but on different children
   * (`<a:solidFill>` for the fill, `<a:ln>` for the stroke), and the
   * writer authors a `<c:spPr>` whenever either knob is set. A caller
   * can pin one without the other; pinning both produces a filled
   * plot area with a colored border.
   *
   * Patterned / gradient strokes are not modelled — only the solid
   * sRGB form lands on the wire. Theme-color references
   * (`<a:schemeClr>`) likewise drop to `undefined` so a parsed value
   * always carries a literal hex Excel will render byte-for-byte.
   * Mirrors `plotAreaFillColor` — same accept-with-or-without-`#` hex
   * grammar, same `<c:spPr>` host element, but lands on the line
   * (`<a:ln>`) child rather than the fill (`<a:solidFill>`) child.
   */
  plotAreaBorderColor?: string;
  /**
   * Plot-area border (stroke) thickness in points (e.g. `1.5`). Maps to
   * the `w` attribute on `<c:plotArea><c:spPr><a:ln w="EMU">` — Excel's
   * "Format Plot Area -> Border -> Width" spinner. The OOXML `w`
   * attribute carries the stroke width in EMU (English Metric Units;
   * `1 pt = 12 700 EMU`) per `CT_LineProperties` (ECMA-376 Part 1,
   * §20.1.2.3.24). The writer multiplies the input by 12 700 and rounds
   * to the nearest integer because the schema types `w` as `xsd:int`.
   *
   * Accepts any finite number; values are clamped to the `0.25..13.5`
   * pt band Excel's UI exposes (the same band used by the series
   * stroke knob `series[i].stroke.width`) and snapped to the
   * 0.25 pt grid so a parsed-then-written width does not drift across
   * round-trips. Non-finite / non-numeric / `NaN` values collapse to
   * `undefined` and the writer omits the `w` attribute (the line keeps
   * Excel's auto-thickness — typically 0.75 pt).
   *
   * Default: omitted — the plot-area border inherits Excel's
   * auto-thickness. Pin a value to thicken the border around the series
   * band (a hairline at 0.25 pt, a heavy frame at several points) or to
   * match the stroke width of neighboring chart tiles.
   *
   * Composes independently with {@link plotAreaBorderColor} — the
   * width attribute lands on the same `<a:ln>` element as the color's
   * `<a:solidFill>` child, but the writer authors `<a:ln>` whenever
   * either knob is set. A caller can pin a width without a color (the
   * border picks Excel's auto-color), pin a color without a width (the
   * border picks Excel's auto-thickness), or pin both. Setting only the
   * width is valid Excel UI — the user can drag the "Width" spinner
   * without picking a custom color.
   *
   * Mirrors the series-line stroke width grammar (`series[i].stroke.width`)
   * and shares the same clamp / snap behavior so width values compose
   * the same way at the call site. The OOXML `w` attribute is also used
   * for axis lines and other strokes — this knob covers only the
   * plot-area border slot.
   */
  plotAreaBorderWidth?: number;
  /**
   * Plot-area border (stroke) preset dash pattern. Maps to the `val`
   * attribute on `<c:plotArea><c:spPr><a:ln><a:prstDash val=".."/>` —
   * Excel's "Format Plot Area -> Border -> Dash type" picker. The OOXML
   * `<a:prstDash>` element carries the `ST_PresetLineDashVal` enum on
   * `CT_PresetLineDashProperties` (ECMA-376 Part 1, §20.1.8.48). Per
   * `CT_LineProperties` schema sequence (§20.1.2.3.24), `<a:prstDash>`
   * follows `<a:solidFill>` and precedes `<a:headEnd>` / `<a:tailEnd>`.
   *
   * Accepts any {@link ChartBorderDash} value; `"solid"` (the OOXML
   * default) collapses to `undefined` for round-trip symmetry — the
   * writer skips the `<a:prstDash>` element entirely so a fresh plot
   * area matches Excel's reference shape byte-for-byte. Unrecognized
   * tokens drop to `undefined` rather than fabricate a value Excel
   * would reject.
   *
   * Default: omitted — the border renders solid. Pin a value to dash /
   * dot the plot-area border, useful for distinguishing the inner band
   * from the outer chart frame.
   *
   * Composes independently with {@link plotAreaBorderColor} and
   * {@link plotAreaBorderWidth} — all three knobs land on the same
   * `<a:ln>` element but on different children / attributes (the color
   * `<a:solidFill>` child, the `w` attribute, and the dash
   * `<a:prstDash>` child); the writer authors `<a:ln>` whenever any
   * of them is set.
   */
  plotAreaBorderDash?: ChartBorderDash;
  /**
   * Chart-space (entire chart background) solid fill color as a 6-digit
   * RGB hex string (e.g. `"F2F2F2"`). Maps to `<c:chartSpace><c:spPr>
   * <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></c:spPr>
   * </c:chartSpace>` (CT_ChartSpace, ECMA-376 Part 1, §21.2.2.29) —
   * Excel's "Format Chart Area -> Fill -> Solid fill -> Color" picker
   * (the same dialog the user reaches by right-clicking the chart's
   * outer frame). The OOXML `<a:srgbClr val=".."/>` carries the
   * 6-character uppercase hex sRGB color (`CT_SRgbColor` inside
   * `CT_ShapeProperties`' fill choice — ECMA-376 Part 1, §20.1.2.3.32 /
   * §20.1.2.3.13). The `<c:spPr>` slot lives at the tail of
   * `<c:chartSpace>` per CT_ChartSpace, after `<c:chart>` /
   * `<c:externalData>` / `<c:printSettings>` / `<c:userShapes>` and
   * before the optional `<c:txPr>` / `<c:extLst>`.
   *
   * Distinct from {@link plotAreaFillColor} — the plot area is the
   * inner band that hosts the series, while chartSpace covers the full
   * frame including the title slot, the legend slot, the axis label
   * margins, and the plot area itself. A caller can pin both knobs
   * (e.g. an off-white frame and a brand-color plot area) since the
   * two `<c:spPr>` blocks land on different host elements.
   *
   * Accepts a leading `#` and any case; the writer collapses to the
   * OOXML canonical uppercase form. Malformed inputs (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` and the writer omits
   * the entire `<c:spPr>` block (Excel's reference serialization for
   * a chart that inherits the auto-fill — typically opaque white from
   * the workbook theme).
   *
   * Default: omitted — the chart inherits the auto-fill Excel picks
   * from the workbook theme. Pin a hex color to mirror Excel's
   * "Format Chart Area -> Fill -> Solid fill" knob and paint a flat
   * background behind the entire chart — useful for dashboard tiles
   * where the chart frame should pick up a brand color or contrast
   * with the surrounding sheet cells.
   *
   * Patterned / gradient / picture fills are not modelled — only the
   * solid sRGB form lands on the wire. Theme-color references
   * (`<a:schemeClr>`) likewise drop to `undefined` so a parsed value
   * always carries a literal hex Excel will render byte-for-byte.
   * Mirrors `plotAreaFillColor` / `legendFillColor` / `titleFillColor`
   * — same accept-with-or-without-`#` hex grammar, same OOXML
   * `<c:spPr><a:solidFill><a:srgbClr val=".."/></a:solidFill>
   * </c:spPr>` mapping — so a caller can thread a single hex string
   * through every `<c:spPr>`-based fill slot.
   */
  chartSpaceFillColor?: string;
  /**
   * Chart-space (entire chart frame) border (stroke) solid color as a
   * 6-digit RGB hex string (e.g. `"1F77B4"`). Maps to `<c:chartSpace>
   * <c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
   * </a:ln></c:spPr></c:chartSpace>` — Excel's "Format Chart Area ->
   * Border -> Solid line -> Color" picker (the same dialog the user
   * reaches by right-clicking the chart's outer frame). The OOXML
   * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
   * sRGB color (`CT_SRgbColor` inside the line's solid fill choice —
   * ECMA-376 Part 1, §20.1.2.3.32 / §20.1.2.3.24). The `<c:spPr>`
   * slot lives at the tail of `<c:chartSpace>` per CT_ChartSpace
   * (§21.2.2.29); `<a:ln>` follows the optional `<a:solidFill>` (fill)
   * child inside `<c:spPr>` per `CT_ShapeProperties` (§20.1.2.3.13).
   *
   * Accepts a leading `#` and any case; the writer collapses to the
   * OOXML canonical uppercase form. Malformed inputs (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` and the writer omits
   * the `<a:ln>` block (Excel's reference serialization for a chart
   * area that inherits the auto-stroke — typically a translucent gray
   * border or no border depending on the theme).
   *
   * Default: omitted — the chart inherits the auto-stroke Excel picks
   * from the workbook theme. Pin a hex color to mirror Excel's
   * "Format Chart Area -> Border -> Solid line" knob and paint a flat
   * border around the entire chart frame — useful for dashboard tiles
   * where the chart should be visually framed against the surrounding
   * sheet cells.
   *
   * Composes independently with {@link chartSpaceFillColor} — the two
   * knobs land on the same `<c:spPr>` block but on different children
   * (`<a:solidFill>` for the fill, `<a:ln>` for the stroke), and the
   * writer authors a `<c:spPr>` whenever either knob is set. A caller
   * can pin one without the other; pinning both produces a filled
   * chart frame with a colored border.
   *
   * Patterned / gradient strokes are not modelled — only the solid
   * sRGB form lands on the wire. Theme-color references
   * (`<a:schemeClr>`) likewise drop to `undefined` so a parsed value
   * always carries a literal hex Excel will render byte-for-byte.
   * Mirrors `plotAreaBorderColor` / `legendBorderColor` /
   * `titleBorderColor` — same accept-with-or-without-`#` hex grammar,
   * same OOXML `<a:ln><a:solidFill><a:srgbClr val=".."/>
   * </a:solidFill></a:ln>` mapping — so a caller can thread a single
   * hex string through every `<a:ln>`-based stroke slot.
   */
  chartSpaceBorderColor?: string;
  /**
   * Chart-space (entire chart frame) border (stroke) thickness in
   * points (e.g. `1.5`). Maps to the `w` attribute on
   * `<c:chartSpace><c:spPr><a:ln w="EMU">` — Excel's "Format Chart
   * Area -> Border -> Width" spinner (the same dialog the user reaches
   * by right-clicking the chart's outer frame). The OOXML `w`
   * attribute carries the stroke width in English Metric Units
   * (1 pt = 12 700 EMU) per `CT_LineProperties` (ECMA-376 Part 1,
   * §20.1.2.3.24). The writer multiplies the input by 12 700 and rounds
   * to the nearest integer because the schema types `w` as `xsd:int`.
   *
   * Accepts any finite number; values are clamped to the `0.25..13.5`
   * pt band Excel's UI exposes (the same band used by
   * {@link plotAreaBorderWidth} / {@link legendBorderWidth} /
   * {@link titleBorderWidth} / the series stroke knob) and snapped to
   * the 0.25 pt grid so a parsed-then-written width does not drift
   * across round-trips. Non-finite / non-numeric / `NaN` values
   * collapse to `undefined` and the writer omits the `w` attribute
   * (the line keeps Excel's auto-thickness — typically 0.75 pt).
   *
   * Default: omitted — the chart-frame border inherits Excel's
   * auto-thickness.
   *
   * Composes independently with {@link chartSpaceBorderColor} — the
   * width attribute lands on the same `<a:ln>` element as the color's
   * `<a:solidFill>` child, but the writer authors `<a:ln>` whenever
   * either knob is set. A caller can pin a width without a color (the
   * border picks Excel's auto-color), pin a color without a width (the
   * border picks Excel's auto-thickness), or pin both. Setting only
   * the width is valid Excel UI — the user can drag the "Width"
   * spinner without picking a custom color.
   *
   * Mirrors {@link plotAreaBorderWidth} / {@link legendBorderWidth} /
   * {@link titleBorderWidth} — same accept-finite-number / clamp /
   * snap grammar, same `<a:ln w="EMU">` host attribute on a different
   * parent (`<c:chartSpace>` rather than `<c:plotArea>` / `<c:legend>`
   * / `<c:title>`).
   */
  chartSpaceBorderWidth?: number;
  /**
   * Chart-space (entire chart frame) border (stroke) preset dash
   * pattern. Maps to the `val` attribute on `<c:chartSpace><c:spPr>
   * <a:ln><a:prstDash val=".."/>`. Mirrors Excel's "Format Chart Area
   * -> Border -> Dash type" picker. Same {@link ChartBorderDash}
   * accept-or-drop grammar as {@link plotAreaBorderDash} — `"solid"`
   * collapses to `undefined` so absence and the OOXML default round-
   * trip identically.
   *
   * Composes independently with {@link chartSpaceBorderColor} and
   * {@link chartSpaceBorderWidth} — all three knobs share the same
   * `<a:ln>` element. Mirrors {@link plotAreaBorderDash} and lands on
   * `<c:chartSpace>`'s own `<c:spPr>` block.
   */
  chartSpaceBorderDash?: ChartBorderDash;
  /** Show the chart-level title element. Default: `true` when `title` is set. */
  showTitle?: boolean;
  /**
   * Whether the chart title overlaps the plot area. Maps to
   * `<c:title><c:overlay val=".."/></c:title>` — Excel's "Format Chart
   * Title -> Show the title without overlapping the chart" toggle (the
   * checkbox is the inverse of this flag — checked means `false`,
   * unchecked means `true`). Default: `false` (the OOXML default Excel
   * itself emits) — the title reserves its own slot above the plot area
   * and the plot area shrinks to accommodate it. Set `true` to draw the
   * title on top of the plot area so the chart series get the full frame.
   *
   * Silently ignored when no title is rendered (`showTitle === false` or
   * `title` is absent) — there is no `<c:title>` element to host the
   * overlay flag in either case. Independent of {@link legendOverlay}:
   * the legend's `<c:overlay>` lives on `<c:legend>`, while this one
   * lives on `<c:title>`, so the two flags compose freely.
   */
  titleOverlay?: boolean;
  /**
   * Chart title rotation in whole degrees, measured clockwise from the
   * normal horizontal baseline. Maps to `<c:title><c:tx><c:rich>
   * <a:bodyPr rot="N"/></c:rich></c:tx></c:title>` — Excel's "Format
   * Chart Title -> Size & Properties -> Alignment -> Custom angle" knob.
   * The OOXML attribute is in 60000ths of a degree, so 45° serializes
   * as `rot="2700000"` and -90° as `rot="-5400000"`; the writer
   * performs the conversion at emit time.
   *
   * Accepted range: `-90..90` (Excel's UI band). Out-of-range inputs
   * clamp to the nearest endpoint; non-integer inputs round to the
   * nearest whole degree (the OOXML attribute is an integer in
   * 60000ths, so a fractional whole-degree value has no meaningful
   * refinement at emit time). `0`, `NaN`, `Infinity`, and non-numeric
   * inputs collapse to `undefined` so the writer falls back to the
   * default horizontal orientation.
   *
   * Default: omitted — the title renders horizontally (the OOXML
   * default `rot="0"` Excel itself emits on a fresh chart). Set a
   * non-zero value to tilt the title diagonally or stand it on its
   * side, useful when composing a dashboard whose chart titles need to
   * fit a tight vertical sidebar slot or to mirror the rotation of an
   * underlying axis label set.
   *
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * rotation in either case. Mirrors the axis-side
   * {@link SheetChart.axes.x.labelRotation} field — same range, same
   * normalization, same OOXML conversion factor — so a caller can
   * thread a single rotation value through both the chart title and an
   * axis label set without bookkeeping the units.
   */
  titleRotation?: number;
  /**
   * Chart title font size in whole or half points. Maps to
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
   * <a:r><a:rPr sz="N"/></a:r></a:p></c:rich></c:tx></c:title>` —
   * Excel's "Format Chart Title -> Font -> Size" knob. The OOXML
   * attribute is in 100ths of a point, so 18pt serializes as
   * `sz="1800"` and 14pt (Excel's reference default) as `sz="1400"`;
   * the writer performs the conversion at emit time and lands the
   * value on both the default-paragraph `<a:defRPr>` and the literal
   * run's `<a:rPr>` so a re-parse picks the size up off either
   * canonical slot.
   *
   * Accepted range: `1..400`pt (the band the OOXML `ST_TextFontSize`
   * schema exposes — `100..400000` in 100ths of a point). Fractional
   * inputs round to the nearest 0.5pt (Excel's UI granularity); inputs
   * outside the band, `NaN`, `Infinity`, and non-numeric inputs all
   * collapse to `undefined` so the writer falls back to the default
   * `14`pt size Excel itself emits on a fresh chart.
   *
   * Default: omitted — the title renders at Excel's default 14pt. Set
   * an explicit value to scale the title up for a hero dashboard tile
   * (e.g. `24`) or down to fit a tight header slot (e.g. `10`).
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * size in either case.
   */
  titleFontSize?: number;
  /**
   * Chart title bold flag. Maps to
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
   * <a:r><a:rPr b=".."/></a:r></a:p></c:rich></c:tx></c:title>` —
   * Excel's "Format Chart Title -> Font -> Bold" toggle. The OOXML
   * attribute is the `xsd:boolean` `b` on `CT_TextCharacterProperties`
   * (ECMA-376 Part 1, §21.1.2.3.7); the writer lands the value on both
   * the default-paragraph `<a:defRPr>` and the literal run's `<a:rPr>`
   * so a re-parse picks the flag up off either canonical slot — Excel
   * keeps the two attributes in sync.
   *
   * Default: omitted — the title renders non-bold (`b="0"`, Excel's
   * reference serialization for a fresh chart title). Set `true` to
   * emit `b="1"` on both slots so the title renders bold; set `false`
   * explicitly to pin the non-default `b="0"` (functionally identical
   * to omission, but useful when overriding a templated title that
   * had bold pinned upstream).
   *
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * flag in either case. Composes independently with
   * {@link titleFontSize} / {@link titleRotation} / {@link titleOverlay}:
   * all four fields land on the same `<c:title>` element so a single
   * configuration call threads cleanly through every chart-title knob
   * Excel exposes.
   */
  titleBold?: boolean;
  /**
   * Chart title italic flag. Maps to
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
   * <a:r><a:rPr i=".."/></a:r></a:p></c:rich></c:tx></c:title>` —
   * Excel's "Format Chart Title -> Font -> Italic" toggle. The OOXML
   * attribute is the `xsd:boolean` `i` on `CT_TextCharacterProperties`
   * (ECMA-376 Part 1, §21.1.2.3.7); the writer lands the value on both
   * the default-paragraph `<a:defRPr>` and the literal run's `<a:rPr>`
   * so a re-parse picks the flag up off either canonical slot — Excel
   * keeps the two attributes in sync.
   *
   * Default: omitted — the title renders non-italic (no `i` attribute,
   * Excel's reference serialization for a fresh chart title; the
   * application-default `false` collapses to absence). Set `true` to
   * emit `i="1"` on both slots so the title renders italic; set `false`
   * explicitly to pin the non-default `i="0"` (functionally identical
   * to omission, but useful when overriding a templated title that had
   * italic pinned upstream).
   *
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * flag in either case. Composes independently with {@link titleBold}
   * / {@link titleFontSize} / {@link titleRotation} /
   * {@link titleOverlay}: all five fields land on the same
   * `<c:title>` element so a single configuration call threads cleanly
   * through every chart-title knob Excel exposes.
   */
  titleItalic?: boolean;
  /**
   * Chart title font color. Maps to
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
   * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr>
   * <a:r><a:rPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
   * </a:rPr></a:r></a:p></c:rich></c:tx></c:title>` — Excel's
   * "Format Chart Title -> Font -> Font Color" picker. The OOXML
   * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
   * sRGB color (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32); the
   * writer lands the fill on both the default-paragraph `<a:defRPr>`
   * and the literal run's `<a:rPr>` so a re-parse picks the color up
   * off either canonical slot — Excel keeps the two values in sync.
   *
   * Accepts the standard 6-character hex string with or without a
   * leading `#` (`"FF0000"` / `"#FF0000"` / `"ff0000"`); the writer
   * normalizes to the OOXML uppercase canonical form
   * (`<a:srgbClr val="FF0000"/>`). The 8-character `#RRGGBBAA` form
   * is *not* accepted — alpha lives on `<a:srgbClr><a:alpha val=".."/>`
   * which is a separate runs-level knob; pinning `titleColor` carries
   * the RGB triple only.
   *
   * Default: omitted — the title renders in Excel's reference
   * inherited theme color (no `<a:solidFill>` element, the writer
   * skips the fill block entirely). Pin a hex value to render the
   * title in that color (e.g. `"1070CA"` for the dashboard hero blue
   * the issue-#136 example reaches for). Malformed inputs (wrong
   * length, non-hex characters, alpha-channel form) collapse to
   * `undefined` so a stray non-hex token never produces a malformed
   * `<a:srgbClr>`.
   *
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * fill in either case. Composes independently with {@link titleBold}
   * / {@link titleItalic} / {@link titleFontSize} /
   * {@link titleRotation} / {@link titleOverlay}: all six fields land
   * on the same `<c:title>` element so a single configuration call
   * threads cleanly through every chart-title knob Excel exposes.
   */
  titleColor?: string;
  /**
   * Chart title strikethrough flag. Maps to
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
   * <a:r><a:rPr strike=".."/></a:r></a:p></c:rich></c:tx></c:title>` —
   * Excel's "Format Chart Title -> Font -> Strikethrough" toggle. The
   * OOXML attribute is the `ST_TextStrikeType` enum on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
   * three values: `"noStrike"` (the OOXML default — no strikethrough),
   * `"sngStrike"` (single horizontal line, the value Excel's UI
   * checkbox emits), and `"dblStrike"` (double horizontal line, a
   * non-UI variant Excel does not surface in its ribbon). The writer
   * lands the value on both the default-paragraph `<a:defRPr>` and
   * the literal run's `<a:rPr>` so a re-parse picks the flag up off
   * either canonical slot — Excel keeps the two attributes in sync.
   *
   * Modeled as a boolean for symmetry with {@link titleBold} /
   * {@link titleItalic}: `true` emits `strike="sngStrike"` (Excel's
   * UI "Strikethrough" checkbox — single line). Absence and
   * non-boolean tokens collapse to omitting the attribute (Excel's
   * reference serialization for a non-strikethrough title — the
   * application-default `"noStrike"` collapses to absence). Set
   * `false` explicitly to pin the non-default `strike="noStrike"`
   * (functionally identical to omission, but useful when overriding a
   * templated title that had strikethrough pinned upstream).
   *
   * Hucre's writer emits only `"sngStrike"` to keep the surfaced shape
   * consistent with what Excel's reference UI authors. The reader
   * collapses the non-UI `"dblStrike"` to `undefined` so a templated
   * chart that pinned the double-line variant in raw OOXML round-trips
   * to the same `undefined` an unmarked chart parses to (i.e. the
   * double-line variant silently downgrades to the single-line write
   * grammar rather than fabricate a value the writer would re-emit
   * incorrectly).
   *
   * Default: omitted — the title renders without strikethrough (no
   * `strike` attribute on either slot, Excel's reference serialization
   * for a fresh chart title). Set `true` to render the title with a
   * single strikethrough line, useful for marking "before / after" or
   * "deprecated" dashboard tile headers.
   *
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * flag in either case. Composes independently with {@link titleBold}
   * / {@link titleItalic} / {@link titleColor} /
   * {@link titleFontSize} / {@link titleRotation} /
   * {@link titleOverlay}: all seven fields land on the same
   * `<c:title>` element so a single configuration call threads
   * cleanly through every chart-title knob Excel exposes.
   */
  titleStrike?: boolean;
  /**
   * Chart title underline flag. Maps to
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
   * <a:r><a:rPr u=".."/></a:r></a:p></c:rich></c:tx></c:title>` —
   * Excel's "Format Chart Title -> Font -> Underline" picker. The
   * OOXML attribute is the `ST_TextUnderlineType` enum on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
   * eighteen values; the two Excel surfaces in the ribbon are
   * `"sng"` (single-line underline — the default checkbox) and
   * `"dbl"` (double-line underline). The remaining sixteen tokens
   * (`"none"`, `"words"`, `"heavy"`, `"dotted"`, `"dottedHeavy"`,
   * `"dash"`, `"dashHeavy"`, `"dashLong"`, `"dashLongHeavy"`,
   * `"dotDash"`, `"dotDashHeavy"`, `"dotDotDash"`,
   * `"dotDotDashHeavy"`, `"wavy"`, `"wavyHeavy"`, `"wavyDbl"`) are
   * non-UI variants Excel does not surface in its ribbon. The writer
   * lands the value on both the default-paragraph `<a:defRPr>` and
   * the literal run's `<a:rPr>` so a re-parse picks the flag up off
   * either canonical slot — Excel keeps the two attributes in sync.
   *
   * Modeled as a boolean for symmetry with {@link titleBold} /
   * {@link titleItalic} / {@link titleStrike}: `true` emits `u="sng"`
   * (Excel's UI "Underline" checkbox — single line). Absence and
   * non-boolean tokens collapse to omitting the attribute (Excel's
   * reference serialization for a non-underlined title — the
   * application-default `"none"` collapses to absence). Set `false`
   * explicitly to pin the non-default omission (functionally
   * identical to omission, but useful when overriding a templated
   * title that had underline pinned upstream).
   *
   * Hucre's writer emits only `"sng"` to keep the surfaced shape
   * consistent with what Excel's reference UI authors. The reader
   * collapses every non-`"sng"` token (the non-UI `"dbl"` variant
   * and the sixteen exotic types) to `undefined` so a templated
   * chart that pinned a non-single underline in raw OOXML round-trips
   * to the same `undefined` an unmarked chart parses to (i.e. the
   * exotic underline silently downgrades to the single-line write
   * grammar rather than fabricate a value the writer would re-emit
   * incorrectly).
   *
   * Default: omitted — the title renders without underline (no `u`
   * attribute on either slot, Excel's reference serialization for a
   * fresh chart title). Set `true` to render the title with a single
   * underline, useful for emphasising "key metric" or "section header"
   * dashboard tile titles.
   *
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * flag in either case. Composes independently with {@link titleBold}
   * / {@link titleItalic} / {@link titleStrike} / {@link titleColor} /
   * {@link titleFontSize} / {@link titleRotation} /
   * {@link titleOverlay}: all eight fields land on the same
   * `<c:title>` element so a single configuration call threads
   * cleanly through every chart-title knob Excel exposes.
   */
  titleUnderline?: boolean;
  /**
   * Chart title font family / typeface. Maps to
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
   * typeface=".."/></a:defRPr></a:pPr><a:r><a:rPr><a:latin
   * typeface=".."/></a:rPr></a:r></a:p></c:rich></c:tx></c:title>` —
   * Excel's "Format Chart Title -> Font -> Font" picker. The OOXML
   * `<a:latin typeface=".."/>` element carries the typeface name
   * (`CT_TextFont`, ECMA-376 Part 1, §21.1.2.3.7); the writer lands
   * the element on both the default-paragraph `<a:defRPr>` and the
   * literal run's `<a:rPr>` so a re-parse picks the typeface up off
   * either canonical slot — Excel keeps the two values in sync.
   *
   * Accepts any non-empty string typeface name (e.g. `"Calibri"`,
   * `"Arial"`, `"Times New Roman"`); the writer trims surrounding
   * whitespace and emits the trimmed value verbatim (XML-escaped)
   * so Excel can resolve the named font from the workbook's font
   * scheme or the host system's installed fonts. Empty / whitespace-
   * only strings and non-string tokens collapse to `undefined` so
   * the writer skips the entire `<a:latin>` element and the title
   * inherits Excel's reference theme typeface (Calibri Light from
   * the default Office theme).
   *
   * Default: omitted — the title renders in Excel's reference theme
   * typeface (no `<a:latin>` element, the writer skips the element
   * entirely). Pin a typeface name to render the title in that font
   * (e.g. `"Arial"` for a corporate dashboard standard).
   *
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * typeface in either case. Composes independently with
   * {@link titleBold} / {@link titleItalic} / {@link titleStrike} /
   * {@link titleUnderline} / {@link titleColor} /
   * {@link titleFontSize} / {@link titleRotation} /
   * {@link titleOverlay}: all nine fields land on the same
   * `<c:title>` element so a single configuration call threads
   * cleanly through every chart-title knob Excel exposes.
   */
  titleFontFamily?: string;
  /**
   * Custom chart-title placement inside the chart frame. Maps to
   * `<c:title><c:layout><c:manualLayout>...</c:manualLayout></c:layout>
   * </c:title>` — Excel's "Format Chart Title -> Title Options ->
   * Position -> Custom" knob (the same drag-handle a user sees when
   * grabbing the title block in Excel's chart editor). The block sits
   * between `<c:tx>` and `<c:overlay>` per `CT_Title` (ECMA-376 Part 1,
   * §21.2.2.210).
   *
   * Each of {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
   * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} is a
   * fraction of the chart frame in the range `0..1` — `(0, 0)` is the
   * upper-left of the chart frame, `(1, 1)` is the lower-right. The
   * coordinates compose independently with {@link titleOverlay} (the
   * overlay flag still records whether the title overlaps the plot
   * area, the manual layout merely overrides where the title block
   * draws). Out-of-range / non-finite / non-numeric coordinates
   * collapse to omitting the matching `<c:x>` / `<c:y>` / `<c:w>` /
   * `<c:h>` slot so a caller can pin only the position
   * ({@link ChartManualLayout.x} / {@link ChartManualLayout.y}) and
   * let the title keep its automatic size, only the size
   * ({@link ChartManualLayout.w} / {@link ChartManualLayout.h}) and
   * let it keep its automatic anchor, or any combination.
   *
   * The writer always emits the matching `<c:xMode>` / `<c:yMode>` /
   * `<c:wMode>` / `<c:hMode>` children with `val="edge"` (Excel's
   * reference shape when the user drags the title to a custom
   * position — the coordinates are absolute fractions of the chart
   * frame, not deltas from the auto-layout baseline).
   *
   * Default: omitted — the title renders at the auto-layout position
   * Excel computes (above the plot area, horizontally centred). Pin a
   * {@link ChartManualLayout} to place the title in a specific
   * quadrant of the chart frame — useful for composing a templated
   * dashboard whose chart title needs to align with an outer header
   * row, or a hero chart whose title should sit in a corner the
   * auto-layout would not pick.
   *
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * layout in that case. An empty layout (every coordinate undefined)
   * collapses to omitting the entire `<c:layout>` block so a fresh
   * chart matches Excel's reference serialization byte-for-byte.
   * Mirrors the writer-side {@link legendLayout} so a caller can
   * thread the same `(x, y, w, h)` 0..1 fractions through both manual-
   * layout slots without bookkeeping a second type.
   */
  titleLayout?: ChartManualLayout;
  /**
   * Chart title background fill color. Maps to `<c:title><c:spPr>
   * <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></c:spPr>
   * </c:title>` (CT_Title, ECMA-376 Part 1, §21.2.2.210) — Excel's
   * "Format Chart Title -> Fill -> Solid fill -> Color" picker (the
   * same dialog the user reaches by right-clicking the title block).
   * The element sits on `<c:title>` between `<c:overlay>` and
   * `<c:txPr>` per the CT_Title schema sequence — distinct from
   * {@link titleColor}, which lands on the run / default-paragraph's
   * `<a:defRPr><a:solidFill>` text-character-properties slot.
   *
   * Accepts a 6-character hex sRGB triple with or without a leading
   * `#` (e.g. `"FFFF00"`, `"#FFFF00"`, `"ffff00"`); the writer
   * normalizes to OOXML's canonical 6-character uppercase form before
   * emit. Empty / whitespace-only strings, alpha-channel forms,
   * non-hex characters, and non-string escapes from an untyped caller
   * all collapse to `undefined` so the writer skips the entire
   * `<c:spPr>` block and the title renders at the theme default fill
   * (Excel's reference behavior for a fresh title that has not had a
   * custom background fill picked — typically a transparent
   * background with no `<c:spPr>` block).
   *
   * Default: omitted — the title renders at the theme default fill
   * (no `<c:spPr>` block, matching Excel's reference serialization
   * for a fresh chart title whose background has not been
   * customized). Pin a hex color to render the title background in
   * that color (e.g. `"FFFF00"` for a highlighted title on a
   * dashboard tile, or `"FFFFFF"` for an opaque white title on a
   * non-white plot area).
   *
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * fill in either case. Mirrors `plotAreaFillColor` /
   * `legendFillColor` — same accept-with-or-without-`#` hex grammar,
   * same OOXML `<c:spPr><a:solidFill><a:srgbClr val=".."/>
   * </a:solidFill></c:spPr>` mapping — so a caller can thread a
   * single hex string through every `<c:spPr>`-based fill slot.
   * Composes independently with {@link titleBold} / {@link titleItalic}
   * / {@link titleStrike} / {@link titleUnderline} / {@link titleColor}
   * / {@link titleFontSize} / {@link titleFontFamily} /
   * {@link titleRotation} / {@link titleOverlay} / {@link titleLayout}:
   * the fill lands on the title's `<c:spPr>` block, the typography
   * knobs on the title's `<c:tx><c:rich><a:p>` rich-text body, the
   * layout on the title's `<c:layout>` block, and the overlay flag on
   * `<c:overlay>` — every knob targets a different child of
   * `<c:title>`.
   */
  titleFillColor?: string;
  /**
   * Chart title border (stroke) solid color as a 6-digit RGB hex string
   * (e.g. `"1F77B4"`). Maps to `<c:title><c:spPr><a:ln><a:solidFill>
   * <a:srgbClr val="RRGGBB"/></a:solidFill></a:ln></c:spPr></c:title>`
   * — Excel's "Format Chart Title -> Border -> Solid line -> Color"
   * picker. The OOXML `<a:srgbClr val=".."/>` carries the 6-character
   * uppercase hex sRGB color (`CT_SRgbColor` inside the line's solid
   * fill choice — ECMA-376 Part 1, §20.1.2.3.32 / §20.1.2.3.24). The
   * `<c:spPr>` slot lives between `<c:overlay>` and `<c:txPr>` /
   * `<c:extLst>` per CT_Title (ECMA-376 Part 1, §21.2.2.210); `<a:ln>`
   * follows the optional `<a:solidFill>` (fill) child inside `<c:spPr>`
   * per `CT_ShapeProperties` (ECMA-376 Part 1, §20.1.2.3.13).
   *
   * Accepts a leading `#` and any case; the writer collapses to the
   * OOXML canonical uppercase form. Malformed inputs (wrong length,
   * non-hex characters, alpha-channel forms, non-string escapes from
   * an untyped caller) collapse to `undefined` and the writer omits
   * the `<a:ln>` block (Excel's reference serialization for a chart
   * title that inherits the auto-stroke — typically no border).
   *
   * Default: omitted — the title inherits the auto-stroke Excel picks
   * from the chart's theme (typically no visible border). Pin a hex
   * color to mirror Excel's "Format Chart Title -> Border -> Solid
   * line" knob and paint a flat border around the title block —
   * useful for dashboard tiles where the title should be visually
   * framed against the chart background, or to highlight the title
   * against a busy theme.
   *
   * Composes independently with {@link titleFillColor} — the two
   * knobs land on the same `<c:spPr>` block but on different children
   * (`<a:solidFill>` for the fill, `<a:ln>` for the stroke), and the
   * writer authors a `<c:spPr>` whenever either knob is set. A caller
   * can pin one without the other; pinning both produces a filled
   * title block with a colored border.
   *
   * Patterned / gradient strokes are not modelled — only the solid
   * sRGB form lands on the wire. Theme-color references
   * (`<a:schemeClr>`) likewise drop to `undefined` so a parsed value
   * always carries a literal hex Excel will render byte-for-byte.
   *
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * stroke in either case. Mirrors `plotAreaBorderColor` — same
   * accept-with-or-without-`#` hex grammar, same OOXML `<c:spPr>
   * <a:ln><a:solidFill><a:srgbClr val=".."/></a:solidFill></a:ln>
   * </c:spPr>` mapping — so a caller can thread a single hex string
   * through every `<c:spPr><a:ln>`-based stroke slot.
   */
  titleBorderColor?: string;
  /**
   * Chart title border (stroke) thickness in points (e.g. `1.5`). Maps
   * to the `w` attribute on `<c:title><c:spPr><a:ln w="EMU">` —
   * Excel's "Format Chart Title -> Border -> Width" spinner. The OOXML
   * `w` attribute carries the stroke width in English Metric Units
   * (`1 pt = 12 700 EMU`) per `CT_LineProperties` (ECMA-376 Part 1,
   * §20.1.2.3.24). The writer multiplies the input by 12 700 and rounds
   * to the nearest integer because the schema types `w` as `xsd:int`.
   *
   * Accepts any finite number; values are clamped to the `0.25..13.5`
   * pt band Excel's UI exposes (the same band used by the series
   * stroke knob `series[i].stroke.width`, the plot-area border width
   * knob {@link plotAreaBorderWidth}, and the legend border width knob
   * {@link legendBorderWidth}) and snapped to the 0.25 pt grid so a
   * parsed-then-written width does not drift across round-trips.
   * Non-finite / non-numeric / `NaN` values collapse to `undefined`
   * and the writer omits the `w` attribute (the line keeps Excel's
   * auto-thickness — typically 0.75 pt).
   *
   * Default: omitted — the title border inherits Excel's
   * auto-thickness. Pin a value to thicken the border around the title
   * block (a hairline at 0.25 pt, a heavy frame at several points) or
   * to match the stroke width of neighboring chart tiles.
   *
   * Composes independently with {@link titleBorderColor} — the width
   * attribute lands on the same `<a:ln>` element as the color's
   * `<a:solidFill>` child, but the writer authors `<a:ln>` whenever
   * either knob is set. A caller can pin a width without a color (the
   * border picks Excel's auto-color), pin a color without a width (the
   * border picks Excel's auto-thickness), or pin both. Setting only the
   * width is valid Excel UI — the user can drag the "Width" spinner
   * without picking a custom color.
   *
   * Silently ignored when no title is rendered (`showTitle === false`
   * or `title` is absent) — there is no `<c:title>` block to host the
   * `<c:spPr>` slot in that case. Mirrors {@link plotAreaBorderWidth}
   * and {@link legendBorderWidth} — same accept-finite-number grammar,
   * same `<a:ln w="EMU">` host attribute on a different parent
   * (`<c:title>` rather than `<c:plotArea>` / `<c:legend>`), and shares
   * the same 0.25..13.5 pt clamp + 0.25 pt snap so width values
   * compose the same way at the call site.
   */
  titleBorderWidth?: number;
  /**
   * Chart title border (stroke) preset dash pattern. Maps to the `val`
   * attribute on `<c:title><c:spPr><a:ln><a:prstDash val=".."/>`.
   * Mirrors Excel's "Format Chart Title -> Border -> Dash type" picker.
   * Same {@link ChartBorderDash} accept-or-drop grammar as
   * {@link plotAreaBorderDash} — `"solid"` collapses to `undefined`
   * so absence and the OOXML default round-trip identically.
   *
   * Composes independently with {@link titleBorderColor} and
   * {@link titleBorderWidth} — all three knobs share the same `<a:ln>`
   * element. Silently ignored when no title is rendered
   * (`showTitle === false` or `title` is absent).
   */
  titleBorderDash?: ChartBorderDash;
  /**
   * Auto-title-deleted flag. Maps to `<c:chart><c:autoTitleDeleted
   * val=".."/>` — Excel's record of whether the user explicitly deleted
   * the auto-generated title that single-series charts synthesise from
   * the series name. Independent of whether a literal {@link title} is
   * authored: pinning `true` suppresses Excel's auto-title even when
   * no `<c:title>` element is emitted.
   *
   * Default: `undefined` — the writer derives the value from the
   * presence of a literal title. When {@link title} is set (and
   * {@link showTitle} is not `false`) the writer emits
   * `<c:autoTitleDeleted val="0"/>` so Excel keeps the literal title
   * visible; when no literal title is rendered the writer emits
   * `<c:autoTitleDeleted val="1"/>` so single-series charts do not
   * silently grow an auto-title from the series name. Pin the field
   * explicitly to override that derivation — e.g. set `false` on a
   * titleless single-series column chart to let Excel synthesise the
   * series-name title, or `true` on a charted dashboard tile that
   * should stay anonymous.
   *
   * The OOXML schema places `<c:autoTitleDeleted>` on `<c:chart>`
   * (between `<c:title>` and `<c:plotArea>` per CT_Chart, ECMA-376
   * Part 1, §21.2.2.4); the writer always emits the element so the
   * rendered intent is explicit on roundtrip — Excel itself includes
   * it on every reference serialization.
   */
  autoTitleDeleted?: boolean;
  /** Alternative text for screen readers (lands in xdr:cNvPr/@descr). */
  altText?: string;
  /** Caption for the chart frame (lands in xdr:cNvPr/@title). */
  frameTitle?: string;
  /**
   * Chart-level data labels applied to every series that does not set
   * its own {@link ChartSeries.dataLabels}. Pass a single
   * {@link ChartDataLabels} object to enable Excel's small in-chart
   * value/category annotations.
   */
  dataLabels?: ChartDataLabels;
  /**
   * How Excel renders missing / blank cells in the source data. Maps
   * to `<c:dispBlanksAs val=".."/>` on `<c:chart>`. Default: `"gap"`
   * (the OOXML default Excel itself emits). Set `"zero"` to anchor the
   * line / bar to the X axis at missing points, or `"span"` to
   * connect across the gap on line and scatter charts. See
   * {@link ChartDisplayBlanksAs} for the accepted set.
   */
  dispBlanksAs?: ChartDisplayBlanksAs;
  /**
   * Vary the color of each data point within the same series. Maps to
   * `<c:varyColors val=".."/>` on the chart-type element
   * (`<c:barChart>`, `<c:lineChart>`, `<c:pieChart>`, ...). Excel
   * exposes the same toggle under "Format Data Series → Fill →
   * Vary colors by point".
   *
   * Excel's per-family defaults differ:
   *   - `pie`, `doughnut`         → `true`  (each slice gets a unique color)
   *   - `bar`, `column`, `line`,
   *     `area`, `scatter`         → `false` (every point on a series
   *                                  shares one color)
   *
   * The writer falls back to those per-family defaults when the field
   * is omitted, so a fresh chart matches Excel's reference
   * serialization. Pin `true` on a single-series bar / column chart to
   * paint each bar a different color, or pin `false` on a doughnut to
   * collapse every wedge to the same color (Excel's "single color"
   * preset).
   *
   * The OOXML schema places `<c:varyColors>` on every chart-type
   * element except `surfaceChart`, `surface3DChart`, and `stockChart`.
   * Hucre's writer emits the element on every authored family, so
   * `varyColors` round-trips on bar / column / line / pie / doughnut /
   * area / scatter charts; surface / stock are not authored by hucre's
   * writer.
   */
  varyColors?: boolean;
  /**
   * Scatter sub-style for `scatter` charts. Maps to
   * `<c:scatterChart><c:scatterStyle val=".."/></c:scatterChart>`.
   * Default: `"lineMarker"` (Excel's chart-picker default — straight
   * lines with markers). Pass `"smooth"` for Excel's "Scatter with
   * Smooth Lines", `"marker"` / `"none"` for "Scatter with Only
   * Markers", `"line"` for "Scatter with Straight Lines", and
   * `"smoothMarker"` for "Scatter with Smooth Lines and Markers". See
   * {@link ChartScatterStyle} for the full preset list.
   *
   * Ignored for every other chart kind — the OOXML schema places
   * `<c:scatterStyle>` exclusively on `<c:scatterChart>`. Use the
   * per-series {@link ChartSeries.smooth} flag to pick a curve on a
   * line chart or pin smoothing on individual scatter series.
   */
  scatterStyle?: ChartScatterStyle;
  /**
   * Whether the chart only plots data from visible cells. Maps to
   * `<c:plotVisOnly val=".."/>` on `<c:chart>`. Mirrors Excel's
   * "Hidden and Empty Cells → Show data in hidden rows and columns"
   * checkbox: when the box is checked, hidden cells stay in the chart
   * and `plotVisOnly` is `false`; when unchecked (the Excel UI
   * default), hidden cells drop out and `plotVisOnly` is `true`.
   *
   * Default: `true` — the OOXML schema default and what every fresh
   * Excel chart emits. Set `false` to keep hidden rows / columns in
   * the rendered chart, useful when the source data range hides helper
   * cells or the dashboard's filter view should not affect the chart.
   *
   * The writer always emits the element so the rendered intent is
   * explicit on roundtrip — Excel itself includes it in every reference
   * serialization.
   */
  plotVisOnly?: boolean;
  /**
   * Whether data labels are shown for points whose values exceed the
   * chart's maximum axis bound. Maps to `<c:showDLblsOverMax val=".."/>`
   * on `<c:chart>`. The element sits at the tail of CT_Chart (after
   * `<c:dispBlanksAs>` and before `<c:extLst>`).
   *
   * Default: `true` — the OOXML schema default. When the value axis
   * is auto-scaled the flag has no observable effect because no point
   * exceeds the max; the toggle only matters when the caller pinned a
   * tight `<c:max>` via {@link ChartAxisScale.max} and a series carries
   * values above it. Setting `false` matches Excel's "Format Axis →
   * Labels → Show data labels for values over maximum scale" checkbox
   * unchecked — the labels for the over-max points disappear while the
   * connector / line still draws above the plot area.
   *
   * The writer always emits the element so the rendered intent is
   * explicit on roundtrip — Excel itself includes it in every reference
   * serialization. Mirrors {@link plotVisOnly} / {@link dispBlanksAs}
   * (the other always-emitted chart-level toggles); a value pinned by
   * the caller round-trips identically through {@link cloneChart}.
   */
  showDLblsOverMax?: boolean;
  /**
   * Whether the chart frame is drawn with rounded corners. Maps to
   * `<c:roundedCorners val=".."/>` on `<c:chartSpace>` (a sibling of
   * `<c:chart>`, not a child). Mirrors Excel's "Format Chart Area →
   * Border → Rounded corners" toggle.
   *
   * Default: `false` — the OOXML schema default and what every fresh
   * Excel chart emits. Set `true` to soften the chart frame's outer
   * edge, useful when matching a dashboard whose other charts already
   * carry the rounded look from a template.
   *
   * The writer always emits the element so the rendered intent is
   * explicit on roundtrip — Excel itself includes it in every reference
   * serialization.
   */
  roundedCorners?: boolean;
  /**
   * Whether to render up / down bars between paired series on a line
   * chart. Maps to `<c:lineChart><c:upDownBars/></c:lineChart>` —
   * Excel's "Add Chart Element -> Up/Down Bars" toggle. The element
   * paints a vertical bar at each category whose top tracks the higher
   * series value and bottom tracks the lower one (typically used to
   * highlight open / close differences on a line-style stock chart).
   *
   * Only meaningful for `line` charts — the OOXML schema places
   * `<c:upDownBars>` on `CT_LineChart`, `CT_Line3DChart`, and
   * `CT_StockChart`; the writer never emits it on bar / column / pie /
   * doughnut / area / scatter, so the field is silently dropped on
   * those families. Default: `false` (no up / down bars; Excel's
   * reference serialization for a fresh line chart omits the element).
   *
   * The writer emits a default `<c:gapWidth val="150"/>` child to
   * mirror Excel's reference serialization — `150` is the OOXML
   * default for `CT_UpDownBars/gapWidth`. Pin
   * {@link upDownBarsGapWidth} to thin or widen the bars.
   */
  upDownBars?: boolean;
  /**
   * Width of the gap between up / down bars as a percentage of the bar
   * width. Maps to `<c:lineChart><c:upDownBars><c:gapWidth val=".."/>
   * </c:upDownBars></c:lineChart>` — Excel's "Format Up Bars / Down
   * Bars -> Gap Width" slider. Accepted range: `0` – `500` (the OOXML
   * `ST_GapAmount` schema). Excel's default is `150` (each up/down
   * bar group's gap equals 1.5× the bar width); smaller values pack
   * the bars tighter, `0` removes the gap entirely.
   *
   * Only meaningful when {@link upDownBars} is `true` — the writer
   * silently drops the value on every other line chart configuration
   * (the OOXML schema places `<c:gapWidth>` on `CT_UpDownBars`, so
   * there is no slot for the value when the parent element is not
   * emitted). The writer also drops the field on bar / column / pie /
   * doughnut / area / scatter chart kinds for the same reason
   * {@link upDownBars} is line-only.
   *
   * Out-of-range or non-finite values fall back to the OOXML default
   * `150` so a fresh chart with a corrupt input still matches Excel's
   * reference serialization. Non-integer values round to the nearest
   * whole percent (Excel's UI accepts integer percentages only).
   *
   * Distinct from {@link gapWidth}: the bar-chart gap width controls
   * spacing between category groups on a `<c:barChart>`, while this
   * field controls the spacing between the up / down bars themselves
   * on a line chart. Both share the same `ST_GapAmount` schema range
   * but are independently scoped.
   */
  upDownBarsGapWidth?: number;
  /**
   * Whether the line chart paints markers at each data point. Maps to
   * `<c:lineChart><c:marker val=".."/></c:lineChart>` — Excel's
   * "Line vs. Line with Markers" chart-type distinction at the
   * chart level. The flag is the chart-level visibility gate; per-
   * series {@link ChartSeries.marker} still picks the symbol / size /
   * color when this gate lets markers render.
   *
   * Default: `true` (the Excel reference behavior — every authored
   * line chart emits `<c:marker val="1"/>`, and hucre's writer mirrors
   * that for back-compat with existing renders). Set `false` to
   * suppress markers chart-wide (the "Line" preset look) — useful when
   * a templated dashboard line chart should render as a clean stroke
   * without the per-point dots.
   *
   * Only meaningful for `line` charts — the OOXML schema places
   * `<c:marker>` (the chart-level CT_Boolean variant) exclusively on
   * `CT_LineChart`. The `line3D` and `stock` chart-type elements have
   * no slot for it; the writer never authors those families, so the
   * field is silently dropped on every other chart kind (`bar` /
   * `column` / `pie` / `doughnut` / `area` / `scatter`).
   *
   * The writer always emits the element so the rendered intent is
   * explicit on roundtrip — Excel's reference serialization includes
   * `<c:marker>` on every line chart, and matching that contract keeps
   * the rendered shape stable. The reader collapses the default
   * `val="1"` (and absence) to `undefined` so a fresh chart and a
   * marker-on chart round-trip identically through {@link cloneChart};
   * only an explicit `val="0"` surfaces `false`.
   */
  showLineMarkers?: boolean;
  /**
   * Built-in chart style preset. Maps to `<c:style val=".."/>` on
   * `<c:chartSpace>` (a sibling of `<c:chart>`, not a child). Mirrors
   * Excel's "Chart Design -> Chart Styles" gallery — each integer
   * picks one of the 48 numbered presets that cycle a colored
   * background, gridline density, border, and label styling across
   * the chart.
   *
   * Default: omitted — when the field is absent the writer skips the
   * element and Excel renders the chart with its application default
   * look. Set an integer in the OOXML range (1–48) to pin a preset;
   * out-of-range and non-integer values are silently dropped rather
   * than emit a token Excel would reject.
   *
   * Useful when matching a dashboard whose other charts already
   * carry a particular preset look from a template — clone-through
   * preserves the parsed value so a fresh chart and a templated chart
   * compose side by side without manual re-styling.
   */
  style?: number;
  /**
   * Editing-locale hint. Maps to `<c:lang val=".."/>` on
   * `<c:chartSpace>` (a sibling of `<c:chart>`, not a child). The
   * value is an RFC-1766 / IETF BCP-47 culture name such as `en-US`,
   * `tr-TR`, or `de-DE` — Excel uses it to drive locale-sensitive
   * defaults within the chart (decimal / group separators on
   * unformatted axis ticks, default text font fallback, and the
   * locale recorded for any in-chart text runs).
   *
   * Default: omitted — when the field is absent the writer skips the
   * element and Excel falls back to the workbook's editing language.
   * Excel's reference serialization for a fresh chart authored on an
   * English locale emits `<c:lang val="en-US"/>`; the writer does
   * not pin a default so an unmarked chart and an `en-US` chart do
   * not silently diverge on roundtrip.
   *
   * Useful when restamping a templated chart for a different locale,
   * or carrying a translated dashboard's `tr-TR` / `de-DE` hint
   * through the parse → clone → write loop. Only well-formed culture
   * names are emitted — unrecognized shapes are silently dropped
   * rather than emit a token Excel would reject. The token has to
   * match `[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*` (the IETF language tag
   * subset `<c:lang>` accepts under `xsd:language`).
   */
  lang?: string;
  /**
   * Date-system hint. Maps to `<c:date1904 val=".."/>` on
   * `<c:chartSpace>` (a sibling of `<c:chart>`, not a child). The flag
   * mirrors the host workbook's `<workbookPr date1904="1"/>` toggle —
   * `true` means the chart's date-axis values are interpreted under
   * the 1904 base (Excel for Mac's legacy epoch where day 0 falls on
   * 1904-01-01); `false` (the OOXML default) is the 1900 base.
   *
   * Default: omitted — when the field is absent the writer skips the
   * element entirely and Excel falls back to the workbook's date
   * system. Excel itself emits `<c:date1904 val="0"/>` on every
   * authored chart even under the 1900 base; the writer does not pin
   * that default so an unmarked chart re-parses to `undefined` (the
   * minimal-shape contract every other chart-space toggle follows).
   *
   * Useful when restamping a chart from a 1904-based template into a
   * 1900-based workbook (or vice versa) — pinning the field keeps the
   * chart's date references anchored to the source's epoch even after
   * the host changes. The value is emitted as `<c:date1904 val="1"/>`
   * for `true` and skipped (rather than `val="0"`) for `false` so the
   * writer's behavior matches absence — a re-parse drops the default
   * back to `undefined` either way.
   *
   * Note: `<c:date1904>` lives on `<c:chartSpace>` (per CT_ChartSpace
   * the element sits at the head of the sequence, before `<c:lang>`
   * and `<c:roundedCorners>`), not inside `<c:chart>` — the toggle
   * governs date interpretation across the whole chart document, not
   * just the plot area.
   */
  date1904?: boolean;
  /**
   * Whether the chart paints a data table beneath the plot area. Maps
   * to `<c:plotArea><c:dTable>...</c:dTable></c:plotArea>` — Excel's
   * "Add Chart Element -> Data Table" toggle, which renders a small
   * table of the underlying series values alongside the plotted shape
   * for a quick read of the numbers behind the picture.
   *
   * Pass `true` to enable the table with Excel's reference defaults
   * (every border drawn, the outline frame on, and the legend keys
   * shown next to each series row). Pass an object to opt individual
   * children in or out — each field maps to the matching `<c:dTable>`
   * boolean child. Pass `false` (or omit the field) to suppress the
   * element entirely so the writer skips emission.
   *
   * Default: omitted — Excel renders no data table on a fresh chart.
   *
   * Only meaningful for chart families that have axes — `bar`, `column`,
   * `line`, `area`, and `scatter`. The OOXML schema places `<c:dTable>`
   * inside `<c:plotArea>` after the axes, and pie / doughnut have no
   * axes at all, so the field is silently dropped on those families.
   * See {@link ChartDataTable}.
   */
  dataTable?: boolean | ChartDataTable;
  /**
   * Chart-space protection. Maps to `<c:chartSpace><c:protection>...
   * </c:protection>` (CT_Protection, ECMA-376 Part 1, §21.2.2.142) —
   * the chart-level lock that Excel honors when the parent worksheet
   * is itself protected via `<sheetProtection>`. Each of the five
   * `CT_Boolean` children (`<c:chartObject>`, `<c:data>`,
   * `<c:formatting>`, `<c:selection>`, `<c:userInterface>`) toggles a
   * separate interaction; `false` (the OOXML default) leaves the
   * action permitted and `true` locks it.
   *
   * Pass `true` to declare the bare `<c:protection/>` element with
   * every flag at its OOXML default `false` — equivalent to passing
   * `protection: {}`. Pass an object to opt individual flags in. Pass
   * `false` (or omit the field) to suppress the element entirely so
   * the writer skips emission.
   *
   * Default: omitted — Excel renders no `<c:protection>` element on
   * a fresh chart.
   *
   * Note: Excel only enforces these flags when the host worksheet is
   * itself protected (the chart inherits the sheet's protection
   * boundary). On an unprotected sheet the element round-trips
   * literally but has no observable runtime effect.
   *
   * The element sits on every chart family — pie / doughnut / bar /
   * column / line / area / scatter — because `<c:protection>` lives
   * on `<c:chartSpace>`, not inside `<c:plotArea>`, so axis-shape
   * has no bearing on whether the slot exists. See
   * {@link ChartProtection}.
   */
  protection?: boolean | ChartProtection;
  /**
   * 3-D view configuration. Maps to `<c:chart><c:view3D>` (CT_View3D,
   * ECMA-376 Part 1, §21.2.2.228) — Excel's "3-D Rotation" pane,
   * which controls the X / Y rotation, height and depth percentages,
   * the right-angle-axes flag, and the perspective foreshortening
   * factor on 3D chart families.
   *
   * Pass an object to pin one or more of the six `CT_View3D` children;
   * each unspecified field falls back to Excel's per-family default
   * (the writer skips emission of any child the object leaves unset
   * so the rendered shape matches absence). Pass an empty object
   * (`{}`) to declare a bare `<c:view3D>` shell, useful for
   * round-trip parity with templates that author the element with no
   * children pinned. Omit the field to suppress the element entirely
   * so the writer skips emission.
   *
   * Default: omitted — Excel renders no `<c:view3D>` element on a
   * fresh chart and falls back to the per-family default rotation /
   * perspective.
   *
   * Although `<c:view3D>` is only meaningful on 3D chart families
   * (`bar3D`, `line3D`, `pie3D`, `area3D`, `surface3D`) and hucre's
   * writer authors only 2D families, the OOXML schema accepts the
   * element on every CT_Chart, so the writer pins it whenever the
   * caller provides a non-empty configuration — Excel silently
   * ignores it on 2D families. Useful primarily for round-tripping a
   * 3D template chart through {@link cloneChart}. See
   * {@link ChartView3D}.
   */
  view3D?: ChartView3D;
  /**
   * 3-D floor thickness, in points. Maps to
   * `<c:chart><c:floor><c:thickness val="N"/></c:floor>` —
   * Excel's "Format Floor -> Floor" pane (the `<c:thickness>` child of
   * `CT_Surface`, ECMA-376 Part 1, §21.2.2.214). Excel renders the
   * floor as a flat plate beneath the plot area on 3D chart families;
   * pinning a positive value extrudes the plate to that depth so the
   * floor reads as a solid slab. Default: omitted — Excel renders no
   * `<c:floor>` element on a fresh chart and the per-family floor
   * default (no extrusion) applies.
   *
   * The OOXML schema (`ST_Thickness`, `xsd:unsignedInt`) accepts any
   * non-negative integer; Excel's UI exposes `0..100` under
   * "Format Floor -> Floor -> Thickness". Out-of-range or non-finite
   * inputs drop at write time rather than emit a token Excel's strict
   * validator would reject. The OOXML default `0` collapses to
   * `undefined` for symmetry with the writer's
   * {@link SheetChart.floorThickness} default — absence and `0` mean
   * the same thing on roundtrip.
   *
   * Although `<c:floor>` is only meaningful on 3D chart families
   * (`bar3D`, `line3D`, `pie3D`, `area3D`, `surface3D`) and hucre's
   * writer authors only 2D families, the OOXML schema accepts the
   * element on every CT_Chart, so the writer pins it whenever the
   * caller provides a positive thickness — Excel silently ignores it
   * on 2D families. Useful primarily for round-tripping a 3D template
   * chart through {@link cloneChart}. The element sits on `<c:chart>`
   * between `<c:view3D>` and `<c:sideWall>` / `<c:backWall>` /
   * `<c:plotArea>` per CT_Chart.
   */
  floorThickness?: number;
  /**
   * 3-D side-wall thickness, in points. Maps to
   * `<c:chart><c:sideWall><c:thickness val="N"/></c:sideWall>` —
   * Excel's "Format Side Wall -> Side Wall" pane (the `<c:thickness>`
   * child of `CT_Surface`, ECMA-376 Part 1, §21.2.2.187). Excel
   * renders the side wall as a flat plate flanking the plot area on
   * 3D chart families; pinning a positive value extrudes the plate to
   * that depth so the wall reads as a solid slab. Default: omitted —
   * Excel renders no `<c:sideWall>` element on a fresh chart and the
   * per-family side-wall default (no extrusion) applies.
   *
   * The OOXML schema (`ST_Thickness`, `xsd:unsignedInt`) accepts any
   * non-negative integer; Excel's UI exposes `0..100` under
   * "Format Side Wall -> Side Wall -> Thickness". Out-of-range or
   * non-finite inputs drop at write time rather than emit a token
   * Excel's strict validator would reject. The OOXML default `0`
   * collapses to `undefined` for symmetry with the writer's
   * {@link SheetChart.sideWallThickness} default — absence and `0`
   * mean the same thing on roundtrip.
   *
   * Although `<c:sideWall>` is only meaningful on 3D chart families
   * (`bar3D`, `line3D`, `pie3D`, `area3D`, `surface3D`) and hucre's
   * writer authors only 2D families, the OOXML schema accepts the
   * element on every CT_Chart, so the writer pins it whenever the
   * caller provides a positive thickness — Excel silently ignores it
   * on 2D families. Useful primarily for round-tripping a 3D template
   * chart through {@link cloneChart}. The element sits on `<c:chart>`
   * between `<c:floor>` and `<c:backWall>` / `<c:plotArea>` per
   * CT_Chart; mirrors {@link SheetChart.view3D} as a chart-level 3D
   * styling knob.
   */
  sideWallThickness?: number;
  /**
   * 3-D back-wall thickness, in points. Maps to
   * `<c:chart><c:backWall><c:thickness val="N"/></c:backWall>` —
   * Excel's "Format Back Wall -> Back Wall" pane (the `<c:thickness>`
   * child of `CT_Surface`, ECMA-376 Part 1, §21.2.2.214). Excel renders
   * the back wall as a flat plate behind the plot area on 3D chart
   * families; pinning a positive value extrudes the plate to that depth
   * so the wall reads as a solid slab. Default: omitted — Excel renders
   * no `<c:backWall>` element on a fresh chart and the per-family
   * back-wall default (no extrusion) applies.
   *
   * The OOXML schema (`ST_Thickness`, `xsd:unsignedInt`) accepts any
   * non-negative integer; Excel's UI exposes `0..100` under
   * "Format Back Wall -> Back Wall -> Thickness". Out-of-range or
   * non-finite inputs drop at write time rather than emit a token
   * Excel's strict validator would reject. The OOXML default `0`
   * collapses to `undefined` for symmetry with the writer's
   * {@link SheetChart.backWallThickness} default — absence and `0` mean
   * the same thing on roundtrip.
   *
   * Although `<c:backWall>` is only meaningful on 3D chart families
   * (`bar3D`, `line3D`, `pie3D`, `area3D`, `surface3D`) and hucre's
   * writer authors only 2D families, the OOXML schema accepts the
   * element on every CT_Chart, so the writer pins it whenever the
   * caller provides a positive thickness — Excel silently ignores it
   * on 2D families. Useful primarily for round-tripping a 3D template
   * chart through {@link cloneChart}. The element sits on `<c:chart>`
   * between `<c:sideWall>` and `<c:plotArea>` per CT_Chart — `<c:floor>`
   * /  `<c:sideWall>` / `<c:backWall>` are independent siblings on
   * `<c:chart>`.
   */
  backWallThickness?: number;
  /**
   * Per-axis configuration rendered alongside the plot area. The `x`
   * axis is the category axis for bar/column/line/area (or the bottom
   * value axis for scatter); the `y` axis is the value axis. Ignored
   * for `pie` and `doughnut` charts because they have no axes in
   * OOXML.
   *
   * `title` maps to a `<c:title>` element nested inside the matching
   * `<c:catAx>` / `<c:valAx>`. Pass an empty string or omit the entry
   * to skip the title — Excel renders no axis label by default.
   *
   * `gridlines` toggles `<c:majorGridlines>` / `<c:minorGridlines>`.
   * Omitting the field skips both — useful when porting a clean look
   * across cloned charts. Set `major: true` to draw the heavier
   * reference lines that Excel shows by default on the value axis;
   * `minor: true` adds the lighter half-step lines.
   *
   * `scale` pins the value axis to explicit `<c:min>` / `<c:max>` /
   * `<c:majorUnit>` / `<c:minorUnit>` / `<c:logBase>` bounds. Excel
   * auto-computes any field omitted from the object. Bar/column/line/
   * area charts apply scaling to the Y axis (`<c:valAx>`); scatter
   * charts apply it to whichever axis the field is set on.
   *
   * `numberFormat` pins the tick-label format via `<c:numFmt>` —
   * useful when the cloned chart needs a different format from the
   * source data range (e.g. forcing `"0.00%"` on a percentage chart
   * whose underlying cells are stored as decimals).
   *
   * `tickLblSkip` and `tickMarkSkip` thin out a crowded category axis.
   * Both map to category-axis-only OOXML elements (`<c:tickLblSkip>` /
   * `<c:tickMarkSkip>` on `CT_CatAx` / `CT_DateAx`); they have no slot
   * on `<c:valAx>` and are silently ignored on the value axis or on
   * scatter charts (whose two axes are both value axes).
   *
   * `hidden` collapses the axis line, tick marks, and tick labels off
   * the rendered chart by emitting `<c:delete val="1"/>`. Maps to
   * Excel's "Format Axis -> Axis Options -> Labels -> Show axis" toggle
   * (and the matching context-menu "Delete" action). Useful for
   * minimal "sparkline-style" dashboard tiles where only the data
   * series should remain visible.
   */
  axes?: {
    /** Category axis (bar/column/line/area) or X value axis (scatter). */
    x?: {
      title?: string;
      /**
       * Rotation of the axis title in whole degrees. Maps to
       * `<c:catAx><c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title></c:catAx>`
       * (or `<c:valAx>` for scatter / value axes) — Excel's "Format Axis
       * Title -> Size & Properties -> Alignment -> Custom angle" knob.
       *
       * Mirrors {@link SheetChart.titleRotation} for axis titles — same
       * `-90..90` band Excel's UI exposes, same conversion factor
       * (60000ths of a degree on the wire). Useful for standing the Y-
       * axis title vertically next to the value labels (the typical
       * "rotated axis label" dashboard look) or pinning a custom angle
       * on a long X-axis title that would otherwise crowd the plot
       * area.
       *
       * Default: `0` (no rotation, Excel's reference look). Out-of-range
       * inputs clamp to the nearest endpoint; non-finite (`NaN`,
       * `Infinity`) and non-numeric inputs drop at write time so the
       * writer never emits a token Excel's strict validator would
       * reject. Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case).
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
       * per the OOXML schema, so the field round-trips on every chart
       * family that has axes (bar / column / line / area / scatter).
       * Pie / doughnut have no axes at all, so the field is silently
       * dropped on those families.
       */
      axisTitleRotation?: number;
      /**
       * Axis title font size in whole or half points. Maps to
       * `<c:catAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
       * <a:r><a:rPr sz="N"/></a:r></a:p></c:rich></c:tx></c:title></c:catAx>`
       * (or `<c:valAx>` for scatter / value axes) — Excel's "Format Axis
       * Title -> Font -> Size" knob. The OOXML attribute is in 100ths
       * of a point, so 12pt serializes as `sz="1200"` and 10pt (Excel's
       * reference default for axis titles) as `sz="1000"`; the writer
       * performs the conversion at emit time and lands the value on
       * both the default-paragraph `<a:defRPr>` and the literal run's
       * `<a:rPr>` so a re-parse picks the size up off either canonical
       * slot.
       *
       * Mirrors {@link SheetChart.titleFontSize} for axis titles — same
       * `1..400`pt band the OOXML `ST_TextFontSize` schema exposes,
       * same 0.5pt half-step granularity Excel's UI exposes, same
       * out-of-range / non-finite drop semantics. Useful for shrinking
       * a long Y-axis unit label so it fits a tight chart frame, or
       * bumping the X-axis title up to match a presentation slide's
       * typography.
       *
       * Default: omitted — the axis title renders at Excel's reference
       * 10pt (the writer's hardcoded default before this field was
       * surfaced). Silently dropped when the axis renders no title
       * (the `<c:title>` element is absent in either case) and on
       * `pie` / `doughnut` charts (no axes at all).
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
       * per the OOXML schema, so the field round-trips on every chart
       * family that has axes (bar / column / line / area / scatter).
       */
      axisTitleFontSize?: number;
      /**
       * Axis title bold flag. Maps to
       * `<c:catAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
       * <a:r><a:rPr b=".."/></a:r></a:p></c:rich></c:tx></c:title></c:catAx>`
       * (or `<c:valAx>` for scatter / value axes) — Excel's "Format
       * Axis Title -> Font -> Bold" toggle. The OOXML attribute is the
       * `xsd:boolean` `b` on `CT_TextCharacterProperties` (ECMA-376
       * Part 1, §21.1.2.3.7); the writer lands the value on both the
       * default-paragraph `<a:defRPr>` and the literal run's `<a:rPr>`
       * so a re-parse picks the flag up off either canonical slot —
       * Excel keeps the two attributes in sync.
       *
       * Mirrors {@link SheetChart.titleBold} for axis titles — same
       * canonical-slot pair, same `boolean | null` clone grammar, same
       * silent drop when the axis renders no title. Useful for matching
       * the chart-level title's emphasis on the axis labels (e.g.
       * bolding the Y-axis "Revenue (USD)" caption to align with a
       * dashboard's heavy-weight typography).
       *
       * Default: omitted — the axis title renders non-bold (`b="0"`,
       * the writer's reference serialization for a fresh axis title).
       * Set `true` to emit `b="1"` on both slots so the axis title
       * renders bold; set `false` explicitly to pin the non-default
       * `b="0"` (functionally identical to omission, but useful when
       * overriding a templated axis title that had bold pinned
       * upstream).
       *
       * Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case) and on `pie` /
       * `doughnut` charts (no axes at all). Composes independently
       * with {@link axisTitleRotation} / {@link axisTitleFontSize} /
       * {@link axisTitleItalic}: all four fields land on the same
       * `<c:title>` body so a single configuration call threads cleanly
       * through every axis-title knob Excel exposes.
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
       * per the OOXML schema, so the field round-trips on every chart
       * family that has axes (bar / column / line / area / scatter).
       */
      axisTitleBold?: boolean;
      /**
       * Axis title italic flag. Maps to
       * `<c:catAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
       * <a:r><a:rPr i=".."/></a:r></a:p></c:rich></c:tx></c:title></c:catAx>`
       * (or `<c:valAx>` for scatter / value axes) — Excel's "Format Axis
       * Title -> Font -> Italic" toggle. The OOXML attribute is the
       * `xsd:boolean` `i` on `CT_TextCharacterProperties` (ECMA-376
       * Part 1, §21.1.2.3.7); the writer lands the value on both the
       * default-paragraph `<a:defRPr>` and the literal run's `<a:rPr>`
       * so a re-parse picks the flag up off either canonical slot —
       * Excel keeps the two attributes in sync.
       *
       * Mirrors {@link SheetChart.titleItalic} for axis titles — same
       * canonical-slot pair, same drop-on-default semantics, same
       * `boolean | null` clone grammar so a single configuration call
       * threads cleanly through both the chart title and either axis
       * title without bookkeeping the canonical OOXML slots.
       *
       * Default: omitted — the axis title renders non-italic (no `i`
       * attribute, Excel's reference serialization for a fresh axis
       * title; the application-default `false` collapses to absence).
       * Set `true` to emit `i="1"` on both slots so the title renders
       * italic; set `false` explicitly to pin the non-default `i="0"`
       * (functionally identical to omission, but useful when overriding
       * a templated axis title that had italic pinned upstream).
       *
       * Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case) and on `pie` /
       * `doughnut` charts (no axes at all).
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
       * per the OOXML schema, so the field round-trips on every chart
       * family that has axes (bar / column / line / area / scatter).
       */
      axisTitleItalic?: boolean;
      /**
       * Axis title font color. Maps to
       * `<c:catAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr>
       * <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr>
       * </a:pPr><a:r><a:rPr><a:solidFill><a:srgbClr val="RRGGBB"/>
       * </a:solidFill></a:rPr></a:r></a:p></c:rich></c:tx></c:title>
       * </c:catAx>` (or `<c:valAx>` for scatter / value axes) — Excel's
       * "Format Axis Title -> Font -> Font Color" picker. The OOXML
       * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
       * sRGB color (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32); the
       * writer lands the fill on both the default-paragraph
       * `<a:defRPr>` and the literal run's `<a:rPr>` so a re-parse
       * picks the color up off either canonical slot — Excel keeps the
       * two values in sync.
       *
       * Mirrors {@link SheetChart.titleColor} for axis titles — same
       * canonical-slot pair, same accept-with-or-without-`#` grammar
       * (`"FF0000"` / `"#FF0000"` / `"ff0000"` all collapse to the
       * uppercase canonical form), same `string | null` clone grammar
       * (`undefined` inherits, `null` drops, a hex string replaces) so
       * a single configuration call threads cleanly through both the
       * chart title and either axis title without bookkeeping the
       * canonical OOXML slots. Mirrors {@link axisTitleRotation} /
       * {@link axisTitleFontSize} / {@link axisTitleBold} /
       * {@link axisTitleItalic} for the same `<c:title>` body, so all
       * five axis-title knobs (rotation, size, bold, italic, color)
       * compose freely on the same axis.
       *
       * Default: omitted — the axis title renders in Excel's reference
       * inherited theme color (no `<a:solidFill>` element, the writer
       * skips the fill block entirely). Pin a hex value to render the
       * axis title in that color (e.g. `"1070CA"` for the dashboard
       * hero blue the issue-#136 example reaches for). The 8-character
       * `#RRGGBBAA` form is *not* accepted — alpha lives on
       * `<a:srgbClr><a:alpha val=".."/>` which is a separate runs-level
       * knob; pinning `axisTitleColor` carries the RGB triple only.
       * Malformed inputs (wrong length, non-hex characters, alpha-
       * channel form) collapse to `undefined` so a stray non-hex token
       * never produces a malformed `<a:srgbClr>`.
       *
       * Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case) and on `pie` /
       * `doughnut` charts (no axes at all).
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
       * per the OOXML schema, so the field round-trips on every chart
       * family that has axes (bar / column / line / area / scatter).
       */
      axisTitleColor?: string;
      /**
       * Axis title strikethrough flag. Maps to
       * `<c:catAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
       * <a:r><a:rPr strike=".."/></a:r></a:p></c:rich></c:tx></c:title></c:catAx>`
       * (or `<c:valAx>` for scatter / value axes) — Excel's "Format Axis
       * Title -> Font -> Strikethrough" toggle. The OOXML attribute is
       * the `ST_TextStrikeType` enum on `CT_TextCharacterProperties`
       * (ECMA-376 Part 1, §21.1.2.3.7) with three values: `"noStrike"`
       * (the OOXML default — no strikethrough), `"sngStrike"` (single
       * horizontal line, the value Excel's UI checkbox emits), and
       * `"dblStrike"` (double horizontal line, a non-UI variant Excel
       * does not surface in its ribbon). The writer lands the value on
       * both the default-paragraph `<a:defRPr>` and the literal run's
       * `<a:rPr>` so a re-parse picks the flag up off either canonical
       * slot — Excel keeps the two attributes in sync.
       *
       * Modeled as a boolean for symmetry with {@link axisTitleBold} /
       * {@link axisTitleItalic}: `true` emits `strike="sngStrike"`
       * (Excel's UI "Strikethrough" checkbox — single line). Absence
       * and non-boolean tokens collapse to omitting the attribute
       * (Excel's reference serialization for a non-strikethrough axis
       * title — the application-default `"noStrike"` collapses to
       * absence). Set `false` explicitly to pin the non-default
       * omission (functionally identical to omission, but useful when
       * overriding a templated axis title that had strikethrough pinned
       * upstream).
       *
       * Hucre's writer emits only `"sngStrike"` to keep the surfaced
       * shape consistent with what Excel's reference UI authors. The
       * reader collapses the non-UI `"dblStrike"` to `undefined` so a
       * templated axis title that pinned the double-line variant in
       * raw OOXML round-trips to the same `undefined` an unmarked axis
       * title parses to (i.e. the double-line variant silently
       * downgrades to the single-line write grammar rather than
       * fabricate a value the writer would re-emit incorrectly).
       *
       * Mirrors {@link SheetChart.titleStrike} for axis titles — same
       * canonical-slot pair, same drop-on-default semantics, same
       * `boolean | null` clone grammar so a single configuration call
       * threads cleanly through both the chart title and either axis
       * title without bookkeeping the canonical OOXML slots. Mirrors
       * {@link axisTitleRotation} / {@link axisTitleFontSize} /
       * {@link axisTitleBold} / {@link axisTitleItalic} /
       * {@link axisTitleColor} for the same `<c:title>` body, so all
       * six axis-title knobs (rotation, size, bold, italic, color,
       * strike) compose freely on the same axis.
       *
       * Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case) and on `pie` /
       * `doughnut` charts (no axes at all).
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
       * per the OOXML schema, so the field round-trips on every chart
       * family that has axes (bar / column / line / area / scatter).
       */
      axisTitleStrike?: boolean;
      /**
       * Axis title underline flag. Maps to
       * `<c:catAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
       * <a:r><a:rPr u=".."/></a:r></a:p></c:rich></c:tx></c:title></c:catAx>`
       * (or `<c:valAx>` for scatter / value axes) — Excel's "Format Axis
       * Title -> Font -> Underline" picker. The OOXML attribute is the
       * `ST_TextUnderlineType` enum on `CT_TextCharacterProperties`
       * (ECMA-376 Part 1, §21.1.2.3.7) with eighteen values; Excel's UI
       * exposes only `"sng"` (single line — the default underline
       * checkbox) and `"dbl"` (double line). The remaining sixteen
       * tokens (`"none"`, `"words"`, `"heavy"`, `"dotted"`,
       * `"dottedHeavy"`, `"dash"`, `"dashHeavy"`, `"dashLong"`,
       * `"dashLongHeavy"`, `"dotDash"`, `"dotDashHeavy"`,
       * `"dotDotDash"`, `"dotDotDashHeavy"`, `"wavy"`, `"wavyHeavy"`,
       * `"wavyDbl"`) are non-UI variants Excel does not surface in its
       * ribbon. The writer lands the value on both the
       * default-paragraph `<a:defRPr>` and the literal run's `<a:rPr>`
       * so a re-parse picks the flag up off either canonical slot —
       * Excel keeps the two attributes in sync.
       *
       * Modeled as a boolean for symmetry with {@link axisTitleBold} /
       * {@link axisTitleItalic} / {@link axisTitleStrike}: `true` emits
       * `u="sng"` (Excel's UI "Underline" checkbox — single line).
       * Absence and non-boolean tokens collapse to omitting the
       * attribute (Excel's reference serialization for a non-underlined
       * axis title — the application-default `"none"` collapses to
       * absence). Set `false` explicitly to pin the non-default
       * omission (functionally identical to omission, but useful when
       * overriding a templated axis title that had underline pinned
       * upstream).
       *
       * Hucre's writer emits only `"sng"` to keep the surfaced shape
       * consistent with what Excel's reference UI authors. The reader
       * collapses every non-`"sng"` token (the non-UI `"dbl"` variant
       * and the sixteen exotic types) to `undefined` so a templated
       * axis title that pinned a non-single underline in raw OOXML
       * round-trips to the same `undefined` an unmarked axis title
       * parses to (i.e. the exotic underline silently downgrades to the
       * single-line write grammar rather than fabricate a value the
       * writer would re-emit incorrectly).
       *
       * Mirrors {@link SheetChart.titleUnderline} for axis titles —
       * same canonical-slot pair, same drop-on-default semantics, same
       * `boolean | null` clone grammar so a single configuration call
       * threads cleanly through both the chart title and either axis
       * title without bookkeeping the canonical OOXML slots. Mirrors
       * {@link axisTitleRotation} / {@link axisTitleFontSize} /
       * {@link axisTitleBold} / {@link axisTitleItalic} /
       * {@link axisTitleColor} / {@link axisTitleStrike} for the same
       * `<c:title>` body, so all seven axis-title typography knobs
       * (rotation, size, bold, italic, color, strike, underline)
       * compose freely on the same axis.
       *
       * Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case) and on `pie` /
       * `doughnut` charts (no axes at all).
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
       * per the OOXML schema, so the field round-trips on every chart
       * family that has axes (bar / column / line / area / scatter).
       */
      axisTitleUnderline?: boolean;
      /**
       * Axis title font family / typeface. Maps to
       * `<c:catAx>` / `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` ->
       * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
       * typeface=".."/></a:defRPr></a:pPr><a:r><a:rPr><a:latin
       * typeface=".."/></a:rPr></a:r></a:p></c:rich></c:tx></c:title>` —
       * Excel's "Format Axis Title -> Font -> Font" picker. The OOXML
       * `<a:latin typeface=".."/>` element carries the typeface name
       * (`CT_TextFont`, ECMA-376 Part 1, §21.1.2.3.7); the writer
       * lands the element on both the default-paragraph `<a:defRPr>`
       * and the literal run's `<a:rPr>` so a re-parse picks the
       * typeface up off either canonical slot — Excel keeps the two
       * values in sync.
       *
       * Accepts any non-empty string typeface name (e.g. `"Calibri"`,
       * `"Arial"`, `"Times New Roman"`); the writer trims surrounding
       * whitespace and emits the trimmed value verbatim (XML-escaped)
       * so Excel can resolve the named font from the workbook's font
       * scheme or the host system's installed fonts. Empty /
       * whitespace-only strings and non-string tokens collapse to
       * `undefined` so the writer skips the entire `<a:latin>`
       * element and the title inherits Excel's reference theme
       * typeface (Calibri Light from the default Office theme).
       *
       * Mirrors {@link SheetChart.titleFontFamily} for axis titles —
       * same canonical-slot pair, same drop-on-empty semantics, same
       * `string | null` clone grammar so a single configuration call
       * threads cleanly through both the chart title and either axis
       * title without bookkeeping the canonical OOXML slots. Mirrors
       * {@link axisTitleRotation} / {@link axisTitleFontSize} /
       * {@link axisTitleBold} / {@link axisTitleItalic} /
       * {@link axisTitleColor} / {@link axisTitleStrike} /
       * {@link axisTitleUnderline} for the same `<c:title>` body, so
       * all eight axis-title typography knobs compose freely on the
       * same axis.
       *
       * Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case) and on `pie` /
       * `doughnut` charts (no axes at all).
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>`
       * shape per the OOXML schema, so the field round-trips on
       * every chart family that has axes (bar / column / line /
       * area / scatter).
       */
      axisTitleFontFamily?: string;
      /**
       * Axis-title overlay flag. Maps to
       * `<c:catAx><c:title><c:overlay val=".."/></c:title></c:catAx>`
       * (or `<c:valAx>` / `<c:dateAx>` / `<c:serAx>`) — the OOXML
       * `<c:overlay>` child of `CT_Title`. The element controls
       * whether the axis title is drawn on top of (and may overlap)
       * the plot area, mirroring how the chart-level
       * {@link SheetChart.titleOverlay} controls the chart title's
       * overlap.
       *
       * Default: omitted — the writer emits the OOXML default
       * `val="0"` (the axis title reserves its own slot adjacent to
       * the axis and the plot area shrinks to make room), matching
       * Excel's reference serialization for a fresh axis title. Pin
       * `axisTitleOverlay: true` to draw the axis title on top of the
       * plot area so the chart series get the full frame
       * (`val="1"`). Excel's UI does not surface this toggle for axis
       * titles directly, but the OOXML schema (CT_Title is shared
       * with the chart-level title) carries the element on every
       * emitted axis-title block — the writer always emits
       * `<c:overlay>` so a hand-edited template that pinned `val="1"`
       * round-trips cleanly.
       *
       * Mirrors {@link SheetChart.titleOverlay} — same `boolean`
       * shape, same OOXML `<c:overlay val=".."/>` mapping — so a
       * caller can thread a single overlay toggle through both the
       * chart and axis titles.
       *
       * Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case) and on `pie` /
       * `doughnut` charts (no axes at all).
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>`
       * shape per the OOXML schema, so the field round-trips on
       * every chart family that has axes (bar / column / line /
       * area / scatter).
       */
      axisTitleOverlay?: boolean;
      /**
       * Axis-title manual placement. Maps to
       * `<c:catAx><c:title><c:layout><c:manualLayout>...</c:manualLayout>
       * </c:layout></c:title></c:catAx>` (or `<c:valAx>` / `<c:dateAx>` /
       * `<c:serAx>`) — Excel's "Format Axis Title -> Title Options ->
       * Position -> Custom" placement, the same drag-handle a user sees
       * when grabbing the axis-title block in Excel's chart editor.
       *
       * The OOXML `CT_ManualLayout` block (ECMA-376 Part 1, §21.2.2.115)
       * sits inside `CT_Title` between `<c:tx>` and `<c:overlay>` and
       * carries the title's `(x, y)` anchor and `(w, h)` size as
       * fractions of the chart frame in the `0..1` band. Each of
       * {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
       * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} is
       * independently optional — pin only the position
       * ({@link ChartManualLayout.x} / {@link ChartManualLayout.y}) and
       * let the title keep its automatic size, only the size
       * ({@link ChartManualLayout.w} / {@link ChartManualLayout.h}) and
       * let it keep its automatic anchor, or any combination.
       *
       * Mirrors {@link SheetChart.legendLayout} /
       * {@link SheetChart.plotAreaLayout} for axis titles — same
       * {@link ChartManualLayout} shape, same accept-or-drop grammar
       * (out-of-range / non-finite / non-numeric coordinates collapse
       * to `undefined` axis-by-axis), same canonical `xMode="edge"`
       * normalization on emit. The writer always emits
       * `xMode="edge"` / `yMode="edge"` / `wMode="edge"` /
       * `hMode="edge"` next to the matching `<c:x>` / `<c:y>` / `<c:w>` /
       * `<c:h>` slot — `"edge"` is Excel's reference shape when the
       * user drags an element to a custom position (the coordinates
       * are absolute fractions of the chart frame, not deltas from
       * the auto-layout baseline).
       *
       * Default: omitted — the writer skips the entire `<c:layout>`
       * block so the axis title renders at Excel's auto-layout
       * position (adjacent to the matching axis, with the plot area
       * shrunk to make room). An empty layout (every coordinate
       * dropped on normalization) collapses back to no `<c:layout>`
       * block so a fresh chart matches Excel's reference shape
       * byte-for-byte.
       *
       * Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case) and on `pie` /
       * `doughnut` charts (no axes at all).
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>`
       * shape per the OOXML schema, so the field round-trips on
       * every chart family that has axes (bar / column / line /
       * area / scatter).
       */
      axisTitleLayout?: ChartManualLayout;
      /**
       * Axis-title background fill (solid). Maps to
       * `<c:catAx><c:title><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
       * </a:solidFill></c:spPr></c:title></c:catAx>` (or `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>`) — Excel's "Format Axis Title -> Fill
       * -> Solid fill -> Color" picker. The OOXML `<c:spPr>` block on
       * `CT_Title` (ECMA-376 Part 1, §21.2.2.210) carries the title's
       * shape properties and sits between `<c:overlay>` and `<c:txPr>` /
       * `<c:extLst>` per the schema sequence.
       *
       * Accepts a 6-character hex string with or without a leading `#`,
       * any case (`"FF0000"`, `"#1070ca"`, `"abcdef"`); the writer
       * normalizes to the OOXML canonical 6-character uppercase form
       * (`"FF0000"`, `"1070CA"`, `"ABCDEF"`) so a re-parse round-trips
       * losslessly. Malformed inputs (wrong length, non-hex characters,
       * alpha-channel forms like `"FFAA0080"`, empty / whitespace-only
       * strings, non-string escapes from an untyped caller) collapse to
       * `undefined` and the writer skips the entire `<c:spPr>` block —
       * the axis title inherits the theme default fill (typically a
       * transparent title background with no `<c:spPr>` block, matching
       * Excel's reference shape byte-for-byte).
       *
       * Mirrors {@link SheetChart.titleFillColor} for axis titles —
       * same `<c:spPr><a:solidFill><a:srgbClr>` slot, same accept-
       * with-or-without-`#` grammar, same drop-on-malformed semantics.
       * Lands in the same `<c:spPr>` family as
       * {@link SheetChart.plotAreaFillColor} /
       * {@link SheetChart.legendFillColor} but on a distinct host
       * element. Composes independently with
       * {@link axisTitleColor} (the font color) — the two knobs target
       * different children of `<c:title>` (`<c:spPr>` for the
       * background fill, `<c:tx><c:rich><a:p><a:pPr><a:defRPr>
       * <a:solidFill>` for the font color).
       *
       * Default: omitted — no `<c:spPr>` block on the axis title (the
       * title inherits the theme default fill, matching Excel's
       * reference shape byte-for-byte).
       *
       * Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case) and on `pie` /
       * `doughnut` charts (no axes at all).
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>`
       * shape per the OOXML schema, so the field round-trips on
       * every chart family that has axes (bar / column / line /
       * area / scatter).
       */
      axisTitleFillColor?: string;
      /**
       * Axis-title border (line stroke) solid color as a 6-digit RGB
       * hex string (e.g. `"1F77B4"`). Maps to
       * `<c:catAx><c:title><c:spPr><a:ln><a:solidFill><a:srgbClr
       * val="RRGGBB"/></a:solidFill></a:ln></c:spPr></c:title></c:catAx>`
       * (or `<c:valAx>` / `<c:dateAx>` / `<c:serAx>`) — Excel's
       * "Format Axis Title -> Border -> Solid line -> Color" picker.
       * The OOXML `<a:srgbClr val=".."/>` carries the 6-character
       * uppercase hex sRGB color (`CT_SRgbColor` inside the line's
       * solid fill choice — ECMA-376 Part 1, §20.1.2.3.32 /
       * §20.1.2.3.24). The `<c:spPr>` slot lives between
       * `<c:overlay>` and `<c:txPr>` / `<c:extLst>` per CT_Title
       * (ECMA-376 Part 1, §21.2.2.210); `<a:ln>` follows the optional
       * `<a:solidFill>` (fill) child inside `<c:spPr>` per
       * `CT_ShapeProperties` (ECMA-376 Part 1, §20.1.2.3.13).
       *
       * Accepts a 6-character hex string with or without a leading
       * `#`, any case (`"FF0000"`, `"#1070ca"`, `"abcdef"`); the
       * writer normalizes to the OOXML canonical 6-character
       * uppercase form (`"FF0000"`, `"1070CA"`, `"ABCDEF"`) so a
       * re-parse round-trips losslessly. Malformed inputs (wrong
       * length, non-hex characters, alpha-channel forms like
       * `"FFAA0080"`, empty / whitespace-only strings, non-string
       * escapes from an untyped caller) collapse to `undefined` and
       * the writer omits the `<a:ln>` block — the axis title
       * inherits the auto-stroke Excel picks from the chart's theme
       * (typically no visible border, matching Excel's reference
       * shape byte-for-byte).
       *
       * Default: omitted — the title inherits the auto-stroke (no
       * `<a:ln>` block). Pin a hex color to mirror Excel's "Format
       * Axis Title -> Border -> Solid line" knob and paint a flat
       * border around the axis-title block — useful for highlighting
       * an axis title against a busy theme or framing it like a
       * dashboard tile.
       *
       * Composes independently with {@link axisTitleFillColor} —
       * the two knobs land on the same `<c:spPr>` block but on
       * different children (`<a:solidFill>` for the fill, `<a:ln>`
       * for the stroke), and the writer authors a `<c:spPr>`
       * whenever either knob is set. A caller can pin one without
       * the other; pinning both produces a filled axis-title block
       * with a colored border.
       *
       * Patterned / gradient strokes are not modelled — only the
       * solid sRGB form lands on the wire. Theme-color references
       * (`<a:schemeClr>`) likewise drop to `undefined` so a parsed
       * value always carries a literal hex Excel will render
       * byte-for-byte.
       *
       * Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case) and on `pie` /
       * `doughnut` charts (no axes at all). Mirrors
       * {@link SheetChart.titleBorderColor} for axis titles — same
       * `<c:spPr><a:ln><a:solidFill><a:srgbClr>` slot, same
       * accept-with-or-without-`#` grammar, distinct host element.
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>`
       * shape per the OOXML schema, so the field round-trips on
       * every chart family that has axes (bar / column / line /
       * area / scatter).
       */
      axisTitleBorderColor?: string;
      /**
       * Axis-title border (stroke) thickness in points (e.g. `1.5`).
       * Maps to the `w` attribute on `<c:catAx><c:title><c:spPr>
       * <a:ln w="EMU">` (or `<c:valAx>` / `<c:dateAx>` / `<c:serAx>`).
       * Mirrors Excel's "Format Axis Title -> Border -> Width"
       * spinner. The OOXML `w` attribute carries the stroke width in
       * English Metric Units (1 pt = 12 700 EMU) per
       * `CT_LineProperties` (ECMA-376 Part 1, §20.1.2.3.24); the
       * writer multiplies by 12 700 and rounds to the nearest integer
       * because the schema types `w` as `xsd:int`.
       *
       * Accepts any finite number; values are clamped to the
       * `0.25..13.5` pt band Excel's UI exposes (the same band used by
       * {@link SheetChart.plotAreaBorderWidth} / {@link SheetChart.legendBorderWidth} /
       * {@link SheetChart.titleBorderWidth} / the series stroke knob)
       * and snapped to the 0.25 pt grid so a parsed-then-written width
       * does not drift across round-trips. Non-finite / non-numeric /
       * `NaN` values collapse to `undefined` and the writer omits the
       * `w` attribute (the line keeps Excel's auto-thickness —
       * typically 0.75 pt).
       *
       * Default: omitted — the border inherits Excel's auto-thickness.
       *
       * Composes independently with {@link axisTitleBorderColor} —
       * the width attribute lands on the same `<a:ln>` element as the
       * color's `<a:solidFill>` child, but the writer authors `<a:ln>`
       * whenever either knob is set.
       *
       * Silently dropped when the axis renders no title (no `<c:title>`
       * to host `<c:spPr>`) and on `pie` / `doughnut` charts (no axes
       * at all).
       */
      axisTitleBorderWidth?: number;
      /**
       * Axis-title border (stroke) preset dash pattern. Maps to the
       * `val` attribute on `<c:catAx><c:title><c:spPr><a:ln>
       * <a:prstDash val=".."/>` (or `<c:valAx>` / `<c:dateAx>` /
       * `<c:serAx>`). Mirrors Excel's "Format Axis Title -> Border ->
       * Dash type" picker. Same {@link ChartBorderDash}
       * accept-or-drop grammar as the chart-level
       * {@link SheetChart.titleBorderDash} — `"solid"` collapses to
       * `undefined` so absence and the OOXML default round-trip
       * identically.
       *
       * Composes independently with {@link axisTitleBorderColor} and
       * {@link axisTitleBorderWidth} — all three knobs share the same
       * `<a:ln>` element. Silently dropped when the axis renders no
       * title and on `pie` / `doughnut` charts.
       */
      axisTitleBorderDash?: ChartBorderDash;
      gridlines?: ChartAxisGridlines;
      scale?: ChartAxisScale;
      numberFormat?: ChartAxisNumberFormat;
      /**
       * Major tick-mark style. Maps to
       * `<c:catAx><c:majorTickMark val=".."/></c:catAx>` (or
       * `<c:valAx>` for scatter). Default: `"out"` — Excel's reference
       * serialization. See {@link ChartAxisTickMark}.
       */
      majorTickMark?: ChartAxisTickMark;
      /**
       * Minor tick-mark style. Maps to
       * `<c:catAx><c:minorTickMark val=".."/></c:catAx>` (or
       * `<c:valAx>` for scatter). Default: `"none"` — Excel's
       * reference serialization. See {@link ChartAxisTickMark}.
       */
      minorTickMark?: ChartAxisTickMark;
      /**
       * Tick-label position. Maps to
       * `<c:catAx><c:tickLblPos val=".."/></c:catAx>` (or
       * `<c:valAx>` for scatter). Default: `"nextTo"` — Excel's
       * reference serialization. See {@link ChartAxisTickLabelPosition}.
       */
      tickLblPos?: ChartAxisTickLabelPosition;
      /**
       * Tick-label rotation in degrees. Maps to
       * `<c:catAx><c:txPr><a:bodyPr rot="N"/></c:txPr></c:catAx>` (or
       * `<c:valAx>` for scatter / value axes) — Excel's "Format Axis ->
       * Alignment -> Custom angle" knob. Useful for rotating long
       * category labels diagonally so they fit underneath a column
       * chart's bars without overlapping their neighbours.
       *
       * Accepted range: `-90..90` (the band Excel's UI exposes; values
       * outside the band clamp to the nearest endpoint). The OOXML
       * `rot` attribute is in 60000ths of a degree — the writer
       * converts at emit time so callers pin the value in degrees.
       *
       * Default: `0` (no rotation, Excel's reference look). Set a
       * positive integer (e.g. `45`) to rotate the labels clockwise
       * (their right edge tilts down — typical "diagonal labels"
       * dashboard look on a column chart). Negative values rotate
       * counter-clockwise.
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
       * the OOXML schema, so the field round-trips on every chart family
       * that has axes (bar / column / line / area / scatter). Pie /
       * doughnut have no axes at all, so the field is silently dropped
       * on those families.
       *
       * Non-finite (`NaN`, `Infinity`) and non-numeric inputs drop at
       * write time so the writer never emits a token Excel's strict
       * validator would reject.
       */
      labelRotation?: number;
      /**
       * Tick-label font size in whole or half points. Maps to
       * `<c:catAx><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr></c:catAx>`
       * (or `<c:valAx>` for scatter / value axes) — Excel's "Format
       * Axis -> Font -> Size" knob applied to the tick labels. The
       * OOXML attribute is in 100ths of a point, so 12pt serializes as
       * `sz="1200"` and 10pt (Excel's reference default for tick
       * labels) as `sz="1000"`; the writer performs the conversion at
       * emit time and lands the value on the default-paragraph
       * `<a:defRPr>` slot.
       *
       * Mirrors {@link SheetChart.axes.x.axisTitleFontSize} for tick
       * labels — same `1..400`pt band the OOXML `ST_TextFontSize`
       * schema exposes, same 0.5pt half-step granularity Excel's UI
       * exposes, same out-of-range / non-finite drop semantics. Useful
       * for shrinking dense numeric tick labels on a value axis to fit
       * a tight chart frame, or bumping category-axis labels up to
       * match a presentation slide's typography.
       *
       * Default: omitted — the tick labels render at Excel's reference
       * size (10pt). Composes independently with
       * {@link labelRotation}: both fields land on the same
       * `<c:txPr>` body, so a single configuration call threads cleanly
       * through both knobs.
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
       * the OOXML schema, so the field round-trips on every chart family
       * that has axes (bar / column / line / area / scatter). Pie /
       * doughnut have no axes at all, so the field is silently dropped
       * on those families.
       */
      labelFontSize?: number;
      /**
       * Tick-label bold flag. Maps to
       * `<c:catAx><c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr></c:catAx>`
       * (or `<c:valAx>` for scatter / value axes) — Excel's "Format
       * Axis -> Font -> Bold" toggle applied to the tick labels. The
       * OOXML attribute is the `xsd:boolean` `b` on
       * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7);
       * the writer lands the value on the default-paragraph
       * `<a:defRPr>` slot inside the same `<c:txPr>` block that
       * carries {@link labelRotation} / {@link labelFontSize}.
       *
       * Mirrors {@link SheetChart.axes.x.axisTitleBold} for tick
       * labels — same `boolean | null` clone grammar, same silent drop
       * on `pie` / `doughnut` (no axes at all). Useful for emphasising
       * dashboard tick labels (e.g. bolding the bottom-axis category
       * names so they read as headers in a busy chart frame).
       *
       * Default: omitted — the tick labels render non-bold (the OOXML
       * default; the writer elides the `b` attribute when no value is
       * pinned). Set `true` to emit `b="1"` on the default-paragraph
       * slot; set `false` explicitly to pin the OOXML default `b="0"`
       * (functionally identical to omission, but useful when overriding
       * a templated chart that had bold pinned upstream).
       *
       * Composes independently with {@link labelRotation} /
       * {@link labelFontSize}: all three knobs land on the same
       * `<c:txPr>` body, so a single configuration call threads cleanly
       * through every tick-label knob the writer exposes.
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
       * the OOXML schema, so the field round-trips on every chart family
       * that has axes (bar / column / line / area / scatter). Pie /
       * doughnut have no axes at all, so the field is silently dropped
       * on those families.
       */
      labelBold?: boolean;
      /**
       * Tick-label italic flag. Maps to
       * `<c:catAx><c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr></c:catAx>`
       * (or `<c:valAx>` for scatter / value axes) — Excel's "Format
       * Axis -> Font -> Italic" toggle applied to the tick labels. The
       * OOXML attribute is the `xsd:boolean` `i` on
       * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7);
       * the writer lands the value on the default-paragraph
       * `<a:defRPr>` slot inside the same `<c:txPr>` block that
       * carries {@link labelRotation} / {@link labelFontSize} /
       * {@link labelBold}.
       *
       * Mirrors {@link SheetChart.axes.x.axisTitleItalic} for tick
       * labels — same `boolean | null` clone grammar, same silent drop
       * on `pie` / `doughnut` (no axes at all). Useful for italicising
       * dashboard tick labels (e.g. emphasising the bottom-axis
       * category names so they stand out against the chart frame
       * without bolding them).
       *
       * Default: omitted — the tick labels render non-italic (the
       * OOXML default; the writer elides the `i` attribute when no
       * value is pinned). Set `true` to emit `i="1"` on the
       * default-paragraph slot; set `false` explicitly to pin the
       * OOXML default `i="0"` (functionally identical to omission, but
       * useful when overriding a templated chart that had italic
       * pinned upstream).
       *
       * Composes independently with {@link labelRotation} /
       * {@link labelFontSize} / {@link labelBold}: all four knobs land
       * on the same `<c:txPr>` body, so a single configuration call
       * threads cleanly through every tick-label knob the writer
       * exposes.
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
       * the OOXML schema, so the field round-trips on every chart family
       * that has axes (bar / column / line / area / scatter). Pie /
       * doughnut have no axes at all, so the field is silently dropped
       * on those families.
       */
      labelItalic?: boolean;
      /**
       * Tick-label font color. Maps to
       * `<c:catAx><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>
       * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
       * </c:txPr></c:catAx>` (or `<c:valAx>` for scatter / value axes) —
       * Excel's "Format Axis -> Font -> Font Color" picker applied to
       * the tick labels. The OOXML `<a:srgbClr val=".."/>` carries the
       * 6-character uppercase hex sRGB color (`CT_SRgbColor`,
       * ECMA-376 Part 1, §20.1.2.3.32); the writer lands the fill on
       * the default-paragraph `<a:defRPr>` slot inside the same
       * `<c:txPr>` block that carries {@link labelRotation} /
       * {@link labelFontSize} / {@link labelBold} / {@link labelItalic}.
       *
       * Mirrors {@link SheetChart.axes.x.axisTitleColor} for tick
       * labels — same accept-with-or-without-`#` grammar (`"FF0000"` /
       * `"#FF0000"` / `"ff0000"` all collapse to the uppercase
       * canonical form), same `string | null` clone grammar
       * (`undefined` inherits, `null` drops, a hex string replaces).
       * Useful for tinting dashboard tick labels (e.g. dimming the
       * value-axis numeric ticks to a muted grey so the data series
       * read as the visual focus of a busy chart frame).
       *
       * Default: omitted — the tick labels render in Excel's reference
       * inherited theme color (no `<a:solidFill>` element, the writer
       * skips the fill block entirely). Pin a hex value to render the
       * tick labels in that color (e.g. `"1070CA"` for the dashboard
       * hero blue the issue-#136 example reaches for). The 8-character
       * `#RRGGBBAA` form is *not* accepted — alpha lives on
       * `<a:srgbClr><a:alpha val=".."/>` which is a separate runs-level
       * knob; pinning `labelColor` carries the RGB triple only.
       * Malformed inputs (wrong length, non-hex characters, alpha-
       * channel form) collapse to `undefined` so a stray non-hex token
       * never produces a malformed `<a:srgbClr>`.
       *
       * Composes independently with {@link labelRotation} /
       * {@link labelFontSize} / {@link labelBold} / {@link labelItalic}:
       * all five knobs land on the same `<c:txPr>` body, so a single
       * configuration call threads cleanly through every tick-label
       * knob the writer exposes.
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
       * the OOXML schema, so the field round-trips on every chart family
       * that has axes (bar / column / line / area / scatter). Pie /
       * doughnut have no axes at all, so the field is silently dropped
       * on those families.
       */
      labelColor?: string;
      /**
       * Tick-label underline flag. Maps to
       * `<c:catAx><c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr></c:catAx>`
       * (or `<c:valAx>` for scatter / value axes) — Excel's "Format
       * Axis -> Font -> Underline" toggle applied to the tick labels.
       * The OOXML attribute is the `ST_TextUnderlineType` enum on
       * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7);
       * the writer lands the value on the default-paragraph
       * `<a:defRPr>` slot inside the same `<c:txPr>` block that
       * carries {@link labelRotation} / {@link labelFontSize} /
       * {@link labelBold} / {@link labelItalic} / {@link labelColor}.
       *
       * Modeled as a boolean for symmetry with {@link labelBold} /
       * {@link labelItalic} and the axis-title counterpart
       * {@link SheetChart.axes.x.axisTitleUnderline}: `true` emits
       * `u="sng"` (Excel's UI checkbox — single line); absence and
       * explicit `false` collapse to omitting the attribute (the
       * OOXML default `"none"` collapses to absence, mirroring how
       * Excel's reference serialization emits a non-underlined tick
       * label). The non-UI variant `"dbl"` and the sixteen exotic
       * types (`"words"`, `"heavy"`, `"dotted"`, `"dottedHeavy"`,
       * `"dash"`, `"dashHeavy"`, `"dashLong"`, `"dashLongHeavy"`,
       * `"dotDash"`, `"dotDashHeavy"`, `"dotDotDash"`,
       * `"dotDotDashHeavy"`, `"wavy"`, `"wavyHeavy"`, `"wavyDbl"`)
       * are read-only — the writer emits only `"sng"` to keep the
       * surfaced shape consistent with what Excel's reference UI
       * authors, and the reader collapses every non-`"sng"` token to
       * `undefined` so a templated tick label that pinned a non-single
       * underline in raw OOXML round-trips lossless rather than
       * silently downgrading on re-emit.
       *
       * Useful for emphasising dashboard tick labels (e.g.
       * underlining the bottom-axis category names so they read as
       * inline links in a busy chart frame, or pairing with
       * {@link labelColor} to land an accented underlined tick on a
       * KPI dashboard).
       *
       * Default: omitted — the tick labels render non-underlined (the
       * OOXML default; the writer elides the `u` attribute when no
       * value is pinned). Set `true` to emit `u="sng"` on the
       * default-paragraph slot; set `false` explicitly to pin the
       * OOXML default behaviour (functionally identical to omission,
       * but useful when overriding a templated chart that had an
       * underline pinned upstream).
       *
       * Composes independently with {@link labelRotation} /
       * {@link labelFontSize} / {@link labelBold} /
       * {@link labelItalic} / {@link labelColor}: all six knobs land
       * on the same `<c:txPr>` body, so a single configuration call
       * threads cleanly through every tick-label knob the writer
       * exposes.
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
       * the OOXML schema, so the field round-trips on every chart family
       * that has axes (bar / column / line / area / scatter). Pie /
       * doughnut have no axes at all, so the field is silently dropped
       * on those families.
       */
      labelUnderline?: boolean;
      /**
       * Tick-label strikethrough flag. Maps to
       * `<c:catAx><c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr></c:catAx>`
       * (or `<c:valAx>` for scatter / value axes) — Excel's "Format
       * Axis -> Font -> Strikethrough" toggle applied to the tick
       * labels. The OOXML attribute is the `ST_TextStrikeType` enum on
       * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7);
       * the writer lands the value on the default-paragraph
       * `<a:defRPr>` slot inside the same `<c:txPr>` block that
       * carries {@link labelRotation} / {@link labelFontSize} /
       * {@link labelBold} / {@link labelItalic} / {@link labelColor} /
       * {@link labelUnderline}.
       *
       * Modeled as a boolean for symmetry with {@link labelBold} /
       * {@link labelItalic} and the axis-title counterpart
       * {@link SheetChart.axes.x.axisTitleStrike}: `true` emits
       * `strike="sngStrike"` (Excel's UI checkbox — single line);
       * absence and explicit `false` collapse to omitting the
       * attribute (the OOXML default `"noStrike"` collapses to absence,
       * mirroring how Excel's reference serialization emits a non-
       * strikethrough tick label). The non-UI variant `"dblStrike"` is
       * read-only — the writer emits only `"sngStrike"` to keep the
       * surfaced shape consistent with what Excel's reference UI
       * authors, and the reader collapses `"dblStrike"` to `undefined`
       * so a templated tick label that pinned the double-line variant
       * in raw OOXML round-trips lossless rather than silently
       * downgrading on re-emit.
       *
       * Useful for striking through dashboard tick labels (e.g.
       * crossing out an obsolete category-axis bucket on a snapshot
       * report, or pairing with {@link labelColor} to render a muted
       * strikethrough on a deprecated value-axis range).
       *
       * Default: omitted — the tick labels render non-strikethrough
       * (the OOXML default; the writer elides the `strike` attribute
       * when no value is pinned). Set `true` to emit
       * `strike="sngStrike"` on the default-paragraph slot; set `false`
       * explicitly to pin the OOXML default behaviour (functionally
       * identical to omission, but useful when overriding a templated
       * chart that had a strikethrough pinned upstream).
       *
       * Composes independently with {@link labelRotation} /
       * {@link labelFontSize} / {@link labelBold} /
       * {@link labelItalic} / {@link labelColor} /
       * {@link labelUnderline}: all seven knobs land on the same
       * `<c:txPr>` body, so a single configuration call threads
       * cleanly through every tick-label knob the writer exposes.
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
       * the OOXML schema, so the field round-trips on every chart family
       * that has axes (bar / column / line / area / scatter). Pie /
       * doughnut have no axes at all, so the field is silently dropped
       * on those families.
       */
      labelStrike?: boolean;
      /**
       * Axis tick-label font family / typeface. Maps to
       * `<c:catAx>` / `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` ->
       * `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/>
       * </a:defRPr></a:pPr></a:p></c:txPr>` — Excel's "Format Axis ->
       * Number / Font -> Font" picker scoped to the axis's tick
       * labels. The OOXML `<a:latin typeface=".."/>` element carries
       * the typeface name (`CT_TextFont`, ECMA-376 Part 1,
       * §21.1.2.3.7); the writer lands the element on the default-
       * paragraph `<a:defRPr>` of the axis-level `<c:txPr>` body.
       *
       * Accepts any non-empty string typeface name (e.g. `"Calibri"`,
       * `"Arial"`, `"Times New Roman"`); the writer trims surrounding
       * whitespace and emits the trimmed value verbatim (XML-escaped)
       * so Excel can resolve the named font from the workbook's font
       * scheme or the host system's installed fonts. Empty /
       * whitespace-only strings and non-string tokens collapse to
       * `undefined` so the writer skips the entire `<a:latin>`
       * element and the tick labels inherit Excel's reference theme
       * typeface.
       *
       * Default: omitted — the tick labels render in Excel's
       * reference theme typeface (no `<a:latin>` element, the writer
       * skips the element entirely). Pin a typeface name to render
       * the labels in that font.
       *
       * Composes independently with {@link labelRotation} /
       * {@link labelFontSize} / {@link labelBold} /
       * {@link labelItalic} / {@link labelColor} /
       * {@link labelUnderline} / {@link labelStrike}: all eight knobs
       * land on the same `<c:txPr>` body, so a single configuration
       * call threads cleanly through every tick-label knob the
       * writer exposes.
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>`
       * per the OOXML schema, so the field round-trips on every
       * chart family that has axes (bar / column / line / area /
       * scatter). Pie / doughnut have no axes at all, so the field
       * is silently dropped on those families.
       */
      labelFontFamily?: string;
      /**
       * Reverse the axis plotting order. Maps to
       * `<c:scaling><c:orientation val="maxMin"/></c:scaling>` —
       * Excel's "Categories in reverse order" / "Values in reverse
       * order" toggle. Default: `false` (the OOXML `"minMax"` default).
       *
       * On a category axis, reversing flips the order in which
       * categories are drawn (right-to-left on a column chart, top-to-
       * bottom on a bar chart). On a value axis, reversing flips the
       * numeric direction so the maximum sits at the origin and the
       * minimum at the far end. Useful when porting templates that
       * pin a specific reading direction (e.g. dates on a horizontal
       * bar chart with the most recent at the top).
       */
      reverse?: boolean;
      /**
       * Show every Nth tick label on a category axis. `1` (the OOXML
       * default) shows every label; `2` shows every other one; `3`
       * shows every third, and so on. Maps to
       * `<c:catAx><c:tickLblSkip val="N"/></c:catAx>`. Only meaningful
       * for bar / column / line / area charts (whose X axis is
       * `<c:catAx>`); silently ignored for scatter (both axes are
       * value axes) and pie / doughnut (no axes at all). Accepted
       * range: positive integers 1..32767 (the OOXML
       * `ST_SkipIntervals` schema). Values outside the range or
       * non-positive are dropped at write time.
       */
      tickLblSkip?: number;
      /**
       * Show every Nth tick mark on a category axis. Same `1`-default
       * semantics as {@link tickLblSkip} but for the short tick lines
       * Excel paints alongside each label. Maps to
       * `<c:catAx><c:tickMarkSkip val="N"/></c:catAx>`. Same
       * scope-restriction as `tickLblSkip` — category axes only.
       */
      tickMarkSkip?: number;
      /**
       * Distance between the tick labels and the axis line on a
       * category axis, expressed as a percentage of the default
       * spacing. `100` (the OOXML default) renders Excel's reference
       * spacing; lower values pull the labels in towards the axis,
       * higher values push them out. Maps to
       * `<c:catAx><c:lblOffset val="N"/></c:catAx>`. Only meaningful
       * for bar / column / line / area charts (whose X axis is
       * `<c:catAx>`); silently ignored for scatter (both axes are
       * value axes) and pie / doughnut (no axes at all). Accepted
       * range: `0..1000` (the OOXML `ST_LblOffsetPercent` schema).
       * Values outside the range are dropped at write time.
       */
      lblOffset?: number;
      /**
       * Suppress Excel's automatic multi-level category labels. Maps
       * to `<c:catAx><c:noMultiLvlLbl val=".."/></c:catAx>`. The OOXML
       * default `false` (Excel groups labels into tiers when the
       * category range spans multiple columns / rows); set `true` to
       * flatten every category into a single line of labels regardless
       * of the source range's shape. Mirrors Excel's "Format Axis ->
       * Multi-level Category Labels" checkbox (the checkbox is the
       * inverse — checked means tiered labels, i.e.
       * `noMultiLvlLbl: false`).
       *
       * Only meaningful for bar / column / line / area charts (whose X
       * axis is `<c:catAx>`); silently ignored for scatter (both axes
       * are value axes) and pie / doughnut (no axes at all). The OOXML
       * schema places the element on `CT_CatAx` only — `CT_ValAx`,
       * `CT_DateAx`, and `CT_SerAx` reject it.
       */
      noMultiLvlLbl?: boolean;
      /**
       * Automatic axis-type detection on a category axis. Maps to
       * `<c:catAx><c:auto val=".."/></c:catAx>` (CT_CatAx, ECMA-376
       * Part 1, §21.2.2.7). The OOXML default `true` lets Excel inspect
       * the axis labels and decide at render time whether to treat the
       * axis as a discrete category axis or a chronological date axis.
       *
       * Set `false` to pin the axis as a literal category axis — Excel
       * keeps every label as-is regardless of whether the cells parse as
       * dates or numerics. Useful when porting a template that pins the
       * "Text axis" radio button under "Format Axis -> Axis Options ->
       * Axis Type" so a date-shaped category range still renders as a
       * flat category axis without any chronological grouping.
       *
       * Only meaningful for bar / column / line / area charts (whose X
       * axis is `<c:catAx>`); silently ignored for scatter (both axes
       * are value axes) and pie / doughnut (no axes at all). The OOXML
       * schema places the element on `CT_CatAx` only — `CT_ValAx`,
       * `CT_DateAx`, and `CT_SerAx` reject it. Mirrors how the writer
       * always emits Excel's reference `<c:auto val="1"/>` on a stock
       * chart and only flips the value when the caller explicitly
       * pins `auto: false`.
       */
      auto?: boolean;
      /**
       * Horizontal alignment of the tick labels on a category axis —
       * `"ctr"` (center, the OOXML default), `"l"` (left), or `"r"`
       * (right). Maps to `<c:catAx><c:lblAlgn val=".."/></c:catAx>`.
       * Useful when category labels are wrapped onto multiple lines
       * and the default centered alignment looks ragged against a
       * column chart's left-aligned bars. Excel's UI exposes the
       * three presets under "Format Axis -> Alignment" on a category
       * axis only.
       *
       * Only meaningful for bar / column / line / area charts (whose X
       * axis is `<c:catAx>`); silently ignored for scatter (both axes
       * are value axes) and pie / doughnut (no axes at all). The OOXML
       * schema (`ST_LblAlgn`) restricts the value to the three tokens
       * above; unknown tokens are dropped at write time. See
       * {@link ChartAxisLabelAlign}.
       */
      lblAlgn?: ChartAxisLabelAlign;
      /**
       * Hide the entire axis (line, tick marks, tick labels). Maps to
       * `<c:catAx><c:delete val="1"/></c:catAx>` (or the matching
       * `<c:valAx>` element on scatter). Default: `false` — Excel
       * paints the axis. Set `true` to collapse a noisy axis off a
       * sparkline-style dashboard tile.
       *
       * Excel still reserves the layout slot the axis would have
       * occupied, so a hidden category axis on a column chart leaves a
       * thin gap at the bottom of the plot area where the labels would
       * have rendered — pair with `<c:layout>` overrides on the parent
       * `<c:plotArea>` if you need to reclaim that space (hucre does
       * not surface a layout knob today; the writer falls back to
       * Excel's auto-layout in either case).
       *
       * The flag is silently ignored on `pie` / `doughnut` charts
       * because the OOXML schema places no axes on those families.
       */
      hidden?: boolean;
      /**
       * Where the perpendicular axis crosses this axis along its own
       * range. Maps to `<c:catAx><c:crosses val=".."/></c:catAx>` (or
       * `<c:valAx>` for scatter). Default: `"autoZero"` — Excel's
       * reference serialization, the perpendicular axis crosses at zero
       * on a value axis or at the first category on a category axis.
       *
       * Set `"min"` / `"max"` to pin the perpendicular axis to the low
       * / high end of this axis (Excel's "Format Axis -> Axis Options
       * -> Vertical axis crosses" toggle). Mutually exclusive with
       * {@link crossesAt} — when both are set the writer favours
       * `crossesAt`. Silently ignored on `pie` / `doughnut` charts
       * because the OOXML schema places no axes on those families. See
       * {@link ChartAxisCrosses}.
       */
      crosses?: ChartAxisCrosses;
      /**
       * Numeric crossing position. Maps to
       * `<c:catAx><c:crossesAt val=".."/></c:catAx>` (or `<c:valAx>` for
       * scatter). When set, takes precedence over {@link crosses}
       * because the OOXML schema (`CT_CatAx` / `CT_ValAx`) places the
       * two elements in an XSD choice — only one may appear at a time.
       *
       * The literal value is preserved (including `0`, which is
       * distinct from the `"autoZero"` default — `crossesAt: 0` pins
       * the crossing point to the numeric value zero, while `crosses:
       * "autoZero"` defers to Excel's auto-placement). Non-finite
       * inputs (`NaN`, `Infinity`) drop at write time. Silently ignored
       * on `pie` / `doughnut` charts.
       */
      crossesAt?: number;
      /**
       * Built-in display-unit preset for the X axis. Maps to
       * `<c:valAx><c:dispUnits><c:builtInUnit val=".."/></c:dispUnits></c:valAx>`.
       *
       * Only meaningful for `scatter` charts — both axes there are value
       * axes (`<c:valAx>`), so `<c:dispUnits>` slots onto the X axis as
       * well. The OOXML schema places the element exclusively on
       * `CT_ValAx`, so the writer drops the field on every other family
       * (the X axis on bar / column / line / area is a category axis,
       * which rejects `<c:dispUnits>`; pie / doughnut have no axes at
       * all). Pass a {@link ChartAxisDispUnit} preset directly as a
       * shorthand for `{ unit: ".." }`; pass an object to opt into the
       * automatic unit annotation via `showLabel: true`.
       *
       * See {@link ChartAxisDispUnits} for the surfaced shape and
       * {@link ChartAxisDispUnit} for the accepted preset tokens.
       */
      dispUnits?: ChartAxisDispUnits | ChartAxisDispUnit;
      /**
       * Cross-between mode for the X axis. Maps to
       * `<c:valAx><c:crossBetween val=".."/></c:valAx>`.
       *
       * Only meaningful for `scatter` charts — both axes there are
       * value axes (`<c:valAx>`), so `<c:crossBetween>` slots onto the
       * X axis as well. The OOXML schema places the element exclusively
       * on `CT_ValAx`, so the writer drops the field on every other
       * family (the X axis on bar / column / line / area is a category
       * axis, which rejects `<c:crossBetween>`; pie / doughnut have no
       * axes at all). See {@link ChartAxisCrossBetween}.
       */
      crossBetween?: ChartAxisCrossBetween;
    };
    /** Value axis. */
    y?: {
      title?: string;
      /**
       * Rotation of the value-axis title in whole degrees. Maps to
       * `<c:valAx><c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.axisTitleRotation} for the
       * value axis — see that field for the full semantics. The OOXML
       * `rot` attribute is in 60000ths of a degree; the writer converts
       * at emit time so callers pin the value in degrees. Range:
       * `-90..90`.
       *
       * Useful for standing the Y-axis title vertically (a common
       * dashboard pattern that reclaims horizontal real estate next to
       * the value labels) or pinning a custom angle on a long axis
       * title that would otherwise crowd the plot area. Silently
       * dropped on `pie` / `doughnut` charts (no axes at all) and on
       * any axis whose `title` is unset (no `<c:title>` block to host
       * the rotation).
       */
      axisTitleRotation?: number;
      /**
       * Value-axis title font size in whole or half points. Maps to
       * `<c:valAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
       * <a:r><a:rPr sz="N"/></a:r></a:p></c:rich></c:tx></c:title></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.axisTitleFontSize} for the
       * value axis — see that field for the full semantics. The OOXML
       * `sz` attribute is in 100ths of a point; the writer converts at
       * emit time so callers pin the value in points. Range: `1..400`.
       *
       * Useful for shrinking a long Y-axis unit label so it fits a
       * tight chart frame, or bumping the Y-axis title up to match
       * the chart-level title's typography on a presentation slide.
       * Silently dropped on `pie` / `doughnut` charts (no axes at all)
       * and on any axis whose `title` is unset (no `<c:title>` block
       * to host the size).
       */
      axisTitleFontSize?: number;
      /**
       * Value-axis title bold flag. Maps to
       * `<c:valAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
       * <a:r><a:rPr b=".."/></a:r></a:p></c:rich></c:tx></c:title></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.axisTitleBold} for the value
       * axis — see that field for the full semantics. The OOXML `b`
       * attribute is the `xsd:boolean` bold flag on
       * `CT_TextCharacterProperties`; the writer lands the value on
       * both the default-paragraph `<a:defRPr>` and the literal run's
       * `<a:rPr>` so a re-parse picks the flag up off either canonical
       * slot.
       *
       * Useful for matching the chart-level title's emphasis on the Y-
       * axis caption (e.g. bolding "Revenue (USD)" to align with a
       * dashboard's heavy-weight typography). Silently dropped on
       * `pie` / `doughnut` charts (no axes at all) and on any axis
       * whose `title` is unset (no `<c:title>` block to host the
       * flag).
       */
      axisTitleBold?: boolean;
      /**
       * Value-axis title italic flag. Maps to
       * `<c:valAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
       * <a:r><a:rPr i=".."/></a:r></a:p></c:rich></c:tx></c:title></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.axisTitleItalic} for the value
       * axis — see that field for the full semantics. The OOXML
       * attribute is the `xsd:boolean` `i` on
       * `CT_TextCharacterProperties`; the writer lands the value on
       * both the default-paragraph `<a:defRPr>` and the literal run's
       * `<a:rPr>` so a re-parse picks the flag up off either canonical
       * slot.
       *
       * Useful for italicising a Y-axis unit label to mark it as a
       * derived measure (a common dashboard pattern that distinguishes
       * an aggregated / computed axis from a raw category axis).
       * Silently dropped on `pie` / `doughnut` charts (no axes at all)
       * and on any axis whose `title` is unset (no `<c:title>` block
       * to host the flag).
       */
      axisTitleItalic?: boolean;
      /**
       * Value-axis title font color. Maps to
       * `<c:valAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr>
       * <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr>
       * </a:pPr><a:r><a:rPr><a:solidFill><a:srgbClr val="RRGGBB"/>
       * </a:solidFill></a:rPr></a:r></a:p></c:rich></c:tx></c:title>
       * </c:valAx>`. Mirrors {@link SheetChart.axes.x.axisTitleColor}
       * for the value axis — see that field for the full semantics.
       * The OOXML `<a:srgbClr val=".."/>` carries the 6-character
       * uppercase hex sRGB color; the writer lands the fill on both
       * the default-paragraph `<a:defRPr>` and the literal run's
       * `<a:rPr>` so a re-parse picks the color up off either
       * canonical slot.
       *
       * Useful for tinting a Y-axis unit label to match a dashboard
       * accent color (e.g. green for revenue, red for expense) without
       * touching the chart title or axis tick labels. Silently dropped
       * on `pie` / `doughnut` charts (no axes at all) and on any axis
       * whose `title` is unset (no `<c:title>` block to host the
       * fill).
       */
      axisTitleColor?: string;
      /**
       * Value-axis title strikethrough flag. Maps to
       * `<c:valAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
       * <a:r><a:rPr strike=".."/></a:r></a:p></c:rich></c:tx></c:title></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.axisTitleStrike} for the value
       * axis — see that field for the full semantics. The OOXML
       * attribute is the `ST_TextStrikeType` enum on
       * `CT_TextCharacterProperties`; the writer emits only the UI
       * variant `"sngStrike"` and lands it on both the default-paragraph
       * `<a:defRPr>` and the literal run's `<a:rPr>` so a re-parse picks
       * the flag up off either canonical slot.
       *
       * Useful for marking a Y-axis unit label as "before" or
       * "deprecated" in dashboard tile typography without touching the
       * chart title. Silently dropped on `pie` / `doughnut` charts (no
       * axes at all) and on any axis whose `title` is unset (no
       * `<c:title>` block to host the flag).
       */
      axisTitleStrike?: boolean;
      /**
       * Value-axis title underline flag. Maps to
       * `<c:valAx><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
       * <a:r><a:rPr u=".."/></a:r></a:p></c:rich></c:tx></c:title></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.axisTitleUnderline} for the value
       * axis — see that field for the full semantics. The OOXML
       * attribute is the `ST_TextUnderlineType` enum on
       * `CT_TextCharacterProperties`; the writer emits only the UI
       * variant `"sng"` and lands it on both the default-paragraph
       * `<a:defRPr>` and the literal run's `<a:rPr>` so a re-parse picks
       * the flag up off either canonical slot.
       *
       * Useful for underlining a Y-axis unit label to highlight it as
       * the dashboard's primary metric (a common dashboard pattern that
       * draws the eye to the rendered value over the categorical
       * sweep). Silently dropped on `pie` / `doughnut` charts (no axes
       * at all) and on any axis whose `title` is unset (no `<c:title>`
       * block to host the flag).
       */
      axisTitleUnderline?: boolean;
      /**
       * Axis title font family / typeface. Maps to
       * `<c:catAx>` / `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` ->
       * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
       * typeface=".."/></a:defRPr></a:pPr><a:r><a:rPr><a:latin
       * typeface=".."/></a:rPr></a:r></a:p></c:rich></c:tx></c:title>` —
       * Excel's "Format Axis Title -> Font -> Font" picker. The OOXML
       * `<a:latin typeface=".."/>` element carries the typeface name
       * (`CT_TextFont`, ECMA-376 Part 1, §21.1.2.3.7); the writer
       * lands the element on both the default-paragraph `<a:defRPr>`
       * and the literal run's `<a:rPr>` so a re-parse picks the
       * typeface up off either canonical slot — Excel keeps the two
       * values in sync.
       *
       * Accepts any non-empty string typeface name (e.g. `"Calibri"`,
       * `"Arial"`, `"Times New Roman"`); the writer trims surrounding
       * whitespace and emits the trimmed value verbatim (XML-escaped)
       * so Excel can resolve the named font from the workbook's font
       * scheme or the host system's installed fonts. Empty /
       * whitespace-only strings and non-string tokens collapse to
       * `undefined` so the writer skips the entire `<a:latin>`
       * element and the title inherits Excel's reference theme
       * typeface (Calibri Light from the default Office theme).
       *
       * Mirrors {@link SheetChart.titleFontFamily} for axis titles —
       * same canonical-slot pair, same drop-on-empty semantics, same
       * `string | null` clone grammar so a single configuration call
       * threads cleanly through both the chart title and either axis
       * title without bookkeeping the canonical OOXML slots. Mirrors
       * {@link axisTitleRotation} / {@link axisTitleFontSize} /
       * {@link axisTitleBold} / {@link axisTitleItalic} /
       * {@link axisTitleColor} / {@link axisTitleStrike} /
       * {@link axisTitleUnderline} for the same `<c:title>` body, so
       * all eight axis-title typography knobs compose freely on the
       * same axis.
       *
       * Silently dropped when the axis renders no title (the
       * `<c:title>` element is absent in either case) and on `pie` /
       * `doughnut` charts (no axes at all).
       *
       * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
       * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>`
       * shape per the OOXML schema, so the field round-trips on
       * every chart family that has axes (bar / column / line /
       * area / scatter).
       */
      axisTitleFontFamily?: string;
      /**
       * Value-axis-title overlay flag. Same canonical `<c:title>
       * <c:overlay val=".."/></c:title>` slot and grammar as
       * {@link SheetChart.axes.x.axisTitleOverlay} — the field
       * round-trips on `<c:valAx>` exactly as it does on `<c:catAx>`.
       * See the X-axis variant for the full semantics.
       */
      axisTitleOverlay?: boolean;
      /**
       * Value-axis-title manual placement. Same canonical
       * `<c:title><c:layout><c:manualLayout>...</c:manualLayout>
       * </c:layout></c:title>` slot and grammar as
       * {@link SheetChart.axes.x.axisTitleLayout} — the field
       * round-trips on `<c:valAx>` exactly as it does on `<c:catAx>`.
       * See the X-axis variant for the full semantics.
       */
      axisTitleLayout?: ChartManualLayout;
      /**
       * Value-axis-title background fill (solid). Same canonical
       * `<c:title><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
       * </a:solidFill></c:spPr></c:title>` slot and grammar as
       * {@link SheetChart.axes.x.axisTitleFillColor} — the field
       * round-trips on `<c:valAx>` exactly as it does on `<c:catAx>`.
       * See the X-axis variant for the full semantics.
       */
      axisTitleFillColor?: string;
      /**
       * Value-axis-title border (line stroke) solid color. Same
       * canonical `<c:title><c:spPr><a:ln><a:solidFill><a:srgbClr
       * val="RRGGBB"/></a:solidFill></a:ln></c:spPr></c:title>` slot
       * and grammar as {@link SheetChart.axes.x.axisTitleBorderColor}
       * — the field round-trips on `<c:valAx>` exactly as it does on
       * `<c:catAx>`. See the X-axis variant for the full semantics.
       */
      axisTitleBorderColor?: string;
      /**
       * Value-axis-title border (stroke) thickness in points. Same
       * canonical `<c:title><c:spPr><a:ln w="EMU">` slot and grammar
       * as {@link SheetChart.axes.x.axisTitleBorderWidth} — the field
       * round-trips on `<c:valAx>` exactly as it does on `<c:catAx>`.
       * See the X-axis variant for the full semantics.
       */
      axisTitleBorderWidth?: number;
      /**
       * Value-axis-title border (stroke) preset dash pattern. Same
       * canonical `<c:title><c:spPr><a:ln><a:prstDash val=".."/>` slot
       * and grammar as {@link SheetChart.axes.x.axisTitleBorderDash} —
       * the field round-trips on `<c:valAx>` exactly as it does on
       * `<c:catAx>`. See the X-axis variant for the full semantics.
       */
      axisTitleBorderDash?: ChartBorderDash;
      gridlines?: ChartAxisGridlines;
      scale?: ChartAxisScale;
      numberFormat?: ChartAxisNumberFormat;
      /**
       * Major tick-mark style for the value axis. Maps to
       * `<c:valAx><c:majorTickMark val=".."/></c:valAx>`. Default:
       * `"out"`. See {@link ChartAxisTickMark}.
       */
      majorTickMark?: ChartAxisTickMark;
      /**
       * Minor tick-mark style for the value axis. Maps to
       * `<c:valAx><c:minorTickMark val=".."/></c:valAx>`. Default:
       * `"none"`. See {@link ChartAxisTickMark}.
       */
      minorTickMark?: ChartAxisTickMark;
      /**
       * Tick-label position for the value axis. Maps to
       * `<c:valAx><c:tickLblPos val=".."/></c:valAx>`. Default:
       * `"nextTo"`. See {@link ChartAxisTickLabelPosition}.
       */
      tickLblPos?: ChartAxisTickLabelPosition;
      /**
       * Tick-label rotation in degrees for the value axis. Maps to
       * `<c:valAx><c:txPr><a:bodyPr rot="N"/></c:txPr></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.labelRotation} for the value
       * axis — see that field for the full semantics. The OOXML `rot`
       * attribute is in 60000ths of a degree; the writer converts at
       * emit time so callers pin the value in degrees. Range: `-90..90`.
       *
       * Useful for tilting Y-axis number labels when a tight chart
       * frame would otherwise crowd them, or for rotating the X-axis
       * value labels on a scatter chart whose long category strings
       * sit on the value axis. Silently dropped on `pie` / `doughnut`
       * charts (no axes at all).
       */
      labelRotation?: number;
      /**
       * Tick-label font size in points for the value axis. Maps to
       * `<c:valAx><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.labelFontSize} for the value
       * axis — see that field for the full semantics. The OOXML `sz`
       * attribute is in 100ths of a point; the writer converts at emit
       * time so callers pin the value in points. Range: `1..400`pt
       * (the OOXML `ST_TextFontSize` band), with 0.5pt granularity.
       *
       * Useful for shrinking dense numeric tick labels (e.g. wide
       * currency totals) so they fit a tight Y-axis column, or
       * bumping the value-axis labels up to match a chart-level
       * typography pin. Silently dropped on `pie` / `doughnut`
       * charts (no axes at all).
       */
      labelFontSize?: number;
      /**
       * Tick-label bold flag for the value axis. Maps to
       * `<c:valAx><c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.labelBold} for the value
       * axis — see that field for the full semantics. The OOXML `b`
       * attribute is the `xsd:boolean` bold flag on
       * `CT_TextCharacterProperties`; the writer emits `1` / `0` at
       * the canonical slot. Absence collapses to the OOXML default
       * (the writer omits the attribute) so a fresh chart inherits
       * the theme-default tick-label weight.
       *
       * Useful for emphasising the value-axis labels on a dashboard
       * (e.g. bolding the Y-axis numeric totals so they read as
       * headers in a busy chart frame). Silently dropped on `pie` /
       * `doughnut` charts (no axes at all).
       */
      labelBold?: boolean;
      /**
       * Tick-label italic flag for the value axis. Maps to
       * `<c:valAx><c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.labelItalic} for the value
       * axis — see that field for the full semantics. The OOXML `i`
       * attribute is the `xsd:boolean` italic flag on
       * `CT_TextCharacterProperties`; the writer emits `1` / `0` at
       * the canonical slot. Absence collapses to the OOXML default
       * (the writer omits the attribute) so a fresh chart inherits
       * the theme-default tick-label slant.
       *
       * Useful for emphasising the value-axis labels on a dashboard
       * (e.g. italicising the Y-axis numeric totals so they read as
       * a stylistic accent in a busy chart frame). Silently dropped
       * on `pie` / `doughnut` charts (no axes at all).
       */
      labelItalic?: boolean;
      /**
       * Tick-label font color for the value axis. Maps to
       * `<c:valAx><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>
       * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
       * </c:txPr></c:valAx>`. Mirrors
       * {@link SheetChart.axes.x.labelColor} for the value axis — see
       * that field for the full semantics. The OOXML
       * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
       * sRGB color; the writer lands the fill on the default-paragraph
       * `<a:defRPr>` slot inside the same `<c:txPr>` block that
       * carries every other tick-label typography knob.
       *
       * Useful for dimming or accenting the Y-axis numeric ticks on a
       * dashboard (e.g. a muted grey so the data series read as the
       * visual focus, or a brand accent so the totals carry a
       * stylistic tint). Silently dropped on `pie` / `doughnut` charts
       * (no axes at all).
       */
      labelColor?: string;
      /**
       * Tick-label underline flag for the value axis. Maps to
       * `<c:valAx><c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.labelUnderline} for the value
       * axis — see that field for the full semantics. The OOXML `u`
       * attribute is the `ST_TextUnderlineType` enum on
       * `CT_TextCharacterProperties`; the writer emits only the UI
       * variant `"sng"` (single line) when the input is `true`.
       * Absence and explicit `false` collapse to omitting the
       * attribute (the OOXML default `"none"` collapses to absence)
       * so a fresh chart inherits Excel's reference non-underlined
       * tick labels.
       *
       * Useful for emphasising the value-axis labels on a dashboard
       * (e.g. underlining the Y-axis numeric totals so they read as
       * inline links in a busy chart frame). Silently dropped on
       * `pie` / `doughnut` charts (no axes at all).
       */
      labelUnderline?: boolean;
      /**
       * Tick-label strikethrough flag for the value axis. Maps to
       * `<c:valAx><c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr></c:valAx>`.
       * Mirrors {@link SheetChart.axes.x.labelStrike} for the value
       * axis — see that field for the full semantics. The OOXML
       * `strike` attribute is the `ST_TextStrikeType` enum on
       * `CT_TextCharacterProperties`; the writer emits only the UI
       * variant `"sngStrike"` (single line) when the input is `true`.
       * Absence and explicit `false` collapse to omitting the
       * attribute (the OOXML default `"noStrike"` collapses to
       * absence) so a fresh chart inherits Excel's reference non-
       * strikethrough tick labels.
       *
       * Useful for striking through obsolete value-axis labels on a
       * dashboard (e.g. crossing out a deprecated numeric range on a
       * snapshot report). Silently dropped on `pie` / `doughnut` charts
       * (no axes at all).
       */
      labelStrike?: boolean;
      /**
       * Value-axis tick-label font family / typeface. Same canonical-
       * slot pair and grammar as {@link SheetChart.axes.x.labelFontFamily} —
       * the field round-trips on `<c:valAx>` exactly as it does on
       * `<c:catAx>`. See the X-axis variant for the full semantics.
       */
      labelFontFamily?: string;
      /**
       * Hide the entire value axis (line, tick marks, tick labels).
       * Maps to `<c:valAx><c:delete val="1"/></c:valAx>`. Default:
       * `false`. See {@link SheetChart.axes.x.hidden} for the full
       * semantics — the value-axis flag mirrors the X-axis flag.
       */
      hidden?: boolean;
      /**
       * Reverse the value axis plotting order. Maps to
       * `<c:valAx><c:scaling><c:orientation val="maxMin"/></c:scaling></c:valAx>`.
       * Default: `false` (the OOXML `"minMax"` default).
       *
       * Mirrors {@link SheetChart.axes.x.reverse} for the value axis —
       * setting `true` flips the numeric direction so the maximum sits
       * at the origin and the minimum at the far end.
       */
      reverse?: boolean;
      /**
       * Where the perpendicular axis crosses the value axis along its
       * own range. Maps to `<c:valAx><c:crosses val=".."/></c:valAx>`.
       * Default: `"autoZero"`. Mirrors
       * {@link SheetChart.axes.x.crosses} for the value axis. Mutually
       * exclusive with {@link crossesAt} — when both are set the writer
       * favours `crossesAt`. See {@link ChartAxisCrosses}.
       */
      crosses?: ChartAxisCrosses;
      /**
       * Numeric crossing position for the value axis. Maps to
       * `<c:valAx><c:crossesAt val=".."/></c:valAx>`. Mirrors
       * {@link SheetChart.axes.x.crossesAt} — when set, takes
       * precedence over {@link crosses}.
       */
      crossesAt?: number;
      /**
       * Built-in display-unit preset for the value axis. Maps to
       * `<c:valAx><c:dispUnits><c:builtInUnit val=".."/></c:dispUnits></c:valAx>`.
       *
       * Excel exposes the same dropdown under "Format Axis -> Display
       * units" — every numeric tick label is divided by the preset's
       * scale before being rendered, so a chart whose source range
       * stores raw amounts (e.g. `1_500_000`) can show compact tick
       * labels (`1.5` with an optional "Millions" annotation) without
       * modifying the underlying cells. The OOXML schema places the
       * element exclusively on `CT_ValAx`, so the writer drops the
       * field on `pie` / `doughnut` charts (no axes at all). Pass a
       * {@link ChartAxisDispUnit} preset directly as a shorthand for
       * `{ unit: ".." }`; pass an object to opt into the automatic
       * unit annotation via `showLabel: true`.
       *
       * See {@link ChartAxisDispUnits} for the surfaced shape and
       * {@link ChartAxisDispUnit} for the accepted preset tokens.
       */
      dispUnits?: ChartAxisDispUnits | ChartAxisDispUnit;
      /**
       * Cross-between mode for the value axis. Maps to
       * `<c:valAx><c:crossBetween val=".."/></c:valAx>`.
       *
       * The Y axis is a value axis on every chart family that has axes —
       * bar / column / line / area / scatter — so the override always
       * takes effect on those families. Pie / doughnut have no axes at
       * all, so the field is silently dropped on those families. See
       * {@link ChartAxisCrossBetween}.
       */
      crossBetween?: ChartAxisCrossBetween;
    };
  };
}

// ── Charts ────────────────────────────────────────────────────────

/**
 * Chart kind reported by {@link Chart.kinds}. Mirrors the OOXML
 * chart-type element local names (`c:barChart`, `c:lineChart`, ...).
 * A single chart can mix multiple kinds (combo chart), in which case
 * every kind appears in `kinds` in the order it's declared.
 */
export type ChartKind =
  | "bar"
  | "bar3D"
  | "line"
  | "line3D"
  | "pie"
  | "pie3D"
  | "doughnut"
  | "area"
  | "area3D"
  | "scatter"
  | "bubble"
  | "radar"
  | "surface"
  | "surface3D"
  | "stock"
  | "ofPie";

/**
 * A single series surfaced from a parsed chart.
 *
 * Field semantics mirror what {@link ChartSeries} accepts on the write
 * side, so a `ChartSeriesInfo` returned by {@link Chart.series} can be
 * used as the basis for cloning a chart with new bindings.
 *
 * `valuesRef` and `categoriesRef` are the raw `<c:f>` formula strings
 * extracted from the chart XML — typically sheet-qualified A1 ranges
 * like `"Sheet1!$B$2:$B$10"`. They may be `undefined` when the series
 * embeds literal numbers (`<c:numLit>`) instead of referencing a range.
 */
export interface ChartSeriesInfo {
  /** Chart kind that owns this series (matches {@link Chart.kinds}). */
  kind: ChartKind;
  /** 0-based position inside the chart-type element. */
  index: number;
  /** Series name pulled from `<c:tx>` (literal `<c:v>` or strRef cache). */
  name?: string;
  /** Raw `<c:f>` for `<c:val>` / `<c:yVal>`. */
  valuesRef?: string;
  /** Raw `<c:f>` for `<c:cat>` / `<c:xVal>`. */
  categoriesRef?: string;
  /** 6-digit RGB hex from `<c:spPr><a:solidFill><a:srgbClr val>`. */
  color?: string;
  /**
   * Series-level data labels parsed from the `<c:ser><c:dLbls>` block.
   * Falls back to the chart-level {@link Chart.dataLabels} when this
   * series carries no override of its own.
   */
  dataLabels?: ChartDataLabelsInfo;
  /**
   * Smoothed-line flag pulled from `<c:ser><c:smooth val=".."/>`.
   * Surfaces only on `line` / `scatter` series — the OOXML schema places
   * `<c:smooth>` exclusively on `CT_LineSer` and `CT_ScatterSer`. `false`
   * collapses to `undefined` because it matches the OOXML default and
   * round-trips identically with absence of the field.
   */
  smooth?: boolean;
  /**
   * Line stroke pulled from `<c:ser><c:spPr><a:ln>` — preset dash
   * pattern and width in points. Surfaces only on `line` / `scatter`
   * series so a dashed-stroke template round-trips through
   * `parseChart` → {@link cloneChart} → `writeXlsx`. Field semantics
   * mirror the write-side {@link ChartLineStroke}, so the value can be
   * fed straight into a clone without transformation.
   */
  stroke?: ChartLineStroke;
  /**
   * Marker styling parsed from `<c:ser><c:marker>`. Surfaces only on
   * `line` / `scatter` series — the OOXML schema places `<c:marker>`
   * exclusively on `CT_LineSer` and `CT_ScatterSer`. Empty marker
   * blocks (no symbol, size, or color) collapse to `undefined` so a
   * round-trip keeps the read-side shape minimal. Field semantics
   * mirror the write-side {@link ChartMarker}, so the value can be
   * fed straight into {@link cloneChart} without transformation.
   */
  marker?: ChartMarker;
  /**
   * Invert-if-negative flag pulled from
   * `<c:ser><c:invertIfNegative val=".."/>`. Surfaces only on `bar`
   * (and `bar3D`) series — the OOXML schema places
   * `<c:invertIfNegative>` exclusively on `CT_BarSer` / `CT_Bar3DSer`.
   * `false` collapses to `undefined` because it matches the OOXML
   * default and round-trips identically with absence of the field.
   */
  invertIfNegative?: boolean;
  /**
   * Slice explosion (in percent of the radius) pulled from
   * `<c:ser><c:explosion val=".."/>`. Surfaces only on `pie`,
   * `pie3D`, `doughnut`, and `ofPie` series — the OOXML schema
   * places `<c:explosion>` exclusively on `CT_PieSer` (which is
   * shared across the pie family via `EG_PieSer`). The OOXML
   * default `0` collapses to `undefined` because absence and `0`
   * round-trip identically through the writer's elision logic.
   */
  explosion?: number;
}

/**
 * Read-side mirror of {@link ChartDataLabels}. Exposes the same toggle
 * fields parsed from `<c:dLbls>` so a `ChartSeriesInfo` returned by
 * `parseChart` can be fed straight into {@link cloneChart} without
 * transformation.
 */
export interface ChartDataLabelsInfo {
  showValue?: boolean;
  showCategoryName?: boolean;
  showSeriesName?: boolean;
  showPercent?: boolean;
  /**
   * Mirror of {@link ChartDataLabels.showLegendKey}. Surfaces `true`
   * only when the source `<c:dLbls>` block pinned
   * `<c:showLegendKey val="1"/>` (Excel's "Format Data Labels ->
   * Legend Key" checkbox). The OOXML default `false` collapses to
   * `undefined` so absence and the default round-trip identically
   * through {@link cloneChart}.
   */
  showLegendKey?: boolean;
  position?: ChartDataLabelPosition;
  separator?: string;
  /**
   * Mirror of {@link ChartDataLabels.numberFormat}. Surfaces the
   * `<c:numFmt formatCode=".." sourceLinked=".."/>` parsed from the
   * source `<c:dLbls>` block — Excel's "Format Data Labels -> Number"
   * panel pin. Same shape as the axis-side
   * {@link ChartAxisNumberFormat} so the parsed value can be fed
   * straight back into {@link cloneChart} or {@link writeXlsx} without
   * transformation. Absent when the source data-labels block has no
   * `<c:numFmt>` element or when the parsed `formatCode` is missing /
   * empty.
   */
  numberFormat?: ChartAxisNumberFormat;
  /**
   * Mirror of {@link ChartDataLabels.showLeaderLines}. Surfaces the
   * `<c:showLeaderLines val=".."/>` flag parsed from the source
   * `<c:dLbls>` block — Excel's "Format Data Labels -> Show Leader
   * Lines" checkbox.
   *
   * The OOXML default is `true`. Surfaces `false` only when the source
   * pinned `<c:showLeaderLines val="0"/>`; absence and the default
   * `val="1"` collapse to `undefined` so a re-parse of a writer that
   * omits the element matches a re-parse of one that pins the default
   * explicitly.
   *
   * The element only appears on pie / doughnut data-labels per the
   * OOXML schema (`EG_DLbls` is scoped to `CT_PieChart` /
   * `CT_DoughnutChart`); the parser surfaces it on every chart family
   * the source emits it for so a templated chart can round-trip
   * cleanly even when the chart-type element ends up coerced.
   */
  showLeaderLines?: boolean;
  /**
   * Data-label font size in points pulled from `<c:dLbls><c:txPr>
   * <a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr></c:dLbls>`.
   * The OOXML attribute is in 100ths of a point on
   * `CT_TextCharacterProperties`' `sz` slot (ECMA-376 Part 1,
   * §21.1.2.3.7); the reader divides by 100 at parse time so the
   * surfaced value matches what the user sees in Excel's "Format Data
   * Labels -> Font -> Size" UI.
   *
   * Out-of-range values (`< 1` or `> 400`) and malformed tokens
   * collapse to `undefined` so absence and a malformed source value
   * round-trip identically through {@link cloneChart} — only an
   * in-range `sz` attribute surfaces a numeric value. Mirrors the
   * writer-side {@link ChartDataLabels.fontSize} so a parsed value
   * slots straight into {@link cloneChart} without conversion.
   */
  fontSize?: number;
  /**
   * Data-label font color pulled from `<c:dLbls><c:txPr><a:p><a:pPr>
   * <a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
   * </a:defRPr></a:pPr></a:p></c:txPr></c:dLbls>`. The OOXML
   * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
   * sRGB color (`CT_SRgbColor` inside `CT_TextCharacterProperties`'
   * fill choice — ECMA-376 Part 1, §20.1.2.3.32 / §21.1.2.3.7).
   *
   * Returned as the canonical 6-character uppercase hex string when
   * the parser walks the full chain and lands on an `<a:srgbClr
   * val="RRGGBB"/>`. Theme references (`<a:schemeClr>`),
   * `<a:hslClr>`, `<a:sysClr>`, `<a:prstClr>`, and malformed `val`
   * tokens (wrong length, non-hex characters) all collapse to
   * `undefined` since only the literal RGB triple round-trips
   * losslessly through {@link writeChart}. Mirrors the writer-side
   * {@link ChartDataLabels.fontColor} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   */
  fontColor?: string;
  /**
   * Data-label bold flag pulled from `<c:dLbls><c:txPr><a:p><a:pPr>
   * <a:defRPr b=".."/></a:pPr></a:p></c:txPr></c:dLbls>`. The OOXML
   * `b` attribute is the `xsd:boolean` bold flag on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7).
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `b="0"` round-trip identically — only an explicit `b="1"` surfaces
   * `true`. Unknown / malformed `b` tokens drop to `undefined` rather
   * than fabricate a value the writer would never emit. Mirrors the
   * writer-side {@link ChartDataLabels.bold} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   */
  bold?: boolean;
  /**
   * Data-label italic flag pulled from `<c:dLbls><c:txPr><a:p><a:pPr>
   * <a:defRPr i=".."/></a:pPr></a:p></c:txPr></c:dLbls>`. The OOXML
   * `i` attribute is the `xsd:boolean` italic flag on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7).
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `i="0"` round-trip identically — only an explicit `i="1"` surfaces
   * `true`. Unknown / malformed `i` tokens drop to `undefined` rather
   * than fabricate a value the writer would never emit. Mirrors the
   * writer-side {@link ChartDataLabels.italic} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   */
  italic?: boolean;
  /**
   * Data-label underline flag pulled from `<c:dLbls><c:txPr><a:p>
   * <a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr></c:dLbls>`. The
   * OOXML `u` attribute is the `ST_TextUnderlineType` enumeration on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7).
   *
   * Only `u="sng"` (Excel's UI variant — single underline) surfaces
   * `true`; the OOXML default `"none"` (and every other variant the
   * schema allows — `"dbl"`, `"heavy"`, `"dotted"`, `"dotDash"`,
   * `"wavy"`, etc.) collapse to `undefined` so absence and `u="none"`
   * round-trip identically through `cloneChart`. Reporting any
   * non-`"sng"` underline as `true` would silently downgrade the
   * choice to a single line on round-trip; the writer emits only
   * `u="sng"` / `u="none"`, matching the boolean shape the UI
   * exposes. Mirrors the writer-side
   * {@link ChartDataLabels.underline} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   */
  underline?: boolean;
  /**
   * Data-label strikethrough flag pulled from `<c:dLbls><c:txPr>
   * <a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr>
   * </c:dLbls>`. The OOXML `strike` attribute is the
   * `ST_TextStrikeType` enumeration on `CT_TextCharacterProperties`
   * (ECMA-376 Part 1, §21.1.2.3.7).
   *
   * Only `strike="sngStrike"` (Excel's UI variant — single line)
   * surfaces `true`; the OOXML default `"noStrike"` and the non-UI
   * variant `"dblStrike"` (and any malformed token) collapse to
   * `undefined` so absence and `"noStrike"` round-trip identically
   * through `cloneChart`. Reporting `"dblStrike"` as `true` would
   * silently downgrade the choice to a single line on round-trip;
   * the writer emits only `"sngStrike"`, matching the boolean shape
   * the UI exposes. Mirrors the writer-side
   * {@link ChartDataLabels.strikethrough} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   */
  strikethrough?: boolean;
  /**
   * Data-label font family / typeface pulled from `<c:dLbls><c:txPr>
   * <a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
   * </a:pPr></a:p></c:txPr></c:dLbls>`. Reflects Excel's "Format Data
   * Labels -> Font -> Font" picker. The OOXML `<a:latin
   * typeface=".."/>` element carries the typeface name (`CT_TextFont`,
   * ECMA-376 Part 1, §21.1.2.3.7).
   *
   * Reports the trimmed typeface string when the source chart pinned
   * a non-empty typeface; absence and empty / whitespace-only
   * `typeface` attributes both collapse to `undefined` so absence
   * and `<a:latin typeface=""/>` round-trip identically through
   * {@link cloneChart}. Non-string `typeface` tokens (defensive — the
   * XML parser only ever surfaces strings) likewise drop to
   * `undefined`. Mirrors the writer-side
   * {@link ChartDataLabels.fontFamily} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   */
  fontFamily?: string;
  /**
   * Data-labels background fill pulled from `<c:dLbls><c:spPr>
   * <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></c:spPr>
   * </c:dLbls>`. The OOXML `<c:spPr>` block sits on `CT_DLbls` between
   * `<c:numFmt>` and `<c:txPr>` per the schema sequence (ECMA-376
   * Part 1, §21.2.2.50); the `<a:srgbClr val=".."/>` carries the
   * 6-character uppercase hex sRGB color (`CT_SRgbColor` inside
   * `CT_ShapeProperties`' fill choice — ECMA-376 Part 1, §20.1.2.3.32 /
   * §20.1.8.54).
   *
   * Returned as the canonical 6-character uppercase hex string when
   * the parser walks the full chain and lands on an `<a:srgbClr
   * val="RRGGBB"/>`. Theme references (`<a:schemeClr>`),
   * `<a:hslClr>`, `<a:sysClr>`, `<a:prstClr>`, non-solid fills
   * (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>` / `<a:blipFill>`),
   * and malformed `val` tokens (wrong length, non-hex characters) all
   * collapse to `undefined` since only the literal RGB triple
   * round-trips losslessly through {@link writeChart}. Mirrors the
   * writer-side {@link ChartDataLabels.fillColor} so a parsed value
   * slots straight into {@link cloneChart} without conversion.
   */
  fillColor?: string;
  /**
   * Data-labels border (line) color pulled from `<c:dLbls><c:spPr>
   * <a:ln><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></a:ln>
   * </c:spPr></c:dLbls>`. The OOXML `<c:spPr>` block sits on
   * `CT_DLbls` between `<c:numFmt>` and `<c:txPr>` per the schema
   * sequence (ECMA-376 Part 1, §21.2.2.50); the `<a:ln>` child sits
   * inside the `<c:spPr>` block alongside the optional `<a:solidFill>`
   * fill child, and the `<a:srgbClr val=".."/>` carries the
   * 6-character uppercase hex sRGB color (`CT_SRgbColor` inside
   * `<a:ln>`'s solid fill choice — ECMA-376 Part 1, §20.1.2.3.32 /
   * §20.1.2.3.24).
   *
   * Returned as the canonical 6-character uppercase hex string when
   * the parser walks the full chain and lands on an `<a:srgbClr
   * val="RRGGBB"/>`. Theme references (`<a:schemeClr>`),
   * `<a:hslClr>`, `<a:sysClr>`, `<a:prstClr>`, non-solid line fills
   * (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>`), and malformed
   * `val` tokens (wrong length, non-hex characters) all collapse to
   * `undefined` since only the literal RGB triple round-trips
   * losslessly through {@link writeChart}. Mirrors the writer-side
   * {@link ChartDataLabels.borderColor} so a parsed value slots
   * straight into {@link cloneChart} without conversion. Independent
   * of {@link fillColor}: the fill lives on `<c:dLbls><c:spPr>
   * <a:solidFill>`, the stroke lives on `<c:dLbls><c:spPr><a:ln>
   * <a:solidFill>` — the two readers walk disjoint children of the
   * same `<c:spPr>` block so a caller can pin both knobs without
   * conflict.
   */
  borderColor?: string;
  /**
   * Data-labels border (stroke) thickness in points pulled from the
   * `w` attribute on `<c:dLbls><c:spPr><a:ln w="EMU">`. Reflects
   * Excel's "Format Data Labels -> Border -> Width" spinner. The OOXML
   * `w` attribute carries the stroke width in English Metric Units
   * (1 pt = 12 700 EMU) per `CT_LineProperties` (ECMA-376 Part 1,
   * §20.1.2.3.24); the reader divides by 12 700 and snaps to the
   * 0.25 pt grid Excel's UI exposes so a parsed-then-cloned width
   * does not drift across round-trips.
   *
   * Reports the point value clamped to the `0.25..13.5` pt band Excel
   * accepts in the UI when the source labels pinned a finite, positive
   * `w` attribute. Absence (no `<a:ln>` or no `w` attribute), zero,
   * negative, and non-numeric `w` values all collapse to `undefined`
   * so absence and an unrenderable width round-trip identically through
   * {@link cloneChart}. Mirrors the writer-side
   * {@link ChartDataLabels.borderWidth} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   */
  borderWidth?: number;
  /**
   * Data-labels border (stroke) preset dash pattern pulled from the
   * `val` attribute on `<c:dLbls><c:spPr><a:ln><a:prstDash val=".."/>`.
   * Reflects Excel's "Format Data Labels -> Border -> Dash type"
   * picker. Reports the {@link ChartBorderDash} value pinned by the
   * source, or `undefined` when the element is absent / the OOXML
   * default `"solid"` was authored / the value is unrecognized. Mirrors
   * the writer-side {@link ChartDataLabels.borderDash}.
   */
  borderDash?: ChartBorderDash;
}

/**
 * Cell-anchored placement for a chart on its host sheet.
 *
 * Mirrors the `<xdr:from>` / `<xdr:to>` pair on the drawing-layer
 * `xdr:twoCellAnchor` (or the `<xdr:from>` alone for a
 * `xdr:oneCellAnchor`). Coordinates are 0-based row/col indices into
 * the worksheet — identical to the convention used by
 * {@link SheetImage.anchor} and {@link SheetChart.anchor}, so a parsed
 * `ChartAnchor` slots straight back into the writer's shape.
 *
 * `to` is optional because Excel also supports `xdr:oneCellAnchor`
 * (chart pinned to a single cell with intrinsic size).
 * `xdr:absoluteAnchor` (EMU-positioned) does not surface here — those
 * charts are reported with `anchor` undefined.
 */
export interface ChartAnchor {
  /** Top-left cell (`<xdr:from>`). */
  from: { row: number; col: number };
  /** Bottom-right cell (`<xdr:to>`). Omitted for one-cell anchors. */
  to?: { row: number; col: number };
}

/**
 * Major / minor gridline visibility for a chart axis.
 *
 * Excel paints horizontal or vertical reference lines across the plot
 * area, anchored to the major or minor tick marks of an axis. The
 * presence of `<c:majorGridlines>` / `<c:minorGridlines>` inside an
 * `<c:catAx>` or `<c:valAx>` toggles them on; absence of the element
 * means the gridline is off (Excel's default for the value axis is
 * major-on/minor-off, but the OOXML serialization is explicit either
 * way — the writer mirrors what the model says).
 */
export interface ChartAxisGridlines {
  /** Whether the axis declares `<c:majorGridlines>`. */
  major?: boolean;
  /** Whether the axis declares `<c:minorGridlines>`. */
  minor?: boolean;
}

/**
 * Per-axis metadata pulled from the chart's `<c:catAx>` / `<c:valAx>`
 * elements.
 *
 * Surfaces the structural pieces that dashboard cloning needs to
 * preserve through a `parseChart` → {@link cloneChart} → `writeXlsx`
 * round-trip — currently the axis title and the gridline visibility.
 */
/**
 * Value-axis scaling pulled from `<c:scaling>` — bounds plus tick
 * spacing. Excel reports a numeric scale for any value-axis chart;
 * absent on category axes (`<c:catAx>` tolerates `<c:scaling>` but
 * populates only `<c:orientation>` there).
 *
 * All four numeric fields are optional — a chart may declare any
 * subset and Excel auto-computes the rest. Hucre surfaces only the
 * explicitly declared values, so a round-trip cannot accidentally pin
 * an axis to numbers Excel would otherwise have inferred.
 */
export interface ChartAxisScale {
  /** `<c:min>` — value where the axis starts. */
  min?: number;
  /** `<c:max>` — value where the axis ends. */
  max?: number;
  /** `<c:majorUnit>` — spacing between major tick marks. Must be > 0. */
  majorUnit?: number;
  /** `<c:minorUnit>` — spacing between minor tick marks. Must be > 0. */
  minorUnit?: number;
  /**
   * `<c:logBase>` — log base for a logarithmic scale. Excel restricts
   * this to 2–1000; the parser does not enforce that range, but the
   * writer rejects values outside it.
   */
  logBase?: number;
}

/**
 * Axis number-format spec pulled from `<c:numFmt>`. Mirrors what Excel
 * emits for tick labels — an OOXML number-format code (e.g.
 * `"#,##0"`, `"0.00%"`, `"$#,##0.00"`) and a `sourceLinked` flag that
 * tells Excel whether to inherit the cell number format from the
 * underlying data range.
 */
export interface ChartAxisNumberFormat {
  /** OOXML format code (e.g. `"#,##0"`, `"0.00%"`). */
  formatCode: string;
  /**
   * When `true`, Excel ignores `formatCode` and pulls the format
   * straight from the source data range. Defaults to `false` when
   * omitted — the pinned `formatCode` wins.
   */
  sourceLinked?: boolean;
}

/**
 * Axis tick-mark style — where Excel paints the short tick lines that
 * mark major or minor unit boundaries on a category or value axis.
 *
 * Maps to the OOXML `ST_TickMark` enumeration which sits inside
 * `<c:catAx>` / `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` as
 * `<c:majorTickMark val=".."/>` and `<c:minorTickMark val=".."/>`:
 *
 * - `"none"`  — no tick marks rendered at all.
 * - `"in"`    — tick marks point inward (toward the plot area).
 * - `"out"`   — tick marks point outward (away from the plot area).
 *               OOXML default for `<c:majorTickMark>`.
 * - `"cross"` — tick marks straddle the axis line.
 *
 * Excel's UI exposes the same four presets under "Format Axis →
 * Tick Marks → Major type / Minor type". The OOXML default for
 * `<c:minorTickMark>` is `"none"` (Excel's UI also defaults to "None"
 * for the minor type on a freshly-drawn axis).
 */
export type ChartAxisTickMark = "none" | "in" | "out" | "cross";

/**
 * Axis tick-label position — where Excel paints the numeric / category
 * labels relative to the axis line.
 *
 * Maps to the OOXML `ST_TickLblPos` enumeration which sits inside
 * `<c:catAx>` / `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` as
 * `<c:tickLblPos val=".."/>`:
 *
 * - `"nextTo"` — labels sit alongside the axis line at the closest
 *                edge of the plot area. OOXML default.
 * - `"low"`    — labels pinned to the low end of the perpendicular
 *                axis (left for value axes, bottom for category axes).
 *                Useful when the axis crosses elsewhere but labels
 *                should stay anchored to the chart edge.
 * - `"high"`   — mirror of `"low"`; labels pinned to the high end.
 * - `"none"`   — no labels rendered. Excel's UI exposes this as
 *                "Format Axis → Labels → Label Position → None".
 */
export type ChartAxisTickLabelPosition = "nextTo" | "low" | "high" | "none";

/**
 * Horizontal alignment for category-axis tick labels — where Excel
 * anchors each label inside its allocated cell along the axis.
 *
 * Maps to the OOXML `ST_LblAlgn` enumeration which sits inside
 * `<c:catAx>` / `<c:dateAx>` as `<c:lblAlgn val=".."/>`. The element
 * does not exist on `<c:valAx>` / `<c:serAx>`:
 *
 * - `"ctr"` — labels centered along the axis. OOXML default and what
 *             Excel paints on a freshly-drawn category axis.
 * - `"l"`   — labels pinned to the left edge of their slot. Useful for
 *             multi-line wrapped labels on a column chart that should
 *             align flush with the leftmost gridline.
 * - `"r"`   — labels pinned to the right edge of their slot.
 *
 * Excel's UI exposes the three presets under "Format Axis ->
 * Alignment -> Text alignment" on a category axis. Pie / doughnut and
 * scatter charts have no category axis, so the field is dropped on
 * those families.
 */
export type ChartAxisLabelAlign = "ctr" | "l" | "r";

/**
 * Axis crossing position — where the perpendicular axis crosses this
 * axis along its own range. Maps to the OOXML `ST_Crosses` enumeration
 * which sits inside `<c:catAx>` / `<c:valAx>` / `<c:dateAx>` /
 * `<c:serAx>` as `<c:crosses val=".."/>`:
 *
 * - `"autoZero"` — the perpendicular axis crosses at zero on a value
 *                  axis (or at the first category on a category axis).
 *                  OOXML default and Excel's reference serialization on
 *                  every freshly-drawn axis.
 * - `"min"`      — the perpendicular axis crosses at the low end of
 *                  this axis (Excel's "Format Axis -> Vertical axis
 *                  crosses -> Automatic / At minimum value" toggle).
 * - `"max"`      — the perpendicular axis crosses at the high end.
 *
 * `<c:crosses>` and `<c:crossesAt>` are mutually exclusive in the OOXML
 * schema (CT_Crosses sits in an XSD choice with CT_Double). The writer
 * favours `crossesAt` whenever the caller pins it; `crosses` is the
 * fallback when only the semantic toggle is set.
 */
export type ChartAxisCrosses = "autoZero" | "min" | "max";

/**
 * Whether the perpendicular axis crosses BETWEEN data points or AT the
 * midpoint of each category on a value axis. Maps to the OOXML
 * `ST_CrossBetween` enumeration which sits inside `<c:valAx>` as
 * `<c:crossBetween val=".."/>`. The element is value-axis-only — the
 * OOXML schema places `<c:crossBetween>` exclusively on `CT_ValAx`, so
 * `<c:catAx>` / `<c:dateAx>` / `<c:serAx>` reject it:
 *
 * - `"between"` — the perpendicular axis crosses between data points,
 *                 leaving a half-category gap on each end of the plot
 *                 area. Excel's reference serialization on bar / column
 *                 charts (so bars sit inside their category slot rather
 *                 than straddling the value-axis line) and the writer's
 *                 default on bar / column / line / area today.
 * - `"midCat"`  — the perpendicular axis crosses at the midpoint of
 *                 each category, so data points (line markers / area
 *                 fill anchors / scatter points) sit ON the
 *                 perpendicular-axis ticks rather than between them.
 *                 Excel's reference serialization on scatter charts —
 *                 useful when porting line / area templates whose first
 *                 / last data point should land flush with the value
 *                 axis instead of inside the plot area.
 *
 * Excel's UI does not expose the toggle as a checkbox — Excel computes
 * it from the chart family on insertion — but Excel preserves the
 * element on round-trip, and a template that pins a non-default value
 * should round-trip through `parseChart -> cloneChart -> writeXlsx`
 * without flattening.
 */
export type ChartAxisCrossBetween = "between" | "midCat";

/**
 * Built-in display-unit preset on a value axis — Excel's "Format Axis ->
 * Display units" dropdown. Every numeric tick label is divided by the
 * preset's scale before being rendered, so a chart whose source range
 * stores raw amounts (e.g. `1_500_000`) can display compact tick labels
 * (`1.5` with a "Millions" annotation) without modifying the underlying
 * cells.
 *
 * Maps to the OOXML `ST_BuiltInUnit` enumeration which sits inside
 * `<c:dispUnits>` on `<c:valAx>` as `<c:builtInUnit val=".."/>`. The
 * tokens mirror Excel's UI labels:
 *
 * - `"hundreds"`         — divide by 1e2.
 * - `"thousands"`        — divide by 1e3.
 * - `"tenThousands"`     — divide by 1e4.
 * - `"hundredThousands"` — divide by 1e5.
 * - `"millions"`         — divide by 1e6.
 * - `"tenMillions"`      — divide by 1e7.
 * - `"hundredMillions"`  — divide by 1e8.
 * - `"billions"`         — divide by 1e9.
 * - `"trillions"`        — divide by 1e12.
 *
 * The OOXML schema also allows a custom numeric divisor via
 * `<c:custUnit val=".."/>`; that variant is not surfaced here — pass a
 * built-in preset instead. Pie / doughnut charts have no value axes, so
 * the field is silently dropped on those families. Category axes
 * (`<c:catAx>`) reject `<c:dispUnits>` entirely, so `dispUnits` only
 * surfaces on the value-axis side of bar / column / line / area
 * charts (the Y axis) and on both axes of scatter charts (both are
 * value axes).
 */
export type ChartAxisDispUnit =
  | "hundreds"
  | "thousands"
  | "tenThousands"
  | "hundredThousands"
  | "millions"
  | "tenMillions"
  | "hundredMillions"
  | "billions"
  | "trillions";

/**
 * Display-unit configuration for a value axis. Maps to the
 * `<c:dispUnits>` element on `<c:valAx>` per ECMA-376 Part 1, §21.2.2.32
 * (CT_ValAx → CT_DispUnits). The element rescales the numeric tick
 * labels by the chosen preset (e.g. `"millions"` divides every label by
 * 1e6) and optionally prints the unit annotation on the chart.
 *
 * The OOXML schema places `<c:builtInUnit>` and `<c:custUnit>` in an
 * `xsd:choice` — exactly one of the two may appear inside `<c:dispUnits>`.
 * Hucre exposes both: pin {@link unit} to pick one of the named OOXML
 * presets (the common path — Excel's "Display units → Hundreds /
 * Thousands / ..." dropdown), or pin {@link custUnit} to declare an
 * arbitrary numeric divisor (Excel's "Display units → Other" path).
 * When both fields are pinned, `custUnit` wins on emit because the
 * OOXML schema forbids emitting both children — the more specific
 * numeric divisor takes precedence so a caller can append a custom unit
 * to a cloned source without manually pruning the inherited preset.
 *
 * `<c:dispUnitsLbl>` is also intentionally minimal: when `showLabel` is
 * `true` the writer emits a bare `<c:dispUnitsLbl/>` so Excel paints its
 * default "Millions" / "Thousands" / ... annotation alongside the axis;
 * the rich-text label customization (`<a:p>` / `<a:r>` inside
 * `<c:dispUnitsLbl>`) is not surfaced. Callers needing a custom label
 * string can layer it on later.
 */
export interface ChartAxisDispUnits {
  /**
   * OOXML `ST_BuiltInUnit` token — the preset divisor. Maps to
   * `<c:dispUnits><c:builtInUnit val=".."/></c:dispUnits>`. Mutually
   * exclusive with {@link custUnit} per the OOXML schema; when both are
   * pinned, `custUnit` wins on emit. Required when `custUnit` is
   * absent — a `ChartAxisDispUnits` object with neither field pinned
   * collapses to no element on emit.
   */
  unit?: ChartAxisDispUnit;
  /**
   * Custom numeric divisor — Excel's "Display units → Other" path.
   * Maps to `<c:dispUnits><c:custUnit val=".."/></c:dispUnits>` (CT_Double
   * per the OOXML schema). The divisor rescales every tick label by the
   * given factor (e.g. `1000` divides labels by 1 000, the same as the
   * `"thousands"` preset; `86400` converts seconds to days).
   *
   * Mutually exclusive with {@link unit} per the OOXML `xsd:choice` —
   * when both are pinned, `custUnit` wins on emit. Must be a finite
   * positive number; `0`, negative, non-finite (`NaN`, `Infinity`), and
   * non-number inputs drop silently rather than emit a token Excel
   * would refuse.
   */
  custUnit?: number;
  /**
   * Whether to print Excel's automatic display-unit annotation
   * alongside the axis (e.g. "Millions" for `unit: "millions"`). Maps
   * to the presence of `<c:dispUnitsLbl/>` inside `<c:dispUnits>`.
   * Default: `false` (no label rendered, the divisor still applies).
   */
  showLabel?: boolean;
}

export interface ChartAxisInfo {
  /** Plain-text title from the axis's `<c:title>`. Omitted when absent. */
  title?: string;
  /**
   * Axis-title rotation in degrees pulled from
   * `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title>`
   * on the axis element. Mirrors Excel's "Format Axis Title -> Size &
   * Properties -> Alignment -> Custom angle" pin.
   *
   * The OOXML `rot` attribute is in 60000ths of a degree; the reader
   * converts to whole degrees and surfaces the literal value (range
   * `-90..90`). The OOXML default `0` (and absence of the element)
   * collapses to `undefined` so absence and the default round-trip
   * identically through {@link cloneChart}. Out-of-range values clamp
   * to the nearest endpoint so a parsed value slots straight back into
   * the writer-side {@link SheetChart.axes.x.axisTitleRotation} field
   * without further transformation. Returns `undefined` when the axis
   * omits `<c:title>` entirely or when the title is a `<c:strRef>`
   * (formula reference) with no `<c:rich>` body.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. The reader surfaces the value on every axis
   * flavour so a parsed chart preserves the rotation regardless of
   * whether the source axis was a category or value axis.
   */
  axisTitleRotation?: number;
  /**
   * Axis-title font size in points (range `1..400`), pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
   * </a:p></c:rich></c:tx></c:title>` on the axis element. Reflects
   * Excel's "Format Axis Title -> Font -> Size" knob.
   *
   * The OOXML attribute is in 100ths of a point; the reader converts
   * to points and rounds to the nearest 0.5pt (Excel's UI exposes the
   * same 0.5pt granularity, e.g. `sz="1000"` surfaces as `10`,
   * `sz="1050"` as `10.5`). Out-of-range values (outside the `1..400`
   * band the OOXML `ST_TextFontSize` schema exposes) drop to
   * `undefined` rather than fabricate a value the writer would never
   * emit. Absence of the attribute (or of `<a:defRPr>` / `<a:pPr>` /
   * `<a:p>` / `<c:rich>` / `<c:tx>` / `<c:title>`) likewise collapses
   * to `undefined`.
   *
   * Reported as `undefined` whenever the axis has no `<c:title>`
   * element at all, or when the title is a `<c:strRef>` (formula
   * reference) with no `<c:rich>` body — there is no `<a:p>` slot to
   * surface the size from in either case. Mirrors the chart-level
   * {@link Chart.titleFontSize} so a parsed value slots straight back
   * into the writer-side {@link SheetChart.axes.x.axisTitleFontSize}
   * without transformation.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. The reader surfaces the value on every axis
   * flavour so a parsed chart preserves the size regardless of
   * whether the source axis was a category or value axis.
   */
  axisTitleFontSize?: number;
  /**
   * Axis-title bold flag pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
   * </a:p></c:rich></c:tx></c:title>` on the axis element. Reflects
   * Excel's "Format Axis Title -> Font -> Bold" toggle.
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `<a:defRPr b="0"/>` round-trip identically through
   * {@link cloneChart} — only an explicit `<a:defRPr b="1"/>` surfaces
   * `true`. The reader accepts the OOXML truthy / falsy spellings
   * (`"1"` / `"true"` / `"0"` / `"false"`); unknown values and missing
   * `b` attributes drop to `undefined`.
   *
   * Reported as `undefined` whenever the axis has no `<c:title>`
   * element at all, or when the title is a `<c:strRef>` (formula
   * reference) with no `<c:rich>` body — there is no `<a:p>` slot to
   * surface the flag from in either case. Mirrors the chart-level
   * {@link Chart.titleBold} so a parsed value slots straight back
   * into the writer-side {@link SheetChart.axes.x.axisTitleBold}
   * without transformation.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. The reader surfaces the value on every axis
   * flavour so a parsed chart preserves the bold state regardless of
   * whether the source axis was a category or value axis.
   */
  axisTitleBold?: boolean;
  /**
   * Axis-title italic flag pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
   * </c:rich></c:tx></c:title>` on the axis element. Reflects Excel's
   * "Format Axis Title -> Font -> Italic" toggle.
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `<a:defRPr i="0"/>` round-trip identically through
   * {@link cloneChart} — only an explicit `<a:defRPr i="1"/>` surfaces
   * `true`. The reader accepts the OOXML truthy / falsy spellings
   * (`"1"` / `"true"` / `"0"` / `"false"`); unknown values and missing
   * `i` attributes drop to `undefined`.
   *
   * Reported as `undefined` whenever the axis has no `<c:title>`
   * element at all, or when the title is a `<c:strRef>` (formula
   * reference) with no `<c:rich>` body — there is no `<a:p>` to host
   * the flag in either case. Mirrors the chart-level
   * {@link Chart.titleItalic} so a parsed value slots straight back
   * into the writer-side {@link SheetChart.axes.x.axisTitleItalic}
   * without transformation.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. The reader surfaces the value on every axis
   * flavour so a parsed chart preserves the flag regardless of
   * whether the source axis was a category or value axis.
   */
  axisTitleItalic?: boolean;
  /**
   * Axis-title font color pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
   * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
   * </c:rich></c:tx></c:title>` on the axis element. Reflects Excel's
   * "Format Axis Title -> Font -> Font Color" picker.
   *
   * Surfaced as the 6-character uppercase hex string the writer round-
   * trips (`"FF0000"` / `"1070CA"`) — the leading `#` is stripped on
   * read so the value threads straight into the writer-side
   * {@link SheetChart.axes.x.axisTitleColor} without transformation.
   * Color picks other than the literal sRGB form (`<a:schemeClr>` theme
   * references, `<a:hslClr>`, `<a:sysClr>`, `<a:prstClr>`) collapse to
   * `undefined` — the reader records only the resolvable RGB triple to
   * keep the round-trip lossless against {@link cloneChart} ->
   * {@link writeXlsx}. Malformed `val` tokens (wrong length, non-hex
   * characters) likewise drop to `undefined` rather than fabricate a
   * value the writer would round-trip into a malformed `<a:srgbClr>`.
   *
   * Reported as `undefined` whenever the axis omits `<c:title>`
   * entirely, when the title is a `<c:strRef>` (formula reference)
   * with no `<c:rich>` body, or when the `<a:defRPr>` slot has no
   * `<a:solidFill>` child (the title inherits the theme's text color
   * in that case). Mirrors the chart-level {@link Chart.titleColor} so
   * a parsed value slots straight back into the writer-side
   * {@link SheetChart.axes.x.axisTitleColor} without transformation.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. The reader surfaces the value on every axis
   * flavour so a parsed chart preserves the color regardless of
   * whether the source axis was a category or value axis.
   */
  axisTitleColor?: string;
  /**
   * Axis-title strikethrough flag pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
   * </a:p></c:rich></c:tx></c:title>` on the axis element. Reflects
   * Excel's "Format Axis Title -> Font -> Strikethrough" toggle.
   *
   * The OOXML attribute is the `ST_TextStrikeType` enum on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
   * three values: `"noStrike"`, `"sngStrike"`, `"dblStrike"`. Only the
   * UI-default `"sngStrike"` (Excel's "Strikethrough" checkbox —
   * single line) surfaces as `true`; `"noStrike"` (the OOXML
   * application default) and absence both collapse to `undefined`,
   * and the non-UI `"dblStrike"` variant likewise collapses to
   * `undefined` rather than surface a value the writer would silently
   * downgrade on round-trip — hucre's writer emits only `"sngStrike"`,
   * so reporting `"dblStrike"` as `true` would round-trip into a lossy
   * single-line replacement.
   *
   * Unknown / malformed `strike` tokens drop to `undefined` rather
   * than fabricate a value the writer would never emit.
   *
   * Reported as `undefined` whenever the axis omits `<c:title>`
   * entirely, or when the title is a `<c:strRef>` (formula reference)
   * with no `<c:rich>` body — there is no `<a:p>` to host the flag in
   * either case. Mirrors the chart-level {@link Chart.titleStrike} so
   * a parsed value slots straight back into the writer-side
   * {@link SheetChart.axes.x.axisTitleStrike} without transformation.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. The reader surfaces the value on every axis
   * flavour so a parsed chart preserves the flag regardless of
   * whether the source axis was a category or value axis.
   */
  axisTitleStrike?: boolean;
  /**
   * Axis-title underline flag pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
   * </a:p></c:rich></c:tx></c:title>` on the axis element. Reflects
   * Excel's "Format Axis Title -> Font -> Underline" picker.
   *
   * The OOXML attribute is the `ST_TextUnderlineType` enum on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
   * eighteen values; Excel's UI exposes only `"sng"` (single line —
   * the default underline checkbox) and `"dbl"` (double line). Only
   * the UI-default `"sng"` (Excel's "Underline" checkbox — single
   * line) surfaces as `true`; `"none"` (the OOXML application
   * default) and absence both collapse to `undefined`, and every
   * other token (`"dbl"` and the sixteen exotic variants) likewise
   * collapses to `undefined` rather than surface a value the writer
   * would silently downgrade on round-trip — hucre's writer emits
   * only `"sng"`, so reporting any non-single underline as `true`
   * would round-trip into a lossy single-line replacement.
   *
   * Unknown / malformed `u` tokens drop to `undefined` rather than
   * fabricate a value the writer would never emit.
   *
   * Reported as `undefined` whenever the axis omits `<c:title>`
   * entirely, or when the title is a `<c:strRef>` (formula reference)
   * with no `<c:rich>` body — there is no `<a:p>` to host the flag in
   * either case. Mirrors the chart-level {@link Chart.titleUnderline}
   * so a parsed value slots straight back into the writer-side
   * {@link SheetChart.axes.x.axisTitleUnderline} without
   * transformation.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. The reader surfaces the value on every axis
   * flavour so a parsed chart preserves the flag regardless of
   * whether the source axis was a category or value axis.
   */
  axisTitleUnderline?: boolean;
  /**
   * Axis title font family / typeface pulled from `<c:catAx>` /
   * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` ->
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
   * typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx>
   * </c:title>`. Reflects Excel's "Format Axis Title -> Font ->
   * Font" picker. The OOXML `<a:latin typeface=".."/>` element
   * carries the typeface name (`CT_TextFont`, ECMA-376 Part 1,
   * §21.1.2.3.7).
   *
   * Reports the trimmed typeface string when the source axis pinned
   * a non-empty typeface; absence and empty / whitespace-only
   * `typeface` attributes both collapse to `undefined` so absence
   * and `<a:latin typeface=""/>` round-trip identically through
   * {@link cloneChart}. Non-string `typeface` tokens (defensive —
   * the XML parser only ever surfaces strings) likewise drop to
   * `undefined`.
   *
   * Reported as `undefined` whenever the source axis has no
   * `<c:title>` element at all, or when the title is a `<c:strRef>`
   * (formula reference) with no `<c:rich>` body — there is no
   * `<a:p>` to host the typeface in either case. Mirrors the
   * chart-level {@link Chart.titleFontFamily} so a parsed value
   * slots straight back into the writer-side
   * {@link SheetChart.axes.x.axisTitleFontFamily} without
   * transformation.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. The reader surfaces the value on every
   * axis flavour so a parsed chart preserves the typeface
   * regardless of whether the source axis was a category or value
   * axis.
   */
  axisTitleFontFamily?: string;
  /**
   * Axis-title overlay flag pulled from
   * `<c:catAx><c:title><c:overlay val=".."/></c:title></c:catAx>` (or
   * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>`). Reflects the OOXML
   * `<c:overlay>` child of `CT_Title` — whether the axis title is
   * drawn on top of (and may overlap) the plot area.
   *
   * The OOXML default `false` (the title reserves its own slot
   * adjacent to the axis, no overlap with the plot area) collapses
   * to `undefined` so absence and `<c:overlay val="0"/>` round-trip
   * identically through {@link cloneChart} — only an explicit
   * `<c:overlay val="1"/>` surfaces `true`. The reader accepts the
   * OOXML truthy / falsy spellings (`"1"` / `"true"` / `"0"` /
   * `"false"`); unknown values and missing `val` attributes drop to
   * `undefined` rather than fabricate a flag Excel would not emit.
   *
   * Returned as `undefined` whenever the axis omits the `<c:title>`
   * element entirely — there is no overlay slot to surface in that
   * case. The element is a sibling of `<c:tx>` inside `<c:title>`
   * per the CT_Title schema, so the lookup is scoped to direct title
   * children.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. Mirrors the chart-level
   * {@link Chart.titleOverlay} so a parsed value slots straight back
   * into the writer-side {@link SheetChart.axes.x.axisTitleOverlay}
   * without transformation.
   */
  axisTitleOverlay?: boolean;
  /**
   * Axis-title manual placement pulled from
   * `<c:catAx><c:title><c:layout><c:manualLayout>...</c:manualLayout>
   * </c:layout></c:title></c:catAx>` (or `<c:valAx>` / `<c:dateAx>` /
   * `<c:serAx>`). Reflects Excel's "Format Axis Title -> Title Options
   * -> Position -> Custom" placement — the `(x, y)` anchor and
   * `(w, h)` size of the axis-title block as fractions of the chart
   * frame in the `0..1` band.
   *
   * Each of {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
   * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} surfaces
   * only when the matching `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` child
   * is present and parses to a finite number in the `0..1` band; out-of-
   * range / non-finite / non-numeric tokens drop on the matching axis
   * so absence and a malformed token round-trip identically through
   * {@link cloneChart}.
   *
   * Both `<c:xMode val="edge"/>` (absolute fraction of the chart frame)
   * and `<c:xMode val="factor"/>` (delta from auto-layout) are accepted
   * — the reader surfaces the same {@link ChartManualLayout} shape
   * regardless, since the writer always normalizes to `"edge"` on emit
   * (Excel itself emits the absolute form when the user drags an
   * element to a custom position).
   *
   * Returned as `undefined` whenever the axis omits the `<c:title>` /
   * `<c:layout>` / `<c:manualLayout>` chain at any link, or when every
   * coordinate dropped on normalization — the field is omitted entirely
   * on a clean parse so absence and an empty layout round-trip
   * identically through the writer. Mirrors the chart-level
   * {@link Chart.titleLayout} (when set on the same chart) and
   * {@link Chart.legendLayout} / {@link Chart.plotAreaLayout} so a
   * parsed value slots straight back into the writer-side
   * {@link SheetChart.axes.x.axisTitleLayout} without transformation.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. The reader surfaces the value on every axis
   * flavour so a parsed chart preserves the placement regardless of
   * whether the source axis was a category or value axis.
   */
  axisTitleLayout?: ChartManualLayout;
  /**
   * Axis-title background fill (solid sRGB) pulled from
   * `<c:catAx><c:title><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
   * </a:solidFill></c:spPr></c:title></c:catAx>` (or `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>`). Reflects Excel's "Format Axis Title
   * -> Fill -> Solid fill -> Color" picker.
   *
   * Surfaced as the 6-character uppercase hex string the writer round-
   * trips (`"FFFF00"` / `"1070CA"`) — the leading `#` is stripped on
   * read so the value threads straight into the writer-side
   * {@link SheetChart.axes.x.axisTitleFillColor} without
   * transformation. Color picks other than the literal sRGB form
   * (`<a:schemeClr>` theme references, `<a:hslClr>`, `<a:sysClr>`,
   * `<a:prstClr>`) collapse to `undefined` — the reader records only
   * the resolvable RGB triple to keep the round-trip lossless against
   * {@link cloneChart} -> {@link writeXlsx}. Non-solid fills
   * (`<a:noFill>`, `<a:gradFill>`, `<a:pattFill>`, `<a:blipFill>`)
   * likewise drop to `undefined`. Malformed `val` tokens (wrong
   * length, non-hex characters) drop to `undefined` rather than
   * fabricate a value the writer would round-trip into a malformed
   * `<a:srgbClr>`.
   *
   * Independent of {@link ChartAxisInfo.axisTitleColor}: the fill
   * lives on `<c:title><c:spPr>`, the font color lives on
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>` —
   * the two readers walk disjoint paths so a caller can pin both
   * knobs without conflict.
   *
   * Reported as `undefined` whenever the axis omits `<c:title>`
   * entirely or when the `<c:spPr><a:solidFill><a:srgbClr>` chain is
   * malformed at any link. Mirrors the chart-level
   * {@link Chart.titleFillColor} so a parsed value slots straight
   * back into the writer-side
   * {@link SheetChart.axes.x.axisTitleFillColor} without
   * transformation.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. The reader surfaces the value on every
   * axis flavour so a parsed chart preserves the fill regardless of
   * whether the source axis was a category or value axis.
   */
  axisTitleFillColor?: string;
  /**
   * Axis-title border (line stroke) solid sRGB color pulled from
   * `<c:catAx><c:title><c:spPr><a:ln><a:solidFill><a:srgbClr
   * val="RRGGBB"/></a:solidFill></a:ln></c:spPr></c:title></c:catAx>`
   * (or `<c:valAx>` / `<c:dateAx>` / `<c:serAx>`). Reflects Excel's
   * "Format Axis Title -> Border -> Solid line -> Color" picker.
   *
   * Surfaced as the 6-character uppercase hex string the writer
   * round-trips (`"FFFF00"` / `"1070CA"`) — the leading `#` is
   * stripped on read so the value threads straight into the
   * writer-side {@link SheetChart.axes.x.axisTitleBorderColor}
   * without transformation. Color picks other than the literal sRGB
   * form (`<a:schemeClr>` theme references, `<a:hslClr>`,
   * `<a:sysClr>`, `<a:prstClr>`) collapse to `undefined` — the
   * reader records only the resolvable RGB triple to keep the round-
   * trip lossless against {@link cloneChart} -> {@link writeXlsx}.
   * Non-solid line fills (`<a:noFill>`, `<a:gradFill>`, `<a:pattFill>`)
   * likewise drop to `undefined`. Malformed `val` tokens (wrong
   * length, non-hex characters) drop to `undefined` rather than
   * fabricate a value the writer would round-trip into a malformed
   * `<a:srgbClr>`.
   *
   * Independent of {@link ChartAxisInfo.axisTitleColor} (font color
   * on `<a:defRPr><a:solidFill>`) and
   * {@link ChartAxisInfo.axisTitleFillColor} (background fill on
   * `<c:spPr><a:solidFill>`): the stroke lives on `<c:spPr><a:ln>
   * <a:solidFill>`, the fill lives on `<c:spPr><a:solidFill>`, and
   * the font color lives on the inner `<a:defRPr><a:solidFill>` —
   * the three readers walk disjoint paths so a caller can pin all
   * three knobs without conflict.
   *
   * Reported as `undefined` whenever the axis omits `<c:title>`
   * entirely or when the `<c:spPr><a:ln><a:solidFill><a:srgbClr>`
   * chain is malformed at any link. Mirrors the chart-level
   * {@link Chart.titleBorderColor} so a parsed value slots straight
   * back into the writer-side
   * {@link SheetChart.axes.x.axisTitleBorderColor} without
   * transformation.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry the same `<c:title>` shape
   * per the OOXML schema. The reader surfaces the value on every
   * axis flavour so a parsed chart preserves the stroke regardless
   * of whether the source axis was a category or value axis.
   */
  axisTitleBorderColor?: string;
  /**
   * Axis-title border (stroke) thickness in points pulled from the
   * `w` attribute on `<c:catAx><c:title><c:spPr><a:ln w="EMU">` (or
   * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>`). Reflects Excel's
   * "Format Axis Title -> Border -> Width" spinner. The OOXML `w`
   * attribute stores the stroke width in English Metric Units
   * (1 pt = 12 700 EMU) per `CT_LineProperties` (ECMA-376 Part 1,
   * §20.1.2.3.24); the reader divides by 12 700 and snaps the result
   * to the 0.25 pt grid Excel's UI exposes so a parsed-then-cloned
   * width does not drift across round-trips.
   *
   * Reports the point value clamped to the `0.25..13.5` pt band
   * Excel accepts in the UI when the source axis pinned a finite,
   * positive `w` attribute. Absence (no `<a:ln>` or no `w`
   * attribute), zero, negative, and non-numeric `w` values all
   * collapse to `undefined` so absence and an unrenderable width
   * round-trip identically through {@link cloneChart}. Mirrors the
   * writer-side {@link SheetChart.axes.x.axisTitleBorderWidth} so a
   * parsed value slots straight into {@link cloneChart} without
   * conversion.
   *
   * Reported as `undefined` whenever the axis omits `<c:title>`
   * entirely. Composes independently with {@link axisTitleBorderColor} —
   * both fields surface from the same `<a:ln>` element but on a
   * different slot (color child vs the width attribute). Mirrors the
   * chart-level {@link Chart.titleBorderWidth} on a different host
   * element.
   */
  axisTitleBorderWidth?: number;
  /**
   * Axis-title border (stroke) preset dash pattern pulled from the
   * `val` attribute on `<c:catAx><c:title><c:spPr><a:ln><a:prstDash
   * val=".."/>` (or `<c:valAx>` / `<c:dateAx>` / `<c:serAx>`). Reflects
   * Excel's "Format Axis Title -> Border -> Dash type" picker. Reports
   * the {@link ChartBorderDash} value pinned by the source, or
   * `undefined` when the element is absent / the OOXML default
   * `"solid"` was authored / the value is unrecognized.
   *
   * Mirrors the writer-side {@link SheetChart.axes.x.axisTitleBorderDash}
   * so a parsed value slots straight into {@link cloneChart} without
   * conversion.
   */
  axisTitleBorderDash?: ChartBorderDash;
  /**
   * Major / minor gridline visibility. Omitted when neither
   * `<c:majorGridlines>` nor `<c:minorGridlines>` is declared on the
   * axis (i.e. Excel's "no gridlines" state for both).
   */
  gridlines?: ChartAxisGridlines;
  /**
   * Numeric scaling (`<c:min>` / `<c:max>` / `<c:majorUnit>` /
   * `<c:minorUnit>` / `<c:logBase>`). Omitted when the axis declared
   * none of those children — Excel auto-computes the bounds in that
   * case and the reader leaves the inference up to the consumer.
   */
  scale?: ChartAxisScale;
  /**
   * Tick-label number format (`<c:numFmt>`). Omitted when the axis
   * does not declare one. Mirrors `formatCode` / `sourceLinked` on
   * the writer side.
   */
  numberFormat?: ChartAxisNumberFormat;
  /**
   * Major tick-mark style pulled from `<c:majorTickMark>`. Omitted
   * when absent or when the axis declared the OOXML default `"out"` —
   * absence and the default round-trip identically through
   * {@link cloneChart}, so collapsing the default keeps the parsed
   * shape minimal. See {@link ChartAxisTickMark}.
   */
  majorTickMark?: ChartAxisTickMark;
  /**
   * Minor tick-mark style pulled from `<c:minorTickMark>`. Omitted
   * when absent or when the axis declared the OOXML default `"none"`.
   * See {@link ChartAxisTickMark}.
   */
  minorTickMark?: ChartAxisTickMark;
  /**
   * Tick-label position pulled from `<c:tickLblPos>`. Omitted when
   * absent or when the axis declared the OOXML default `"nextTo"` —
   * absence and the default round-trip identically through
   * {@link cloneChart}, so collapsing the default keeps the parsed
   * shape minimal. See {@link ChartAxisTickLabelPosition}.
   */
  tickLblPos?: ChartAxisTickLabelPosition;
  /**
   * Tick-label rotation in degrees pulled from
   * `<c:txPr><a:bodyPr rot="N"/></c:txPr>` on the axis element.
   * Mirrors Excel's "Format Axis -> Alignment -> Custom angle" pin.
   *
   * The OOXML `rot` attribute is in 60000ths of a degree; the reader
   * converts to whole degrees and surfaces the literal value (range
   * `-90..90`). The OOXML default `0` (and absence of the element)
   * collapses to `undefined` so absence and the default round-trip
   * identically through {@link cloneChart}. Out-of-range values are
   * clamped to the nearest endpoint so a parsed value slots straight
   * back into the writer-side {@link SheetChart.axes.x.labelRotation}
   * field without further transformation.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
   * the OOXML schema. The reader surfaces the value on every axis
   * flavour so a parsed chart preserves the rotation regardless of
   * whether the source axis was a category or value axis.
   */
  labelRotation?: number;
  /**
   * Tick-label font size in points pulled from
   * `<c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr>`
   * on the axis element. Reflects Excel's "Format Axis -> Font ->
   * Size" knob applied to tick labels.
   *
   * The OOXML attribute is in 100ths of a point; the reader converts
   * to points and rounds to the nearest 0.5pt (Excel's UI exposes the
   * same 0.5pt granularity, e.g. `sz="1000"` surfaces as `10`,
   * `sz="1050"` as `10.5`). Out-of-range values (outside the `1..400`
   * band the OOXML `ST_TextFontSize` schema exposes) drop to
   * `undefined` rather than fabricate a value the writer would never
   * emit. Absence of the attribute / element / chain likewise
   * collapses to `undefined`.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
   * the OOXML schema. Mirrors the writer-side
   * {@link SheetChart.axes.x.labelFontSize} so a parsed value slots
   * straight back into a clone target without transformation. The
   * lookup is scoped to the axis-level `<c:txPr>` so a stray
   * `<a:defRPr>` inside `<c:title>` (surfaced by
   * {@link ChartAxisInfo.axisTitleFontSize}) cannot leak in.
   */
  labelFontSize?: number;
  /**
   * Tick-label bold flag pulled from
   * `<c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr>`
   * on the axis element. Reflects Excel's "Format Axis -> Font ->
   * Bold" toggle applied to tick labels.
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `<a:defRPr b="0"/>` round-trip identically through
   * {@link cloneChart} — only an explicit `<a:defRPr b="1"/>` surfaces
   * `true`. The reader accepts the OOXML truthy / falsy spellings
   * (`"1"` / `"true"` / `"0"` / `"false"`); unknown values and missing
   * `b` attributes drop to `undefined`.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
   * the OOXML schema. Mirrors the writer-side
   * {@link SheetChart.axes.x.labelBold} so a parsed value slots
   * straight back into a clone target without transformation. The
   * lookup is scoped to the axis-level `<c:txPr>` so a stray
   * `<a:defRPr b=".."/>` inside `<c:title>` (surfaced by
   * {@link ChartAxisInfo.axisTitleBold}) cannot leak in.
   */
  labelBold?: boolean;
  /**
   * Tick-label italic flag pulled from
   * `<c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr>`
   * on the axis element. Reflects Excel's "Format Axis -> Font ->
   * Italic" toggle applied to tick labels.
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `<a:defRPr i="0"/>` round-trip identically through
   * {@link cloneChart} — only an explicit `<a:defRPr i="1"/>` surfaces
   * `true`. The reader accepts the OOXML truthy / falsy spellings
   * (`"1"` / `"true"` / `"0"` / `"false"`); unknown values and missing
   * `i` attributes drop to `undefined`.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
   * the OOXML schema. Mirrors the writer-side
   * {@link SheetChart.axes.x.labelItalic} so a parsed value slots
   * straight back into a clone target without transformation. The
   * lookup is scoped to the axis-level `<c:txPr>` so a stray
   * `<a:defRPr i=".."/>` inside `<c:title>` (surfaced by
   * {@link ChartAxisInfo.axisTitleItalic}) cannot leak in.
   */
  labelItalic?: boolean;
  /**
   * Tick-label font color pulled from
   * `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>
   * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
   * </c:txPr>` on the axis element. Reflects Excel's "Format Axis ->
   * Font -> Font Color" picker applied to tick labels.
   *
   * Surfaced as the 6-character uppercase hex string the writer round-
   * trips (`"FF0000"` / `"1070CA"`) — the leading `#` is stripped on
   * read so the value threads straight into the writer-side
   * {@link SheetChart.axes.x.labelColor} without transformation. Color
   * picks other than the literal sRGB form (`<a:schemeClr>` theme
   * references, `<a:hslClr>`, `<a:sysClr>`, `<a:prstClr>`) collapse
   * to `undefined` — the reader records only the resolvable RGB
   * triple to keep the round-trip lossless against {@link cloneChart}
   * -> {@link writeXlsx}. Malformed `val` tokens (wrong length,
   * non-hex characters) likewise drop to `undefined` rather than
   * fabricate a value the writer would round-trip into a malformed
   * `<a:srgbClr>`.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
   * the OOXML schema. Mirrors the writer-side
   * {@link SheetChart.axes.x.labelColor} so a parsed value slots
   * straight back into a clone target without transformation. The
   * lookup is scoped to the axis-level `<c:txPr>` so a stray
   * `<a:solidFill>` inside `<c:title>` (surfaced by
   * {@link ChartAxisInfo.axisTitleColor}) or on a `<c:spPr>` series
   * fill cannot leak in.
   */
  labelColor?: string;
  /**
   * Tick-label underline flag pulled from
   * `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>`
   * on the axis element. Reflects Excel's "Format Axis -> Font ->
   * Underline" toggle applied to tick labels.
   *
   * The OOXML `u` attribute is the `ST_TextUnderlineType` enum on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
   * eighteen values; Excel's UI exposes only `"sng"` (single line —
   * the default underline checkbox) and `"dbl"` (double line). The
   * reader surfaces only the UI-default `"sng"` as `true`; `"none"`
   * (the OOXML application default), absence, the non-UI `"dbl"`
   * variant, and the sixteen exotic tokens (`"words"`, `"heavy"`,
   * `"dotted"`, `"dottedHeavy"`, `"dash"`, `"dashHeavy"`,
   * `"dashLong"`, `"dashLongHeavy"`, `"dotDash"`, `"dotDashHeavy"`,
   * `"dotDotDash"`, `"dotDotDashHeavy"`, `"wavy"`, `"wavyHeavy"`,
   * `"wavyDbl"`) all collapse to `undefined` — the writer emits only
   * `"sng"`, so reporting any non-single underline as `true` would
   * silently downgrade the choice to a single line on round-trip.
   * Unknown / malformed `u` tokens likewise drop to `undefined`.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
   * the OOXML schema. Mirrors the writer-side
   * {@link SheetChart.axes.x.labelUnderline} so a parsed value slots
   * straight back into a clone target without transformation. The
   * lookup is scoped to the axis-level `<c:txPr>` so a stray
   * `<a:defRPr u=".."/>` inside `<c:title>` (surfaced by
   * {@link ChartAxisInfo.axisTitleUnderline}) cannot leak in.
   */
  labelUnderline?: boolean;
  /**
   * Tick-label strikethrough flag pulled from
   * `<c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr>`
   * on the axis element. Reflects Excel's "Format Axis -> Font ->
   * Strikethrough" toggle applied to tick labels.
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
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
   * the OOXML schema. Mirrors the writer-side
   * {@link SheetChart.axes.x.labelStrike} so a parsed value slots
   * straight back into a clone target without transformation. The
   * lookup is scoped to the axis-level `<c:txPr>` so a stray
   * `<a:defRPr strike=".."/>` inside `<c:title>` (surfaced by
   * {@link ChartAxisInfo.axisTitleStrike}) cannot leak in.
   */
  labelStrike?: boolean;
  /**
   * Axis tick-label font family / typeface pulled from `<c:catAx>` /
   * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` ->
   * `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/>
   * </a:defRPr></a:pPr></a:p></c:txPr>`. Reflects Excel's "Format
   * Axis -> Number / Font -> Font" picker scoped to the axis's tick
   * labels. The OOXML `<a:latin typeface=".."/>` element carries the
   * typeface name (`CT_TextFont`, ECMA-376 Part 1, §21.1.2.3.7).
   *
   * Reports the trimmed typeface string when the source axis pinned
   * a non-empty typeface; absence and empty / whitespace-only
   * `typeface` attributes both collapse to `undefined` so absence
   * and `<a:latin typeface=""/>` round-trip identically through
   * {@link cloneChart}. Non-string `typeface` tokens (defensive —
   * the XML parser only ever surfaces strings) likewise drop to
   * `undefined`.
   *
   * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` all carry an optional `<c:txPr>` per
   * the OOXML schema. Mirrors the writer-side
   * {@link SheetChart.axes.x.labelFontFamily} so a parsed value slots
   * straight back into a clone target without transformation. The
   * lookup is scoped to the axis-level `<c:txPr>` so a stray
   * `<a:latin>` inside `<c:title>` (surfaced by
   * {@link ChartAxisInfo.axisTitleFontFamily}) cannot leak in.
   */
  labelFontFamily?: string;
  /**
   * Reverse-axis flag pulled from
   * `<c:scaling><c:orientation val=".."/></c:scaling>`. Surfaces `true`
   * only when the axis pinned `"maxMin"` (Excel's "Categories /
   * Values in reverse order" toggle); the OOXML default `"minMax"`
   * collapses to `undefined` so absence and the default round-trip
   * identically through {@link cloneChart}. Mirrors the writer-side
   * {@link SheetChart.axes.x.reverse} field, so a parsed value slots
   * straight back into a clone target without transformation.
   */
  reverse?: boolean;
  /**
   * Tick-label skip interval pulled from `<c:tickLblSkip val=".."/>`.
   * Surfaces only on category axes (`<c:catAx>` / `<c:dateAx>`) — the
   * OOXML schema does not place this element on `<c:valAx>`. The
   * default `1` (show every label) collapses to `undefined` so absence
   * and the default round-trip identically through {@link cloneChart}.
   * Out-of-range values (non-positive or > 32767) are dropped rather
   * than fabricated.
   */
  tickLblSkip?: number;
  /**
   * Tick-mark skip interval pulled from `<c:tickMarkSkip val=".."/>`.
   * Same scope (category axes only) and default-collapse semantics as
   * {@link tickLblSkip}.
   */
  tickMarkSkip?: number;
  /**
   * Label offset pulled from `<c:lblOffset val=".."/>`, expressed as a
   * percentage of the default axis-label spacing. Surfaces only on
   * category axes (`<c:catAx>` / `<c:dateAx>`) — the OOXML schema
   * (`ST_LblOffsetPercent`) does not place this element on `<c:valAx>`
   * or `<c:serAx>`. The default `100` (Excel's reference spacing)
   * collapses to `undefined` so absence and the default round-trip
   * identically through {@link cloneChart}. Accepted range is `0..1000`;
   * out-of-range values are dropped rather than fabricated.
   */
  lblOffset?: number;
  /**
   * Tick-label horizontal alignment pulled from `<c:lblAlgn val=".."/>`.
   * Surfaces only on category axes (`<c:catAx>` / `<c:dateAx>`) — the
   * OOXML schema (`ST_LblAlgn`) does not place this element on
   * `<c:valAx>` or `<c:serAx>`. The default `"ctr"` (Excel's reference
   * centered alignment) collapses to `undefined` so absence and the
   * default round-trip identically through {@link cloneChart}. Unknown
   * tokens drop to `undefined` rather than fabricate a value the
   * writer would never emit. See {@link ChartAxisLabelAlign}.
   */
  lblAlgn?: ChartAxisLabelAlign;
  /**
   * Multi-level-label suppression flag pulled from
   * `<c:noMultiLvlLbl val=".."/>`. Surfaces `true` only when the axis
   * pinned `val="1"` (Excel's "Multi-level Category Labels" checkbox
   * unchecked — every category collapses onto one line). The OOXML
   * default `val="0"` (and absence of the element) collapse to
   * `undefined` so absence and the default round-trip identically
   * through {@link cloneChart}.
   *
   * Surfaces only on category axes (`<c:catAx>`) — the OOXML schema
   * places the element on `CT_CatAx` exclusively (it has no slot on
   * `CT_ValAx`, `CT_DateAx`, or `CT_SerAx`). The reader accepts the
   * OOXML truthy / falsy spellings (`"1"` / `"true"` / `"0"` /
   * `"false"`); unknown values and missing `val` attributes drop to
   * `undefined`.
   */
  noMultiLvlLbl?: boolean;
  /**
   * Automatic axis-type detection flag pulled from
   * `<c:auto val=".."/>`. Surfaces `false` only when the axis pinned
   * `val="0"` (Excel's "Text axis" radio under "Format Axis -> Axis
   * Options -> Axis Type" — Excel keeps every label as-is regardless
   * of whether the cells parse as dates / numerics). The OOXML default
   * `val="1"` (and absence of the element) collapse to `undefined` so
   * absence and the default round-trip identically through
   * {@link cloneChart}.
   *
   * Surfaces only on category axes (`<c:catAx>`) — the OOXML schema
   * places the element on `CT_CatAx` exclusively. The reader accepts
   * the OOXML truthy / falsy spellings (`"1"` / `"true"` / `"0"` /
   * `"false"`); unknown values and missing `val` attributes drop to
   * `undefined`.
   */
  auto?: boolean;
  /**
   * Axis hidden flag pulled from `<c:delete val=".."/>`. Surfaces
   * `true` when the axis pinned `val="1"` (Excel's "Format Axis ->
   * Show axis = off" toggle). The OOXML default `val="0"` (and absence
   * of the element) collapse to `undefined` so absence and the default
   * round-trip identically through {@link cloneChart}. The reader
   * accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` /
   * `"0"` / `"false"`); unknown values and missing `val` attributes
   * drop to `undefined`.
   */
  hidden?: boolean;
  /**
   * Semantic crossing position pulled from `<c:crosses val=".."/>`.
   * Surfaces only when the axis pinned a non-default token — the OOXML
   * default `"autoZero"` collapses to `undefined` so absence and the
   * default round-trip identically through {@link cloneChart}. Unknown
   * tokens drop rather than fabricate a value the writer would never
   * emit. See {@link ChartAxisCrosses}.
   *
   * Mutually exclusive with {@link crossesAt} in the OOXML schema —
   * when both elements appear on the same axis (a malformed template),
   * the reader keeps `crossesAt` and drops `crosses` to mirror the
   * writer's preference.
   */
  crosses?: ChartAxisCrosses;
  /**
   * Numeric crossing position pulled from `<c:crossesAt val=".."/>`.
   * Surfaces the literal value Excel paints — `0` is preserved (it is a
   * valid pin, distinct from the `"autoZero"` default). Non-numeric
   * `val` attributes drop to `undefined` rather than fabricate a value
   * the writer would never emit.
   *
   * Mutually exclusive with {@link crosses} in the OOXML schema (CT_Double
   * sits in an XSD choice with CT_Crosses). When both elements appear on
   * the same axis (a malformed template) the reader keeps `crossesAt`
   * and drops `crosses` to mirror the writer's preference.
   */
  crossesAt?: number;
  /**
   * Built-in display-unit preset pulled from
   * `<c:dispUnits><c:builtInUnit val=".."/><c:dispUnitsLbl?/></c:dispUnits>`.
   * Surfaces only on value axes — the OOXML schema places `<c:dispUnits>`
   * exclusively on `CT_ValAx`, so `<c:catAx>` / `<c:dateAx>` / `<c:serAx>`
   * never carry one. The reader keeps the parsed `unit` token and the
   * presence of `<c:dispUnitsLbl>` (`showLabel`); the OOXML alternative
   * `<c:custUnit val=".."/>` (custom numeric divisor) and any rich-text
   * `<c:dispUnitsLbl>` body are intentionally not surfaced.
   *
   * The OOXML schema accepts the nine `ST_BuiltInUnit` tokens listed in
   * {@link ChartAxisDispUnit}; unknown tokens drop to `undefined` rather
   * than fabricate a value the writer would never emit. Absence
   * (and any unrecognized payload) collapses to `undefined` so a
   * round-trip leaves Excel's default "no display unit" state untouched.
   */
  dispUnits?: ChartAxisDispUnits;
  /**
   * Cross-between mode pulled from `<c:crossBetween val=".."/>`.
   * Surfaces only on value axes — the OOXML schema places the element
   * exclusively on `CT_ValAx`, so `<c:catAx>` / `<c:dateAx>` /
   * `<c:serAx>` never carry one. Unknown / typo'd tokens drop to
   * `undefined` rather than fabricate a value the writer would never
   * emit; absence likewise collapses to `undefined` so a chart that
   * inherited Excel's default still round-trips minimally through
   * {@link cloneChart}. See {@link ChartAxisCrossBetween}.
   */
  crossBetween?: ChartAxisCrossBetween;
}

/**
 * A chart anchored on a sheet via the sheet's drawing part.
 *
 * Charts come from `xl/charts/chartN.xml`. Hucre exposes the
 * structural metadata needed to recognize, introspect, and clone the
 * chart; the chart body is preserved verbatim through roundtrip.
 */
/**
 * Legend placement reported by {@link Chart.legend}.
 *
 * Values mirror the {@link SheetChart.legend} options on the writer
 * side, so a parsed legend position slots straight back into a clone
 * target. `false` is reported when the chart explicitly omits the
 * legend element (Excel's "no legend" state); `undefined` means the
 * chart did not declare a legend at all.
 */
export type ChartLegendPosition = "top" | "bottom" | "left" | "right" | "topRight";

/**
 * Bar/column grouping reported by {@link Chart.barGrouping}.
 *
 * Pulled from `<c:barChart><c:grouping val="..."/></c:barChart>`.
 * `"standard"` is the OOXML value for non-stacked, non-percent layouts
 * — it is excluded here because the writer's
 * {@link SheetChart.barGrouping} models the same default as the
 * absence of the field. Only the stacked variants surface, which is
 * what callers need to detect when cloning a stacked template.
 */
export type ChartBarGrouping = "clustered" | "stacked" | "percentStacked";

/**
 * Line/area grouping reported by {@link Chart.lineGrouping} and
 * {@link Chart.areaGrouping}.
 *
 * Pulled from `<c:lineChart><c:grouping val="..."/></c:lineChart>` or
 * `<c:areaChart><c:grouping val="..."/></c:areaChart>`. Only the
 * stacked variants surface — `"standard"` is the OOXML default and
 * is collapsed to `undefined` for symmetry with the writer's
 * {@link SheetChart.lineGrouping} / {@link SheetChart.areaGrouping}
 * defaults.
 */
export type ChartLineAreaGrouping = "stacked" | "percentStacked";

export interface Chart {
  /** Chart-type elements present in `<c:plotArea>`, in declaration order. */
  kinds: ChartKind[];
  /** Number of `<c:ser>` series across every chart-type element. */
  seriesCount: number;
  /** Plain-text title pulled from `<c:title>`, when present. */
  title?: string;
  /**
   * Per-series metadata across every chart-type element, in
   * declaration order. Empty when the chart has no `<c:ser>` children.
   */
  series?: ChartSeriesInfo[];
  /**
   * Cell anchor pulled from the host drawing's `<xdr:twoCellAnchor>` /
   * `<xdr:oneCellAnchor>`. Undefined when the drawing positions the
   * chart with `<xdr:absoluteAnchor>` (EMU-positioned, no cell anchor)
   * or when the drawing's anchor element is missing the `from` block.
   */
  anchor?: ChartAnchor;
  /**
   * Legend placement pulled from `<c:legend><c:legendPos val=".."/>`.
   * Reported as `false` when the chart explicitly omits the legend
   * element (Excel's "no legend" state). `undefined` means the chart
   * did not declare a legend at all — Excel falls back to its default
   * placement in that case.
   */
  legend?: false | ChartLegendPosition;
  /**
   * Legend-overlay flag pulled from `<c:legend><c:overlay val=".."/>`.
   * Reflects Excel's "Format Legend -> Show the legend without
   * overlapping the chart" toggle (the checkbox is the inverse — checked
   * means `false`, unchecked means `true`).
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `<c:overlay val="0"/>` round-trip identically through
   * {@link cloneChart} — only an explicit `<c:overlay val="1"/>` surfaces
   * `true`. The reader accepts the OOXML truthy / falsy spellings (`"1"`
   * / `"true"` / `"0"` / `"false"`); unknown values and missing `val`
   * attributes drop to `undefined`.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there is no
   * overlay flag to surface in either case.
   */
  legendOverlay?: boolean;
  /**
   * Per-series legend-entry overrides pulled from
   * `<c:legend><c:legendEntry>` children. Each entry surfaces the
   * 0-based series {@link ChartLegendEntry.idx} the chart targeted and
   * the {@link ChartLegendEntry.delete} flag.
   *
   * The reader emits an entry only when the source chart actually
   * declares a `<c:legendEntry>` block — absence collapses to
   * `undefined` (the field is omitted entirely) so a chart with no
   * overrides round-trips minimally through {@link cloneChart}. Entries
   * whose `<c:idx>` is missing or invalid are dropped rather than
   * surface a fabricated index.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there are no
   * legend entries to surface in either case.
   */
  legendEntries?: ChartLegendEntry[];
  /**
   * Legend font size in points pulled from
   * `<c:legend><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
   * </c:txPr></c:legend>`. The OOXML `sz` attribute is in 100ths of a
   * point — the reader converts to points and rounds to the nearest
   * 0.5pt (Excel's UI exposes the same 0.5pt granularity). Range:
   * `1..400`pt (the band the OOXML `ST_TextFontSize` schema exposes).
   *
   * Absence of the element / attribute and out-of-range / non-numeric
   * / non-finite values all collapse to `undefined` so a fresh chart
   * and a chart that pinned an out-of-range size both round-trip to
   * the writer's "skip the size attribute" path.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there is no
   * `<c:txPr>` slot to surface the size from in either case. Mirrors
   * the writer-side {@link SheetChart.legendFontSize} so a parsed value
   * slots straight into {@link cloneChart} without conversion.
   */
  legendFontSize?: number;
  /**
   * Legend bold flag pulled from `<c:legend><c:txPr><a:p><a:pPr>
   * <a:defRPr b=".."/></a:pPr></a:p></c:txPr></c:legend>`. The OOXML
   * `b` attribute is the `xsd:boolean` bold flag on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7).
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `b="0"` round-trip identically — only an explicit `b="1"` surfaces
   * `true`. Unknown / malformed `b` tokens drop to `undefined` rather
   * than fabricate a value the writer would never emit.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there is no
   * `<c:txPr>` slot to surface the flag from in either case. Mirrors
   * the writer-side {@link SheetChart.legendBold} so a parsed value
   * slots straight into {@link cloneChart} without conversion.
   */
  legendBold?: boolean;
  /**
   * Legend italic flag pulled from `<c:legend><c:txPr><a:p><a:pPr>
   * <a:defRPr i=".."/></a:pPr></a:p></c:txPr></c:legend>`. The OOXML
   * `i` attribute is the `xsd:boolean` italic flag on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7).
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `i="0"` round-trip identically — only an explicit `i="1"` surfaces
   * `true`. Unknown / malformed `i` tokens drop to `undefined` rather
   * than fabricate a value the writer would never emit.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there is no
   * `<c:txPr>` slot to surface the flag from in either case. Mirrors
   * the writer-side {@link SheetChart.legendItalic} so a parsed value
   * slots straight into {@link cloneChart} without conversion.
   */
  legendItalic?: boolean;
  /**
   * Legend underline flag pulled from `<c:legend><c:txPr><a:p><a:pPr>
   * <a:defRPr u=".."/></a:pPr></a:p></c:txPr></c:legend>`. The OOXML
   * `u` attribute is the `ST_TextUnderlineType` enumeration on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7).
   *
   * The OOXML default `"none"` collapses to `undefined` so absence and
   * `u="none"` round-trip identically — only `u="sng"` (Excel's UI
   * variant — single underline) surfaces `true`. Unknown / malformed
   * `u` tokens (`"dbl"`, `"dotted"`, etc.) drop to `undefined` rather
   * than fabricate a value the writer would never emit; the writer
   * only emits `u="sng"` / `u="none"`, matching the boolean shape the
   * UI exposes.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there is no
   * `<c:txPr>` slot to surface the flag from in either case. Mirrors
   * the writer-side {@link SheetChart.legendUnderline} so a parsed
   * value slots straight into {@link cloneChart} without conversion.
   */
  legendUnderline?: boolean;
  /**
   * Legend strikethrough flag pulled from `<c:legend><c:txPr><a:p>
   * <a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr></c:legend>`.
   * The OOXML `strike` attribute is the `ST_TextStrikeType` enumeration
   * on `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7).
   *
   * The OOXML default `"noStrike"` collapses to `undefined` so absence
   * and `strike="noStrike"` round-trip identically — only
   * `strike="sngStrike"` (Excel's UI variant — single line) surfaces
   * `true`. The non-UI variant `strike="dblStrike"` (double line) is
   * read-only and collapses to `undefined` so a templated chart that
   * pinned the double-line variant in raw OOXML round-trips lossless
   * rather than silently downgrading on re-emit; the writer emits only
   * `"sngStrike"`, matching the boolean shape the UI exposes.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there is no
   * `<c:txPr>` slot to surface the flag from in either case. Mirrors
   * the writer-side {@link SheetChart.legendStrikethrough} so a parsed
   * value slots straight into {@link cloneChart} without conversion.
   */
  legendStrikethrough?: boolean;
  /**
   * Legend font color pulled from `<c:legend><c:txPr><a:p><a:pPr>
   * <a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
   * </a:defRPr></a:pPr></a:p></c:txPr></c:legend>`. The OOXML
   * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
   * sRGB color (`CT_SRgbColor` inside `CT_TextCharacterProperties`'
   * fill choice — ECMA-376 Part 1, §20.1.2.3.32 / §21.1.2.3.7).
   *
   * Returned as the canonical 6-character uppercase hex string when
   * the parser walks the full chain and lands on an `<a:srgbClr
   * val="RRGGBB"/>`. Theme references (`<a:schemeClr>`),
   * `<a:hslClr>`, `<a:sysClr>`, `<a:prstClr>`, and malformed `val`
   * tokens (wrong length, non-hex characters) all collapse to
   * `undefined` since only the literal RGB triple round-trips
   * losslessly through {@link writeChart}.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there is no
   * `<c:txPr>` slot to surface the fill from in either case. Mirrors
   * the writer-side {@link SheetChart.legendFontColor} so a parsed
   * value slots straight into {@link cloneChart} without conversion.
   */
  legendFontColor?: string;
  /**
   * Legend font family / typeface pulled from `<c:legend><c:txPr>
   * <a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
   * </a:pPr></a:p></c:txPr></c:legend>`. Reflects Excel's "Format
   * Legend -> Font -> Font" picker. The OOXML `<a:latin
   * typeface=".."/>` element carries the typeface name (`CT_TextFont`,
   * ECMA-376 Part 1, §21.1.2.3.7).
   *
   * Reports the trimmed typeface string when the source chart pinned
   * a non-empty typeface; absence and empty / whitespace-only
   * `typeface` attributes both collapse to `undefined` so absence and
   * `<a:latin typeface=""/>` round-trip identically through
   * {@link cloneChart}. Non-string `typeface` tokens (defensive — the
   * XML parser only ever surfaces strings) likewise drop to
   * `undefined`.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there is no
   * `<c:txPr>` slot to surface the typeface from in either case.
   * Mirrors the writer-side {@link SheetChart.legendFontFamily} so a
   * parsed value slots straight into {@link cloneChart} without
   * conversion.
   */
  legendFontFamily?: string;
  /**
   * Custom legend placement pulled from `<c:legend><c:layout>
   * <c:manualLayout>...</c:manualLayout></c:layout></c:legend>`.
   * Reflects Excel's "Format Legend -> Position -> Custom" knob — the
   * `(x, y)` anchor and `(w, h)` size of the legend block as fractions
   * of the chart frame in the `0..1` band.
   *
   * Each of {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
   * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} surfaces
   * the literal `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` value when the
   * source chart pins one; absence / non-numeric / non-finite tokens
   * collapse to `undefined` on the matching field so absence and a
   * malformed token round-trip identically through {@link cloneChart}.
   * The reader accepts both `xMode="edge"` (absolute fraction of the
   * chart frame) and `xMode="factor"` (delta from auto-layout) and
   * surfaces the same shape; the writer normalizes to `"edge"` on emit
   * since that is the form Excel itself emits when the user drags a
   * legend to a custom position.
   *
   * Reported as `undefined` whenever {@link legend} is `false`, the
   * source chart has no `<c:legend>` element at all, or every
   * coordinate the source pinned drops to `undefined` — there is no
   * meaningful layout to surface in any of those cases. Mirrors the
   * writer-side {@link SheetChart.legendLayout} so a parsed value
   * slots straight into {@link cloneChart} without conversion.
   */
  legendLayout?: ChartManualLayout;
  /**
   * Legend background fill color pulled from `<c:legend><c:spPr>
   * <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></c:spPr>
   * </c:legend>`. Reflects Excel's "Format Legend -> Fill -> Solid
   * fill -> Color" picker (the same dialog the user reaches by
   * right-clicking the legend background). The element sits on
   * `<c:legend>` between `<c:overlay>` and `<c:txPr>` per the
   * CT_Legend schema sequence (ECMA-376 Part 1, §21.2.2.114).
   *
   * Reports the 6-character uppercase hex string when the source
   * chart pins a literal `<a:srgbClr val="RRGGBB"/>` fill on the
   * legend's `<c:spPr>` block. Theme references (`<a:schemeClr>`),
   * non-solid fills (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>` /
   * `<a:blipFill>`), and the OOXML system / preset color forms
   * (`<a:sysClr>` / `<a:hslClr>` / `<a:prstClr>`) all collapse to
   * `undefined` — only the literal RGB triple round-trips losslessly
   * through {@link writeChart}. Malformed `val` tokens (wrong length,
   * non-hex characters) likewise drop to `undefined` rather than
   * fabricate a value the writer would round-trip into a malformed
   * `<a:srgbClr>`.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there is no
   * `<c:spPr>` slot to surface the fill from in either case. Mirrors
   * the writer-side {@link SheetChart.legendFillColor} so a parsed
   * value slots straight into {@link cloneChart} without conversion.
   */
  legendFillColor?: string;
  /**
   * Legend border (stroke) solid color pulled from `<c:legend><c:spPr>
   * <a:ln><a:solidFill><a:srgbClr val=".."/></a:solidFill></a:ln>
   * </c:spPr></c:legend>`. Reflects Excel's "Format Legend -> Border
   * -> Solid line -> Color" picker. The OOXML `<a:srgbClr val=".."/>`
   * carries the 6-character uppercase hex sRGB color (`CT_SRgbColor`
   * inside the line's solid fill choice — ECMA-376 Part 1, §20.1.2.3.32
   * / §20.1.2.3.24). The `<c:spPr>` slot sits between `<c:overlay>` and
   * `<c:txPr>` per the CT_Legend schema sequence (ECMA-376 Part 1,
   * §21.2.2.114); `<a:ln>` follows the optional `<a:solidFill>` (fill)
   * child inside `<c:spPr>` per `CT_ShapeProperties` (ECMA-376 Part 1,
   * §20.1.2.3.13).
   *
   * Reports the 6-character uppercase hex form when the source chart
   * pinned a literal `<a:srgbClr>` color on `<a:ln>`; absence,
   * malformed `val` tokens (wrong length, non-hex characters,
   * alpha-channel forms), non-solid line fills (`<a:noFill>`,
   * `<a:gradFill>`, `<a:pattFill>`), and theme-color references
   * (`<a:schemeClr>`) all collapse to `undefined` so absence and an
   * unrenderable stroke round-trip identically through
   * {@link cloneChart}. Mirrors the writer-side
   * {@link SheetChart.legendBorderColor} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there is no
   * `<c:spPr>` slot to surface the stroke from in either case. Composes
   * independently with {@link legendFillColor} — the two fields surface
   * from the same `<c:spPr>` block but on different children
   * (`<a:solidFill>` for fill, `<a:ln><a:solidFill>` for stroke).
   */
  legendBorderColor?: string;
  /**
   * Legend border (stroke) thickness in points pulled from the `w`
   * attribute on `<c:legend><c:spPr><a:ln w="EMU">`. Reflects Excel's
   * "Format Legend -> Border -> Width" spinner. The OOXML `w`
   * attribute stores the stroke width in English Metric Units
   * (1 pt = 12 700 EMU) per `CT_LineProperties` (ECMA-376 Part 1,
   * §20.1.2.3.24); the reader divides by 12 700 and snaps the result
   * to the 0.25 pt grid Excel's UI exposes so a parsed-then-cloned
   * width does not drift across round-trips.
   *
   * Reports the point value clamped to the `0.25..13.5` pt band Excel
   * accepts in the UI when the source chart pinned a finite, positive
   * `w` attribute. Absence (no `<a:ln>` or `<a:ln>` without a `w`
   * attribute), zero, negative, and non-numeric `w` values all collapse
   * to `undefined` so absence and an unrenderable width round-trip
   * identically through {@link cloneChart}. Mirrors the writer-side
   * {@link SheetChart.legendBorderWidth} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   *
   * Reported as `undefined` whenever {@link legend} is `false` or the
   * source chart has no `<c:legend>` element at all — there is no
   * `<c:spPr>` slot to surface the stroke from in either case. Composes
   * independently with {@link legendBorderColor} — both fields surface
   * from the same `<a:ln>` element but on a different slot (the color
   * child versus the width attribute).
   */
  legendBorderWidth?: number;
  /**
   * Chart legend border (stroke) preset dash pattern pulled from the
   * `val` attribute on `<c:legend><c:spPr><a:ln><a:prstDash val=".."/>`.
   * Reflects Excel's "Format Legend -> Border -> Dash type" picker.
   * Reports the {@link ChartBorderDash} value pinned by the source, or
   * `undefined` when the element is absent / the OOXML default
   * `"solid"` was authored / the value is unrecognized.
   *
   * Mirrors the writer-side {@link SheetChart.legendBorderDash} so a
   * parsed value slots straight into {@link cloneChart} without
   * conversion.
   */
  legendBorderDash?: ChartBorderDash;
  /**
   * Custom plot-area placement pulled from `<c:plotArea><c:layout>
   * <c:manualLayout>...</c:manualLayout></c:layout></c:plotArea>`.
   * Reflects Excel's "Format Plot Area -> Position -> Custom" placement —
   * the `(x, y)` anchor and `(w, h)` size of the plot area as fractions
   * of the chart frame in the `0..1` band.
   *
   * Each of {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
   * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} surfaces
   * the literal `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` value when the
   * source chart pins one; absence / non-numeric / non-finite tokens
   * collapse to `undefined` on the matching field so absence and a
   * malformed token round-trip identically through {@link cloneChart}.
   * The reader accepts both `xMode="edge"` (absolute fraction of the
   * chart frame) and `xMode="factor"` (delta from auto-layout) and
   * surfaces the same shape; the writer normalizes to `"edge"` on emit
   * since that is the form Excel itself emits when the user drags the
   * plot area to a custom position.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:plotArea>` element at all, the `<c:plotArea>` carries only the
   * bare `<c:layout/>` placeholder (Excel's reference shape for an
   * auto-layout chart), or every coordinate the source pinned drops to
   * `undefined` — there is no meaningful layout to surface in any of
   * those cases. Mirrors the writer-side {@link SheetChart.plotAreaLayout}
   * so a parsed value slots straight into {@link cloneChart} without
   * conversion.
   */
  plotAreaLayout?: ChartManualLayout;
  /**
   * Plot-area solid fill color pulled from `<c:plotArea><c:spPr>
   * <a:solidFill><a:srgbClr val=".."/></a:solidFill></c:spPr></c:plotArea>`.
   * Reflects Excel's "Format Plot Area -> Fill -> Solid fill -> Color"
   * picker. The OOXML `<a:srgbClr val=".."/>` carries the 6-character
   * uppercase hex sRGB color (`CT_SRgbColor` inside `CT_ShapeProperties`'
   * fill choice — ECMA-376 Part 1, §20.1.2.3.32 / §20.1.2.3.13). The
   * `<c:spPr>` slot lives at the tail of `<c:plotArea>` per
   * `CT_PlotArea` (ECMA-376 Part 1, §21.2.2.145), after every chart-type
   * element / axes / `<c:dTable>`.
   *
   * Reports the 6-character uppercase hex form when the source chart
   * pinned a literal `<a:srgbClr>` color; absence, malformed `val`
   * tokens (wrong length, non-hex characters, alpha-channel forms),
   * non-solid fills (`<a:noFill>`, `<a:gradFill>`, `<a:pattFill>`,
   * `<a:blipFill>`), and theme-color references (`<a:schemeClr>`) all
   * collapse to `undefined` so absence and an unrenderable fill
   * round-trip identically through {@link cloneChart}. Mirrors the
   * writer-side {@link SheetChart.plotAreaFillColor} so a parsed value
   * slots straight into {@link cloneChart} without conversion.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:plotArea>` element at all — there is no `<c:spPr>` slot to
   * surface the fill from in that case.
   */
  plotAreaFillColor?: string;
  /**
   * Plot-area border (stroke) solid color pulled from
   * `<c:plotArea><c:spPr><a:ln><a:solidFill><a:srgbClr val=".."/>
   * </a:solidFill></a:ln></c:spPr></c:plotArea>`. Reflects Excel's
   * "Format Plot Area -> Border -> Solid line -> Color" picker. The
   * OOXML `<a:srgbClr val=".."/>` carries the 6-character uppercase
   * hex sRGB color (`CT_SRgbColor` inside the line's solid fill choice
   * — ECMA-376 Part 1, §20.1.2.3.32 / §20.1.2.3.24). The `<c:spPr>`
   * slot lives at the tail of `<c:plotArea>` per `CT_PlotArea`
   * (ECMA-376 Part 1, §21.2.2.145), after every chart-type element /
   * axes / `<c:dTable>`; `<a:ln>` follows the optional `<a:solidFill>`
   * fill child inside `<c:spPr>` per `CT_ShapeProperties` (ECMA-376
   * Part 1, §20.1.2.3.13).
   *
   * Reports the 6-character uppercase hex form when the source chart
   * pinned a literal `<a:srgbClr>` color on `<a:ln>`; absence,
   * malformed `val` tokens (wrong length, non-hex characters,
   * alpha-channel forms), non-solid line fills (`<a:noFill>`,
   * `<a:gradFill>`, `<a:pattFill>`), and theme-color references
   * (`<a:schemeClr>`) all collapse to `undefined` so absence and an
   * unrenderable stroke round-trip identically through
   * {@link cloneChart}. Mirrors the writer-side
   * {@link SheetChart.plotAreaBorderColor} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:plotArea>` element at all — there is no `<c:spPr>` slot to
   * surface the stroke from in that case. Composes independently with
   * {@link plotAreaFillColor} — the two fields surface from the same
   * `<c:spPr>` block but on different children (`<a:solidFill>` for
   * fill, `<a:ln><a:solidFill>` for stroke).
   */
  plotAreaBorderColor?: string;
  /**
   * Plot-area border (stroke) thickness in points pulled from the `w`
   * attribute on `<c:plotArea><c:spPr><a:ln w="EMU">`. Reflects Excel's
   * "Format Plot Area -> Border -> Width" spinner. The OOXML `w`
   * attribute stores the stroke width in English Metric Units
   * (1 pt = 12 700 EMU) per `CT_LineProperties` (ECMA-376 Part 1,
   * §20.1.2.3.24); the reader divides by 12 700 and snaps the result
   * to the 0.25 pt grid Excel's UI exposes so a parsed-then-cloned
   * width does not drift across round-trips.
   *
   * Reports the point value clamped to the `0.25..13.5` pt band Excel
   * accepts in the UI when the source chart pinned a finite, positive
   * `w` attribute. Absence (no `<a:ln>` or `<a:ln>` without a `w`
   * attribute), zero, negative, and non-numeric `w` values all collapse
   * to `undefined` so absence and an unrenderable width round-trip
   * identically through {@link cloneChart}. Mirrors the writer-side
   * {@link SheetChart.plotAreaBorderWidth} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:plotArea>` element at all — there is no `<c:spPr>` slot to
   * surface the stroke from in that case. Composes independently with
   * {@link plotAreaBorderColor} — both fields surface from the same
   * `<a:ln>` element but on a different slot (the color child versus
   * the width attribute).
   */
  plotAreaBorderWidth?: number;
  /**
   * Plot-area border (stroke) preset dash pattern pulled from the
   * `val` attribute on `<c:plotArea><c:spPr><a:ln><a:prstDash
   * val=".."/>`. Reflects Excel's "Format Plot Area -> Border -> Dash
   * type" picker. Reports the {@link ChartBorderDash} value that the
   * source chart pinned, or `undefined` when the element is absent /
   * the source chart authored the OOXML default `"solid"` / the value
   * is unrecognized.
   *
   * Mirrors the writer-side {@link SheetChart.plotAreaBorderDash} so a
   * parsed value slots straight into {@link cloneChart} without
   * conversion.
   */
  plotAreaBorderDash?: ChartBorderDash;
  /**
   * Chart-space (entire chart background) solid fill color pulled
   * from `<c:chartSpace><c:spPr><a:solidFill><a:srgbClr val=".."/>
   * </a:solidFill></c:spPr></c:chartSpace>`. Reflects Excel's
   * "Format Chart Area -> Fill -> Solid fill -> Color" picker — the
   * fill that paints the entire chart frame (title slot, legend slot,
   * axis label margins, and plot area together). The OOXML
   * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
   * sRGB color (`CT_SRgbColor` inside `CT_ShapeProperties`' fill
   * choice — ECMA-376 Part 1, §20.1.2.3.32 / §20.1.2.3.13). The
   * `<c:spPr>` slot lives at the tail of `<c:chartSpace>` per
   * CT_ChartSpace (§21.2.2.29), after `<c:chart>` and the optional
   * `<c:externalData>` / `<c:printSettings>` / `<c:userShapes>` and
   * before the optional `<c:txPr>` / `<c:extLst>`.
   *
   * Reports the 6-character uppercase hex form when the source chart
   * pinned a literal `<a:srgbClr>` color; absence, malformed `val`
   * tokens (wrong length, non-hex characters, alpha-channel forms),
   * non-solid fills (`<a:noFill>`, `<a:gradFill>`, `<a:pattFill>`,
   * `<a:blipFill>`), and theme-color references (`<a:schemeClr>`) all
   * collapse to `undefined` so absence and an unrenderable fill
   * round-trip identically through {@link cloneChart}. Mirrors the
   * writer-side {@link SheetChart.chartSpaceFillColor} so a parsed
   * value slots straight into {@link cloneChart} without conversion.
   *
   * Distinct from {@link plotAreaFillColor} — the lookup is scoped to
   * direct children of `<c:chartSpace>` (the document root) so a
   * stray `<c:spPr>` on `<c:plotArea>` / `<c:legend>` / `<c:title>`
   * cannot leak into this field. Mirrors the chart-title /
   * plot-area / legend `<c:spPr>` slots — same accept-or-drop grammar
   * — so a parsed value flows through the same hex shape regardless
   * of which `<c:spPr>`-based fill slot the source chart pinned.
   */
  chartSpaceFillColor?: string;
  /**
   * Chart-space (entire chart frame) border (stroke) color pulled
   * from `<c:chartSpace><c:spPr><a:ln><a:solidFill><a:srgbClr val=".."/>
   * </a:solidFill></a:ln></c:spPr></c:chartSpace>`. Reflects Excel's
   * "Format Chart Area -> Border -> Solid line -> Color" picker — the
   * stroke that paints the outer border of the entire chart frame.
   * The OOXML `<a:srgbClr val=".."/>` carries the 6-character
   * uppercase hex sRGB color (`CT_SRgbColor` inside the line's solid
   * fill choice — ECMA-376 Part 1, §20.1.2.3.32 / §20.1.2.3.24). The
   * `<c:spPr>` slot lives at the tail of `<c:chartSpace>` per
   * CT_ChartSpace (§21.2.2.29); `<a:ln>` follows the optional
   * `<a:solidFill>` (fill) child inside `<c:spPr>` per
   * `CT_ShapeProperties` (§20.1.2.3.13).
   *
   * Reports the 6-character uppercase hex form when the source chart
   * pinned a literal `<a:srgbClr>` color; absence, malformed `val`
   * tokens (wrong length, non-hex characters, alpha-channel forms),
   * non-solid line fills (`<a:noFill>`, `<a:gradFill>`, `<a:pattFill>`,
   * `<a:blipFill>`), and theme-color references (`<a:schemeClr>`) all
   * collapse to `undefined` so absence and an unrenderable stroke
   * round-trip identically through {@link cloneChart}. Mirrors the
   * writer-side {@link SheetChart.chartSpaceBorderColor} so a parsed
   * value slots straight into {@link cloneChart} without conversion.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:chartSpace><c:spPr>` block at all — there is no `<a:ln>` slot
   * to surface the stroke from in that case. Composes independently
   * with {@link chartSpaceFillColor} — the two fields surface from the
   * same `<c:spPr>` block but on different children (`<a:solidFill>`
   * for fill, `<a:ln><a:solidFill>` for stroke). The lookup is scoped
   * to direct children of `<c:chartSpace>` so a stray `<c:spPr>`
   * elsewhere (e.g. on `<c:plotArea>` / `<c:legend>` / `<c:title>` /
   * a series) cannot leak into this field.
   */
  chartSpaceBorderColor?: string;
  /**
   * Chart-space (entire chart frame) border (stroke) thickness in
   * points pulled from the `w` attribute on `<c:chartSpace><c:spPr>
   * <a:ln w="EMU">`. Reflects Excel's "Format Chart Area -> Border ->
   * Width" spinner. The OOXML `w` attribute stores the stroke width
   * in English Metric Units (1 pt = 12 700 EMU) per
   * `CT_LineProperties` (ECMA-376 Part 1, §20.1.2.3.24); the reader
   * divides by 12 700 and snaps the result to the 0.25 pt grid Excel's
   * UI exposes so a parsed-then-cloned width does not drift across
   * round-trips.
   *
   * Reports the point value clamped to the `0.25..13.5` pt band Excel
   * accepts in the UI when the source chart pinned a finite, positive
   * `w` attribute. Absence (no `<a:ln>` or no `w` attribute), zero,
   * negative, and non-numeric `w` values all collapse to `undefined`
   * so absence and an unrenderable width round-trip identically
   * through {@link cloneChart}. Mirrors the writer-side
   * {@link SheetChart.chartSpaceBorderWidth} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:chartSpace><c:spPr>` block at all — there is no `<a:ln>` slot
   * to surface the width from in that case. Composes independently
   * with {@link chartSpaceBorderColor} — both fields surface from the
   * same `<a:ln>` element but on a different slot (the color child
   * versus the width attribute). Mirrors {@link plotAreaBorderWidth} /
   * {@link legendBorderWidth} / {@link titleBorderWidth} — same EMU
   * encoding, same `<a:ln>` host — but lands on `<c:chartSpace>`'s own
   * `<c:spPr>` block.
   */
  chartSpaceBorderWidth?: number;
  /**
   * Chart-space (entire chart frame) border (stroke) preset dash
   * pattern pulled from the `val` attribute on `<c:chartSpace><c:spPr>
   * <a:ln><a:prstDash val=".."/>`. Reflects Excel's "Format Chart Area
   * -> Border -> Dash type" picker. Reports the {@link ChartBorderDash}
   * value that the source chart pinned, or `undefined` when the element
   * is absent / the source chart authored the OOXML default `"solid"` /
   * the value is unrecognized.
   *
   * Mirrors the writer-side {@link SheetChart.chartSpaceBorderDash} so
   * a parsed value slots straight into {@link cloneChart} without
   * conversion.
   */
  chartSpaceBorderDash?: ChartBorderDash;
  /**
   * Title-overlay flag pulled from `<c:title><c:overlay val=".."/>`.
   * Reflects Excel's "Format Chart Title -> Show the title without
   * overlapping the chart" toggle (the checkbox is the inverse — checked
   * means `false`, unchecked means `true`).
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `<c:overlay val="0"/>` round-trip identically through
   * {@link cloneChart} — only an explicit `<c:overlay val="1"/>` surfaces
   * `true`. The reader accepts the OOXML truthy / falsy spellings (`"1"`
   * / `"true"` / `"0"` / `"false"`); unknown values and missing `val`
   * attributes drop to `undefined`.
   *
   * Reported as `undefined` whenever the source chart has no `<c:title>`
   * element at all — there is no overlay flag to surface in that case.
   */
  titleOverlay?: boolean;
  /**
   * Chart title rotation in whole degrees, pulled from
   * `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title>`.
   * Reflects Excel's "Format Chart Title -> Size & Properties ->
   * Alignment -> Custom angle" knob.
   *
   * The OOXML attribute is in 60000ths of a degree; the reader divides
   * by 60000 (and rounds) to surface a whole-degree value in the
   * `-90..90` band Excel's UI exposes. The OOXML default `0` (and
   * absence of the `<a:bodyPr>` element / `rot` attribute) all collapse
   * to `undefined` so absence and the default round-trip identically
   * through {@link cloneChart}. Out-of-range values clamp to the
   * nearest endpoint of the `-90..90` band; non-numeric tokens drop
   * back to `undefined`.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all — there is no rotation to surface in
   * that case. Mirrors the axis-side {@link ChartAxisInfo.labelRotation}
   * field — same range, same conversion factor — so a parsed value
   * threads straight back into the writer-side
   * {@link SheetChart.titleRotation} without transformation.
   */
  titleRotation?: number;
  /**
   * Chart title font size in points (range `1..400`), pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
   * </a:p></c:rich></c:tx></c:title>`. Reflects Excel's "Format Chart
   * Title -> Font -> Size" knob.
   *
   * The OOXML attribute is in 100ths of a point; the reader converts
   * to points and rounds to the nearest 0.5pt (Excel's UI exposes the
   * same 0.5pt granularity, e.g. `sz="1400"` surfaces as `14`,
   * `sz="1450"` as `14.5`). Out-of-range values (outside the `1..400`
   * band the OOXML `ST_TextFontSize` schema exposes) drop to
   * `undefined` rather than fabricate a value the writer would never
   * emit. Absence of the attribute (or of `<a:defRPr>` / `<a:pPr>` /
   * `<a:p>` / `<c:rich>` / `<c:tx>` / `<c:title>`) likewise collapses
   * to `undefined` so a fresh chart and a chart that pinned an
   * out-of-range size both round-trip to the writer's "skip the size
   * attribute" path.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all, or when the title is a `<c:strRef>`
   * (formula reference) with no `<c:rich>` body — there is no `<a:p>`
   * to host the size in either case. The parsed value threads
   * straight back into the writer-side
   * {@link SheetChart.titleFontSize} without transformation.
   */
  titleFontSize?: number;
  /**
   * Chart title bold flag pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
   * </a:p></c:rich></c:tx></c:title>`. Reflects Excel's "Format Chart
   * Title -> Font -> Bold" toggle.
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `<a:defRPr b="0"/>` round-trip identically through
   * {@link cloneChart} — only an explicit `<a:defRPr b="1"/>` surfaces
   * `true`. The reader accepts the OOXML truthy / falsy spellings
   * (`"1"` / `"true"` / `"0"` / `"false"`); unknown values and missing
   * `b` attributes drop to `undefined`.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all, or when the title is a `<c:strRef>`
   * (formula reference) with no `<c:rich>` body — there is no `<a:p>`
   * to host the flag in either case. The parsed value threads
   * straight back into the writer-side {@link SheetChart.titleBold}
   * without transformation.
   */
  titleBold?: boolean;
  /**
   * Chart title italic flag pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
   * </a:p></c:rich></c:tx></c:title>`. Reflects Excel's "Format Chart
   * Title -> Font -> Italic" toggle.
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `<a:defRPr i="0"/>` round-trip identically through
   * {@link cloneChart} — only an explicit `<a:defRPr i="1"/>` surfaces
   * `true`. The reader accepts the OOXML truthy / falsy spellings
   * (`"1"` / `"true"` / `"0"` / `"false"`); unknown values and missing
   * `i` attributes drop to `undefined`.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all, or when the title is a `<c:strRef>`
   * (formula reference) with no `<c:rich>` body — there is no `<a:p>`
   * to host the flag in either case. The parsed value threads
   * straight back into the writer-side {@link SheetChart.titleItalic}
   * without transformation.
   */
  titleItalic?: boolean;
  /**
   * Chart title font color pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
   * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr>
   * </a:p></c:rich></c:tx></c:title>`. Reflects Excel's "Format Chart
   * Title -> Font -> Font Color" picker.
   *
   * Surfaced as the 6-character uppercase hex string the writer round-
   * trips (`"FF0000"` / `"1070CA"`) — the leading `#` is stripped on
   * read so the value threads straight into the writer-side
   * {@link SheetChart.titleColor} without transformation. Color picks
   * other than the literal sRGB form (`<a:schemeClr>` theme references,
   * `<a:hslClr>`, `<a:sysClr>`, `<a:prstClr>`) collapse to `undefined`
   * — the reader records only the resolvable RGB triple to keep the
   * round-trip lossless against {@link cloneChart} -> {@link writeXlsx}.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all, when the title is a `<c:strRef>`
   * (formula reference) with no `<c:rich>` body, when the
   * `<a:defRPr>` slot has no `<a:solidFill>` child (the title inherits
   * the theme's text color in that case), or when the `<a:srgbClr>`
   * `val` is malformed (wrong length, non-hex characters). There is
   * no `<a:p>` to host the fill in any of those cases.
   */
  titleColor?: string;
  /**
   * Chart title strikethrough flag pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
   * </a:p></c:rich></c:tx></c:title>`. Reflects Excel's "Format Chart
   * Title -> Font -> Strikethrough" toggle.
   *
   * The OOXML attribute is the `ST_TextStrikeType` enum on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
   * three values: `"noStrike"`, `"sngStrike"`, `"dblStrike"`. Only
   * the UI-default `"sngStrike"` (Excel's "Strikethrough" checkbox —
   * single line) surfaces as `true`; `"noStrike"` (the OOXML
   * application default) and absence both collapse to `undefined`,
   * and the non-UI `"dblStrike"` variant likewise collapses to
   * `undefined` rather than surface a value the writer would silently
   * downgrade on round-trip — hucre's writer emits only `"sngStrike"`,
   * so reporting `"dblStrike"` as `true` would round-trip into a
   * lossy single-line replacement.
   *
   * Unknown / malformed `strike` tokens drop to `undefined` rather
   * than fabricate a value the writer would never emit.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all, or when the title is a `<c:strRef>`
   * (formula reference) with no `<c:rich>` body — there is no `<a:p>`
   * to host the flag in either case. The parsed value threads
   * straight back into the writer-side {@link SheetChart.titleStrike}
   * without transformation.
   */
  titleStrike?: boolean;
  /**
   * Chart title underline flag pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
   * </a:p></c:rich></c:tx></c:title>`. Reflects Excel's "Format Chart
   * Title -> Font -> Underline" picker.
   *
   * The OOXML attribute is the `ST_TextUnderlineType` enum on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
   * eighteen values; Excel's UI exposes only `"sng"` (single line —
   * the default underline checkbox) and `"dbl"` (double line). Only
   * the UI-default `"sng"` (Excel's "Underline" checkbox — single
   * line) surfaces as `true`; `"none"` (the OOXML application
   * default) and absence both collapse to `undefined`, and every
   * other token (`"dbl"` and the sixteen exotic variants) likewise
   * collapses to `undefined` rather than surface a value the writer
   * would silently downgrade on round-trip — hucre's writer emits
   * only `"sng"`, so reporting any non-single underline as `true`
   * would round-trip into a lossy single-line replacement.
   *
   * Unknown / malformed `u` tokens drop to `undefined` rather than
   * fabricate a value the writer would never emit.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all, or when the title is a `<c:strRef>`
   * (formula reference) with no `<c:rich>` body — there is no `<a:p>`
   * to host the flag in either case. The parsed value threads
   * straight back into the writer-side {@link SheetChart.titleUnderline}
   * without transformation.
   */
  titleUnderline?: boolean;
  /**
   * Chart title font family / typeface pulled from
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
   * typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx>
   * </c:title>`. Reflects Excel's "Format Chart Title -> Font -> Font"
   * picker. The OOXML `<a:latin typeface=".."/>` element carries the
   * typeface name (`CT_TextFont`, ECMA-376 Part 1, §21.1.2.3.7).
   *
   * Reports the trimmed typeface string when the source chart pinned
   * a non-empty typeface; absence and empty / whitespace-only
   * `typeface` attributes both collapse to `undefined` so absence
   * and `<a:latin typeface=""/>` round-trip identically through
   * {@link cloneChart}. Non-string `typeface` tokens (defensive — the
   * XML parser only ever surfaces strings) likewise drop to
   * `undefined`.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all, or when the title is a `<c:strRef>`
   * (formula reference) with no `<c:rich>` body — there is no `<a:p>`
   * to host the typeface in either case. The parsed value threads
   * straight back into the writer-side
   * {@link SheetChart.titleFontFamily} without transformation.
   */
  titleFontFamily?: string;
  /**
   * Custom chart-title placement pulled from `<c:title><c:layout>
   * <c:manualLayout>...</c:manualLayout></c:layout></c:title>`. Reflects
   * Excel's "Format Chart Title -> Title Options -> Position -> Custom"
   * knob — the `(x, y)` anchor and `(w, h)` size of the title block as
   * fractions of the chart frame in the `0..1` band.
   *
   * Each of {@link ChartManualLayout.x} / {@link ChartManualLayout.y} /
   * {@link ChartManualLayout.w} / {@link ChartManualLayout.h} surfaces
   * the literal `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` value when the
   * source chart pins one; absence / non-numeric / non-finite tokens
   * collapse to `undefined` on the matching field so absence and a
   * malformed token round-trip identically through {@link cloneChart}.
   * The reader accepts both `xMode="edge"` (absolute fraction of the
   * chart frame) and `xMode="factor"` (delta from auto-layout) and
   * surfaces the same shape; the writer normalizes to `"edge"` on emit
   * since that is the form Excel itself emits when the user drags a
   * title to a custom position.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all, or when every coordinate the source
   * pinned drops to `undefined` — there is no meaningful layout to
   * surface in either case. Mirrors the writer-side
   * {@link SheetChart.titleLayout} so a parsed value slots straight
   * into {@link cloneChart} without conversion.
   */
  titleLayout?: ChartManualLayout;
  /**
   * Chart title background fill color pulled from `<c:title><c:spPr>
   * <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></c:spPr>
   * </c:title>`. Reflects Excel's "Format Chart Title -> Fill -> Solid
   * fill -> Color" picker (the same dialog the user reaches by
   * right-clicking the title block). The element sits on `<c:title>`
   * between `<c:overlay>` and `<c:txPr>` per the CT_Title schema
   * sequence (ECMA-376 Part 1, §21.2.2.210).
   *
   * Reports the 6-character uppercase hex string when the source
   * chart pins a literal `<a:srgbClr val="RRGGBB"/>` fill on the
   * title's `<c:spPr>` block. Theme references (`<a:schemeClr>`),
   * non-solid fills (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>` /
   * `<a:blipFill>`), and the OOXML system / preset color forms
   * (`<a:sysClr>` / `<a:hslClr>` / `<a:prstClr>`) all collapse to
   * `undefined` — only the literal RGB triple round-trips losslessly
   * through {@link writeChart}. Malformed `val` tokens (wrong length,
   * non-hex characters) likewise drop to `undefined` rather than
   * fabricate a value the writer would round-trip into a malformed
   * `<a:srgbClr>`.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all, or when the title is a `<c:strRef>`
   * (formula reference) with no `<c:rich>` body — Excel's "Format
   * Title -> Fill" dialog is still authored against `<c:spPr>` even
   * when the text body is a formula reference, so the lookup is on
   * `<c:title>` directly rather than gated on `<c:rich>`. Mirrors the
   * writer-side {@link SheetChart.titleFillColor} so a parsed value
   * slots straight into {@link cloneChart} without conversion.
   */
  titleFillColor?: string;
  /**
   * Chart title border (stroke) solid color pulled from
   * `<c:title><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
   * </a:solidFill></a:ln></c:spPr></c:title>`. Reflects Excel's
   * "Format Chart Title -> Border -> Solid line -> Color" picker.
   *
   * The OOXML `<a:srgbClr>` element carries the literal sRGB color
   * (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32) inside the line's
   * solid fill choice (`CT_LineProperties`, §20.1.2.3.24) which itself
   * sits inside `<c:spPr>` (`CT_ShapeProperties`, §20.1.2.3.13). The
   * `<c:spPr>` slot lives between `<c:overlay>` and `<c:txPr>` /
   * `<c:extLst>` per CT_Title (§21.2.2.210); `<a:ln>` follows the
   * optional `<a:solidFill>` (fill) child inside `<c:spPr>`.
   *
   * The reader surfaces only the literal `<a:srgbClr>` form — absence,
   * non-solid line fills (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>`),
   * and theme-color references (`<a:schemeClr>` / `<a:sysClr>` /
   * `<a:hslClr>` / `<a:prstClr>`) all collapse to `undefined` so a
   * chart that pinned a stroke the writer cannot reproduce on emit
   * drops the field rather than fabricate one Excel would render
   * differently. Malformed `val` tokens (wrong length, non-hex
   * characters, alpha-channel forms, non-string escapes) likewise
   * drop to `undefined`.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all, or when the title is a `<c:strRef>`
   * (formula reference) with no `<c:rich>` body — Excel's "Format
   * Title -> Border" dialog is still authored against `<c:spPr>` even
   * when the text body is a formula reference, so the lookup is on
   * `<c:title>` directly rather than gated on `<c:rich>`. Mirrors the
   * writer-side {@link SheetChart.titleBorderColor} so a parsed value
   * slots straight into {@link cloneChart} without conversion.
   *
   * Independent of {@link titleFillColor} — the two fields surface
   * from the same `<c:spPr>` host but on different children
   * (`<a:solidFill>` for the fill, `<a:ln>` for the stroke), so a
   * caller can read both simultaneously. Mirrors
   * {@link plotAreaBorderColor} — same `<a:ln><a:solidFill><a:srgbClr>`
   * chain on a different host element.
   */
  titleBorderColor?: string;
  /**
   * Chart title border (stroke) thickness in points pulled from the
   * `w` attribute on `<c:title><c:spPr><a:ln w="EMU">`. Reflects
   * Excel's "Format Chart Title -> Border -> Width" spinner. The OOXML
   * `w` attribute stores the stroke width in English Metric Units
   * (1 pt = 12 700 EMU) per `CT_LineProperties` (ECMA-376 Part 1,
   * §20.1.2.3.24); the reader divides by 12 700 and snaps the result
   * to the 0.25 pt grid Excel's UI exposes so a parsed-then-cloned
   * width does not drift across round-trips.
   *
   * Reports the point value clamped to the `0.25..13.5` pt band Excel
   * accepts in the UI when the source chart pinned a finite, positive
   * `w` attribute. Absence (no `<a:ln>` or `<a:ln>` without a `w`
   * attribute), zero, negative, and non-numeric `w` values all collapse
   * to `undefined` so absence and an unrenderable width round-trip
   * identically through {@link cloneChart}. Mirrors the writer-side
   * {@link SheetChart.titleBorderWidth} so a parsed value slots
   * straight into {@link cloneChart} without conversion.
   *
   * Reported as `undefined` whenever the source chart has no
   * `<c:title>` element at all — there is no `<c:spPr>` slot to
   * surface the stroke from in that case. Composes independently with
   * {@link titleBorderColor} — both fields surface from the same
   * `<a:ln>` element but on a different slot (the color child versus
   * the width attribute). Mirrors {@link plotAreaBorderWidth} and
   * {@link legendBorderWidth} — same EMU encoding, same `<a:ln>` host —
   * but lands on `<c:title>`'s own `<c:spPr>` block.
   */
  titleBorderWidth?: number;
  /**
   * Chart title border (stroke) preset dash pattern pulled from the
   * `val` attribute on `<c:title><c:spPr><a:ln><a:prstDash val=".."/>`.
   * Reflects Excel's "Format Chart Title -> Border -> Dash type"
   * picker. Reports the {@link ChartBorderDash} value pinned by the
   * source, or `undefined` when the element is absent / the OOXML
   * default `"solid"` was authored / the value is unrecognized.
   *
   * Mirrors the writer-side {@link SheetChart.titleBorderDash} so a
   * parsed value slots straight into {@link cloneChart} without
   * conversion.
   */
  titleBorderDash?: ChartBorderDash;
  /**
   * Auto-title-deleted flag pulled from `<c:chart><c:autoTitleDeleted
   * val=".."/>`. Reflects Excel's "the user explicitly deleted the
   * auto-generated title" state — single-series charts where Excel
   * normally synthesises a title from the series name leave the flag
   * `false` (the OOXML default) so the auto-title can render; clicking
   * "Delete" on that auto-title flips it to `true` and suppresses the
   * synthesis even though no `<c:title>` element is emitted.
   *
   * The flag is independent of {@link title} — a chart with an explicit
   * `<c:title>` typically pins `false` (the user has not deleted the
   * auto-title because they overrode it with a literal one), while a
   * chart with no `<c:title>` may be `true` (auto-title suppressed)
   * or `false` (auto-title not suppressed; Excel may still synthesise
   * one for a single-series chart).
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * `<c:autoTitleDeleted val="0"/>` round-trip identically through
   * {@link cloneChart} — only an explicit `<c:autoTitleDeleted val="1"/>`
   * surfaces `true`. The reader accepts the OOXML truthy / falsy
   * spellings (`"1"` / `"true"` / `"0"` / `"false"`); unknown values
   * and missing `val` attributes drop to `undefined`.
   */
  autoTitleDeleted?: boolean;
  /**
   * Grouping pulled from the first `<c:barChart>` element, when the
   * chart has one. Surfaces only the stacked variants — the OOXML
   * `"standard"` / `"clustered"` values both round-trip cleanly to
   * the writer's `"clustered"` default, but only the explicit
   * `clustered` value is reported here for symmetry with the writer's
   * {@link SheetChart.barGrouping} field.
   */
  barGrouping?: ChartBarGrouping;
  /**
   * Grouping pulled from the first `<c:lineChart>` element, when the
   * chart has one. Surfaces only `"stacked"` / `"percentStacked"` —
   * the OOXML `"standard"` value is the writer default and collapses
   * to `undefined` here.
   */
  lineGrouping?: ChartLineAreaGrouping;
  /**
   * Grouping pulled from the first `<c:areaChart>` element, when the
   * chart has one. Surfaces only `"stacked"` / `"percentStacked"` —
   * the OOXML `"standard"` value is the writer default and collapses
   * to `undefined` here.
   */
  areaGrouping?: ChartLineAreaGrouping;
  /**
   * Drop-lines flag pulled from the first `<c:lineChart>` /
   * `<c:areaChart>` element's `<c:dropLines/>` child. Reflects
   * Excel's "Add Chart Element -> Lines -> Drop Lines" toggle. The
   * element is bare (it has no `val` attribute) — its mere presence
   * paints the connector lines, so this field surfaces `true` when the
   * element is present and `undefined` when it is absent.
   *
   * The OOXML schema places `<c:dropLines>` exclusively on
   * `<c:lineChart>`, `<c:line3DChart>`, `<c:areaChart>`, and
   * `<c:area3DChart>`. Surfaces `undefined` on every other chart
   * family.
   */
  dropLines?: boolean;
  /**
   * High-low-lines flag pulled from the first `<c:lineChart>`
   * element's `<c:hiLowLines/>` child. Reflects Excel's "Add Chart
   * Element -> Lines -> High-Low Lines" toggle. Like `<c:dropLines>`,
   * the element is bare — its mere presence paints the connectors, so
   * this field surfaces `true` when the element is present and
   * `undefined` when it is absent.
   *
   * The OOXML schema places `<c:hiLowLines>` exclusively on
   * `<c:lineChart>`, `<c:line3DChart>`, and `<c:stockChart>`. Surfaces
   * `undefined` on every other chart family.
   */
  hiLowLines?: boolean;
  /**
   * Series-lines flag pulled from the first `<c:barChart>` /
   * `<c:ofPieChart>` element's `<c:serLines/>` child. Reflects Excel's
   * "Add Chart Element -> Lines -> Series Lines" toggle — connector
   * lines drawn between paired data points across consecutive series in
   * a stacked bar / column chart. Like `<c:dropLines>` /
   * `<c:hiLowLines>`, the element is bare — its mere presence paints
   * the connectors, so this field surfaces `true` when the element is
   * present and `undefined` when it is absent.
   *
   * The OOXML schema places `<c:serLines>` exclusively on
   * `<c:barChart>` and `<c:ofPieChart>`. Surfaces `undefined` on every
   * other chart family.
   */
  serLines?: boolean;
  /**
   * Chart-level data label defaults parsed from the first chart-type
   * element's `<c:dLbls>` block. Series-level overrides on
   * {@link ChartSeriesInfo.dataLabels} take precedence.
   */
  dataLabels?: ChartDataLabelsInfo;
  /**
   * Per-axis metadata. `x` corresponds to the chart's `<c:catAx>`
   * (category axis on bar/column/line/area) or the first `<c:valAx>`
   * on scatter. `y` corresponds to the value axis. Both fields are
   * omitted on charts that have no axes (e.g. pie/doughnut) or when
   * neither axis carries a title.
   */
  axes?: {
    x?: ChartAxisInfo;
    y?: ChartAxisInfo;
  };
  /**
   * Doughnut hole size pulled from the chart's `<c:doughnutChart>
   * <c:holeSize val=".."/>`, expressed as a percentage of the outer
   * radius (1–99). Omitted on non-doughnut charts and on doughnut
   * charts that do not declare the element.
   */
  holeSize?: number;
  /**
   * Bar/column gap width pulled from the first `<c:barChart>` /
   * `<c:bar3DChart>` element's `<c:gapWidth val=".."/>`, expressed as a
   * percentage of the bar width. Range: 0–500. The OOXML default of
   * `150` collapses to `undefined` so absence and the default
   * round-trip identically — symmetric with how the writer's
   * {@link SheetChart.gapWidth} treats the absence of the field.
   * Omitted on non-bar / non-column charts.
   */
  gapWidth?: number;
  /**
   * Bar/column series overlap pulled from the first `<c:barChart>` /
   * `<c:bar3DChart>` element's `<c:overlap val=".."/>`, expressed as a
   * percentage of the bar width. Range: -100..100. The OOXML default of
   * `0` collapses to `undefined` so absence and the default round-trip
   * identically — symmetric with how the writer's
   * {@link SheetChart.overlap} treats the absence of the field.
   * Omitted on non-bar / non-column charts.
   */
  overlap?: number;
  /**
   * Pie / doughnut starting angle in degrees pulled from the first
   * `<c:pieChart>` / `<c:doughnutChart>` element's
   * `<c:firstSliceAng val=".."/>`. Range: 0–360. `0` collapses to
   * `undefined` because it is the OOXML default (first slice at the
   * 12 o'clock position) — the writer's
   * {@link SheetChart.firstSliceAng} treats the absence of the field
   * the same way. Omitted on non-pie / non-doughnut charts.
   */
  firstSliceAng?: number;
  /**
   * How the chart renders missing / blank cells, pulled from
   * `<c:chart><c:dispBlanksAs val=".."/>`. The OOXML default of
   * `"gap"` collapses to `undefined` so absence and the default
   * round-trip identically through {@link cloneChart} — symmetric with
   * the writer's {@link SheetChart.dispBlanksAs} field. Surfaces
   * `"zero"` and `"span"` literally; unknown values are dropped rather
   * than fabricated.
   */
  dispBlanksAs?: ChartDisplayBlanksAs;
  /**
   * Vary-colors-by-point flag pulled from the first chart-type
   * element's `<c:varyColors val=".."/>`. Reflects Excel's
   * per-family default by collapsing matching values to `undefined`:
   *
   *   - On `pie`, `pie3D`, `doughnut`, `ofPie` charts, the OOXML
   *     default is `true` — `<c:varyColors val="1"/>` and absence both
   *     collapse to `undefined`; only an explicit `<c:varyColors val="0"/>`
   *     surfaces `false`.
   *   - On every other chart family the OOXML default is `false` —
   *     `<c:varyColors val="0"/>` and absence both collapse to
   *     `undefined`; only an explicit `<c:varyColors val="1"/>`
   *     surfaces `true`.
   *
   * The asymmetric collapse keeps the parsed shape minimal — a pure
   * round-trip of a stock chart returns no `varyColors` field, while
   * a template that overrides the per-family default surfaces the
   * non-default value so {@link cloneChart} can carry it through.
   * Omitted on chart families that have no `<c:varyColors>` slot
   * (`surface`, `surface3D`, `stock`).
   */
  varyColors?: boolean;
  /**
   * Scatter sub-style pulled from `<c:scatterChart><c:scatterStyle
   * val=".."/></c:scatterChart>`. Reflects which of Excel's six XY
   * scatter presets the chart was authored with — `"none"`, `"line"`,
   * `"lineMarker"`, `"marker"`, `"smooth"`, or `"smoothMarker"`. The
   * OOXML default `"marker"` collapses to `undefined` (Excel's reference
   * serialization actually emits `"lineMarker"` even at the UI default,
   * so the reader does not pin a default of its own — both `"marker"`
   * and `"lineMarker"` surface literally so a clone preserves what the
   * template said).
   *
   * Omitted on every chart family except `scatter`; the OOXML schema
   * places `<c:scatterStyle>` exclusively on `<c:scatterChart>`.
   */
  scatterStyle?: ChartScatterStyle;
  /**
   * Plot-visible-only flag pulled from
   * `<c:chart><c:plotVisOnly val=".."/>`. Reflects Excel's "Hidden and
   * Empty Cells → Show data in hidden rows and columns" toggle (the
   * checkbox is the inverse of this flag — checked means `false`,
   * unchecked means `true`).
   *
   * The OOXML default `true` collapses to `undefined` so absence and
   * the default round-trip identically through {@link cloneChart} —
   * only an explicit `<c:plotVisOnly val="0"/>` surfaces `false`. The
   * reader accepts the OOXML truthy / falsy spellings (`"1"` / `"true"`
   * / `"0"` / `"false"`); unknown values and missing `val` attributes
   * drop to `undefined`.
   */
  plotVisOnly?: boolean;
  /**
   * Show-data-labels-over-max flag pulled from
   * `<c:chart><c:showDLblsOverMax val=".."/>`. Reflects Excel's "Format
   * Axis → Labels → Show data labels for values over maximum scale"
   * checkbox — when the box is unchecked, labels are suppressed for any
   * point whose value exceeds the pinned `<c:max>` axis bound and the
   * field surfaces `false`.
   *
   * The OOXML default `true` collapses to `undefined` so absence and
   * the default round-trip identically through {@link cloneChart} —
   * only an explicit `<c:showDLblsOverMax val="0"/>` surfaces `false`.
   * The reader accepts the OOXML truthy / falsy spellings (`"1"` /
   * `"true"` / `"0"` / `"false"`); unknown values and missing `val`
   * attributes drop to `undefined`. Mirrors the parsing semantics of
   * {@link plotVisOnly}.
   *
   * `<c:showDLblsOverMax>` lives on `<c:chart>` at the tail of CT_Chart
   * (after `<c:dispBlanksAs>` and before `<c:extLst>`). The toggle has
   * no observable effect on a chart whose value axis auto-scales (no
   * point exceeds the auto-computed max); it only matters when the
   * caller pinned a tighter axis ceiling.
   */
  showDLblsOverMax?: boolean;
  /**
   * Rounded-corners flag pulled from
   * `<c:chartSpace><c:roundedCorners val=".."/>`. Reflects Excel's
   * "Format Chart Area → Border → Rounded corners" toggle, which paints
   * the chart frame with rounded edges instead of the default square
   * border.
   *
   * The OOXML default `false` collapses to `undefined` so absence and
   * the default round-trip identically through {@link cloneChart} —
   * only an explicit `<c:roundedCorners val="1"/>` surfaces `true`.
   * The reader accepts the OOXML truthy / falsy spellings (`"1"` /
   * `"true"` / `"0"` / `"false"`); unknown values and missing `val`
   * attributes drop to `undefined`.
   *
   * Note: `<c:roundedCorners>` lives on `<c:chartSpace>`, not inside
   * `<c:chart>` — the toggle styles the outer frame, not the plot area.
   */
  roundedCorners?: boolean;
  /**
   * Up / down bars flag pulled from the first `<c:lineChart>` element's
   * `<c:upDownBars>` child. Reflects Excel's "Add Chart Element ->
   * Up/Down Bars" toggle on a line chart — vertical bars connecting
   * paired series at each category, typically used to visualize open /
   * close differences on a line-style stock chart.
   *
   * Surfaces `true` whenever the element is present (with or without
   * the optional `<c:upBars>` / `<c:downBars>` children — the per-bar
   * styling is not modelled at this layer). The optional
   * `<c:gapWidth>` child is surfaced separately via
   * {@link upDownBarsGapWidth}. Absence collapses to `undefined`. Only
   * line-flavored chart types surface the field; the OOXML schema
   * places `<c:upDownBars>` on `CT_LineChart`, `CT_Line3DChart`, and
   * `CT_StockChart`, so the reader ignores any stray element on bar /
   * column / pie / doughnut / area / scatter chart-type elements.
   */
  upDownBars?: boolean;
  /**
   * Up / down bars gap width pulled from
   * `<c:lineChart><c:upDownBars><c:gapWidth val=".."/></c:upDownBars>
   * </c:lineChart>`. The value is a percentage of the bar width
   * (the OOXML `ST_GapAmount` schema, `0..500`).
   *
   * Surfaces the literal value carried by the file when it falls in
   * the schema band and differs from the OOXML default of `150`. The
   * default and absence both collapse to `undefined` so absence and
   * `<c:gapWidth val="150"/>` round-trip identically through
   * {@link cloneChart}. Out-of-range or non-numeric values are dropped
   * rather than clamped so a corrupt template does not silently surface
   * a value the writer would never emit.
   *
   * Only meaningful when {@link upDownBars} is `true` — the OOXML
   * schema scopes `<c:gapWidth>` exclusively to `<c:upDownBars>`, so
   * the reader only inspects the element when the parent toggle is
   * present.
   */
  upDownBarsGapWidth?: number;
  /**
   * Chart-level marker visibility flag pulled from
   * `<c:lineChart><c:marker val=".."/></c:lineChart>`. Reflects Excel's
   * "Line vs. Line with Markers" chart-type distinction — the flag
   * gates whether per-series markers paint at all on a line chart.
   *
   * Surfaces `false` only when the chart pinned `<c:marker val="0"/>`
   * (the non-default state — the "Line" preset, no per-point dots).
   * The Excel-default `val="1"` and absence both collapse to
   * `undefined` so absence and the default round-trip identically
   * through {@link cloneChart}. The reader accepts the OOXML truthy /
   * falsy spellings (`"1"` / `"true"` / `"0"` / `"false"`); unknown
   * values and missing `val` attributes drop to `undefined`.
   *
   * Only line-flavored chart types surface the field — the OOXML
   * schema places the chart-level `<c:marker>` (CT_Boolean) exclusively
   * on `CT_LineChart`. `CT_Line3DChart` / `CT_StockChart` have no
   * slot for it, and a stray `<c:marker>` on bar / column / pie /
   * doughnut / area / scatter chart-type elements is ignored — the
   * reader only inspects the line-flavored body.
   *
   * Note: this flag is independent of per-series
   * {@link ChartSeriesInfo.marker} — the chart-level toggle gates
   * marker rendering across every series, while the per-series block
   * picks the symbol / size / fill that paints when the gate is open.
   */
  showLineMarkers?: boolean;
  /**
   * Built-in chart style preset pulled from `<c:chartSpace><c:style
   * val=".."/>`. Reflects Excel's "Chart Design -> Chart Styles"
   * gallery — each value picks one of the 48 numbered presets that
   * cycle a colored background, gridline density, border, and label
   * styling across the chart.
   *
   * Surfaces the integer value verbatim when `val` is an integer in
   * the OOXML range (1–48); absence and out-of-range / non-integer
   * values drop to `undefined`. The reader does not pin a default —
   * Excel's reference serialization for a fresh chart emits `<c:style
   * val="2"/>`, but a chart that omits the element renders identically
   * (Excel falls back to its application default). Surfacing only the
   * non-default values keeps the parsed shape minimal and lets a
   * roundtrip of a templated chart preserve its preset while a fresh
   * chart stays unmarked.
   *
   * Note: `<c:style>` lives on `<c:chartSpace>`, not inside
   * `<c:chart>` — the preset styles the outer chart space (frame
   * fill, plot area look, default text font), not just the plot area.
   */
  style?: number;
  /**
   * Editing-locale hint pulled from `<c:chartSpace><c:lang val=".."/>`.
   * The value is an IETF BCP-47 culture name such as `en-US`, `tr-TR`,
   * or `de-DE` — Excel records the editing locale on every authored
   * chart and uses it to drive locale-sensitive defaults (decimal /
   * group separators on unformatted axis ticks, default text font
   * fallback, and the locale recorded for any in-chart text runs).
   *
   * Surfaces the value verbatim when `val` matches the IETF subset
   * Excel emits (`[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*`); absence and
   * malformed tokens drop to `undefined` rather than fabricate a
   * default. Excel's reference serialization for a fresh chart
   * authored on an English locale emits `<c:lang val="en-US"/>`,
   * but the reader does not pin that — only the value the file
   * actually carries surfaces, so the round-trip stays minimal.
   *
   * Note: `<c:lang>` lives on `<c:chartSpace>` (per CT_ChartSpace
   * the element sits between `<c:date1904>` and `<c:roundedCorners>`),
   * not inside `<c:chart>` — the locale governs the entire chart
   * document, not just the plot area.
   */
  lang?: string;
  /**
   * Date-system flag pulled from `<c:chartSpace><c:date1904 val=".."/>`.
   * Mirrors the host workbook's `<workbookPr date1904="1"/>` toggle —
   * `true` signals that date-axis values inside the chart are
   * interpreted under the 1904 base (Excel for Mac's legacy epoch
   * where day 0 falls on 1904-01-01); the OOXML default `false` is
   * the 1900 base.
   *
   * Surfaces `true` only when the chart pinned `<c:date1904 val="1"/>`
   * (the non-default state). The default `val="0"` and absence both
   * collapse to `undefined` so absence and the default round-trip
   * identically through {@link cloneChart}. The reader accepts the
   * OOXML truthy / falsy spellings (`"1"` / `"true"` / `"0"` /
   * `"false"`); unknown values and missing `val` attributes drop to
   * `undefined`.
   *
   * Useful when copying a chart authored on Excel for Mac (or any
   * 1904-based template) into a 1900-based workbook — pinning the
   * flag keeps the chart's date references anchored to the source's
   * epoch instead of silently shifting by 1462 days when the host
   * date system flips. Excel's reference serialization for a fresh
   * chart authored on a 1900-based workbook emits `<c:date1904
   * val="0"/>`, but a chart that omits the element renders identically;
   * surfacing only the non-default value preserves the minimal-shape
   * contract the rest of {@link Chart} follows.
   *
   * Note: `<c:date1904>` lives on `<c:chartSpace>` (per CT_ChartSpace
   * the element sits at the head of the sequence, before `<c:lang>`
   * and `<c:roundedCorners>`), not inside `<c:chart>` — the toggle
   * governs date interpretation across the whole chart document, not
   * just the plot area.
   */
  date1904?: boolean;
  /**
   * Data-table configuration pulled from
   * `<c:plotArea><c:dTable>...</c:dTable></c:plotArea>`. Reflects
   * Excel's "Add Chart Element -> Data Table" toggle, which paints a
   * small table of the underlying series values beneath the plot area.
   *
   * Surfaces a {@link ChartDataTable} object whenever the source chart
   * declares the element. Each of the four boolean children
   * (`<c:showHorzBorder>`, `<c:showVertBorder>`, `<c:showOutline>`,
   * `<c:showKeys>`) round-trips literally — the reader does not collapse
   * any per-field default because every field is required on
   * `CT_DTable` and Excel always emits all four. Absent / unknown
   * `val` attributes drop the matching field to `undefined` rather than
   * fabricate a flag the file did not pin.
   *
   * Surfaces `undefined` when the chart has no `<c:dTable>` element at
   * all. Only chart families with axes (`bar`, `column`, `line`,
   * `area`, `scatter`) carry a data table because the OOXML schema
   * places `<c:dTable>` inside `<c:plotArea>` after the axes — pie /
   * doughnut have no axes and surface `undefined`.
   */
  dataTable?: ChartDataTable;
  /**
   * Chart-space protection pulled from
   * `<c:chartSpace><c:protection>...</c:protection>`. Reflects the
   * chart-level lock Excel honors when the parent worksheet is
   * protected via `<sheetProtection>`.
   *
   * Surfaces a {@link ChartProtection} object whenever the source
   * chart declares the element. Each of the five boolean children
   * (`<c:chartObject>`, `<c:data>`, `<c:formatting>`, `<c:selection>`,
   * `<c:userInterface>`) is independently optional on `CT_Protection`,
   * so the reader only surfaces the flags the file actually pinned.
   * A child that is missing or carries an unknown `val` attribute drops
   * to `undefined` for that field rather than fabricate a value the
   * file did not declare. The element itself is the gating signal — a
   * `<c:protection>` block with no resolvable children surfaces as an
   * empty `{}`, mirroring how `dataTable` handles a malformed `<c:dTable>`.
   *
   * Surfaces `undefined` when the chart has no `<c:protection>` element
   * at all. The element lives on `<c:chartSpace>` (a sibling of
   * `<c:chart>`, between `<c:style>` / `<c:pivotSource>` and
   * `<c:chart>` per CT_ChartSpace), so every chart family — including
   * pie / doughnut — can carry it.
   */
  protection?: ChartProtection;
  /**
   * 3-D view configuration pulled from `<c:chart><c:view3D>` (CT_View3D,
   * ECMA-376 Part 1, §21.2.2.228). Reflects Excel's "3-D Rotation"
   * pane on 3D chart families — the X / Y rotation, height and depth
   * percentages, the right-angle-axes flag, and the perspective
   * foreshortening factor.
   *
   * Surfaces a {@link ChartView3D} object whenever the source chart
   * declares the element. Each of the six children (`<c:rotX>`,
   * `<c:hPercent>`, `<c:rotY>`, `<c:depthPercent>`, `<c:rAngAx>`,
   * `<c:perspective>`) is independently optional on CT_View3D, so the
   * reader only surfaces the fields the file actually pinned. A child
   * that is missing or carries an out-of-range / unparseable `val`
   * attribute drops to `undefined` for that field rather than fabricate
   * a value the file did not declare. The element itself is the gating
   * signal — a `<c:view3D>` block with no resolvable children surfaces
   * as an empty `{}`, mirroring how `dataTable` / `protection` handle
   * a malformed inner block.
   *
   * Surfaces `undefined` when the chart has no `<c:view3D>` element at
   * all. The element lives on `<c:chart>` (between `<c:autoTitleDeleted>`
   * / `<c:pivotFmts>` and `<c:floor>` / `<c:plotArea>` per CT_Chart),
   * so the OOXML schema accepts it on every chart family — though it
   * is only meaningful on 3D families (`bar3D`, `line3D`, `pie3D`,
   * `area3D`, `surface3D`); a stray element on a 2D chart still
   * surfaces here so the round-trip through {@link cloneChart} stays
   * lossless.
   */
  view3D?: ChartView3D;
  /**
   * 3-D floor thickness pulled from
   * `<c:chart><c:floor><c:thickness val="N"/></c:floor>` (the
   * `<c:thickness>` child of `CT_Surface`, ECMA-376 Part 1, §21.2.2.214).
   * Reflects Excel's "Format Floor -> Floor -> Thickness" pin on 3D
   * chart families.
   *
   * Surfaces the integer pinned by the source chart. The OOXML default
   * `0` (and absence of the element) collapses to `undefined` so absence
   * and the default round-trip identically through {@link cloneChart} —
   * only an explicit positive thickness surfaces here. Out-of-range or
   * unparseable values also drop to `undefined` rather than fabricate a
   * value the file did not declare.
   *
   * The element lives on `<c:chart>` between `<c:view3D>` and
   * `<c:sideWall>` / `<c:backWall>` / `<c:plotArea>` per CT_Chart, so
   * the OOXML schema accepts it on every chart family — though it is
   * only meaningful on 3D families (`bar3D`, `line3D`, `pie3D`,
   * `area3D`, `surface3D`); a stray element on a 2D chart still
   * surfaces here so the round-trip through {@link cloneChart} stays
   * lossless.
   */
  floorThickness?: number;
  /**
   * 3-D side-wall thickness pulled from
   * `<c:chart><c:sideWall><c:thickness val="N"/></c:sideWall>` (the
   * `<c:thickness>` child of `CT_Surface`, ECMA-376 Part 1,
   * §21.2.2.187). Reflects Excel's "Format Side Wall -> Side Wall ->
   * Thickness" pin on 3D chart families.
   *
   * Surfaces the integer pinned by the source chart. The OOXML default
   * `0` (and absence of the element) collapses to `undefined` so absence
   * and the default round-trip identically through {@link cloneChart} —
   * only an explicit positive thickness surfaces here. Out-of-range or
   * unparseable values also drop to `undefined` rather than fabricate a
   * value the file did not declare.
   *
   * The element lives on `<c:chart>` between `<c:floor>` and
   * `<c:backWall>` / `<c:plotArea>` per CT_Chart, so the OOXML schema
   * accepts it on every chart family — though it is only meaningful
   * on 3D families (`bar3D`, `line3D`, `pie3D`, `area3D`,
   * `surface3D`); a stray element on a 2D chart still surfaces here
   * so the round-trip through {@link cloneChart} stays lossless.
   */
  sideWallThickness?: number;
  /**
   * 3-D back-wall thickness pulled from
   * `<c:chart><c:backWall><c:thickness val="N"/></c:backWall>` (the
   * `<c:thickness>` child of `CT_Surface`, ECMA-376 Part 1, §21.2.2.214).
   * Reflects Excel's "Format Back Wall -> Back Wall -> Thickness" pin
   * on 3D chart families.
   *
   * Surfaces the integer pinned by the source chart. The OOXML default
   * `0` (and absence of the element) collapses to `undefined` so absence
   * and the default round-trip identically through {@link cloneChart} —
   * only an explicit positive thickness surfaces here. Out-of-range or
   * unparseable values also drop to `undefined` rather than fabricate a
   * value the file did not declare.
   *
   * The element lives on `<c:chart>` between `<c:sideWall>` and
   * `<c:plotArea>` per CT_Chart, so the OOXML schema accepts it on every
   * chart family — though it is only meaningful on 3D families (`bar3D`,
   * `line3D`, `pie3D`, `area3D`, `surface3D`); a stray element on a 2D
   * chart still surfaces here so the round-trip through
   * {@link cloneChart} stays lossless. `<c:floor>` / `<c:sideWall>` /
   * `<c:backWall>` are independent siblings on `<c:chart>`.
   */
  backWallThickness?: number;
}

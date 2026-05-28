// ── Chart Series Extras ────────────────────────────────────────────
// Per-series "extras" that layer on top of `<c:ser>`:
//   - `<c:dPt>`        — per-data-point overrides (CT_DPt, §21.2.2.52)
//   - `<c:trendline>`  — trendline definitions  (CT_Trendline, §21.2.2.211)
//   - `<c:errBars>`    — error bars             (CT_ErrBars, §21.2.2.55)
//   - `<c:bubbleSize>` — bubble size reference  (CT_NumDataSource, §21.2.2.30)
//   - `<c:shape>`      — bar3D shape variant    (ST_Shape, §21.2.3.34)
//
// Reader / writer / clone helpers all live here so the per-feature
// concerns stay co-located. The shape primitives (`normalizeRgbHex`,
// `clampStrokeWidthPt`, `normalizeBorderDash`) are reused from
// `./shape`; the per-series module (`./series`) imports the helpers
// from here when threading the new fields through `buildSeries`.

import type {
  ChartBorderDash,
  ChartDataPoint,
  ChartErrorBarDirection,
  ChartErrorBarType,
  ChartErrorBarValType,
  ChartErrorBars,
  ChartKind,
  ChartShape3D,
  ChartTrendline,
  ChartTrendlineType,
} from "./types"
import type { XmlElement } from "../../xml/parser"
import { xmlElement, xmlSelfClose } from "../../xml/writer"
import {
  EMU_PER_PT,
  STROKE_WIDTH_MAX_PT,
  STROKE_WIDTH_MIN_PT,
  VALID_BORDER_DASHES,
  clampStrokeWidthPt,
  normalizeBorderDash,
  normalizeRgbHex,
} from "./shape"
import { childElements, findChild, formulaText, readBoolAttr } from "./util"
import { buildSeriesMarker, cloneMarker, parseMarker } from "./series"

// ── Constants ─────────────────────────────────────────────────────

/** Recognized 3D shape tokens — mirrors OOXML `ST_Shape`. */
export const VALID_SHAPE_3D: ReadonlySet<ChartShape3D> = new Set([
  "cone",
  "coneToMax",
  "box",
  "cylinder",
  "pyramid",
  "pyramidToMax",
])

/** Recognized trendline type tokens — mirrors OOXML `ST_TrendlineType`. */
export const VALID_TRENDLINE_TYPES: ReadonlySet<ChartTrendlineType> = new Set([
  "linear",
  "log",
  "exp",
  "power",
  "poly",
  "movingAvg",
])

/** Recognized error-bar direction tokens — mirrors OOXML `ST_ErrDir`. */
export const VALID_ERR_DIRECTIONS: ReadonlySet<ChartErrorBarDirection> = new Set(["x", "y"])

/** Recognized error-bar type tokens — mirrors OOXML `ST_ErrBarType`. */
export const VALID_ERR_TYPES: ReadonlySet<ChartErrorBarType> = new Set(["both", "minus", "plus"])

/** Recognized error-bar value-source tokens — mirrors OOXML `ST_ErrValType`. */
export const VALID_ERR_VAL_TYPES: ReadonlySet<ChartErrorBarValType> = new Set([
  "cust",
  "fixedVal",
  "percentage",
  "stdDev",
  "stdErr",
])

/** Polynomial order range Excel's UI exposes (`<c:order>`). */
export const TRENDLINE_ORDER_MIN = 2
export const TRENDLINE_ORDER_MAX = 6

/** Moving-average period range Excel's UI exposes (`<c:period>`). */
export const TRENDLINE_PERIOD_MIN = 2
export const TRENDLINE_PERIOD_MAX = 100

/** Per-point explosion clamp (mirrors series-level explosion). */
const DPT_EXPLOSION_MAX = 400

// ── 3D shape ──────────────────────────────────────────────────────

/**
 * Pull `<c:ser><c:shape val=".."/>` off a bar3D series. Returns
 * `undefined` for absent / malformed / unrecognized tokens.
 */
export function parseShape3D(ser: XmlElement): ChartShape3D | undefined {
  const el = findChild(ser, "shape")
  if (!el) return undefined
  const raw = el.attrs.val
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim() as ChartShape3D
  return VALID_SHAPE_3D.has(trimmed) ? trimmed : undefined
}

/**
 * Validate a 3D-shape value for the writer. Returns the recognized
 * token or `undefined` so the writer can elide the element.
 */
export function normalizeShape3D(value: ChartShape3D | undefined): ChartShape3D | undefined {
  if (typeof value !== "string") return undefined
  return VALID_SHAPE_3D.has(value) ? value : undefined
}

/**
 * Build `<c:shape val=".."/>` for a bar3D series. Returns `undefined`
 * when the value is absent or unrecognized so the caller can skip the
 * element entirely (Excel falls back to `"box"`).
 */
export function buildShape3D(value: ChartShape3D | undefined): string | undefined {
  const normalized = normalizeShape3D(value)
  if (normalized === undefined) return undefined
  return xmlSelfClose("c:shape", { val: normalized })
}

// ── Bubble size ───────────────────────────────────────────────────

/**
 * Pull `<c:ser><c:bubbleSize>` off a bubble series. Returns the raw
 * `<c:f>` formula text or `undefined`.
 */
export function parseBubbleSizeRef(ser: XmlElement): string | undefined {
  return formulaText(findChild(ser, "bubbleSize"))
}

// ── Per-data-point ────────────────────────────────────────────────

/**
 * Pull every `<c:ser><c:dPt>` block off the supplied `<c:ser>`.
 * Returns `undefined` when the array is empty so a series with no
 * overrides matches absence on the parsed shape.
 */
export function parseDataPoints(ser: XmlElement): ChartDataPoint[] | undefined {
  const out: ChartDataPoint[] = []
  for (const child of childElements(ser)) {
    if (child.local !== "dPt") continue
    const idxEl = findChild(child, "idx")
    if (!idxEl) continue
    const rawIdx = idxEl.attrs.val
    if (typeof rawIdx !== "string") continue
    const idx = Number.parseInt(rawIdx, 10)
    if (!Number.isFinite(idx) || idx < 0) continue

    const dp: ChartDataPoint = { idx }

    const explosionEl = findChild(child, "explosion")
    if (explosionEl) {
      const raw = explosionEl.attrs.val
      if (typeof raw === "string") {
        const n = Number.parseFloat(raw)
        if (Number.isFinite(n) && n > 0) {
          const rounded = Math.round(n)
          if (rounded > 0) dp.explosion = rounded
        }
      }
    }

    const bubble3DEl = findChild(child, "bubble3D")
    if (bubble3DEl) {
      const v = readBoolAttr(bubble3DEl)
      if (v === true) dp.bubble3D = true
    }

    const fill = parseDPtFill(child)
    if (fill !== undefined) dp.fillColor = fill

    const borderColor = parseDPtBorderColor(child)
    if (borderColor !== undefined) dp.borderColor = borderColor

    const borderWidth = parseDPtBorderWidth(child)
    if (borderWidth !== undefined) dp.borderWidth = borderWidth

    const borderDash = parseDPtBorderDash(child)
    if (borderDash !== undefined) dp.borderDash = borderDash

    const marker = parseMarker(child)
    if (marker !== undefined) dp.marker = marker

    out.push(dp)
  }
  return out.length > 0 ? out : undefined
}

function parseDPtFill(dPt: XmlElement): string | undefined {
  const spPr = findChild(dPt, "spPr")
  if (!spPr) return undefined
  const fill = findChild(spPr, "solidFill")
  if (!fill) return undefined
  const srgb = findChild(fill, "srgbClr")
  if (!srgb) return undefined
  return normalizeRgbHex(srgb.attrs.val)
}

function parseDPtBorderColor(dPt: XmlElement): string | undefined {
  const spPr = findChild(dPt, "spPr")
  if (!spPr) return undefined
  const ln = findChild(spPr, "ln")
  if (!ln) return undefined
  const fill = findChild(ln, "solidFill")
  if (!fill) return undefined
  const srgb = findChild(fill, "srgbClr")
  if (!srgb) return undefined
  return normalizeRgbHex(srgb.attrs.val)
}

function parseDPtBorderWidth(dPt: XmlElement): number | undefined {
  const spPr = findChild(dPt, "spPr")
  if (!spPr) return undefined
  const ln = findChild(spPr, "ln")
  if (!ln) return undefined
  const wAttr = ln.attrs.w
  if (typeof wAttr !== "string") return undefined
  const emu = Number.parseFloat(wAttr)
  if (!Number.isFinite(emu) || emu <= 0) return undefined
  const pt = Math.round((emu / EMU_PER_PT) * 4) / 4
  if (pt < STROKE_WIDTH_MIN_PT) return STROKE_WIDTH_MIN_PT
  if (pt > STROKE_WIDTH_MAX_PT) return STROKE_WIDTH_MAX_PT
  return pt
}

function parseDPtBorderDash(dPt: XmlElement): ChartBorderDash | undefined {
  const spPr = findChild(dPt, "spPr")
  if (!spPr) return undefined
  const ln = findChild(spPr, "ln")
  if (!ln) return undefined
  const prstDash = findChild(ln, "prstDash")
  if (!prstDash) return undefined
  const raw = prstDash.attrs.val
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim() as ChartBorderDash
  if (!VALID_BORDER_DASHES.has(trimmed)) return undefined
  if (trimmed === "solid") return undefined
  return trimmed
}

/**
 * Build `<c:dPt>` blocks for each entry in {@link ChartDataPoint}[].
 * Returns the joined XML or `undefined` when the array is empty.
 *
 * Accepts the chart kind so pie / doughnut-only fields like
 * `<c:explosion>` can scope-guard themselves: the writer drops
 * explosion on every other family even when the caller pinned it.
 */
export function buildDataPoints(
  dataPoints: ChartDataPoint[] | undefined,
  chartKind: ChartKind | "bar" | "column" | "line" | "pie" | "doughnut" | "scatter" | "area",
): string[] {
  if (!dataPoints || dataPoints.length === 0) return []

  const isPieFamily = chartKind === "pie" || chartKind === "doughnut" || chartKind === "pie3D"

  const out: string[] = []
  for (const dp of dataPoints) {
    if (!Number.isFinite(dp.idx) || dp.idx < 0) continue
    const idx = Math.floor(dp.idx)
    const children: string[] = []

    children.push(xmlSelfClose("c:idx", { val: idx }))

    // `<c:dPt>` schema sequence: idx → invertIfNegative? → marker? →
    // bubble3D? → explosion? → spPr? → pictureOptions? → extLst?
    // We do not model invertIfNegative on per-point yet.

    // Marker (line/scatter/bubble only — emit when present).
    const markerXml = buildSeriesMarker(dp.marker)
    if (markerXml) children.push(markerXml)

    // bubble3D — required on CT_DPt (val="0" by default). Always emit
    // when the chart is bubble or bar3D-ish; for plain charts emit
    // only when the caller asked for it. We emit `val="0"` explicitly
    // because OOXML lists the element as required on CT_DPt.
    children.push(xmlSelfClose("c:bubble3D", { val: dp.bubble3D === true ? 1 : 0 }))

    // explosion — pie family only.
    if (isPieFamily && dp.explosion !== undefined && Number.isFinite(dp.explosion)) {
      const rounded = Math.max(0, Math.min(DPT_EXPLOSION_MAX, Math.round(dp.explosion)))
      if (rounded > 0) children.push(xmlSelfClose("c:explosion", { val: rounded }))
    }

    // spPr — fill / border.
    const fillHex = normalizeRgbHex(dp.fillColor)
    const borderHex = normalizeRgbHex(dp.borderColor)
    const borderWidth = clampStrokeWidthPt(dp.borderWidth)
    const borderDash = normalizeBorderDash(dp.borderDash)
    if (fillHex || borderHex || borderWidth !== undefined || borderDash !== undefined) {
      const spPrChildren: string[] = []
      if (fillHex) {
        spPrChildren.push(
          xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: fillHex })]),
        )
      }
      if (borderHex || borderWidth !== undefined || borderDash !== undefined) {
        const lnAttrs: Record<string, string | number> = {}
        if (borderWidth !== undefined) lnAttrs.w = Math.round(borderWidth * EMU_PER_PT)
        const lnChildren: string[] = []
        if (borderHex) {
          lnChildren.push(
            xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: borderHex })]),
          )
        }
        if (borderDash !== undefined) {
          lnChildren.push(xmlSelfClose("a:prstDash", { val: borderDash }))
        }
        spPrChildren.push(
          lnChildren.length === 0
            ? xmlSelfClose("a:ln", lnAttrs)
            : xmlElement("a:ln", Object.keys(lnAttrs).length > 0 ? lnAttrs : undefined, lnChildren),
        )
      }
      children.push(xmlElement("c:spPr", undefined, spPrChildren))
    }

    out.push(xmlElement("c:dPt", undefined, children))
  }
  return out
}

export function cloneDataPoint(dp: ChartDataPoint): ChartDataPoint {
  const out: ChartDataPoint = { idx: dp.idx }
  if (dp.explosion !== undefined && Number.isFinite(dp.explosion)) out.explosion = dp.explosion
  if (dp.bubble3D === true) out.bubble3D = true
  if (typeof dp.fillColor === "string") out.fillColor = dp.fillColor
  if (typeof dp.borderColor === "string") out.borderColor = dp.borderColor
  if (typeof dp.borderWidth === "number" && Number.isFinite(dp.borderWidth)) {
    out.borderWidth = dp.borderWidth
  }
  if (typeof dp.borderDash === "string") out.borderDash = dp.borderDash
  if (dp.marker) {
    const m = cloneMarker(dp.marker)
    if (m) out.marker = m
  }
  return out
}

export function cloneDataPoints(arr: ChartDataPoint[] | undefined): ChartDataPoint[] | undefined {
  if (!arr || arr.length === 0) return undefined
  const out = arr.map(cloneDataPoint)
  return out.length > 0 ? out : undefined
}

export function resolveDataPoints(
  source: ChartDataPoint[] | undefined,
  override: ChartDataPoint[] | null | undefined,
): ChartDataPoint[] | undefined {
  if (override === undefined) return cloneDataPoints(source)
  if (override === null) return undefined
  return cloneDataPoints(override)
}

// ── Trendlines ────────────────────────────────────────────────────

/**
 * Pull every `<c:ser><c:trendline>` block off the supplied `<c:ser>`.
 * Returns `undefined` when the array is empty.
 */
export function parseTrendlines(ser: XmlElement): ChartTrendline[] | undefined {
  const out: ChartTrendline[] = []
  for (const child of childElements(ser)) {
    if (child.local !== "trendline") continue

    const typeEl = findChild(child, "trendlineType")
    if (!typeEl) continue
    const rawType = typeEl.attrs.val
    if (typeof rawType !== "string") continue
    const type = rawType.trim() as ChartTrendlineType
    if (!VALID_TRENDLINE_TYPES.has(type)) continue

    const t: ChartTrendline = { type }

    const nameEl = findChild(child, "name")
    if (nameEl) {
      // <c:name> is plain text content
      let buf = ""
      for (const c of nameEl.children) {
        if (typeof c === "string") buf += c
      }
      const name = buf.trim()
      if (name.length > 0) t.name = name
    }

    const orderEl = findChild(child, "order")
    if (orderEl) {
      const v = parseFloat(String(orderEl.attrs.val ?? ""))
      if (Number.isFinite(v)) {
        const rounded = Math.round(v)
        if (rounded >= TRENDLINE_ORDER_MIN && rounded <= TRENDLINE_ORDER_MAX) t.order = rounded
      }
    }

    const periodEl = findChild(child, "period")
    if (periodEl) {
      const v = parseFloat(String(periodEl.attrs.val ?? ""))
      if (Number.isFinite(v)) {
        const rounded = Math.round(v)
        if (rounded >= TRENDLINE_PERIOD_MIN && rounded <= TRENDLINE_PERIOD_MAX) {
          t.period = rounded
        }
      }
    }

    const fwdEl = findChild(child, "forward")
    if (fwdEl) {
      const v = parseFloat(String(fwdEl.attrs.val ?? ""))
      if (Number.isFinite(v)) t.forward = v
    }

    const bwdEl = findChild(child, "backward")
    if (bwdEl) {
      const v = parseFloat(String(bwdEl.attrs.val ?? ""))
      if (Number.isFinite(v)) t.backward = v
    }

    const interceptEl = findChild(child, "intercept")
    if (interceptEl) {
      const v = parseFloat(String(interceptEl.attrs.val ?? ""))
      if (Number.isFinite(v)) t.intercept = v
    }

    const dispEqEl = findChild(child, "dispEq")
    if (dispEqEl) {
      const v = readBoolAttr(dispEqEl)
      if (v === true) t.dispEquation = true
    }

    const dispRSqrEl = findChild(child, "dispRSqr")
    if (dispRSqrEl) {
      const v = readBoolAttr(dispRSqrEl)
      if (v === true) t.dispRSquared = true
    }

    // Stroke
    const spPr = findChild(child, "spPr")
    if (spPr) {
      const ln = findChild(spPr, "ln")
      if (ln) {
        const fill = findChild(ln, "solidFill")
        if (fill) {
          const srgb = findChild(fill, "srgbClr")
          if (srgb) {
            const hex = normalizeRgbHex(srgb.attrs.val)
            if (hex) t.lineColor = hex
          }
        }
        const wAttr = ln.attrs.w
        if (typeof wAttr === "string") {
          const emu = Number.parseFloat(wAttr)
          if (Number.isFinite(emu) && emu > 0) {
            const pt = Math.round((emu / EMU_PER_PT) * 4) / 4
            if (pt >= STROKE_WIDTH_MIN_PT && pt <= STROKE_WIDTH_MAX_PT) t.lineWidth = pt
            else if (pt < STROKE_WIDTH_MIN_PT) t.lineWidth = STROKE_WIDTH_MIN_PT
            else t.lineWidth = STROKE_WIDTH_MAX_PT
          }
        }
        const dashEl = findChild(ln, "prstDash")
        if (dashEl) {
          const raw = dashEl.attrs.val
          if (typeof raw === "string") {
            const trimmed = raw.trim() as ChartBorderDash
            if (VALID_BORDER_DASHES.has(trimmed) && trimmed !== "solid") t.lineDash = trimmed
          }
        }
      }
    }

    out.push(t)
  }
  return out.length > 0 ? out : undefined
}

/** Build a `<c:trendline>` block for the writer. */
export function buildTrendline(t: ChartTrendline): string | undefined {
  const type = typeof t.type === "string" && VALID_TRENDLINE_TYPES.has(t.type) ? t.type : undefined
  if (type === undefined) return undefined

  const children: string[] = []

  if (typeof t.name === "string" && t.name.trim().length > 0) {
    children.push(xmlElement("c:name", undefined, escapeXmlText(t.name)))
  }

  // spPr (line color / width / dash)
  const spPr = buildTrendlineSpPr(t)
  if (spPr) children.push(spPr)

  children.push(xmlSelfClose("c:trendlineType", { val: type }))

  if (type === "poly" && typeof t.order === "number" && Number.isFinite(t.order)) {
    const rounded = Math.max(
      TRENDLINE_ORDER_MIN,
      Math.min(TRENDLINE_ORDER_MAX, Math.round(t.order)),
    )
    children.push(xmlSelfClose("c:order", { val: rounded }))
  }

  if (type === "movingAvg" && typeof t.period === "number" && Number.isFinite(t.period)) {
    const rounded = Math.max(
      TRENDLINE_PERIOD_MIN,
      Math.min(TRENDLINE_PERIOD_MAX, Math.round(t.period)),
    )
    children.push(xmlSelfClose("c:period", { val: rounded }))
  }

  if (typeof t.forward === "number" && Number.isFinite(t.forward)) {
    children.push(xmlSelfClose("c:forward", { val: t.forward }))
  }

  if (typeof t.backward === "number" && Number.isFinite(t.backward)) {
    children.push(xmlSelfClose("c:backward", { val: t.backward }))
  }

  if (typeof t.intercept === "number" && Number.isFinite(t.intercept)) {
    children.push(xmlSelfClose("c:intercept", { val: t.intercept }))
  }

  // dispRSqr / dispEq order per CT_Trendline
  if (t.dispRSquared === true) {
    children.push(xmlSelfClose("c:dispRSqr", { val: 1 }))
  }
  if (t.dispEquation === true) {
    children.push(xmlSelfClose("c:dispEq", { val: 1 }))
  }

  return xmlElement("c:trendline", undefined, children)
}

function buildTrendlineSpPr(t: ChartTrendline): string | undefined {
  const lineHex = normalizeRgbHex(t.lineColor)
  const widthPt = clampStrokeWidthPt(t.lineWidth)
  const dash = normalizeBorderDash(t.lineDash)
  if (!lineHex && widthPt === undefined && dash === undefined) return undefined

  const lnAttrs: Record<string, string | number> = {}
  if (widthPt !== undefined) lnAttrs.w = Math.round(widthPt * EMU_PER_PT)
  const lnChildren: string[] = []
  if (lineHex) {
    lnChildren.push(
      xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: lineHex })]),
    )
  }
  if (dash !== undefined) {
    lnChildren.push(xmlSelfClose("a:prstDash", { val: dash }))
  }
  const ln =
    lnChildren.length === 0
      ? xmlSelfClose("a:ln", lnAttrs)
      : xmlElement("a:ln", Object.keys(lnAttrs).length > 0 ? lnAttrs : undefined, lnChildren)
  return xmlElement("c:spPr", undefined, [ln])
}

export function buildTrendlines(trendlines: ChartTrendline[] | undefined): string[] {
  if (!trendlines || trendlines.length === 0) return []
  const out: string[] = []
  for (const t of trendlines) {
    const xml = buildTrendline(t)
    if (xml) out.push(xml)
  }
  return out
}

export function cloneTrendline(t: ChartTrendline): ChartTrendline | undefined {
  if (!VALID_TRENDLINE_TYPES.has(t.type)) return undefined
  const out: ChartTrendline = { type: t.type }
  if (typeof t.name === "string" && t.name.length > 0) out.name = t.name
  if (typeof t.order === "number" && Number.isFinite(t.order)) out.order = t.order
  if (typeof t.period === "number" && Number.isFinite(t.period)) out.period = t.period
  if (typeof t.forward === "number" && Number.isFinite(t.forward)) out.forward = t.forward
  if (typeof t.backward === "number" && Number.isFinite(t.backward)) out.backward = t.backward
  if (typeof t.intercept === "number" && Number.isFinite(t.intercept)) out.intercept = t.intercept
  if (t.dispEquation === true) out.dispEquation = true
  if (t.dispRSquared === true) out.dispRSquared = true
  if (typeof t.lineColor === "string") out.lineColor = t.lineColor
  if (typeof t.lineWidth === "number" && Number.isFinite(t.lineWidth)) out.lineWidth = t.lineWidth
  if (typeof t.lineDash === "string") out.lineDash = t.lineDash
  return out
}

export function cloneTrendlines(arr: ChartTrendline[] | undefined): ChartTrendline[] | undefined {
  if (!arr || arr.length === 0) return undefined
  const out: ChartTrendline[] = []
  for (const t of arr) {
    const c = cloneTrendline(t)
    if (c) out.push(c)
  }
  return out.length > 0 ? out : undefined
}

export function resolveTrendlines(
  source: ChartTrendline[] | undefined,
  override: ChartTrendline[] | null | undefined,
): ChartTrendline[] | undefined {
  if (override === undefined) return cloneTrendlines(source)
  if (override === null) return undefined
  return cloneTrendlines(override)
}

// ── Error bars ────────────────────────────────────────────────────

/**
 * Pull every `<c:ser><c:errBars>` block off the supplied `<c:ser>`.
 * Returns `undefined` when the array is empty.
 */
export function parseErrorBars(ser: XmlElement): ChartErrorBars[] | undefined {
  const out: ChartErrorBars[] = []
  for (const child of childElements(ser)) {
    if (child.local !== "errBars") continue

    const dirEl = findChild(child, "errDir")
    if (!dirEl) continue
    const rawDir = dirEl.attrs.val
    if (typeof rawDir !== "string") continue
    const direction = rawDir.trim() as ChartErrorBarDirection
    if (!VALID_ERR_DIRECTIONS.has(direction)) continue

    const typeEl = findChild(child, "errBarType")
    if (!typeEl) continue
    const rawType = typeEl.attrs.val
    if (typeof rawType !== "string") continue
    const type = rawType.trim() as ChartErrorBarType
    if (!VALID_ERR_TYPES.has(type)) continue

    const valTypeEl = findChild(child, "errValType")
    if (!valTypeEl) continue
    const rawVT = valTypeEl.attrs.val
    if (typeof rawVT !== "string") continue
    const valType = rawVT.trim() as ChartErrorBarValType
    if (!VALID_ERR_VAL_TYPES.has(valType)) continue

    const eb: ChartErrorBars = { direction, type, valType }

    const valEl = findChild(child, "val")
    if (valEl) {
      const v = parseFloat(String(valEl.attrs.val ?? ""))
      if (Number.isFinite(v)) eb.value = v
    }

    const noEndCapEl = findChild(child, "noEndCap")
    if (noEndCapEl) {
      const v = readBoolAttr(noEndCapEl)
      if (v === true) eb.noEndCap = true
    }

    const spPr = findChild(child, "spPr")
    if (spPr) {
      const ln = findChild(spPr, "ln")
      if (ln) {
        const fill = findChild(ln, "solidFill")
        if (fill) {
          const srgb = findChild(fill, "srgbClr")
          if (srgb) {
            const hex = normalizeRgbHex(srgb.attrs.val)
            if (hex) eb.lineColor = hex
          }
        }
        const wAttr = ln.attrs.w
        if (typeof wAttr === "string") {
          const emu = Number.parseFloat(wAttr)
          if (Number.isFinite(emu) && emu > 0) {
            const pt = Math.round((emu / EMU_PER_PT) * 4) / 4
            if (pt < STROKE_WIDTH_MIN_PT) eb.lineWidth = STROKE_WIDTH_MIN_PT
            else if (pt > STROKE_WIDTH_MAX_PT) eb.lineWidth = STROKE_WIDTH_MAX_PT
            else eb.lineWidth = pt
          }
        }
        const dashEl = findChild(ln, "prstDash")
        if (dashEl) {
          const raw = dashEl.attrs.val
          if (typeof raw === "string") {
            const trimmed = raw.trim() as ChartBorderDash
            if (VALID_BORDER_DASHES.has(trimmed) && trimmed !== "solid") eb.lineDash = trimmed
          }
        }
      }
    }

    out.push(eb)
  }
  return out.length > 0 ? out : undefined
}

export function buildErrorBars(eb: ChartErrorBars): string | undefined {
  if (!VALID_ERR_DIRECTIONS.has(eb.direction)) return undefined
  if (!VALID_ERR_TYPES.has(eb.type)) return undefined
  if (!VALID_ERR_VAL_TYPES.has(eb.valType)) return undefined

  const children: string[] = [
    xmlSelfClose("c:errDir", { val: eb.direction }),
    xmlSelfClose("c:errBarType", { val: eb.type }),
    xmlSelfClose("c:errValType", { val: eb.valType }),
  ]

  if (eb.noEndCap === true) {
    children.push(xmlSelfClose("c:noEndCap", { val: 1 }))
  }

  if (eb.valType !== "stdErr" && eb.valType !== "cust") {
    if (typeof eb.value === "number" && Number.isFinite(eb.value)) {
      children.push(xmlSelfClose("c:val", { val: eb.value }))
    }
  }

  // spPr
  const lineHex = normalizeRgbHex(eb.lineColor)
  const widthPt = clampStrokeWidthPt(eb.lineWidth)
  const dash = normalizeBorderDash(eb.lineDash)
  if (lineHex || widthPt !== undefined || dash !== undefined) {
    const lnAttrs: Record<string, string | number> = {}
    if (widthPt !== undefined) lnAttrs.w = Math.round(widthPt * EMU_PER_PT)
    const lnChildren: string[] = []
    if (lineHex) {
      lnChildren.push(
        xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: lineHex })]),
      )
    }
    if (dash !== undefined) {
      lnChildren.push(xmlSelfClose("a:prstDash", { val: dash }))
    }
    const ln =
      lnChildren.length === 0
        ? xmlSelfClose("a:ln", lnAttrs)
        : xmlElement("a:ln", Object.keys(lnAttrs).length > 0 ? lnAttrs : undefined, lnChildren)
    children.push(xmlElement("c:spPr", undefined, [ln]))
  }

  return xmlElement("c:errBars", undefined, children)
}

export function buildAllErrorBars(arr: ChartErrorBars[] | undefined): string[] {
  if (!arr || arr.length === 0) return []
  const out: string[] = []
  for (const eb of arr) {
    const xml = buildErrorBars(eb)
    if (xml) out.push(xml)
  }
  return out
}

export function cloneErrorBars(eb: ChartErrorBars): ChartErrorBars | undefined {
  if (!VALID_ERR_DIRECTIONS.has(eb.direction)) return undefined
  if (!VALID_ERR_TYPES.has(eb.type)) return undefined
  if (!VALID_ERR_VAL_TYPES.has(eb.valType)) return undefined
  const out: ChartErrorBars = {
    direction: eb.direction,
    type: eb.type,
    valType: eb.valType,
  }
  if (typeof eb.value === "number" && Number.isFinite(eb.value)) out.value = eb.value
  if (eb.noEndCap === true) out.noEndCap = true
  if (typeof eb.lineColor === "string") out.lineColor = eb.lineColor
  if (typeof eb.lineWidth === "number" && Number.isFinite(eb.lineWidth))
    out.lineWidth = eb.lineWidth
  if (typeof eb.lineDash === "string") out.lineDash = eb.lineDash
  return out
}

export function cloneAllErrorBars(arr: ChartErrorBars[] | undefined): ChartErrorBars[] | undefined {
  if (!arr || arr.length === 0) return undefined
  const out: ChartErrorBars[] = []
  for (const eb of arr) {
    const c = cloneErrorBars(eb)
    if (c) out.push(c)
  }
  return out.length > 0 ? out : undefined
}

export function resolveErrorBars(
  source: ChartErrorBars[] | undefined,
  override: ChartErrorBars[] | null | undefined,
): ChartErrorBars[] | undefined {
  if (override === undefined) return cloneAllErrorBars(source)
  if (override === null) return undefined
  return cloneAllErrorBars(override)
}

// ── Local helpers ─────────────────────────────────────────────────

function escapeXmlText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
}

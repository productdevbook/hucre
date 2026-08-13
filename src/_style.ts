// ── Cell style copying ───────────────────────────────────────────────
//
// One deep copy of the cell-style tree, for the places that need one:
//
// • `styles-writer` snapshots on registration, so a caller mutating a
//   style after use cannot retroactively restyle earlier cells (#437).
// • `sheet-ops.cloneSheet` copies so a cloned sheet is independent.
// • `xlsx/worksheet` hands a column-level style to a cell as its own
//   object rather than as the column's.
// • `cloneCellStyle` is public, for a caller about to edit one cell.
//
// `styles.resolveStyle` is deliberately **not** one of them. It hands
// back the parsed font / fill / border records themselves, shared by
// every cell whose xf indexes them, so editing one cell's font through a
// resolved style restyles the rest (#439 §P). Copying per cell nearly
// doubles peak memory on a styled read — 407 MB against 787 MB over
// 720,000 styled cells — so the sharing stayed and `cloneCellStyle` is
// the answer instead. `docs/PARITY.md` states it as a contract.
//
// Several walks over the same tree drift the way two cell serializers
// do. This is the one walk.

import type {
  BorderSide,
  BorderStyle,
  CellStyle,
  Color,
  FillStyle,
  FontStyle,
  PatternFill,
} from "./_types"

export function cloneColor(color?: Color): Color | undefined {
  return color === undefined ? undefined : { ...color }
}

export function cloneFont(font: FontStyle): FontStyle {
  const copy: FontStyle = { ...font }
  if (font.color) copy.color = cloneColor(font.color)
  return copy
}

export function cloneFill(fill: FillStyle): FillStyle {
  if (fill.type === "gradient") {
    return {
      ...fill,
      stops: fill.stops.map((stop) => ({ ...stop, color: { ...stop.color } })),
    }
  }
  const copy: PatternFill = { ...fill }
  if (fill.fgColor) copy.fgColor = cloneColor(fill.fgColor)
  if (fill.bgColor) copy.bgColor = cloneColor(fill.bgColor)
  return copy
}

export function cloneBorderSide(side?: BorderSide): BorderSide | undefined {
  if (side === undefined) return undefined
  const copy: BorderSide = { ...side }
  if (side.color) copy.color = cloneColor(side.color)
  return copy
}

export function cloneBorder(border: BorderStyle): BorderStyle {
  const copy: BorderStyle = { ...border }
  for (const side of ["top", "right", "bottom", "left", "diagonal"] as const) {
    if (border[side]) copy[side] = cloneBorderSide(border[side])
  }
  return copy
}

/** Deep copy of a whole cell style. Absent facets stay absent. */
export function cloneCellStyle(style: CellStyle): CellStyle {
  const copy: CellStyle = { ...style }
  if (style.font) copy.font = cloneFont(style.font)
  if (style.fill) copy.fill = cloneFill(style.fill)
  if (style.border) copy.border = cloneBorder(style.border)
  if (style.alignment) copy.alignment = { ...style.alignment }
  if (style.protection) copy.protection = { ...style.protection }
  return copy
}

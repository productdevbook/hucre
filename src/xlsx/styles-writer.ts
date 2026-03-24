// ── Styles XML Writer ────────────────────────────────────────────────
// Generates xl/styles.xml for an XLSX package.

import type {
  CellStyle,
  FontStyle,
  FillStyle,
  BorderStyle,
  AlignmentStyle,
  Color,
  BorderSide,
  CellProtection,
} from "../_types";
import { xmlDocument, xmlElement, xmlSelfClose } from "../xml/writer";

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

// ── Serialization Helpers ──────────────────────────────────────────

function serializeColor(tagName: string, color: Color): string {
  const attrs: Record<string, string | number> = {};
  if (color.rgb !== undefined) {
    // Excel expects ARGB (8 hex chars). If user provides RGB (6), prefix with FF (opaque).
    attrs["rgb"] = color.rgb.length === 6 ? `FF${color.rgb}` : color.rgb;
  }
  if (color.theme !== undefined) {
    attrs["theme"] = color.theme;
  }
  if (color.tint !== undefined) {
    attrs["tint"] = color.tint;
  }
  if (color.indexed !== undefined) {
    attrs["indexed"] = color.indexed;
  }
  return xmlSelfClose(tagName, attrs);
}

function serializeFont(font: FontStyle): string {
  const children: string[] = [];

  if (font.bold) children.push(xmlSelfClose("b"));
  if (font.italic) children.push(xmlSelfClose("i"));
  if (font.strikethrough) children.push(xmlSelfClose("strike"));

  if (font.underline !== undefined && font.underline !== false) {
    if (font.underline === true || font.underline === "single") {
      children.push(xmlSelfClose("u"));
    } else {
      children.push(xmlSelfClose("u", { val: font.underline }));
    }
  }

  if (font.vertAlign) {
    children.push(xmlSelfClose("vertAlign", { val: font.vertAlign }));
  }

  if (font.size !== undefined) {
    children.push(xmlSelfClose("sz", { val: font.size }));
  }

  if (font.color) {
    children.push(serializeColor("color", font.color));
  }

  if (font.name) {
    children.push(xmlSelfClose("name", { val: font.name }));
  }

  if (font.family !== undefined) {
    children.push(xmlSelfClose("family", { val: font.family }));
  }

  if (font.charset !== undefined) {
    children.push(xmlSelfClose("charset", { val: font.charset }));
  }

  if (font.scheme) {
    children.push(xmlSelfClose("scheme", { val: font.scheme }));
  }

  return xmlElement("font", undefined, children);
}

function serializeFill(fill: FillStyle): string {
  if (fill.type === "pattern") {
    const children: string[] = [];

    if (fill.fgColor) {
      children.push(serializeColor("fgColor", fill.fgColor));
    }
    if (fill.bgColor) {
      children.push(serializeColor("bgColor", fill.bgColor));
    }

    const patternFill = xmlElement(
      "patternFill",
      { patternType: fill.pattern },
      children.length > 0 ? children : undefined,
    );

    return xmlElement("fill", undefined, [patternFill]);
  }

  // Gradient fill
  const stops: string[] = [];
  for (const stop of fill.stops) {
    stops.push(
      xmlElement("stop", { position: stop.position }, [serializeColor("color", stop.color)]),
    );
  }

  const attrs: Record<string, string | number> = {};
  if (fill.degree !== undefined) {
    attrs["degree"] = fill.degree;
  }

  const gradientFill = xmlElement("gradientFill", attrs, stops);
  return xmlElement("fill", undefined, [gradientFill]);
}

function serializeBorderSide(tagName: string, side?: BorderSide): string {
  if (!side) {
    return xmlSelfClose(tagName);
  }

  if (side.color) {
    return xmlElement("" + tagName, { style: side.style }, [serializeColor("color", side.color)]);
  }

  return xmlSelfClose(tagName, { style: side.style });
}

function serializeBorder(border: BorderStyle): string {
  const attrs: Record<string, string | boolean> = {};
  if (border.diagonalUp) attrs["diagonalUp"] = true;
  if (border.diagonalDown) attrs["diagonalDown"] = true;

  const children: string[] = [
    serializeBorderSide("left", border.left),
    serializeBorderSide("right", border.right),
    serializeBorderSide("top", border.top),
    serializeBorderSide("bottom", border.bottom),
    serializeBorderSide("diagonal", border.diagonal),
  ];

  return xmlElement("border", Object.keys(attrs).length > 0 ? attrs : undefined, children);
}

// ── Key Generation (for deduplication) ─────────────────────────────

function colorKey(c?: Color): string {
  if (!c) return "";
  const parts: string[] = [];
  if (c.rgb !== undefined) parts.push(`rgb:${c.rgb}`);
  if (c.theme !== undefined) parts.push(`th:${c.theme}`);
  if (c.tint !== undefined) parts.push(`tint:${c.tint}`);
  if (c.indexed !== undefined) parts.push(`idx:${c.indexed}`);
  return parts.join("|");
}

function fontKey(f: FontStyle): string {
  const parts: string[] = [];
  if (f.name) parts.push(`n:${f.name}`);
  if (f.size !== undefined) parts.push(`s:${f.size}`);
  if (f.bold) parts.push("b");
  if (f.italic) parts.push("i");
  if (f.underline !== undefined && f.underline !== false) {
    parts.push(`u:${f.underline}`);
  }
  if (f.strikethrough) parts.push("st");
  if (f.color) parts.push(`c:${colorKey(f.color)}`);
  if (f.vertAlign) parts.push(`va:${f.vertAlign}`);
  if (f.family !== undefined) parts.push(`fam:${f.family}`);
  if (f.charset !== undefined) parts.push(`cs:${f.charset}`);
  if (f.scheme) parts.push(`sch:${f.scheme}`);
  return parts.join(",");
}

function fillKey(f: FillStyle): string {
  if (f.type === "pattern") {
    return `p:${f.pattern}|fg:${colorKey(f.fgColor)}|bg:${colorKey(f.bgColor)}`;
  }
  return `g:${f.degree ?? ""}|${f.stops.map((s) => `${s.position}:${colorKey(s.color)}`).join(";")}`;
}

function borderSideKey(side?: BorderSide): string {
  if (!side) return "";
  return `${side.style}:${colorKey(side.color)}`;
}

function borderKey(b: BorderStyle): string {
  return [
    `l:${borderSideKey(b.left)}`,
    `r:${borderSideKey(b.right)}`,
    `t:${borderSideKey(b.top)}`,
    `b:${borderSideKey(b.bottom)}`,
    `d:${borderSideKey(b.diagonal)}`,
    b.diagonalUp ? "du" : "",
    b.diagonalDown ? "dd" : "",
  ].join("|");
}

// ── Styles Collector ───────────────────────────────────────────────

export interface StylesCollector {
  /** Register a cell style, return its xf index */
  addStyle(style: CellStyle): number;
  /** Register a number format, return its numFmt id */
  addNumFmt(format: string): number;
  /** Register a differential format (for conditional formatting), return its dxfId */
  addDxf(style: CellStyle): number;
  /** Generate the complete styles.xml */
  toXml(): string;
}

interface FontEntry {
  key: string;
  font: FontStyle;
}

interface FillEntry {
  key: string;
  fill: FillStyle;
}

interface BorderEntry {
  key: string;
  border: BorderStyle;
}

interface NumFmtEntry {
  id: number;
  formatCode: string;
}

interface XfEntry {
  key: string;
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  alignment?: AlignmentStyle;
  protection?: CellProtection;
}

export function createStylesCollector(defaultFont?: FontStyle): StylesCollector {
  // ── Defaults ──
  // Default font (Calibri 11)
  const baseFont: FontStyle = {
    name: defaultFont?.name ?? "Calibri",
    size: defaultFont?.size ?? 11,
    ...defaultFont,
  };

  const fonts: FontEntry[] = [{ key: fontKey(baseFont), font: baseFont }];
  const fontMap = new Map<string, number>([[fonts[0].key, 0]]);

  // Excel requires at least 2 fills: "none" and "gray125"
  const fills: FillEntry[] = [
    { key: "p:none|fg:|bg:", fill: { type: "pattern", pattern: "none" } },
    { key: "p:gray125|fg:|bg:", fill: { type: "pattern", pattern: "gray125" } },
  ];
  const fillMap = new Map<string, number>([
    [fills[0].key, 0],
    [fills[1].key, 1],
  ]);

  // Excel requires at least 1 border (empty)
  const emptyBorder: BorderStyle = {};
  const borders: BorderEntry[] = [{ key: borderKey(emptyBorder), border: emptyBorder }];
  const borderMap = new Map<string, number>([[borders[0].key, 0]]);

  // Number formats — custom start at 164
  const numFmts: NumFmtEntry[] = [];
  const numFmtMap = new Map<string, number>();
  let nextNumFmtId = 164;

  // Cell XFs — default xf at index 0
  const defaultXfKey = "nf:0|f:0|fl:0|b:0|a:|p:";
  const xfs: XfEntry[] = [
    {
      key: defaultXfKey,
      numFmtId: 0,
      fontId: 0,
      fillId: 0,
      borderId: 0,
    },
  ];
  const xfMap = new Map<string, number>([[defaultXfKey, 0]]);

  function addFont(font: FontStyle): number {
    const key = fontKey(font);
    const existing = fontMap.get(key);
    if (existing !== undefined) return existing;

    const id = fonts.length;
    fonts.push({ key, font });
    fontMap.set(key, id);
    return id;
  }

  function addFill(fill: FillStyle): number {
    const key = fillKey(fill);
    const existing = fillMap.get(key);
    if (existing !== undefined) return existing;

    const id = fills.length;
    fills.push({ key, fill });
    fillMap.set(key, id);
    return id;
  }

  function addBorder(border: BorderStyle): number {
    const key = borderKey(border);
    const existing = borderMap.get(key);
    if (existing !== undefined) return existing;

    const id = borders.length;
    borders.push({ key, border });
    borderMap.set(key, id);
    return id;
  }

  function addNumFmt(format: string): number {
    const existing = numFmtMap.get(format);
    if (existing !== undefined) return existing;

    const id = nextNumFmtId++;
    numFmts.push({ id, formatCode: format });
    numFmtMap.set(format, id);
    return id;
  }

  function alignmentKey(a?: AlignmentStyle): string {
    if (!a) return "";
    const parts: string[] = [];
    if (a.horizontal) parts.push(`h:${a.horizontal}`);
    if (a.vertical) parts.push(`v:${a.vertical}`);
    if (a.wrapText) parts.push("w");
    if (a.shrinkToFit) parts.push("sf");
    if (a.textRotation !== undefined) parts.push(`r:${a.textRotation}`);
    if (a.indent !== undefined) parts.push(`i:${a.indent}`);
    if (a.readingOrder) parts.push(`ro:${a.readingOrder}`);
    return parts.join(",");
  }

  function protectionKey(p?: CellProtection): string {
    if (!p) return "";
    const parts: string[] = [];
    if (p.locked !== undefined) parts.push(`l:${p.locked}`);
    if (p.hidden !== undefined) parts.push(`h:${p.hidden}`);
    return parts.join(",");
  }

  // ── Differential Formats (dxf) ──
  const dxfs: CellStyle[] = [];
  const dxfMap = new Map<string, number>();

  function dxfKey(style: CellStyle): string {
    const parts: string[] = [];
    if (style.font) parts.push(`f:${fontKey(style.font)}`);
    if (style.fill) parts.push(`fl:${fillKey(style.fill)}`);
    if (style.border) parts.push(`b:${borderKey(style.border)}`);
    if (style.numFmt) parts.push(`nf:${style.numFmt}`);
    if (style.alignment) parts.push(`a:${alignmentKey(style.alignment)}`);
    return parts.join("|");
  }

  function addDxf(style: CellStyle): number {
    const key = dxfKey(style);
    const existing = dxfMap.get(key);
    if (existing !== undefined) return existing;

    const id = dxfs.length;
    dxfs.push(style);
    dxfMap.set(key, id);
    return id;
  }

  function addStyle(style: CellStyle): number {
    const fontId = style.font ? addFont(style.font) : 0;
    const fillId = style.fill ? addFill(style.fill) : 0;
    const borderId = style.border ? addBorder(style.border) : 0;
    const numFmtId = style.numFmt ? addNumFmt(style.numFmt) : 0;

    const key = [
      `nf:${numFmtId}`,
      `f:${fontId}`,
      `fl:${fillId}`,
      `b:${borderId}`,
      `a:${alignmentKey(style.alignment)}`,
      `p:${protectionKey(style.protection)}`,
    ].join("|");

    const existing = xfMap.get(key);
    if (existing !== undefined) return existing;

    const id = xfs.length;
    xfs.push({
      key,
      numFmtId,
      fontId,
      fillId,
      borderId,
      alignment: style.alignment,
      protection: style.protection,
    });
    xfMap.set(key, id);
    return id;
  }

  function serializeAlignment(a: AlignmentStyle): string {
    const attrs: Record<string, string | number | boolean> = {};
    if (a.horizontal) attrs["horizontal"] = a.horizontal;
    if (a.vertical) attrs["vertical"] = a.vertical;
    if (a.wrapText) attrs["wrapText"] = true;
    if (a.shrinkToFit) attrs["shrinkToFit"] = true;
    if (a.textRotation !== undefined) attrs["textRotation"] = a.textRotation;
    if (a.indent !== undefined) attrs["indent"] = a.indent;
    if (a.readingOrder) {
      const roMap = { ltr: 1, rtl: 2, context: 0 } as const;
      attrs["readingOrder"] = roMap[a.readingOrder];
    }
    return xmlSelfClose("alignment", attrs);
  }

  function serializeProtection(p: CellProtection): string {
    const attrs: Record<string, string | number> = {};
    if (p.locked !== undefined) attrs["locked"] = p.locked ? 1 : 0;
    if (p.hidden !== undefined) attrs["hidden"] = p.hidden ? 1 : 0;
    return xmlSelfClose("protection", attrs);
  }

  function toXml(): string {
    const parts: string[] = [];

    // numFmts
    if (numFmts.length > 0) {
      const fmtChildren = numFmts.map((nf) =>
        xmlSelfClose("numFmt", {
          numFmtId: nf.id,
          formatCode: nf.formatCode,
        }),
      );
      parts.push(xmlElement("numFmts", { count: numFmts.length }, fmtChildren));
    }

    // fonts
    const fontChildren = fonts.map((f) => serializeFont(f.font));
    parts.push(xmlElement("fonts", { count: fonts.length }, fontChildren));

    // fills
    const fillChildren = fills.map((f) => serializeFill(f.fill));
    parts.push(xmlElement("fills", { count: fills.length }, fillChildren));

    // borders
    const borderChildren = borders.map((b) => serializeBorder(b.border));
    parts.push(xmlElement("borders", { count: borders.length }, borderChildren));

    // cellStyleXfs (one default entry)
    parts.push(
      xmlElement("cellStyleXfs", { count: 1 }, [
        xmlSelfClose("xf", {
          numFmtId: 0,
          fontId: 0,
          fillId: 0,
          borderId: 0,
        }),
      ]),
    );

    // cellXfs
    const xfChildren = xfs.map((xf) => {
      const attrs: Record<string, string | number | boolean> = {
        numFmtId: xf.numFmtId,
        fontId: xf.fontId,
        fillId: xf.fillId,
        borderId: xf.borderId,
        xfId: 0,
      };

      if (xf.numFmtId !== 0) attrs["applyNumberFormat"] = true;
      if (xf.fontId !== 0) attrs["applyFont"] = true;
      if (xf.fillId !== 0) attrs["applyFill"] = true;
      if (xf.borderId !== 0) attrs["applyBorder"] = true;

      const hasAlignment = xf.alignment !== undefined;
      const hasProtection = xf.protection !== undefined;

      if (hasAlignment) attrs["applyAlignment"] = true;
      if (hasProtection) attrs["applyProtection"] = true;

      if (hasAlignment || hasProtection) {
        const innerChildren: string[] = [];
        if (hasAlignment) innerChildren.push(serializeAlignment(xf.alignment!));
        if (hasProtection) innerChildren.push(serializeProtection(xf.protection!));
        return xmlElement("xf", attrs, innerChildren);
      }

      return xmlSelfClose("xf", attrs);
    });
    parts.push(xmlElement("cellXfs", { count: xfs.length }, xfChildren));

    // dxfs (differential formatting for conditional formatting)
    if (dxfs.length > 0) {
      const dxfChildren = dxfs.map((style) => {
        const children: string[] = [];
        if (style.font) children.push(serializeFont(style.font));
        if (style.numFmt) {
          const numFmtId = addNumFmt(style.numFmt);
          children.push(
            xmlSelfClose("numFmt", {
              numFmtId,
              formatCode: style.numFmt,
            }),
          );
        }
        if (style.fill) children.push(serializeFill(style.fill));
        if (style.border) children.push(serializeBorder(style.border));
        if (style.alignment) children.push(serializeAlignment(style.alignment));
        return xmlElement("dxf", undefined, children);
      });
      parts.push(xmlElement("dxfs", { count: dxfs.length }, dxfChildren));
    }

    return xmlDocument("styleSheet", { xmlns: NS_SPREADSHEET }, parts);
  }

  return {
    addStyle,
    addNumFmt,
    addDxf,
    toXml,
  };
}

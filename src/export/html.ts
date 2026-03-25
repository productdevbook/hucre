import type { Sheet, CellValue, CellStyle, Color, MergeRange } from "../_types";

export interface HtmlExportOptions {
  /** Include inline CSS styles from cell styles. Default: false */
  styles?: boolean;
  /** Add CSS classes for cell types (num, bool, date, null). Default: true */
  classes?: boolean;
  /** Use first row as <thead>. Default: false */
  headerRow?: boolean;
  /** Custom CSS class prefix. Default: "hucre" */
  classPrefix?: string;
  /** Include a minimal <style> block. Default: false */
  includeStyleTag?: boolean;
  /** Table caption text. Adds <caption> as first child of <table>. */
  caption?: string;
  /** ARIA label for the table. Adds aria-label attribute. */
  ariaLabel?: string;
}

/** Escape HTML entities */
function escapeHtml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** Convert a Color to a CSS hex color string */
function colorToCss(color: Color | undefined): string | null {
  if (!color) return null;
  if (color.rgb) {
    // rgb is hex without '#', e.g. "FF0000"
    return `#${color.rgb}`;
  }
  return null;
}

/** Convert a CellStyle to an inline CSS string */
function styleToCss(style: CellStyle): string {
  const parts: string[] = [];

  // Font
  if (style.font) {
    if (style.font.bold) parts.push("font-weight:bold");
    if (style.font.italic) parts.push("font-style:italic");
    if (style.font.underline) parts.push("text-decoration:underline");
    if (style.font.strikethrough) {
      // If already has underline, combine
      const idx = parts.findIndex((p) => p.startsWith("text-decoration:"));
      if (idx >= 0) {
        parts[idx] = "text-decoration:underline line-through";
      } else {
        parts.push("text-decoration:line-through");
      }
    }
    if (style.font.size) parts.push(`font-size:${style.font.size}pt`);
    if (style.font.name) parts.push(`font-family:${style.font.name}`);
    const fontColor = colorToCss(style.font.color);
    if (fontColor) parts.push(`color:${fontColor}`);
  }

  // Fill (background)
  if (style.fill && style.fill.type === "pattern" && style.fill.pattern === "solid") {
    const bgColor = colorToCss(style.fill.fgColor);
    if (bgColor) parts.push(`background-color:${bgColor}`);
  }

  // Alignment
  if (style.alignment) {
    if (style.alignment.horizontal && style.alignment.horizontal !== "general") {
      parts.push(`text-align:${style.alignment.horizontal}`);
    }
    if (style.alignment.vertical) {
      parts.push(`vertical-align:${style.alignment.vertical}`);
    }
    if (style.alignment.wrapText) {
      parts.push("white-space:pre-wrap");
    }
  }

  // Border
  if (style.border) {
    const sides = ["top", "right", "bottom", "left"] as const;
    for (const side of sides) {
      const b = style.border[side];
      if (b) {
        const borderColor = colorToCss(b.color) || "#000000";
        let width = "1px";
        if (
          b.style === "medium" ||
          b.style === "mediumDashed" ||
          b.style === "mediumDashDot" ||
          b.style === "mediumDashDotDot"
        ) {
          width = "2px";
        } else if (b.style === "thick") {
          width = "3px";
        }
        let cssStyle = "solid";
        if (b.style === "dashed" || b.style === "mediumDashed") cssStyle = "dashed";
        else if (b.style === "dotted") cssStyle = "dotted";
        else if (b.style === "double") cssStyle = "double";
        parts.push(`border-${side}:${width} ${cssStyle} ${borderColor}`);
      }
    }
  }

  return parts.join(";");
}

/** Format a cell value as a string for HTML output */
function formatCellValue(value: CellValue): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === "boolean") return String(value);
  if (typeof value === "number") return String(value);
  return escapeHtml(String(value));
}

/** Get the CSS class for a cell value type */
function getCellClass(value: CellValue, prefix: string): string | null {
  if (value === null || value === undefined) return `${prefix}-null`;
  if (value instanceof Date) return `${prefix}-date`;
  if (typeof value === "number") return `${prefix}-num`;
  if (typeof value === "boolean") return `${prefix}-bool`;
  return null; // strings get no special class
}

/**
 * Build a lookup for merged cells.
 * Returns a map from "row,col" to merge info:
 * - For the top-left cell: { colspan, rowspan }
 * - For other cells in the range: { hidden: true }
 */
function buildMergeMap(
  merges: MergeRange[] | undefined,
): Map<string, { colspan?: number; rowspan?: number; hidden?: boolean }> {
  const map = new Map<string, { colspan?: number; rowspan?: number; hidden?: boolean }>();
  if (!merges) return map;

  for (const merge of merges) {
    const colspan = merge.endCol - merge.startCol + 1;
    const rowspan = merge.endRow - merge.startRow + 1;

    // Top-left cell gets colspan/rowspan
    map.set(`${merge.startRow},${merge.startCol}`, {
      colspan: colspan > 1 ? colspan : undefined,
      rowspan: rowspan > 1 ? rowspan : undefined,
    });

    // All other cells in the range are hidden
    for (let r = merge.startRow; r <= merge.endRow; r++) {
      for (let c = merge.startCol; c <= merge.endCol; c++) {
        if (r === merge.startRow && c === merge.startCol) continue;
        map.set(`${r},${c}`, { hidden: true });
      }
    }
  }

  return map;
}

/**
 * Export a sheet as an HTML <table> string.
 */
export function toHtml(sheet: Sheet, options?: HtmlExportOptions): string {
  const opts: Required<Omit<HtmlExportOptions, "caption" | "ariaLabel">> &
    Pick<HtmlExportOptions, "caption" | "ariaLabel"> = {
    styles: options?.styles ?? false,
    classes: options?.classes ?? true,
    headerRow: options?.headerRow ?? false,
    classPrefix: options?.classPrefix ?? "hucre",
    includeStyleTag: options?.includeStyleTag ?? false,
    caption: options?.caption,
    ariaLabel: options?.ariaLabel,
  };

  const prefix = opts.classPrefix;
  const tableAttrs: string[] = [];
  if (opts.includeStyleTag) tableAttrs.push(`class="${prefix}-table"`);
  if (opts.headerRow) tableAttrs.push(`role="table"`);
  if (opts.ariaLabel) tableAttrs.push(`aria-label="${escapeHtml(opts.ariaLabel)}"`);
  const tableAttrStr = tableAttrs.length > 0 ? " " + tableAttrs.join(" ") : "";

  const rows = sheet.rows;
  if (!rows || rows.length === 0) {
    let empty = `<table${tableAttrStr}>`;
    if (opts.caption) empty += `<caption>${escapeHtml(opts.caption)}</caption>`;
    empty += "</table>";
    if (opts.includeStyleTag) {
      return buildStyleTag(prefix) + empty;
    }
    return empty;
  }

  const mergeMap = buildMergeMap(sheet.merges);
  const parts: string[] = [];

  if (opts.includeStyleTag) {
    parts.push(buildStyleTag(prefix));
  }

  parts.push(`<table${tableAttrStr}>`);

  // Caption (first child of table)
  if (opts.caption) {
    parts.push(`<caption>${escapeHtml(opts.caption)}</caption>`);
  }

  const startRow = opts.headerRow ? 1 : 0;

  // Header row
  if (opts.headerRow && rows.length > 0) {
    parts.push("<thead>");
    parts.push("<tr>");
    const row = rows[0];
    for (let c = 0; c < row.length; c++) {
      const mergeInfo = mergeMap.get(`0,${c}`);
      if (mergeInfo?.hidden) continue;

      const value = row[c];
      const attrs = buildCellAttrs(value, 0, c, sheet, opts, mergeInfo);
      parts.push(`<th scope="col"${attrs}>${formatCellValue(value)}</th>`);
    }
    parts.push("</tr>");
    parts.push("</thead>");
  }

  // Body rows
  parts.push("<tbody>");
  for (let r = startRow; r < rows.length; r++) {
    parts.push("<tr>");
    const row = rows[r];
    for (let c = 0; c < row.length; c++) {
      const mergeInfo = mergeMap.get(`${r},${c}`);
      if (mergeInfo?.hidden) continue;

      const value = row[c];
      const attrs = buildCellAttrs(value, r, c, sheet, opts, mergeInfo);
      parts.push(`<td${attrs}>${formatCellValue(value)}</td>`);
    }
    parts.push("</tr>");
  }
  parts.push("</tbody>");

  parts.push("</table>");

  return parts.join("");
}

/** Build HTML attributes for a cell element */
function buildCellAttrs(
  value: CellValue,
  row: number,
  col: number,
  sheet: Sheet,
  opts: Required<Omit<HtmlExportOptions, "caption" | "ariaLabel">> &
    Pick<HtmlExportOptions, "caption" | "ariaLabel">,
  mergeInfo: { colspan?: number; rowspan?: number; hidden?: boolean } | undefined,
): string {
  const attrs: string[] = [];

  // Classes
  if (opts.classes) {
    const cls = getCellClass(value, opts.classPrefix);
    if (cls) attrs.push(`class="${cls}"`);
  }

  // Inline styles
  if (opts.styles) {
    const cell = sheet.cells?.get(`${row},${col}`);
    if (cell?.style) {
      const css = styleToCss(cell.style);
      if (css) attrs.push(`style="${css}"`);
    }
  }

  // Merge attributes
  if (mergeInfo) {
    if (mergeInfo.colspan) attrs.push(`colspan="${mergeInfo.colspan}"`);
    if (mergeInfo.rowspan) attrs.push(`rowspan="${mergeInfo.rowspan}"`);
  }

  return attrs.length > 0 ? " " + attrs.join(" ") : "";
}

/** Build a dark/light responsive <style> block */
function buildStyleTag(prefix: string): string {
  return [
    "<style>",
    `.${prefix}-table{border-collapse:collapse;width:100%;font-family:system-ui,-apple-system,sans-serif;font-size:14px;color:#1a1a1a;background:#fff}`,
    `.${prefix}-table th,.${prefix}-table td{border:1px solid #e0e0e0;padding:6px 10px;text-align:left}`,
    `.${prefix}-table th{background:#f5f5f5;font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:0.03em}`,
    `.${prefix}-table tr:hover td{background:#f8f8f8}`,
    `.${prefix}-num{text-align:right;font-variant-numeric:tabular-nums}`,
    `.${prefix}-bool{text-align:center}`,
    `.${prefix}-null{color:#999;font-style:italic}`,
    `@media(prefers-color-scheme:dark){`,
    `.${prefix}-table{color:#e0e0e0;background:#1a1a1a}`,
    `.${prefix}-table th,.${prefix}-table td{border-color:#333}`,
    `.${prefix}-table th{background:#252525}`,
    `.${prefix}-table tr:hover td{background:#222}`,
    `.${prefix}-null{color:#666}`,
    `}`,
    "</style>",
  ].join("");
}

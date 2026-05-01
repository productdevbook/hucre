// ── Accessibility Helpers ──────────────────────────────────────────
// Audit and helpers for generating WCAG 2.1 AA-compliant spreadsheets.
//
// What screen readers see in a spreadsheet:
//   • The cell pointer reads `<address> <value>` left-to-right from A1.
//   • Drawings (images, charts, text boxes) announce their `descr` attribute
//     on `xdr:cNvPr` — that is the alt text.
//   • The workbook description in docProps/core.xml is announced when the
//     file is opened. Tables (`xl/tables/tableN.xml`) carry an explicit
//     header row that screen readers honor.
//
// The audit covers only what hucre can derive from the in-memory workbook:
// missing alt text, no header row marking, low font-vs-fill contrast,
// merged cells overlapping a header row, blank rows splitting data, and
// missing document-level title/description.

import type {
  A11yCode,
  A11yIssue,
  Cell,
  CellValue,
  Sheet,
  Workbook,
  WorkbookProperties,
} from "./_types";
import { cellRef } from "./xlsx/worksheet-writer";

// ── Public API ─────────────────────────────────────────────────────

export interface AuditOptions {
  /**
   * Minimum contrast ratio for normal-size text. Default: 4.5
   * (WCAG 2.1 AA). Use 7.0 for AAA.
   */
  minContrast?: number;
  /**
   * Skip color contrast checking. Useful when fonts/fills are theme-driven
   * and the resolved colors are not yet known. Default: false.
   */
  skipContrast?: boolean;
  /**
   * Maximum number of cells to inspect for contrast issues. Default: 5000.
   * Avoids walking very large sheets for what is essentially an advisory check.
   */
  contrastSampleLimit?: number;
}

/**
 * Audit a workbook for common WCAG 2.1 AA accessibility issues.
 * Returns a list of findings; an empty array means no issues were detected.
 *
 * @example
 * ```ts
 * import { a11y } from "hucre";
 * const issues = a11y.audit(workbook);
 * for (const issue of issues) console.log(issue.type, issue.message);
 * ```
 */
export function audit(workbook: Workbook, options: AuditOptions = {}): A11yIssue[] {
  const issues: A11yIssue[] = [];
  const minContrast = options.minContrast ?? 4.5;
  const sampleLimit = options.contrastSampleLimit ?? 5000;

  auditWorkbookProperties(workbook.properties, workbook.sheets, issues);

  for (const sheet of workbook.sheets) {
    auditSheet(sheet, issues);
    if (!options.skipContrast) {
      auditSheetContrast(sheet, minContrast, sampleLimit, issues);
    }
  }

  return issues;
}

// ── Side-effecting helper for the writer ───────────────────────────

/**
 * Copy the first non-empty `sheet.a11y.summary` to
 * `workbook.properties.description` when the workbook does not already
 * declare one. Mutates and returns the workbook so writers can simply call
 * `applyA11ySummary(options)` before serialization.
 */
export function applyA11ySummary(workbook: Workbook): Workbook {
  const props = (workbook.properties ?? {}) as WorkbookProperties;
  if (props.description !== undefined && props.description !== "") return workbook;

  for (const sheet of workbook.sheets) {
    const summary = sheet.a11y?.summary;
    if (summary && summary.trim().length > 0) {
      workbook.properties = { ...props, description: summary };
      return workbook;
    }
  }

  return workbook;
}

// ── Color helpers ──────────────────────────────────────────────────

/**
 * WCAG 2.1 relative luminance for an sRGB color. Accepts a 6-digit hex
 * string with or without a leading `#`. Returns a number in [0, 1].
 *
 * Reference: https://www.w3.org/WAI/GL/wiki/Relative_luminance
 */
export function relativeLuminance(hex: string): number {
  const { r, g, b } = parseHex(hex);
  const lin = (c: number): number => {
    const v = c / 255;
    return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
}

/**
 * WCAG 2.1 contrast ratio between two sRGB colors. Returns a value in
 * `[1, 21]`. WCAG 2.1 AA requires `>= 4.5` for normal text and `>= 3.0`
 * for large text (≥ 18pt or ≥ 14pt bold).
 */
export function contrastRatio(fgHex: string, bgHex: string): number {
  const lf = relativeLuminance(fgHex);
  const lb = relativeLuminance(bgHex);
  const [light, dark] = lf > lb ? [lf, lb] : [lb, lf];
  return (light + 0.05) / (dark + 0.05);
}

function parseHex(hex: string): { r: number; g: number; b: number } {
  let h = hex.startsWith("#") ? hex.slice(1) : hex;
  // Excel theme colors sometimes include an alpha prefix ("FFRRGGBB").
  if (h.length === 8) h = h.slice(2);
  if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  if (h.length !== 6 || /[^0-9a-f]/i.test(h)) {
    return { r: 0, g: 0, b: 0 };
  }
  return {
    r: parseInt(h.slice(0, 2), 16),
    g: parseInt(h.slice(2, 4), 16),
    b: parseInt(h.slice(4, 6), 16),
  };
}

// ── Internal audit primitives ──────────────────────────────────────

function push(
  issues: A11yIssue[],
  type: "error" | "warning" | "info",
  code: A11yCode,
  message: string,
  location?: A11yIssue["location"],
): void {
  issues.push({ type, code, message, ...(location ? { location } : {}) });
}

function auditWorkbookProperties(
  props: WorkbookProperties | undefined,
  sheets: Sheet[],
  issues: A11yIssue[],
): void {
  if (!props?.title) {
    push(issues, "info", "no-doc-title", "Workbook has no title in document properties");
  }

  const hasDescription = !!props?.description;
  const hasAnySummary = sheets.some((s) => s.a11y?.summary && s.a11y.summary.trim().length > 0);
  if (!hasDescription && !hasAnySummary) {
    push(
      issues,
      "warning",
      "no-doc-description",
      "Workbook has no description; screen readers cannot announce its purpose",
    );
  }
}

function auditSheet(sheet: Sheet, issues: A11yIssue[]): void {
  const rows = sheet.rows;
  const isEmpty = (rows?.length ?? 0) === 0 && !(sheet.cells && sheet.cells.size > 0);
  if (isEmpty) {
    push(issues, "info", "empty-sheet", `Sheet "${sheet.name}" is empty`, { sheet: sheet.name });
    return;
  }

  // Header row: the audit accepts either an explicit `a11y.headerRow`,
  // an Excel table whose totalsRowShown excludes the header, or simply
  // the user telling us where headers live. Without any signal, we warn
  // because screen readers cannot identify the header otherwise.
  const hasTableHeader = (sheet.tables ?? []).length > 0;
  const headerRow = sheet.a11y?.headerRow;
  if (!hasTableHeader && headerRow === undefined) {
    push(
      issues,
      "warning",
      "no-header-row",
      `Sheet "${sheet.name}" has no header row marked (set sheet.a11y.headerRow or define a table)`,
      { sheet: sheet.name },
    );
  }

  // Header row should not contain merged cells — merged headers
  // confuse screen readers and break the column-by-column read order.
  if (headerRow !== undefined && sheet.merges) {
    for (const merge of sheet.merges) {
      const top = Math.min(merge.startRow, merge.endRow);
      const bottom = Math.max(merge.startRow, merge.endRow);
      if (headerRow >= top && headerRow <= bottom) {
        const ref = `${cellRef(merge.startRow, merge.startCol)}:${cellRef(merge.endRow, merge.endCol)}`;
        push(
          issues,
          "warning",
          "merged-header-row",
          `Sheet "${sheet.name}" has a merged cell overlapping the header row at ${ref}`,
          { sheet: sheet.name, ref },
        );
      }
    }
  }

  // Image alt text. Charts and decorative shapes also live as images,
  // so missing alt text is treated as an error — every screen-reader
  // user will silently skip the cell otherwise.
  if (sheet.images) {
    for (let i = 0; i < sheet.images.length; i++) {
      const img = sheet.images[i];
      if (!img.altText || img.altText.trim().length === 0) {
        const ref = cellRef(img.anchor.from.row, img.anchor.from.col);
        push(
          issues,
          "error",
          "missing-alt-text",
          `Sheet "${sheet.name}" has an image at ${ref} with no alt text`,
          { sheet: sheet.name, ref, image: i },
        );
      }
    }
  }

  if (sheet.textBoxes) {
    for (let t = 0; t < sheet.textBoxes.length; t++) {
      const tb = sheet.textBoxes[t];
      if (!tb.altText || tb.altText.trim().length === 0) {
        const ref = cellRef(tb.anchor.from.row, tb.anchor.from.col);
        push(
          issues,
          "warning",
          "missing-alt-text",
          `Sheet "${sheet.name}" has a text box at ${ref} with no alt text`,
          { sheet: sheet.name, ref, textBox: t },
        );
      }
    }
  }

  // Blank rows in the middle of data — JAWS/NVDA stop reading at the
  // first blank row in a contiguous range and assume the table ended.
  detectBlankRows(sheet, issues);
}

function detectBlankRows(sheet: Sheet, issues: A11yIssue[]): void {
  const rows = sheet.rows;
  if (!rows || rows.length === 0) return;

  let firstNonEmpty = -1;
  let lastNonEmpty = -1;
  for (let r = 0; r < rows.length; r++) {
    if (rowHasContent(rows[r])) {
      if (firstNonEmpty === -1) firstNonEmpty = r;
      lastNonEmpty = r;
    }
  }
  if (firstNonEmpty === -1) return;

  for (let r = firstNonEmpty + 1; r < lastNonEmpty; r++) {
    if (!rowHasContent(rows[r])) {
      const ref = `${r + 1}:${r + 1}`;
      push(
        issues,
        "info",
        "blank-row-in-data",
        `Sheet "${sheet.name}" has a blank row at row ${r + 1}; screen readers may assume the table ended`,
        { sheet: sheet.name, ref },
      );
    }
  }
}

function rowHasContent(row: CellValue[] | undefined): boolean {
  if (!row) return false;
  for (const v of row) {
    if (v !== null && v !== undefined && v !== "") return true;
  }
  return false;
}

function auditSheetContrast(
  sheet: Sheet,
  minContrast: number,
  sampleLimit: number,
  issues: A11yIssue[],
): void {
  if (!sheet.cells || sheet.cells.size === 0) return;

  let inspected = 0;
  for (const [key, cell] of sheet.cells) {
    if (inspected >= sampleLimit) break;
    inspected++;
    if (!hasUserText(cell)) continue;

    const fill = cell.style?.fill;
    if (!fill || fill.type !== "pattern") continue;
    const fg = resolveRgb(cell.style?.font?.color?.rgb);
    const bg = resolveRgb(fill.fgColor?.rgb);
    if (!fg || !bg) continue;

    const ratio = contrastRatio(fg, bg);
    if (ratio < minContrast) {
      const [rowStr, colStr] = key.split(",");
      const row = parseInt(rowStr, 10);
      const col = parseInt(colStr, 10);
      const ref = cellRef(row, col);
      push(
        issues,
        "warning",
        "low-contrast",
        `Sheet "${sheet.name}" cell ${ref} contrast ${ratio.toFixed(2)}:1 below ${minContrast}:1`,
        { sheet: sheet.name, ref },
      );
    }
  }
}

function hasUserText(cell: Partial<Cell>): boolean {
  const v = cell.value;
  return v !== null && v !== undefined && v !== "";
}

function resolveRgb(rgb: string | undefined): string | null {
  if (!rgb) return null;
  // Excel commonly stores 8-digit ARGB; strip the alpha prefix.
  const cleaned = rgb.length === 8 ? rgb.slice(2) : rgb;
  if (cleaned.length !== 6 || /[^0-9a-f]/i.test(cleaned)) return null;
  return cleaned;
}

// ── Style Presets ───────────────────────────────────────────────────
// Reusable CellStyle objects for common report patterns.
// Plain objects — composable with spread operator.

import type { CellStyle, ColumnDef, StylePreset } from "./_types";

// ── Slate ──────────────────────────────────────────────────────────

export const slate: StylePreset = {
  header: {
    font: { bold: true, color: { rgb: "FFFFFF" } },
    fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "374151" } },
    alignment: { horizontal: "center" },
  },
  data: {
    border: { bottom: { style: "thin", color: { rgb: "E5E7EB" } } },
  },
  altData: {
    fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "F9FAFB" } },
    border: { bottom: { style: "thin", color: { rgb: "E5E7EB" } } },
  },
  summary: {
    font: { bold: true },
    border: { top: { style: "medium", color: { rgb: "374151" } } },
  },
};

// ── Ocean ──────────────────────────────────────────────────────────

export const ocean: StylePreset = {
  header: {
    font: { bold: true, color: { rgb: "FFFFFF" } },
    fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "1E40AF" } },
    alignment: { horizontal: "center" },
  },
  data: {
    border: { bottom: { style: "thin", color: { rgb: "DBEAFE" } } },
  },
  altData: {
    fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "EFF6FF" } },
    border: { bottom: { style: "thin", color: { rgb: "DBEAFE" } } },
  },
  summary: {
    font: { bold: true, color: { rgb: "1E40AF" } },
    border: { top: { style: "medium", color: { rgb: "1E40AF" } } },
  },
};

// ── Forest ─────────────────────────────────────────────────────────

export const forest: StylePreset = {
  header: {
    font: { bold: true, color: { rgb: "FFFFFF" } },
    fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "166534" } },
    alignment: { horizontal: "center" },
  },
  data: {
    border: { bottom: { style: "thin", color: { rgb: "D1FAE5" } } },
  },
  altData: {
    fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "F0FDF4" } },
    border: { bottom: { style: "thin", color: { rgb: "D1FAE5" } } },
  },
  summary: {
    font: { bold: true, color: { rgb: "166534" } },
    border: { top: { style: "medium", color: { rgb: "166534" } } },
  },
};

// ── Rose ───────────────────────────────────────────────────────────

export const rose: StylePreset = {
  header: {
    font: { bold: true, color: { rgb: "FFFFFF" } },
    fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "BE123C" } },
    alignment: { horizontal: "center" },
  },
  data: {
    border: { bottom: { style: "thin", color: { rgb: "FECDD3" } } },
  },
  altData: {
    fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFF1F2" } },
    border: { bottom: { style: "thin", color: { rgb: "FECDD3" } } },
  },
  summary: {
    font: { bold: true, color: { rgb: "BE123C" } },
    border: { top: { style: "medium", color: { rgb: "BE123C" } } },
  },
};

// ── Minimal ────────────────────────────────────────────────────────

export const minimal: StylePreset = {
  header: {
    font: { bold: true },
    border: { bottom: { style: "medium", color: { rgb: "000000" } } },
  },
  data: {},
  summary: {
    font: { bold: true },
    border: { top: { style: "thin", color: { rgb: "000000" } } },
  },
};

// ── Helper ─────────────────────────────────────────────────────────

/** Apply a style preset to columns. Sets headerStyle, style, and summary.style where not already set. */
export function applyPreset<T>(columns: ColumnDef<T>[], preset: StylePreset): ColumnDef<T>[] {
  return columns.map((col) => {
    const result: ColumnDef<T> = { ...col };
    if (!result.headerStyle) result.headerStyle = preset.header;
    if (!result.style) result.style = preset.data as CellStyle;
    if (result.summary && !result.summary.style && preset.summary) {
      result.summary = { ...result.summary, style: preset.summary };
    }
    if (result.children) {
      result.children = applyPreset(result.children, preset);
    }
    return result;
  });
}

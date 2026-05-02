// ── Slicer & Timeline Reader ──────────────────────────────────────
// Parses xl/slicers/slicerN.xml, xl/slicerCaches/slicerCacheN.xml,
// xl/timelines/timelineN.xml, and xl/timelineCaches/timelineCacheN.xml
// into structured records that callers can inspect.
//
// Slicers were introduced with Excel 2010 and live in the `x14` extension
// namespace, hence the XML files use the `x14:slicer` element. Timeline
// slicers were added in Excel 2013 under `x15`.
//
// OOXML reference: ECMA-376 Part 1, §18.10 (Slicers) and §18.13 (Timelines).

import type {
  Slicer,
  SlicerCache,
  SlicerCachePivotTable,
  SlicerCacheTableSource,
  Timeline,
  TimelineCache,
} from "../_types";
import { parseXml } from "../xml/parser";
import type { XmlElement } from "../xml/parser";

// ── Slicers (per-sheet) ────────────────────────────────────────────

/**
 * Parse a slicer file (`xl/slicers/slicerN.xml`). One file may declare
 * multiple `<slicer>` elements, so the result is an array.
 */
export function parseSlicers(xml: string): Slicer[] {
  const root = parseXml(xml);
  const out: Slicer[] = [];
  for (const child of childElements(root)) {
    if (child.local !== "slicer") continue;
    const name = child.attrs.name;
    const cache = child.attrs.cache;
    if (!name || !cache) continue;
    const slicer: Slicer = { name, cache };
    if (child.attrs.caption) slicer.caption = child.attrs.caption;
    if (child.attrs.columnCount !== undefined) {
      const n = parseIntSafe(child.attrs.columnCount, NaN);
      if (!Number.isNaN(n)) slicer.columnCount = n;
    }
    if (child.attrs.style) slicer.style = child.attrs.style;
    if (child.attrs.sortOrder) slicer.sortOrder = child.attrs.sortOrder;
    if (child.attrs.rowHeight !== undefined) {
      const n = parseIntSafe(child.attrs.rowHeight, NaN);
      if (!Number.isNaN(n)) slicer.rowHeight = n;
    }
    out.push(slicer);
  }
  return out;
}

// ── Slicer Caches (workbook-level) ─────────────────────────────────

/**
 * Parse a slicer cache file (`xl/slicerCaches/slicerCacheN.xml`).
 *
 * Structure:
 *   <slicerCacheDefinition name="..." sourceName="...">
 *     <pivotTables>
 *       <pivotTable tabId="0" name="PivotTable1"/>
 *     </pivotTables>
 *     <data>
 *       <tabular pivotCacheId="..."/>
 *     </data>
 *     <extLst>
 *       <ext>
 *         <x15:tableSlicerCache tableId="1" column="..."/>
 *       </ext>
 *     </extLst>
 *   </slicerCacheDefinition>
 */
export function parseSlicerCache(xml: string): SlicerCache | undefined {
  const root = parseXml(xml);
  const def =
    root.local === "slicerCacheDefinition" ? root : findChild(root, "slicerCacheDefinition");
  if (!def) return undefined;
  const name = def.attrs.name;
  if (!name) return undefined;

  const cache: SlicerCache = { name };
  if (def.attrs.sourceName) cache.sourceName = def.attrs.sourceName;

  const pivotTables = parsePivotTables(def);
  if (pivotTables.length > 0) cache.pivotTables = pivotTables;

  const tableSource = parseTableSlicerCache(def);
  if (tableSource) cache.tableSource = tableSource;

  return cache;
}

// ── Timelines (per-sheet) ──────────────────────────────────────────

/**
 * Parse a timeline file (`xl/timelines/timelineN.xml`). One file may
 * declare multiple `<timeline>` elements.
 */
export function parseTimelines(xml: string): Timeline[] {
  const root = parseXml(xml);
  const out: Timeline[] = [];
  for (const child of childElements(root)) {
    if (child.local !== "timeline") continue;
    const name = child.attrs.name;
    const cache = child.attrs.cache;
    if (!name || !cache) continue;
    const tl: Timeline = { name, cache };
    if (child.attrs.caption) tl.caption = child.attrs.caption;
    if (child.attrs.style) tl.style = child.attrs.style;
    if (child.attrs.level) tl.level = child.attrs.level;
    if (child.attrs.showHeader !== undefined) tl.showHeader = parseBool(child.attrs.showHeader);
    if (child.attrs.showSelectionLabel !== undefined)
      tl.showSelectionLabel = parseBool(child.attrs.showSelectionLabel);
    if (child.attrs.showTimeLevel !== undefined)
      tl.showTimeLevel = parseBool(child.attrs.showTimeLevel);
    if (child.attrs.showHorizontalScrollbar !== undefined)
      tl.showHorizontalScrollbar = parseBool(child.attrs.showHorizontalScrollbar);
    out.push(tl);
  }
  return out;
}

// ── Timeline Caches (workbook-level) ───────────────────────────────

/**
 * Parse a timeline cache file (`xl/timelineCaches/timelineCacheN.xml`).
 */
export function parseTimelineCache(xml: string): TimelineCache | undefined {
  const root = parseXml(xml);
  const def =
    root.local === "timelineCacheDefinition" ? root : findChild(root, "timelineCacheDefinition");
  if (!def) return undefined;
  const name = def.attrs.name;
  if (!name) return undefined;
  const cache: TimelineCache = { name };
  if (def.attrs.sourceName) cache.sourceName = def.attrs.sourceName;

  const pivotTables = parsePivotTables(def);
  if (pivotTables.length > 0) cache.pivotTables = pivotTables;

  return cache;
}

// ── Internals ─────────────────────────────────────────────────────

function parsePivotTables(def: XmlElement): SlicerCachePivotTable[] {
  const pivotTables = findChild(def, "pivotTables");
  if (!pivotTables) return [];
  const out: SlicerCachePivotTable[] = [];
  for (const child of childElements(pivotTables)) {
    if (child.local !== "pivotTable") continue;
    const name = child.attrs.name;
    if (!name) continue;
    const tabId = parseIntSafe(child.attrs.tabId, NaN);
    if (Number.isNaN(tabId)) continue;
    out.push({ tabId, name });
  }
  return out;
}

function parseTableSlicerCache(def: XmlElement): SlicerCacheTableSource | undefined {
  const extLst = findChild(def, "extLst");
  if (!extLst) return undefined;
  for (const ext of childElements(extLst)) {
    if (ext.local !== "ext") continue;
    const tableSlicer = findChild(ext, "tableSlicerCache");
    if (!tableSlicer) continue;
    const name = tableSlicer.attrs.name ?? tableSlicer.attrs.tableId;
    if (!name) continue;
    const src: SlicerCacheTableSource = { name };
    if (tableSlicer.attrs.column) src.column = tableSlicer.attrs.column;
    return src;
  }
  return undefined;
}

function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c;
  }
  return undefined;
}

function childElements(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of el.children) {
    if (typeof c !== "string") out.push(c);
  }
  return out;
}

function parseIntSafe(s: string | undefined, fallback: number): number {
  if (s === undefined) return fallback;
  const n = parseInt(s, 10);
  return Number.isNaN(n) ? fallback : n;
}

function parseBool(s: string): boolean {
  return s === "1" || s === "true";
}

// ── Pivot Table Reader ────────────────────────────────────────────
// Parses xl/pivotTables/pivotTable{N}.xml and
// xl/pivotCache/pivotCacheDefinition{N}.xml into structured records.
//
// OOXML reference: ECMA-376 Part 1, §18.10 (PivotTables) and §18.11
// (PivotCache).

import type {
  PivotCache,
  PivotDataFieldFunction,
  PivotField,
  PivotFieldAxis,
  PivotTable,
} from "../_types";
import { parseXml } from "../xml/parser";
import type { XmlElement } from "../xml/parser";

// ── Pivot Cache Definition (workbook-level) ────────────────────────

/**
 * Parse a pivot cache definition file
 * (`xl/pivotCache/pivotCacheDefinition{N}.xml`).
 *
 * The OOXML root element is `<pivotCacheDefinition>`. The two pieces
 * we surface are the source range — read off `<cacheSource>` and its
 * `<worksheetSource>` child — and the cache field names, which line up
 * 1:1 with the field indexes used by `<pivotField>` entries on the
 * pivot table side.
 */
export function parsePivotCacheDefinition(xml: string): PivotCache | undefined {
  const root = parseXml(xml);
  const def =
    root.local === "pivotCacheDefinition" ? root : findChild(root, "pivotCacheDefinition");
  if (!def) return undefined;

  // cacheId is not stored on the definition itself — it lives in
  // workbook.xml. Default to 0; the reader fills in the real value
  // from the workbook's <pivotCaches> block.
  const cache: PivotCache = { cacheId: 0, fieldNames: [] };

  // Source: <cacheSource type="..."><worksheetSource ref="..." sheet="..."/></cacheSource>
  const source = findChild(def, "cacheSource");
  if (source) {
    const t = source.attrs.type;
    if (t === "worksheet" || t === "external" || t === "consolidation" || t === "scenario") {
      cache.sourceType = t;
    }
    const worksheetSource = findChild(source, "worksheetSource");
    if (worksheetSource) {
      // `ref` is the cell range. `name` (a defined name) is the
      // alternative — both stash into `sourceRef`. `sheet` is optional.
      if (worksheetSource.attrs.ref) cache.sourceRef = worksheetSource.attrs.ref;
      else if (worksheetSource.attrs.name) cache.sourceRef = worksheetSource.attrs.name;
      if (worksheetSource.attrs.sheet) cache.sourceSheet = worksheetSource.attrs.sheet;
    }
  }

  // Cache fields: <cacheFields count="N"><cacheField name="..." ...>...</cacheField></cacheFields>
  const cacheFields = findChild(def, "cacheFields");
  if (cacheFields) {
    for (const child of childElements(cacheFields)) {
      if (child.local !== "cacheField") continue;
      const name = child.attrs.name ?? "";
      cache.fieldNames.push(name);
    }
  }

  return cache;
}

// ── Pivot Table Definition (per-sheet) ─────────────────────────────

/**
 * Parse a pivot table definition file
 * (`xl/pivotTables/pivotTable{N}.xml`).
 *
 * The body is dense — we keep just the layout-relevant attributes so
 * roundtrip can carry them forward unmodified, and surface the field
 * placement (row / col / page / data) that callers most often want to
 * inspect.
 */
export function parsePivotTable(xml: string): PivotTable | undefined {
  const root = parseXml(xml);
  const def =
    root.local === "pivotTableDefinition" ? root : findChild(root, "pivotTableDefinition");
  if (!def) return undefined;

  const name = def.attrs.name;
  if (!name) return undefined;

  const cacheIdRaw = parseIntSafe(def.attrs.cacheId, NaN);

  const table: PivotTable = {
    name,
    // The per-table cacheId attribute matches the workbook's
    // <pivotCache cacheId="..."> entry. Fall back to 0 when missing
    // rather than throwing — Excel tolerates that.
    cacheId: Number.isNaN(cacheIdRaw) ? 0 : cacheIdRaw,
    location: "",
    fields: [],
  };

  // Location block: <location ref="A3:D20" firstHeaderRow="0" firstDataRow="1" firstDataCol="0"/>
  const location = findChild(def, "location");
  if (location) {
    if (location.attrs.ref) table.location = location.attrs.ref;
    const fh = parseIntSafe(location.attrs.firstHeaderRow, NaN);
    if (!Number.isNaN(fh)) table.firstHeaderRow = fh;
    const fdr = parseIntSafe(location.attrs.firstDataRow, NaN);
    if (!Number.isNaN(fdr)) table.firstDataRow = fdr;
    const fdc = parseIntSafe(location.attrs.firstDataCol, NaN);
    if (!Number.isNaN(fdc)) table.firstDataCol = fdc;
    const rpc = parseIntSafe(location.attrs.rowPageCount, NaN);
    if (!Number.isNaN(rpc)) table.rowPageCount = rpc;
    const cpc = parseIntSafe(location.attrs.colPageCount, NaN);
    if (!Number.isNaN(cpc)) table.colPageCount = cpc;
  }

  // Field declarations: <pivotFields><pivotField axis="..."/></pivotFields>
  // The position in this list is the field's index everywhere else.
  const pivotFields = findChild(def, "pivotFields");
  const fieldDefs: Array<{ axis: PivotFieldAxis; raw: XmlElement }> = [];
  if (pivotFields) {
    for (const child of childElements(pivotFields)) {
      if (child.local !== "pivotField") continue;
      const axisAttr = child.attrs.axis;
      const dataFlag = child.attrs.dataField === "1" || child.attrs.dataField === "true";
      let axis: PivotFieldAxis = "hidden";
      if (axisAttr === "axisRow") axis = "row";
      else if (axisAttr === "axisCol") axis = "col";
      else if (axisAttr === "axisPage") axis = "page";
      else if (axisAttr === "axisValues" || dataFlag) axis = "data";
      fieldDefs.push({ axis, raw: child });
    }
  }

  // Walk the dataFields block to recover per-data-field aggregation +
  // display name overrides. Each <dataField fld="N" subtotal="sum"/>
  // points back into the pivotField at index `fld`.
  const dataFields = findChild(def, "dataFields");
  const dataFieldOverrides = new Map<
    number,
    { name?: string; subtotal?: PivotDataFieldFunction }
  >();
  if (dataFields) {
    for (const child of childElements(dataFields)) {
      if (child.local !== "dataField") continue;
      const fldRaw = parseIntSafe(child.attrs.fld, NaN);
      if (Number.isNaN(fldRaw)) continue;
      const entry: { name?: string; subtotal?: PivotDataFieldFunction } = {};
      if (child.attrs.name) entry.name = child.attrs.name;
      const subtotal = mapAggregateFunction(child.attrs.subtotal);
      if (subtotal) entry.subtotal = subtotal;
      // dataFields may legally repeat the same fld — last write wins
      // here so the surfaced override matches what Excel itself shows.
      dataFieldOverrides.set(fldRaw, entry);
    }
  }

  // Need the cache field names to populate PivotField.name. The reader
  // doesn't have those at this layer (the cache definition is parsed
  // separately), so emit synthetic names based on the field index. The
  // caller (xlsx/reader.ts) overlays the real names afterwards.
  for (let i = 0; i < fieldDefs.length; i++) {
    const { axis } = fieldDefs[i];
    const f: PivotField = { name: `field${i + 1}`, axis };
    if (axis === "data") {
      const ov = dataFieldOverrides.get(i);
      if (ov?.subtotal) f.function = ov.subtotal;
      else f.function = "sum"; // OOXML default
      if (ov?.name) f.displayName = ov.name;
    }
    table.fields.push(f);
  }

  // Style info
  const styleInfo = findChild(def, "pivotTableStyleInfo");
  if (styleInfo?.attrs.name) table.styleName = styleInfo.attrs.name;

  if (def.attrs.dataCaption) table.dataCaption = def.attrs.dataCaption;

  return table;
}

/**
 * Overlay cache field names onto a PivotTable's synthetic field names.
 * Mutates `table.fields[i].name` for indexes the cache covers; out-of-
 * range entries keep their `fieldN` placeholder.
 */
export function attachPivotCacheFields(table: PivotTable, cache: PivotCache): void {
  for (let i = 0; i < table.fields.length; i++) {
    const cacheName = cache.fieldNames[i];
    if (cacheName) table.fields[i].name = cacheName;
  }
}

// ── Internals ─────────────────────────────────────────────────────

/**
 * Map the OOXML `subtotal` enum (e.g. `"sum"`, `"countA"`) onto the
 * narrower `PivotDataFieldFunction` we expose. Returns `undefined`
 * for values we don't recognise — caller falls back to the spec default.
 */
function mapAggregateFunction(s: string | undefined): PivotDataFieldFunction | undefined {
  if (!s) return undefined;
  switch (s) {
    case "sum":
    case "count":
    case "average":
    case "max":
    case "min":
    case "product":
    case "countNums":
    case "stdDev":
    case "stdDevp":
    case "var":
    case "varp":
      return s;
    case "countA":
      // OOXML's countA collapses to count for our purposes.
      return "count";
    default:
      return undefined;
  }
}

function childElements(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of el.children) {
    if (typeof c !== "string") out.push(c);
  }
  return out;
}

function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c;
  }
  return undefined;
}

function parseIntSafe(s: string | undefined, fallback: number): number {
  if (s === undefined) return fallback;
  const n = parseInt(s, 10);
  return Number.isNaN(n) ? fallback : n;
}

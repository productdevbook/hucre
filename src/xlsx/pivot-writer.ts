// ── Pivot Table Writer ────────────────────────────────────────────────
// Generates the four OOXML parts that back a pivot table:
//   xl/pivotCache/pivotCacheDefinition{N}.xml — field declarations + shared items
//   xl/pivotCache/pivotCacheRecords{N}.xml    — cached source rows
//   xl/pivotCache/_rels/pivotCacheDefinition{N}.xml.rels — links cache → records
//   xl/pivotTables/pivotTable{N}.xml          — pivot layout
//   xl/pivotTables/_rels/pivotTable{N}.xml.rels — links pivot → cache definition
//
// Phase 1 of issue #159. The aim is a *structurally valid* pivot table
// that Excel can populate on first open via the recompute that
// `<calcPr fullCalcOnLoad="1"/>` triggers — the writer does not
// pre-compute row totals, value cells, or expanded item layouts.
//
// OOXML reference: ECMA-376 Part 1, §18.10 (PivotTables) and §18.11
// (PivotCache).

import type { CellValue, PivotDataFieldFunction, WritePivotTable } from "../_types";
import { xmlDocument, xmlElement, xmlSelfClose } from "../xml/writer";

// ── Namespaces ────────────────────────────────────────────────────────

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";

const REL_PIVOT_CACHE_RECORDS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
const REL_PIVOT_CACHE_DEFINITION =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";

// ── Public API ────────────────────────────────────────────────────────

/**
 * A pivot data source resolved against the workbook. The header row
 * gives the field names; data rows feed both the cache records and the
 * shared-items table on each `<cacheField>`.
 */
export interface ResolvedPivotSource {
  /** Source sheet name as it appears in `xl/workbook.xml`. */
  sheetName: string;
  /** Field names (in declaration order). */
  fieldNames: string[];
  /** Data rows aligned 1:1 with `fieldNames`. */
  dataRows: CellValue[][];
  /**
   * The OOXML `<worksheetSource ref>` value, e.g. `"A1:C100"`. Already
   * resolved against the source sheet's data length when the user did
   * not supply a range.
   */
  ref: string;
}

export interface PivotWriteResult {
  /** Body of `xl/pivotCache/pivotCacheDefinition{N}.xml`. */
  cacheDefinitionXml: string;
  /** Body of `xl/pivotCache/_rels/pivotCacheDefinition{N}.xml.rels`. */
  cacheDefinitionRels: string;
  /** Body of `xl/pivotCache/pivotCacheRecords{N}.xml`. */
  cacheRecordsXml: string;
  /** Body of `xl/pivotTables/pivotTable{N}.xml`. */
  pivotTableXml: string;
  /** Body of `xl/pivotTables/_rels/pivotTable{N}.xml.rels`. */
  pivotTableRels: string;
}

/**
 * Generate the cache + table OOXML for a single pivot table.
 *
 * @param pivot - User-supplied pivot definition.
 * @param source - Resolved source data (header row + data rows).
 * @param cacheId - Workbook-level cacheId. Mirrors the value emitted in
 *                  `<workbook><pivotCaches><pivotCache cacheId="..."/>`.
 */
export function writePivotTable(
  pivot: WritePivotTable,
  source: ResolvedPivotSource,
  cacheId: number,
): PivotWriteResult {
  const fieldsMeta = buildFieldMetadata(pivot, source);

  const cacheDefinitionXml = buildCacheDefinition(source, fieldsMeta);
  const cacheDefinitionRels = buildCacheDefinitionRels();
  const cacheRecordsXml = buildCacheRecords(source, fieldsMeta);
  const pivotTableXml = buildPivotTable(pivot, fieldsMeta, cacheId);
  const pivotTableRels = buildPivotTableRels();

  return {
    cacheDefinitionXml,
    cacheDefinitionRels,
    cacheRecordsXml,
    pivotTableXml,
    pivotTableRels,
  };
}

/**
 * Resolve a {@link WritePivotTable} against the workbook's source sheet.
 *
 * Throws when the source sheet is missing, has fewer than two rows
 * (header + at least one data row), or when a referenced field is not
 * present in the header row.
 */
export function resolvePivotSource(
  pivot: WritePivotTable,
  sourceSheetName: string,
  sourceRows: ReadonlyArray<ReadonlyArray<CellValue>>,
): ResolvedPivotSource {
  if (sourceRows.length < 2) {
    throw new Error(
      `Pivot "${pivot.name}" source sheet "${sourceSheetName}" needs at least a header row plus one data row (got ${sourceRows.length}).`,
    );
  }

  const header = sourceRows[0];
  const fieldNames: string[] = [];
  for (let i = 0; i < header.length; i++) {
    const v = header[i];
    fieldNames.push(v === null || v === undefined ? `Column${i + 1}` : String(v));
  }

  // Data rows: pad shorter rows with null so every row aligns with the
  // header. Trim rows that are longer than the header — the cache only
  // tracks the declared columns.
  const dataRows: CellValue[][] = [];
  for (let r = 1; r < sourceRows.length; r++) {
    const row = sourceRows[r];
    const padded: CellValue[] = [];
    for (let c = 0; c < fieldNames.length; c++) {
      padded.push(c < row.length ? row[c] : null);
    }
    dataRows.push(padded);
  }

  // Validate that every named field exists in the header row.
  const headerSet = new Set(fieldNames);
  const namedFields = [...(pivot.rows ?? []), ...(pivot.columns ?? []), ...(pivot.pages ?? [])];
  for (const v of pivot.values) namedFields.push(v.field);
  for (const name of namedFields) {
    if (!headerSet.has(name)) {
      throw new Error(
        `Pivot "${pivot.name}" references field "${name}" which is not in the source header (have: ${fieldNames.join(", ")}).`,
      );
    }
  }

  const ref = pivot.sourceRange ?? autoRange(fieldNames.length, sourceRows.length);

  return {
    sheetName: sourceSheetName,
    fieldNames,
    dataRows,
    ref,
  };
}

// ── Field Metadata ────────────────────────────────────────────────────

type FieldKind = "string" | "number";

interface FieldMeta {
  /** Field index. Matches the position in `cache.fieldNames`. */
  index: number;
  /** Field name. */
  name: string;
  /** OOXML axis the field is placed on; `hidden` when unused. */
  axis: "row" | "col" | "page" | "data" | "hidden";
  /** Order within its axis (0-based). Only meaningful when `axis !== "hidden"`. */
  axisOrder: number;
  /**
   * Inferred data type. Numeric fields skip the shared-items table
   * (Excel computes ranges from records); string / mixed fields keep a
   * sorted shared-items list so row/col items can resolve to indexes.
   */
  kind: FieldKind;
  /**
   * Sorted unique string values, present when `kind === "string"`. The
   * record-side serialiser maps cell values to indexes into this list
   * via `<x v="N"/>` tokens.
   */
  sharedStrings?: string[];
  /** Lookup helper: shared-string value → index. */
  sharedIndex?: Map<string, number>;
}

interface PivotFieldsMeta {
  fields: FieldMeta[];
  /** All data-field placements, in declaration order. */
  dataFieldDefs: Array<{
    fieldIndex: number;
    function: PivotDataFieldFunction;
    displayName: string;
    numberFormat?: string;
  }>;
}

function buildFieldMetadata(pivot: WritePivotTable, source: ResolvedPivotSource): PivotFieldsMeta {
  const fields: FieldMeta[] = source.fieldNames.map((name, index) => ({
    index,
    name,
    axis: "hidden" as FieldMeta["axis"],
    axisOrder: 0,
    kind: inferFieldKind(source.dataRows, index),
  }));

  const placeAxis = (
    names: ReadonlyArray<string> | undefined,
    axis: "row" | "col" | "page",
  ): void => {
    if (!names) return;
    for (let i = 0; i < names.length; i++) {
      const f = fields.find((entry) => entry.name === names[i]);
      // Validation in resolvePivotSource guarantees a match.
      if (f) {
        f.axis = axis;
        f.axisOrder = i;
      }
    }
  };

  placeAxis(pivot.rows, "row");
  placeAxis(pivot.columns, "col");
  placeAxis(pivot.pages, "page");

  const dataFieldDefs: PivotFieldsMeta["dataFieldDefs"] = [];
  for (let i = 0; i < pivot.values.length; i++) {
    const v = pivot.values[i];
    const idx = source.fieldNames.indexOf(v.field);
    // Already validated in resolvePivotSource.
    if (idx === -1) continue;

    fields[idx].axis = "data";
    fields[idx].axisOrder = i;

    const fn: PivotDataFieldFunction = v.function ?? "sum";
    const displayName = v.displayName ?? defaultDataFieldName(fn, v.field);
    const entry: PivotFieldsMeta["dataFieldDefs"][number] = {
      fieldIndex: idx,
      function: fn,
      displayName,
    };
    if (v.numberFormat !== undefined) entry.numberFormat = v.numberFormat;
    dataFieldDefs.push(entry);
  }

  // Build shared-items tables for string fields placed on row / column /
  // page axes. Data-axis numeric fields skip this — their items are
  // streamed into the records as `<n v="..."/>` literals.
  for (const f of fields) {
    if (f.axis === "data" || f.axis === "hidden") continue;
    if (f.kind !== "string") continue;
    const seen = new Set<string>();
    const ordered: string[] = [];
    for (const row of source.dataRows) {
      const cell = row[f.index];
      const text = cell === null || cell === undefined ? "" : String(cell);
      if (!seen.has(text)) {
        seen.add(text);
        ordered.push(text);
      }
    }
    f.sharedStrings = ordered;
    f.sharedIndex = new Map(ordered.map((s, i) => [s, i]));
  }

  return { fields, dataFieldDefs };
}

function inferFieldKind(rows: ReadonlyArray<ReadonlyArray<CellValue>>, col: number): FieldKind {
  // A field is numeric when *every non-empty cell* parses as a finite
  // number. A single non-numeric value forces the field to `string`,
  // which sends every value through the shared-items table. This
  // mirrors what Excel does on `Refresh` — it picks the type from the
  // most permissive column.
  for (const row of rows) {
    const cell = row[col];
    if (cell === null || cell === undefined || cell === "") continue;
    if (typeof cell === "number" && Number.isFinite(cell)) continue;
    return "string";
  }
  return "number";
}

// ── Cache Definition ──────────────────────────────────────────────────

function buildCacheDefinition(source: ResolvedPivotSource, meta: PivotFieldsMeta): string {
  const cacheSourceEl = xmlElement("cacheSource", { type: "worksheet" }, [
    xmlSelfClose("worksheetSource", {
      ref: source.ref,
      sheet: source.sheetName,
    }),
  ]);

  const cacheFieldElements: string[] = [];
  for (const f of meta.fields) {
    cacheFieldElements.push(buildCacheField(f, source.dataRows));
  }

  return xmlDocument(
    "pivotCacheDefinition",
    {
      xmlns: NS_SPREADSHEET,
      "xmlns:r": NS_R,
      "r:id": "rId1",
      refreshOnLoad: 1,
      refreshedBy: "hucre",
      refreshedDate: 0,
      createdVersion: 6,
      refreshedVersion: 6,
      minRefreshableVersion: 3,
      recordCount: source.dataRows.length,
    },
    [cacheSourceEl, xmlElement("cacheFields", { count: meta.fields.length }, cacheFieldElements)],
  );
}

function buildCacheField(
  field: FieldMeta,
  dataRows: ReadonlyArray<ReadonlyArray<CellValue>>,
): string {
  if (field.kind === "string") {
    const items: string[] = [];
    if (field.sharedStrings) {
      for (const s of field.sharedStrings) {
        items.push(xmlSelfClose("s", { v: s }));
      }
    }
    return xmlElement("cacheField", { name: field.name, numFmtId: 0 }, [
      xmlElement("sharedItems", { count: items.length }, items),
    ]);
  }

  // Numeric field — compute min/max and emit a containsNumber marker
  // so Excel does not try to inflate `<sharedItems>` itself on refresh.
  let min = Infinity;
  let max = -Infinity;
  let containsBlank = false;
  for (const row of dataRows) {
    const v = row[field.index];
    if (v === null || v === undefined || v === "") {
      containsBlank = true;
      continue;
    }
    if (typeof v === "number") {
      if (v < min) min = v;
      if (v > max) max = v;
    }
  }
  const attrs: Record<string, string | number> = {
    containsSemiMixedTypes: 0,
    containsString: 0,
    containsNumber: 1,
    containsInteger: Number.isInteger(min) && Number.isInteger(max) ? 1 : 0,
  };
  if (containsBlank) attrs.containsBlank = 1;
  if (Number.isFinite(min)) attrs.minValue = min;
  if (Number.isFinite(max)) attrs.maxValue = max;

  return xmlElement("cacheField", { name: field.name, numFmtId: 0 }, [
    xmlSelfClose("sharedItems", attrs),
  ]);
}

function buildCacheDefinitionRels(): string {
  return xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, [
    xmlSelfClose("Relationship", {
      Id: "rId1",
      Type: REL_PIVOT_CACHE_RECORDS,
      Target: "pivotCacheRecords1.xml",
    }),
  ]);
}

// ── Cache Records ─────────────────────────────────────────────────────

function buildCacheRecords(source: ResolvedPivotSource, meta: PivotFieldsMeta): string {
  const recordElements: string[] = [];
  for (const row of source.dataRows) {
    const cells: string[] = [];
    for (const f of meta.fields) {
      cells.push(buildRecordCell(row[f.index], f));
    }
    recordElements.push(xmlElement("r", undefined, cells));
  }
  return xmlDocument(
    "pivotCacheRecords",
    {
      xmlns: NS_SPREADSHEET,
      "xmlns:r": NS_R,
      count: source.dataRows.length,
    },
    recordElements,
  );
}

function buildRecordCell(value: CellValue, field: FieldMeta): string {
  if (value === null || value === undefined || value === "") {
    return xmlSelfClose("m");
  }

  if (field.kind === "string") {
    // Shared-items index lookup. The string was registered during
    // metadata building, so the lookup is always present.
    const idx = field.sharedIndex?.get(String(value));
    if (idx !== undefined) {
      return xmlSelfClose("x", { v: idx });
    }
    // Fallback: emit as inline string. Should not happen unless the
    // caller mutates the source data between resolve and write.
    return xmlSelfClose("s", { v: String(value) });
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return xmlSelfClose("n", { v: value });
  }

  if (typeof value === "boolean") {
    return xmlSelfClose("b", { v: value ? 1 : 0 });
  }

  // Date or other — fall back to inline string.
  return xmlSelfClose("s", { v: String(value) });
}

// ── Pivot Table Definition ────────────────────────────────────────────

function buildPivotTable(pivot: WritePivotTable, meta: PivotFieldsMeta, cacheId: number): string {
  const rowFields = meta.fields
    .filter((f) => f.axis === "row")
    .sort((a, b) => a.axisOrder - b.axisOrder);
  const colFields = meta.fields
    .filter((f) => f.axis === "col")
    .sort((a, b) => a.axisOrder - b.axisOrder);
  const pageFields = meta.fields
    .filter((f) => f.axis === "page")
    .sort((a, b) => a.axisOrder - b.axisOrder);

  const targetCell = pivot.targetCell ?? "A1";
  const styleName = pivot.styleName ?? "PivotStyleLight16";
  const dataCaption = pivot.dataCaption ?? "Values";

  // The output range must cover at least the header rows. Excel
  // resizes the area on refresh, but a sensible default avoids a
  // collapsed-frame look on first open.
  const location = computeLocation(targetCell, rowFields.length, colFields.length, meta);

  const parts: string[] = [];

  parts.push(xmlSelfClose("location", location));

  // ── pivotFields ──
  const pivotFieldElements: string[] = [];
  for (const f of meta.fields) {
    pivotFieldElements.push(buildPivotFieldElement(f));
  }
  parts.push(xmlElement("pivotFields", { count: meta.fields.length }, pivotFieldElements));

  // ── rowFields / rowItems ──
  if (rowFields.length > 0) {
    parts.push(
      xmlElement(
        "rowFields",
        { count: rowFields.length },
        rowFields.map((f) => xmlSelfClose("field", { x: f.index })),
      ),
    );
    // One <i/> per row item placeholder. Excel rebuilds the real items
    // on refresh; emitting a single grand-total row keeps the layout
    // valid in the meantime.
    parts.push(
      xmlElement("rowItems", { count: 1 }, [xmlElement("i", undefined, [xmlSelfClose("x")])]),
    );
  }

  // ── colFields / colItems ──
  if (colFields.length > 0) {
    parts.push(
      xmlElement(
        "colFields",
        { count: colFields.length },
        colFields.map((f) => xmlSelfClose("field", { x: f.index })),
      ),
    );
    parts.push(
      xmlElement("colItems", { count: 1 }, [xmlElement("i", undefined, [xmlSelfClose("x")])]),
    );
  } else if (meta.dataFieldDefs.length > 1) {
    // When there is no explicit column axis but multiple data fields,
    // Excel still emits a colFields placeholder for the data axis.
    parts.push(xmlElement("colFields", { count: 1 }, [xmlSelfClose("field", { x: -2 })]));
    const colItemElements: string[] = [];
    for (let i = 0; i < meta.dataFieldDefs.length; i++) {
      colItemElements.push(
        xmlElement("i", i === 0 ? undefined : { i }, [xmlSelfClose("x", { v: i })]),
      );
    }
    parts.push(xmlElement("colItems", { count: meta.dataFieldDefs.length }, colItemElements));
  }

  // ── pageFields ──
  if (pageFields.length > 0) {
    parts.push(
      xmlElement(
        "pageFields",
        { count: pageFields.length },
        pageFields.map((f) => xmlSelfClose("pageField", { fld: f.index, hier: -1 })),
      ),
    );
  }

  // ── dataFields ──
  if (meta.dataFieldDefs.length > 0) {
    const dataFieldElements: string[] = [];
    for (const d of meta.dataFieldDefs) {
      const attrs: Record<string, string | number> = {
        name: d.displayName,
        fld: d.fieldIndex,
        baseField: 0,
        baseItem: 0,
      };
      if (d.function !== "sum") {
        attrs.subtotal = d.function;
      }
      if (d.numberFormat !== undefined) {
        attrs.numFmtId = 0; // numFmtId is style-driven; the format is lost on roundtrip
      }
      dataFieldElements.push(xmlSelfClose("dataField", attrs));
    }
    parts.push(xmlElement("dataFields", { count: meta.dataFieldDefs.length }, dataFieldElements));
  }

  // ── pivotTableStyleInfo ──
  parts.push(
    xmlSelfClose("pivotTableStyleInfo", {
      name: styleName,
      showRowHeaders: 1,
      showColHeaders: 1,
      showRowStripes: 0,
      showColStripes: 0,
      showLastColumn: 1,
    }),
  );

  return xmlDocument(
    "pivotTableDefinition",
    {
      xmlns: NS_SPREADSHEET,
      "xmlns:r": NS_R,
      name: pivot.name,
      cacheId,
      applyNumberFormats: 0,
      applyBorderFormats: 0,
      applyFontFormats: 0,
      applyPatternFormats: 0,
      applyAlignmentFormats: 0,
      applyWidthHeightFormats: 1,
      dataCaption,
      updatedVersion: 6,
      minRefreshableVersion: 3,
      useAutoFormatting: 1,
      itemPrintTitles: 1,
      createdVersion: 6,
      indent: 0,
      outline: 1,
      outlineData: 1,
      multipleFieldFilters: 0,
    },
    parts,
  );
}

function buildPivotFieldElement(field: FieldMeta): string {
  if (field.axis === "data") {
    return xmlSelfClose("pivotField", { dataField: 1, showAll: 0 });
  }
  if (field.axis === "hidden") {
    return xmlSelfClose("pivotField", { showAll: 0 });
  }

  // Row / col / page axis. Emit one <item/> per shared-item entry plus
  // a trailing <item t="default"/> for the subtotal row Excel injects.
  const axisAttr = field.axis === "row" ? "axisRow" : field.axis === "col" ? "axisCol" : "axisPage";

  const itemElements: string[] = [];
  if (field.sharedStrings) {
    for (let i = 0; i < field.sharedStrings.length; i++) {
      itemElements.push(xmlSelfClose("item", { x: i }));
    }
  }
  itemElements.push(xmlSelfClose("item", { t: "default" }));

  return xmlElement("pivotField", { axis: axisAttr, showAll: 0 }, [
    xmlElement("items", { count: itemElements.length }, itemElements),
  ]);
}

function buildPivotTableRels(): string {
  return xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, [
    xmlSelfClose("Relationship", {
      Id: "rId1",
      Type: REL_PIVOT_CACHE_DEFINITION,
      Target: "../pivotCache/pivotCacheDefinition1.xml",
    }),
  ]);
}

// ── Helpers ───────────────────────────────────────────────────────────

/**
 * Build the `<location>` block. Phase 1 reserves a 4×N rectangle below
 * `targetCell` — Excel resizes on first refresh. The exact size doesn't
 * matter for validity, only for the initial empty layout footprint.
 */
function computeLocation(
  targetCell: string,
  rowFieldCount: number,
  colFieldCount: number,
  meta: PivotFieldsMeta,
): Record<string, string | number> {
  const { col, row } = parseCellRef(targetCell);

  // Header rows: 1 (page filters) + 1 (column field header per col field) + 1 (data field header)
  const firstHeaderRow = 0;
  const firstDataRow = Math.max(1, colFieldCount) + (meta.dataFieldDefs.length > 1 ? 1 : 0);
  const firstDataCol = Math.max(1, rowFieldCount);

  const widthCols = Math.max(2, rowFieldCount + Math.max(1, meta.dataFieldDefs.length));
  const heightRows = firstDataRow + 2; // 2 placeholder data rows

  const startRef = encodeCellRef(row, col);
  const endRef = encodeCellRef(row + heightRows - 1, col + widthCols - 1);

  return {
    ref: `${startRef}:${endRef}`,
    firstHeaderRow,
    firstDataRow,
    firstDataCol,
  };
}

/** Parse `"B3"` → `{col: 1, row: 2}` (0-based). */
function parseCellRef(cell: string): { col: number; row: number } {
  const m = /^([A-Z]+)(\d+)$/i.exec(cell.trim());
  if (!m) {
    throw new Error(`Invalid pivot targetCell "${cell}" — expected an A1-style reference`);
  }
  const colLetters = m[1].toUpperCase();
  let col = 0;
  for (let i = 0; i < colLetters.length; i++) {
    col = col * 26 + (colLetters.charCodeAt(i) - 64);
  }
  return { col: col - 1, row: parseInt(m[2], 10) - 1 };
}

function encodeCellRef(row: number, col: number): string {
  let n = col;
  let letters = "";
  while (n >= 0) {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  }
  return `${letters}${row + 1}`;
}

/** Auto-fit a `<worksheetSource ref>` to the source sheet's row count. */
function autoRange(colCount: number, rowCount: number): string {
  if (colCount <= 0 || rowCount <= 0) {
    throw new Error("Pivot source range must contain at least one column and row");
  }
  const start = encodeCellRef(0, 0);
  const end = encodeCellRef(rowCount - 1, colCount - 1);
  return `${start}:${end}`;
}

/** Excel's auto-label for a data field: `"Sum of Revenue"`, `"Count of Region"`, etc. */
function defaultDataFieldName(fn: PivotDataFieldFunction, field: string): string {
  const label =
    fn === "sum"
      ? "Sum"
      : fn === "count"
        ? "Count"
        : fn === "average"
          ? "Average"
          : fn === "max"
            ? "Max"
            : fn === "min"
              ? "Min"
              : fn === "product"
                ? "Product"
                : fn === "countNums"
                  ? "Count Nums"
                  : fn === "stdDev"
                    ? "StdDev"
                    : fn === "stdDevp"
                      ? "StdDevp"
                      : fn === "var"
                        ? "Var"
                        : "Varp";
  return `${label} of ${field}`;
}

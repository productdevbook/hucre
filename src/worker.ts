// ── Web Worker Serialization Helpers ────────────────────────────────
// Utilities for passing Workbook objects across the Web Worker boundary.
// postMessage uses the structured clone algorithm, which handles most
// types but NOT Map instances. Date objects ARE supported by structured
// clone, but some environments or wrappers may serialize to JSON
// intermediarily, losing Date instances.  These helpers guarantee safe
// transfer in all environments.
// ────────────────────────────────────────────────────────────────────

import type {
  Workbook,
  Sheet,
  Cell,
  CellValue,
  RowDef,
  WorkbookProperties,
  SheetImage,
} from "./_types";

// ── Serialized Types ────────────────────────────────────────────────

/** A Cell with Date values converted to `{ __date: string }` markers. */
export interface SerializedCell {
  value: SerializedCellValue;
  type: Cell["type"];
  style?: Cell["style"];
  formula?: Cell["formula"];
  formulaResult?: SerializedCellValue;
  richText?: Cell["richText"];
  hyperlink?: Cell["hyperlink"];
  comment?: Cell["comment"];
}

/** CellValue with Dates replaced by ISO-string markers. */
export type SerializedCellValue = string | number | boolean | null | { __date: string };

/** SheetImage with `data` as a plain array (Uint8Array is structured-clone safe, but JSON isn't). */
export interface SerializedSheetImage {
  data: number[];
  type: SheetImage["type"];
  anchor: SheetImage["anchor"];
  width?: SheetImage["width"];
  height?: SheetImage["height"];
  altText?: SheetImage["altText"];
  title?: SheetImage["title"];
}

/** A Sheet with Maps converted to plain objects/arrays and Dates serialized. */
export interface SerializedSheet {
  name: string;
  rows: SerializedCellValue[][];
  cells?: Record<string, SerializedCell>;
  columns?: Sheet["columns"];
  rowDefs?: Array<[number, RowDef]>;
  merges?: Sheet["merges"];
  dataValidations?: Sheet["dataValidations"];
  conditionalRules?: Sheet["conditionalRules"];
  autoFilter?: Sheet["autoFilter"];
  freezePane?: Sheet["freezePane"];
  images?: SerializedSheetImage[];
  protection?: Sheet["protection"];
  pageSetup?: Sheet["pageSetup"];
  headerFooter?: Sheet["headerFooter"];
  view?: Sheet["view"];
  hidden?: Sheet["hidden"];
  veryHidden?: Sheet["veryHidden"];
  tables?: Sheet["tables"];
  a11y?: Sheet["a11y"];
}

/** Serialized WorkbookProperties with Dates as ISO markers. */
export interface SerializedWorkbookProperties {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  created?: { __date: string };
  modified?: { __date: string };
  company?: string;
  manager?: string;
  category?: string;
  custom?: Record<string, string | number | boolean | { __date: string }>;
}

/** A Workbook safe to pass through postMessage / JSON. */
export interface SerializedWorkbook {
  sheets: SerializedSheet[];
  properties?: SerializedWorkbookProperties;
  namedRanges?: Workbook["namedRanges"];
  dateSystem?: Workbook["dateSystem"];
  defaultFont?: Workbook["defaultFont"];
  activeSheet?: Workbook["activeSheet"];
  externalLinks?: Workbook["externalLinks"];
}

// ── Serialize ───────────────────────────────────────────────────────

/**
 * Serialize a CellValue, converting Date instances to ISO-string markers.
 */
function serializeCellValue(v: CellValue): SerializedCellValue {
  if (v instanceof Date) {
    return { __date: v.toISOString() };
  }
  return v;
}

/**
 * Serialize a Cell object.
 */
function serializeCell(cell: Cell): SerializedCell {
  const out: SerializedCell = {
    value: serializeCellValue(cell.value),
    type: cell.type,
  };
  if (cell.style !== undefined) out.style = cell.style;
  if (cell.formula !== undefined) out.formula = cell.formula;
  if (cell.formulaResult !== undefined) {
    out.formulaResult = serializeCellValue(cell.formulaResult);
  }
  if (cell.richText !== undefined) out.richText = cell.richText;
  if (cell.hyperlink !== undefined) out.hyperlink = cell.hyperlink;
  if (cell.comment !== undefined) out.comment = cell.comment;
  return out;
}

/**
 * Serialize a SheetImage, converting Uint8Array to a plain number array.
 */
function serializeImage(img: SheetImage): SerializedSheetImage {
  const out: SerializedSheetImage = {
    data: Array.from(img.data),
    type: img.type,
    anchor: img.anchor,
  };
  if (img.width !== undefined) out.width = img.width;
  if (img.height !== undefined) out.height = img.height;
  if (img.altText !== undefined) out.altText = img.altText;
  if (img.title !== undefined) out.title = img.title;
  return out;
}

/**
 * Serialize a Sheet, converting Maps to plain objects/arrays and Dates
 * to ISO-string markers.
 */
function serializeSheet(sheet: Sheet): SerializedSheet {
  const out: SerializedSheet = {
    name: sheet.name,
    rows: sheet.rows.map((row) => row.map(serializeCellValue)),
  };

  if (sheet.cells) {
    const cells: Record<string, SerializedCell> = {};
    for (const [key, cell] of sheet.cells) {
      cells[key] = serializeCell(cell);
    }
    out.cells = cells;
  }

  if (sheet.columns) out.columns = sheet.columns;

  if (sheet.rowDefs) {
    out.rowDefs = Array.from(sheet.rowDefs.entries());
  }

  if (sheet.merges) out.merges = sheet.merges;
  if (sheet.dataValidations) out.dataValidations = sheet.dataValidations;
  if (sheet.conditionalRules) out.conditionalRules = sheet.conditionalRules;
  if (sheet.autoFilter) out.autoFilter = sheet.autoFilter;
  if (sheet.freezePane) out.freezePane = sheet.freezePane;

  if (sheet.images) {
    out.images = sheet.images.map(serializeImage);
  }

  if (sheet.protection) out.protection = sheet.protection;
  if (sheet.pageSetup) out.pageSetup = sheet.pageSetup;
  if (sheet.headerFooter) out.headerFooter = sheet.headerFooter;
  if (sheet.view) out.view = sheet.view;
  if (sheet.hidden) out.hidden = sheet.hidden;
  if (sheet.veryHidden) out.veryHidden = sheet.veryHidden;
  if (sheet.tables) out.tables = sheet.tables;
  if (sheet.a11y) out.a11y = sheet.a11y;

  return out;
}

/**
 * Serialize WorkbookProperties, converting Date fields to ISO markers.
 */
function serializeProperties(props: WorkbookProperties): SerializedWorkbookProperties {
  const out: SerializedWorkbookProperties = {};

  if (props.title !== undefined) out.title = props.title;
  if (props.subject !== undefined) out.subject = props.subject;
  if (props.creator !== undefined) out.creator = props.creator;
  if (props.keywords !== undefined) out.keywords = props.keywords;
  if (props.description !== undefined) out.description = props.description;
  if (props.lastModifiedBy !== undefined) out.lastModifiedBy = props.lastModifiedBy;
  if (props.company !== undefined) out.company = props.company;
  if (props.manager !== undefined) out.manager = props.manager;
  if (props.category !== undefined) out.category = props.category;

  if (props.created instanceof Date) {
    out.created = { __date: props.created.toISOString() };
  }
  if (props.modified instanceof Date) {
    out.modified = { __date: props.modified.toISOString() };
  }

  if (props.custom) {
    const custom: Record<string, string | number | boolean | { __date: string }> = {};
    for (const [k, v] of Object.entries(props.custom)) {
      custom[k] = v instanceof Date ? { __date: v.toISOString() } : v;
    }
    out.custom = custom;
  }

  return out;
}

/**
 * Serialize a Workbook to a transferable format for `postMessage`.
 *
 * Converts:
 * - `Map` instances to plain objects / arrays
 * - `Date` values to `{ __date: "ISO string" }` markers
 * - `Uint8Array` image data to plain number arrays
 *
 * Use {@link deserializeWorkbook} on the receiving side.
 */
export function serializeWorkbook(wb: Workbook): SerializedWorkbook {
  const out: SerializedWorkbook = {
    sheets: wb.sheets.map(serializeSheet),
  };

  if (wb.properties) {
    out.properties = serializeProperties(wb.properties);
  }
  if (wb.namedRanges) out.namedRanges = wb.namedRanges;
  if (wb.dateSystem) out.dateSystem = wb.dateSystem;
  if (wb.defaultFont) out.defaultFont = wb.defaultFont;
  if (wb.activeSheet !== undefined) out.activeSheet = wb.activeSheet;
  if (wb.externalLinks) out.externalLinks = wb.externalLinks;

  return out;
}

// ── Deserialize ─────────────────────────────────────────────────────

/**
 * Deserialize a CellValue, restoring `{ __date }` markers to Date objects.
 */
function deserializeCellValue(v: SerializedCellValue): CellValue {
  if (v !== null && typeof v === "object" && "__date" in v) {
    return new Date(v.__date);
  }
  return v;
}

/**
 * Deserialize a Cell.
 */
function deserializeCell(sc: SerializedCell): Cell {
  const cell: Cell = {
    value: deserializeCellValue(sc.value),
    type: sc.type,
  };
  if (sc.style !== undefined) cell.style = sc.style;
  if (sc.formula !== undefined) cell.formula = sc.formula;
  if (sc.formulaResult !== undefined) {
    cell.formulaResult = deserializeCellValue(sc.formulaResult);
  }
  if (sc.richText !== undefined) cell.richText = sc.richText;
  if (sc.hyperlink !== undefined) cell.hyperlink = sc.hyperlink;
  if (sc.comment !== undefined) cell.comment = sc.comment;
  return cell;
}

/**
 * Deserialize a SheetImage, restoring Uint8Array from plain array.
 */
function deserializeImage(si: SerializedSheetImage): SheetImage {
  const img: SheetImage = {
    data: new Uint8Array(si.data),
    type: si.type,
    anchor: si.anchor,
  };
  if (si.width !== undefined) img.width = si.width;
  if (si.height !== undefined) img.height = si.height;
  if (si.altText !== undefined) img.altText = si.altText;
  if (si.title !== undefined) img.title = si.title;
  return img;
}

/**
 * Deserialize a Sheet, restoring Maps and Date objects.
 */
function deserializeSheet(ss: SerializedSheet): Sheet {
  const sheet: Sheet = {
    name: ss.name,
    rows: ss.rows.map((row) => row.map(deserializeCellValue)),
  };

  if (ss.cells) {
    const cells = new Map<string, Cell>();
    for (const [key, sc] of Object.entries(ss.cells)) {
      cells.set(key, deserializeCell(sc));
    }
    sheet.cells = cells;
  }

  if (ss.columns) sheet.columns = ss.columns;

  if (ss.rowDefs) {
    sheet.rowDefs = new Map(ss.rowDefs);
  }

  if (ss.merges) sheet.merges = ss.merges;
  if (ss.dataValidations) sheet.dataValidations = ss.dataValidations;
  if (ss.conditionalRules) sheet.conditionalRules = ss.conditionalRules;
  if (ss.autoFilter) sheet.autoFilter = ss.autoFilter;
  if (ss.freezePane) sheet.freezePane = ss.freezePane;

  if (ss.images) {
    sheet.images = ss.images.map(deserializeImage);
  }

  if (ss.protection) sheet.protection = ss.protection;
  if (ss.pageSetup) sheet.pageSetup = ss.pageSetup;
  if (ss.headerFooter) sheet.headerFooter = ss.headerFooter;
  if (ss.view) sheet.view = ss.view;
  if (ss.hidden) sheet.hidden = ss.hidden;
  if (ss.veryHidden) sheet.veryHidden = ss.veryHidden;
  if (ss.tables) sheet.tables = ss.tables;
  if (ss.a11y) sheet.a11y = ss.a11y;

  return sheet;
}

/**
 * Deserialize WorkbookProperties, restoring Date fields from ISO markers.
 */
function deserializeProperties(sp: SerializedWorkbookProperties): WorkbookProperties {
  const props: WorkbookProperties = {};

  if (sp.title !== undefined) props.title = sp.title;
  if (sp.subject !== undefined) props.subject = sp.subject;
  if (sp.creator !== undefined) props.creator = sp.creator;
  if (sp.keywords !== undefined) props.keywords = sp.keywords;
  if (sp.description !== undefined) props.description = sp.description;
  if (sp.lastModifiedBy !== undefined) props.lastModifiedBy = sp.lastModifiedBy;
  if (sp.company !== undefined) props.company = sp.company;
  if (sp.manager !== undefined) props.manager = sp.manager;
  if (sp.category !== undefined) props.category = sp.category;

  if (sp.created) {
    props.created = new Date(sp.created.__date);
  }
  if (sp.modified) {
    props.modified = new Date(sp.modified.__date);
  }

  if (sp.custom) {
    const custom: Record<string, string | number | boolean | Date> = {};
    for (const [k, v] of Object.entries(sp.custom)) {
      custom[k] = v !== null && typeof v === "object" && "__date" in v ? new Date(v.__date) : v;
    }
    props.custom = custom;
  }

  return props;
}

/**
 * Deserialize a SerializedWorkbook back to a Workbook.
 *
 * Restores:
 * - Plain objects / arrays back to `Map` instances
 * - `{ __date: "ISO string" }` markers back to `Date` objects
 * - Plain number arrays back to `Uint8Array` for image data
 *
 * Use after receiving a serialized workbook via `postMessage`.
 */
export function deserializeWorkbook(data: SerializedWorkbook): Workbook {
  const wb: Workbook = {
    sheets: data.sheets.map(deserializeSheet),
  };

  if (data.properties) {
    wb.properties = deserializeProperties(data.properties);
  }
  if (data.namedRanges) wb.namedRanges = data.namedRanges;
  if (data.dateSystem) wb.dateSystem = data.dateSystem;
  if (data.defaultFont) wb.defaultFont = data.defaultFont;
  if (data.activeSheet !== undefined) wb.activeSheet = data.activeSheet;
  if (data.externalLinks) wb.externalLinks = data.externalLinks;

  return wb;
}

// ── Worker-safe Function List ───────────────────────────────────────

/**
 * List of defter export names that are safe to call inside a Web Worker.
 * All core functions work in Web Workers since defter has zero DOM
 * dependencies. This list is provided for documentation and tooling.
 */
export const WORKER_SAFE_FUNCTIONS: string[] = [
  // High-level API
  "read",
  "write",
  "readObjects",
  "writeObjects",
  // XLSX
  "readXlsx",
  "writeXlsx",
  "openXlsx",
  "saveXlsx",
  "hashSheetPassword",
  "streamXlsxRows",
  "XlsxStreamWriter",
  // ODS
  "readOds",
  "writeOds",
  // CSV
  "parseCsv",
  "parseCsvObjects",
  "detectDelimiter",
  "stripBom",
  "writeCsv",
  "writeCsvObjects",
  "formatCsvValue",
  "streamCsvRows",
  "CsvStreamWriter",
  // Schema
  "validateWithSchema",
  // Date utilities
  "serialToDate",
  "dateToSerial",
  "isDateFormat",
  "formatDate",
  "parseDate",
  "serialToTime",
  "timeToSerial",
  // Sheet operations
  "insertRows",
  "deleteRows",
  "insertColumns",
  "deleteColumns",
  "moveRows",
  "hideRows",
  "hideColumns",
  "groupRows",
];

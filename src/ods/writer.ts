// ── ODS Writer ──────────────────────────────────────────────────────
// Generates valid OpenDocument Spreadsheet (.ods) files.

import type {
  WriteOptions,
  WriteOutput,
  CellValue,
  WorkbookProperties,
  WriteSheet,
  Cell,
  CellStyle,
  MergeRange,
} from "../_types";
import { ZipWriter } from "../zip/writer";
import { xmlDocument, xmlElement, xmlSelfClose, xmlEscape } from "../xml/writer";

const encoder = /* @__PURE__ */ new TextEncoder();

// ── ODS Namespaces ──────────────────────────────────────────────────

const NS_OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const NS_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const NS_TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const NS_STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const NS_FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const NS_NUMBER = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const NS_SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const NS_META = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const NS_DC = "http://purl.org/dc/elements/1.1/";
const NS_XLINK = "http://www.w3.org/1999/xlink";
const NS_OF = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";

const MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet";

// ── Helpers ─────────────────────────────────────────────────────────

/**
 * Format a number for display in <text:p>.
 * - Integers (including floats with no fractional part like 12.0) → "12"
 * - Floats → reasonable decimal places, no floating-point artifacts
 */
function formatNumberDisplay(value: number): string {
  if (Number.isInteger(value)) return String(value);
  // Use toPrecision(15) to avoid floating-point artifacts (JS has ~17 significant digits),
  // then parseFloat to strip trailing zeros
  return String(parseFloat(value.toPrecision(15)));
}

function formatOdsDate(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, "Z");
}

function formatOdsDateValue(date: Date): string {
  // ODS date values use ISO 8601 without time zone: YYYY-MM-DDTHH:MM:SS
  // Must use UTC methods to avoid local timezone offset corruption
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, "0");
  const d = String(date.getUTCDate()).padStart(2, "0");
  const hh = String(date.getUTCHours()).padStart(2, "0");
  const mm = String(date.getUTCMinutes()).padStart(2, "0");
  const ss = String(date.getUTCSeconds()).padStart(2, "0");
  return `${y}-${m}-${d}T${hh}:${mm}:${ss}`;
}

// ── Style Generation ────────────────────────────────────────────────

/** Maps a CellStyle to a unique string key for deduplication */
function styleKey(style: CellStyle): string {
  const parts: string[] = [];
  if (style.font?.bold) parts.push("b");
  if (style.font?.italic) parts.push("i");
  if (style.font?.size) parts.push(`sz${style.font.size}`);
  if (style.font?.color?.rgb) parts.push(`fc${style.font.color.rgb}`);
  if (style.fill?.type === "pattern" && style.fill.fgColor?.rgb) {
    parts.push(`bg${style.fill.fgColor.rgb}`);
  }
  return parts.join("|");
}

/** Generate a <style:style> element for a cell style */
function generateStyleElement(name: string, style: CellStyle): string {
  const textProps: Record<string, string> = {};
  const cellProps: Record<string, string> = {};

  if (style.font?.bold) {
    textProps["fo:font-weight"] = "bold";
  }
  if (style.font?.italic) {
    textProps["fo:font-style"] = "italic";
  }
  if (style.font?.size) {
    textProps["fo:font-size"] = `${style.font.size}pt`;
  }
  if (style.font?.color?.rgb) {
    textProps["fo:color"] = `#${style.font.color.rgb}`;
  }

  if (style.fill?.type === "pattern" && style.fill.fgColor?.rgb) {
    cellProps["fo:background-color"] = `#${style.fill.fgColor.rgb}`;
  }

  const children: string[] = [];
  if (Object.keys(textProps).length > 0) {
    children.push(xmlSelfClose("style:text-properties", textProps));
  }
  if (Object.keys(cellProps).length > 0) {
    children.push(xmlSelfClose("style:table-cell-properties", cellProps));
  }

  return xmlElement("style:style", { "style:name": name, "style:family": "table-cell" }, children);
}

// ── Style Collector ────────────────────────────────────────────────

interface StyleCollector {
  /** Map from style key → style name (e.g. "ce1") */
  styleMap: Map<string, string>;
  /** Map from style name → XML element string */
  styleElements: Map<string, string>;
  /** Counter for generating unique names */
  counter: number;
}

function createStyleCollector(): StyleCollector {
  return { styleMap: new Map(), styleElements: new Map(), counter: 1 };
}

function getOrCreateStyleName(collector: StyleCollector, style: CellStyle): string {
  const key = styleKey(style);
  if (!key) return ""; // No style properties

  const existing = collector.styleMap.get(key);
  if (existing) return existing;

  const name = `ce${collector.counter++}`;
  collector.styleMap.set(key, name);
  collector.styleElements.set(name, generateStyleElement(name, style));
  return name;
}

// ── Formula Conversion ──────────────────────────────────────────────

/**
 * Convert an Excel-style formula to ODS formula syntax.
 * ODS formulas use `of:=` prefix and `[.A1]` cell references.
 */
function excelFormulaToOds(formula: string): string {
  // Convert cell references like A1, $A$1, A1:B2 to ODS [.A1] notation
  // Handle range references like A1:B2 → [.A1:.B2]
  const converted = formula.replace(
    /(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?/g,
    (_match, ref1: string, ref2?: string) => {
      if (ref2) {
        return `[.${ref1}:.${ref2}]`;
      }
      return `[.${ref1}]`;
    },
  );
  return `of:=${converted}`;
}

// ── Cell Serialization ──────────────────────────────────────────────

interface CellContext {
  /** Cell override from sheet.cells */
  cellOverride?: Partial<Cell>;
  /** Style name to apply (from style collector) */
  styleName?: string;
  /** Merge span attributes */
  colSpan?: number;
  rowSpan?: number;
}

function cellToOds(value: CellValue, ctx?: CellContext): string {
  const attrs: Record<string, string> = {};
  const children: string[] = [];

  if (ctx?.styleName) {
    attrs["table:style-name"] = ctx.styleName;
  }
  if (ctx?.colSpan && ctx.colSpan > 1) {
    attrs["table:number-columns-spanned"] = String(ctx.colSpan);
  }
  if (ctx?.rowSpan && ctx.rowSpan > 1) {
    attrs["table:number-rows-spanned"] = String(ctx.rowSpan);
  }

  // Formula
  const formula = ctx?.cellOverride?.formula;
  if (formula) {
    attrs["table:formula"] = excelFormulaToOds(formula);
  }

  // Hyperlink
  const hyperlink = ctx?.cellOverride?.hyperlink;

  if (value === null || value === undefined) {
    if (Object.keys(attrs).length === 0) {
      return xmlSelfClose("table:table-cell");
    }
    return xmlElement("table:table-cell", attrs, children);
  }

  if (typeof value === "string") {
    attrs["office:value-type"] = "string";
    if (hyperlink) {
      const linkEl = xmlElement(
        "text:a",
        { "xlink:href": hyperlink.target, "xlink:type": "simple" },
        xmlEscape(value),
      );
      children.push(xmlElement("text:p", undefined, linkEl));
    } else {
      children.push(xmlElement("text:p", undefined, xmlEscape(value)));
    }
    return xmlElement("table:table-cell", attrs, children);
  }

  if (typeof value === "number") {
    attrs["office:value-type"] = "float";
    attrs["office:value"] = String(value);
    children.push(xmlElement("text:p", undefined, formatNumberDisplay(value)));
    return xmlElement("table:table-cell", attrs, children);
  }

  if (typeof value === "boolean") {
    attrs["office:value-type"] = "boolean";
    attrs["office:boolean-value"] = value ? "true" : "false";
    children.push(xmlElement("text:p", undefined, value ? "TRUE" : "FALSE"));
    return xmlElement("table:table-cell", attrs, children);
  }

  if (value instanceof Date) {
    const dateStr = formatOdsDateValue(value);
    attrs["office:value-type"] = "date";
    attrs["office:date-value"] = dateStr;
    children.push(xmlElement("text:p", undefined, dateStr));
    return xmlElement("table:table-cell", attrs, children);
  }

  if (Object.keys(attrs).length === 0) {
    return xmlSelfClose("table:table-cell");
  }
  return xmlElement("table:table-cell", attrs, children);
}

// ── Merge helpers ───────────────────────────────────────────────────

/** Build a set of covered cell positions from merge ranges */
function buildMergeMap(merges: MergeRange[] | undefined): {
  /** Cells that are the start of a merge: "row,col" → { colSpan, rowSpan } */
  starts: Map<string, { colSpan: number; rowSpan: number }>;
  /** Cells covered by a merge (not the start cell) */
  covered: Set<string>;
} {
  const starts = new Map<string, { colSpan: number; rowSpan: number }>();
  const covered = new Set<string>();

  if (!merges) return { starts, covered };

  for (const m of merges) {
    const colSpan = m.endCol - m.startCol + 1;
    const rowSpan = m.endRow - m.startRow + 1;
    starts.set(`${m.startRow},${m.startCol}`, { colSpan, rowSpan });

    for (let r = m.startRow; r <= m.endRow; r++) {
      for (let c = m.startCol; c <= m.endCol; c++) {
        if (r === m.startRow && c === m.startCol) continue;
        covered.add(`${r},${c}`);
      }
    }
  }

  return { starts, covered };
}

// ── Row serialization with merge and cell override support ──────────

function rowToOds(
  row: CellValue[],
  rowIndex: number,
  sheet: WriteSheet,
  mergeMap: { starts: Map<string, { colSpan: number; rowSpan: number }>; covered: Set<string> },
  styleCollector: StyleCollector,
  maxCol: number,
): string {
  const cellElements: string[] = [];

  // We need to emit cells for the full width including merge-covered columns
  const effectiveMax = Math.max(row.length - 1, maxCol);

  // Find the last column that has meaningful content (value, merge start, covered cell)
  let lastMeaningful = row.length - 1;
  while (
    lastMeaningful >= 0 &&
    (row[lastMeaningful] === null || row[lastMeaningful] === undefined)
  ) {
    lastMeaningful--;
  }
  // Also consider merge starts and covered cells beyond data
  for (let c = lastMeaningful + 1; c <= effectiveMax; c++) {
    const key = `${rowIndex},${c}`;
    if (mergeMap.starts.has(key) || mergeMap.covered.has(key)) {
      lastMeaningful = c;
    }
  }

  let i = 0;
  while (i <= lastMeaningful) {
    const key = `${rowIndex},${i}`;

    // Check if this cell is covered by a merge
    if (mergeMap.covered.has(key)) {
      // Count consecutive covered cells
      let count = 1;
      while (i + count <= lastMeaningful && mergeMap.covered.has(`${rowIndex},${i + count}`)) {
        count++;
      }
      if (count > 1) {
        cellElements.push(
          xmlSelfClose("table:covered-table-cell", {
            "table:number-columns-repeated": String(count),
          }),
        );
      } else {
        cellElements.push(xmlSelfClose("table:covered-table-cell"));
      }
      i += count;
      continue;
    }

    const cell = i < row.length ? row[i] : null;

    // Get cell override for formulas, hyperlinks, styles
    const cellOverride = sheet.cells?.get(key);

    // Build cell context
    const ctx: CellContext = {};
    if (cellOverride) ctx.cellOverride = cellOverride;

    // Merge span
    const mergeInfo = mergeMap.starts.get(key);
    if (mergeInfo) {
      ctx.colSpan = mergeInfo.colSpan;
      ctx.rowSpan = mergeInfo.rowSpan;
    }

    // Style from cell override
    const style = cellOverride?.style;
    if (style) {
      const name = getOrCreateStyleName(styleCollector, style);
      if (name) ctx.styleName = name;
    }

    if (cell === null || cell === undefined) {
      if (Object.keys(ctx).length === 0 && !ctx.cellOverride && !mergeInfo) {
        // Plain empty cell — count consecutive empties
        let count = 1;
        while (
          i + count <= lastMeaningful &&
          (i + count >= row.length || row[i + count] === null || row[i + count] === undefined) &&
          !mergeMap.covered.has(`${rowIndex},${i + count}`) &&
          !mergeMap.starts.has(`${rowIndex},${i + count}`) &&
          !sheet.cells?.has(`${rowIndex},${i + count}`)
        ) {
          count++;
        }
        if (count > 1) {
          cellElements.push(
            xmlSelfClose("table:table-cell", {
              "table:number-columns-repeated": String(count),
            }),
          );
        } else {
          cellElements.push(xmlSelfClose("table:table-cell"));
        }
        i += count;
        continue;
      }
    }

    cellElements.push(cellToOds(cell, ctx));
    i++;
  }

  return xmlElement("table:table-row", undefined, cellElements);
}

// ── content.xml ─────────────────────────────────────────────────────

function writeContentXml(options: WriteOptions): string {
  const { sheets } = options;

  const styleCollector = createStyleCollector();
  const tableElements: string[] = [];

  // First pass: collect styles and build table XML (deferred because styles go before body)
  const sheetXmlParts: string[][] = [];

  for (const sheet of sheets) {
    const children: string[] = [];

    // Resolve rows from rows or data
    let rows: CellValue[][] = [];
    if (sheet.rows) {
      rows = sheet.rows;
    } else if (sheet.data && sheet.columns) {
      // Generate header row + data rows from objects
      const keys = sheet.columns.map((c) => c.key ?? c.header ?? "");
      const hasHeaders = sheet.columns.some((c) => c.header);

      if (hasHeaders) {
        const headerRow = sheet.columns.map((c) => c.header ?? c.key ?? "");
        rows.push(headerRow);
      }

      for (const item of sheet.data) {
        const row = keys.map((k) => (k in item ? (item[k] as CellValue) : null));
        rows.push(row);
      }
    }

    // Build merge map
    const mergeMap = buildMergeMap(sheet.merges);

    // Determine column count (max width across all rows, considering merges)
    let colCount = 0;
    for (const row of rows) {
      if (row.length > colCount) colCount = row.length;
    }
    if (sheet.merges) {
      for (const m of sheet.merges) {
        if (m.endCol + 1 > colCount) colCount = m.endCol + 1;
      }
    }

    // Determine max row needed (considering merges)
    let rowCount = rows.length;
    if (sheet.merges) {
      for (const m of sheet.merges) {
        if (m.endRow + 1 > rowCount) rowCount = m.endRow + 1;
      }
    }

    // Emit table:table-column element to declare column count
    if (colCount > 0) {
      if (colCount > 1) {
        children.push(
          xmlSelfClose("table:table-column", {
            "table:number-columns-repeated": String(colCount),
          }),
        );
      } else {
        children.push(xmlSelfClose("table:table-column"));
      }
    }

    // Emit rows (extend to cover merged rows beyond data)
    for (let r = 0; r < rowCount; r++) {
      const row = r < rows.length ? rows[r] : [];
      children.push(rowToOds(row, r, sheet, mergeMap, styleCollector, colCount - 1));
    }

    sheetXmlParts.push(children);
  }

  // Now build the final XML with styles collected during serialization
  for (let i = 0; i < sheets.length; i++) {
    tableElements.push(
      xmlElement("table:table", { "table:name": sheets[i].name }, sheetXmlParts[i]),
    );
  }

  const spreadsheetBody = xmlElement("office:spreadsheet", undefined, tableElements);
  const body = xmlElement("office:body", undefined, spreadsheetBody);

  // Build automatic styles from collected styles
  const styleXml =
    styleCollector.styleElements.size > 0
      ? xmlElement("office:automatic-styles", undefined, [...styleCollector.styleElements.values()])
      : xmlElement("office:automatic-styles", undefined, "");

  // Build content sections in order per ODS spec:
  // office:scripts, office:font-face-decls, office:automatic-styles, office:body
  const contentParts: string[] = [];
  contentParts.push(xmlSelfClose("office:scripts"));
  contentParts.push(xmlElement("office:font-face-decls", undefined, ""));
  contentParts.push(styleXml);
  contentParts.push(body);

  return xmlDocument(
    "office:document-content",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:table": NS_TABLE,
      "xmlns:text": NS_TEXT,
      "xmlns:style": NS_STYLE,
      "xmlns:fo": NS_FO,
      "xmlns:number": NS_NUMBER,
      "xmlns:svg": NS_SVG,
      "xmlns:xlink": NS_XLINK,
      "xmlns:of": NS_OF,
      "office:version": "1.2",
    },
    contentParts,
  );
}

// ── meta.xml ────────────────────────────────────────────────────────

function writeMetaXml(props?: WorkbookProperties): string {
  const children: string[] = [];

  if (props?.title) {
    children.push(xmlElement("dc:title", undefined, xmlEscape(props.title)));
  }
  if (props?.subject) {
    children.push(xmlElement("dc:subject", undefined, xmlEscape(props.subject)));
  }
  if (props?.creator) {
    children.push(xmlElement("meta:initial-creator", undefined, xmlEscape(props.creator)));
  }
  if (props?.description) {
    children.push(xmlElement("dc:description", undefined, xmlEscape(props.description)));
  }
  if (props?.keywords) {
    children.push(xmlElement("meta:keyword", undefined, xmlEscape(props.keywords)));
  }
  if (props?.created) {
    children.push(xmlElement("meta:creation-date", undefined, formatOdsDate(props.created)));
  }

  const modified = props?.modified ?? new Date();
  children.push(xmlElement("dc:date", undefined, formatOdsDate(modified)));

  children.push(xmlElement("meta:generator", undefined, "defter"));

  const metaContent = xmlElement("office:meta", undefined, children);

  return xmlDocument(
    "office:document-meta",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:meta": NS_META,
      "xmlns:dc": NS_DC,
      "office:version": "1.2",
    },
    metaContent,
  );
}

// ── styles.xml ──────────────────────────────────────────────────────

function writeStylesXml(): string {
  // ODS spec requires these child elements even if empty:
  // office:font-face-decls, office:styles, office:automatic-styles, office:master-styles
  const children: string[] = [];
  children.push(xmlElement("office:font-face-decls", undefined, ""));
  children.push(xmlElement("office:styles", undefined, ""));
  children.push(xmlElement("office:automatic-styles", undefined, ""));
  children.push(xmlElement("office:master-styles", undefined, ""));

  return xmlDocument(
    "office:document-styles",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:style": NS_STYLE,
      "xmlns:text": NS_TEXT,
      "xmlns:table": NS_TABLE,
      "xmlns:fo": NS_FO,
      "xmlns:number": NS_NUMBER,
      "xmlns:svg": NS_SVG,
      "office:version": "1.2",
    },
    children,
  );
}

// ── settings.xml ─────────────────────────────────────────────────────

function writeSettingsXml(): string {
  const NS_CONFIG = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";

  return xmlDocument(
    "office:document-settings",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:config": NS_CONFIG,
      "office:version": "1.2",
    },
    xmlElement("office:settings", undefined, ""),
  );
}

// ── manifest.xml ────────────────────────────────────────────────────

function writeManifestXml(): string {
  const NS_MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

  const entries: string[] = [];
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "/",
      "manifest:version": "1.2",
      "manifest:media-type": MIMETYPE,
    }),
  );
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "content.xml",
      "manifest:media-type": "text/xml",
    }),
  );
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "meta.xml",
      "manifest:media-type": "text/xml",
    }),
  );
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "styles.xml",
      "manifest:media-type": "text/xml",
    }),
  );
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "settings.xml",
      "manifest:media-type": "text/xml",
    }),
  );

  return xmlDocument(
    "manifest:manifest",
    {
      "xmlns:manifest": NS_MANIFEST,
      "manifest:version": "1.2",
    },
    entries,
  );
}

// ── Main Writer ─────────────────────────────────────────────────────

/**
 * Write a workbook to ODS format.
 * Returns a Uint8Array containing the ZIP archive.
 */
export async function writeOds(options: WriteOptions): Promise<WriteOutput> {
  const zip = new ZipWriter();

  // mimetype MUST be the first entry and MUST be stored uncompressed
  zip.add("mimetype", encoder.encode(MIMETYPE), { compress: false });

  // META-INF/manifest.xml
  zip.add("META-INF/manifest.xml", encoder.encode(writeManifestXml()));

  // content.xml — main spreadsheet data
  zip.add("content.xml", encoder.encode(writeContentXml(options)));

  // meta.xml — document metadata
  zip.add("meta.xml", encoder.encode(writeMetaXml(options.properties)));

  // styles.xml — style definitions
  zip.add("styles.xml", encoder.encode(writeStylesXml()));

  // settings.xml — document settings
  zip.add("settings.xml", encoder.encode(writeSettingsXml()));

  return zip.build();
}

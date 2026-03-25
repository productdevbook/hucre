// ── ODS Reader ──────────────────────────────────────────────────────
// Reads OpenDocument Spreadsheet (.ods) files.

import type {
  Workbook,
  ReadOptions,
  ReadInput,
  Sheet,
  CellValue,
  WorkbookProperties,
} from "../_types";
import { ParseError, ZipError } from "../errors";
import { ZipReader } from "../zip/reader";
import { parseXml } from "../xml/parser";
import type { XmlElement } from "../xml/parser";

// ── Helpers ─────────────────────────────────────────────────────────

function toUint8Array(input: ReadInput): Uint8Array {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  throw new ParseError("Unsupported input type. Expected Uint8Array or ArrayBuffer.");
}

function decodeUtf8(data: Uint8Array): string {
  return new TextDecoder("utf-8").decode(data);
}

function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === localName) return child;
  }
  return undefined;
}

function findChildren(el: XmlElement, localName: string): XmlElement[] {
  const result: XmlElement[] = [];
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === localName) result.push(child);
  }
  return result;
}

// ── Cell Value Parsing ──────────────────────────────────────────────

function parseCellValue(cell: XmlElement): CellValue {
  const valueType = cell.attrs["office:value-type"] ?? cell.attrs["calcext:value-type"] ?? "";

  switch (valueType) {
    case "string": {
      // Get text from <text:p> children
      const textP = findChild(cell, "p");
      if (textP) {
        return textP.children.filter((c: unknown) => typeof c === "string").join("");
      }
      // Check office:string-value attribute
      const strVal = cell.attrs["office:string-value"];
      if (strVal !== undefined) return strVal;
      return "";
    }

    case "float":
    case "currency":
    case "percentage": {
      const val = cell.attrs["office:value"];
      if (val !== undefined) return Number(val);
      return null;
    }

    case "boolean": {
      const boolVal = cell.attrs["office:boolean-value"];
      if (boolVal === "true") return true;
      if (boolVal === "false") return false;
      return null;
    }

    case "date": {
      const dateVal = cell.attrs["office:date-value"];
      if (dateVal) {
        const d = new Date(dateVal);
        if (!Number.isNaN(d.getTime())) return d;
      }
      return null;
    }

    case "time": {
      // ODS time values are ISO 8601 durations like PT12H30M
      const timeVal = cell.attrs["office:time-value"];
      if (timeVal) {
        return timeVal;
      }
      return null;
    }

    default: {
      // No explicit type — try to extract text
      const textP = findChild(cell, "p");
      if (textP) {
        const text = textP.children.filter((c: unknown) => typeof c === "string").join("");
        if (text) return text;
      }
      return null;
    }
  }
}

// ── Content XML Parsing ─────────────────────────────────────────────

function parseContentXml(xml: string, options?: ReadOptions): Sheet[] {
  const doc = parseXml(xml);
  const sheets: Sheet[] = [];

  // Navigate: document-content > body > spreadsheet > table
  const body = findChild(doc, "body");
  if (!body) return sheets;

  const spreadsheet = findChild(body, "spreadsheet");
  if (!spreadsheet) return sheets;

  const tables = findChildren(spreadsheet, "table");

  for (const table of tables) {
    const name = table.attrs["table:name"] ?? `Sheet${sheets.length + 1}`;

    // Filter sheets if specified
    if (options?.sheets && options.sheets.length > 0) {
      const shouldRead = options.sheets.some((spec) => {
        if (typeof spec === "string") return spec === name;
        if (typeof spec === "number") return spec === sheets.length;
        return false;
      });
      if (!shouldRead) {
        sheets.push({ name, rows: [] }); // placeholder to maintain index
        continue;
      }
    }

    const rows: CellValue[][] = [];
    const tableRows = findChildren(table, "table-row");

    for (const tableRow of tableRows) {
      const rowRepeat = Number(tableRow.attrs["table:number-rows-repeated"] ?? "1");

      // Collect cell entries with their repeat counts first,
      // so we can trim trailing nulls before expanding
      const cellEntries: Array<{ value: CellValue; repeat: number }> = [];

      for (const child of tableRow.children) {
        if (typeof child === "string") continue;
        const local = child.local || child.tag;

        if (local === "table-cell" || local === "covered-table-cell") {
          const colRepeat = Number(child.attrs["table:number-columns-repeated"] ?? "1");
          const value = parseCellValue(child);
          cellEntries.push({ value, repeat: colRepeat });
        }
      }

      // Trim trailing null entries (avoids expanding huge repeat counts like 16384)
      while (cellEntries.length > 0 && cellEntries[cellEntries.length - 1].value === null) {
        cellEntries.pop();
      }

      // Expand into row data
      const rowData: CellValue[] = [];
      for (const entry of cellEntries) {
        for (let r = 0; r < entry.repeat; r++) {
          rowData.push(entry.value);
        }
      }

      // Cap row repeats for empty rows to avoid memory issues
      // (LibreOffice may emit large row repeats for trailing empty rows)
      const effectiveRowRepeat = rowData.length > 0 ? rowRepeat : 0;

      for (let r = 0; r < effectiveRowRepeat; r++) {
        rows.push(effectiveRowRepeat === 1 ? rowData : [...rowData]);
      }
    }

    // Trim trailing empty rows
    while (rows.length > 0 && rows[rows.length - 1].length === 0) {
      rows.pop();
    }

    sheets.push({ name, rows });
  }

  // If filter was applied, remove placeholder sheets with empty rows
  if (options?.sheets && options.sheets.length > 0) {
    return sheets.filter(
      (s) =>
        s.rows.length > 0 ||
        options.sheets!.some((spec) => {
          if (typeof spec === "string") return spec === s.name;
          return false;
        }),
    );
  }

  return sheets;
}

// ── Meta XML Parsing ────────────────────────────────────────────────

function parseMetaXml(xml: string): Partial<WorkbookProperties> {
  const doc = parseXml(xml);
  const props: Partial<WorkbookProperties> = {};

  // Navigate to office:meta element
  const meta = findChild(doc, "meta");
  if (!meta) return props;

  for (const child of meta.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    const text = child.children.filter((c: unknown) => typeof c === "string").join("");

    switch (local) {
      case "title":
        if (text) props.title = text;
        break;
      case "subject":
        if (text) props.subject = text;
        break;
      case "initial-creator":
        if (text) props.creator = text;
        break;
      case "description":
        if (text) props.description = text;
        break;
      case "keyword":
        if (text) props.keywords = text;
        break;
      case "creation-date":
        if (text) {
          const d = new Date(text);
          if (!Number.isNaN(d.getTime())) props.created = d;
        }
        break;
      case "date":
        if (text) {
          const d = new Date(text);
          if (!Number.isNaN(d.getTime())) props.modified = d;
        }
        break;
    }
  }

  return props;
}

// ── Main Reader ─────────────────────────────────────────────────────

/**
 * Read an ODS file and return a Workbook.
 * Input can be Uint8Array or ArrayBuffer.
 */
export async function readOds(input: ReadInput, options?: ReadOptions): Promise<Workbook> {
  const data = toUint8Array(input);

  // 1. Open ZIP archive
  let zip: ZipReader;
  try {
    zip = new ZipReader(data);
  } catch (err) {
    if (err instanceof ZipError) throw err;
    throw new ParseError("Failed to open ODS file: not a valid ZIP archive", undefined, {
      cause: err,
    });
  }

  // 2. Verify mimetype
  if (zip.has("mimetype")) {
    const mimeData = await zip.extract("mimetype");
    const mime = decodeUtf8(mimeData).trim();
    if (mime !== "application/vnd.oasis.opendocument.spreadsheet") {
      throw new ParseError(`Invalid ODS mimetype: ${mime}`);
    }
  }

  // 3. Parse content.xml (required)
  if (!zip.has("content.xml")) {
    throw new ParseError("Invalid ODS: missing content.xml");
  }
  const contentXml = decodeUtf8(await zip.extract("content.xml"));
  const sheets = parseContentXml(contentXml, options);

  // 4. Parse meta.xml (optional)
  let properties: WorkbookProperties | undefined;
  if (zip.has("meta.xml")) {
    const metaXml = decodeUtf8(await zip.extract("meta.xml"));
    const metaProps = parseMetaXml(metaXml);
    if (Object.keys(metaProps).length > 0) {
      properties = { ...metaProps };
    }
  }

  // 5. Build workbook
  const workbook: Workbook = {
    sheets,
  };

  if (properties) {
    workbook.properties = properties;
  }

  return workbook;
}

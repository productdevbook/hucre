<p align="center">
  <br>
  <img src=".github/assets/cover.svg" alt="hucre — Zero-dependency spreadsheet engine" width="100%">
  <br><br>
  <b style="font-size: 2em;">hucre</b>
  <br><br>
  Zero-dependency spreadsheet engine.
  <br>
  Read & write XLSX, CSV, ODS, JSON, NDJSON, XML. Schema validation, streaming, round-trip preservation. Pure TypeScript, works everywhere.
  <br><br>
  <a href="https://npmjs.com/package/hucre"><img src="https://img.shields.io/npm/v/hucre?style=flat&colorA=18181B&colorB=34d399" alt="npm version"></a>
  <a href="https://npmjs.com/package/hucre"><img src="https://img.shields.io/npm/dm/hucre?style=flat&colorA=18181B&colorB=34d399" alt="npm downloads"></a>
  <a href="https://bundlephobia.com/result?p=hucre"><img src="https://img.shields.io/bundlephobia/minzip/hucre?style=flat&colorA=18181B&colorB=34d399" alt="bundle size"></a>
  <a href="https://github.com/productdevbook/hucre/blob/main/LICENSE"><img src="https://img.shields.io/github/license/productdevbook/hucre?style=flat&colorA=18181B&colorB=34d399" alt="license"></a>
</p>

## Quick Start

```sh
npm install hucre
```

```ts
import { readXlsx, writeXlsx } from "hucre";

// Read an XLSX file
const workbook = await readXlsx(buffer);
console.log(workbook.sheets[0].rows);

// Write an XLSX file
const xlsx = await writeXlsx({
  sheets: [
    {
      name: "Products",
      columns: [
        { header: "Name", key: "name", width: 25 },
        { header: "Price", key: "price", width: 12, numFmt: "$#,##0.00" },
        { header: "Stock", key: "stock", width: 10 },
      ],
      data: [
        { name: "Widget", price: 9.99, stock: 142 },
        { name: "Gadget", price: 24.5, stock: 87 },
      ],
    },
  ],
});
```

## Tree Shaking

Import only what you need:

```ts
import { readXlsx, writeXlsx } from "hucre/xlsx"; // XLSX only
import { parseCsv, writeCsv } from "hucre/csv"; // CSV only (~2 KB gzipped)
import { readOds, writeOds } from "hucre/ods"; // ODS only
import { parseJson, writeNdjson } from "hucre/json"; // JSON / NDJSON
import { readXml, writeXml } from "hucre/xml"; // Tabular XML
```

## Why hucre?

### vs JavaScript / TypeScript Libraries

|                         | hucre     | SheetJS CE    | ExcelJS   | xlsx-js-style |
| ----------------------- | --------- | ------------- | --------- | ------------- |
| **Dependencies**        | 0         | 0\*           | 12 (CVEs) | 0\*           |
| **Bundle (gzip)**       | ~18 KB    | ~300 KB       | ~500 KB   | ~300 KB       |
| **ESM native**          | Yes       | Partial       | No (CJS)  | Partial       |
| **TypeScript**          | Native    | Bolted-on     | Bolted-on | Bolted-on     |
| **Edge runtime**        | Yes       | No            | No        | No            |
| **CSP compliant**       | Yes       | Yes           | No (eval) | Yes           |
| **npm published**       | Yes       | No (CDN only) | Stale     | Yes           |
| **Read + Write**        | Yes       | Yes (Pro $)   | Yes       | Yes           |
| **Styling**             | Yes       | No (Pro $)    | Yes       | Yes           |
| **Cond. formatting**    | Yes (all) | No (Pro $)    | Partial   | No            |
| **Stream read + write** | Yes       | CSV only      | Yes       | CSV only      |
| **ODS support**         | Yes       | Yes           | No        | Yes           |
| **Round-trip**          | Yes       | Partial       | Partial   | Partial       |
| **Sparklines**          | Yes       | No            | No        | No            |
| **Tables**              | Yes       | Yes           | Yes       | Yes           |
| **Images**              | Yes       | No (Pro $)    | Yes       | No            |

\* SheetJS removed itself from npm; must install from CDN tarball.

### vs Libraries in Other Languages

|                       | hucre (TS)   | openpyxl (Py) | XlsxWriter (Py) | rust_xlsxwriter | Apache POI (Java) |
| --------------------- | ------------ | ------------- | --------------- | --------------- | ----------------- |
| **Read XLSX**         | Yes          | Yes           | No              | No              | Yes               |
| **Write XLSX**        | Yes          | Yes           | Yes             | Yes             | Yes               |
| **Streaming**         | Read+Write   | Write-only    | No              | const_memory    | SXSSF (write)     |
| **Charts**            | Round-trip   | 15+ types     | 9 types         | 12+ types       | Limited           |
| **Pivot tables**      | No           | Read-only     | No              | No              | Limited           |
| **Cond. formatting**  | Yes (all)    | Yes           | Yes             | Yes             | Yes               |
| **Sparklines**        | Yes          | No            | Yes             | Yes             | No                |
| **Formula eval**      | No           | No            | No              | No              | Yes               |
| **Multi-format**      | XLSX/ODS/CSV | XLSX only     | XLSX only       | XLSX only       | XLS/XLSX          |
| **Zero dependencies** | Yes          | lxml optional | No              | Yes             | No                |

## Features

### Reading

```ts
import { readXlsx } from "hucre/xlsx";

const wb = await readXlsx(uint8Array, {
  sheets: [0, "Products"], // Filter sheets by index or name
  readStyles: true, // Parse cell styles
  dateSystem: "auto", // Auto-detect 1900/1904
});

for (const sheet of wb.sheets) {
  console.log(sheet.name); // "Products"
  console.log(sheet.rows); // CellValue[][]
  console.log(sheet.merges); // MergeRange[]
}
```

Supported cell types: strings, numbers, booleans, dates, formulas, rich text, errors, inline strings.

### Writing

```ts
import { writeXlsx } from "hucre/xlsx";

const buffer = await writeXlsx({
  sheets: [
    {
      name: "Report",
      columns: [
        { header: "Date", key: "date", width: 15, numFmt: "yyyy-mm-dd" },
        { header: "Revenue", key: "revenue", width: 15, numFmt: "$#,##0.00" },
        { header: "Active", key: "active", width: 10 },
      ],
      data: [
        { date: new Date("2026-01-15"), revenue: 12500, active: true },
        { date: new Date("2026-01-16"), revenue: 8900, active: false },
      ],
      freezePane: { rows: 1 },
      autoFilter: { range: "A1:C3" },
    },
  ],
  defaultFont: { name: "Calibri", size: 11 },
});
```

Features: cell styles, auto column widths, merged cells, freeze/split panes, auto-filter with criteria, data validation, hyperlinks, images (PNG/JPEG/GIF/SVG/WebP), comments, tables, conditional formatting (cellIs/colorScale/dataBar/iconSet), named ranges, print settings, page breaks, sheet protection, workbook protection, rich text, shared/array/dynamic formulas, sparklines, textboxes, background images, number formats, hidden sheets, HTML/Markdown/JSON/TSV export, template engine.

### Auto Column Width

```ts
const buffer = await writeXlsx({
  sheets: [
    {
      name: "Products",
      columns: [
        { header: "Name", key: "name", autoWidth: true },
        { header: "Price", key: "price", autoWidth: true, numFmt: "$#,##0.00" },
        { header: "SKU", key: "sku", autoWidth: true },
      ],
      data: products,
    },
  ],
});
```

Calculates optimal column widths from cell content — font-aware, handles CJK double-width characters, number formats, min/max constraints.

### Data Validation

```ts
const buffer = await writeXlsx({
  sheets: [
    {
      name: "Sheet1",
      rows: [
        ["Status", "Quantity"],
        ["active", 10],
      ],
      dataValidations: [
        {
          type: "list",
          values: ["active", "inactive", "draft"],
          range: "A2:A100",
          showErrorMessage: true,
          errorTitle: "Invalid",
          errorMessage: "Pick from the list",
        },
        {
          type: "whole",
          operator: "between",
          formula1: "0",
          formula2: "1000",
          range: "B2:B100",
        },
      ],
    },
  ],
});
```

### Hyperlinks

```ts
const buffer = await writeXlsx({
  sheets: [
    {
      name: "Links",
      rows: [["Visit Google", "Go to Sheet2"]],
      cells: new Map([
        [
          "0,0",
          {
            value: "Visit Google",
            type: "string",
            hyperlink: { target: "https://google.com", tooltip: "Open Google" },
          },
        ],
        [
          "0,1",
          {
            value: "Go to Sheet2",
            type: "string",
            hyperlink: { target: "", location: "Sheet2!A1" },
          },
        ],
      ]),
    },
  ],
});
```

### Streaming

Process large files row-by-row without loading everything into memory:

```ts
import { streamXlsxRows, XlsxStreamWriter } from "hucre/xlsx";

// Stream read — async generator yields rows one at a time
for await (const row of streamXlsxRows(buffer)) {
  console.log(row.index, row.values);
}

// Stream write — add rows incrementally
const writer = new XlsxStreamWriter({
  name: "BigData",
  columns: [{ header: "ID" }, { header: "Value" }],
  freezePane: { rows: 1 },
});
for (let i = 0; i < 100_000; i++) {
  writer.addRow([i + 1, Math.random()]);
}
const buffer = await writer.finish();
```

### ODS (OpenDocument)

```ts
import { readOds, writeOds } from "hucre/ods";

const wb = await readOds(buffer);
const ods = await writeOds({ sheets: [{ name: "Sheet1", rows: [["Hello", 42]] }] });
```

### Round-trip Preservation

Open, modify, save — without losing charts, macros, or features hucre doesn't natively handle:

```ts
import { openXlsx, saveXlsx } from "hucre/xlsx";

const workbook = await openXlsx(buffer);
workbook.sheets[0].rows[0][0] = "Updated!";
const output = await saveXlsx(workbook); // Charts, VBA, themes preserved
```

### Unified API

Auto-detect format and work with simple helpers:

```ts
import { read, write, readObjects, writeObjects } from "hucre";

// Auto-detect XLSX vs ODS
const wb = await read(buffer);

// Quick: file → array of objects
const products = await readObjects<{ name: string; price: number }>(buffer);

// Quick: objects → XLSX
const xlsx = await writeObjects(products, { sheetName: "Products" });
```

### CLI

```bash
npx hucre convert input.xlsx output.csv
npx hucre convert input.csv output.xlsx
npx hucre inspect file.xlsx
npx hucre inspect file.xlsx --sheet 0
npx hucre validate data.xlsx --schema schema.json
```

### Sheet Operations

Manipulate sheet data in memory:

```ts
import { insertRows, deleteRows, cloneSheet, moveSheet } from "hucre";

insertRows(sheet, 5, 3); // Insert 3 rows at position 5
deleteRows(sheet, 0, 1); // Delete first row
const copy = cloneSheet(sheet, "Copy"); // Deep clone
moveSheet(workbook, 0, 2); // Reorder sheets
```

### HTML & Markdown Export

```ts
import { toHtml, toMarkdown } from "hucre";

const html = toHtml(workbook.sheets[0], {
  headerRow: true,
  styles: true,
  classes: true,
});

const md = toMarkdown(workbook.sheets[0]);
// | Name   | Price  | Stock |
// |--------|-------:|------:|
// | Widget |   9.99 |   142 |
```

### Number Format Renderer

```ts
import { formatValue } from "hucre";

formatValue(1234.5, "#,##0.00"); // "1,234.50"
formatValue(0.15, "0%"); // "15%"
formatValue(44197, "yyyy-mm-dd"); // "2021-01-01"
formatValue(1234, "$#,##0"); // "$1,234"
formatValue(0.333, "# ?/?"); // "1/3"
```

### Cell Utilities

```ts
import { parseCellRef, cellRef, colToLetter, rangeRef } from "hucre";

parseCellRef("AA15"); // { row: 14, col: 26 }
cellRef(14, 26); // "AA15"
colToLetter(26); // "AA"
rangeRef(0, 0, 9, 3); // "A1:D10"
```

### Builder API

Fluent method-chaining interface:

```ts
import { WorkbookBuilder } from "hucre";

const xlsx = await WorkbookBuilder.create()
  .addSheet("Products")
  .columns([
    { header: "Name", key: "name", autoWidth: true },
    { header: "Price", key: "price", numFmt: "$#,##0.00" },
  ])
  .row(["Widget", 9.99])
  .row(["Gadget", 24.5])
  .freeze(1)
  .done()
  .build();
```

### Template Engine

Fill `{{placeholders}}` in existing XLSX templates:

```ts
import { openXlsx, saveXlsx, fillTemplate } from "hucre";

const workbook = await openXlsx(templateBuffer);
fillTemplate(workbook, {
  company: "Acme Inc",
  date: new Date(),
  total: 12500,
});
const output = await saveXlsx(workbook);
```

### Object Shorthand (XLSX / ODS)

Skip the `wb.sheets[0].rows[0] as headers, slice(1) as data` boilerplate — return objects directly, mirror of `parseCsvObjects`:

```ts
import { readXlsxObjects, writeXlsxObjects } from "hucre/xlsx";
import { readOdsObjects, writeOdsObjects } from "hucre/ods";

const { data, headers } = await readXlsxObjects(buffer, {
  sheet: 0, // index or name (default: 0)
  headerRow: 0, // 0-based (default: 0)
  skipEmptyRows: true,
  transformHeader: (h) => h.toLowerCase().replace(/ /g, "_"),
  transformValue: (v, header) => (header === "price" ? Number(v) : v),
});

// Symmetric write — headers come from the first object's keys when omitted
const xlsx = await writeXlsxObjects(
  [
    { Name: "Widget", Price: 9.99 },
    { Name: "Gadget", Price: 24.5 },
  ],
  { sheetName: "Products" },
);
```

### JSON / NDJSON

```ts
import {
  parseJson,
  parseNdjson,
  writeJson,
  writeNdjson,
  workbookToJson,
  NdjsonStreamWriter,
  readNdjsonStream,
} from "hucre/json";

// Read — top-level array, { products: [...] } shape, or single object
const { data, headers } = parseJson(jsonString);

// Pick rows from a deeper path
parseJson(text, { rowsAt: "data.rows" });

// Flatten nested objects with dot-path keys (default: true)
parseJson('[{"sku":"P1","pricing":{"cost":100}}]');
// → data: [{ sku: "P1", "pricing.cost": 100 }]

// NDJSON / JSON Lines — one object per line
const out = parseNdjson(ndjsonText, {
  onError: (line, ln) => console.warn(`bad line ${ln}`), // skip + report
});

// Round-trip a workbook (single sheet → array, multi-sheet → { Sheet: [...] })
import { readXlsx } from "hucre/xlsx";
const wb = await readXlsx(buffer);
const json = workbookToJson(wb, { pretty: true });

// Streaming write — works in Cloudflare Workers / Deno / Node 18+
const writer = new NdjsonStreamWriter();
for await (const row of source) writer.write(row);
writer.end();
return new Response(writer.toStream(), {
  headers: { "content-type": "application/x-ndjson" },
});

// Streaming read
for await (const row of readNdjsonStream(request.body!)) {
  console.log(row);
}
```

### XML

Read and write tabular XML — product feeds (GS1 GDSN, Trendyol, marketplace exports), ERP dumps (SAP B1, Logo GO, Netsis), CRM catalogs. SAX-based: 50–200 MB feeds don't load into memory.

```ts
import { readXml, writeXml } from "hucre/xml";

// Auto-detects the most-frequently-repeating direct child of root as the row tag
const { data, headers, rowTag } = readXml(`
  <Catalog>
    <Product code="P1">
      <Name>Oak</Name>
      <Pricing currency="USD">
        <Cost>100</Cost>
        <Retail>180</Retail>
      </Pricing>
    </Product>
    <Product code="P2"><Name>Pine</Name></Product>
  </Catalog>
`);
// rowTag: "Product"
// data: [{ "@code": "P1", Name: "Oak", "Pricing.@currency": "USD",
//         "Pricing.Cost": "100", "Pricing.Retail": "180" }, ...]

// Override auto-detect with rowTag, strip namespace prefixes, control flatten
readXml(xml, { rowTag: "ns:Product", stripNamespaces: true, flatten: true });

// Write — @-keyed fields become XML attributes, dot-paths reconstruct elements
const xml = writeXml(
  [
    { "@code": "P1", Name: "Oak", "Pricing.Cost": 100 },
    { "@code": "P2", Name: "Pine", "Pricing.Cost": 90 },
  ],
  { rootTag: "Catalog", rowTag: "Product", pretty: true },
);
```

### JSON Export (legacy)

```ts
import { toJson } from "hucre";

toJson(sheet, { format: "objects" }); // [{Name:"Widget", Price:9.99}, ...]
toJson(sheet, { format: "columns" }); // {Name:["Widget"], Price:[9.99]}
toJson(sheet, { format: "arrays" }); // {headers:[...], data:[[...]]}
```

For new code prefer `writeJson` / `workbookToJson` from `hucre/json` — same result, consistent with `parseJson`/`parseNdjson`/`writeNdjson`.

### CSV

```ts
import { parseCsv, parseCsvObjects, writeCsv, detectDelimiter } from "hucre/csv";

// Parse — auto-detects delimiter, handles RFC 4180 edge cases
const rows = parseCsv(csvString, { typeInference: true });

// Parse with headers — returns typed objects
const { data, headers } = parseCsvObjects(csvString, { header: true });

// Write
const csv = writeCsv(rows, { delimiter: ";", bom: true });

// Detect delimiter
detectDelimiter(csvString); // "," or ";" or "\t" or "|"
```

### Schema Validation

Validate imported data with type coercion, pattern matching, and error collection:

```ts
import { validateWithSchema } from "hucre";
import { parseCsv } from "hucre/csv";

const rows = parseCsv(csvString);

const result = validateWithSchema(
  rows,
  {
    "Product Name": { type: "string", required: true },
    Price: { type: "number", required: true, min: 0 },
    SKU: { type: "string", pattern: /^[A-Z]{3}-\d{4}$/ },
    Stock: { type: "integer", min: 0, default: 0 },
    Status: { type: "string", enum: ["active", "inactive", "draft"] },
  },
  { headerRow: 1 },
);

console.log(result.data); // Validated & coerced objects
console.log(result.errors); // [{ row: 3, field: "Price", message: "...", value: "abc" }]
```

Schema field options:

| Option        | Type                                                       | Description                             |
| ------------- | ---------------------------------------------------------- | --------------------------------------- |
| `type`        | `"string" \| "number" \| "integer" \| "boolean" \| "date"` | Target type (with coercion)             |
| `required`    | `boolean`                                                  | Reject null/empty values                |
| `pattern`     | `RegExp`                                                   | Regex validation (strings)              |
| `min`         | `number`                                                   | Min value (numbers) or length (strings) |
| `max`         | `number`                                                   | Max value (numbers) or length (strings) |
| `enum`        | `unknown[]`                                                | Allowed values                          |
| `default`     | `unknown`                                                  | Default for null/empty                  |
| `validate`    | `(v) => boolean \| string`                                 | Custom validator                        |
| `transform`   | `(v) => unknown`                                           | Post-validation transform               |
| `column`      | `string`                                                   | Column header name                      |
| `columnIndex` | `number`                                                   | Column index (0-based)                  |

### Date Utilities

Timezone-safe Excel date serial number conversion:

```ts
import { serialToDate, dateToSerial, isDateFormat, formatDate } from "hucre";

serialToDate(44197); // 2021-01-01T00:00:00.000Z
dateToSerial(new Date("2021-01-01")); // 44197
isDateFormat("yyyy-mm-dd"); // true
isDateFormat("#,##0.00"); // false
formatDate(new Date(), "yyyy-mm-dd"); // "2026-03-24"
```

Handles the Lotus 1-2-3 bug (serial 60), 1900/1904 date systems, and time fractions correctly.

## Platform Support

hucre works everywhere — no Node.js APIs (`fs`, `crypto`, `Buffer`) in core.

| Runtime               | Status       |
| --------------------- | ------------ |
| Node.js 18+           | Full support |
| Deno                  | Full support |
| Bun                   | Full support |
| Modern browsers       | Full support |
| Cloudflare Workers    | Full support |
| Vercel Edge Functions | Full support |
| Web Workers           | Full support |

## Architecture

```
hucre (~37 KB gzipped)
├── zip/            Zero-dep DEFLATE/inflate + ZIP read/write
├── xml/            SAX parser + XML writer (CSP-compliant, no eval)
├── xlsx/
│   ├── reader      Shared strings, styles, worksheets, relationships
│   ├── writer      Styles, shared strings, drawing, tables, comments
│   ├── roundtrip   Open → modify → save with preservation
│   ├── stream-*    Streaming reader (AsyncGenerator) + writer
│   └── auto-width  Font-aware column width calculation
├── ods/            OpenDocument Spreadsheet read/write
├── csv/            RFC 4180 parser/writer + streaming
├── export/         HTML, Markdown, JSON, TSV output + HTML import
├── hucre           Unified read/write API, format auto-detect
├── builder         Fluent WorkbookBuilder / SheetBuilder API
├── template        {{placeholder}} template engine
├── sheet-ops       Insert/delete/move/sort/find/replace, clone, copy
├── cell-utils      parseCellRef, colToLetter, parseRange, isInRange
├── image           imageFromBase64 utility
├── worker          Web Worker serialization helpers
├── _date           Timezone-safe serial ↔ Date, Lotus bug, 1900/1904
├── _format         Number format renderer (locale-aware)
├── _schema         Schema validation, type coercion, error collection
└── cli             Convert, inspect, validate (citty + consola)
```

Zero dependencies. Pure TypeScript. The ZIP engine uses `CompressionStream`/`DecompressionStream` Web APIs with a pure TS fallback.

## API Reference

### High-level

| Function                       | Description                                       |
| ------------------------------ | ------------------------------------------------- |
| `read(input, options?)`        | Auto-detect format (XLSX/ODS), returns `Workbook` |
| `write(options)`               | Write XLSX or ODS (via `format` option)           |
| `readObjects(input, options?)` | File → array of objects (first row = headers)     |
| `writeObjects(data, options?)` | Objects → XLSX/ODS                                |

### XLSX

| Function                           | Description                                            |
| ---------------------------------- | ------------------------------------------------------ |
| `readXlsx(input, options?)`        | Parse XLSX from `Uint8Array \| ArrayBuffer`            |
| `writeXlsx(options)`               | Generate XLSX, returns `Uint8Array`                    |
| `readXlsxObjects(input, options?)` | Read sheet as `{ data, headers }` — mirror of CSV      |
| `writeXlsxObjects(data, options?)` | Write objects to XLSX (auto-derives headers from keys) |
| `openXlsx(input, options?)`        | Open for round-trip (preserves unknown parts)          |
| `saveXlsx(workbook)`               | Save round-trip workbook back to XLSX                  |
| `streamXlsxRows(input, options?)`  | AsyncGenerator yielding rows one at a time             |
| `XlsxStreamWriter`                 | Class for incremental row-by-row XLSX writing          |

### ODS

| Function                          | Description                          |
| --------------------------------- | ------------------------------------ |
| `readOds(input, options?)`        | Parse ODS (OpenDocument Spreadsheet) |
| `writeOds(options)`               | Generate ODS                         |
| `readOdsObjects(input, options?)` | Read sheet as `{ data, headers }`    |
| `writeOdsObjects(data, options?)` | Write objects to ODS                 |
| `streamOdsRows(input)`            | AsyncGenerator yielding ODS rows     |

### CSV

| Function                           | Description                                  |
| ---------------------------------- | -------------------------------------------- |
| `parseCsv(input, options?)`        | Parse CSV string → `CellValue[][]`           |
| `parseCsvObjects(input, options?)` | Parse CSV with headers → `{ data, headers }` |
| `writeCsv(rows, options?)`         | Write `CellValue[][]` → CSV string           |
| `writeCsvObjects(data, options?)`  | Write objects → CSV string                   |
| `detectDelimiter(input)`           | Auto-detect delimiter character              |
| `streamCsvRows(input, options?)`   | Generator yielding CSV rows                  |
| `CsvStreamWriter`                  | Class for incremental CSV writing            |
| `writeTsv(rows, options?)`         | Write TSV (tab-separated)                    |
| `fetchCsv(url, options?)`          | Fetch and parse CSV from URL                 |

### JSON

| Function                          | Description                                                    |
| --------------------------------- | -------------------------------------------------------------- |
| `parseJson(input, options?)`      | Parse JSON string/Uint8Array → `{ data, headers }`             |
| `parseValue(value, options?)`     | Same on already-parsed JSON                                    |
| `parseNdjson(input, options?)`    | Parse NDJSON / JSON Lines (`onError` skips invalid)            |
| `writeJson(data, options?)`       | Serialize rows to a JSON string                                |
| `writeNdjson(data, options?)`     | Serialize rows to NDJSON, one object per line                  |
| `workbookToJson(wb, options?)`    | Convert a `Workbook` to JSON (single-sheet array or per-sheet) |
| `readNdjsonStream(stream, opts?)` | Async generator over a `ReadableStream<Uint8Array>`            |
| `NdjsonStreamWriter`              | Incremental writer with `toStream(): ReadableStream`           |

### XML

| Function                   | Description                                              |
| -------------------------- | -------------------------------------------------------- |
| `readXml(input, options?)` | SAX-based XML reader, auto-detects repeating row element |
| `writeXml(data, options?)` | Serialize rows to XML; `@`-keys → attributes             |

### Sheet Operations

| Function                                | Description                     |
| --------------------------------------- | ------------------------------- |
| `insertRows(sheet, index, count)`       | Insert rows, shift down         |
| `deleteRows(sheet, index, count)`       | Delete rows, shift up           |
| `insertColumns(sheet, index, count)`    | Insert columns, shift right     |
| `deleteColumns(sheet, index, count)`    | Delete columns, shift left      |
| `moveRows(sheet, from, count, to)`      | Move rows                       |
| `cloneSheet(sheet, name)`               | Deep clone a sheet              |
| `copySheetToWorkbook(sheet, wb, name?)` | Copy sheet between workbooks    |
| `copyRange(sheet, source, target)`      | Copy cell range within sheet    |
| `moveSheet(wb, from, to)`               | Reorder sheets                  |
| `removeSheet(wb, index)`                | Remove a sheet                  |
| `sortRows(sheet, col, order?)`          | Sort rows by column             |
| `findCells(sheet, predicate)`           | Find cells by value or function |
| `replaceCells(sheet, find, replace)`    | Find and replace values         |

### Export

| Function                      | Description                                      |
| ----------------------------- | ------------------------------------------------ |
| `toHtml(sheet, options?)`     | HTML `<table>` with styles, a11y, dark/light CSS |
| `toMarkdown(sheet, options?)` | Markdown table with auto-alignment               |
| `toJson(sheet, options?)`     | JSON (objects, arrays, or columns format)        |
| `fromHtml(html, options?)`    | Parse HTML table string → Sheet                  |
| `writeTsv(rows, options?)`    | Write TSV (tab-separated)                        |

### Builder

| Function                       | Description                             |
| ------------------------------ | --------------------------------------- |
| `WorkbookBuilder.create()`     | Fluent API for building workbooks       |
| `fillTemplate(workbook, data)` | Replace `{{placeholders}}` in templates |

### Formatting & Utilities

| Function                                     | Description                              |
| -------------------------------------------- | ---------------------------------------- |
| `formatValue(value, numFmt, options?)`       | Apply Excel number format (locale-aware) |
| `validateWithSchema(rows, schema, options?)` | Validate & coerce data with schema       |
| `serialToDate(serial, is1904?)`              | Excel serial → Date (UTC)                |
| `dateToSerial(date, is1904?)`                | Date → Excel serial                      |
| `isDateFormat(numFmt)`                       | Check if format string is date           |
| `formatDate(date, format)`                   | Format Date with Excel format string     |
| `parseCellRef(ref)`                          | "AA15" → `{ row: 14, col: 26 }`          |
| `cellRef(row, col)`                          | `(14, 26)` → "AA15"                      |
| `colToLetter(col)`                           | `26` → "AA"                              |
| `rangeRef(r1, c1, r2, c2)`                   | `(0,0,9,3)` → "A1:D10"                   |

### Web Worker Helpers

| Function                    | Description                                                          |
| --------------------------- | -------------------------------------------------------------------- |
| `serializeWorkbook(wb)`     | Convert Workbook for `postMessage` (Maps → objects, Dates → strings) |
| `deserializeWorkbook(data)` | Restore Workbook from serialized form                                |
| `WORKER_SAFE_FUNCTIONS`     | List of all hucre functions safe for Web Workers (all of them)       |

## Development

```sh
pnpm install
pnpm dev          # vitest watch
pnpm test         # lint + typecheck + test
pnpm build        # obuild (minified, tree-shaken)
pnpm lint:fix     # oxlint + oxfmt
pnpm typecheck    # tsgo
```

## Contributing

Contributions are welcome! Please [open an issue](https://github.com/productdevbook/hucre/issues) or submit a PR.

127 of 135 tracked features are implemented. See the [issue tracker](https://github.com/productdevbook/hucre/issues) for the roadmap.

### Roadmap

**Upcoming Engine Features:**

- Chart creation (bar, line, pie, scatter, area + subtypes)
- XLS BIFF8 read (legacy Excel 97-2003)
- XLSB binary format read
- Formula evaluation engine
- File encryption/decryption (AES-256, MS-OFFCRYPTO)
- Pivot table creation
- Threaded comments (Excel 365+)
- Checkboxes (Excel 2024+)
- VBA/macro injection
- Slicers & timeline filters
- R1C1 notation support
- Accessibility helpers (WCAG 2.1 AA)

## Alternatives

Looking for a different approach? These libraries may fit your use case:

- **[SheetJS (xlsx)](https://github.com/SheetJS/sheetjs)** — The most popular spreadsheet library. Feature-rich but large bundle (~300 KB), removed from npm (CDN-only), styling requires Pro license.
- **[ExcelJS](https://github.com/exceljs/exceljs)** — Read/write/stream XLSX with styling. Mature but has 12 dependencies (some with CVEs), CJS-only, no ESM.
- **[xlsx-js-style](https://github.com/gitbrent/xlsx-js-style)** — SheetJS fork that adds cell styling. Same bundle size and limitations as SheetJS.
- **[xlsmith](https://github.com/ChronicStone/xlsmith)** — Schema-driven Excel report builder with typed column definitions, formula helpers, conditional styles, and summary rows. Great for structured report generation.
- **[xlsx-populate](https://github.com/dtjohnson/xlsx-populate)** — Template-based XLSX manipulation. Good for filling existing templates, limited write-from-scratch support.
- **[better-xlsx](https://github.com/nichenqin/better-xlsx)** — Lightweight XLSX writer with styling. Write-only, no read support.

## License

[MIT](./LICENSE) — Made by [productdevbook](https://github.com/productdevbook)

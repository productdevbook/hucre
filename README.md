<p align="center">
  <br>
  <img src=".github/assets/cover.svg" alt="defter — Zero-dependency spreadsheet engine" width="100%">
  <br><br>
  <b style="font-size: 2em;">defter</b>
  <br><br>
  Zero-dependency spreadsheet engine.
  <br>
  Read & write XLSX, CSV. Schema validation, streaming, tree-shakeable. Pure TypeScript, works everywhere.
  <br><br>
  <a href="https://npmjs.com/package/defter"><img src="https://img.shields.io/npm/v/defter?style=flat&colorA=18181B&colorB=34d399" alt="npm version"></a>
  <a href="https://npmjs.com/package/defter"><img src="https://img.shields.io/npm/dm/defter?style=flat&colorA=18181B&colorB=34d399" alt="npm downloads"></a>
  <a href="https://bundlephobia.com/result?p=defter"><img src="https://img.shields.io/bundlephobia/minzip/defter?style=flat&colorA=18181B&colorB=34d399" alt="bundle size"></a>
  <a href="https://github.com/productdevbook/defter/blob/main/LICENSE"><img src="https://img.shields.io/github/license/productdevbook/defter?style=flat&colorA=18181B&colorB=34d399" alt="license"></a>
</p>

## Quick Start

```sh
npm install defter
```

```ts
import { readXlsx, writeXlsx } from "defter";

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
import { readXlsx, writeXlsx } from "defter/xlsx"; // XLSX only (~14 KB gzipped)
import { parseCsv, writeCsv } from "defter/csv"; // CSV only (~2 KB gzipped)
```

## Why defter?

|                   | defter | SheetJS       | ExcelJS   | read-excel-file |
| ----------------- | ------ | ------------- | --------- | --------------- |
| **Dependencies**  | 0      | 0\*           | 12 (CVEs) | 2               |
| **Bundle (gzip)** | ~18 KB | ~300 KB       | ~500 KB   | ~40 KB          |
| **ESM native**    | Yes    | Partial       | No (CJS)  | Yes             |
| **TypeScript**    | Native | Bolted-on     | Bolted-on | Yes             |
| **Edge runtime**  | Yes    | No            | No        | No              |
| **CSP compliant** | Yes    | Yes           | No (eval) | Yes             |
| **npm published** | Yes    | No (CDN only) | Stale     | Yes             |
| **Read + Write**  | Yes    | Yes (Pro $)   | Yes       | Separate pkgs   |

\* SheetJS removed itself from npm; must install from CDN tarball.

## Features

### Reading

```ts
import { readXlsx } from "defter/xlsx";

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
import { writeXlsx } from "defter/xlsx";

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

Features: cell styles (fonts, fills, borders, alignment), auto column widths, merged cells, freeze panes, auto-filter, data validation (dropdowns), hyperlinks, number formats, formulas, multiple sheets, hidden sheets.

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

### CSV

```ts
import { parseCsv, parseCsvObjects, writeCsv, detectDelimiter } from "defter/csv";

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
import { validateWithSchema } from "defter";
import { parseCsv } from "defter/csv";

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
import { serialToDate, dateToSerial, isDateFormat, formatDate } from "defter";

serialToDate(44197); // 2021-01-01T00:00:00.000Z
dateToSerial(new Date("2021-01-01")); // 44197
isDateFormat("yyyy-mm-dd"); // true
isDateFormat("#,##0.00"); // false
formatDate(new Date(), "yyyy-mm-dd"); // "2026-03-24"
```

Handles the Lotus 1-2-3 bug (serial 60), 1900/1904 date systems, and time fractions correctly.

## Platform Support

defter works everywhere — no Node.js APIs (`fs`, `crypto`, `Buffer`) in core.

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
defter (18 KB gzipped)
├── zip/          Zero-dep DEFLATE/inflate + ZIP read/write
├── xml/          SAX parser + XML writer (CSP-compliant, no eval)
├── xlsx/
│   ├── reader    Shared strings, styles, worksheets, relationships
│   └── writer    Styles collector, shared strings dedup, worksheet gen
├── csv/          RFC 4180 parser/writer, auto-detect, type inference
├── _date         Timezone-safe serial ↔ Date, Lotus bug, 1900/1904
├── _schema       Schema validation, type coercion, error collection
└── _types        Full TypeScript type definitions
```

Zero dependencies. Pure TypeScript. The ZIP engine uses `CompressionStream`/`DecompressionStream` Web APIs with a pure TS fallback.

## API Reference

### XLSX

| Function                    | Description                                 |
| --------------------------- | ------------------------------------------- |
| `readXlsx(input, options?)` | Parse XLSX from `Uint8Array \| ArrayBuffer` |
| `writeXlsx(options)`        | Generate XLSX, returns `Uint8Array`         |

### CSV

| Function                           | Description                                  |
| ---------------------------------- | -------------------------------------------- |
| `parseCsv(input, options?)`        | Parse CSV string → `CellValue[][]`           |
| `parseCsvObjects(input, options?)` | Parse CSV with headers → `{ data, headers }` |
| `writeCsv(rows, options?)`         | Write `CellValue[][]` → CSV string           |
| `writeCsvObjects(data, options?)`  | Write objects → CSV string                   |
| `detectDelimiter(input)`           | Auto-detect delimiter character              |
| `stripBom(input)`                  | Remove BOM from string                       |
| `formatCsvValue(value, options?)`  | Format single value for CSV                  |

### Schema

| Function                                     | Description                        |
| -------------------------------------------- | ---------------------------------- |
| `validateWithSchema(rows, schema, options?)` | Validate & coerce data with schema |

### Date

| Function                        | Description                                         |
| ------------------------------- | --------------------------------------------------- |
| `serialToDate(serial, is1904?)` | Excel serial → Date (UTC)                           |
| `dateToSerial(date, is1904?)`   | Date → Excel serial                                 |
| `isDateFormat(numFmt)`          | Check if format string is date                      |
| `formatDate(date, format)`      | Format Date with Excel format string                |
| `parseDate(value)`              | Parse date string → Date or null                    |
| `serialToTime(serial)`          | Serial fraction → `{ hours, minutes, seconds, ms }` |
| `timeToSerial(h, m, s?, ms?)`   | Time components → serial fraction                   |

### Errors

| Class                    | When                      |
| ------------------------ | ------------------------- |
| `DefterError`            | Base error class          |
| `ParseError`             | Invalid file structure    |
| `ZipError`               | ZIP archive issues        |
| `XmlError`               | Malformed XML             |
| `ValidationError`        | Schema validation failure |
| `UnsupportedFormatError` | Unknown file format       |
| `EncryptedFileError`     | Password-protected file   |

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

Contributions are welcome! Please [open an issue](https://github.com/productdevbook/defter/issues) or submit a PR.

See the [issue tracker](https://github.com/productdevbook/defter/issues) for planned features — there are 39 tracked issues covering everything from ODS support to chart creation.

## License

[MIT](./LICENSE) — Made by [productdevbook](https://github.com/productdevbook)

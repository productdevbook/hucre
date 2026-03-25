#!/usr/bin/env node

// ── CLI Tool ────────────────────────────────────────────────────────
// defter convert input.xlsx output.csv
// defter convert input.csv output.xlsx
// defter convert input.xlsx output.ods
// defter inspect file.xlsx
// defter inspect file.xlsx --sheet 0
// defter validate data.xlsx --schema schema.json
// ─────────────────────────────────────────────────────────────────────

import { defineCommand, runMain } from "citty";
import { consola } from "consola";
import { readFileSync, writeFileSync } from "node:fs";
import { extname } from "node:path";
import { readXlsx } from "./xlsx/reader";
import { writeXlsx } from "./xlsx/writer";
import { readOds } from "./ods/reader";
import { writeOds } from "./ods/writer";
import { parseCsv } from "./csv/reader";
import { writeCsv } from "./csv/writer";
import { validateWithSchema } from "./_schema";
import type { Workbook, CellValue, WriteOptions, SchemaDefinition } from "./_types";

// ── Helpers ─────────────────────────────────────────────────────────

type Format = "xlsx" | "ods" | "csv";

function detectFormatFromExtension(filePath: string): Format {
  const ext = extname(filePath).toLowerCase();
  switch (ext) {
    case ".xlsx":
      return "xlsx";
    case ".ods":
      return "ods";
    case ".csv":
    case ".tsv":
      return "csv";
    default:
      consola.error(`Unsupported file extension: ${ext}`);
      process.exit(1);
  }
}

async function readFile(filePath: string): Promise<Workbook> {
  const format = detectFormatFromExtension(filePath);
  const data = readFileSync(filePath);
  const input = new Uint8Array(data);

  switch (format) {
    case "xlsx":
      return readXlsx(input);
    case "ods":
      return readOds(input);
    case "csv": {
      const text = new TextDecoder("utf-8").decode(input);
      const rows = parseCsv(text);
      return {
        sheets: [{ name: "Sheet1", rows }],
      };
    }
  }
}

function formatCellValue(value: CellValue): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString();
  return String(value);
}

// ── Convert Command ─────────────────────────────────────────────────

const convertCommand = defineCommand({
  meta: {
    name: "convert",
    description: "Convert between spreadsheet formats",
  },
  args: {
    input: {
      type: "positional",
      description: "Input file path",
      required: true,
    },
    output: {
      type: "positional",
      description: "Output file path",
      required: true,
    },
  },
  async run({ args }) {
    const inputPath = args.input as string;
    const outputPath = args.output as string;
    const outputFormat = detectFormatFromExtension(outputPath);

    consola.start(`Reading ${inputPath}...`);
    const workbook = await readFile(inputPath);
    consola.success(`Read ${workbook.sheets.length} sheet(s)`);

    consola.start(`Writing ${outputPath}...`);

    if (outputFormat === "csv") {
      // CSV: use first sheet only
      const sheet = workbook.sheets[0];
      if (!sheet) {
        consola.error("No sheets found in input file");
        process.exit(1);
      }
      const headers = sheet.rows[0]?.map((v) => formatCellValue(v));
      const dataRows = sheet.rows.slice(headers ? 1 : 0);
      const csv = writeCsv(dataRows, {
        headers: headers ?? undefined,
      });
      writeFileSync(outputPath, csv, "utf-8");
    } else {
      // XLSX or ODS
      const writeOptions: WriteOptions = {
        sheets: workbook.sheets.map((sheet) => ({
          name: sheet.name,
          rows: sheet.rows,
        })),
        properties: workbook.properties,
      };

      let output: Uint8Array;
      if (outputFormat === "ods") {
        output = await writeOds(writeOptions);
      } else {
        output = await writeXlsx(writeOptions);
      }

      writeFileSync(outputPath, output);
    }

    consola.success(`Written to ${outputPath}`);
  },
});

// ── Inspect Command ─────────────────────────────────────────────────

const inspectCommand = defineCommand({
  meta: {
    name: "inspect",
    description: "Inspect a spreadsheet file",
  },
  args: {
    file: {
      type: "positional",
      description: "File to inspect",
      required: true,
    },
    sheet: {
      type: "string",
      description: "Sheet index to show detailed data (0-based)",
    },
  },
  async run({ args }) {
    const filePath = args.file as string;

    consola.start(`Inspecting ${filePath}...`);
    const workbook = await readFile(filePath);

    consola.info(`Sheets: ${workbook.sheets.length}`);

    for (let i = 0; i < workbook.sheets.length; i++) {
      const sheet = workbook.sheets[i]!;
      const rowCount = sheet.rows.length;
      const colCount = sheet.rows.reduce((max, row) => Math.max(max, row.length), 0);

      // Count cell types
      const typeCounts: Record<string, number> = {};
      for (const row of sheet.rows) {
        for (const cell of row) {
          let type: string;
          if (cell === null || cell === undefined) type = "empty";
          else if (typeof cell === "string") type = "string";
          else if (typeof cell === "number") type = "number";
          else if (typeof cell === "boolean") type = "boolean";
          else if (cell instanceof Date) type = "date";
          else type = "unknown";

          typeCounts[type] = (typeCounts[type] ?? 0) + 1;
        }
      }

      const typeStr = Object.entries(typeCounts)
        .map(([t, c]) => `${t}: ${c}`)
        .join(", ");

      consola.log(`  [${i}] "${sheet.name}" - ${rowCount} rows x ${colCount} cols (${typeStr})`);
    }

    if (workbook.properties) {
      const props = workbook.properties;
      if (props.title) consola.log(`  Title: ${props.title}`);
      if (props.creator) consola.log(`  Creator: ${props.creator}`);
      if (props.created) consola.log(`  Created: ${props.created.toISOString()}`);
    }

    // Show detailed sheet data if --sheet is specified
    if (args.sheet !== undefined) {
      const sheetIdx = Number(args.sheet);
      if (Number.isNaN(sheetIdx) || sheetIdx < 0 || sheetIdx >= workbook.sheets.length) {
        consola.error(
          `Invalid sheet index: ${args.sheet}. Valid range: 0-${workbook.sheets.length - 1}`,
        );
        process.exit(1);
      }

      const sheet = workbook.sheets[sheetIdx]!;
      consola.info(`\nSheet "${sheet.name}" (first 10 rows):`);

      const previewRows = sheet.rows.slice(0, 10);
      if (previewRows.length === 0) {
        consola.log("  (empty sheet)");
      } else {
        // Build column widths for formatting
        const maxCols = previewRows.reduce((max, row) => Math.max(max, row.length), 0);
        const colWidths: number[] = Array.from({ length: maxCols }, () => 0);

        const formatted = previewRows.map((row) =>
          Array.from({ length: maxCols }, (_, j) => {
            const val = j < row.length ? formatCellValue(row[j]!) : "";
            const str = val.length > 40 ? `${val.substring(0, 37)}...` : val;
            if (str.length > colWidths[j]!) colWidths[j] = str.length;
            return str;
          }),
        );

        for (let i = 0; i < formatted.length; i++) {
          const row = formatted[i]!;
          const line = row.map((val, j) => val.padEnd(colWidths[j]! + 2)).join("");
          consola.log(`  ${String(i).padStart(3)}| ${line}`);
        }

        if (sheet.rows.length > 10) {
          consola.log(`  ... and ${sheet.rows.length - 10} more rows`);
        }
      }
    }
  },
});

// ── Validate Command ────────────────────────────────────────────────

const validateCommand = defineCommand({
  meta: {
    name: "validate",
    description: "Validate a spreadsheet against a JSON schema",
  },
  args: {
    file: {
      type: "positional",
      description: "Spreadsheet file to validate",
      required: true,
    },
    schema: {
      type: "string",
      description: "Path to JSON schema file",
      required: true,
    },
    sheet: {
      type: "string",
      description: "Sheet index to validate (0-based, default: 0)",
      default: "0",
    },
  },
  async run({ args }) {
    const filePath = args.file as string;
    const schemaPath = args.schema as string;
    const sheetIdx = Number(args.sheet ?? "0");

    consola.start(`Validating ${filePath} with schema ${schemaPath}...`);

    // Read schema
    const schemaJson = readFileSync(schemaPath, "utf-8");
    let schema: SchemaDefinition;
    try {
      schema = JSON.parse(schemaJson) as SchemaDefinition;
    } catch {
      consola.error("Invalid JSON schema file");
      process.exit(1);
    }

    // Read spreadsheet
    const workbook = await readFile(filePath);

    if (sheetIdx < 0 || sheetIdx >= workbook.sheets.length) {
      consola.error(
        `Invalid sheet index: ${sheetIdx}. File has ${workbook.sheets.length} sheet(s)`,
      );
      process.exit(1);
    }

    const sheet = workbook.sheets[sheetIdx]!;
    const result = validateWithSchema(sheet.rows, schema, { headerRow: 1 });

    if (result.errors.length === 0) {
      consola.success(`Valid! ${result.data.length} row(s) passed validation.`);
    } else {
      consola.error(`Found ${result.errors.length} error(s) in ${result.data.length} row(s):`);
      for (const err of result.errors.slice(0, 20)) {
        consola.log(`  Row ${err.row}, Column "${err.column}": ${err.message}`);
      }
      if (result.errors.length > 20) {
        consola.log(`  ... and ${result.errors.length - 20} more errors`);
      }
      process.exit(1);
    }
  },
});

// ── Main Command ────────────────────────────────────────────────────

const main = defineCommand({
  meta: {
    name: "hucre",
    version: "0.0.1",
    description:
      "Spreadsheet Swiss Army knife. Convert, inspect, and validate XLSX, CSV, and ODS files.",
  },
  subCommands: {
    convert: convertCommand,
    inspect: inspectCommand,
    validate: validateCommand,
  },
});

runMain(main);

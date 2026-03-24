// ── Schema Validation Engine ─────────────────────────────────────────
//
// Validates and transforms parsed spreadsheet rows using a SchemaDefinition.
// Used after parsing rows from any format (XLSX, CSV, ODS).
// ─────────────────────────────────────────────────────────────────────

import type {
  CellValue,
  SchemaDefinition,
  SchemaField,
  SchemaFieldType,
  ValidationError as ValidationErrorType,
} from "./_types";
import { ValidationError } from "./errors";
import { serialToDate, parseDate } from "./_date";

/**
 * Validate and transform parsed rows using a schema definition.
 *
 * @param rows - Raw 2D array of cell values (first row may be headers)
 * @param schema - Schema definition mapping field names to validation rules
 * @param options - Validation options
 * @returns Object with validated data and any errors
 */
export function validateWithSchema<T extends Record<string, unknown> = Record<string, unknown>>(
  rows: CellValue[][],
  schema: SchemaDefinition,
  options?: {
    headerRow?: number;
    skipEmptyRows?: boolean;
    errorMode?: "collect" | "throw";
  },
): { data: T[]; errors: ValidationErrorType[] } {
  const headerRowNum = options?.headerRow ?? 1;
  const skipEmptyRows = options?.skipEmptyRows ?? false;
  const errorMode = options?.errorMode ?? "collect";

  const fieldNames = Object.keys(schema);

  // No fields in schema → return empty data
  if (fieldNames.length === 0) {
    return { data: [], errors: [] };
  }

  // No rows → return empty result
  if (rows.length === 0) {
    return { data: [], errors: [] };
  }

  // ── Step 1: Build column index mapping ──────────────────────────

  // Extract headers from the header row (0-based index)
  const headerRowIndex = headerRowNum - 1;
  const headerRow =
    headerRowIndex >= 0 && headerRowIndex < rows.length ? rows[headerRowIndex]! : [];

  // Build a map: normalized header name → column index
  const headerMap = new Map<string, number>();
  for (let i = 0; i < headerRow.length; i++) {
    const raw = headerRow[i];
    if (raw != null) {
      const normalized = String(raw).trim().toLowerCase();
      if (normalized !== "") {
        headerMap.set(normalized, i);
      }
    }
  }

  // Resolve each schema field to a column index
  const fieldColumnMap = new Map<string, number>();
  const errors: ValidationErrorType[] = [];

  for (const fieldName of fieldNames) {
    const field = schema[fieldName]!;

    if (field.columnIndex != null) {
      // Explicit column index — use directly
      fieldColumnMap.set(fieldName, field.columnIndex);
    } else {
      // Look up by column name or field name
      const lookupName = (field.column ?? fieldName).trim().toLowerCase();
      const colIdx = headerMap.get(lookupName);

      if (colIdx != null) {
        fieldColumnMap.set(fieldName, colIdx);
      } else {
        // Column not found in headers
        // If the field is required, we need to report errors for every data row
        // We still continue; individual row validation will report missing values
        // Set to -1 as a sentinel for "column not found"
        fieldColumnMap.set(fieldName, -1);
      }
    }
  }

  // ── Step 2: Validate data rows ──────────────────────────────────

  const data: T[] = [];
  const dataStartIndex = headerRowIndex + 1;

  for (let rowIdx = dataStartIndex; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx]!;

    // Skip empty rows if requested
    if (skipEmptyRows && isRowEmpty(row)) {
      continue;
    }

    const record: Record<string, unknown> = {};

    for (const fieldName of fieldNames) {
      const field = schema[fieldName]!;
      const colIdx = fieldColumnMap.get(fieldName)!;

      // Get raw value
      const rawValue: CellValue = colIdx >= 0 && colIdx < row.length ? (row[colIdx] ?? null) : null;

      // 1-based row number for error reporting
      const displayRow = rowIdx + 1;
      const displayColumn = field.column ?? fieldName;

      // Check required
      if (field.required && isEmpty(rawValue)) {
        const err: ValidationErrorType = {
          row: displayRow,
          column: displayColumn,
          message: `Required field '${displayColumn}' is empty`,
          value: rawValue,
          field: fieldName,
        };
        errors.push(err);
        if (errorMode === "throw") {
          throw new ValidationError(err.message, [err]);
        }

        record[fieldName] = null;
        continue;
      }

      // If value is empty and not required, apply default or set null
      if (isEmpty(rawValue)) {
        if (field.default !== undefined) {
          record[fieldName] = field.default;
        } else {
          record[fieldName] = null;
        }
        continue;
      }

      // Type coercion
      let coerced: unknown = rawValue;
      let coercionError = false;

      if (field.type) {
        const result = coerceValue(rawValue, field.type, displayColumn);
        if (result.error) {
          const err: ValidationErrorType = {
            row: displayRow,
            column: displayColumn,
            message: result.error,
            value: rawValue,
            field: fieldName,
          };
          errors.push(err);
          if (errorMode === "throw") {
            throw new ValidationError(err.message, [err]);
          }
          coercionError = true;

          record[fieldName] = null;
          continue;
        }
        coerced = result.value;
      }

      if (coercionError) {
        continue;
      }

      // Pattern validation (strings only)
      if (field.pattern && typeof coerced === "string") {
        if (!field.pattern.test(coerced)) {
          const err: ValidationErrorType = {
            row: displayRow,
            column: displayColumn,
            message: `'${displayColumn}' does not match pattern`,
            value: rawValue,
            field: fieldName,
          };
          errors.push(err);
          if (errorMode === "throw") {
            throw new ValidationError(err.message, [err]);
          }

          record[fieldName] = null;
          continue;
        }
      }

      // Min/max validation
      if (field.min != null || field.max != null) {
        const minMaxErr = validateMinMax(
          coerced,
          field,
          displayColumn,
          displayRow,
          rawValue,
          fieldName,
        );
        if (minMaxErr) {
          errors.push(minMaxErr);
          if (errorMode === "throw") {
            throw new ValidationError(minMaxErr.message, [minMaxErr]);
          }

          record[fieldName] = null;
          continue;
        }
      }

      // Enum validation
      if (field.enum) {
        if (!field.enum.includes(coerced as never)) {
          const allowed = field.enum.map((v) => String(v)).join(", ");
          const err: ValidationErrorType = {
            row: displayRow,
            column: displayColumn,
            message: `'${displayColumn}' must be one of: ${allowed}`,
            value: rawValue,
            field: fieldName,
          };
          errors.push(err);
          if (errorMode === "throw") {
            throw new ValidationError(err.message, [err]);
          }

          record[fieldName] = null;
          continue;
        }
      }

      // Custom validate function
      if (field.validate) {
        const result = field.validate(coerced);
        if (result !== true) {
          const message =
            typeof result === "string" ? result : `Custom validation failed for '${displayColumn}'`;
          const err: ValidationErrorType = {
            row: displayRow,
            column: displayColumn,
            message,
            value: rawValue,
            field: fieldName,
          };
          errors.push(err);
          if (errorMode === "throw") {
            throw new ValidationError(err.message, [err]);
          }

          record[fieldName] = null;
          continue;
        }
      }

      // Transform function
      if (field.transform) {
        coerced = field.transform(coerced);
      }

      record[fieldName] = coerced;
    }

    data.push(record as T);
  }

  return { data, errors };
}

// ── Helpers ───────────────────────────────────────────────────────────

/** Check if a cell value is considered "empty" */
function isEmpty(value: CellValue): boolean {
  if (value === null || value === undefined) return true;
  if (typeof value === "string" && value.trim() === "") return true;
  return false;
}

/** Check if an entire row is empty */
function isRowEmpty(row: CellValue[]): boolean {
  if (row.length === 0) return true;
  return row.every(
    (cell) =>
      cell === null || cell === undefined || (typeof cell === "string" && cell.trim() === ""),
  );
}

/** Coerce a raw cell value to the target type. Returns { value } or { error }. */
function coerceValue(
  raw: CellValue,
  type: SchemaFieldType,
  columnName: string,
): { value: unknown; error?: undefined } | { value?: undefined; error: string } {
  switch (type) {
    case "string":
      return coerceToString(raw);
    case "number":
      return coerceToNumber(raw, columnName);
    case "integer":
      return coerceToInteger(raw, columnName);
    case "boolean":
      return coerceToBoolean(raw, columnName);
    case "date":
      return coerceToDate(raw, columnName);
    default:
      return { value: raw };
  }
}

function coerceToString(raw: CellValue): { value: string } {
  if (raw === null || raw === undefined) {
    return { value: "" };
  }
  if (typeof raw === "string") {
    return { value: raw.trim() };
  }
  if (typeof raw === "boolean") {
    return { value: String(raw) };
  }
  if (typeof raw === "number") {
    return { value: String(raw) };
  }
  if (raw instanceof Date) {
    return { value: raw.toISOString() };
  }
  return { value: String(raw) };
}

function coerceToNumber(
  raw: CellValue,
  columnName: string,
): { value: number; error?: undefined } | { value?: undefined; error: string } {
  if (typeof raw === "number") {
    return { value: raw };
  }
  if (typeof raw === "boolean") {
    return { value: raw ? 1 : 0 };
  }
  if (typeof raw === "string") {
    // Strip commas used as thousands separators
    const cleaned = raw.replace(/,/g, "").trim();
    if (cleaned === "") {
      // Empty string treated as null — should have been caught by required check
      return {
        error: `Expected number for '${columnName}', got ''`,
      };
    }
    const num = Number.parseFloat(cleaned);
    if (Number.isNaN(num)) {
      return {
        error: `Expected number for '${columnName}', got '${raw}'`,
      };
    }
    return { value: num };
  }
  return {
    error: `Expected number for '${columnName}', got '${String(raw)}'`,
  };
}

function coerceToInteger(
  raw: CellValue,
  columnName: string,
): { value: number; error?: undefined } | { value?: undefined; error: string } {
  if (typeof raw === "number") {
    // Allow .0 (e.g. 42.0 → 42) but reject actual decimals
    if (!Number.isFinite(raw)) {
      return {
        error: `Expected integer for '${columnName}', got '${raw}'`,
      };
    }
    if (raw % 1 !== 0) {
      return {
        error: `Expected integer for '${columnName}', got '${raw}'`,
      };
    }
    return { value: Math.trunc(raw) };
  }
  if (typeof raw === "string") {
    const cleaned = raw.replace(/,/g, "").trim();
    if (cleaned === "") {
      return {
        error: `Expected integer for '${columnName}', got ''`,
      };
    }
    const num = Number(cleaned);
    if (Number.isNaN(num)) {
      return {
        error: `Expected integer for '${columnName}', got '${raw}'`,
      };
    }
    if (num % 1 !== 0) {
      return {
        error: `Expected integer for '${columnName}', got '${raw}'`,
      };
    }
    return { value: Math.trunc(num) };
  }
  return {
    error: `Expected integer for '${columnName}', got '${String(raw)}'`,
  };
}

function coerceToBoolean(
  raw: CellValue,
  columnName: string,
): { value: boolean; error?: undefined } | { value?: undefined; error: string } {
  if (typeof raw === "boolean") {
    return { value: raw };
  }
  if (typeof raw === "number") {
    if (raw === 1) return { value: true };
    if (raw === 0) return { value: false };
    return {
      error: `Expected boolean for '${columnName}', got '${raw}'`,
    };
  }
  if (typeof raw === "string") {
    const lower = raw.trim().toLowerCase();
    if (lower === "true" || lower === "yes" || lower === "1") {
      return { value: true };
    }
    if (lower === "false" || lower === "no" || lower === "0") {
      return { value: false };
    }
    return {
      error: `Expected boolean for '${columnName}', got '${raw}'`,
    };
  }
  return {
    error: `Expected boolean for '${columnName}', got '${String(raw)}'`,
  };
}

function coerceToDate(
  raw: CellValue,
  columnName: string,
): { value: Date; error?: undefined } | { value?: undefined; error: string } {
  if (raw instanceof Date) {
    return { value: raw };
  }
  if (typeof raw === "number") {
    // Treat as Excel serial number
    return { value: serialToDate(raw) };
  }
  if (typeof raw === "string") {
    const parsed = parseDate(raw);
    if (parsed === null) {
      return {
        error: `Expected date for '${columnName}', got '${raw}'`,
      };
    }
    return { value: parsed };
  }
  return {
    error: `Expected date for '${columnName}', got '${String(raw)}'`,
  };
}

function validateMinMax(
  value: unknown,
  field: SchemaField,
  columnName: string,
  row: number,
  rawValue: unknown,
  fieldName: string,
): ValidationErrorType | null {
  if (typeof value === "number") {
    if (field.min != null && value < field.min) {
      return {
        row,
        column: columnName,
        message: `Value ${value} for '${columnName}' is below minimum ${field.min}`,
        value: rawValue,
        field: fieldName,
      };
    }
    if (field.max != null && value > field.max) {
      return {
        row,
        column: columnName,
        message: `Value ${value} for '${columnName}' exceeds maximum ${field.max}`,
        value: rawValue,
        field: fieldName,
      };
    }
  } else if (typeof value === "string") {
    if (field.min != null && value.length < field.min) {
      return {
        row,
        column: columnName,
        message: `'${columnName}' length ${value.length} is below minimum ${field.min}`,
        value: rawValue,
        field: fieldName,
      };
    }
    if (field.max != null && value.length > field.max) {
      return {
        row,
        column: columnName,
        message: `'${columnName}' length ${value.length} exceeds maximum ${field.max}`,
        value: rawValue,
        field: fieldName,
      };
    }
  }
  return null;
}

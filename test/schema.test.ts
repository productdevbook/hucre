import { describe, expect, it } from "vitest";
import { validateWithSchema } from "../src/_schema";
import { ValidationError } from "../src/errors";
import type { CellValue, SchemaDefinition } from "../src/_types";
import { serialToDate } from "../src/_date";

// ── Header Matching ───────────────────────────────────────────────────

describe("header matching", () => {
  it("should match by exact header name", () => {
    const rows: CellValue[][] = [
      ["Name", "Price"],
      ["Widget", 9.99],
    ];
    const schema: SchemaDefinition = {
      name: { column: "Name", type: "string" },
      price: { column: "Price", type: "number" },
    };
    const { data, errors } = validateWithSchema(rows, schema);
    expect(errors).toHaveLength(0);
    expect(data).toEqual([{ name: "Widget", price: 9.99 }]);
  });

  it("should match by case-insensitive header name", () => {
    const rows: CellValue[][] = [
      ["NAME", "price"],
      ["Widget", 9.99],
    ];
    const schema: SchemaDefinition = {
      name: { column: "name", type: "string" },
      price: { column: "Price", type: "number" },
    };
    const { data, errors } = validateWithSchema(rows, schema);
    expect(errors).toHaveLength(0);
    expect(data).toEqual([{ name: "Widget", price: 9.99 }]);
  });

  it("should match by columnIndex (ignore headers)", () => {
    const rows: CellValue[][] = [
      ["Anything", "Whatever"],
      ["Widget", 9.99],
    ];
    const schema: SchemaDefinition = {
      name: { columnIndex: 0, type: "string" },
      price: { columnIndex: 1, type: "number" },
    };
    const { data, errors } = validateWithSchema(rows, schema);
    expect(errors).toHaveLength(0);
    expect(data).toEqual([{ name: "Widget", price: 9.99 }]);
  });

  it("should match by field name when no column specified", () => {
    const rows: CellValue[][] = [
      ["name", "price"],
      ["Widget", 9.99],
    ];
    const schema: SchemaDefinition = {
      name: { type: "string" },
      price: { type: "number" },
    };
    const { data, errors } = validateWithSchema(rows, schema);
    expect(errors).toHaveLength(0);
    expect(data).toEqual([{ name: "Widget", price: 9.99 }]);
  });

  it("should match header with extra whitespace", () => {
    const rows: CellValue[][] = [
      ["  Name  ", " Price "],
      ["Widget", 9.99],
    ];
    const schema: SchemaDefinition = {
      name: { column: "Name", type: "string" },
      price: { column: "Price", type: "number" },
    };
    const { data, errors } = validateWithSchema(rows, schema);
    expect(errors).toHaveLength(0);
    expect(data).toEqual([{ name: "Widget", price: 9.99 }]);
  });

  it("should produce null for missing header on required field", () => {
    const rows: CellValue[][] = [["Name"], ["Widget"]];
    const schema: SchemaDefinition = {
      name: { column: "Name", type: "string" },
      price: { column: "Price", type: "number", required: true },
    };
    const { data, errors } = validateWithSchema(rows, schema);
    expect(errors).toHaveLength(1);
    expect(errors[0]!.field).toBe("price");
    expect(errors[0]!.message).toContain("Required field");
    expect(data[0]!.price).toBeNull();
  });

  it("should use header row at different position (headerRow: 2)", () => {
    const rows: CellValue[][] = [["This is a title row"], ["Name", "Price"], ["Widget", 9.99]];
    const schema: SchemaDefinition = {
      name: { column: "Name", type: "string" },
      price: { column: "Price", type: "number" },
    };
    const { data, errors } = validateWithSchema(rows, schema, {
      headerRow: 2,
    });
    expect(errors).toHaveLength(0);
    expect(data).toEqual([{ name: "Widget", price: 9.99 }]);
  });
});

// ── Type Coercion: String ─────────────────────────────────────────────

describe("type coercion — string", () => {
  const mkSchema = (): SchemaDefinition => ({
    val: { columnIndex: 0, type: "string" },
  });

  it("should convert number to string: 42 → '42'", () => {
    const { data } = validateWithSchema([[42]], mkSchema(), { headerRow: 0 });
    expect(data[0]!.val).toBe("42");
  });

  it("should convert boolean to string: true → 'true'", () => {
    const { data } = validateWithSchema([[true]], mkSchema(), { headerRow: 0 });
    expect(data[0]!.val).toBe("true");
  });

  it("should convert Date to ISO string", () => {
    const d = new Date(Date.UTC(2024, 0, 15));
    const { data } = validateWithSchema([[d]], mkSchema(), { headerRow: 0 });
    expect(data[0]!.val).toBe(d.toISOString());
  });

  it("should convert null to null (empty string → null since isEmpty)", () => {
    const { data } = validateWithSchema([[null]], mkSchema(), { headerRow: 0 });
    expect(data[0]!.val).toBeNull();
  });

  it("should keep string as-is but trim", () => {
    const { data } = validateWithSchema([["  hello  "]], mkSchema(), {
      headerRow: 0,
    });
    expect(data[0]!.val).toBe("hello");
  });
});

// ── Type Coercion: Number ─────────────────────────────────────────────

describe("type coercion — number", () => {
  const mkSchema = (): SchemaDefinition => ({
    val: { columnIndex: 0, type: "number" },
  });

  it('should convert string "42" → 42', () => {
    const { data, errors } = validateWithSchema([["42"]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.val).toBe(42);
  });

  it('should convert string "3.14" → 3.14', () => {
    const { data } = validateWithSchema([["3.14"]], mkSchema(), {
      headerRow: 0,
    });
    expect(data[0]!.val).toBe(3.14);
  });

  it('should convert string "-10" → -10', () => {
    const { data } = validateWithSchema([["-10"]], mkSchema(), {
      headerRow: 0,
    });
    expect(data[0]!.val).toBe(-10);
  });

  it('should error on string "abc"', () => {
    const { errors } = validateWithSchema([["abc"]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("Expected number");
    expect(errors[0]!.message).toContain("abc");
  });

  it("should return null for empty string (not required)", () => {
    const { data, errors } = validateWithSchema([[""]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.val).toBeNull();
  });

  it("should convert boolean true → 1, false → 0", () => {
    const schema: SchemaDefinition = {
      a: { columnIndex: 0, type: "number" },
      b: { columnIndex: 1, type: "number" },
    };
    const { data } = validateWithSchema([[true, false]], schema, {
      headerRow: 0,
    });
    expect(data[0]!.a).toBe(1);
    expect(data[0]!.b).toBe(0);
  });

  it("should keep number as-is", () => {
    const { data } = validateWithSchema([[42.5]], mkSchema(), {
      headerRow: 0,
    });
    expect(data[0]!.val).toBe(42.5);
  });

  it("should return null for null (not required)", () => {
    const { data, errors } = validateWithSchema([[null]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.val).toBeNull();
  });

  it('should strip commas: "1,234.56" → 1234.56', () => {
    const { data, errors } = validateWithSchema([["1,234.56"]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.val).toBe(1234.56);
  });
});

// ── Type Coercion: Integer ────────────────────────────────────────────

describe("type coercion — integer", () => {
  const mkSchema = (): SchemaDefinition => ({
    val: { columnIndex: 0, type: "integer" },
  });

  it('should convert string "42" → 42', () => {
    const { data, errors } = validateWithSchema([["42"]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.val).toBe(42);
  });

  it('should error on string "3.14" (not integer)', () => {
    const { errors } = validateWithSchema([["3.14"]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("Expected integer");
  });

  it("should error on number 3.14", () => {
    const { errors } = validateWithSchema([[3.14]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("Expected integer");
  });

  it("should allow number 42.0 → 42", () => {
    const { data, errors } = validateWithSchema([[42.0]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.val).toBe(42);
  });

  it('should error on string "abc"', () => {
    const { errors } = validateWithSchema([["abc"]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("Expected integer");
  });
});

// ── Type Coercion: Boolean ────────────────────────────────────────────

describe("type coercion — boolean", () => {
  const mkSchema = (): SchemaDefinition => ({
    val: { columnIndex: 0, type: "boolean" },
  });

  it('should convert "true"/"TRUE"/"True" → true', () => {
    for (const v of ["true", "TRUE", "True"]) {
      const { data, errors } = validateWithSchema([[v]], mkSchema(), {
        headerRow: 0,
      });
      expect(errors).toHaveLength(0);
      expect(data[0]!.val).toBe(true);
    }
  });

  it('should convert "false"/"FALSE"/"False" → false', () => {
    for (const v of ["false", "FALSE", "False"]) {
      const { data, errors } = validateWithSchema([[v]], mkSchema(), {
        headerRow: 0,
      });
      expect(errors).toHaveLength(0);
      expect(data[0]!.val).toBe(false);
    }
  });

  it('should convert "yes"/"no" → true/false', () => {
    const schema: SchemaDefinition = {
      a: { columnIndex: 0, type: "boolean" },
      b: { columnIndex: 1, type: "boolean" },
    };
    const { data } = validateWithSchema([["yes", "no"]], schema, {
      headerRow: 0,
    });
    expect(data[0]!.a).toBe(true);
    expect(data[0]!.b).toBe(false);
  });

  it('should convert "1"/"0" → true/false', () => {
    const schema: SchemaDefinition = {
      a: { columnIndex: 0, type: "boolean" },
      b: { columnIndex: 1, type: "boolean" },
    };
    const { data } = validateWithSchema([["1", "0"]], schema, {
      headerRow: 0,
    });
    expect(data[0]!.a).toBe(true);
    expect(data[0]!.b).toBe(false);
  });

  it("should convert number 1/0 → true/false", () => {
    const schema: SchemaDefinition = {
      a: { columnIndex: 0, type: "boolean" },
      b: { columnIndex: 1, type: "boolean" },
    };
    const { data } = validateWithSchema([[1, 0]], schema, { headerRow: 0 });
    expect(data[0]!.a).toBe(true);
    expect(data[0]!.b).toBe(false);
  });

  it('should error on string "abc"', () => {
    const { errors } = validateWithSchema([["abc"]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("Expected boolean");
  });

  it("should pass through boolean values unchanged", () => {
    const schema: SchemaDefinition = {
      a: { columnIndex: 0, type: "boolean" },
      b: { columnIndex: 1, type: "boolean" },
    };
    const { data, errors } = validateWithSchema([[true, false]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.a).toBe(true);
    expect(data[0]!.b).toBe(false);
  });
});

// ── Type Coercion: Date ───────────────────────────────────────────────

describe("type coercion — date", () => {
  const mkSchema = (): SchemaDefinition => ({
    val: { columnIndex: 0, type: "date" },
  });

  it("should pass through Date objects", () => {
    const d = new Date(Date.UTC(2024, 0, 15));
    const { data, errors } = validateWithSchema([[d]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.val).toBeInstanceOf(Date);
    expect((data[0]!.val as Date).toISOString()).toBe(d.toISOString());
  });

  it("should convert serial number to Date", () => {
    const serial = 45307; // some Excel serial
    const expected = serialToDate(serial);
    const { data, errors } = validateWithSchema([[serial]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.val).toBeInstanceOf(Date);
    expect((data[0]!.val as Date).toISOString()).toBe(expected.toISOString());
  });

  it("should parse ISO date string", () => {
    const { data, errors } = validateWithSchema([["2024-01-15"]], mkSchema(), { headerRow: 0 });
    expect(errors).toHaveLength(0);
    expect(data[0]!.val).toBeInstanceOf(Date);
    expect((data[0]!.val as Date).toISOString()).toContain("2024-01-15");
  });

  it('should error on unparseable string "abc"', () => {
    const { errors } = validateWithSchema([["abc"]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("Expected date");
  });

  it("should return null for null (not required)", () => {
    const { data, errors } = validateWithSchema([[null]], mkSchema(), {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.val).toBeNull();
  });
});

// ── Required Validation ───────────────────────────────────────────────

describe("required validation", () => {
  it("should error on missing required field", () => {
    const schema: SchemaDefinition = {
      name: { columnIndex: 0, type: "string", required: true },
    };
    const { errors } = validateWithSchema([[null]], schema, { headerRow: 0 });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.row).toBe(1);
    expect(errors[0]!.message).toContain("Required field");
  });

  it("should not error on missing optional field", () => {
    const schema: SchemaDefinition = {
      name: { columnIndex: 0, type: "string" },
    };
    const { data, errors } = validateWithSchema([[null]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.name).toBeNull();
  });

  it("should error on empty string for required field", () => {
    const schema: SchemaDefinition = {
      name: { columnIndex: 0, type: "string", required: true },
    };
    const { errors } = validateWithSchema([[""]], schema, { headerRow: 0 });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("Required field");
  });

  it("should NOT error on zero for required number (0 is valid)", () => {
    const schema: SchemaDefinition = {
      count: { columnIndex: 0, type: "number", required: true },
    };
    const { data, errors } = validateWithSchema([[0]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.count).toBe(0);
  });

  it("should NOT error on false for required boolean (false is valid)", () => {
    const schema: SchemaDefinition = {
      active: { columnIndex: 0, type: "boolean", required: true },
    };
    const { data, errors } = validateWithSchema([[false]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.active).toBe(false);
  });
});

// ── Pattern Validation ────────────────────────────────────────────────

describe("pattern validation", () => {
  it("should pass when pattern matches", () => {
    const schema: SchemaDefinition = {
      email: {
        columnIndex: 0,
        type: "string",
        pattern: /^[^@]+@[^@]+$/,
      },
    };
    const { errors } = validateWithSchema([["user@example.com"]], schema, { headerRow: 0 });
    expect(errors).toHaveLength(0);
  });

  it("should error when pattern does not match", () => {
    const schema: SchemaDefinition = {
      email: {
        columnIndex: 0,
        type: "string",
        pattern: /^[^@]+@[^@]+$/,
      },
    };
    const { errors } = validateWithSchema([["invalid"]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("does not match pattern");
  });

  it("should only apply pattern to string values", () => {
    const schema: SchemaDefinition = {
      count: {
        columnIndex: 0,
        type: "number",
        pattern: /^\d+$/,
      },
    };
    // After number coercion, value is a number, not a string. Pattern is skipped.
    const { errors } = validateWithSchema([["42"]], schema, { headerRow: 0 });
    expect(errors).toHaveLength(0);
  });
});

// ── Min/Max Validation ────────────────────────────────────────────────

describe("min/max validation", () => {
  it("should error on number below min", () => {
    const schema: SchemaDefinition = {
      price: { columnIndex: 0, type: "number", min: 0 },
    };
    const { errors } = validateWithSchema([[-1]], schema, { headerRow: 0 });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("below minimum");
  });

  it("should error on number above max", () => {
    const schema: SchemaDefinition = {
      qty: { columnIndex: 0, type: "number", max: 10 },
    };
    const { errors } = validateWithSchema([[42]], schema, { headerRow: 0 });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("exceeds maximum");
  });

  it("should pass on number in range", () => {
    const schema: SchemaDefinition = {
      qty: { columnIndex: 0, type: "number", min: 0, max: 100 },
    };
    const { errors } = validateWithSchema([[50]], schema, { headerRow: 0 });
    expect(errors).toHaveLength(0);
  });

  it("should error on string shorter than min length", () => {
    const schema: SchemaDefinition = {
      code: { columnIndex: 0, type: "string", min: 3 },
    };
    const { errors } = validateWithSchema([["ab"]], schema, { headerRow: 0 });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("below minimum");
  });

  it("should error on string longer than max length", () => {
    const schema: SchemaDefinition = {
      code: { columnIndex: 0, type: "string", max: 5 },
    };
    const { errors } = validateWithSchema([["toolong"]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("exceeds maximum");
  });

  it("should pass on string within length range", () => {
    const schema: SchemaDefinition = {
      code: { columnIndex: 0, type: "string", min: 2, max: 10 },
    };
    const { errors } = validateWithSchema([["hello"]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
  });
});

// ── Enum Validation ───────────────────────────────────────────────────

describe("enum validation", () => {
  it("should pass when value is in enum", () => {
    const schema: SchemaDefinition = {
      status: {
        columnIndex: 0,
        type: "string",
        enum: ["active", "inactive"],
      },
    };
    const { errors } = validateWithSchema([["active"]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
  });

  it("should error when value is not in enum", () => {
    const schema: SchemaDefinition = {
      status: {
        columnIndex: 0,
        type: "string",
        enum: ["active", "inactive"],
      },
    };
    const { errors } = validateWithSchema([["deleted"]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("must be one of");
    expect(errors[0]!.message).toContain("active");
    expect(errors[0]!.message).toContain("inactive");
  });
});

// ── Custom Validate Function ──────────────────────────────────────────

describe("custom validate function", () => {
  it("should pass when validate returns true", () => {
    const schema: SchemaDefinition = {
      sku: {
        columnIndex: 0,
        type: "string",
        validate: (v) => typeof v === "string" && v.startsWith("SKU-"),
      },
    };
    const { errors } = validateWithSchema([["SKU-001"]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
  });

  it("should error with generic message when validate returns false", () => {
    const schema: SchemaDefinition = {
      sku: {
        columnIndex: 0,
        type: "string",
        validate: () => false,
      },
    };
    const { errors } = validateWithSchema([["BAD"]], schema, { headerRow: 0 });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toContain("Custom validation failed");
  });

  it("should use returned string as error message", () => {
    const schema: SchemaDefinition = {
      sku: {
        columnIndex: 0,
        type: "string",
        validate: (v) =>
          typeof v === "string" && v.startsWith("SKU-") ? true : "SKU must start with 'SKU-'",
      },
    };
    const { errors } = validateWithSchema([["BAD"]], schema, { headerRow: 0 });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.message).toBe("SKU must start with 'SKU-'");
  });
});

// ── Transform Function ────────────────────────────────────────────────

describe("transform function", () => {
  it("should apply transform after validation", () => {
    const schema: SchemaDefinition = {
      code: {
        columnIndex: 0,
        type: "string",
        transform: (v) => String(v).toUpperCase(),
      },
    };
    const { data, errors } = validateWithSchema([["hello"]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(0);
    expect(data[0]!.code).toBe("HELLO");
  });

  it("should use transform return value in output", () => {
    const schema: SchemaDefinition = {
      price: {
        columnIndex: 0,
        type: "number",
        transform: (v) => Math.round((v as number) * 100),
      },
    };
    const { data } = validateWithSchema([[9.99]], schema, { headerRow: 0 });
    expect(data[0]!.price).toBe(999);
  });
});

// ── Default Values ────────────────────────────────────────────────────

describe("default values", () => {
  it("should use default for null value", () => {
    const schema: SchemaDefinition = {
      status: { columnIndex: 0, type: "string", default: "active" },
    };
    const { data } = validateWithSchema([[null]], schema, { headerRow: 0 });
    expect(data[0]!.status).toBe("active");
  });

  it("should use default for undefined (missing column)", () => {
    const schema: SchemaDefinition = {
      status: { columnIndex: 5, type: "string", default: "active" },
    };
    // Row only has 1 column, so index 5 is out of bounds (→ null)
    const { data } = validateWithSchema([["Widget"]], schema, {
      headerRow: 0,
    });
    expect(data[0]!.status).toBe("active");
  });

  it("should use default for empty string on optional field", () => {
    const schema: SchemaDefinition = {
      status: { columnIndex: 0, type: "string", default: "active" },
    };
    const { data } = validateWithSchema([[""]], schema, { headerRow: 0 });
    expect(data[0]!.status).toBe("active");
  });

  it("should NOT use default when value is present", () => {
    const schema: SchemaDefinition = {
      status: { columnIndex: 0, type: "string", default: "active" },
    };
    const { data } = validateWithSchema([["inactive"]], schema, {
      headerRow: 0,
    });
    expect(data[0]!.status).toBe("inactive");
  });
});

// ── Skip Empty Rows ───────────────────────────────────────────────────

describe("skip empty rows", () => {
  const schema: SchemaDefinition = {
    name: { columnIndex: 0, type: "string" },
  };

  it("should skip row with all nulls", () => {
    const rows: CellValue[][] = [["Alice"], [null], ["Bob"]];
    const { data } = validateWithSchema(rows, schema, {
      headerRow: 0,
      skipEmptyRows: true,
    });
    expect(data).toHaveLength(2);
    expect(data[0]!.name).toBe("Alice");
    expect(data[1]!.name).toBe("Bob");
  });

  it("should skip row with all empty strings", () => {
    const rows: CellValue[][] = [["Alice"], [""], ["Bob"]];
    const { data } = validateWithSchema(rows, schema, {
      headerRow: 0,
      skipEmptyRows: true,
    });
    expect(data).toHaveLength(2);
  });

  it("should NOT skip row with some values", () => {
    const schema2: SchemaDefinition = {
      a: { columnIndex: 0, type: "string" },
      b: { columnIndex: 1, type: "string" },
    };
    const rows: CellValue[][] = [
      ["Alice", "Smith"],
      [null, "Jones"],
      ["Bob", "Brown"],
    ];
    const { data } = validateWithSchema(rows, schema2, {
      headerRow: 0,
      skipEmptyRows: true,
    });
    expect(data).toHaveLength(3);
  });
});

// ── Error Collection ──────────────────────────────────────────────────

describe("error collection", () => {
  it("should collect multiple errors per row", () => {
    const schema: SchemaDefinition = {
      name: { columnIndex: 0, type: "string", required: true },
      price: { columnIndex: 1, type: "number", required: true },
    };
    const { errors } = validateWithSchema([[null, null]], schema, {
      headerRow: 0,
    });
    expect(errors).toHaveLength(2);
  });

  it("should collect errors from multiple rows", () => {
    const schema: SchemaDefinition = {
      price: { columnIndex: 0, type: "number" },
    };
    const rows: CellValue[][] = [["abc"], ["def"]];
    const { errors } = validateWithSchema(rows, schema, { headerRow: 0 });
    expect(errors).toHaveLength(2);
    expect(errors[0]!.row).toBe(1);
    expect(errors[1]!.row).toBe(2);
  });

  it("should include row, column, message, value, field in errors", () => {
    const schema: SchemaDefinition = {
      price: { column: "Price", columnIndex: 0, type: "number" },
    };
    const { errors } = validateWithSchema([["abc"]], schema, { headerRow: 0 });
    expect(errors).toHaveLength(1);
    const err = errors[0]!;
    expect(err.row).toBe(1);
    expect(err.column).toBe("Price");
    expect(err.message).toContain("Expected number");
    expect(err.value).toBe("abc");
    expect(err.field).toBe("price");
  });

  it('should throw on first error when errorMode is "throw"', () => {
    const schema: SchemaDefinition = {
      name: { columnIndex: 0, type: "string", required: true },
      price: { columnIndex: 1, type: "number", required: true },
    };
    expect(() =>
      validateWithSchema([[null, null]], schema, {
        headerRow: 0,
        errorMode: "throw",
      }),
    ).toThrow(ValidationError);
  });

  it('should return all errors when errorMode is "collect"', () => {
    const schema: SchemaDefinition = {
      name: { columnIndex: 0, type: "string", required: true },
      price: { columnIndex: 1, type: "number", required: true },
    };
    const { errors } = validateWithSchema([[null, null]], schema, {
      headerRow: 0,
      errorMode: "collect",
    });
    expect(errors).toHaveLength(2);
  });
});

// ── Edge Cases ────────────────────────────────────────────────────────

describe("edge cases", () => {
  it("should return empty result for empty data (no rows)", () => {
    const schema: SchemaDefinition = {
      name: { columnIndex: 0, type: "string" },
    };
    const { data, errors } = validateWithSchema([], schema);
    expect(data).toEqual([]);
    expect(errors).toEqual([]);
  });

  it("should return empty result for schema with no fields", () => {
    const { data, errors } = validateWithSchema(
      [
        ["a", "b"],
        [1, 2],
      ],
      {},
    );
    expect(data).toEqual([]);
    expect(errors).toEqual([]);
  });

  it("should ignore extra columns beyond schema", () => {
    const schema: SchemaDefinition = {
      name: { columnIndex: 0, type: "string" },
    };
    const rows: CellValue[][] = [["Alice", "extra1", "extra2"]];
    const { data } = validateWithSchema(rows, schema, { headerRow: 0 });
    expect(data).toHaveLength(1);
    expect(data[0]!.name).toBe("Alice");
    expect(Object.keys(data[0]!)).toEqual(["name"]);
  });

  it("should treat missing columns as null (row shorter than schema)", () => {
    const schema: SchemaDefinition = {
      a: { columnIndex: 0, type: "string" },
      b: { columnIndex: 1, type: "string" },
      c: { columnIndex: 2, type: "string" },
    };
    const rows: CellValue[][] = [["only-one"]];
    const { data } = validateWithSchema(rows, schema, { headerRow: 0 });
    expect(data[0]!.a).toBe("only-one");
    expect(data[0]!.b).toBeNull();
    expect(data[0]!.c).toBeNull();
  });

  it("should handle column that doesn't exist in headers (not required)", () => {
    const rows: CellValue[][] = [["Name"], ["Widget"]];
    const schema: SchemaDefinition = {
      name: { column: "Name", type: "string" },
      color: { column: "Color", type: "string" },
    };
    const { data, errors } = validateWithSchema(rows, schema);
    expect(errors).toHaveLength(0);
    expect(data[0]!.name).toBe("Widget");
    expect(data[0]!.color).toBeNull();
  });
});

// ── Integration: Product Import ───────────────────────────────────────

describe("integration — product import", () => {
  const schema: SchemaDefinition = {
    name: { column: "Name", type: "string", required: true, min: 1 },
    price: { column: "Price", type: "number", required: true, min: 0 },
    sku: {
      column: "SKU",
      type: "string",
      pattern: /^[A-Z]{2,4}-\d+$/,
    },
    active: { column: "Active", type: "boolean", default: true },
  };

  it("should validate and transform a complete product import", () => {
    const rows: CellValue[][] = [
      ["Name", "Price", "SKU", "Active"],
      ["Widget", 9.99, "WDG-001", "true"],
      ["Gadget", 24.5, "GDG-002", "false"],
      ["Doohickey", 1.5, null, null],
    ];

    const { data, errors } = validateWithSchema(rows, schema);
    expect(errors).toHaveLength(0);
    expect(data).toHaveLength(3);

    expect(data[0]).toEqual({
      name: "Widget",
      price: 9.99,
      sku: "WDG-001",
      active: true,
    });
    expect(data[1]).toEqual({
      name: "Gadget",
      price: 24.5,
      sku: "GDG-002",
      active: false,
    });
    expect(data[2]).toEqual({
      name: "Doohickey",
      price: 1.5,
      sku: null,
      active: true, // default
    });
  });

  it("should report errors for invalid product data", () => {
    const rows: CellValue[][] = [
      ["Name", "Price", "SKU", "Active"],
      ["", -5, "bad-sku", "maybe"],
    ];

    const { errors } = validateWithSchema(rows, schema);
    // Expected errors:
    // 1. name required + empty
    // 2. price below min 0
    // 3. SKU pattern mismatch
    // 4. Active invalid boolean
    expect(errors.length).toBeGreaterThanOrEqual(3);

    const fieldNames = errors.map((e) => e.field);
    expect(fieldNames).toContain("name");
    expect(fieldNames).toContain("price");
    expect(fieldNames).toContain("active");
  });
});

// ── Integration: Employee Import ──────────────────────────────────────

describe("integration — employee import", () => {
  const schema: SchemaDefinition = {
    id: { column: "ID", type: "integer", required: true },
    name: { column: "Full Name", type: "string", required: true, min: 2 },
    email: {
      column: "Email",
      type: "string",
      required: true,
      pattern: /^[^@]+@[^@]+\.[^@]+$/,
    },
    salary: { column: "Salary", type: "number", min: 0, max: 1_000_000 },
    department: {
      column: "Dept",
      type: "string",
      enum: ["Engineering", "Sales", "HR", "Marketing"],
      transform: (v) => String(v).trim(),
    },
    startDate: { column: "Start Date", type: "date" },
    isManager: { column: "Manager", type: "boolean", default: false },
  };

  it("should process a valid employee spreadsheet", () => {
    const rows: CellValue[][] = [
      ["ID", "Full Name", "Email", "Salary", "Dept", "Start Date", "Manager"],
      [1, "Alice Johnson", "alice@corp.com", "85,000", "Engineering", "2024-01-15", "yes"],
      [2, "Bob Smith", "bob@corp.com", 72000, "Sales", "2023-06-01", "no"],
      [3, "Charlie Brown", "charlie@corp.com", null, "HR", null, null],
    ];

    const { data, errors } = validateWithSchema(rows, schema);
    expect(errors).toHaveLength(0);
    expect(data).toHaveLength(3);

    expect(data[0]!.id).toBe(1);
    expect(data[0]!.name).toBe("Alice Johnson");
    expect(data[0]!.salary).toBe(85000);
    expect(data[0]!.department).toBe("Engineering");
    expect(data[0]!.startDate).toBeInstanceOf(Date);
    expect(data[0]!.isManager).toBe(true);

    expect(data[2]!.salary).toBeNull();
    expect(data[2]!.startDate).toBeNull();
    expect(data[2]!.isManager).toBe(false); // default
  });

  it("should report multiple validation failures", () => {
    const rows: CellValue[][] = [
      ["ID", "Full Name", "Email", "Salary", "Dept", "Start Date", "Manager"],
      ["abc", "A", "invalid-email", 2_000_000, "Finance", "not-a-date", "maybe"],
    ];

    const { errors } = validateWithSchema(rows, schema);
    // id: not integer, name: min 2, email: pattern, salary: max, dept: enum, date: invalid, manager: boolean
    expect(errors.length).toBeGreaterThanOrEqual(5);

    const fields = errors.map((e) => e.field);
    expect(fields).toContain("id");
    expect(fields).toContain("email");
    expect(fields).toContain("salary");
    expect(fields).toContain("department");
    expect(fields).toContain("isManager");
  });
});

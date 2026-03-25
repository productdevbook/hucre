import { describe, it, expect } from "vitest";
import { fillTemplate } from "../src/template";
import type { Workbook, Sheet } from "../src/_types";

// ── Helper ──────────────────────────────────────────────────────────

function makeWorkbook(rows: (string | number | boolean | Date | null)[][]): Workbook {
  const sheet: Sheet = {
    name: "Sheet1",
    rows,
  };
  return { sheets: [sheet] };
}

// ── Tests ────────────────────────────────────────────────────────────

describe("fillTemplate — basic placeholders", () => {
  it("replaces {{name}} with string value", () => {
    const wb = makeWorkbook([["Hello {{name}}", "other"]]);
    const result = fillTemplate(wb, { name: "World" });
    expect(result.sheets[0].rows[0][0]).toBe("Hello World");
    expect(result.sheets[0].rows[0][1]).toBe("other");
  });

  it("replaces a cell that is only a placeholder", () => {
    const wb = makeWorkbook([["{{company}}"]]);
    const result = fillTemplate(wb, { company: "Acme Corp" });
    expect(result.sheets[0].rows[0][0]).toBe("Acme Corp");
  });

  it("handles multiple placeholders in one cell", () => {
    const wb = makeWorkbook([["{{first}} {{last}}"]]);
    const result = fillTemplate(wb, { first: "John", last: "Doe" });
    expect(result.sheets[0].rows[0][0]).toBe("John Doe");
  });

  it("handles repeated placeholders", () => {
    const wb = makeWorkbook([["{{x}} and {{x}}"]]);
    const result = fillTemplate(wb, { x: "A" });
    expect(result.sheets[0].rows[0][0]).toBe("A and A");
  });
});

describe("fillTemplate — placeholder not found", () => {
  it("leaves unmatched placeholders as-is", () => {
    const wb = makeWorkbook([["Hello {{unknown}}"]]);
    const result = fillTemplate(wb, { name: "World" });
    expect(result.sheets[0].rows[0][0]).toBe("Hello {{unknown}}");
  });

  it("replaces known and keeps unknown in same cell", () => {
    const wb = makeWorkbook([["{{known}} and {{unknown}}"]]);
    const result = fillTemplate(wb, { known: "yes" });
    expect(result.sheets[0].rows[0][0]).toBe("yes and {{unknown}}");
  });
});

describe("fillTemplate — typed values", () => {
  it("replaces single placeholder with number value directly", () => {
    const wb = makeWorkbook([["{{total}}"]]);
    const result = fillTemplate(wb, { total: 12500 });
    expect(result.sheets[0].rows[0][0]).toBe(12500);
    expect(typeof result.sheets[0].rows[0][0]).toBe("number");
  });

  it("replaces single placeholder with boolean value directly", () => {
    const wb = makeWorkbook([["{{active}}"]]);
    const result = fillTemplate(wb, { active: true });
    expect(result.sheets[0].rows[0][0]).toBe(true);
    expect(typeof result.sheets[0].rows[0][0]).toBe("boolean");
  });

  it("replaces single placeholder with Date value directly", () => {
    const date = new Date("2025-06-15T00:00:00Z");
    const wb = makeWorkbook([["{{date}}"]]);
    const result = fillTemplate(wb, { date });
    expect(result.sheets[0].rows[0][0]).toBeInstanceOf(Date);
    expect((result.sheets[0].rows[0][0] as Date).toISOString()).toBe("2025-06-15T00:00:00.000Z");
  });

  it("replaces single placeholder with null", () => {
    const wb = makeWorkbook([["{{empty}}"]]);
    const result = fillTemplate(wb, { empty: null });
    expect(result.sheets[0].rows[0][0]).toBe(null);
  });

  it("converts number to string when mixed with text", () => {
    const wb = makeWorkbook([["Total: {{amount}} USD"]]);
    const result = fillTemplate(wb, { amount: 500 });
    expect(result.sheets[0].rows[0][0]).toBe("Total: 500 USD");
    expect(typeof result.sheets[0].rows[0][0]).toBe("string");
  });

  it("converts Date to ISO string when mixed with text", () => {
    const date = new Date("2025-01-01T00:00:00Z");
    const wb = makeWorkbook([["Date: {{date}}"]]);
    const result = fillTemplate(wb, { date });
    expect(result.sheets[0].rows[0][0]).toBe("Date: 2025-01-01T00:00:00.000Z");
  });

  it("converts null to empty string when mixed with text", () => {
    const wb = makeWorkbook([["Value: {{val}}!"]]);
    const result = fillTemplate(wb, { val: null });
    expect(result.sheets[0].rows[0][0]).toBe("Value: !");
  });
});

describe("fillTemplate — non-string cells are untouched", () => {
  it("skips number cells", () => {
    const wb = makeWorkbook([[42, "{{name}}"]]);
    const result = fillTemplate(wb, { name: "Alice" });
    expect(result.sheets[0].rows[0][0]).toBe(42);
    expect(result.sheets[0].rows[0][1]).toBe("Alice");
  });

  it("skips boolean cells", () => {
    const wb = makeWorkbook([[true, "{{flag}}"]]);
    const result = fillTemplate(wb, { flag: "yes" });
    expect(result.sheets[0].rows[0][0]).toBe(true);
  });

  it("skips null cells", () => {
    const wb = makeWorkbook([[null, "{{val}}"]]);
    const result = fillTemplate(wb, { val: "ok" });
    expect(result.sheets[0].rows[0][0]).toBe(null);
  });
});

describe("fillTemplate — multiple sheets", () => {
  it("fills placeholders across all sheets", () => {
    const wb: Workbook = {
      sheets: [
        { name: "Sheet1", rows: [["{{name}}"]] },
        { name: "Sheet2", rows: [["Hi {{name}}"]] },
      ],
    };
    const result = fillTemplate(wb, { name: "Test" });
    expect(result.sheets[0].rows[0][0]).toBe("Test");
    expect(result.sheets[1].rows[0][0]).toBe("Hi Test");
  });
});

describe("fillTemplate — whitespace in placeholder", () => {
  it("trims whitespace inside braces", () => {
    const wb = makeWorkbook([["{{ name }}"]]);
    const result = fillTemplate(wb, { name: "Trimmed" });
    expect(result.sheets[0].rows[0][0]).toBe("Trimmed");
  });
});

describe("fillTemplate — cells with no placeholders", () => {
  it("leaves cells without {{ untouched", () => {
    const wb = makeWorkbook([["No placeholders here", "Just text"]]);
    const result = fillTemplate(wb, { anything: "value" });
    expect(result.sheets[0].rows[0][0]).toBe("No placeholders here");
    expect(result.sheets[0].rows[0][1]).toBe("Just text");
  });
});

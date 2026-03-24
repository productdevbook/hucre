import { describe, expect, it } from "vitest";
import { parseCsv, parseCsvObjects, detectDelimiter, stripBom } from "../src/csv/index";

describe("stripBom", () => {
  it("should strip UTF-8 BOM", () => {
    const input = "\uFEFFhello";
    expect(stripBom(input)).toBe("hello");
  });

  it("should strip UTF-16 LE BOM", () => {
    const input = "\uFFFEhello";
    expect(stripBom(input)).toBe("hello");
  });

  it("should strip UTF-16 BE BOM", () => {
    const input = "\uFEFFhello";
    expect(stripBom(input)).toBe("hello");
  });

  it("should return empty string unchanged", () => {
    expect(stripBom("")).toBe("");
  });

  it("should leave strings without BOM unchanged", () => {
    expect(stripBom("hello")).toBe("hello");
  });
});

describe("detectDelimiter", () => {
  it("should detect comma delimiter", () => {
    expect(detectDelimiter("a,b,c\n1,2,3")).toBe(",");
  });

  it("should detect semicolon delimiter", () => {
    expect(detectDelimiter("a;b;c\n1;2;3")).toBe(";");
  });

  it("should detect tab delimiter", () => {
    expect(detectDelimiter("a\tb\tc\n1\t2\t3")).toBe("\t");
  });

  it("should detect pipe delimiter", () => {
    expect(detectDelimiter("a|b|c\n1|2|3")).toBe("|");
  });

  it("should default to comma for empty input", () => {
    expect(detectDelimiter("")).toBe(",");
  });

  it("should default to comma for single-column data", () => {
    expect(detectDelimiter("hello\nworld")).toBe(",");
  });
});

describe("parseCsv", () => {
  // ── Basic parsing ──────────────────────────────────────────────

  it("should parse simple CSV", () => {
    const result = parseCsv("a,b,c\n1,2,3");
    expect(result).toEqual([
      ["a", "b", "c"],
      ["1", "2", "3"],
    ]);
  });

  it("should parse single column CSV", () => {
    const result = parseCsv("a\nb\nc");
    expect(result).toEqual([["a"], ["b"], ["c"]]);
  });

  it("should parse single row CSV (no newline)", () => {
    const result = parseCsv("a,b,c");
    expect(result).toEqual([["a", "b", "c"]]);
  });

  it("should return empty array for empty input", () => {
    expect(parseCsv("")).toEqual([]);
  });

  it("should handle only headers, no data", () => {
    const result = parseCsv("name,age,city\n");
    expect(result).toEqual([["name", "age", "city"]]);
  });

  // ── Quoting and escaping (RFC 4180) ────────────────────────────

  it("should handle quoted fields with embedded commas", () => {
    const result = parseCsv('"hello, world",foo');
    expect(result).toEqual([["hello, world", "foo"]]);
  });

  it("should handle quoted fields with escaped quotes", () => {
    const result = parseCsv('"foo""bar",baz');
    expect(result).toEqual([['foo"bar', "baz"]]);
  });

  it("should handle embedded newlines in quoted fields", () => {
    const result = parseCsv('"line1\nline2",b\nc,d');
    expect(result).toEqual([
      ["line1\nline2", "b"],
      ["c", "d"],
    ]);
  });

  it("should handle field with only quotes", () => {
    const result = parseCsv('"""",b');
    expect(result).toEqual([['"', "b"]]);
  });

  // ── Empty fields ───────────────────────────────────────────────

  it("should handle empty fields", () => {
    const result = parseCsv("a,,c");
    expect(result).toEqual([["a", "", "c"]]);
  });

  it("should handle consecutive delimiters", () => {
    const result = parseCsv("a,,,b");
    expect(result).toEqual([["a", "", "", "b"]]);
  });

  it("should handle empty rows", () => {
    const result = parseCsv("a,b\n\nc,d");
    expect(result).toEqual([["a", "b"], [""], ["c", "d"]]);
  });

  // ── Whitespace-only fields ─────────────────────────────────────

  it("should preserve whitespace-only fields", () => {
    const result = parseCsv("  ,b, c ");
    expect(result).toEqual([["  ", "b", " c "]]);
  });

  // ── Line endings ───────────────────────────────────────────────

  it("should handle \\r\\n line endings", () => {
    const result = parseCsv("a,b\r\nc,d");
    expect(result).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("should handle \\r line endings (classic Mac)", () => {
    const result = parseCsv("a,b\rc,d");
    expect(result).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("should handle mixed line endings", () => {
    const result = parseCsv("a,b\nc,d\r\ne,f\rg,h");
    expect(result).toEqual([
      ["a", "b"],
      ["c", "d"],
      ["e", "f"],
      ["g", "h"],
    ]);
  });

  it("should handle trailing newline at end of file", () => {
    const result = parseCsv("a,b\nc,d\n");
    expect(result).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("should handle trailing \\r\\n at end of file", () => {
    const result = parseCsv("a,b\r\nc,d\r\n");
    expect(result).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  // ── BOM ────────────────────────────────────────────────────────

  it("should strip UTF-8 BOM by default", () => {
    const result = parseCsv("\uFEFFa,b\n1,2");
    expect(result).toEqual([
      ["a", "b"],
      ["1", "2"],
    ]);
  });

  it("should not strip BOM when skipBom is false", () => {
    const result = parseCsv("\uFEFFa,b", { skipBom: false });
    expect(result[0]![0]).toBe("\uFEFFa");
  });

  // ── Delimiter options ──────────────────────────────────────────

  it("should use explicit delimiter override", () => {
    const result = parseCsv("a;b;c\n1;2;3", { delimiter: ";" });
    expect(result).toEqual([
      ["a", "b", "c"],
      ["1", "2", "3"],
    ]);
  });

  it("should auto-detect semicolon delimiter", () => {
    const result = parseCsv("a;b;c\n1;2;3");
    expect(result).toEqual([
      ["a", "b", "c"],
      ["1", "2", "3"],
    ]);
  });

  it("should auto-detect tab delimiter", () => {
    const result = parseCsv("a\tb\tc\n1\t2\t3");
    expect(result).toEqual([
      ["a", "b", "c"],
      ["1", "2", "3"],
    ]);
  });

  it("should auto-detect pipe delimiter", () => {
    const result = parseCsv("a|b|c\n1|2|3");
    expect(result).toEqual([
      ["a", "b", "c"],
      ["1", "2", "3"],
    ]);
  });

  // ── Comment lines ──────────────────────────────────────────────

  it("should skip comment lines with # prefix", () => {
    const result = parseCsv("# comment\na,b\n1,2", { comment: "#" });
    expect(result).toEqual([
      ["a", "b"],
      ["1", "2"],
    ]);
  });

  it("should not skip lines when comment is not set", () => {
    const result = parseCsv("# not a comment\na,b");
    expect(result).toEqual([["# not a comment"], ["a", "b"]]);
  });

  // ── Skip empty rows ───────────────────────────────────────────

  it("should skip empty rows when skipEmptyRows is true", () => {
    const result = parseCsv("a,b\n\n\nc,d", { skipEmptyRows: true });
    expect(result).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("should keep empty rows by default", () => {
    const result = parseCsv("a,b\n\nc,d");
    expect(result).toEqual([["a", "b"], [""], ["c", "d"]]);
  });

  // ── Type inference ─────────────────────────────────────────────

  it("should infer integer numbers", () => {
    const result = parseCsv("123", { typeInference: true });
    expect(result).toEqual([[123]]);
  });

  it("should infer float numbers", () => {
    const result = parseCsv("1.5", { typeInference: true });
    expect(result).toEqual([[1.5]]);
  });

  it("should infer negative numbers", () => {
    const result = parseCsv("-42", { typeInference: true });
    expect(result).toEqual([[-42]]);
  });

  it("should infer scientific notation numbers", () => {
    const result = parseCsv("1e10", { typeInference: true });
    expect(result).toEqual([[1e10]]);
  });

  it("should infer locale-aware numbers with commas", () => {
    // "1,234.56" has comma as thousands separator
    const result = parseCsv("1,234.56", {
      typeInference: true,
      delimiter: ";",
    });
    expect(result).toEqual([[1234.56]]);
  });

  it("should infer booleans: true/false", () => {
    const result = parseCsv("true,false", { typeInference: true });
    expect(result).toEqual([[true, false]]);
  });

  it("should infer booleans: TRUE/FALSE", () => {
    const result = parseCsv("TRUE,FALSE", { typeInference: true });
    expect(result).toEqual([[true, false]]);
  });

  it("should infer booleans: yes/no", () => {
    const result = parseCsv("yes,no", { typeInference: true });
    expect(result).toEqual([[true, false]]);
  });

  it("should infer booleans: 1/0", () => {
    const result = parseCsv("1,0", { typeInference: true });
    expect(result).toEqual([[true, false]]);
  });

  it("should infer ISO 8601 dates", () => {
    const result = parseCsv("2024-01-15", { typeInference: true });
    expect(result[0]![0]).toBeInstanceOf(Date);
    expect((result[0]![0] as Date).toISOString()).toContain("2024-01-15");
  });

  it("should infer ISO 8601 datetime with timezone", () => {
    const result = parseCsv("2024-01-15T10:30:00Z", { typeInference: true });
    expect(result[0]![0]).toBeInstanceOf(Date);
  });

  it("should keep strings as strings when typeInference is disabled", () => {
    const result = parseCsv("123,true,2024-01-15", { typeInference: false });
    expect(result).toEqual([["123", "true", "2024-01-15"]]);
  });

  it("should keep non-numeric strings as strings", () => {
    const result = parseCsv("hello,world", { typeInference: true });
    expect(result).toEqual([["hello", "world"]]);
  });

  // ── parseCsvObjects ────────────────────────────────────────────

  it("should parse CSV objects with headers", () => {
    const { data, headers } = parseCsvObjects("name,age\nAlice,30\nBob,25", {
      header: true,
    });
    expect(headers).toEqual(["name", "age"]);
    expect(data).toEqual([
      { name: "Alice", age: "30" },
      { name: "Bob", age: "25" },
    ]);
  });

  it("should handle header row with trailing spaces", () => {
    const { headers } = parseCsvObjects(" name , age \nAlice,30", {
      header: true,
    });
    expect(headers).toEqual(["name", "age"]);
  });

  it("should return empty data and headers for empty input", () => {
    const { data, headers } = parseCsvObjects("", { header: true });
    expect(data).toEqual([]);
    expect(headers).toEqual([]);
  });

  it("should handle only headers, no data rows", () => {
    const { data, headers } = parseCsvObjects("name,age\n", { header: true });
    expect(headers).toEqual(["name", "age"]);
    expect(data).toEqual([]);
  });

  it("should apply type inference on object values", () => {
    const { data } = parseCsvObjects("name,age\nAlice,30", {
      header: true,
      typeInference: true,
    });
    expect(data[0]!.age).toBe(30);
  });

  // ── Unicode content ────────────────────────────────────────────

  it("should handle Chinese characters", () => {
    const result = parseCsv("名前,年齢\n太郎,25");
    expect(result).toEqual([
      ["名前", "年齢"],
      ["太郎", "25"],
    ]);
  });

  it("should handle Arabic characters", () => {
    const result = parseCsv("اسم,عمر\nأحمد,30");
    expect(result).toEqual([
      ["اسم", "عمر"],
      ["أحمد", "30"],
    ]);
  });

  it("should handle emoji content", () => {
    const result = parseCsv("emoji,name\n🎉,party\n🚀,rocket");
    expect(result).toEqual([
      ["emoji", "name"],
      ["🎉", "party"],
      ["🚀", "rocket"],
    ]);
  });

  // ── Performance ────────────────────────────────────────────────

  it("should handle large CSV (10,000 rows) without being slow", () => {
    const header = "a,b,c,d,e";
    const row = "1,2,3,4,5";
    const lines = [header];
    for (let i = 0; i < 10_000; i++) {
      lines.push(row);
    }
    const input = lines.join("\n");

    const start = performance.now();
    const result = parseCsv(input);
    const elapsed = performance.now() - start;

    expect(result.length).toBe(10_001);
    expect(result[0]).toEqual(["a", "b", "c", "d", "e"]);
    // Should complete in under 1 second (generous limit)
    expect(elapsed).toBeLessThan(1000);
  });
});

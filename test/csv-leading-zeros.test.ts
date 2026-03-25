import { describe, expect, it } from "vitest";
import { parseCsv } from "../src/csv/index";

describe("CSV preserveLeadingZeros", () => {
  it('should keep "0123" as string when preserveLeadingZeros is true', () => {
    const result = parseCsv("0123", { typeInference: true, preserveLeadingZeros: true });
    expect(result[0]![0]).toBe("0123");
    expect(typeof result[0]![0]).toBe("string");
  });

  it('should convert "0123" to number 123 when preserveLeadingZeros is false', () => {
    const result = parseCsv("0123", { typeInference: true, preserveLeadingZeros: false });
    expect(result[0]![0]).toBe(123);
    expect(typeof result[0]![0]).toBe("number");
  });

  it('"0" should become boolean false (always, regardless of preserveLeadingZeros)', () => {
    const result = parseCsv("0", { typeInference: true, preserveLeadingZeros: true });
    expect(result[0]![0]).toBe(false);
  });

  it('"0.5" should become number 0.5 (always)', () => {
    const result = parseCsv("0.5", { typeInference: true, preserveLeadingZeros: true });
    expect(result[0]![0]).toBe(0.5);
    expect(typeof result[0]![0]).toBe("number");
  });

  it('"00" should stay as string "00"', () => {
    const result = parseCsv("00", { typeInference: true, preserveLeadingZeros: true });
    expect(result[0]![0]).toBe("00");
    expect(typeof result[0]![0]).toBe("string");
  });

  it('"007" should stay as string "007"', () => {
    const result = parseCsv("007", { typeInference: true, preserveLeadingZeros: true });
    expect(result[0]![0]).toBe("007");
    expect(typeof result[0]![0]).toBe("string");
  });

  it("default (no option) should preserve leading zeros (preserveLeadingZeros defaults to true)", () => {
    const result = parseCsv("0123", { typeInference: true });
    expect(result[0]![0]).toBe("0123");
    expect(typeof result[0]![0]).toBe("string");
  });

  it('"0.123" should become number 0.123 with leading zero preservation', () => {
    const result = parseCsv("0.123", { typeInference: true, preserveLeadingZeros: true });
    expect(result[0]![0]).toBe(0.123);
    expect(typeof result[0]![0]).toBe("number");
  });

  it("should handle mixed values in a row", () => {
    const result = parseCsv("0123,456,007,0.5,hello", {
      typeInference: true,
      preserveLeadingZeros: true,
    });
    const row = result[0]!;
    expect(row[0]).toBe("0123"); // leading zero preserved
    expect(row[1]).toBe(456); // regular number
    expect(row[2]).toBe("007"); // leading zero preserved
    expect(row[3]).toBe(0.5); // decimal starting with 0
    expect(row[4]).toBe("hello"); // string
  });
});

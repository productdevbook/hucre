import { describe, expect, it } from "vitest";
import { readXml, writeXml } from "../src/xml";

describe("writeXml", () => {
  it("emits a basic root/row structure", () => {
    const out = writeXml([{ sku: "P1" }, { sku: "P2" }]);
    expect(out).toContain("<row>");
    expect(out).toContain("<sku>P1</sku>");
    expect(out).toContain("<sku>P2</sku>");
  });

  it("uses custom rootTag and rowTag", () => {
    const out = writeXml([{ name: "A" }], { rootTag: "Catalog", rowTag: "Product" });
    expect(out).toContain("<Catalog>");
    expect(out).toContain("<Product>");
    expect(out).toContain("<name>A</name>");
  });

  it("emits @-prefixed keys as attributes", () => {
    const out = writeXml([{ "@code": "P1", name: "Oak" }], { rowTag: "Product" });
    expect(out).toContain('<Product code="P1">');
    expect(out).toContain("<name>Oak</name>");
  });

  it("reconstructs nested elements from dot-path keys", () => {
    const out = writeXml([{ sku: "P1", "Pricing.Cost": 100, "Pricing.Retail": 180 }], {
      rowTag: "Product",
      pretty: true,
    });
    expect(out).toContain("<Pricing>");
    expect(out).toContain("<Cost>100</Cost>");
    expect(out).toContain("<Retail>180</Retail>");
  });

  it("emits self-closing for null values without children", () => {
    const out = writeXml([{ name: null }]);
    expect(out).toContain("<name/>");
  });

  it("escapes special characters in text", () => {
    const out = writeXml([{ note: "A & B < C" }]);
    expect(out).toContain("A &amp; B &lt; C");
  });

  it("escapes special characters in attribute values", () => {
    const out = writeXml([{ "@title": 'He said "hi"' }]);
    expect(out).toContain('title="He said &quot;hi&quot;"');
  });

  it("optionally omits XML declaration", () => {
    const out = writeXml([{ a: 1 }], { declaration: false });
    expect(out.startsWith("<?xml")).toBe(false);
  });

  it("pretty-prints with indentation", () => {
    const out = writeXml([{ a: 1 }], { pretty: true });
    expect(out).toMatch(/\n {2}<row>/);
  });

  it("converts Date values to ISO 8601", () => {
    const d = new Date("2025-04-25T00:00:00Z");
    const out = writeXml([{ at: d }]);
    expect(out).toContain(`<at>${d.toISOString()}</at>`);
  });

  it("throws when a key is not a valid XML name", () => {
    expect(() => writeXml([{ "has spaces": 1 }])).toThrow(/Invalid XML name/);
  });

  it("round-trips through readXml for simple shapes", () => {
    const data = [
      { "@code": "P1", name: "Oak", "Pricing.Cost": 100 },
      { "@code": "P2", name: "Pine", "Pricing.Cost": 90 },
    ];
    const out = writeXml(data, { rootTag: "Catalog", rowTag: "Product" });
    const back = readXml(out);

    expect(back.rowTag).toBe("Product");
    expect(back.data).toEqual([
      { "@code": "P1", name: "Oak", "Pricing.Cost": "100" },
      { "@code": "P2", name: "Pine", "Pricing.Cost": "90" },
    ]);
  });
});

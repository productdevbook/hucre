import { describe, expect, it } from "vitest";
import { readXml } from "../src/xml";

describe("readXml — auto-detect", () => {
  it("auto-detects the most-frequent repeating row tag", () => {
    const xml = `
      <Catalog>
        <Product><sku>P1</sku></Product>
        <Product><sku>P2</sku></Product>
        <Header>meta</Header>
      </Catalog>
    `;
    const r = readXml(xml);
    expect(r.rowTag).toBe("Product");
    expect(r.data).toEqual([{ sku: "P1" }, { sku: "P2" }]);
  });

  it("uses first-seen tag when ties happen", () => {
    const xml = `<Root><A/><B/><A/><B/></Root>`;
    const r = readXml(xml);
    expect(r.rowTag).toBe("A");
  });

  it("treats a single non-repeating child as one row", () => {
    const xml = `<Root><Item><x>1</x></Item></Root>`;
    const r = readXml(xml);
    expect(r.data).toEqual([{ x: "1" }]);
  });

  it("returns empty data for empty input", () => {
    expect(readXml("")).toEqual({ data: [], headers: [], rowTag: "" });
  });

  it("throws when root has no child elements", () => {
    expect(() => readXml("<Root>text only</Root>")).toThrow(/no child elements/);
  });
});

describe("readXml — flatten", () => {
  it("flattens nested elements with dot-path keys", () => {
    const xml = `
      <Catalog>
        <Product>
          <sku>P1</sku>
          <Pricing>
            <Cost>100</Cost>
            <Retail>180</Retail>
          </Pricing>
        </Product>
      </Catalog>
    `;
    const r = readXml(xml);
    expect(r.data[0]).toEqual({
      sku: "P1",
      "Pricing.Cost": "100",
      "Pricing.Retail": "180",
    });
  });

  it("prefixes attributes with @ by default", () => {
    const xml = `<Root><Product code="P1"><Name>Oak</Name></Product></Root>`;
    const r = readXml(xml);
    expect(r.data[0]).toEqual({
      "@code": "P1",
      Name: "Oak",
    });
  });

  it("nested element attributes get dot-path + @ prefix", () => {
    const xml = `
      <Root>
        <Product>
          <Pricing currency="USD"><Cost>100</Cost></Pricing>
        </Product>
      </Root>
    `;
    const r = readXml(xml);
    expect(r.data[0]).toEqual({
      "Pricing.@currency": "USD",
      "Pricing.Cost": "100",
    });
  });

  it("custom attrPrefix", () => {
    const xml = `<Root><Item id="1"/><Item id="2"/></Root>`;
    const r = readXml(xml, { attrPrefix: "_" });
    expect(r.headers).toEqual(["_id"]);
    expect(r.data).toEqual([{ _id: "1" }, { _id: "2" }]);
  });

  it("flatten:false stringifies nested children", () => {
    const xml = `
      <Root>
        <Product>
          <sku>P1</sku>
          <Pricing><Cost>100</Cost></Pricing>
        </Product>
      </Root>
    `;
    const r = readXml(xml, { flatten: false });
    const row = r.data[0]!;
    expect(row.sku).toBe("P1");
    expect(typeof row["#text"]).toBe("undefined");
    // Pricing serialized as JSON
    const parsed = JSON.parse(row["Pricing"] as string) as Record<string, unknown>;
    expect(parsed.Cost).toBe("100");
  });
});

describe("readXml — namespaces", () => {
  it("keeps prefixed tag names by default", () => {
    const xml = `
      <ns:Catalog xmlns:ns="urn:test">
        <ns:Product><ns:sku>P1</ns:sku></ns:Product>
        <ns:Product><ns:sku>P2</ns:sku></ns:Product>
      </ns:Catalog>
    `;
    const r = readXml(xml);
    expect(r.rowTag).toBe("ns:Product");
    expect(r.headers).toEqual(["ns:sku"]);
  });

  it("strips namespaces when stripNamespaces:true", () => {
    const xml = `
      <ns:Catalog xmlns:ns="urn:test">
        <ns:Product><ns:sku>P1</ns:sku></ns:Product>
        <ns:Product><ns:sku>P2</ns:sku></ns:Product>
      </ns:Catalog>
    `;
    const r = readXml(xml, { stripNamespaces: true });
    expect(r.rowTag).toBe("Product");
    expect(r.headers).toEqual(["sku"]);
    expect(r.data).toEqual([{ sku: "P1" }, { sku: "P2" }]);
  });
});

describe("readXml — special inputs", () => {
  it("decodes XML entities in text content", () => {
    const xml = `<Root><Item><name>A &amp; B</name></Item><Item><name>C &lt; D</name></Item></Root>`;
    const r = readXml(xml);
    expect(r.data).toEqual([{ name: "A & B" }, { name: "C < D" }]);
  });

  it("preserves CDATA content", () => {
    const xml = `<Root><Item><body><![CDATA[<html>hi</html>]]></body></Item><Item><body>plain</body></Item></Root>`;
    const r = readXml(xml);
    expect(r.data[0]!.body).toBe("<html>hi</html>");
    expect(r.data[1]!.body).toBe("plain");
  });

  it("self-closing leaf becomes null", () => {
    const xml = `<Root><Item><Name/></Item><Item><Name>X</Name></Item></Root>`;
    const r = readXml(xml);
    expect(r.data[0]!.Name).toBeNull();
    expect(r.data[1]!.Name).toBe("X");
  });

  it("union of headers across heterogeneous rows", () => {
    const xml = `<Root><Item><a>1</a></Item><Item><b>2</b></Item></Root>`;
    const r = readXml(xml);
    expect(r.headers).toEqual(["a", "b"]);
    expect(r.data).toEqual([
      { a: "1", b: null },
      { a: null, b: "2" },
    ]);
  });
});

describe("readXml — options", () => {
  it("respects explicit rowTag", () => {
    const xml = `<Root><A><x>1</x></A><B><x>2</x></B><A><x>3</x></A></Root>`;
    const r = readXml(xml, { rowTag: "A" });
    expect(r.rowTag).toBe("A");
    expect(r.data).toEqual([{ x: "1" }, { x: "3" }]);
  });

  it("respects maxRows", () => {
    const xml = `<R><I><x>1</x></I><I><x>2</x></I><I><x>3</x></I></R>`;
    const r = readXml(xml, { maxRows: 2 });
    expect(r.data).toHaveLength(2);
  });

  it("applies transformHeader", () => {
    const xml = `<R><I><FirstName>A</FirstName></I></R>`;
    const r = readXml(xml, { transformHeader: (h) => h.toLowerCase() });
    expect(r.headers).toEqual(["firstname"]);
    expect(r.data[0]).toEqual({ firstname: "A" });
  });

  it("applies transformValue", () => {
    const xml = `<R><I><price>10</price></I><I><price>20</price></I></R>`;
    const r = readXml(xml, {
      transformValue: (v, h) => (h === "price" && typeof v === "string" ? Number(v) : v),
    });
    expect(r.data).toEqual([{ price: 10 }, { price: 20 }]);
  });

  it("throws XmlError on malformed XML", () => {
    // Unterminated comment is reliably caught by the SAX parser
    expect(() => readXml("<Root><Item>x</Item><!-- unterminated")).toThrow();
  });
});

import { describe, it, expect } from "vitest";
import { writeOds } from "../src/ods/writer";
import { readOds } from "../src/ods/reader";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import type { Cell, WriteSheet } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");

async function extractFile(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data);
  const raw = await zip.extract(path);
  return decoder.decode(raw);
}

async function parseXmlFromZip(data: Uint8Array, path: string) {
  const xml = await extractFile(data, path);
  return parseXml(xml);
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

async function getAutomaticStyles(data: Uint8Array) {
  const contentDoc = await parseXmlFromZip(data, "content.xml");
  return findChild(contentDoc, "automatic-styles");
}

function singleCellSheet(numFmt: string, value: number): WriteSheet {
  const cells = new Map<string, Partial<Cell>>();
  cells.set("0,0", { value, style: { numFmt } });
  return { name: "Sheet1", rows: [[value]], cells };
}

// ── Writer: data style emission ─────────────────────────────────────

describe("ODS writer — numFmt data styles", () => {
  it("emits a number-style for a plain numeric format", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("0.00", 3.14)] });
    const autoStyles = await getAutomaticStyles(data);

    const numberStyle = findChild(autoStyles, "number-style");
    expect(numberStyle).toBeDefined();
    expect(typeof numberStyle.attrs["style:name"]).toBe("string");

    const numEl = findChild(numberStyle, "number");
    expect(numEl.attrs["number:decimal-places"]).toBe("2");
    expect(numEl.attrs["number:min-integer-digits"]).toBe("1");
    expect(numEl.attrs["number:grouping"]).toBeUndefined();
  });

  it("propagates the thousands-separator into number:grouping", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("#,##0.00", 1234.5)] });
    const autoStyles = await getAutomaticStyles(data);
    const numberStyle = findChild(autoStyles, "number-style");
    const numEl = findChild(numberStyle, "number");
    expect(numEl.attrs["number:grouping"]).toBe("true");
    expect(numEl.attrs["number:decimal-places"]).toBe("2");
  });

  it("emits a percentage-style for percent codes", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("0.00%", 0.5)] });
    const autoStyles = await getAutomaticStyles(data);
    const pctStyle = findChild(autoStyles, "percentage-style");
    expect(pctStyle).toBeDefined();
    const numEl = findChild(pctStyle, "number");
    expect(numEl.attrs["number:decimal-places"]).toBe("2");
    // Trailing literal "%" sits in <number:text>%</number:text>
    const textEls = findChildren(pctStyle, "text");
    const collected = textEls
      .map((t: any) => t.children.filter((c: unknown) => typeof c === "string").join(""))
      .join("");
    expect(collected).toContain("%");
  });

  it("emits a date-style for yyyy-mm-dd", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("yyyy-mm-dd", 45000)] });
    const autoStyles = await getAutomaticStyles(data);
    const dateStyle = findChild(autoStyles, "date-style");
    expect(dateStyle).toBeDefined();

    const year = findChild(dateStyle, "year");
    expect(year).toBeDefined();
    expect(year.attrs["number:style"]).toBe("long");

    // Two month elements? No — there is exactly one `<number:month>` between the
    // two literal `-` separators. We assert it's there with style="long".
    const month = findChild(dateStyle, "month");
    expect(month).toBeDefined();
    expect(month.attrs["number:style"]).toBe("long");

    const day = findChild(dateStyle, "day");
    expect(day).toBeDefined();
    expect(day.attrs["number:style"]).toBe("long");
  });

  it("emits a time-style for hh:mm:ss", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("hh:mm:ss", 0.5)] });
    const autoStyles = await getAutomaticStyles(data);
    const timeStyle = findChild(autoStyles, "time-style");
    expect(timeStyle).toBeDefined();

    expect(findChild(timeStyle, "hours")).toBeDefined();
    expect(findChild(timeStyle, "minutes")).toBeDefined();
    expect(findChild(timeStyle, "seconds")).toBeDefined();

    // No truncate-on-overflow for plain "hh:mm:ss" — it wraps at 24h
    expect(timeStyle.attrs["number:truncate-on-overflow"]).toBeUndefined();
  });

  it("flags duration formats with truncate-on-overflow=false", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("[HH]:MM", 0.020833)] });
    const autoStyles = await getAutomaticStyles(data);
    const timeStyle = findChild(autoStyles, "time-style");
    expect(timeStyle).toBeDefined();
    expect(timeStyle.attrs["number:truncate-on-overflow"]).toBe("false");

    const hours = findChild(timeStyle, "hours");
    expect(hours).toBeDefined();
    const minutes = findChild(timeStyle, "minutes");
    expect(minutes).toBeDefined();
  });

  it("emits a currency-style for $#,##0.00", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("$#,##0.00", 99.95)] });
    const autoStyles = await getAutomaticStyles(data);
    const currencyStyle = findChild(autoStyles, "currency-style");
    expect(currencyStyle).toBeDefined();

    const symbol = findChild(currencyStyle, "currency-symbol");
    expect(symbol).toBeDefined();
    const symbolText = symbol.children.filter((c: unknown) => typeof c === "string").join("");
    expect(symbolText).toBe("$");

    const numEl = findChild(currencyStyle, "number");
    expect(numEl).toBeDefined();
    expect(numEl.attrs["number:decimal-places"]).toBe("2");
    expect(numEl.attrs["number:grouping"]).toBe("true");
  });

  it("references the data style from the cell <style:style> via style:data-style-name", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("0.00", 3.14)] });
    const autoStyles = await getAutomaticStyles(data);

    const numberStyle = findChild(autoStyles, "number-style");
    const dataStyleName = numberStyle.attrs["style:name"];
    expect(dataStyleName).toBeTruthy();

    const cellStyle = findChild(autoStyles, "style");
    expect(cellStyle).toBeDefined();
    expect(cellStyle.attrs["style:family"]).toBe("table-cell");
    expect(cellStyle.attrs["style:data-style-name"]).toBe(dataStyleName);
  });

  it("orders data styles before cell styles inside automatic-styles", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("0.00", 3.14)] });
    const autoStyles = await getAutomaticStyles(data);

    const order: string[] = [];
    for (const child of autoStyles.children) {
      if (typeof child === "string") continue;
      order.push((child as any).local || (child as any).tag);
    }
    const dataIdx = order.indexOf("number-style");
    const styleIdx = order.indexOf("style");
    expect(dataIdx).toBeGreaterThanOrEqual(0);
    expect(styleIdx).toBeGreaterThanOrEqual(0);
    expect(dataIdx).toBeLessThan(styleIdx);
  });

  it("deduplicates styles that share a numFmt", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", { value: 1, style: { numFmt: "0.00" } });
    cells.set("0,1", { value: 2, style: { numFmt: "0.00" } });
    cells.set("0,2", { value: 3, style: { numFmt: "0.00%" } });
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[1, 2, 3]], cells }],
    });
    const autoStyles = await getAutomaticStyles(data);

    expect(findChildren(autoStyles, "number-style")).toHaveLength(1);
    expect(findChildren(autoStyles, "percentage-style")).toHaveLength(1);
    // Two distinct cell styles (number-style vs percentage-style)
    expect(findChildren(autoStyles, "style")).toHaveLength(2);
  });

  it("ignores numFmt = 'General' / '@'", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", { value: 1, style: { numFmt: "General" } });
    cells.set("0,1", { value: 2, style: { numFmt: "@" } });
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[1, 2]], cells }],
    });
    const autoStyles = await getAutomaticStyles(data);
    expect(findChildren(autoStyles, "number-style")).toHaveLength(0);
    expect(findChildren(autoStyles, "percentage-style")).toHaveLength(0);
    expect(findChildren(autoStyles, "date-style")).toHaveLength(0);
    expect(findChildren(autoStyles, "time-style")).toHaveLength(0);
    // The cell-level <style:style> is also dropped because it carries no
    // visual properties — same behaviour as an unstyled cell.
    expect(findChildren(autoStyles, "style")).toHaveLength(0);
  });

  it("co-exists with font and fill properties on the same cell style", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: 0.5,
      style: {
        numFmt: "0.00%",
        font: { bold: true },
        fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } },
      },
    });
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[0.5]], cells }],
    });
    const autoStyles = await getAutomaticStyles(data);

    const pctStyle = findChild(autoStyles, "percentage-style");
    expect(pctStyle).toBeDefined();

    const cellStyle = findChild(autoStyles, "style");
    expect(cellStyle.attrs["style:data-style-name"]).toBe(pctStyle.attrs["style:name"]);
    expect(findChild(cellStyle, "text-properties").attrs["fo:font-weight"]).toBe("bold");
    expect(findChild(cellStyle, "table-cell-properties").attrs["fo:background-color"]).toBe(
      "#FFFF00",
    );
  });
});

// ── Roundtrip: write → read with readStyles ─────────────────────────

describe("ODS parity — numFmt roundtrip", () => {
  it("plain number format roundtrips through write/read", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("0.00", 3.14)] });
    const wb = await readOds(data, { readStyles: true });
    const cell = wb.sheets[0].cells!.get("0,0");
    expect(cell?.style?.numFmt).toBe("0.00");
  });

  it("grouping number format roundtrips", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("#,##0.00", 1234.5)] });
    const wb = await readOds(data, { readStyles: true });
    expect(wb.sheets[0].cells!.get("0,0")?.style?.numFmt).toBe("#,##0.00");
  });

  it("percentage format roundtrips", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("0.00%", 0.5)] });
    const wb = await readOds(data, { readStyles: true });
    expect(wb.sheets[0].cells!.get("0,0")?.style?.numFmt).toBe("0.00%");
  });

  it("date format yyyy-mm-dd roundtrips", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("yyyy-mm-dd", 45000)] });
    const wb = await readOds(data, { readStyles: true });
    expect(wb.sheets[0].cells!.get("0,0")?.style?.numFmt).toBe("yyyy-mm-dd");
  });

  it("time format hh:mm:ss roundtrips", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("hh:mm:ss", 0.5)] });
    const wb = await readOds(data, { readStyles: true });
    expect(wb.sheets[0].cells!.get("0,0")?.style?.numFmt).toBe("hh:mm:ss");
  });

  it("duration format [HH]:MM roundtrips", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("[HH]:MM", 0.020833)] });
    const wb = await readOds(data, { readStyles: true });
    // Output normalises to lowercase brackets and short minute token
    const code = wb.sheets[0].cells!.get("0,0")?.style?.numFmt;
    expect(code).toBeDefined();
    expect(code!.startsWith("[h")).toBe(true);
    expect(code!.toLowerCase()).toContain("m");
  });

  it("readStyles=false keeps numFmt out of the cell metadata", async () => {
    const data = await writeOds({ sheets: [singleCellSheet("0.00%", 0.5)] });
    const wb = await readOds(data); // default: readStyles=false
    const cell = wb.sheets[0].cells?.get("0,0");
    // No style key collected when readStyles is off
    expect(cell?.style?.numFmt).toBeUndefined();
  });
});

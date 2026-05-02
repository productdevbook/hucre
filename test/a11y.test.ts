import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { audit, contrastRatio, relativeLuminance, applyA11ySummary } from "../src/a11y";
import type { WriteOptions, Workbook } from "../src/_types";

const decoder = new TextDecoder("utf-8");

async function extractText(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data);
  return decoder.decode(await zip.extract(path));
}

function fakePng(size = 64): Uint8Array {
  const d = new Uint8Array(size);
  d[0] = 0x89;
  d[1] = 0x50;
  d[2] = 0x4e;
  d[3] = 0x47;
  d[4] = 0x0d;
  d[5] = 0x0a;
  d[6] = 0x1a;
  d[7] = 0x0a;
  for (let i = 8; i < size; i++) d[i] = i % 256;
  return d;
}

// ── Color helpers ──────────────────────────────────────────────────

describe("a11y.relativeLuminance", () => {
  it("returns 0 for pure black", () => {
    expect(relativeLuminance("000000")).toBeCloseTo(0, 5);
  });
  it("returns 1 for pure white", () => {
    expect(relativeLuminance("FFFFFF")).toBeCloseTo(1, 5);
  });
  it("accepts a leading hash", () => {
    expect(relativeLuminance("#FFFFFF")).toBeCloseTo(1, 5);
  });
  it("accepts 8-digit ARGB by stripping the alpha prefix", () => {
    expect(relativeLuminance("FFFFFFFF")).toBeCloseTo(1, 5);
  });
  it("expands 3-digit shorthand", () => {
    expect(relativeLuminance("FFF")).toBeCloseTo(1, 5);
  });
  it("returns 0 for malformed input", () => {
    expect(relativeLuminance("zzzzzz")).toBeCloseTo(0, 5);
  });
});

describe("a11y.contrastRatio", () => {
  it("returns 21:1 for black-on-white", () => {
    expect(contrastRatio("000000", "FFFFFF")).toBeCloseTo(21, 1);
  });
  it("returns 1:1 for identical colors", () => {
    expect(contrastRatio("808080", "808080")).toBeCloseTo(1, 5);
  });
  it("is symmetric (fg/bg order does not matter)", () => {
    expect(contrastRatio("336699", "FFFFFF")).toBeCloseTo(contrastRatio("FFFFFF", "336699"), 5);
  });
  it("flags well-known low-contrast pairs as below WCAG AA", () => {
    // Light gray on white — classic accessibility offender.
    expect(contrastRatio("AAAAAA", "FFFFFF")).toBeLessThan(4.5);
  });
  it("approves a known AA-passing pair", () => {
    // GitHub's #0969da link blue on white is well above 4.5:1.
    expect(contrastRatio("0969DA", "FFFFFF")).toBeGreaterThan(4.5);
  });
});

// ── audit() ────────────────────────────────────────────────────────

describe("a11y.audit — workbook-level", () => {
  it("flags missing document description", () => {
    const wb: Workbook = {
      sheets: [{ name: "S", rows: [["a"]] }],
    };
    const issues = audit(wb);
    expect(issues.some((i) => i.code === "no-doc-description")).toBe(true);
  });

  it("does not flag missing description when a sheet supplies a summary", () => {
    const wb: Workbook = {
      sheets: [{ name: "S", rows: [["a"]], a11y: { summary: "Quarterly report" } }],
    };
    const issues = audit(wb);
    expect(issues.some((i) => i.code === "no-doc-description")).toBe(false);
  });

  it("emits info (not warning) for missing title", () => {
    const wb: Workbook = {
      sheets: [{ name: "S", rows: [["a"]] }],
      properties: { description: "x" },
    };
    const issues = audit(wb);
    const titleIssue = issues.find((i) => i.code === "no-doc-title");
    expect(titleIssue?.type).toBe("info");
  });
});

describe("a11y.audit — sheet-level", () => {
  it("warns when a populated sheet has no header row marked", () => {
    const wb: Workbook = {
      sheets: [{ name: "Data", rows: [["a", "b"]], a11y: { summary: "x" } }],
      properties: { title: "t", description: "d" },
    };
    const issue = audit(wb).find((i) => i.code === "no-header-row");
    expect(issue?.type).toBe("warning");
    expect(issue?.location?.sheet).toBe("Data");
  });

  it("does not warn when an Excel table covers the data", () => {
    const wb: Workbook = {
      sheets: [
        {
          name: "T",
          rows: [
            ["h1", "h2"],
            ["a", "b"],
          ],
          tables: [
            {
              name: "T1",
              range: "A1:B2",
              columns: [{ name: "h1" }, { name: "h2" }],
            },
          ],
        },
      ],
      properties: { title: "t", description: "d" },
    };
    const issues = audit(wb);
    expect(issues.some((i) => i.code === "no-header-row")).toBe(false);
  });

  it("flags an empty sheet", () => {
    const wb: Workbook = {
      sheets: [{ name: "Empty", rows: [] }],
      properties: { title: "t", description: "d" },
    };
    const issue = audit(wb).find((i) => i.code === "empty-sheet");
    expect(issue).toBeDefined();
    expect(issue?.type).toBe("info");
  });

  it("flags a merged cell overlapping the marked header row", () => {
    const wb: Workbook = {
      sheets: [
        {
          name: "S",
          rows: [
            ["h", ""],
            ["a", "b"],
          ],
          a11y: { headerRow: 0 },
          merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
        },
      ],
      properties: { title: "t", description: "d" },
    };
    const issue = audit(wb).find((i) => i.code === "merged-header-row");
    expect(issue).toBeDefined();
    expect(issue?.location?.ref).toBe("A1:B1");
  });

  it("does not flag a merged cell outside the header row", () => {
    const wb: Workbook = {
      sheets: [
        {
          name: "S",
          rows: [
            ["h1", "h2"],
            ["a", "b"],
            ["c", "d"],
          ],
          a11y: { headerRow: 0 },
          merges: [{ startRow: 1, startCol: 0, endRow: 2, endCol: 0 }],
        },
      ],
      properties: { title: "t", description: "d" },
    };
    const issues = audit(wb);
    expect(issues.some((i) => i.code === "merged-header-row")).toBe(false);
  });

  it("flags blank rows in the middle of populated data", () => {
    const wb: Workbook = {
      sheets: [
        {
          name: "S",
          rows: [["a"], [], ["b"]],
          a11y: { headerRow: 0 },
        },
      ],
      properties: { title: "t", description: "d" },
    };
    const issue = audit(wb).find((i) => i.code === "blank-row-in-data");
    expect(issue?.location?.ref).toBe("2:2");
  });

  it("does not flag trailing blank rows", () => {
    const wb: Workbook = {
      sheets: [
        {
          name: "S",
          rows: [["a"], ["b"], []],
          a11y: { headerRow: 0 },
        },
      ],
      properties: { title: "t", description: "d" },
    };
    const issues = audit(wb);
    expect(issues.some((i) => i.code === "blank-row-in-data")).toBe(false);
  });
});

describe("a11y.audit — images", () => {
  it("errors on an image with no altText", () => {
    const wb: Workbook = {
      sheets: [
        {
          name: "S",
          rows: [["x"]],
          a11y: { headerRow: 0 },
          images: [
            {
              data: fakePng(),
              type: "png",
              anchor: { from: { row: 4, col: 1 } },
            },
          ],
        },
      ],
      properties: { title: "t", description: "d" },
    };
    const issue = audit(wb).find((i) => i.code === "missing-alt-text");
    expect(issue?.type).toBe("error");
    expect(issue?.location?.ref).toBe("B5");
    expect(issue?.location?.image).toBe(0);
  });

  it("does not error when altText is present", () => {
    const wb: Workbook = {
      sheets: [
        {
          name: "S",
          rows: [["x"]],
          a11y: { headerRow: 0 },
          images: [
            {
              data: fakePng(),
              type: "png",
              anchor: { from: { row: 0, col: 0 } },
              altText: "Sales chart",
            },
          ],
        },
      ],
      properties: { title: "t", description: "d" },
    };
    const issues = audit(wb);
    expect(issues.some((i) => i.code === "missing-alt-text")).toBe(false);
  });
});

describe("a11y.audit — color contrast", () => {
  it("flags low-contrast cells via cell.style.font.color and cell.style.fill.fgColor", () => {
    const cells = new Map();
    cells.set("0,0", {
      value: "low",
      style: {
        font: { color: { rgb: "AAAAAA" } },
        fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFFFF" } },
      },
    });
    const wb: Workbook = {
      sheets: [
        {
          name: "S",
          rows: [["low"]],
          cells,
          a11y: { headerRow: 0 },
        },
      ],
      properties: { title: "t", description: "d" },
    };
    const issue = audit(wb).find((i) => i.code === "low-contrast");
    expect(issue?.type).toBe("warning");
    expect(issue?.location?.ref).toBe("A1");
  });

  it("does not flag high-contrast cells", () => {
    const cells = new Map();
    cells.set("0,0", {
      value: "ok",
      style: {
        font: { color: { rgb: "000000" } },
        fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFFFF" } },
      },
    });
    const wb: Workbook = {
      sheets: [
        {
          name: "S",
          rows: [["ok"]],
          cells,
          a11y: { headerRow: 0 },
        },
      ],
      properties: { title: "t", description: "d" },
    };
    const issues = audit(wb);
    expect(issues.some((i) => i.code === "low-contrast")).toBe(false);
  });

  it("respects skipContrast", () => {
    const cells = new Map();
    cells.set("0,0", {
      value: "low",
      style: {
        font: { color: { rgb: "AAAAAA" } },
        fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFFFF" } },
      },
    });
    const wb: Workbook = {
      sheets: [{ name: "S", rows: [["low"]], cells, a11y: { headerRow: 0 } }],
      properties: { title: "t", description: "d" },
    };
    const issues = audit(wb, { skipContrast: true });
    expect(issues.some((i) => i.code === "low-contrast")).toBe(false);
  });
});

// ── applyA11ySummary ───────────────────────────────────────────────

describe("a11y.applyA11ySummary", () => {
  it("copies the first sheet summary to workbook description", () => {
    const wb: Workbook = {
      sheets: [
        { name: "S1", rows: [["a"]], a11y: { summary: "Q1 figures" } },
        { name: "S2", rows: [["b"]] },
      ],
    };
    applyA11ySummary(wb);
    expect(wb.properties?.description).toBe("Q1 figures");
  });

  it("does not overwrite an explicit description", () => {
    const wb: Workbook = {
      sheets: [{ name: "S", rows: [["a"]], a11y: { summary: "ignore me" } }],
      properties: { description: "explicit" },
    };
    applyA11ySummary(wb);
    expect(wb.properties?.description).toBe("explicit");
  });

  it("is a no-op when no sheet has a summary", () => {
    const wb: Workbook = {
      sheets: [{ name: "S", rows: [["a"]] }],
    };
    applyA11ySummary(wb);
    expect(wb.properties?.description).toBeUndefined();
  });
});

// ── End-to-end: written file actually carries the metadata ─────────

describe("writeXlsx — a11y integration", () => {
  it("emits descr= and title= on xdr:cNvPr for images with altText/title", async () => {
    const opts: WriteOptions = {
      sheets: [
        {
          name: "S",
          rows: [["x"]],
          images: [
            {
              data: fakePng(),
              type: "png",
              anchor: { from: { row: 0, col: 0 } },
              altText: "Bar chart of revenue & cost",
              title: "Revenue",
            },
          ],
        },
      ],
    };
    const out = await writeXlsx(opts);
    const drawing = await extractText(out, "xl/drawings/drawing1.xml");
    // Ampersands inside attributes must remain escaped (&amp;) so the file is well-formed.
    expect(drawing).toContain('descr="Bar chart of revenue &amp; cost"');
    expect(drawing).toContain('title="Revenue"');
  });

  it("promotes the first sheet a11y.summary into docProps/core.xml when no description is set", async () => {
    const opts: WriteOptions = {
      sheets: [{ name: "S1", rows: [["a"]], a11y: { summary: "Quarterly sales report" } }],
    };
    const out = await writeXlsx(opts);
    const core = await extractText(out, "docProps/core.xml");
    expect(core).toContain("Quarterly sales report");
  });

  it("does not override an explicit workbook description", async () => {
    const opts: WriteOptions = {
      sheets: [{ name: "S1", rows: [["a"]], a11y: { summary: "from sheet" } }],
      properties: { description: "from properties" },
    };
    const out = await writeXlsx(opts);
    const core = await extractText(out, "docProps/core.xml");
    expect(core).toContain("from properties");
    expect(core).not.toContain("from sheet");
  });
});

// ── Roundtrip: alt text / title survive read → re-read ─────────────

describe("readXlsx — drawing alt text / title roundtrip", () => {
  it("recovers altText and title from xdr:cNvPr on images", async () => {
    const opts: WriteOptions = {
      sheets: [
        {
          name: "S",
          rows: [["x"]],
          images: [
            {
              data: fakePng(),
              type: "png",
              anchor: { from: { row: 0, col: 0 } },
              altText: "Bar chart of Q1 revenue",
              title: "Q1 Revenue",
            },
          ],
        },
      ],
    };
    const out = await writeXlsx(opts);
    const wb = await readXlsx(out);
    const img = wb.sheets[0].images?.[0];
    expect(img).toBeDefined();
    expect(img?.altText).toBe("Bar chart of Q1 revenue");
    expect(img?.title).toBe("Q1 Revenue");
  });

  it("recovers altText and title from xdr:cNvPr on text boxes", async () => {
    const opts: WriteOptions = {
      sheets: [
        {
          name: "S",
          rows: [["x"]],
          textBoxes: [
            {
              text: "Note",
              anchor: {
                from: { row: 0, col: 0 },
                to: { row: 2, col: 2 },
              },
              altText: "Disclaimer about quarterly figures",
              title: "Disclaimer",
            },
          ],
        },
      ],
    };
    const out = await writeXlsx(opts);
    const wb = await readXlsx(out);
    const tb = wb.sheets[0].textBoxes?.[0];
    expect(tb).toBeDefined();
    expect(tb?.altText).toBe("Disclaimer about quarterly figures");
    expect(tb?.title).toBe("Disclaimer");
  });

  it("leaves altText/title undefined when the source XML has no descr/title", async () => {
    // Image written without altText/title — both should remain absent on re-read.
    const opts: WriteOptions = {
      sheets: [
        {
          name: "S",
          rows: [["x"]],
          images: [
            {
              data: fakePng(),
              type: "png",
              anchor: { from: { row: 0, col: 0 } },
            },
          ],
        },
      ],
    };
    const out = await writeXlsx(opts);
    const wb = await readXlsx(out);
    const img = wb.sheets[0].images?.[0];
    expect(img?.altText).toBeUndefined();
    expect(img?.title).toBeUndefined();
  });
});

import { describe, expect, it } from "vitest";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { streamXlsxRows } from "../src/xlsx/stream-reader";
import type { StreamRow } from "../src/xlsx/stream-reader";
import { XlsxStreamWriter } from "../src/xlsx/stream-writer";
import { streamCsvRows, CsvStreamWriter } from "../src/csv/stream";
import { parseCsv } from "../src/csv/reader";
import { writeCsv } from "../src/csv/writer";
import type { CellValue, WriteSheet } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

async function collectStreamRows(
  gen: AsyncGenerator<StreamRow, void, undefined>,
): Promise<StreamRow[]> {
  const rows: StreamRow[] = [];
  for await (const row of gen) {
    rows.push(row);
  }
  return rows;
}

function collectSyncRows(gen: Generator<CellValue[], void, undefined>): CellValue[][] {
  const rows: CellValue[][] = [];
  for (const row of gen) {
    rows.push(row);
  }
  return rows;
}

/** Create a simple test XLSX via the regular writer */
async function createTestXlsx(sheets: WriteSheet[]): Promise<Uint8Array> {
  return writeXlsx({ sheets });
}

// ═══════════════════════════════════════════════════════════════════════
// XLSX Stream Reader
// ═══════════════════════════════════════════════════════════════════════

describe("streamXlsxRows", () => {
  it("streams rows from basic XLSX", async () => {
    const xlsx = await createTestXlsx([
      {
        name: "Sheet1",
        rows: [
          ["Name", "Age", "Active"],
          ["Alice", 30, true],
          ["Bob", 25, false],
        ],
      },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(xlsx));

    expect(rows).toHaveLength(3);
    expect(rows[0].index).toBe(0);
    expect(rows[0].values).toEqual(["Name", "Age", "Active"]);
    expect(rows[1].index).toBe(1);
    expect(rows[1].values).toEqual(["Alice", 30, true]);
    expect(rows[2].index).toBe(2);
    expect(rows[2].values).toEqual(["Bob", 25, false]);
  });

  it("yields correct 0-based row indices", async () => {
    const xlsx = await createTestXlsx([
      {
        name: "Sheet1",
        rows: [["Row0"], ["Row1"], ["Row2"], ["Row3"], ["Row4"]],
      },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(xlsx));

    expect(rows).toHaveLength(5);
    for (let i = 0; i < 5; i++) {
      expect(rows[i].index).toBe(i);
    }
  });

  it("cell values match written data", async () => {
    const testRows: CellValue[][] = [
      ["Hello", 42, true, null, "World"],
      [null, 3.14, false, "Test", null],
      ["A", "B", "C", "D", "E"],
    ];

    const xlsx = await createTestXlsx([{ name: "Data", rows: testRows }]);

    // Streaming read
    const streamRows = await collectStreamRows(streamXlsxRows(xlsx));

    expect(streamRows).toHaveLength(testRows.length);
    for (let i = 0; i < testRows.length; i++) {
      const streamValues = streamRows[i].values;
      const expectedValues = testRows[i];

      // Both should contain the same actual data
      const maxLen = Math.max(streamValues.length, expectedValues.length);
      for (let j = 0; j < maxLen; j++) {
        const sv = j < streamValues.length ? streamValues[j] : null;
        const ev = j < expectedValues.length ? expectedValues[j] : null;
        expect(sv).toEqual(ev);
      }
    }
  });

  it("resolves shared strings correctly", async () => {
    const xlsx = await createTestXlsx([
      {
        name: "Strings",
        rows: [
          ["Apple", "Banana", "Cherry"],
          ["Apple", "Date", "Elderberry"],
          ["Banana", "Banana", "Cherry"],
        ],
      },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(xlsx));

    expect(rows[0].values).toEqual(["Apple", "Banana", "Cherry"]);
    expect(rows[1].values).toEqual(["Apple", "Date", "Elderberry"]);
    expect(rows[2].values).toEqual(["Banana", "Banana", "Cherry"]);
  });

  it("detects date cells via style", async () => {
    const date1 = new Date(Date.UTC(2024, 0, 15)); // Jan 15, 2024
    const date2 = new Date(Date.UTC(2024, 5, 30)); // Jun 30, 2024

    const xlsx = await createTestXlsx([
      {
        name: "Dates",
        rows: [
          ["Date", "Value"],
          [date1, 100],
          [date2, 200],
        ],
      },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(xlsx));

    expect(rows).toHaveLength(3);
    // First row is headers
    expect(rows[0].values[0]).toBe("Date");

    // Date values should be Date objects
    const val1 = rows[1].values[0];
    expect(val1).toBeInstanceOf(Date);
    expect((val1 as Date).getUTCFullYear()).toBe(2024);
    expect((val1 as Date).getUTCMonth()).toBe(0);
    expect((val1 as Date).getUTCDate()).toBe(15);

    const val2 = rows[2].values[0];
    expect(val2).toBeInstanceOf(Date);
    expect((val2 as Date).getUTCFullYear()).toBe(2024);
    expect((val2 as Date).getUTCMonth()).toBe(5);
    expect((val2 as Date).getUTCDate()).toBe(30);
  });

  it("streams specific sheet by index", async () => {
    const xlsx = await createTestXlsx([
      { name: "First", rows: [["Sheet1Data"]] },
      { name: "Second", rows: [["Sheet2Data"]] },
      { name: "Third", rows: [["Sheet3Data"]] },
    ]);

    // Stream the second sheet (index 1)
    const rows = await collectStreamRows(streamXlsxRows(xlsx, { sheet: 1 }));

    expect(rows).toHaveLength(1);
    expect(rows[0].values[0]).toBe("Sheet2Data");
  });

  it("streams specific sheet by name", async () => {
    const xlsx = await createTestXlsx([
      { name: "Alpha", rows: [["AlphaData"]] },
      { name: "Beta", rows: [["BetaData"]] },
      { name: "Gamma", rows: [["GammaData"]] },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(xlsx, { sheet: "Gamma" }));

    expect(rows).toHaveLength(1);
    expect(rows[0].values[0]).toBe("GammaData");
  });

  it("handles large sheet (5000 rows) without issues", async () => {
    const largeRows: CellValue[][] = [];
    for (let i = 0; i < 5000; i++) {
      largeRows.push([`Row${i}`, i, i % 2 === 0]);
    }

    const xlsx = await createTestXlsx([{ name: "Large", rows: largeRows }]);

    let count = 0;
    for await (const row of streamXlsxRows(xlsx)) {
      expect(row.index).toBe(count);
      expect(row.values[0]).toBe(`Row${count}`);
      expect(row.values[1]).toBe(count);
      count++;
    }

    expect(count).toBe(5000);
  });

  it("empty sheet yields no rows", async () => {
    const xlsx = await createTestXlsx([{ name: "Empty", rows: [] }]);

    const rows = await collectStreamRows(streamXlsxRows(xlsx));

    expect(rows).toHaveLength(0);
  });

  it("yields no rows for non-existent sheet name", async () => {
    const xlsx = await createTestXlsx([{ name: "Sheet1", rows: [["data"]] }]);

    const rows = await collectStreamRows(streamXlsxRows(xlsx, { sheet: "NonExistent" }));

    expect(rows).toHaveLength(0);
  });

  it("defaults to first sheet when no sheet option given", async () => {
    const xlsx = await createTestXlsx([
      { name: "First", rows: [["FirstData"]] },
      { name: "Second", rows: [["SecondData"]] },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(xlsx));

    expect(rows).toHaveLength(1);
    expect(rows[0].values[0]).toBe("FirstData");
  });

  it("handles mixed types in cells", async () => {
    const xlsx = await createTestXlsx([
      {
        name: "Mixed",
        rows: [
          ["text", 42, true, null, 3.14],
          [null, null, false, "hello", 0],
        ],
      },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(xlsx));

    expect(rows[0].values).toEqual(["text", 42, true, null, 3.14]);
    expect(rows[1].values).toEqual([null, null, false, "hello", 0]);
  });

  it("accepts ArrayBuffer input", async () => {
    const xlsx = await createTestXlsx([{ name: "Sheet1", rows: [["test"]] }]);

    // Convert Uint8Array to ArrayBuffer
    const arrayBuffer = xlsx.buffer.slice(
      xlsx.byteOffset,
      xlsx.byteOffset + xlsx.byteLength,
    ) as ArrayBuffer;

    const rows = await collectStreamRows(streamXlsxRows(arrayBuffer));

    expect(rows).toHaveLength(1);
    expect(rows[0].values[0]).toBe("test");
  });
});

// ═══════════════════════════════════════════════════════════════════════
// XLSX Stream Reader — ReadableStream Input
// ═══════════════════════════════════════════════════════════════════════

/** Convert Uint8Array to ReadableStream<Uint8Array> for testing */
function toReadableStream(data: Uint8Array): ReadableStream<Uint8Array> {
  return new ReadableStream({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  });
}

/** Convert Uint8Array to ReadableStream with small chunks for testing chunk boundaries */
function toChunkedReadableStream(data: Uint8Array, chunkSize: number): ReadableStream<Uint8Array> {
  let offset = 0;
  return new ReadableStream({
    pull(controller) {
      if (offset >= data.length) {
        controller.close();
        return;
      }
      const end = Math.min(offset + chunkSize, data.length);
      controller.enqueue(data.subarray(offset, end));
      offset = end;
    },
  });
}

describe("streamXlsxRows — ReadableStream input", () => {
  it("streams rows from ReadableStream", async () => {
    const xlsx = await createTestXlsx([
      {
        name: "Sheet1",
        rows: [
          ["Name", "Age", "Active"],
          ["Alice", 30, true],
          ["Bob", 25, false],
        ],
      },
    ]);

    const stream = toReadableStream(xlsx);
    const rows = await collectStreamRows(streamXlsxRows(stream));

    expect(rows).toHaveLength(3);
    expect(rows[0].values).toEqual(["Name", "Age", "Active"]);
    expect(rows[1].values).toEqual(["Alice", 30, true]);
    expect(rows[2].values).toEqual(["Bob", 25, false]);
  });

  it("ReadableStream output matches Uint8Array output", async () => {
    const xlsx = await createTestXlsx([
      {
        name: "Data",
        rows: [
          ["Hello", 42, true, null, "World"],
          [null, 3.14, false, "Test", null],
          ["A", "B", "C", "D", "E"],
        ],
      },
    ]);

    const uint8Rows = await collectStreamRows(streamXlsxRows(xlsx));
    const streamRows = await collectStreamRows(streamXlsxRows(toReadableStream(xlsx)));

    expect(streamRows).toHaveLength(uint8Rows.length);
    for (let i = 0; i < uint8Rows.length; i++) {
      expect(streamRows[i].index).toBe(uint8Rows[i].index);
      expect(streamRows[i].values).toEqual(uint8Rows[i].values);
    }
  });

  it("resolves shared strings from ReadableStream", async () => {
    const xlsx = await createTestXlsx([
      {
        name: "Strings",
        rows: [
          ["Apple", "Banana", "Cherry"],
          ["Apple", "Date", "Elderberry"],
          ["Banana", "Banana", "Cherry"],
        ],
      },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(toReadableStream(xlsx)));

    expect(rows[0].values).toEqual(["Apple", "Banana", "Cherry"]);
    expect(rows[1].values).toEqual(["Apple", "Date", "Elderberry"]);
    expect(rows[2].values).toEqual(["Banana", "Banana", "Cherry"]);
  });

  it("detects date cells from ReadableStream", async () => {
    const date1 = new Date(Date.UTC(2024, 0, 15));
    const date2 = new Date(Date.UTC(2024, 5, 30));

    const xlsx = await createTestXlsx([
      {
        name: "Dates",
        rows: [
          ["Date", "Value"],
          [date1, 100],
          [date2, 200],
        ],
      },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(toReadableStream(xlsx)));

    expect(rows).toHaveLength(3);
    const val1 = rows[1].values[0];
    expect(val1).toBeInstanceOf(Date);
    expect((val1 as Date).getUTCFullYear()).toBe(2024);
    expect((val1 as Date).getUTCMonth()).toBe(0);
    expect((val1 as Date).getUTCDate()).toBe(15);

    const val2 = rows[2].values[0];
    expect(val2).toBeInstanceOf(Date);
    expect((val2 as Date).getUTCMonth()).toBe(5);
  });

  it("streams specific sheet by name from ReadableStream", async () => {
    const xlsx = await createTestXlsx([
      { name: "Alpha", rows: [["AlphaData"]] },
      { name: "Beta", rows: [["BetaData"]] },
      { name: "Gamma", rows: [["GammaData"]] },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(toReadableStream(xlsx), { sheet: "Beta" }));

    expect(rows).toHaveLength(1);
    expect(rows[0].values[0]).toBe("BetaData");
  });

  it("streams specific sheet by index from ReadableStream", async () => {
    const xlsx = await createTestXlsx([
      { name: "First", rows: [["Sheet1Data"]] },
      { name: "Second", rows: [["Sheet2Data"]] },
      { name: "Third", rows: [["Sheet3Data"]] },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(toReadableStream(xlsx), { sheet: 2 }));

    expect(rows).toHaveLength(1);
    expect(rows[0].values[0]).toBe("Sheet3Data");
  });

  it("handles empty sheet from ReadableStream", async () => {
    const xlsx = await createTestXlsx([{ name: "Empty", rows: [] }]);

    const rows = await collectStreamRows(streamXlsxRows(toReadableStream(xlsx)));

    expect(rows).toHaveLength(0);
  });

  it("handles single row from ReadableStream", async () => {
    const xlsx = await createTestXlsx([{ name: "One", rows: [["only row"]] }]);

    const rows = await collectStreamRows(streamXlsxRows(toReadableStream(xlsx)));

    expect(rows).toHaveLength(1);
    expect(rows[0].index).toBe(0);
    expect(rows[0].values).toEqual(["only row"]);
  });

  it("handles small chunks (tests chunk boundary handling)", async () => {
    const xlsx = await createTestXlsx([
      {
        name: "Chunked",
        rows: [
          ["Name", "Value"],
          ["test", 123],
          ["data", 456],
        ],
      },
    ]);

    // Use very small chunks (64 bytes) to stress chunk boundary handling
    const stream = toChunkedReadableStream(xlsx, 64);
    const rows = await collectStreamRows(streamXlsxRows(stream));

    expect(rows).toHaveLength(3);
    expect(rows[0].values).toEqual(["Name", "Value"]);
    expect(rows[1].values).toEqual(["test", 123]);
    expect(rows[2].values).toEqual(["data", 456]);
  });

  it("handles mixed types from ReadableStream", async () => {
    const xlsx = await createTestXlsx([
      {
        name: "Mixed",
        rows: [
          ["text", 42, true, null, 3.14],
          [null, null, false, "hello", 0],
        ],
      },
    ]);

    const rows = await collectStreamRows(streamXlsxRows(toReadableStream(xlsx)));

    expect(rows[0].values).toEqual(["text", 42, true, null, 3.14]);
    expect(rows[1].values).toEqual([null, null, false, "hello", 0]);
  });

  it("yields no rows for non-existent sheet from ReadableStream", async () => {
    const xlsx = await createTestXlsx([{ name: "Sheet1", rows: [["data"]] }]);

    const rows = await collectStreamRows(
      streamXlsxRows(toReadableStream(xlsx), { sheet: "NonExistent" }),
    );

    expect(rows).toHaveLength(0);
  });

  it("handles large sheet (100k rows) from ReadableStream", async () => {
    const largeRows: CellValue[][] = [];
    for (let i = 0; i < 100_000; i++) {
      largeRows.push([`Row${i}`, i, i % 2 === 0]);
    }

    const xlsx = await createTestXlsx([{ name: "Large", rows: largeRows }]);

    let count = 0;
    for await (const row of streamXlsxRows(toReadableStream(xlsx))) {
      if (count === 0) {
        expect(row.values[0]).toBe("Row0");
        expect(row.values[1]).toBe(0);
      }
      if (count === 99_999) {
        expect(row.values[0]).toBe("Row99999");
        expect(row.values[1]).toBe(99_999);
      }
      count++;
    }

    expect(count).toBe(100_000);
  }, 30_000);

  it("formula cells return cached result from ReadableStream", async () => {
    const cells = new Map<string, { formula: string; formulaResult: number }>();
    cells.set("0,2", { formula: "A1+B1", formulaResult: 30 });
    cells.set("1,2", { formula: "A2+B2", formulaResult: 70 });

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Formulas",
          rows: [
            [10, 20],
            [30, 40],
          ],
          cells,
        },
      ],
    });

    const rows = await collectStreamRows(streamXlsxRows(toReadableStream(xlsx)));

    expect(rows).toHaveLength(2);
    expect(rows[0].values[0]).toBe(10);
    expect(rows[0].values[1]).toBe(20);
    // Formula result should appear in column C
    expect(rows[0].values[2]).toBe(30);
    expect(rows[1].values[0]).toBe(30);
    expect(rows[1].values[1]).toBe(40);
    expect(rows[1].values[2]).toBe(70);
  });

  it("backward compatibility: Uint8Array path still works identically", async () => {
    const xlsx = await createTestXlsx([
      {
        name: "Compat",
        rows: [
          ["a", 1, true],
          ["b", 2, false],
          ["c", 3, null],
        ],
      },
    ]);

    // Test all three input types produce identical results
    const uint8Rows = await collectStreamRows(streamXlsxRows(xlsx));
    const arrayBufRows = await collectStreamRows(
      streamXlsxRows(
        xlsx.buffer.slice(xlsx.byteOffset, xlsx.byteOffset + xlsx.byteLength) as ArrayBuffer,
      ),
    );
    const streamRows = await collectStreamRows(streamXlsxRows(toReadableStream(xlsx)));

    expect(uint8Rows).toHaveLength(3);
    expect(arrayBufRows).toHaveLength(3);
    expect(streamRows).toHaveLength(3);

    for (let i = 0; i < 3; i++) {
      expect(uint8Rows[i].index).toBe(arrayBufRows[i].index);
      expect(uint8Rows[i].index).toBe(streamRows[i].index);
      expect(uint8Rows[i].values).toEqual(arrayBufRows[i].values);
      expect(uint8Rows[i].values).toEqual(streamRows[i].values);
    }
  });
});

// ═══════════════════════════════════════════════════════════════════════
// XLSX Stream Writer
// ═══════════════════════════════════════════════════════════════════════

describe("XlsxStreamWriter", () => {
  it("writes basic rows and produces valid XLSX", async () => {
    const writer = new XlsxStreamWriter({ name: "Sheet1" });
    writer.addRow(["Hello", 42, true]);
    writer.addRow(["World", 99, false]);

    const xlsx = await writer.finish();

    // Verify it's a valid XLSX by reading it back
    const workbook = await readXlsx(xlsx);

    expect(workbook.sheets).toHaveLength(1);
    expect(workbook.sheets[0].name).toBe("Sheet1");
    expect(workbook.sheets[0].rows).toHaveLength(2);
    expect(workbook.sheets[0].rows[0]).toEqual(["Hello", 42, true]);
    expect(workbook.sheets[0].rows[1]).toEqual(["World", 99, false]);
  });

  it("writes with column headers", async () => {
    const writer = new XlsxStreamWriter({
      name: "Data",
      columns: [
        { header: "Name", key: "name" },
        { header: "Age", key: "age" },
        { header: "Active", key: "active" },
      ],
    });
    writer.addRow(["Alice", 30, true]);
    writer.addRow(["Bob", 25, false]);

    const xlsx = await writer.finish();
    const workbook = await readXlsx(xlsx);

    expect(workbook.sheets[0].rows).toHaveLength(3);
    // Header row auto-added by constructor
    expect(workbook.sheets[0].rows[0]).toEqual(["Name", "Age", "Active"]);
    expect(workbook.sheets[0].rows[1]).toEqual(["Alice", 30, true]);
    expect(workbook.sheets[0].rows[2]).toEqual(["Bob", 25, false]);
  });

  it("writes with freeze pane", async () => {
    const writer = new XlsxStreamWriter({
      name: "Frozen",
      freezePane: { rows: 1 },
    });
    writer.addRow(["Header1", "Header2"]);
    writer.addRow(["Data1", "Data2"]);

    const xlsx = await writer.finish();
    const workbook = await readXlsx(xlsx);

    expect(workbook.sheets[0].rows).toHaveLength(2);
    // Note: freeze pane is in the XML but readXlsx doesn't currently
    // extract it to the Sheet object, so we just verify the data is intact
    expect(workbook.sheets[0].rows[0]).toEqual(["Header1", "Header2"]);
    expect(workbook.sheets[0].rows[1]).toEqual(["Data1", "Data2"]);
  });

  it("read back streamed output with regular reader — data matches", async () => {
    const originalRows: CellValue[][] = [
      ["Name", "Score", "Pass"],
      ["Alice", 95, true],
      ["Bob", 72, true],
      ["Charlie", 45, false],
      [null, 0, null],
    ];

    const writer = new XlsxStreamWriter({ name: "Test" });
    for (const row of originalRows) {
      writer.addRow(row);
    }

    const xlsx = await writer.finish();
    const workbook = await readXlsx(xlsx);

    expect(workbook.sheets[0].rows).toHaveLength(originalRows.length);
    for (let i = 0; i < originalRows.length; i++) {
      const actual = workbook.sheets[0].rows[i];
      const expected = originalRows[i];
      // Compare meaningful values (regular reader pads with null)
      for (let j = 0; j < expected.length; j++) {
        expect(actual[j]).toEqual(expected[j]);
      }
    }
  });

  it("writes 1000 rows — all present", async () => {
    const writer = new XlsxStreamWriter({ name: "Bulk" });
    for (let i = 0; i < 1000; i++) {
      writer.addRow([`Row${i}`, i]);
    }

    const xlsx = await writer.finish();
    const workbook = await readXlsx(xlsx);

    expect(workbook.sheets[0].rows).toHaveLength(1000);
    expect(workbook.sheets[0].rows[0][0]).toBe("Row0");
    expect(workbook.sheets[0].rows[0][1]).toBe(0);
    expect(workbook.sheets[0].rows[999][0]).toBe("Row999");
    expect(workbook.sheets[0].rows[999][1]).toBe(999);
  });

  it("handles mixed types (string, number, boolean, date, null)", async () => {
    const date = new Date(Date.UTC(2024, 6, 4)); // Jul 4, 2024
    const writer = new XlsxStreamWriter({ name: "Types" });
    writer.addRow(["text", 42, true, date, null]);
    writer.addRow([null, 0, false, date, "end"]);

    const xlsx = await writer.finish();
    const workbook = await readXlsx(xlsx);

    const rows = workbook.sheets[0].rows;
    expect(rows).toHaveLength(2);

    // First row
    expect(rows[0][0]).toBe("text");
    expect(rows[0][1]).toBe(42);
    expect(rows[0][2]).toBe(true);
    expect(rows[0][3]).toBeInstanceOf(Date);
    expect((rows[0][3] as Date).getUTCFullYear()).toBe(2024);
    expect((rows[0][3] as Date).getUTCMonth()).toBe(6);
    expect(rows[0][4]).toBe(null);

    // Second row
    expect(rows[1][0]).toBe(null);
    expect(rows[1][1]).toBe(0);
    expect(rows[1][2]).toBe(false);
    expect(rows[1][3]).toBeInstanceOf(Date);
    expect(rows[1][4]).toBe("end");
  });

  it("round-trips through stream reader", async () => {
    const writer = new XlsxStreamWriter({ name: "RoundTrip" });
    writer.addRow(["A", 1, true]);
    writer.addRow(["B", 2, false]);
    writer.addRow(["C", 3, true]);

    const xlsx = await writer.finish();

    // Read back via streaming reader
    const rows = await collectStreamRows(streamXlsxRows(xlsx));

    expect(rows).toHaveLength(3);
    expect(rows[0].values).toEqual(["A", 1, true]);
    expect(rows[1].values).toEqual(["B", 2, false]);
    expect(rows[2].values).toEqual(["C", 3, true]);
  });

  it("produces valid XLSX with no rows", async () => {
    const writer = new XlsxStreamWriter({ name: "Empty" });
    const xlsx = await writer.finish();

    const workbook = await readXlsx(xlsx);
    expect(workbook.sheets).toHaveLength(1);
    expect(workbook.sheets[0].name).toBe("Empty");
    expect(workbook.sheets[0].rows).toHaveLength(0);
  });
});

// ═══════════════════════════════════════════════════════════════════════
// CSV Stream Reader
// ═══════════════════════════════════════════════════════════════════════

describe("streamCsvRows", () => {
  it("streams rows from CSV string", () => {
    const csv = "a,b,c\n1,2,3\n4,5,6";
    const rows = collectSyncRows(streamCsvRows(csv));

    expect(rows).toHaveLength(3);
    expect(rows[0]).toEqual(["a", "b", "c"]);
    expect(rows[1]).toEqual(["1", "2", "3"]);
    expect(rows[2]).toEqual(["4", "5", "6"]);
  });

  it("values match non-streaming parse", () => {
    const csv = 'name,age,city\n"Alice",30,"New York"\nBob,25,London';

    const streamRows = collectSyncRows(streamCsvRows(csv));
    const regularRows = parseCsv(csv);

    expect(streamRows).toEqual(regularRows);
  });

  it("handles quoted fields", () => {
    const csv = '"hello, world",simple,"with ""quotes"""\na,b,c';
    const rows = collectSyncRows(streamCsvRows(csv));

    expect(rows).toHaveLength(2);
    expect(rows[0][0]).toBe("hello, world");
    expect(rows[0][1]).toBe("simple");
    expect(rows[0][2]).toBe('with "quotes"');
  });

  it("type inference works per-row", () => {
    const csv = "true,42,hello,2024-01-15\nfalse,3.14,world,not-a-date";
    const rows = collectSyncRows(streamCsvRows(csv, { typeInference: true }));

    expect(rows).toHaveLength(2);
    expect(rows[0][0]).toBe(true);
    expect(rows[0][1]).toBe(42);
    expect(rows[0][2]).toBe("hello");
    expect(rows[0][3]).toBeInstanceOf(Date);

    expect(rows[1][0]).toBe(false);
    expect(rows[1][1]).toBeCloseTo(3.14);
    expect(rows[1][2]).toBe("world");
    expect(rows[1][3]).toBe("not-a-date");
  });

  it("header row handling", () => {
    const csv = "name,age\nAlice,30\nBob,25";
    const rows = collectSyncRows(streamCsvRows(csv, { header: true }));

    // Header row should be consumed, not yielded
    expect(rows).toHaveLength(2);
    expect(rows[0]).toEqual(["Alice", "30"]);
    expect(rows[1]).toEqual(["Bob", "25"]);
  });

  it("empty input yields no rows", () => {
    const rows = collectSyncRows(streamCsvRows(""));
    expect(rows).toHaveLength(0);
  });

  it("handles CRLF line endings", () => {
    const csv = "a,b\r\n1,2\r\n3,4";
    const rows = collectSyncRows(streamCsvRows(csv));

    expect(rows).toHaveLength(3);
    expect(rows[0]).toEqual(["a", "b"]);
    expect(rows[1]).toEqual(["1", "2"]);
    expect(rows[2]).toEqual(["3", "4"]);
  });

  it("handles trailing newline without extra empty row", () => {
    const csv = "a,b\n1,2\n";
    const rows = collectSyncRows(streamCsvRows(csv));

    expect(rows).toHaveLength(2);
    expect(rows[0]).toEqual(["a", "b"]);
    expect(rows[1]).toEqual(["1", "2"]);
  });

  it("skips BOM by default", () => {
    const csv = "\uFEFFa,b\n1,2";
    const rows = collectSyncRows(streamCsvRows(csv));

    expect(rows).toHaveLength(2);
    expect(rows[0][0]).toBe("a");
  });

  it("skips comment rows", () => {
    const csv = "# comment\na,b\n# another\n1,2";
    const rows = collectSyncRows(streamCsvRows(csv, { comment: "#" }));

    expect(rows).toHaveLength(2);
    expect(rows[0]).toEqual(["a", "b"]);
    expect(rows[1]).toEqual(["1", "2"]);
  });

  it("skips empty rows when configured", () => {
    const csv = "a,b\n\n1,2\n\n3,4";
    const rows = collectSyncRows(streamCsvRows(csv, { skipEmptyRows: true }));

    expect(rows).toHaveLength(3);
    expect(rows[0]).toEqual(["a", "b"]);
    expect(rows[1]).toEqual(["1", "2"]);
    expect(rows[2]).toEqual(["3", "4"]);
  });

  it("handles custom delimiter", () => {
    const csv = "a;b;c\n1;2;3";
    const rows = collectSyncRows(streamCsvRows(csv, { delimiter: ";" }));

    expect(rows).toHaveLength(2);
    expect(rows[0]).toEqual(["a", "b", "c"]);
    expect(rows[1]).toEqual(["1", "2", "3"]);
  });

  it("handles quoted fields with newlines inside", () => {
    const csv = '"line1\nline2",b\nc,d';
    const rows = collectSyncRows(streamCsvRows(csv));

    expect(rows).toHaveLength(2);
    expect(rows[0][0]).toBe("line1\nline2");
    expect(rows[0][1]).toBe("b");
    expect(rows[1][0]).toBe("c");
  });
});

// ═══════════════════════════════════════════════════════════════════════
// CSV Stream Writer
// ═══════════════════════════════════════════════════════════════════════

describe("CsvStreamWriter", () => {
  it("writes rows incrementally", () => {
    const writer = new CsvStreamWriter();
    writer.addRow(["a", "b", "c"]);
    writer.addRow(["1", "2", "3"]);

    const result = writer.finish();
    expect(result).toBe("a,b,c\r\n1,2,3");
  });

  it("output matches non-streaming writeCsv", () => {
    const rows: CellValue[][] = [
      ["Name", "Age", "City"],
      ["Alice", 30, "New York"],
      ["Bob", 25, "London"],
    ];

    // Non-streaming
    const expected = writeCsv(rows);

    // Streaming
    const writer = new CsvStreamWriter();
    for (const row of rows) {
      writer.addRow(row);
    }
    const result = writer.finish();

    expect(result).toBe(expected);
  });

  it("writes with headers", () => {
    const writer = new CsvStreamWriter({
      headers: ["Name", "Age"],
    });
    writer.addRow(["Alice", 30]);
    writer.addRow(["Bob", 25]);

    const result = writer.finish();
    expect(result).toBe("Name,Age\r\nAlice,30\r\nBob,25");
  });

  it("writes with BOM", () => {
    const writer = new CsvStreamWriter({ bom: true });
    writer.addRow(["a", "b"]);

    const result = writer.finish();
    expect(result).toBe("\uFEFFa,b");
  });

  it("handles mixed types", () => {
    const writer = new CsvStreamWriter();
    writer.addRow(["text", 42, true, null, false]);

    const result = writer.finish();
    expect(result).toBe("text,42,true,,false");
  });

  it("quotes fields containing delimiter", () => {
    const writer = new CsvStreamWriter();
    writer.addRow(["hello, world", "simple"]);

    const result = writer.finish();
    expect(result).toBe('"hello, world",simple');
  });

  it("quotes fields containing newlines", () => {
    const writer = new CsvStreamWriter();
    writer.addRow(["line1\nline2", "ok"]);

    const result = writer.finish();
    expect(result).toBe('"line1\nline2",ok');
  });

  it("escapes quote characters by doubling", () => {
    const writer = new CsvStreamWriter();
    writer.addRow(['say "hello"', "ok"]);

    const result = writer.finish();
    expect(result).toBe('"say ""hello""",ok');
  });

  it("uses custom delimiter", () => {
    const writer = new CsvStreamWriter({ delimiter: ";" });
    writer.addRow(["a", "b", "c"]);

    const result = writer.finish();
    expect(result).toBe("a;b;c");
  });

  it("uses custom line separator", () => {
    const writer = new CsvStreamWriter({ lineSeparator: "\r\n" });
    writer.addRow(["a", "b"]);
    writer.addRow(["1", "2"]);

    const result = writer.finish();
    expect(result).toBe("a,b\r\n1,2");
  });

  it("handles date values", () => {
    const date = new Date("2024-07-04T00:00:00.000Z");
    const writer = new CsvStreamWriter();
    writer.addRow([date]);

    const result = writer.finish();
    expect(result).toBe("2024-07-04T00:00:00.000Z");
  });

  it("handles empty output", () => {
    const writer = new CsvStreamWriter();
    const result = writer.finish();
    expect(result).toBe("");
  });

  it("quote style all wraps every field", () => {
    const writer = new CsvStreamWriter({ quoteStyle: "all" });
    writer.addRow(["a", "b"]);

    const result = writer.finish();
    expect(result).toBe('"a","b"');
  });

  it("BOM with headers", () => {
    const writer = new CsvStreamWriter({
      bom: true,
      headers: ["X", "Y"],
    });
    writer.addRow([1, 2]);

    const result = writer.finish();
    expect(result).toBe("\uFEFFX,Y\r\n1,2");
  });
});

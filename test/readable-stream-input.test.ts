// ── ReadableStream<Uint8Array> input across reader entry points ────────
//
// `ReadInput` is documented as `Uint8Array | ArrayBuffer |
// ReadableStream<Uint8Array>`, but until now only `streamXlsxRows` honored
// the third arm. These tests lock in stream support across `readXlsx`,
// `readOds`, the unified `read()` dispatcher, and the object shorthands.
// ──────────────────────────────────────────────────────────────────────

import { describe, expect, it } from "vitest";
import { readXlsx } from "../src/xlsx/reader";
import { writeXlsx } from "../src/xlsx/writer";
import { readOds } from "../src/ods/reader";
import { writeOds } from "../src/ods/writer";
import { read, readObjects } from "../src/defter";
import { readXlsxObjects } from "../src/xlsx/objects";
import { readOdsObjects } from "../src/ods/objects";
import { ParseError } from "../src/errors";

// ── Helpers ─────────────────────────────────────────────────────────

/** Wrap a Uint8Array in a ReadableStream that delivers it as a single chunk. */
function toReadableStream(data: Uint8Array): ReadableStream<Uint8Array> {
  return new ReadableStream({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  });
}

/**
 * Wrap a Uint8Array in a ReadableStream that delivers it across many
 * fixed-size chunks. Exercises the multi-chunk path of
 * `bufferReadableStream`.
 */
function toChunkedReadableStream(data: Uint8Array, chunkSize: number): ReadableStream<Uint8Array> {
  let offset = 0;
  return new ReadableStream({
    pull(controller) {
      if (offset >= data.length) {
        controller.close();
        return;
      }
      const end = Math.min(offset + chunkSize, data.length);
      controller.enqueue(data.slice(offset, end));
      offset = end;
    },
  });
}

// ── Fixtures ────────────────────────────────────────────────────────

const SAMPLE_SHEETS = [
  {
    name: "Data",
    rows: [
      ["id", "name", "amount"],
      [1, "alpha", 10.5],
      [2, "beta", 20.25],
    ],
  },
];

// ── readXlsx ────────────────────────────────────────────────────────

describe("readXlsx — ReadableStream input", () => {
  it("reads a workbook from a single-chunk ReadableStream", async () => {
    const xlsx = await writeXlsx({ sheets: SAMPLE_SHEETS });

    const wb = await readXlsx(toReadableStream(xlsx));

    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0]!.name).toBe("Data");
    expect(wb.sheets[0]!.rows).toEqual([
      ["id", "name", "amount"],
      [1, "alpha", 10.5],
      [2, "beta", 20.25],
    ]);
  });

  it("matches Uint8Array output when input arrives as many small chunks", async () => {
    const xlsx = await writeXlsx({ sheets: SAMPLE_SHEETS });

    const fromBytes = await readXlsx(xlsx);
    const fromStream = await readXlsx(toChunkedReadableStream(xlsx, 64));

    expect(fromStream.sheets[0]!.rows).toEqual(fromBytes.sheets[0]!.rows);
    expect(fromStream.sheets[0]!.name).toEqual(fromBytes.sheets[0]!.name);
  });

  it("accepts ArrayBuffer alongside Uint8Array (regression guard)", async () => {
    const xlsx = await writeXlsx({ sheets: SAMPLE_SHEETS });
    const ab = xlsx.buffer.slice(xlsx.byteOffset, xlsx.byteOffset + xlsx.byteLength) as ArrayBuffer;

    const wb = await readXlsx(ab);
    expect(wb.sheets[0]!.rows[1]).toEqual([1, "alpha", 10.5]);
  });

  it("throws ParseError for unsupported input shapes", async () => {
    await expect(readXlsx("not-a-buffer" as unknown as Uint8Array)).rejects.toBeInstanceOf(
      ParseError,
    );
  });
});

// ── readOds ─────────────────────────────────────────────────────────

describe("readOds — ReadableStream input", () => {
  it("reads an ODS workbook from a ReadableStream", async () => {
    const ods = await writeOds({ sheets: SAMPLE_SHEETS });

    const wb = await readOds(toReadableStream(ods));

    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0]!.name).toBe("Data");
    // ODS preserves cell content but not always exact numeric typing —
    // assert string-equivalent values rather than strict equality.
    const row = wb.sheets[0]!.rows[1]!;
    expect(String(row[1])).toBe("alpha");
  });

  it("matches Uint8Array output when input arrives chunked", async () => {
    const ods = await writeOds({ sheets: SAMPLE_SHEETS });

    const fromBytes = await readOds(ods);
    const fromStream = await readOds(toChunkedReadableStream(ods, 128));

    expect(fromStream.sheets[0]!.rows).toEqual(fromBytes.sheets[0]!.rows);
  });
});

// ── unified read() ──────────────────────────────────────────────────

describe("read() — ReadableStream input", () => {
  it("auto-detects XLSX from a stream", async () => {
    const xlsx = await writeXlsx({ sheets: SAMPLE_SHEETS });

    const wb = await read(toReadableStream(xlsx));

    expect(wb.sheets[0]!.name).toBe("Data");
    expect(wb.sheets[0]!.rows[2]).toEqual([2, "beta", 20.25]);
  });

  it("auto-detects ODS from a stream", async () => {
    const ods = await writeOds({ sheets: SAMPLE_SHEETS });

    const wb = await read(toReadableStream(ods));

    expect(wb.sheets[0]!.name).toBe("Data");
    expect(wb.sheets[0]!.rows[0]).toEqual(["id", "name", "amount"]);
  });

  it("readObjects accepts a stream", async () => {
    const xlsx = await writeXlsx({ sheets: SAMPLE_SHEETS });

    const objects = await readObjects(toReadableStream(xlsx));

    expect(objects).toEqual([
      { id: 1, name: "alpha", amount: 10.5 },
      { id: 2, name: "beta", amount: 20.25 },
    ]);
  });
});

// ── *Objects shorthands ─────────────────────────────────────────────

describe("readXlsxObjects / readOdsObjects — ReadableStream input", () => {
  it("readXlsxObjects accepts a stream", async () => {
    const xlsx = await writeXlsx({ sheets: SAMPLE_SHEETS });

    const { headers, data } = await readXlsxObjects(toChunkedReadableStream(xlsx, 256));

    expect(headers).toEqual(["id", "name", "amount"]);
    expect(data).toEqual([
      { id: 1, name: "alpha", amount: 10.5 },
      { id: 2, name: "beta", amount: 20.25 },
    ]);
  });

  it("readOdsObjects accepts a stream", async () => {
    const ods = await writeOds({ sheets: SAMPLE_SHEETS });

    const { headers, data } = await readOdsObjects(toReadableStream(ods));

    expect(headers).toEqual(["id", "name", "amount"]);
    expect(data).toHaveLength(2);
    expect(data[0]!.name).toBe("alpha");
    expect(data[1]!.name).toBe("beta");
  });
});

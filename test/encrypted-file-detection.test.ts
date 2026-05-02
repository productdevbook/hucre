// ── Encrypted-workbook detection across reader entry points ────────────
//
// Office encrypts password-protected XLSX / XLSM / ODS files inside an
// OLE2 / Compound File Binary container (MS-OFFCRYPTO). Until full
// decryption lands (#156), every reader should recognize the OLE2
// header up front and surface a typed `EncryptedFileError` rather than
// the confusing `"not a valid ZIP archive"` ParseError that the ZIP
// reader throws when it tries to open one.
//
// Each test here builds a synthetic OLE2 buffer (just the 8-byte magic
// padded with zeros — enough to hit the byte-sniff path) and exercises
// every public reader entry point.
// ──────────────────────────────────────────────────────────────────────

import { describe, expect, it } from "vitest";
import { readXlsx } from "../src/xlsx/reader";
import { streamXlsxRows } from "../src/xlsx/stream-reader";
import { readOds } from "../src/ods/reader";
import { read, readObjects } from "../src/defter";
import { readXlsxObjects } from "../src/xlsx/objects";
import { readOdsObjects } from "../src/ods/objects";
import { isOle2Container } from "../src/_input";
import { EncryptedFileError, DefterError, ZipError, ParseError } from "../src/errors";

// ── Helpers ─────────────────────────────────────────────────────────

/** OLE2 / CFB compound-document magic header. */
const OLE2_MAGIC = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

/**
 * Build a minimal OLE2 container — the 8-byte magic header plus enough
 * padding zero bytes to clear the ZIP "data too small" guard. The byte
 * sniff only inspects the first 8 bytes, so the rest of the buffer is
 * irrelevant to the detection path under test.
 */
function makeOle2Container(extraBytes: number = 256): Uint8Array {
  const buf = new Uint8Array(OLE2_MAGIC.length + extraBytes);
  buf.set(OLE2_MAGIC, 0);
  return buf;
}

function toReadableStream(data: Uint8Array): ReadableStream<Uint8Array> {
  return new ReadableStream({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  });
}

// ── isOle2Container ─────────────────────────────────────────────────

describe("isOle2Container — byte-sniff helper", () => {
  it("returns true for the 8-byte CFB magic header", () => {
    expect(isOle2Container(OLE2_MAGIC)).toBe(true);
  });

  it("returns true for the magic followed by trailing bytes", () => {
    expect(isOle2Container(makeOle2Container(64))).toBe(true);
  });

  it("returns false for a plain ZIP archive (PK\\x03\\x04)", () => {
    const zipMagic = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0]);
    expect(isOle2Container(zipMagic)).toBe(false);
  });

  it("returns false for buffers shorter than the magic", () => {
    expect(isOle2Container(new Uint8Array(0))).toBe(false);
    expect(isOle2Container(new Uint8Array([0xd0, 0xcf, 0x11]))).toBe(false);
  });

  it("returns false when only the first byte matches", () => {
    const almost = new Uint8Array([0xd0, 0, 0, 0, 0, 0, 0, 0]);
    expect(isOle2Container(almost)).toBe(false);
  });
});

// ── readXlsx ────────────────────────────────────────────────────────

describe("readXlsx — encrypted XLSX detection", () => {
  it("throws EncryptedFileError for the OLE2 magic instead of a ZIP ParseError", async () => {
    const data = makeOle2Container();

    await expect(readXlsx(data)).rejects.toBeInstanceOf(EncryptedFileError);
  });

  it('attaches format="xlsx" to the error so callers can branch', async () => {
    const data = makeOle2Container();

    try {
      await readXlsx(data);
      throw new Error("readXlsx should have thrown");
    } catch (err) {
      expect(err).toBeInstanceOf(EncryptedFileError);
      const enc = err as EncryptedFileError;
      expect(enc.format).toBe("xlsx");
      expect(enc.message).toContain("password-protected");
      expect(enc.message).toContain("XLSX");
      // Must extend the project's base error so existing catch-all
      // handlers `instanceof DefterError` keep working.
      expect(err).toBeInstanceOf(DefterError);
      // EncryptedFileError is intentionally distinct from ZipError /
      // ParseError so callers can branch on it.
      expect(err).not.toBeInstanceOf(ZipError);
      expect(err).not.toBeInstanceOf(ParseError);
    }
  });

  it("detects encryption when the input arrives as a ReadableStream", async () => {
    const data = makeOle2Container();

    await expect(readXlsx(toReadableStream(data))).rejects.toBeInstanceOf(EncryptedFileError);
  });

  it("accepts ArrayBuffer encoding of the OLE2 container", async () => {
    const data = makeOle2Container();
    const ab = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;

    await expect(readXlsx(ab)).rejects.toBeInstanceOf(EncryptedFileError);
  });
});

// ── streamXlsxRows ──────────────────────────────────────────────────

describe("streamXlsxRows — encrypted XLSX detection", () => {
  it("rejects the generator with EncryptedFileError before ZIP parsing runs", async () => {
    const data = makeOle2Container();

    const gen = streamXlsxRows(data);
    await expect(gen.next()).rejects.toBeInstanceOf(EncryptedFileError);
  });

  it("detects encryption from a chunked ReadableStream input", async () => {
    const data = makeOle2Container();

    const gen = streamXlsxRows(toReadableStream(data));
    await expect(gen.next()).rejects.toBeInstanceOf(EncryptedFileError);
  });
});

// ── readOds ─────────────────────────────────────────────────────────

describe("readOds — encrypted ODS detection", () => {
  it('throws EncryptedFileError with format="ods" for the OLE2 magic', async () => {
    const data = makeOle2Container();

    try {
      await readOds(data);
      throw new Error("readOds should have thrown");
    } catch (err) {
      expect(err).toBeInstanceOf(EncryptedFileError);
      const enc = err as EncryptedFileError;
      expect(enc.format).toBe("ods");
      expect(enc.message).toContain("ODS");
    }
  });

  it("detects encryption from a ReadableStream input", async () => {
    const data = makeOle2Container();

    await expect(readOds(toReadableStream(data))).rejects.toBeInstanceOf(EncryptedFileError);
  });
});

// ── unified read() / readObjects() ──────────────────────────────────

describe("read() — encrypted workbook detection (auto-detect path)", () => {
  it("throws EncryptedFileError before format auto-detection runs", async () => {
    const data = makeOle2Container();

    try {
      await read(data);
      throw new Error("read() should have thrown");
    } catch (err) {
      expect(err).toBeInstanceOf(EncryptedFileError);
      // The unified entry point can't tell whether the encrypted
      // package is XLSX or ODS without decrypting, so `format` is
      // intentionally undefined.
      expect((err as EncryptedFileError).format).toBeUndefined();
    }
  });

  it("readObjects() inherits the same detection path", async () => {
    const data = makeOle2Container();

    await expect(readObjects(data)).rejects.toBeInstanceOf(EncryptedFileError);
  });

  it("detects encryption from a ReadableStream via read()", async () => {
    const data = makeOle2Container();

    await expect(read(toReadableStream(data))).rejects.toBeInstanceOf(EncryptedFileError);
  });
});

// ── Object shorthands ──────────────────────────────────────────────

describe("readXlsxObjects / readOdsObjects — encrypted detection", () => {
  it("readXlsxObjects throws EncryptedFileError", async () => {
    const data = makeOle2Container();

    await expect(readXlsxObjects(data)).rejects.toBeInstanceOf(EncryptedFileError);
  });

  it("readOdsObjects throws EncryptedFileError", async () => {
    const data = makeOle2Container();

    await expect(readOdsObjects(data)).rejects.toBeInstanceOf(EncryptedFileError);
  });
});

// ── EncryptedFileError API ─────────────────────────────────────────

describe("EncryptedFileError — constructor surface", () => {
  it("supports the legacy zero-arg constructor with no format hint", () => {
    const err = new EncryptedFileError();
    expect(err.name).toBe("EncryptedFileError");
    expect(err.format).toBeUndefined();
    expect(err.message).toBe("File is password-protected. Provide a password in options.");
  });

  it("includes the format hint and a uppercase format token in the default message", () => {
    const err = new EncryptedFileError("xlsx");
    expect(err.format).toBe("xlsx");
    expect(err.message).toContain("XLSX");
    expect(err.message).toContain("password-protected");
  });

  it("accepts a fully custom message override while still tracking the format", () => {
    const err = new EncryptedFileError("ods", "custom note");
    expect(err.format).toBe("ods");
    expect(err.message).toBe("custom note");
  });
});

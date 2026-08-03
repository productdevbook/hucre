import { describe, expect, it } from "vitest"

import {
  DecryptionError,
  DefterError,
  EncryptedFileError,
  HucreError,
  InvalidArgumentError,
  ParseError,
  UnsupportedFormatError,
  ValidationError,
  XmlError,
  ZipError,
} from "../src/errors"
import { DefterError as rootDefterError, HucreError as rootHucreError } from "../src/index"

// ── The v1 rename ────────────────────────────────────────────────────
// `DefterError` was the frozen root of the hierarchy in a package called
// `hucre`. Renaming it after v1 would break every `instanceof` catch-all
// in the wild, so the rename lands now and the old name stays as an
// alias to the *same class object* — not a subclass, not a copy.

describe("HucreError / DefterError", () => {
  it("keeps DefterError as an alias of the same class object", () => {
    expect(DefterError).toBe(HucreError)
    expect(rootDefterError).toBe(rootHucreError)
    expect(rootHucreError).toBe(HucreError)
  })

  it("keeps `instanceof DefterError` working for every subclass", () => {
    const subclasses = [
      new ParseError("x"),
      new ZipError("x"),
      new XmlError("x"),
      new ValidationError("x", []),
      new InvalidArgumentError("x"),
      new UnsupportedFormatError("x"),
      new EncryptedFileError("xlsx"),
      new DecryptionError(),
    ]
    for (const err of subclasses) {
      expect(err).toBeInstanceOf(DefterError)
      expect(err).toBeInstanceOf(HucreError)
      expect(err).toBeInstanceOf(Error)
    }
  })

  it("reports the new name on the base class", () => {
    expect(new HucreError("boom").name).toBe("HucreError")
    expect(new DefterError("boom").name).toBe("HucreError")
    expect(new ParseError("boom").name).toBe("ParseError")
  })

  it("still carries message and cause through the base constructor", () => {
    const cause = new Error("root cause")
    const err = new HucreError("wrapped", { cause })
    expect(err.message).toBe("wrapped")
    expect(err.cause).toBe(cause)
  })
})

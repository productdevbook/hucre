import { describe, expect, it } from "vitest"

import {
  DecryptionError,
  EncryptedFileError,
  HucreError,
  InvalidArgumentError,
  ParseError,
  UnsupportedFormatError,
  ValidationError,
  XmlError,
  ZipError,
} from "../src/errors"
import { HucreError as rootHucreError } from "../src/index"

describe("HucreError", () => {
  it("is the same class object from the root and from ./errors", () => {
    expect(rootHucreError).toBe(HucreError)
  })

  it("is the base of every subclass", () => {
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
      expect(err).toBeInstanceOf(HucreError)
      expect(err).toBeInstanceOf(Error)
    }
  })

  it("reports the new name on the base class", () => {
    expect(new HucreError("boom").name).toBe("HucreError")
    expect(new ParseError("boom").name).toBe("ParseError")
  })

  it("still carries message and cause through the base constructor", () => {
    const cause = new Error("root cause")
    const err = new HucreError("wrapped", { cause })
    expect(err.message).toBe("wrapped")
    expect(err.cause).toBe(cause)
  })
})

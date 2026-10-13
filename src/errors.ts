/**
 * Root of the error hierarchy. Every error this library throws is an
 * instance of it, so `catch (e) { if (e instanceof HucreError) … }` is
 * the catch-all.
 */
export class HucreError extends Error {
  override name = "HucreError"

  constructor(message: string, options?: ErrorOptions) {
    super(message, options)
  }
}

export class ParseError extends HucreError {
  override name = "ParseError"

  constructor(
    message: string,
    public readonly details?: {
      file?: string
      line?: number
      column?: number
    },
    options?: ErrorOptions,
  ) {
    super(message, options)
  }
}

export class ZipError extends HucreError {
  override name = "ZipError"
}

export class XmlError extends HucreError {
  override name = "XmlError"
}

export class ValidationError extends HucreError {
  override name = "ValidationError"

  constructor(
    message: string,
    public readonly errors: Array<{
      row: number
      column: string | number
      message: string
      value: unknown
      field: string
    }>,
  ) {
    super(message)
  }
}

/**
 * A caller passed something the format cannot represent — an illegal
 * sheet name, an out-of-range coordinate, a value past a hard limit.
 *
 * Distinct from {@link ValidationError}, whose constructor takes a list
 * of row/column schema failures and models *imported data* being wrong.
 * This one models *arguments* being wrong, and is thrown before any
 * output is produced rather than collected alongside it.
 */
export class InvalidArgumentError extends HucreError {
  override name = "InvalidArgumentError"
}

export class UnsupportedFormatError extends HucreError {
  override name = "UnsupportedFormatError"

  constructor(format: string) {
    super(`Unsupported format: ${format}`)
  }
}

export class EncryptedFileError extends HucreError {
  override name = "EncryptedFileError"

  /**
   * Format hint for the encrypted container, when known. `"xlsx"` /
   * `"ods"` mean the caller's reader detected the OLE2 / CFB envelope
   * that Office uses for password-protected workbooks. Older callers
   * that constructed `new EncryptedFileError()` without a hint still
   * see `undefined` here.
   */
  readonly format?: "xlsx" | "ods"

  constructor(format?: "xlsx" | "ods", message?: string) {
    super(
      message ??
        (format
          ? `File is password-protected (${format.toUpperCase()} encrypted with the OLE2/CFB container). Provide a password in options to decrypt it.`
          : "File is password-protected. Provide a password in options."),
    )
    this.format = format
  }
}

/**
 * Thrown when a password WAS supplied for an encrypted workbook but
 * decryption failed — almost always a wrong password, occasionally a
 * corrupt or unsupported encryption blob. Distinct from
 * {@link EncryptedFileError} ("encrypted, and no password was given").
 */
export class DecryptionError extends HucreError {
  override name = "DecryptionError"

  constructor(
    message = "Failed to decrypt: incorrect password or corrupt encryption data.",
    options?: ErrorOptions,
  ) {
    super(message, options)
  }
}

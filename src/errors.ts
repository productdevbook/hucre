export class DefterError extends Error {
  override name = "DefterError";

  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
  }
}

export class ParseError extends DefterError {
  override name = "ParseError";

  constructor(
    message: string,
    public readonly details?: {
      file?: string;
      line?: number;
      column?: number;
    },
    options?: ErrorOptions,
  ) {
    super(message, options);
  }
}

export class ZipError extends DefterError {
  override name = "ZipError";
}

export class XmlError extends DefterError {
  override name = "XmlError";
}

export class ValidationError extends DefterError {
  override name = "ValidationError";

  constructor(
    message: string,
    public readonly errors: Array<{
      row: number;
      column: string | number;
      message: string;
      value: unknown;
      field: string;
    }>,
  ) {
    super(message);
  }
}

export class UnsupportedFormatError extends DefterError {
  override name = "UnsupportedFormatError";

  constructor(format: string) {
    super(`Unsupported format: ${format}`);
  }
}

export class EncryptedFileError extends DefterError {
  override name = "EncryptedFileError";

  /**
   * Format hint for the encrypted container, when known. `"xlsx"` /
   * `"ods"` mean the caller's reader detected the OLE2 / CFB envelope
   * that Office uses for password-protected workbooks. Older callers
   * that constructed `new EncryptedFileError()` without a hint still
   * see `undefined` here.
   */
  readonly format?: "xlsx" | "ods";

  constructor(format?: "xlsx" | "ods", message?: string) {
    super(
      message ??
        (format
          ? `File is password-protected (${format.toUpperCase()} encrypted with the OLE2/CFB container). Reading password-protected files is not yet supported.`
          : "File is password-protected. Provide a password in options."),
    );
    this.format = format;
  }
}

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

  constructor() {
    super("File is password-protected. Provide a password in options.");
  }
}

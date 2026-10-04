// ── BIFF5 code pages ─────────────────────────────────────────────────
// A BIFF5/7 workbook stores every string as bytes in one Windows code
// page, named once by the globals' CODEPAGE record. Decoding is the Web
// `TextDecoder`, so the page has to be one WHATWG names.

import { ParseError } from "../errors"

const CODEPAGE_LABELS: Record<number, string> = {
  866: "ibm866",
  874: "windows-874",
  932: "shift_jis",
  936: "gbk",
  949: "euc-kr",
  950: "big5",
  1200: "utf-16le",
  1250: "windows-1250",
  1251: "windows-1251",
  1252: "windows-1252",
  1253: "windows-1253",
  1254: "windows-1254",
  1255: "windows-1255",
  1256: "windows-1256",
  1257: "windows-1257",
  1258: "windows-1258",
  10000: "macintosh",
  10007: "x-mac-cyrillic",
  20866: "koi8-r",
  21866: "koi8-u",
  // 0x8000 / 0x8001: the Macintosh and Windows ANSI markers Excel for Mac
  // wrote instead of a real page number.
  32768: "macintosh",
  32769: "windows-1252",
  65001: "utf-8",
}

/** The WHATWG encoding label for a Windows code page, or undefined. */
export function codepageLabel(codepage: number): string | undefined {
  const label = CODEPAGE_LABELS[codepage]
  if (label) return label
  // ISO-8859-n is the contiguous block 28591-28606.
  if (codepage > 28590 && codepage <= 28606) return `iso-8859-${codepage - 28590}`
  return undefined
}

/**
 * A byte → string decoder for one code page. A page this runtime cannot
 * decode is an error naming it, never silently Latin-1: the text would
 * come back as the wrong letters with no sign anything went wrong.
 */
export function codepageDecoder(codepage: number): (bytes: Uint8Array) => string {
  const label = codepageLabel(codepage)
  if (!label) throw new ParseError(`Unsupported XLS codepage ${codepage}`)
  let decoder: TextDecoder
  try {
    decoder = new TextDecoder(label)
  } catch (err) {
    throw new ParseError(
      `XLS codepage ${codepage} (${label}) is not decodable in this runtime`,
      undefined,
      { cause: err },
    )
  }
  return (bytes) => decoder.decode(bytes)
}

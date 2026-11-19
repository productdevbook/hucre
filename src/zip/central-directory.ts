// ── ZIP central directory parsing ────────────────────────────────────
// The end-of-central-directory record, its ZIP64 counterpart, and the
// central directory itself, parsed out of whatever bytes the caller has.
//
// `ZipReader` holds the whole archive and passes it in whole. The
// random-access reader (`random-access.ts`) holds only the tail and the
// directory, read from a source it can seek in. Both go through these
// functions, so there is one reading of the format rather than two that
// drift — the bounds checks below are the ones `ZipReader` has always
// applied, with the offsets made relative to the bytes actually held.

import { ZipError } from "../errors"

export const SIG_LOCAL_FILE = 0x04034b50
export const SIG_CENTRAL_DIR = 0x02014b50
export const SIG_END_OF_CENTRAL_DIR = 0x06054b50
export const SIG_DATA_DESCRIPTOR = 0x08074b50
export const SIG_ZIP64_EOCD = 0x06064b50
export const SIG_ZIP64_EOCD_LOCATOR = 0x07064b50

/** 0xFFFFFFFF marker that signals a 32-bit ZIP field overflowed into ZIP64. */
export const ZIP64_SENTINEL = 0xffffffff
/** 0xFFFF marker for the 16-bit entry-count fields. */
export const ZIP64_SENTINEL_16 = 0xffff
/** Header id of the ZIP64 extended information extra field. */
export const ZIP64_EXTRA_ID = 0x0001

/** Smallest classic EOCD record (no comment). */
export const EOCD_MIN_BYTES = 22
/** Furthest back the EOCD can start: the record plus a maximal comment. */
export const EOCD_SEARCH_BYTES: number = 22 + 65535
/** The ZIP64 EOCD locator sits immediately before the classic EOCD. */
export const ZIP64_LOCATOR_BYTES = 20
/** Fixed part of a ZIP64 EOCD record. */
export const ZIP64_EOCD_BYTES = 56
/** Fixed part of one central directory record. */
export const CENTRAL_DIR_ENTRY_MIN_BYTES = 46
/** Fixed part of a local file header. */
export const LOCAL_HEADER_BYTES = 30

/** STORE (0) or DEFLATE (8): the compression methods hucre reads. */
export function isSupportedMethod(method: number): boolean {
  return method === 0 || method === 8
}

export interface CentralDirEntry {
  fileName: string
  compressionMethod: number
  compressedSize: number
  uncompressedSize: number
  crc32: number
  localHeaderOffset: number
  /** Bit 3 of general purpose flag — data descriptor present */
  hasDataDescriptor: boolean
}

/** Where the central directory is, as the end records state it. */
export interface DirectoryBounds {
  entryCount: number
  centralDirSize: number
  centralDirOffset: number
}

/** The classic EOCD record, with the ZIP64 question still open. */
export interface EocdRecord extends DirectoryBounds {
  /** Absolute offset of the EOCD signature. */
  eocdOffset: number
  /** A field holds a sentinel, so the real bounds are in the ZIP64 record. */
  needsZip64: boolean
}

/**
 * Find and decode the classic EOCD record in `bytes`, which hold the
 * archive from absolute offset `base` to its end.
 *
 * The record is at least 22 bytes and may be followed by a comment of up
 * to 65,535, so the search runs backwards over that window.
 */
export function readEocd(bytes: Uint8Array, base: number): EocdRecord {
  const view = viewOf(bytes)
  const minOffset = Math.max(0, bytes.length - EOCD_SEARCH_BYTES)

  let at = -1
  for (let i = bytes.length - EOCD_MIN_BYTES; i >= minOffset; i--) {
    if (view.getUint32(i, true) === SIG_END_OF_CENTRAL_DIR) {
      at = i
      break
    }
  }
  if (at === -1) {
    throw new ZipError("End of Central Directory not found — not a valid ZIP file")
  }

  const centralDirSize = view.getUint32(at + 12, true)
  const centralDirOffset = view.getUint32(at + 16, true)
  const entryCount = view.getUint16(at + 10, true)

  // ZIP64: when the entry count or central-directory size/offset overflows
  // the 16-/32-bit EOCD fields they hold a 0xFFFF / 0xFFFFFFFF sentinel and
  // the real values live in a ZIP64 EOCD record. Producers also emit these
  // records defensively on small archives, so a sentinel is not by itself
  // evidence that the archive is huge.
  const needsZip64 =
    entryCount === ZIP64_SENTINEL_16 ||
    centralDirSize === ZIP64_SENTINEL ||
    centralDirOffset === ZIP64_SENTINEL

  return { eocdOffset: base + at, entryCount, centralDirSize, centralDirOffset, needsZip64 }
}

/**
 * Decode the ZIP64 EOCD locator (20 bytes, immediately before the classic
 * EOCD) and return the absolute offset of the ZIP64 EOCD record.
 */
export function readZip64Locator(locator: Uint8Array): number {
  if (
    locator.length < ZIP64_LOCATOR_BYTES ||
    viewOf(locator).getUint32(0, true) !== SIG_ZIP64_EOCD_LOCATOR
  ) {
    throw new ZipError("ZIP64 End of Central Directory locator not found")
  }
  return readUint64(locator, 8, "ZIP64 EOCD offset")
}

/** Decode a ZIP64 EOCD record, which repeats the bounds at 64-bit width. */
export function readZip64Eocd(record: Uint8Array): DirectoryBounds {
  if (viewOf(record).getUint32(0, true) !== SIG_ZIP64_EOCD) {
    throw new ZipError("Invalid ZIP64 End of Central Directory signature")
  }
  return {
    entryCount: readUint64(record, 32, "ZIP64 entry count"),
    centralDirSize: readUint64(record, 40, "ZIP64 central directory size"),
    centralDirOffset: readUint64(record, 48, "ZIP64 central directory offset"),
  }
}

/**
 * Parse `expectedCount` central directory records from `bytes`, starting
 * at `start`. Every read is bounds-checked against `bytes`, so a count
 * the directory cannot hold fails on the first record past its end
 * rather than walking off into whatever follows.
 */
export function readCentralDirectory(
  bytes: Uint8Array,
  start: number,
  expectedCount: number,
): CentralDirEntry[] {
  const view = viewOf(bytes)
  const entries: CentralDirEntry[] = []
  const decoder = new TextDecoder()
  let pos = start

  for (let i = 0; i < expectedCount; i++) {
    if (pos + CENTRAL_DIR_ENTRY_MIN_BYTES > bytes.length) {
      throw new ZipError("Central Directory entry extends beyond file")
    }

    const sig = view.getUint32(pos, true)
    if (sig !== SIG_CENTRAL_DIR) {
      throw new ZipError(
        `Invalid Central Directory signature at offset ${pos}: 0x${sig.toString(16)}`,
      )
    }

    const generalFlag = view.getUint16(pos + 8, true)
    const compressionMethod = view.getUint16(pos + 10, true)
    const entryCrc32 = view.getUint32(pos + 16, true)
    let compressedSize = view.getUint32(pos + 20, true)
    let uncompressedSize = view.getUint32(pos + 24, true)
    const fileNameLength = view.getUint16(pos + 28, true)
    const extraFieldLength = view.getUint16(pos + 30, true)
    const commentLength = view.getUint16(pos + 32, true)
    let localHeaderOffset = view.getUint32(pos + 42, true)

    const fileNameBytes = bytes.subarray(pos + 46, pos + 46 + fileNameLength)
    const fileName = decoder.decode(fileNameBytes)

    // Any field left at the sentinel is carried at 64-bit width in the
    // ZIP64 extra field instead.
    if (
      compressedSize === ZIP64_SENTINEL ||
      uncompressedSize === ZIP64_SENTINEL ||
      localHeaderOffset === ZIP64_SENTINEL
    ) {
      const zip64 = readZip64Extra(
        bytes,
        pos + 46 + fileNameLength,
        extraFieldLength,
        uncompressedSize === ZIP64_SENTINEL,
        compressedSize === ZIP64_SENTINEL,
        localHeaderOffset === ZIP64_SENTINEL,
        fileName,
      )
      if (zip64.uncompressedSize !== undefined) uncompressedSize = zip64.uncompressedSize
      if (zip64.compressedSize !== undefined) compressedSize = zip64.compressedSize
      if (zip64.localHeaderOffset !== undefined) localHeaderOffset = zip64.localHeaderOffset
    }

    entries.push({
      fileName,
      compressionMethod,
      compressedSize,
      uncompressedSize,
      crc32: entryCrc32,
      localHeaderOffset,
      hasDataDescriptor: (generalFlag & 0x08) !== 0,
    })

    pos += 46 + fileNameLength + extraFieldLength + commentLength
  }

  return entries
}

/**
 * Pull the overflowed fields out of a ZIP64 extended-information extra
 * field — central directory or local header, the layout is the same.
 *
 * The record is positional: only the fields whose 32-bit counterpart held
 * the sentinel are present, always in the order size → compressed size →
 * local header offset → disk number.
 */
export function readZip64Extra(
  bytes: Uint8Array,
  offset: number,
  length: number,
  wantUncompressed: boolean,
  wantCompressed: boolean,
  wantOffset: boolean,
  fileName: string,
): { uncompressedSize?: number; compressedSize?: number; localHeaderOffset?: number } {
  const view = viewOf(bytes)
  const end = Math.min(offset + length, bytes.length)
  let pos = offset

  while (pos + 4 <= end) {
    const headerId = view.getUint16(pos, true)
    const dataSize = view.getUint16(pos + 2, true)
    const dataStart = pos + 4

    if (headerId !== ZIP64_EXTRA_ID) {
      pos = dataStart + dataSize
      continue
    }

    const dataEnd = dataStart + dataSize
    let cursor = dataStart
    const result: {
      uncompressedSize?: number
      compressedSize?: number
      localHeaderOffset?: number
    } = {}

    const take = (what: string): number => {
      if (cursor + 8 > dataEnd) {
        throw new ZipError(`Truncated ZIP64 extra field for entry: ${fileName}`)
      }
      const value = readUint64(bytes, cursor, what)
      cursor += 8
      return value
    }

    if (wantUncompressed) result.uncompressedSize = take("ZIP64 uncompressed size")
    if (wantCompressed) result.compressedSize = take("ZIP64 compressed size")
    if (wantOffset) result.localHeaderOffset = take("ZIP64 local header offset")

    return result
  }

  throw new ZipError(`Missing ZIP64 extra field for entry: ${fileName}`)
}

/** Read a 64-bit little-endian field, refusing values JS can't index with. */
export function readUint64(bytes: Uint8Array, offset: number, what: string): number {
  if (offset + 8 > bytes.length) {
    throw new ZipError(`${what} extends beyond file`)
  }
  const value = viewOf(bytes).getBigUint64(offset, true)
  if (value > BigInt(Number.MAX_SAFE_INTEGER)) {
    throw new ZipError(`${what} (${value}) exceeds the addressable range`)
  }
  return Number(value)
}

function viewOf(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

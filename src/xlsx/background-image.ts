// ── Sheet background images ──────────────────────────────────────────
// Shared by the authoring writer and the round-trip writer. They used to
// hold one copy of this logic each, which is how they came to share one
// bug: both hard-coded `.png` for every background image, so a JPEG was
// stored at `xl/media/imageN.png` and declared `image/png`. See #427.

import type { SheetImage } from "../_types"

/** The formats hucre declares a content type for. Matches `SheetImage["type"]`. */
type ImageFormat = SheetImage["type"]

/**
 * Identify an image from its leading bytes.
 *
 * A background image arrives as a bare `Uint8Array` — `Sheet.background
 * Image` carries no type, and the reader discards the source extension —
 * so the bytes are the only thing left to be faithful to. They are enough:
 * every format below is self-describing in its first few bytes, which is
 * also why adding a `type` field to the public shape would be API surface
 * for something derivable, and would do nothing for the round-trip path.
 *
 * Falls back to `"png"` for anything unrecognised, which is what the code
 * did unconditionally before — so an exotic format is no worse off than it
 * already was.
 */
export function sniffImageFormat(data: Uint8Array): ImageFormat {
  const b = data

  // \x89 P N G \r \n \x1a \n
  if (b.length >= 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
    return "png"
  }

  // SOI marker, then any JFIF/Exif/raw variant.
  if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return "jpeg"

  // "GIF87a" / "GIF89a"
  if (b.length >= 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return "gif"

  // "RIFF" .... "WEBP" — the size field sits between them, so check both.
  if (
    b.length >= 12 &&
    b[0] === 0x52 &&
    b[1] === 0x49 &&
    b[2] === 0x46 &&
    b[3] === 0x46 &&
    b[8] === 0x57 &&
    b[9] === 0x45 &&
    b[10] === 0x42 &&
    b[11] === 0x50
  ) {
    return "webp"
  }

  // SVG is text, and may lead with an XML declaration, a doctype, a BOM
  // or whitespace before the root element — so look for the tag rather
  // than anchoring at byte 0. Bounded to the head: a long binary file
  // must not be scanned end to end on the off-chance it spells "<svg".
  if (isSvg(b)) return "svg"

  return "png"
}

/** Bytes to inspect when sniffing SVG. Room for a declaration and a doctype. */
const SVG_SNIFF_BYTES = 1024

function isSvg(data: Uint8Array): boolean {
  const head = data.subarray(0, SVG_SNIFF_BYTES)
  // Latin-1 rather than UTF-8: the markers are ASCII, and decoding a
  // truncated multi-byte sequence must not throw.
  let text = ""
  for (let i = 0; i < head.length; i++) text += String.fromCharCode(head[i]!)
  const trimmed = text.trimStart().replace(/^﻿/, "")
  if (!trimmed.startsWith("<")) return false
  return /<svg[\s>]/i.test(trimmed)
}

/**
 * Assign a media path to every sheet that has a background image, in the
 * shared `xl/media/imageN` numbering.
 *
 * The counter is shared with drawing images on purpose: both land in
 * `xl/media`, so numbering them independently would collide and one would
 * overwrite the other.
 *
 * Returns one entry per sheet — `null` where the sheet has no background —
 * plus the advanced counter, and records each extension used so the
 * content types can declare it.
 */
export function assignBackgroundImagePaths(
  sheets: ReadonlyArray<{ backgroundImage?: Uint8Array }>,
  startIndex: number,
  imageExtensions: Set<string>,
): { paths: Array<string | null>; nextIndex: number } {
  const paths: Array<string | null> = []
  let index = startIndex

  for (const sheet of sheets) {
    if (!sheet.backgroundImage) {
      paths.push(null)
      continue
    }
    const ext = sniffImageFormat(sheet.backgroundImage)
    paths.push(`xl/media/image${index}.${ext}`)
    imageExtensions.add(ext)
    index++
  }

  return { paths, nextIndex: index }
}

import { describe, expect, it } from "vitest"
import { readFileSync, readdirSync } from "node:fs"
import { parseWorksheet, parseWorksheetStream, type WorksheetContext } from "../src/xlsx/worksheet"
import { ZipReader } from "../src/zip/reader"
import { parseSharedStrings } from "../src/xlsx/shared-strings"
import { parseStyles } from "../src/xlsx/styles"

// ═══════════════════════════════════════════════════════════════════════
// #503 gave the worksheet reader a second driver: a part over the string
// ceiling is parsed from the ZIP entry's stream, a chunk at a time, by
// the same handlers the buffered parse uses. Sharing the handlers means
// the two cannot disagree about a *field* — there is only one
// implementation of each.
//
// What they can still disagree about is the chunk boundary. A tag, an
// attribute, an entity, a multibyte sequence or a CRLF split across two
// chunks is state the buffered parser never has to carry, and every one
// of those is a place to lose a character or gain one. #536 found a real
// instance while writing the driver — a cut between a CR and its LF
// became two line endings, because `normalizeEol` runs per piece.
//
// That PR tests the boundary byte-at-a-time on XML it wrote. This runs
// the same comparison over **files hucre did not write** — the #464
// corpus, Excel 16.0 and openpyxl and ExcelJS — at ten chunk sizes
// chosen to be hostile, on the parts as they really are: real shared
// strings, real styles, real inline runs, real unicode.
//
// It is a characterisation test. It asserts that the two drivers agree,
// not what they agree on, because what they agree on is already asserted
// everywhere else. Checked for teeth rather than assumed: dropping the
// carried tail in `parseSaxStream` — the classic streaming bug — turns it
// red.
//
// What it does *not* reach, so nobody writes it twice: the guards inside
// `safeTextSplit`. Those need a single text run over 256 KiB and a chunk
// boundary landing on the exact character, and no real fixture has a cell
// that long. #536 covers the CRLF one directly, with the run lengths
// worked out to put the break on the CR. The surrogate one is defensive —
// `TextDecoder` in stream mode completes a character before emitting it,
// so the buffer never ends on a lone high surrogate.
// ═══════════════════════════════════════════════════════════════════════

const dec = new TextDecoder()

/**
 * Sizes chosen to be hostile. `1` splits everything; the primes land
 * boundaries at offsets a power of two never reaches, which is where a
 * multibyte sequence or an entity gets cut.
 */
const CHUNK_SIZES = [1, 2, 3, 7, 13, 31, 61, 127, 251, 1021]

function chunked(data: Uint8Array, size: number): ReadableStream<Uint8Array> {
  let at = 0
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (at >= data.length) return controller.close()
      const end = Math.min(at + size, data.length)
      controller.enqueue(data.subarray(at, end))
      at = end
    },
  })
}

/** `Map` and `Date` do not survive `JSON.stringify` unaided. */
function replacer(_key: string, value: unknown): unknown {
  if (value instanceof Map) return { __map: [...value.entries()].sort() }
  if (value instanceof Date) return `__date:${value.toISOString()}`
  return value
}

function corpus(): string[] {
  const out: string[] = []
  for (const dir of ["test/fixtures", "test/fixtures/third-party"]) {
    for (const file of readdirSync(dir)) {
      if (file.endsWith(".xlsx")) out.push(`${dir}/${file}`)
    }
  }
  return out.sort()
}

describe("the streaming worksheet driver, cut anywhere", () => {
  it("builds what the buffered driver builds, on every third-party fixture", async () => {
    const diffs: string[] = []

    for (const path of corpus()) {
      const zip = new ZipReader(new Uint8Array(readFileSync(path)))

      const sharedStrings = zip.has("xl/sharedStrings.xml")
        ? parseSharedStrings(dec.decode(await zip.extract("xl/sharedStrings.xml")))
        : []
      const styles = zip.has("xl/styles.xml")
        ? parseStyles(dec.decode(await zip.extract("xl/styles.xml")))
        : null

      for (const entry of zip.entries()) {
        if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(entry)) continue

        const raw = await zip.extract(entry)
        // `readStyles: true` so the comparison covers the style index on
        // every cell, not just its value.
        const ctx = (): WorksheetContext => ({
          sharedStrings,
          styles,
          readStyles: true,
          dateSystem: "1900",
        })

        // A fixture over the bounding-box limit throws on both drivers,
        // and agreeing about the failure is as much a claim as agreeing
        // about the Sheet — `excel-sparse.xlsx` is the one that does.
        let want: string
        try {
          want = JSON.stringify(parseWorksheet(dec.decode(raw), "S", ctx()), replacer)
        } catch (error) {
          want = `THREW ${(error as Error).message}`
        }

        for (const size of CHUNK_SIZES) {
          let got: string
          try {
            got = JSON.stringify(
              await parseWorksheetStream(chunked(raw, size), "S", ctx()),
              replacer,
            )
          } catch (error) {
            got = `THREW ${(error as Error).message}`
          }

          if (got !== want && diffs.length < 6) {
            diffs.push(
              `${path} ${entry} chunk=${size}\n  want ${want.slice(0, 220)}\n  got  ${got.slice(0, 220)}`,
            )
          }
        }
      }
    }

    expect(diffs).toEqual([])
  })
})

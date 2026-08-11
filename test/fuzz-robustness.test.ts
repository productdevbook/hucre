import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { writeOds } from "../src/ods/writer"
import { readXlsx } from "../src/xlsx/reader"
import { readOds } from "../src/ods/reader"
import { openXlsx } from "../src/xlsx/roundtrip"
import { parseCsv } from "../src/csv/reader"
import { read } from "../src/defter"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { HucreError } from "../src/errors"
import { byteMutations, seeded, XML_MUTATORS } from "./_fuzz"

// ═══════════════════════════════════════════════════════════════════════
// #473 — every bound in `src/limits.ts` is tested with a hand-written
// case that reaches exactly that bound. The interesting failures are the
// ones nobody thought of: a byte flipped in a ZIP central directory, a
// `<c>` whose `r` is `"A"` with no row, a part that ends mid-element.
//
// The bar is the one limits.ts already states — "a typed error instead of
// killing the process" — so that is what these assert, and nothing more.
// A fuzzer that pinned the *value* a corrupt file reads back would be
// asserting an accident.
//
// Two rules for everything here:
//   - Seeded. A failure you cannot reproduce has told you nothing, and a
//     suite that runs different cases each time makes a red build a coin
//     toss.
//   - Bounded. A few hundred cases is seconds, which is what keeps this
//     in CI rather than in a nightly nobody reads.
//
// This found one thing on its first run: corrupt DEFLATE data threw a
// bare `Error`, so a caller doing `catch (e) { if (e instanceof
// HucreError) }` — the documented contract — missed it entirely.
// ═══════════════════════════════════════════════════════════════════════

const SEED = 0xc0ffee
const ITERATIONS = 240

async function fixtureXlsx(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "S",
        rows: [
          ["a", 1, true],
          ["b", 2.5, null],
          ["c", "text", new Date("2024-01-15T00:00:00Z")],
        ],
        merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
        conditionalRules: [
          { type: "cellIs", priority: 1, range: "A1:B2", operator: "greaterThan", formula: "1" },
        ],
      },
    ],
  })
}

/**
 * The whole contract, in one function.
 *
 * Succeeding is fine — leniency is the documented default. Throwing is
 * fine, as long as it is a `HucreError`. Anything else is the bug: a raw
 * `TypeError` or `RangeError` is the library falling over rather than
 * refusing.
 */
async function mustBeTypedOrFine(run: () => Promise<unknown>, label: string): Promise<void> {
  try {
    await run()
  } catch (error) {
    expect(
      error instanceof HucreError,
      `${label} threw ${(error as Error)?.constructor?.name}: ${(error as Error)?.message}`,
    ).toBe(true)
  }
}

describe("byte-level corruption never escapes the error hierarchy", () => {
  it("readXlsx", async () => {
    const base = await fixtureXlsx()
    for (const { label, bytes } of byteMutations(base, ITERATIONS, SEED)) {
      await mustBeTypedOrFine(() => readXlsx(bytes), `readXlsx ${label}`)
    }
  })

  it("openXlsx, which keeps far more of the package", async () => {
    const base = await fixtureXlsx()
    for (const { label, bytes } of byteMutations(base, ITERATIONS, SEED)) {
      await mustBeTypedOrFine(() => openXlsx(bytes), `openXlsx ${label}`)
    }
  })

  it("readOds", async () => {
    const base = await writeOds({
      sheets: [
        {
          name: "S",
          rows: [
            ["a", 1],
            ["b", 2],
          ],
        },
      ],
    })
    for (const { label, bytes } of byteMutations(base, ITERATIONS, SEED)) {
      await mustBeTypedOrFine(() => readOds(bytes), `readOds ${label}`)
    }
  })

  it("read(), which has to sniff before it can dispatch", async () => {
    const base = await fixtureXlsx()
    for (const { label, bytes } of byteMutations(base, ITERATIONS, SEED)) {
      await mustBeTypedOrFine(() => read(bytes), `read ${label}`)
    }
  })

  it("parseCsv, where every byte sequence is arguably valid", async () => {
    const base = new TextEncoder().encode('name,qty\n"a""b",1\nc,2\n')
    for (const { label, bytes } of byteMutations(base, ITERATIONS, SEED)) {
      await mustBeTypedOrFine(async () => parseCsv(bytes), `parseCsv ${label}`)
    }
  })
})

describe("structure-aware corruption reaches the parsers, not just the ZIP", () => {
  /** Rebuild a valid archive with one part replaced. */
  async function repackage(
    all: Map<string, Uint8Array>,
    path: string,
    xml: string,
  ): Promise<Uint8Array> {
    const zw = new ZipWriter()
    for (const [name, data] of all) {
      zw.add(name, name === path ? new TextEncoder().encode(xml) : data)
    }
    return zw.build()
  }

  it("every XML part, every mutator, still typed", async () => {
    // Byte mutations mostly break the DEFLATE stream, which exercises one
    // layer over and over. These arrive at the parsers intact-but-wrong,
    // which is where the interesting failures live.
    const base = await fixtureXlsx()
    const all = await new ZipReader(base).extractAll()
    const parts = [...all.keys()].filter((p) => p.endsWith(".xml") || p.endsWith(".rels"))
    const rnd = seeded(SEED)

    expect(parts.length).toBeGreaterThan(4)

    for (const part of parts) {
      const original = new TextDecoder().decode(all.get(part)!)
      for (const [label, mutate] of XML_MUTATORS) {
        const mutated = mutate(original, rnd)
        if (mutated === original) continue
        const bytes = await repackage(all, part, mutated)

        await mustBeTypedOrFine(() => readXlsx(bytes), `readXlsx ${part} ${label}`)
        await mustBeTypedOrFine(() => openXlsx(bytes), `openXlsx ${part} ${label}`)
      }
    }
  }, 60_000)
})

describe("corruption is refused promptly", () => {
  it("no mutation takes materially longer than the clean file", async () => {
    // The failure mode `limits.ts` exists to prevent is not a wrong
    // answer, it is a process that never comes back. A quadratic parse or
    // an unbounded allocation shows up here as a wall-clock outlier.
    const base = await fixtureXlsx()

    const t0 = performance.now()
    await readXlsx(base)
    const clean = performance.now() - t0

    let worst = 0
    for (const { bytes } of byteMutations(base, 120, SEED)) {
      const start = performance.now()
      try {
        await readXlsx(bytes)
      } catch {
        // The point is the clock, not the outcome.
      }
      worst = Math.max(worst, performance.now() - start)
    }

    // Generous on purpose: this is a hang detector, not a benchmark, and
    // a tight bound here would fail on a loaded CI runner rather than on
    // a real regression.
    expect(worst).toBeLessThan(Math.max(2000, clean * 200))
  }, 60_000)
})

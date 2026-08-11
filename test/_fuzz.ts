// ── Deterministic fuzzing helpers ───────────────────────────────────
//
// `src/limits.ts` and `SECURITY.md` describe a library that expects
// hostile files, and every bound in them is tested — with a hand-written
// case that reaches exactly that bound. The interesting failures are the
// ones nobody thought of. See #473.
//
// Everything here is seeded. A fuzzer that finds a failure you cannot
// reproduce has told you something useless, and one that runs different
// cases every CI run makes a red build a coin toss. The seed is in the
// test; change it deliberately to search elsewhere.

/**
 * mulberry32 — a small, fast, well-distributed PRNG.
 *
 * `Math.random()` cannot be seeded, and a fuzzer without a seed cannot
 * hand you back the case that broke.
 */
export function seeded(seed: number): () => number {
  let a = seed
  return () => {
    a |= 0
    a = (a + 0x6d2b79f5) | 0
    let t = Math.imul(a ^ (a >>> 15), 1 | a)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

/** One mutated input, with a label that says how to rebuild it. */
export interface Mutation {
  label: string
  bytes: Uint8Array
}

/**
 * Byte-level mutations of a valid file.
 *
 * Four kinds, because they break different layers: a flipped bit usually
 * lands in compressed data, a truncation usually kills the central
 * directory, a random byte can do either, and a scatter reaches several
 * at once.
 */
export function* byteMutations(
  source: Uint8Array,
  count: number,
  seed: number,
): Generator<Mutation> {
  const rnd = seeded(seed)
  for (let i = 0; i < count; i++) {
    const copy = new Uint8Array(source)
    const at = Math.floor(rnd() * Math.max(1, copy.length))
    switch (i % 4) {
      case 0:
        copy[at] ^= 1 << Math.floor(rnd() * 8)
        yield { label: `flip@${at}`, bytes: copy }
        break
      case 1:
        yield { label: `truncate@${at}`, bytes: copy.subarray(0, at) }
        break
      case 2:
        copy[at] = Math.floor(rnd() * 256)
        yield { label: `set@${at}`, bytes: copy }
        break
      default: {
        const n = 1 + Math.floor(rnd() * 8)
        for (let k = 0; k < n; k++) {
          copy[Math.floor(rnd() * copy.length)] = Math.floor(rnd() * 256)
        }
        yield { label: `scatter${n}`, bytes: copy }
        break
      }
    }
  }
}

/**
 * Text mutations of one XML part, applied inside an otherwise valid
 * package.
 *
 * Byte mutations mostly break the DEFLATE stream, which exercises one
 * layer over and over. These reach the parsers: a `<c>` whose `r` is
 * `"A"` with no row, a part that ends mid-element, a `count` of
 * 999999999999, a reference past the last legal cell.
 */
export const XML_MUTATORS: Array<[string, (xml: string, rnd: () => number) => string]> = [
  ["truncate", (x, rnd) => x.slice(0, Math.floor(rnd() * x.length))],
  [
    "drop-a-close-tag",
    (x, rnd) => {
      const all = [...x.matchAll(/<\/[\w:]+>/g)]
      if (all.length === 0) return x
      const pick = all[Math.floor(rnd() * all.length)]!
      return x.slice(0, pick.index) + x.slice(pick.index! + pick[0].length)
    },
  ],
  ["ref-without-row", (x) => x.replace(/ r="[A-Z]+\d+"/g, ' r="A"')],
  ["ref-past-the-grid", (x) => x.replace(/ r="[A-Z]+\d+"/g, ' r="XFD1048577"')],
  ["negative-numbers", (x) => x.replace(/="(\d+)"/g, '="-$1"')],
  ["huge-numbers", (x) => x.replace(/="(\d+)"/g, '="999999999999"')],
  ["not-a-number", (x) => x.replace(/="(\d+)"/g, '="NaN"')],
  ["every-attribute-empty", (x) => x.replace(/="[^"]*"/g, '=""')],
  [
    "unknown-element",
    (x, rnd) => {
      const all = [...x.matchAll(/<([\w:]+)[ >]/g)]
      if (all.length < 2) return x
      const pick = all[Math.floor(rnd() * all.length)]!
      return `${x.slice(0, pick.index)}<zzz ${x.slice(pick.index! + pick[0].length)}`
    },
  ],
  ["no-xml-declaration", (x) => x.replace(/<\?xml[^>]*\?>/, "")],
  [
    "entity-expansion",
    (x) =>
      x.replace(
        "<sheetData>",
        '<!DOCTYPE t [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">' +
          '<!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;"><!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;&c;&c;">]>' +
          "<sheetData>",
      ),
  ],
]

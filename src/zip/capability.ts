// ── Compression capability ───────────────────────────────────────────
//
// `deflate-raw` is the format a ZIP needs, and it is not the same
// question as "does this runtime have CompressionStream". Node 18 had the
// constructor and rejected `deflate-raw` — it shipped with `gzip` and
// `deflate`, and the raw format arrived in Node 20.
//
// Node is no longer the case this defends against: the floor is 24 (see
// CLAUDE.md). It stays because hucre runs in browsers and on Workers too,
// where the same split exists and the version is not ours to choose.
//
// Four modules each kept their own memoized flag, and every one of them
// probed the constructor's *existence*. The buffered writer got away with
// it because it also wrapped the construction in a try/catch and fell back
// to the pure-TypeScript DEFLATE; the streaming writer constructed
// outside any try and threw `ERR_INVALID_ARG_VALUE` on the first chunk —
// so `writeXlsxStream` did not work on the Node version the package
// claims as its floor. See #439 §AN.
//
// The only probe that answers the question is to build one and see.

let deflateRaw: boolean | undefined
let inflateRaw: boolean | undefined

/** Whether this runtime can DEFLATE with the raw format a ZIP entry needs. */
export function canDeflateRaw(): boolean {
  if (deflateRaw === undefined) {
    try {
      // Constructed, not just referenced: Node 18 has CompressionStream
      // and throws here.
      new CompressionStream("deflate-raw")
      deflateRaw = typeof ReadableStream !== "undefined"
    } catch {
      deflateRaw = false
    }
  }
  return deflateRaw
}

/** Whether this runtime can inflate the raw format a ZIP entry carries. */
export function canInflateRaw(): boolean {
  if (inflateRaw === undefined) {
    try {
      new DecompressionStream("deflate-raw")
      inflateRaw = typeof ReadableStream !== "undefined"
    } catch {
      inflateRaw = false
    }
  }
  return inflateRaw
}

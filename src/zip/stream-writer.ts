// ── Streaming ZIP Writer ─────────────────────────────────────────────
// Emits a ZIP archive as a byte stream: for every entry a local file
// header goes out first, then the (optionally DEFLATE-compressed) bytes
// as they arrive, then a trailing data descriptor. The central directory
// and EOCD close the archive.
//
// Because an entry's size and CRC are unknown when its header is
// written, every entry sets general-purpose bit 3 and carries a data
// descriptor — the same layout `archiver`/`zip-stream` produce, which
// Excel, LibreOffice, and the platform unzip tools all accept.
//
// Nothing is buffered beyond the current chunk plus one central
// directory record per entry, so peak memory is O(entries), not
// O(archive size).

import { CRC32_INIT, crc32Update, crc32Final } from "./deflate"
import { ZipError } from "../errors"

// ── ZIP Signatures ──────────────────────────────────────────────────

const SIG_LOCAL_FILE = 0x04034b50
const SIG_DATA_DESCRIPTOR = 0x08074b50
const SIG_CENTRAL_DIR = 0x02014b50
const SIG_END_OF_CENTRAL_DIR = 0x06054b50

/** General purpose bit 3 — sizes and CRC live in a trailing descriptor. */
const FLAG_DATA_DESCRIPTOR = 0x0008

/** Largest value a 32-bit ZIP size/offset field can hold. */
const MAX_UINT32 = 0xffffffff

const METHOD_STORE = 0
const METHOD_DEFLATE = 8

const encoder = /* @__PURE__ */ new TextEncoder()

// ── Types ───────────────────────────────────────────────────────────

/** Bytes for one entry — a whole buffer, or chunks pulled on demand. */
export type ZipStreamData = Uint8Array | AsyncIterable<Uint8Array> | Iterable<Uint8Array>

export interface ZipStreamEntry {
  /** Path inside the archive, e.g. `xl/worksheets/sheet1.xml` */
  path: string
  /** Entry contents */
  data: ZipStreamData
  /**
   * DEFLATE the entry. Default `true`. Falls back to STORE when the
   * platform has no `CompressionStream` — the pure-TS DEFLATE needs the
   * whole buffer up front, which would defeat streaming.
   */
  compress?: boolean
}

/** Entries may themselves be produced lazily (that's the point). */
export type ZipStreamEntries = AsyncIterable<ZipStreamEntry> | Iterable<ZipStreamEntry>

// ── Compression ─────────────────────────────────────────────────────

let hasCompressionStream: boolean | undefined

function checkCompressionStream(): boolean {
  if (hasCompressionStream === undefined) {
    try {
      hasCompressionStream = typeof CompressionStream !== "undefined"
    } catch {
      hasCompressionStream = false
    }
  }
  return hasCompressionStream
}

/** Pipe chunks through `deflate-raw`, preserving backpressure both ways. */
async function* deflateRawChunks(source: AsyncIterable<Uint8Array>): AsyncGenerator<Uint8Array> {
  const cs = new CompressionStream("deflate-raw")
  const writer = cs.writable.getWriter()

  let pumpError: unknown
  const pump = (async () => {
    try {
      for await (const chunk of source) {
        // `ready` is what makes the producer wait when the consumer
        // stops pulling — this is the backpressure hinge.
        await writer.ready
        await writer.write(chunk as unknown as BufferSource)
      }
      await writer.close()
    } catch (err) {
      pumpError = err
      try {
        await writer.abort(err)
      } catch {
        // Already errored — nothing to salvage.
      }
    }
  })()

  const reader = cs.readable.getReader()
  let drained = false
  try {
    for (;;) {
      const { done, value } = await reader.read()
      if (done) {
        drained = true
        break
      }
      if (value && value.length > 0) yield value
    }
  } finally {
    if (!drained) {
      try {
        await reader.cancel()
      } catch {
        // Consumer walked away; the pump rejection is swallowed above.
      }
    }
  }

  await pump
  if (pumpError) throw pumpError
}

/** Normalize entry data into an async chunk source. */
async function* toChunks(data: ZipStreamData): AsyncGenerator<Uint8Array> {
  if (data instanceof Uint8Array) {
    if (data.length > 0) yield data
    return
  }
  for await (const chunk of data as AsyncIterable<Uint8Array>) {
    if (chunk.length > 0) yield chunk
  }
}

/** Pass chunks through untouched while feeding them to `onChunk`. */
async function* tapChunks(
  source: AsyncIterable<Uint8Array>,
  onChunk: (chunk: Uint8Array) => void,
): AsyncGenerator<Uint8Array> {
  for await (const chunk of source) {
    onChunk(chunk)
    yield chunk
  }
}

// ── Header / trailer builders ───────────────────────────────────────

interface CentralRecord {
  pathBytes: Uint8Array
  method: number
  crc: number
  compressedSize: number
  uncompressedSize: number
  localOffset: number
}

function buildLocalHeader(pathBytes: Uint8Array, method: number): Uint8Array {
  const buf = new Uint8Array(30 + pathBytes.length)
  const view = new DataView(buf.buffer)

  view.setUint32(0, SIG_LOCAL_FILE, true)
  view.setUint16(4, 20, true) // Version needed (2.0)
  view.setUint16(6, FLAG_DATA_DESCRIPTOR, true)
  view.setUint16(8, method, true)
  view.setUint16(10, 0, true) // Mod time
  view.setUint16(12, 0x0021, true) // Mod date (1980-01-01)
  view.setUint32(14, 0, true) // CRC — in the data descriptor
  view.setUint32(18, 0, true) // Compressed size — in the data descriptor
  view.setUint32(22, 0, true) // Uncompressed size — in the data descriptor
  view.setUint16(26, pathBytes.length, true)
  view.setUint16(28, 0, true) // Extra field length

  buf.set(pathBytes, 30)
  return buf
}

function buildDataDescriptor(crc: number, compressedSize: number, uncompressedSize: number) {
  const buf = new Uint8Array(16)
  const view = new DataView(buf.buffer)

  view.setUint32(0, SIG_DATA_DESCRIPTOR, true)
  view.setUint32(4, crc, true)
  view.setUint32(8, compressedSize, true)
  view.setUint32(12, uncompressedSize, true)
  return buf
}

function buildCentralDirectory(records: CentralRecord[], centralDirOffset: number): Uint8Array {
  let size = 22 // EOCD
  for (const rec of records) size += 46 + rec.pathBytes.length

  const buf = new Uint8Array(size)
  const view = new DataView(buf.buffer)
  let offset = 0

  for (const rec of records) {
    view.setUint32(offset, SIG_CENTRAL_DIR, true)
    view.setUint16(offset + 4, 20, true) // Version made by (2.0)
    view.setUint16(offset + 6, 20, true) // Version needed (2.0)
    view.setUint16(offset + 8, FLAG_DATA_DESCRIPTOR, true)
    view.setUint16(offset + 10, rec.method, true)
    view.setUint16(offset + 12, 0, true) // Mod time
    view.setUint16(offset + 14, 0x0021, true) // Mod date (1980-01-01)
    view.setUint32(offset + 16, rec.crc, true)
    view.setUint32(offset + 20, rec.compressedSize, true)
    view.setUint32(offset + 24, rec.uncompressedSize, true)
    view.setUint16(offset + 28, rec.pathBytes.length, true)
    view.setUint16(offset + 30, 0, true) // Extra field length
    view.setUint16(offset + 32, 0, true) // File comment length
    view.setUint16(offset + 34, 0, true) // Disk number start
    view.setUint16(offset + 36, 0, true) // Internal file attributes
    view.setUint32(offset + 38, 0, true) // External file attributes
    view.setUint32(offset + 42, rec.localOffset, true)
    buf.set(rec.pathBytes, offset + 46)
    offset += 46 + rec.pathBytes.length
  }

  const centralDirSize = offset

  view.setUint32(offset, SIG_END_OF_CENTRAL_DIR, true)
  view.setUint16(offset + 4, 0, true) // Disk number
  view.setUint16(offset + 6, 0, true) // Central dir start disk
  view.setUint16(offset + 8, records.length, true)
  view.setUint16(offset + 10, records.length, true)
  view.setUint32(offset + 12, centralDirSize, true)
  view.setUint32(offset + 16, centralDirOffset, true)
  view.setUint16(offset + 20, 0, true) // Comment length

  return buf
}

// ── Streaming writer ────────────────────────────────────────────────

function assertFits(value: number, what: string, path: string): void {
  if (value > MAX_UINT32) {
    throw new ZipError(
      `ZIP ${what} exceeds the 4 GiB limit at "${path}". ` +
        `Zip64 is not supported yet — split the data across more parts.`,
    )
  }
}

/**
 * Emit a ZIP archive chunk by chunk.
 *
 * Each entry's data is consumed to completion before the next entry is
 * pulled from `entries`, so an async generator can safely compute a
 * later entry (styles, shared strings, the workbook index) from state
 * accumulated while earlier entries streamed out.
 */
export async function* zipStreamChunks(entries: ZipStreamEntries): AsyncGenerator<Uint8Array> {
  const records: CentralRecord[] = []
  const seen = new Set<string>()
  let offset = 0
  const canDeflate = checkCompressionStream()

  for await (const entry of entries) {
    if (seen.has(entry.path)) {
      throw new ZipError(`Duplicate ZIP entry: "${entry.path}"`)
    }
    seen.add(entry.path)

    const pathBytes = encoder.encode(entry.path)
    const method = entry.compress !== false && canDeflate ? METHOD_DEFLATE : METHOD_STORE
    const localOffset = offset

    const header = buildLocalHeader(pathBytes, method)
    offset += header.length
    yield header

    let crcState = CRC32_INIT
    let uncompressedSize = 0
    let compressedSize = 0

    const raw = tapChunks(toChunks(entry.data), (chunk) => {
      crcState = crc32Update(crcState, chunk)
      uncompressedSize += chunk.length
    })

    const body = method === METHOD_DEFLATE ? deflateRawChunks(raw) : raw

    for await (const chunk of body) {
      compressedSize += chunk.length
      offset += chunk.length
      yield chunk
    }

    assertFits(uncompressedSize, "uncompressed size", entry.path)
    assertFits(compressedSize, "compressed size", entry.path)

    const descriptor = buildDataDescriptor(crc32Final(crcState), compressedSize, uncompressedSize)
    offset += descriptor.length
    yield descriptor

    assertFits(offset, "archive size", entry.path)

    records.push({
      pathBytes,
      method,
      crc: crc32Final(crcState),
      compressedSize,
      uncompressedSize,
      localOffset,
    })
  }

  yield buildCentralDirectory(records, offset)
}

/** Wrap {@link zipStreamChunks} in a `ReadableStream` of bytes. */
export function zipStream(entries: ZipStreamEntries): ReadableStream<Uint8Array> {
  const chunks = zipStreamChunks(entries)

  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      try {
        const { done, value } = await chunks.next()
        if (done) {
          controller.close()
          return
        }
        controller.enqueue(value)
      } catch (err) {
        controller.error(err)
      }
    },
    async cancel(reason) {
      await chunks.return?.(reason)
    },
  })
}

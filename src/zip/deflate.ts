// ── Pure TypeScript DEFLATE (RFC 1951) ─────────────────────────────
// Fallback for environments without CompressionStream/DecompressionStream.

// ── CRC-32 ──────────────────────────────────────────────────────────

const crcTable = /* @__PURE__ */ (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c;
  }
  return table;
})();

export function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) {
    crc = crcTable[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

// ── Bit Reader (for inflate) ────────────────────────────────────────

class BitReader {
  private pos = 0;
  private bitBuf = 0;
  private bitCount = 0;

  constructor(private data: Uint8Array) {}

  readBits(n: number): number {
    while (this.bitCount < n) {
      if (this.pos >= this.data.length) {
        throw new Error("Unexpected end of DEFLATE data");
      }
      this.bitBuf |= this.data[this.pos++] << this.bitCount;
      this.bitCount += 8;
    }
    const val = this.bitBuf & ((1 << n) - 1);
    this.bitBuf >>>= n;
    this.bitCount -= n;
    return val;
  }

  /** Align to byte boundary (discard remaining bits in current byte) */
  alignToByte(): void {
    this.bitBuf = 0;
    this.bitCount = 0;
  }

  readByte(): number {
    if (this.pos >= this.data.length) {
      throw new Error("Unexpected end of DEFLATE data");
    }
    return this.data[this.pos++];
  }

  readUint16LE(): number {
    const lo = this.readByte();
    const hi = this.readByte();
    return lo | (hi << 8);
  }

  get offset(): number {
    return this.pos;
  }

  get available(): boolean {
    return this.pos < this.data.length || this.bitCount > 0;
  }
}

// ── Huffman Tree ────────────────────────────────────────────────────

interface HuffmanTree {
  /** For each bit length, the starting code value */
  counts: Uint16Array;
  /** Symbol table indexed by canonical code */
  symbols: Uint16Array;
}

function buildHuffmanTree(codeLengths: Uint8Array | number[], maxSymbol: number): HuffmanTree {
  const maxBits = 15;
  const counts = new Uint16Array(maxBits + 1);

  // Count code lengths
  for (let i = 0; i < maxSymbol; i++) {
    if (codeLengths[i]) {
      counts[codeLengths[i]]++;
    }
  }

  // Build offset table
  const offsets = new Uint16Array(maxBits + 1);
  for (let i = 1; i < maxBits; i++) {
    offsets[i + 1] = offsets[i] + counts[i];
  }

  // Build symbol table
  const totalSymbols = offsets[maxBits] + counts[maxBits];
  const symbols = new Uint16Array(totalSymbols);

  for (let i = 0; i < maxSymbol; i++) {
    if (codeLengths[i]) {
      symbols[offsets[codeLengths[i]]++] = i;
    }
  }

  return { counts, symbols };
}

function decodeSymbol(reader: BitReader, tree: HuffmanTree): number {
  let code = 0;
  let first = 0;
  let index = 0;

  for (let len = 1; len <= 15; len++) {
    code |= reader.readBits(1);
    const count = tree.counts[len];
    if (code < first + count) {
      return tree.symbols[index + (code - first)];
    }
    index += count;
    first = (first + count) << 1;
    code <<= 1;
  }

  throw new Error("Invalid Huffman code");
}

// ── Fixed Huffman Tables (RFC 1951 Section 3.2.6) ───────────────────

const fixedLitLenTree = /* @__PURE__ */ (() => {
  const lengths = new Uint8Array(288);
  // 0-143: 8 bits
  for (let i = 0; i <= 143; i++) lengths[i] = 8;
  // 144-255: 9 bits
  for (let i = 144; i <= 255; i++) lengths[i] = 9;
  // 256-279: 7 bits
  for (let i = 256; i <= 279; i++) lengths[i] = 7;
  // 280-287: 8 bits
  for (let i = 280; i <= 287; i++) lengths[i] = 8;
  return buildHuffmanTree(lengths, 288);
})();

const fixedDistTree = /* @__PURE__ */ (() => {
  const lengths = new Uint8Array(32);
  for (let i = 0; i < 32; i++) lengths[i] = 5;
  return buildHuffmanTree(lengths, 32);
})();

// ── Length / Distance Extra Bits Tables ─────────────────────────────

const lengthBase = [
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131,
  163, 195, 227, 258,
];

const lengthExtra = [
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
];

const distBase = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049,
  3073, 4097, 6145, 8193, 12289, 16385, 24577,
];

const distExtra = [
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13,
];

// ── Code Length Order (RFC 1951 Section 3.2.7) ──────────────────────

const codeLengthOrder = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

// ── Inflate ─────────────────────────────────────────────────────────

/**
 * Decompress raw DEFLATE data (no zlib/gzip header).
 * Pure TypeScript implementation of RFC 1951.
 */
export function inflate(data: Uint8Array): Uint8Array {
  const reader = new BitReader(data);
  // Dynamic output buffer
  let output = new Uint8Array(data.length * 3 || 1024);
  let outPos = 0;

  function ensureCapacity(needed: number): void {
    if (outPos + needed > output.length) {
      let newSize = output.length * 2;
      while (outPos + needed > newSize) {
        newSize *= 2;
      }
      const newBuf = new Uint8Array(newSize);
      newBuf.set(output);
      output = newBuf;
    }
  }

  function writeByte(b: number): void {
    ensureCapacity(1);
    output[outPos++] = b;
  }

  function copyFromOutput(distance: number, length: number): void {
    ensureCapacity(length);
    let srcPos = outPos - distance;
    // Byte-by-byte copy needed for overlapping references
    for (let i = 0; i < length; i++) {
      output[outPos++] = output[srcPos++];
    }
  }

  function inflateBlock(litLenTree: HuffmanTree, distTree: HuffmanTree): void {
    for (;;) {
      const sym = decodeSymbol(reader, litLenTree);

      if (sym < 256) {
        // Literal byte
        writeByte(sym);
      } else if (sym === 256) {
        // End of block
        return;
      } else {
        // Length/distance pair
        const lenIdx = sym - 257;
        const length = lengthBase[lenIdx] + reader.readBits(lengthExtra[lenIdx]);

        const distSym = decodeSymbol(reader, distTree);
        const distance = distBase[distSym] + reader.readBits(distExtra[distSym]);

        copyFromOutput(distance, length);
      }
    }
  }

  let bfinal = 0;
  do {
    bfinal = reader.readBits(1);
    const btype = reader.readBits(2);

    if (btype === 0) {
      // No compression (stored block)
      reader.alignToByte();
      const len = reader.readUint16LE();
      const _nlen = reader.readUint16LE();
      ensureCapacity(len);
      for (let i = 0; i < len; i++) {
        output[outPos++] = reader.readByte();
      }
    } else if (btype === 1) {
      // Fixed Huffman codes
      inflateBlock(fixedLitLenTree, fixedDistTree);
    } else if (btype === 2) {
      // Dynamic Huffman codes
      const hlit = reader.readBits(5) + 257;
      const hdist = reader.readBits(5) + 1;
      const hclen = reader.readBits(4) + 4;

      // Read code length code lengths
      const codeLenCodeLens = new Uint8Array(19);
      for (let i = 0; i < hclen; i++) {
        codeLenCodeLens[codeLengthOrder[i]] = reader.readBits(3);
      }

      const codeLenTree = buildHuffmanTree(codeLenCodeLens, 19);

      // Read literal/length + distance code lengths
      const totalCodes = hlit + hdist;
      const allLengths = new Uint8Array(totalCodes);
      let idx = 0;

      while (idx < totalCodes) {
        const sym = decodeSymbol(reader, codeLenTree);

        if (sym < 16) {
          allLengths[idx++] = sym;
        } else if (sym === 16) {
          // Repeat previous length 3-6 times
          const repeat = reader.readBits(2) + 3;
          const prev = idx > 0 ? allLengths[idx - 1] : 0;
          for (let i = 0; i < repeat; i++) {
            allLengths[idx++] = prev;
          }
        } else if (sym === 17) {
          // Repeat 0 for 3-10 times
          const repeat = reader.readBits(3) + 3;
          for (let i = 0; i < repeat; i++) {
            allLengths[idx++] = 0;
          }
        } else if (sym === 18) {
          // Repeat 0 for 11-138 times
          const repeat = reader.readBits(7) + 11;
          for (let i = 0; i < repeat; i++) {
            allLengths[idx++] = 0;
          }
        }
      }

      const litLenLengths = allLengths.subarray(0, hlit);
      const distLengths = allLengths.subarray(hlit, hlit + hdist);

      const litLenTree = buildHuffmanTree(litLenLengths, hlit);
      const distTree = buildHuffmanTree(distLengths, hdist);

      inflateBlock(litLenTree, distTree);
    } else {
      throw new Error(`Invalid DEFLATE block type: ${btype}`);
    }
  } while (bfinal === 0);

  return output.subarray(0, outPos);
}

// ── Deflate ─────────────────────────────────────────────────────────

/**
 * Compress data using raw DEFLATE (no zlib/gzip header).
 * Uses fixed Huffman codes with basic LZ77 matching.
 */
export function deflate(data: Uint8Array): Uint8Array {
  if (data.length === 0) {
    // Empty input: emit a single final stored block with length 0
    return new Uint8Array([0x03, 0x00]);
  }

  // For small data, use stored blocks
  if (data.length <= 64) {
    return deflateStored(data);
  }

  return deflateLZ77(data);
}

/** Emit a single final stored (uncompressed) block */
function deflateStored(data: Uint8Array): Uint8Array {
  // We can fit up to 65535 bytes per stored block.
  // For simplicity, handle data <= 65535 as one block.
  const blocks: Uint8Array[] = [];
  let offset = 0;

  while (offset < data.length) {
    const remaining = data.length - offset;
    const blockLen = Math.min(remaining, 65535);
    const isFinal = offset + blockLen >= data.length;

    const block = new Uint8Array(5 + blockLen);
    block[0] = isFinal ? 0x01 : 0x00; // BFINAL=1/0, BTYPE=00 (stored)
    block[1] = blockLen & 0xff;
    block[2] = (blockLen >> 8) & 0xff;
    block[3] = ~blockLen & 0xff;
    block[4] = (~blockLen >> 8) & 0xff;
    block.set(data.subarray(offset, offset + blockLen), 5);
    blocks.push(block);
    offset += blockLen;
  }

  const totalLen = blocks.reduce((sum, b) => sum + b.length, 0);
  const result = new Uint8Array(totalLen);
  let pos = 0;
  for (const block of blocks) {
    result.set(block, pos);
    pos += block.length;
  }
  return result;
}

/** Deflate with LZ77 + fixed Huffman codes */
function deflateLZ77(data: Uint8Array): Uint8Array {
  // Output bit buffer
  let outBuf = new Uint8Array(data.length + 512);
  let outPos = 0;
  let bitBuf = 0;
  let bitCount = 0;

  function ensureOut(needed: number): void {
    if (outPos + needed > outBuf.length) {
      let newSize = outBuf.length * 2;
      while (outPos + needed > newSize) {
        newSize *= 2;
      }
      const newArr = new Uint8Array(newSize);
      newArr.set(outBuf);
      outBuf = newArr;
    }
  }

  function writeBits(value: number, bits: number): void {
    bitBuf |= value << bitCount;
    bitCount += bits;
    while (bitCount >= 8) {
      ensureOut(1);
      outBuf[outPos++] = bitBuf & 0xff;
      bitBuf >>>= 8;
      bitCount -= 8;
    }
  }

  function flushBits(): void {
    if (bitCount > 0) {
      ensureOut(1);
      outBuf[outPos++] = bitBuf & 0xff;
      bitBuf = 0;
      bitCount = 0;
    }
  }

  /** Encode a literal/length code using fixed Huffman tables */
  function writeFixedLitLen(code: number): void {
    if (code <= 143) {
      // 8 bits: 00110000 + code (reversed)
      writeBits(reverseBits(0x30 + code, 8), 8);
    } else if (code <= 255) {
      // 9 bits: 110010000 + (code - 144) (reversed)
      writeBits(reverseBits(0x190 + (code - 144), 9), 9);
    } else if (code <= 279) {
      // 7 bits: 0000000 + (code - 256) (reversed)
      writeBits(reverseBits(code - 256, 7), 7);
    } else {
      // 8 bits: 11000000 + (code - 280) (reversed)
      writeBits(reverseBits(0xc0 + (code - 280), 8), 8);
    }
  }

  /** Encode a distance code using fixed Huffman (5 bits) */
  function writeFixedDist(code: number): void {
    writeBits(reverseBits(code, 5), 5);
  }

  // BFINAL=1, BTYPE=01 (fixed Huffman)
  writeBits(1, 1); // BFINAL
  writeBits(1, 2); // BTYPE = 01

  // Simple hash-chain LZ77
  const WINDOW = 32768;
  const MAX_MATCH = 258;
  const MIN_MATCH = 3;
  const HASH_SIZE = 1 << 15;
  const HASH_MASK = HASH_SIZE - 1;

  const head = new Int32Array(HASH_SIZE).fill(-1);
  const prev = new Int32Array(WINDOW);

  function hash3(pos: number): number {
    if (pos + 2 >= data.length) return 0;
    return ((data[pos] << 10) ^ (data[pos + 1] << 5) ^ data[pos + 2]) & HASH_MASK;
  }

  let i = 0;
  while (i < data.length) {
    if (i + MIN_MATCH > data.length) {
      // Not enough bytes for a match, emit literal
      writeFixedLitLen(data[i]);
      i++;
      continue;
    }

    const h = hash3(i);
    let bestLen = MIN_MATCH - 1;
    let bestDist = 0;

    // Search hash chain
    let chainPos = head[h];
    let chainLen = 0;
    const maxChain = 128;
    const minPos = Math.max(0, i - WINDOW);

    while (chainPos >= minPos && chainLen < maxChain) {
      const dist = i - chainPos;
      if (dist > 0 && dist <= WINDOW) {
        // Check match length
        let len = 0;
        const maxLen = Math.min(MAX_MATCH, data.length - i);
        while (len < maxLen && data[chainPos + len] === data[i + len]) {
          len++;
        }
        if (len > bestLen) {
          bestLen = len;
          bestDist = dist;
          if (len === MAX_MATCH) break;
        }
      }
      chainPos = prev[chainPos & (WINDOW - 1)];
      chainLen++;
    }

    // Update hash chain
    prev[i & (WINDOW - 1)] = head[h];
    head[h] = i;

    if (bestLen >= MIN_MATCH) {
      // Emit length/distance pair
      const lenCode = findLengthCode(bestLen);
      writeFixedLitLen(lenCode + 257);
      writeBits(bestLen - lengthBase[lenCode], lengthExtra[lenCode]);

      const distCode = findDistCode(bestDist);
      writeFixedDist(distCode);
      writeBits(bestDist - distBase[distCode], distExtra[distCode]);

      // Update hash chain for skipped positions
      for (let j = 1; j < bestLen; j++) {
        const pos = i + j;
        if (pos + MIN_MATCH <= data.length) {
          const h2 = hash3(pos);
          prev[pos & (WINDOW - 1)] = head[h2];
          head[h2] = pos;
        }
      }

      i += bestLen;
    } else {
      // Emit literal
      writeFixedLitLen(data[i]);
      i++;
    }
  }

  // End of block
  writeFixedLitLen(256);
  flushBits();

  return outBuf.subarray(0, outPos);
}

function reverseBits(value: number, bits: number): number {
  let result = 0;
  for (let i = 0; i < bits; i++) {
    result = (result << 1) | (value & 1);
    value >>= 1;
  }
  return result;
}

function findLengthCode(length: number): number {
  for (let i = lengthBase.length - 1; i >= 0; i--) {
    if (length >= lengthBase[i]) return i;
  }
  return 0;
}

function findDistCode(dist: number): number {
  for (let i = distBase.length - 1; i >= 0; i--) {
    if (dist >= distBase[i]) return i;
  }
  return 0;
}

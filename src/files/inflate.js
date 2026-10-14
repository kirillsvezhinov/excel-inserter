// Raw DEFLATE (RFC 1951) decoder. Used when the platform has no
// DecompressionStream('deflate-raw') (older Safari, some test runtimes).

const LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258]
const LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0]
const DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577]
const DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13]
const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]

/** Canonical Huffman table from code lengths: counts per length + symbols sorted by code. */
function buildTable(lengths) {
  const counts = new Uint16Array(16)
  for (const l of lengths) counts[l]++
  counts[0] = 0
  const offsets = new Uint16Array(16)
  for (let i = 1; i < 16; i++) offsets[i] = offsets[i - 1] + counts[i - 1]
  const symbols = new Uint16Array(lengths.length)
  lengths.forEach((l, sym) => { if (l) symbols[offsets[l]++] = sym })
  return { counts, symbols }
}

let FIXED = null
function fixedTables() {
  if (!FIXED) {
    const lit = new Uint8Array(288)
    lit.fill(8, 0, 144); lit.fill(9, 144, 256); lit.fill(7, 256, 280); lit.fill(8, 280, 288)
    FIXED = [buildTable(lit), buildTable(new Uint8Array(30).fill(5))]
  }
  return FIXED
}

export function inflateRaw(data, sizeHint = 0) {
  let pos = 0
  let bitBuf = 0
  let bitCnt = 0
  let out = new Uint8Array(Math.max(sizeHint, data.length * 4, 1024))
  let len = 0

  const ensure = (n) => {
    if (len + n <= out.length) return
    const next = new Uint8Array(Math.max(out.length * 2, len + n))
    next.set(out.subarray(0, len))
    out = next
  }
  const bits = (n) => {
    while (bitCnt < n) {
      if (pos >= data.length) throw new Error('inflate: unexpected end of data')
      bitBuf |= data[pos++] << bitCnt
      bitCnt += 8
    }
    const v = bitBuf & ((1 << n) - 1)
    bitBuf >>>= n
    bitCnt -= n
    return v
  }
  const decode = ({ counts, symbols }) => {
    let code = 0
    let first = 0
    let index = 0
    for (let l = 1; l < 16; l++) {
      code |= bits(1)
      const count = counts[l]
      if (code - first < count) return symbols[index + code - first]
      index += count
      first = (first + count) << 1
      code <<= 1
    }
    throw new Error('inflate: bad Huffman code')
  }

  let final = 0
  while (!final) {
    final = bits(1)
    const type = bits(2)
    if (type === 0) {
      bitBuf = 0
      bitCnt = 0
      if (pos + 4 > data.length) throw new Error('inflate: truncated stored block')
      const n = data[pos] | (data[pos + 1] << 8)
      pos += 4
      if (pos + n > data.length) throw new Error('inflate: truncated stored block')
      ensure(n)
      out.set(data.subarray(pos, pos + n), len)
      len += n
      pos += n
      continue
    }
    let lit, dist
    if (type === 1) [lit, dist] = fixedTables()
    else if (type === 2) {
      const hlit = bits(5) + 257
      const hdist = bits(5) + 1
      const hclen = bits(4) + 4
      const clLens = new Uint8Array(19)
      for (let i = 0; i < hclen; i++) clLens[CL_ORDER[i]] = bits(3)
      const clTable = buildTable(clLens)
      const lens = new Uint8Array(hlit + hdist)
      for (let i = 0; i < hlit + hdist;) {
        const sym = decode(clTable)
        if (sym < 16) lens[i++] = sym
        else {
          let repeat, value = 0
          if (sym === 16) {
            if (!i) throw new Error('inflate: repeat with no previous length')
            value = lens[i - 1]
            repeat = 3 + bits(2)
          } else if (sym === 17) repeat = 3 + bits(3)
          else repeat = 11 + bits(7)
          if (i + repeat > lens.length) throw new Error('inflate: too many code lengths')
          lens.fill(value, i, i + repeat)
          i += repeat
        }
      }
      lit = buildTable(lens.subarray(0, hlit))
      dist = buildTable(lens.subarray(hlit))
    } else throw new Error('inflate: invalid block type')

    for (;;) {
      const sym = decode(lit)
      if (sym < 256) { ensure(1); out[len++] = sym; continue }
      if (sym === 256) break
      const li = sym - 257
      if (li >= LEN_BASE.length) throw new Error('inflate: bad length symbol')
      const length = LEN_BASE[li] + bits(LEN_EXTRA[li])
      const di = decode(dist)
      if (di >= DIST_BASE.length) throw new Error('inflate: bad distance symbol')
      const distance = DIST_BASE[di] + bits(DIST_EXTRA[di])
      if (distance > len) throw new Error('inflate: distance too far back')
      ensure(length)
      for (let k = 0; k < length; k++, len++) out[len] = out[len - distance]
    }
  }
  return out.subarray(0, len)
}

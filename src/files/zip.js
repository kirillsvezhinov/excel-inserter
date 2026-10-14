import { inflateRaw } from './inflate.js'

const MAX_ENTRY_SIZE = 256 * 1024 * 1024 // refuse zip bombs

async function inflateNative(data) {
  const stream = new ReadableStream({ start(c) { c.enqueue(data); c.close() } })
    .pipeThrough(new DecompressionStream('deflate-raw'))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

let nativeOk = typeof DecompressionStream !== 'undefined' && typeof Response !== 'undefined'

async function inflate(data, size) {
  if (nativeOk) {
    try {
      return await inflateNative(data)
    } catch (err) {
      if (!(err instanceof TypeError) && !(err instanceof RangeError)) throw err
      nativeOk = false // 'deflate-raw' unsupported here
    }
  }
  return inflateRaw(data, size)
}

/**
 * Minimal ZIP reader (stored + deflate, no ZIP64/encryption) — enough for OOXML files.
 * @param {ArrayBuffer|Uint8Array} buffer
 * @returns {Map<string, () => Promise<Uint8Array>>} entry name → lazy reader
 */
export function readZip(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = -1
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 22 - 65535); i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break }
  }
  if (eocd < 0) throw new Error('Not a ZIP file')
  const count = view.getUint16(eocd + 10, true)
  let p = view.getUint32(eocd + 16, true)
  if (p === 0xffffffff || count === 0xffff) throw new Error('ZIP64 is not supported')

  const entries = new Map()
  const decoder = new TextDecoder()
  for (let i = 0; i < count; i++) {
    if (p + 46 > bytes.length || view.getUint32(p, true) !== 0x02014b50) throw new Error('Corrupt ZIP directory')
    const flags = view.getUint16(p + 8, true)
    const method = view.getUint16(p + 10, true)
    const compSize = view.getUint32(p + 20, true)
    const size = view.getUint32(p + 24, true)
    const nameLen = view.getUint16(p + 28, true)
    const extraLen = view.getUint16(p + 30, true)
    const commentLen = view.getUint16(p + 32, true)
    const local = view.getUint32(p + 42, true)
    const name = decoder.decode(bytes.subarray(p + 46, p + 46 + nameLen))
    p += 46 + nameLen + extraLen + commentLen

    entries.set(name.replace(/^\/+/, ''), async () => {
      if (flags & 1) throw new Error(`Encrypted ZIP entry: ${name}`)
      if (size > MAX_ENTRY_SIZE) throw new Error(`ZIP entry too large: ${name}`)
      if (view.getUint32(local, true) !== 0x04034b50) throw new Error(`Corrupt ZIP entry: ${name}`)
      const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true)
      const data = bytes.subarray(start, start + compSize)
      if (method === 0) return data
      if (method === 8) {
        const out = await inflate(data, size)
        if (out.length > MAX_ENTRY_SIZE) throw new Error(`ZIP entry too large: ${name}`)
        return out
      }
      throw new Error(`Unsupported ZIP compression method ${method}: ${name}`)
    })
  }
  return entries
}

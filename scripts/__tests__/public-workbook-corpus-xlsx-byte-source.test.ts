import { createHash } from 'node:crypto'
import { describe, expect, it } from 'vitest'

import { sha256XlsxZipByteSourceHex } from '../public-workbook-corpus-xlsx-byte-source.ts'

describe('public workbook corpus XLSX byte source helpers', () => {
  it('hashes byte sources through reusable readRangeInto scratch buffers', () => {
    const bytes = patternedBytes(150_000)
    const source = new InstrumentedByteSource(bytes)

    expect(sha256XlsxZipByteSourceHex(source)).toBe(sha256Hex(bytes))

    expect(source.readIntoCount).toBe(3)
    expect(source.rangeCount).toBe(0)
  })

  it('falls back to readRange when a byte source does not support readRangeInto', () => {
    const bytes = patternedBytes(140_000)
    const source = new RangeOnlyByteSource(bytes)

    expect(sha256XlsxZipByteSourceHex(source)).toBe(sha256Hex(bytes))

    expect(source.rangeCount).toBe(3)
  })
})

class InstrumentedByteSource {
  readonly byteLength: number
  rangeCount = 0
  readIntoCount = 0

  constructor(private readonly bytes: Uint8Array) {
    this.byteLength = bytes.byteLength
  }

  readRange(start: number, end: number): Uint8Array {
    this.rangeCount += 1
    return this.bytes.subarray(start, end)
  }

  readRangeInto(start: number, end: number, target: Uint8Array): Uint8Array {
    this.readIntoCount += 1
    target.set(this.bytes.subarray(start, end), 0)
    return target.subarray(0, end - start)
  }
}

class RangeOnlyByteSource {
  readonly byteLength: number
  rangeCount = 0

  constructor(private readonly bytes: Uint8Array) {
    this.byteLength = bytes.byteLength
  }

  readRange(start: number, end: number): Uint8Array {
    this.rangeCount += 1
    return this.bytes.subarray(start, end)
  }
}

function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

function patternedBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length)
  for (let index = 0; index < bytes.length; index += 1) {
    bytes[index] = index % 251
  }
  return bytes
}

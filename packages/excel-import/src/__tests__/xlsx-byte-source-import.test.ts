import { describe, expect, it } from 'vitest'

import { borrowXlsxZipByteSource } from '../xlsx-byte-source-import.js'

describe('XLSX byte-source import', () => {
  it('preserves reusable readRangeInto support on borrowed ZIP byte sources', () => {
    const source = new InstrumentedByteSource(new Uint8Array([1, 2, 3, 4]))
    const borrowed = borrowXlsxZipByteSource(source)
    const scratch = new Uint8Array(2)

    expect(Array.from(borrowed.readRangeInto?.(1, 3, scratch) ?? [])).toEqual([2, 3])

    expect(source.readIntoCount).toBe(1)
    expect(source.rangeCount).toBe(0)
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

import { describe, expect, it } from 'vitest'

import { borrowXlsxZipByteSource as borrowLargeSimpleVerifierXlsxZipByteSource } from '../public-workbook-corpus-large-simple-compact.ts'
import {
  buildVerificationWorkerProcessArgs,
  disableBunSmolVerificationWorkerEnvVar,
  shouldUseBunSmolForVerificationWorker,
} from '../public-workbook-corpus-verify-isolated.ts'

describe('public workbook corpus isolated verification worker runtime', () => {
  it('uses Bun smol mode for memory-sensitive isolated verification workers', () => {
    expect(shouldUseBunSmolForVerificationWorker({ versions: { bun: '1.3.0' }, env: {} })).toBe(true)
    expect(buildVerificationWorkerProcessArgs(['worker.ts', 'verify-artifact-worker'], { versions: { bun: '1.3.0' }, env: {} })).toEqual([
      '--smol',
      'worker.ts',
      'verify-artifact-worker',
    ])
  })

  it('does not add Bun runtime flags under Node or when explicitly disabled', () => {
    expect(shouldUseBunSmolForVerificationWorker({ versions: { node: '24.0.0' }, env: {} })).toBe(false)
    expect(
      buildVerificationWorkerProcessArgs(['worker.ts'], {
        versions: { bun: '1.3.0' },
        env: { [disableBunSmolVerificationWorkerEnvVar]: 'true' },
      }),
    ).toEqual(['worker.ts'])
  })

  it('keeps readRangeInto support on compact verifier borrowed byte sources', () => {
    const source = new InstrumentedByteSource(new Uint8Array([1, 2, 3, 4]))
    const borrowed = borrowLargeSimpleVerifierXlsxZipByteSource(source)
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

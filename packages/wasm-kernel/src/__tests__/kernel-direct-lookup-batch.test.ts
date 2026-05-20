import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

describe('wasm kernel direct lookup batch', () => {
  it('evaluates exact and approximate uniform numeric lookup batches', async () => {
    const kernel = await createKernel()
    const outTags = new Uint8Array(5)
    const outNumbers = new Float64Array(5)
    const outErrors = new Uint16Array(5)

    kernel.evalUniformNumericLookupBatch(
      Uint8Array.from([1, 1, 2, 2, 2]),
      Uint8Array.from([1, 1, 1, 2, 1]),
      Float64Array.from([2, 2, 1, 10, 1]),
      Float64Array.from([2, 2, 1, -1, 1]),
      Uint32Array.from([8, 8, 10, 10, 10]),
      Uint32Array.from([0, 0, 0, 0, 2]),
      Uint8Array.from([ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number]),
      Float64Array.from([8, 9, 4.5, 7.5, 3.5]),
      outTags,
      outNumbers,
      outErrors,
    )

    expect([...outTags]).toEqual([ValueTag.Number, ValueTag.Error, ValueTag.Number, ValueTag.Number, ValueTag.Number])
    expect(outNumbers[0]).toBe(4)
    expect(outErrors[1]).toBe(ErrorCode.NA)
    expect(outNumbers[2]).toBe(4)
    expect(outNumbers[3]).toBe(3)
    expect(outNumbers[4]).toBe(6)
  })
})

import { describe, expect, it } from 'vitest'
import { createKernel } from '../index.js'

describe('wasm kernel direct aggregate batch', () => {
  it('evaluates dense numeric row aggregate batches', async () => {
    const kernel = await createKernel()
    const outNumbers = new Float64Array(3)

    kernel.evalDenseNumericRowAggregateBatch(1, Float64Array.from([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]), 3, 4, 1, 2, 10, outNumbers)

    expect([...outNumbers]).toEqual([15, 23, 31])
  })

  it('evaluates anchored prefix aggregate batches with mixed cell tags', async () => {
    const kernel = await createKernel()
    const outTags = new Uint8Array(4)
    const outNumbers = new Float64Array(4)
    const outErrors = new Uint16Array(4)

    kernel.evalAnchoredPrefixAggregateBatch(
      1,
      Uint8Array.from([1, 2, 0, 3, 4]),
      Float64Array.from([5, 1, 0, 0, 0]),
      Uint16Array.from([0, 0, 0, 0, 3]),
      5,
      1,
      Uint32Array.from([0, 1, 2, 4]),
      Float64Array.from([0, 10, 0, 0]),
      outTags,
      outNumbers,
      outErrors,
    )

    expect([...outTags]).toEqual([1, 1, 1, 4])
    expect([...outNumbers]).toEqual([5, 16, 6, 0])
    expect([...outErrors]).toEqual([0, 0, 0, 3])
  })
})

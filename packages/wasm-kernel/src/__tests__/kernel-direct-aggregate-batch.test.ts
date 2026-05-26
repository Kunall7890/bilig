import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

const DIRECT_AGGREGATE_OP_SUM = 1
const DIRECT_AGGREGATE_OP_AVERAGE = 2
const DIRECT_AGGREGATE_OP_COUNT = 3
const DIRECT_AGGREGATE_OP_MIN = 4
const DIRECT_AGGREGATE_OP_MAX = 5

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
      DIRECT_AGGREGATE_OP_SUM,
      Uint8Array.from([ValueTag.Number, ValueTag.Boolean, ValueTag.Empty, ValueTag.String, ValueTag.Error]),
      Float64Array.from([5, 1, 0, 0, 0]),
      Uint16Array.from([0, 0, 0, 0, ErrorCode.Value]),
      5,
      1,
      Uint32Array.from([0, 1, 2, 4]),
      Float64Array.from([0, 10, 0, 0]),
      outTags,
      outNumbers,
      outErrors,
    )

    expect([...outTags]).toEqual([ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Error])
    expect([...outNumbers]).toEqual([5, 15, 5, 0])
    expect([...outErrors]).toEqual([0, 0, 0, ErrorCode.Value])
  })

  it('ignores non-numeric prefix references for every direct aggregate kind', async () => {
    const kernel = await createKernel()
    const tags = Uint8Array.from([ValueTag.Number, ValueTag.Boolean, ValueTag.Empty, ValueTag.Number, ValueTag.String])
    const numbers = Float64Array.from([5, 1, 0, -2, 0])
    const errors = new Uint16Array(5)
    const formulaRowEnds = Uint32Array.from([0, 1, 2, 3, 4])
    const resultOffsets = new Float64Array(5)

    const cases = [
      { aggregateKind: DIRECT_AGGREGATE_OP_SUM, expected: [5, 5, 5, 3, 3] },
      { aggregateKind: DIRECT_AGGREGATE_OP_AVERAGE, expected: [5, 5, 5, 1.5, 1.5] },
      { aggregateKind: DIRECT_AGGREGATE_OP_COUNT, expected: [1, 1, 1, 2, 2] },
      { aggregateKind: DIRECT_AGGREGATE_OP_MIN, expected: [5, 5, 5, -2, -2] },
      { aggregateKind: DIRECT_AGGREGATE_OP_MAX, expected: [5, 5, 5, 5, 5] },
    ]

    for (const { aggregateKind, expected } of cases) {
      const outTags = new Uint8Array(5)
      const outNumbers = new Float64Array(5)
      const outErrors = new Uint16Array(5)

      kernel.evalAnchoredPrefixAggregateBatch(
        aggregateKind,
        tags,
        numbers,
        errors,
        5,
        1,
        formulaRowEnds,
        resultOffsets,
        outTags,
        outNumbers,
        outErrors,
      )

      expect([...outTags]).toEqual([ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number])
      expect([...outNumbers]).toEqual(expected)
      expect([...outErrors]).toEqual([0, 0, 0, 0, 0])
    }
  })
})

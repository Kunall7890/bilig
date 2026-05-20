import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

const DIRECT_AGGREGATE_OP_SUM = 1
const DIRECT_AGGREGATE_OP_AVERAGE = 2
const DIRECT_AGGREGATE_OP_COUNT = 3
const DIRECT_AGGREGATE_OP_MIN = 4
const DIRECT_AGGREGATE_OP_MAX = 5

describe('wasm kernel direct criteria aggregate batch', () => {
  it('reduces matched row offsets with criteria aggregate semantics', async () => {
    const kernel = await createKernel()
    const outTags = new Uint8Array(7)
    const outNumbers = new Float64Array(7)
    const outErrors = new Uint16Array(7)

    kernel.evalDirectCriteriaMatchedAggregateBatch(
      Uint8Array.from([
        DIRECT_AGGREGATE_OP_SUM,
        DIRECT_AGGREGATE_OP_AVERAGE,
        DIRECT_AGGREGATE_OP_MIN,
        DIRECT_AGGREGATE_OP_MAX,
        DIRECT_AGGREGATE_OP_COUNT,
        DIRECT_AGGREGATE_OP_SUM,
        DIRECT_AGGREGATE_OP_AVERAGE,
      ]),
      Uint32Array.from([0, 0, 0, 0, 0, 4, 3]),
      Uint32Array.from([4, 4, 4, 4, 4, 3, 1]),
      Uint32Array.from([0, 1, 2, 3, 4, 5, 0, 4]),
      Uint8Array.from([ValueTag.Number, ValueTag.Boolean, ValueTag.Empty, ValueTag.String, ValueTag.Error, ValueTag.Number]),
      Float64Array.from([5, 1, 0, 0, 0, 9]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.Ref, ErrorCode.None]),
      outTags,
      outNumbers,
      outErrors,
    )

    expect([...outTags]).toEqual([
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Error,
      ValueTag.Error,
    ])
    expect(outNumbers[0]).toBe(6)
    expect(outNumbers[1]).toBe(2)
    expect(outNumbers[2]).toBe(5)
    expect(outNumbers[3]).toBe(5)
    expect(outNumbers[4]).toBe(4)
    expect(outErrors[5]).toBe(ErrorCode.Ref)
    expect(outErrors[6]).toBe(ErrorCode.Div0)
  })
})

import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

describe('wasm kernel direct scalar batch', () => {
  it('evaluates value batches with dependencies on earlier batch outputs', async () => {
    const kernel = await createKernel()
    const outTags = new Uint8Array(3)
    const outNumbers = new Float64Array(3)
    const outErrors = new Uint16Array(3)

    kernel.evalDirectScalarValueBatch(
      Uint8Array.from([1, 3, 1]),
      Uint32Array.from([0xffffffff, 0, 0xffffffff]),
      Uint8Array.from([ValueTag.Number, ValueTag.Empty, ValueTag.String]),
      Float64Array.from([2, 0, 0]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None]),
      Uint32Array.from([0xffffffff, 0xffffffff, 0xffffffff]),
      Uint8Array.from([ValueTag.Number, ValueTag.Number, ValueTag.Number]),
      Float64Array.from([3, 2, 1]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None]),
      Float64Array.from([0, 5, 0]),
      outTags,
      outNumbers,
      outErrors,
    )

    expect(outTags[0]).toBe(ValueTag.Number)
    expect(outNumbers[0]).toBe(5)
    expect(outTags[1]).toBe(ValueTag.Number)
    expect(outNumbers[1]).toBe(15)
    expect(outTags[2]).toBe(ValueTag.Error)
    expect(outErrors[2]).toBe(ErrorCode.Value)
  })

  it('writes store-target batches into resident cell arrays', async () => {
    const kernel = await createKernel()
    kernel.init(8, 1, 1, 1, 1)

    kernel.evalDirectScalarStoreTargetBatch(
      Uint32Array.from([2, 3, 4]),
      Uint8Array.from([1, 3, 1]),
      Uint32Array.from([0xffffffff, 0, 0xffffffff]),
      Uint8Array.from([ValueTag.Number, ValueTag.Empty, ValueTag.String]),
      Float64Array.from([2, 0, 0]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None]),
      Uint32Array.from([0xffffffff, 0xffffffff, 0xffffffff]),
      Uint8Array.from([ValueTag.Number, ValueTag.Number, ValueTag.Number]),
      Float64Array.from([3, 2, 1]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None]),
      Float64Array.from([0, 5, 0]),
    )

    expect(kernel.readTags()[2]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[2]).toBe(5)
    expect(kernel.readTags()[3]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[3]).toBe(15)
    expect(kernel.readTags()[4]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[4]).toBe(ErrorCode.Value)
    expect(kernel.readStringIds()[4]).toBe(0)
  })

  it('evaluates direct conditional branch-pick batches', async () => {
    const kernel = await createKernel()
    const outTags = new Uint8Array(4)
    const outNumbers = new Float64Array(4)
    const outStringIds = new Uint32Array(4)
    const outErrors = new Uint16Array(4)

    kernel.evalDirectConditionalPickBatch(
      Uint32Array.from([0, 2, 3, 3]),
      Uint32Array.from([2, 1, 0, 1]),
      Uint8Array.from([2, 2, 4, 4]),
      Uint8Array.from([ValueTag.String, ValueTag.String, ValueTag.Number, ValueTag.String]),
      Float64Array.from([0, 0, 8, 0]),
      Uint32Array.from([1, 2, 0, 4]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None]),
      Uint8Array.from([ValueTag.String, ValueTag.String, ValueTag.Number, ValueTag.String]),
      Float64Array.from([0, 0, 5, 0]),
      Uint32Array.from([3, 2, 0, 5]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None]),
      Uint8Array.from([ValueTag.Number, ValueTag.String, ValueTag.Boolean, ValueTag.Number]),
      Float64Array.from([1, 0, 1, 99]),
      Uint32Array.from([0, 7, 0, 0]),
      Uint16Array.from([ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None]),
      Uint8Array.from([ValueTag.Error, ValueTag.String, ValueTag.Error, ValueTag.Number]),
      Float64Array.from([0, 0, 0, 0]),
      Uint32Array.from([0, 9, 0, 0]),
      Uint16Array.from([ErrorCode.NA, ErrorCode.None, ErrorCode.NA, ErrorCode.None]),
      outTags,
      outNumbers,
      outStringIds,
      outErrors,
    )

    expect([...outTags]).toEqual([ValueTag.String, ValueTag.Boolean, ValueTag.Error, ValueTag.Error])
    expect([...outNumbers]).toEqual([0, 1, 0, 0])
    expect([...outStringIds]).toEqual([7, 0, 0, 0])
    expect([...outErrors]).toEqual([ErrorCode.None, ErrorCode.None, ErrorCode.NA, ErrorCode.Value])
  })

  it('writes dense two-formula row chains into resident cell arrays', async () => {
    const kernel = await createKernel()
    kernel.init(8, 1, 1, 1, 1)

    kernel.evalDenseDirectScalarRowChainStoreTargetBatch(
      Float64Array.from([2, 5, 7]),
      Float64Array.from([3, 4, 8]),
      Uint32Array.from([0, 2, 4]),
      Uint32Array.from([1, 3, 5]),
      3,
      1,
      2,
      1,
    )

    expect(Array.from(kernel.readTags().slice(0, 6))).toEqual(Array(6).fill(ValueTag.Number))
    expect(Array.from(kernel.readNumbers().slice(0, 6))).toEqual([5, 11, 9, 19, 15, 31])
    expect(Array.from(kernel.readStringIds().slice(0, 6))).toEqual(Array(6).fill(0))
    expect(Array.from(kernel.readErrors().slice(0, 6))).toEqual(Array(6).fill(ErrorCode.None))
  })

  it('supports division in dense two-formula row chains', async () => {
    const kernel = await createKernel()
    kernel.init(4, 1, 1, 1, 1)

    kernel.evalDenseDirectScalarRowChainStoreTargetBatch(
      Float64Array.from([450, 84]),
      Float64Array.from([16, 7]),
      Uint32Array.from([0, 2]),
      Uint32Array.from([1, 3]),
      2,
      5,
      6,
      0,
    )

    expect(Array.from(kernel.readTags().slice(0, 4))).toEqual(Array(4).fill(ValueTag.Number))
    expect(Array.from(kernel.readNumbers().slice(0, 4))).toEqual([28.125, 168.75, 12, 72])
    expect(Array.from(kernel.readErrors().slice(0, 4))).toEqual(Array(4).fill(ErrorCode.None))
  })

  it('writes dense row chains whose second formula divides by a row denominator', async () => {
    const kernel = await createKernel()
    kernel.init(6, 1, 1, 1, 1)

    kernel.evalDenseDirectScalarRowChainDivideStoreTargetBatch(
      Float64Array.from([900, 120]),
      Float64Array.from([300, 30]),
      Float64Array.from([4, 5]),
      Uint32Array.from([0, 2]),
      Uint32Array.from([1, 3]),
      2,
      1,
    )

    expect(Array.from(kernel.readTags().slice(0, 4))).toEqual(Array(4).fill(ValueTag.Number))
    expect(Array.from(kernel.readNumbers().slice(0, 4))).toEqual([1200, 300, 150, 30])
    expect(Array.from(kernel.readStringIds().slice(0, 4))).toEqual(Array(4).fill(0))
    expect(Array.from(kernel.readErrors().slice(0, 4))).toEqual(Array(4).fill(ErrorCode.None))
  })
})

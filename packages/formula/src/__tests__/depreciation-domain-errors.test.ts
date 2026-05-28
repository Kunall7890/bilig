import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const numError = { tag: ValueTag.Error, code: ErrorCode.Num } as const
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const
const div0Error = { tag: ValueTag.Error, code: ErrorCode.Div0 } as const

function num(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function text(value: string, stringId = 1): CellValue {
  return { tag: ValueTag.String, value, stringId }
}

function err(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

describe('depreciation formula domain errors', () => {
  it('returns documented depreciation domain errors for DB, DDB, SLN, and SYD', () => {
    const DB = getBuiltin('DB')!
    const DDB = getBuiltin('DDB')!
    const SLN = getBuiltin('SLN')!
    const SYD = getBuiltin('SYD')!

    expect(DB(num(10000), num(1000), num(5), num(6))).toEqual(numError)
    expect(DB(num(10000), num(1000), num(5), num(6), num(12))).toEqual(numError)
    expect(DB(num(10000), num(1000), num(5), num(6), num(6))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(238.5271245878818, 12),
    })
    expect(DB(num(10000), num(1000), num(5), num(1), num(0))).toEqual(numError)
    expect(DB(num(10000), num(1000), num(5), num(1), num(13))).toEqual(numError)
    expect(DB(text('bad'), num(1000), num(5), num(1), num(12))).toEqual(valueError)

    expect(DDB(num(2400), num(300), num(10), num(11))).toEqual(numError)

    expect(SLN(num(10000), num(1000), num(0))).toEqual(div0Error)
    expect(SLN(text('bad'), num(1000), num(9))).toEqual(valueError)

    expect(SYD(num(10000), num(1000), num(9), num(10))).toEqual(numError)
    expect(SYD(num(10000), num(1000), num(0), num(1))).toEqual(numError)
    expect(SYD(num(10000), num(1000), num(9), num(0))).toEqual(numError)
    expect(SYD(text('bad'), num(1000), num(9), num(1))).toEqual(valueError)
  })

  it('returns #NUM! for DDB and VDB numeric-domain errors while preserving #VALUE! coercion errors', () => {
    const DDB = getBuiltin('DDB')!
    const VDB = getBuiltin('VDB')!

    expect(DDB(num(2400), num(300), num(10), num(2), num(0))).toEqual(numError)
    expect(DDB(num(2400), num(300), num(0), num(2), num(2))).toEqual(numError)
    expect(DDB(num(2400), num(300), num(10), num(0), num(2))).toEqual(numError)
    expect(DDB(text('bad'), num(300), num(10), num(2), num(2))).toEqual(valueError)

    expect(VDB(num(2400), num(300), num(10), num(3), num(1), num(2))).toEqual(numError)
    expect(VDB(num(2400), num(300), num(0), num(1), num(3), num(2))).toEqual(numError)
    expect(VDB(num(2400), num(300), num(10), num(1), num(3), num(0))).toEqual(numError)
    expect(VDB(text('bad', 2), num(300), num(10), num(1), num(3), num(2))).toEqual(valueError)
    expect(VDB(num(2400), num(300), num(10), num(1), num(3), num(2), text('bad', 3))).toEqual(valueError)
  })

  it('preserves incoming depreciation errors before scalar coercion', () => {
    const DB = getBuiltin('DB')!
    const DDB = getBuiltin('DDB')!
    const VDB = getBuiltin('VDB')!

    expect(DB(num(2400), num(300), num(10), num(2), err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(DDB(err(ErrorCode.Ref), num(300), num(10), num(2), num(2))).toEqual(err(ErrorCode.Ref))
    expect(VDB(num(2400), num(300), num(10), num(1), num(3), num(2), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
  })
})

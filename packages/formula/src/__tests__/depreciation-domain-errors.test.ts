import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const numError = { tag: ValueTag.Error, code: ErrorCode.Num } as const
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const

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

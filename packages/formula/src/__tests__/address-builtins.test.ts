import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const stringResult = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const

describe('ADDRESS builtins', () => {
  it('uses the documented logical A1 flag for A1 and R1C1 styles', () => {
    const ADDRESS = getBuiltin('ADDRESS')!

    expect(ADDRESS(num(2), num(3), num(2), bool(false))).toEqual(stringResult('R2C[3]'))
    expect(ADDRESS(num(2), num(3), num(2), num(0))).toEqual(stringResult('R2C[3]'))
    expect(ADDRESS(num(2), num(3), num(2), bool(true))).toEqual(stringResult('C$2'))
    expect(ADDRESS(num(2), num(3), num(2), num(2))).toEqual(stringResult('C$2'))
    expect(ADDRESS(num(2), num(3), num(2), text('not logical'))).toEqual(valueError)
  })
})

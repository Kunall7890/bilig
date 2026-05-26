import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string, stringId = 1): CellValue => ({ tag: ValueTag.String, value, stringId })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })

const numError = { tag: ValueTag.Error, code: ErrorCode.Num } as const
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const

describe('statistical builtins', () => {
  it('matches Microsoft-documented normal-family numeric-domain errors', () => {
    expect(getBuiltin('STANDARDIZE')?.(num(42), num(40), num(0))).toEqual(numError)
    expect(getBuiltin('NORMDIST')?.(num(1), num(0), num(0), bool(true))).toEqual(numError)
    expect(getBuiltin('NORM.DIST')?.(num(1), num(0), num(0), bool(false))).toEqual(numError)
    expect(getBuiltin('NORMINV')?.(num(0), num(0), num(1))).toEqual(numError)
    expect(getBuiltin('NORM.INV')?.(num(1), num(0), num(1))).toEqual(numError)
    expect(getBuiltin('NORMINV')?.(num(0.5), num(0), num(0))).toEqual(numError)
    expect(getBuiltin('NORMSINV')?.(num(0))).toEqual(numError)
    expect(getBuiltin('NORM.S.INV')?.(num(1))).toEqual(numError)

    expect(getBuiltin('NORMDIST')?.(num(1), text('bad'), num(1), bool(true))).toEqual(valueError)
    expect(getBuiltin('NORMINV')?.(text('bad'), num(0), num(1))).toEqual(valueError)
    expect(getBuiltin('NORMSINV')?.(text('bad'))).toEqual(valueError)
  })

  it('matches Microsoft-documented lognormal numeric-domain errors', () => {
    expect(getBuiltin('LOGINV')?.(num(0), num(0), num(1))).toEqual(numError)
    expect(getBuiltin('LOGINV')?.(num(0.5), num(0), num(0))).toEqual(numError)
    expect(getBuiltin('LOGNORM.INV')?.(num(1), num(0), num(1))).toEqual(numError)
    expect(getBuiltin('LOGNORMDIST')?.(num(0), num(0), num(1))).toEqual(numError)
    expect(getBuiltin('LOGNORM.DIST')?.(num(1), num(0), num(0), bool(true))).toEqual(numError)

    expect(getBuiltin('LOGINV')?.(text('bad'), num(0), num(1))).toEqual(valueError)
    expect(getBuiltin('LOGNORMDIST')?.(text('bad'), num(0), num(1))).toEqual(valueError)
    expect(getBuiltin('LOGNORM.DIST')?.(num(1), num(0), num(1), text('bad'))).toEqual(valueError)
  })
})

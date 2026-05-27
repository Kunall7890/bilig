import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string, stringId = 1): CellValue => ({ tag: ValueTag.String, value, stringId })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })

const numError = { tag: ValueTag.Error, code: ErrorCode.Num } as const
const naError = { tag: ValueTag.Error, code: ErrorCode.NA } as const
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const
const div0Error = { tag: ValueTag.Error, code: ErrorCode.Div0 } as const

describe('statistical builtins', () => {
  it('coerces direct numeric text for scalar normal-family functions', () => {
    expect(getBuiltin('GAUSS')?.(text('0'))).toEqual(num(0))
    expect(getBuiltin('PHI')?.(text('0'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(0.3989422804014327, 12),
    })
    expect(getBuiltin('STANDARDIZE')?.(text('2'), text('1'), text('2'))).toEqual(num(0.5))
    expect(getBuiltin('NORMINV')?.(text('0.5'), text('0'), text('1'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 8),
    })
    expect(getBuiltin('NORM.S.INV')?.(text('0.5'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 8),
    })
  })

  it('coerces direct numeric and logical text for normal and lognormal distribution forms', () => {
    expect(getBuiltin('NORMDIST')?.(text('1'), text('0'), text('1'), text('TRUE'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 8),
    })
    expect(getBuiltin('NORM.DIST')?.(text('1'), text('0'), text('1'), text('TRUE'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 8),
    })
    expect(getBuiltin('NORM.S.DIST')?.(text('1'), text('FALSE'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(0.24197072451914337, 8),
    })
    expect(getBuiltin('LOGINV')?.(text('0.5'), text('0'), text('1'))).toEqual(num(1))
    expect(getBuiltin('LOGNORM.DIST')?.(text('1'), text('0'), text('1'), text('TRUE'))).toEqual(num(0.5))
  })

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

  it('preserves incoming errors before statistical distribution coercion', () => {
    const cases = [
      getBuiltin('GAUSS')?.(naError),
      getBuiltin('PHI')?.(naError),
      getBuiltin('STANDARDIZE')?.(num(4), naError, num(2)),
      getBuiltin('NORMINV')?.(num(0.5), num(0), naError),
      getBuiltin('NORM.INV')?.(naError, num(0), num(1)),
      getBuiltin('NORMSDIST')?.(naError),
      getBuiltin('NORMSINV')?.(naError),
      getBuiltin('NORM.S.INV')?.(naError),
      getBuiltin('LOGINV')?.(num(0.5), naError, num(1)),
      getBuiltin('LOGNORM.INV')?.(num(0.5), num(0), naError),
      getBuiltin('LOGNORMDIST')?.(num(1), naError, num(1)),
    ]

    for (const actual of cases) {
      expect(actual).toEqual(naError)
    }
  })

  it('matches Microsoft-documented SKEW, SKEW.P, and KURT invalid-dispersion errors', () => {
    expect(getBuiltin('SKEW')?.(num(1), num(2))).toEqual(div0Error)
    expect(getBuiltin('SKEW')?.(num(1), num(1), num(1))).toEqual(div0Error)
    expect(getBuiltin('SKEW.P')?.(num(1), num(2))).toEqual(div0Error)
    expect(getBuiltin('SKEW.P')?.(num(1), num(1), num(1))).toEqual(div0Error)
    expect(getBuiltin('KURT')?.(num(1), num(2), num(3))).toEqual(div0Error)
    expect(getBuiltin('KURT')?.(num(1), num(1), num(1), num(1))).toEqual(div0Error)
  })
})

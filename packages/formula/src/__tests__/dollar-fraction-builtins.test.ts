import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const str = (value: string, stringId = 1): CellValue => ({ tag: ValueTag.String, value, stringId })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

describe('dollar fraction builtins', () => {
  it('matches Microsoft DOLLARDE and DOLLARFR denominator semantics', () => {
    const DOLLARDE = getBuiltin('DOLLARDE')!
    const DOLLARFR = getBuiltin('DOLLARFR')!

    expect(DOLLARDE(num(1.06), num(12))).toEqual(num(1.5))
    expect(DOLLARFR(num(1.5), num(12))).toEqual(num(1.06))

    expect(DOLLARDE(num(1.02), num(12.9))).toEqual(num(1 + 2 / 12))
    expect(DOLLARFR(num(1.5), num(12.9))).toEqual(num(1.06))

    expect(DOLLARDE(num(1.5), num(-1))).toEqual(err(ErrorCode.Num))
    expect(DOLLARFR(num(1.5), num(-1))).toEqual(err(ErrorCode.Num))
    expect(DOLLARDE(num(1.5), num(-0.5))).toEqual(err(ErrorCode.Num))
    expect(DOLLARFR(num(1.5), num(-0.5))).toEqual(err(ErrorCode.Num))

    expect(DOLLARDE(num(1.5), num(0))).toEqual(err(ErrorCode.Div0))
    expect(DOLLARFR(num(1.5), num(0))).toEqual(err(ErrorCode.Div0))
    expect(DOLLARDE(num(1.5), num(0.5))).toEqual(err(ErrorCode.Div0))
    expect(DOLLARFR(num(1.5), num(0.5))).toEqual(err(ErrorCode.Div0))

    expect(DOLLARDE(num(1.5), str('bad'))).toEqual(err(ErrorCode.Value))
    expect(DOLLARFR(num(1.5), str('bad'))).toEqual(err(ErrorCode.Value))
  })
})

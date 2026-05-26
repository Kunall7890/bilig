import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const str = (value: string, stringId = 1): CellValue => ({ tag: ValueTag.String, value, stringId })
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const
const div0Error = { tag: ValueTag.Error, code: ErrorCode.Div0 } as const
const numError = { tag: ValueTag.Error, code: ErrorCode.Num } as const

describe('math builtins', () => {
  it('rejects invalid scalar coercions instead of defaulting to zero or base 10', () => {
    expect(getBuiltin('ACOT')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('ACOTH')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('COT')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('SEC')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('LOG')?.(num(100), str('bad'))).toEqual(valueError)
  })

  it('keeps representative rounding, combinatoric, and bitwise behavior intact', () => {
    expect(getBuiltin('ROUNDUP')?.(num(12.341), num(2))).toEqual(num(12.35))
    expect(getBuiltin('COMBINA')?.(num(4), num(3))).toEqual(num(20))
    expect(getBuiltin('MROUND')?.(num(10), num(3))).toEqual(num(9))
    expect(getBuiltin('BITLSHIFT')?.(num(3), num(2))).toEqual(num(12))
  })

  it('matches Excel scalar math text coercion and overflow errors', () => {
    expect(getBuiltin('SIN')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('COS')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('POWER')?.(str('bad'), num(2))).toEqual(valueError)

    expect(getBuiltin('SIN')?.(str('1'))).toEqual(num(Math.sin(1)))
    expect(getBuiltin('POWER')?.(str('2'), str('3'))).toEqual(num(8))
    expect(getBuiltin('EXP')?.(str(''))).toEqual(num(1))

    expect(getBuiltin('EXP')?.(num(1000))).toEqual(numError)
    expect(getBuiltin('POWER')?.(num(10), num(400))).toEqual(numError)
    expect(getBuiltin('SINH')?.(num(1000))).toEqual(numError)
    expect(getBuiltin('COSH')?.(num(1000))).toEqual(numError)
  })

  it('coerces direct numeric text across scalar math builtins', () => {
    expect(getBuiltin('ABS')?.(str('-2'))).toEqual(num(2))
    expect(getBuiltin('ROUND')?.(str('2.5'), num(0))).toEqual(num(3))
    expect(getBuiltin('FLOOR')?.(str('5.5'), str('2'))).toEqual(num(4))
    expect(getBuiltin('CEILING')?.(str('5.5'), str('2'))).toEqual(num(6))
    expect(getBuiltin('MOD')?.(str('7'), str('2'))).toEqual(num(1))
    expect(getBuiltin('QUOTIENT')?.(str('7'), str('2'))).toEqual(num(3))
    expect(getBuiltin('MROUND')?.(str('10'), str('3'))).toEqual(num(9))
    expect(getBuiltin('LN')?.(str('2'))).toEqual(num(Math.log(2)))
    expect(getBuiltin('LOG')?.(str('100'), str('10'))).toEqual(num(2))
    expect(getBuiltin('SQRT')?.(str('4'))).toEqual(num(2))
    expect(getBuiltin('ROUNDUP')?.(str('12.34'), num(1))).toEqual(num(12.4))
    expect(getBuiltin('ROUNDDOWN')?.(str('12.34'), num(1))).toEqual(num(12.3))
    expect(getBuiltin('TRUNC')?.(str('12.34'), num(1))).toEqual(num(12.3))
    expect(getBuiltin('EVEN')?.(str('3'))).toEqual(num(4))
    expect(getBuiltin('ODD')?.(str('3'))).toEqual(num(3))
    expect(getBuiltin('FACT')?.(str('5'))).toEqual(num(120))
    expect(getBuiltin('COMBIN')?.(str('5'), str('2'))).toEqual(num(10))
    expect(getBuiltin('COMBINA')?.(str('4'), str('3'))).toEqual(num(20))
    expect(getBuiltin('PERMUT')?.(str('5'), str('2'))).toEqual(num(20))
    expect(getBuiltin('GCD')?.(str('18'), str('24'))).toEqual(num(6))
    expect(getBuiltin('LCM')?.(str('6'), str('8'))).toEqual(num(24))
  })

  it('matches Microsoft Excel GCD and LCM domain errors', () => {
    expect(getBuiltin('GCD')?.(num(-18), num(24))).toEqual(numError)
    expect(getBuiltin('GCD')?.(num(2 ** 53), num(2))).toEqual(numError)
    expect(getBuiltin('GCD')?.(num(18), str('bad'))).toEqual(valueError)

    expect(getBuiltin('LCM')?.(num(6), num(-8))).toEqual(numError)
    expect(getBuiltin('LCM')?.(num(2 ** 52), num(3))).toEqual(numError)
    expect(getBuiltin('LCM')?.(num(6), str('bad'))).toEqual(valueError)
  })

  it('matches Microsoft Excel MROUND sign-domain errors', () => {
    expect(getBuiltin('MROUND')?.(num(-10), num(3))).toEqual(numError)
    expect(getBuiltin('MROUND')?.(num(5), num(-2))).toEqual(numError)
  })

  it('matches Desktop Excel MOD sign semantics for negative operands', () => {
    expect(getBuiltin('MOD')?.(num(-3), num(2))).toEqual(num(1))
    expect(getBuiltin('MOD')?.(num(3), num(-2))).toEqual(num(-1))
    expect(getBuiltin('MOD')?.(num(-3), num(-2))).toEqual(num(-1))
  })

  it('matches Microsoft Excel ATAN2 coordinate argument order and zero-origin error', () => {
    expect(getBuiltin('ATAN2')?.(num(-1), num(1))).toEqual(num((3 * Math.PI) / 4))
    expect(getBuiltin('ATAN2')?.(num(1), num(-1))).toEqual(num(-Math.PI / 4))
    expect(getBuiltin('ATAN2')?.(num(0), num(0))).toEqual(div0Error)
  })

  it('matches Microsoft Excel FLOOR positive-number negative-significance error semantics', () => {
    expect(getBuiltin('FLOOR')?.(num(2.5), num(-2))).toEqual(numError)
    expect(getBuiltin('FLOOR')?.(num(-2.5), num(2))).toEqual(num(-4))
    expect(getBuiltin('FLOOR')?.(num(-2.5), num(-2))).toEqual(num(-2))
  })

  it('matches Microsoft Excel square-root negative-domain errors', () => {
    expect(getBuiltin('SQRT')?.(num(-1))).toEqual(numError)
    expect(getBuiltin('SQRTPI')?.(num(-1))).toEqual(numError)
    expect(getBuiltin('SQRT')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('SQRTPI')?.(str('bad'))).toEqual(valueError)
  })

  it('matches Microsoft Excel inverse trigonometric numeric-domain errors', () => {
    expect(getBuiltin('ASIN')?.(num(2))).toEqual(numError)
    expect(getBuiltin('ACOS')?.(num(2))).toEqual(numError)
    expect(getBuiltin('ACOSH')?.(num(0.5))).toEqual(numError)
    expect(getBuiltin('ATANH')?.(num(1))).toEqual(numError)
    expect(getBuiltin('ATANH')?.(num(-1))).toEqual(numError)
    expect(getBuiltin('ACOTH')?.(num(0.5))).toEqual(numError)
    expect(getBuiltin('ACOTH')?.(num(-0.5))).toEqual(numError)

    expect(getBuiltin('ASIN')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('ACOS')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('ACOSH')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('ATANH')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('ACOTH')?.(str('bad'))).toEqual(valueError)
  })

  it('matches Microsoft Excel numeric-domain errors for combinatoric functions', () => {
    expect(getBuiltin('FACT')?.(num(-1))).toEqual(numError)
    expect(getBuiltin('FACTDOUBLE')?.(num(-1))).toEqual(numError)
    expect(getBuiltin('COMBIN')?.(num(-1), num(0))).toEqual(numError)
    expect(getBuiltin('COMBIN')?.(num(2), num(3))).toEqual(numError)
    expect(getBuiltin('COMBINA')?.(num(-1), num(1))).toEqual(numError)
    expect(getBuiltin('COMBINA')?.(num(1), num(-1))).toEqual(numError)
    expect(getBuiltin('PERMUT')?.(num(0), num(1))).toEqual(numError)
    expect(getBuiltin('PERMUT')?.(num(3), num(4))).toEqual(numError)
    expect(getBuiltin('PERMUTATIONA')?.(num(0), num(1))).toEqual(numError)
    expect(getBuiltin('PERMUTATIONA')?.(num(-1), num(1))).toEqual(numError)
    expect(getBuiltin('MULTINOMIAL')?.(num(1), num(-1))).toEqual(numError)
  })

  it('matches Microsoft Excel Bessel order-domain errors', () => {
    expect(getBuiltin('BESSELI')?.(num(1), num(-1))).toEqual(numError)
    expect(getBuiltin('BESSELJ')?.(num(1), num(-1))).toEqual(numError)
    expect(getBuiltin('BESSELK')?.(num(1), num(-1))).toEqual(numError)
    expect(getBuiltin('BESSELY')?.(num(1), num(-1))).toEqual(numError)

    expect(getBuiltin('BESSELI')?.(str('bad'), num(1))).toEqual(valueError)
    expect(getBuiltin('BESSELJ')?.(num(1), str('bad'))).toEqual(valueError)
    expect(getBuiltin('BESSELK')?.(str('bad'), num(1))).toEqual(valueError)
    expect(getBuiltin('BESSELY')?.(num(1), str('bad'))).toEqual(valueError)
  })
})

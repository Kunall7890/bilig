import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'
import { evaluateAst } from '../js-evaluator.js'
import { parseFormula } from '../parser.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const str = (value: string, stringId = 1): CellValue => ({ tag: ValueTag.String, value, stringId })
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const
const div0Error = { tag: ValueTag.Error, code: ErrorCode.Div0 } as const
const numError = { tag: ValueTag.Error, code: ErrorCode.Num } as const
const nameError = { tag: ValueTag.Error, code: ErrorCode.Name } as const
const naError = { tag: ValueTag.Error, code: ErrorCode.NA } as const

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

  it('propagates scalar math argument errors before coercion and domain checks', () => {
    expect(getBuiltin('ACOT')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('COT')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('COTH')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('CSC')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('CSCH')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('SEC')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('SIGN')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('FLOOR.MATH')?.(num(1), naError)).toEqual(naError)
    expect(getBuiltin('CEILING.MATH')?.(num(1), num(1), naError)).toEqual(naError)
    expect(getBuiltin('FLOOR.PRECISE')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('CEILING.PRECISE')?.(num(1), nameError)).toEqual(nameError)
    expect(getBuiltin('ISO.CEILING')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('MOD')?.(nameError, num(0))).toEqual(nameError)
    expect(getBuiltin('INT')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('EVEN')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('ODD')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('FACT')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('FACTDOUBLE')?.(nameError)).toEqual(nameError)
    expect(getBuiltin('COMBIN')?.(nameError, num(-1))).toEqual(nameError)
    expect(getBuiltin('COMBINA')?.(num(1), naError)).toEqual(naError)
    expect(getBuiltin('QUOTIENT')?.(nameError, num(0))).toEqual(nameError)
    expect(getBuiltin('BESSELK')?.(nameError, num(-1))).toEqual(nameError)
    expect(getBuiltin('BESSELY')?.(num(0), naError)).toEqual(naError)
    expect(getBuiltin('PERMUT')?.(nameError, num(1))).toEqual(nameError)
    expect(getBuiltin('PERMUTATIONA')?.(num(1), naError)).toEqual(naError)
    expect(getBuiltin('SERIESSUM')?.(naError, num(1), num(1), num(1))).toEqual(naError)
    expect(getBuiltin('SERIESSUM')?.(num(1), num(1), num(1), naError)).toEqual(naError)
    expect(getBuiltin('DELTA')?.(naError, num(0))).toEqual(naError)
    expect(getBuiltin('GESTEP')?.(num(1), naError)).toEqual(naError)
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

  it('matches spreadsheet zero-multiple rounding semantics', () => {
    expect(getBuiltin('FLOOR')?.(num(2.5), num(0))).toEqual(div0Error)
    expect(getBuiltin('CEILING')?.(num(2.5), num(0))).toEqual(div0Error)
    expect(getBuiltin('FLOOR.MATH')?.(num(2.5), num(0))).toEqual(num(0))
    expect(getBuiltin('FLOOR.PRECISE')?.(num(2.5), num(0))).toEqual(num(0))
    expect(getBuiltin('CEILING.MATH')?.(num(2.5), num(0))).toEqual(num(0))
    expect(getBuiltin('CEILING.PRECISE')?.(num(2.5), num(0))).toEqual(num(0))
    expect(getBuiltin('ISO.CEILING')?.(num(2.5), num(0))).toEqual(num(0))
    expect(getBuiltin('MROUND')?.(num(10), num(0))).toEqual(num(0))
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

  it('matches Microsoft Excel ROUND and CEILING negative-number edge semantics', () => {
    expect(getBuiltin('ROUND')?.(num(-2.5), num(0))).toEqual(num(-3))
    expect(getBuiltin('ROUND')?.(num(-1.475), num(2))).toEqual(num(-1.48))
    expect(getBuiltin('ROUND')?.(num(-50.55), num(-2))).toEqual(num(-100))
    expect(getBuiltin('CEILING')?.(num(2.5), num(-2))).toEqual(numError)
    expect(getBuiltin('CEILING')?.(num(-2.5), num(2))).toEqual(num(-2))
    expect(getBuiltin('CEILING')?.(num(-2.5), num(-2))).toEqual(num(-4))
  })

  it('matches Microsoft Excel square-root negative-domain errors', () => {
    expect(getBuiltin('SQRT')?.(num(-1))).toEqual(numError)
    expect(getBuiltin('SQRTPI')?.(num(-1))).toEqual(numError)
    expect(getBuiltin('SQRT')?.(str('bad'))).toEqual(valueError)
    expect(getBuiltin('SQRTPI')?.(str('bad'))).toEqual(valueError)
  })

  it('matches Microsoft Excel GEOMEAN and HARMEAN direct text and positive-domain semantics', () => {
    expect(getBuiltin('GEOMEAN')?.(num(0), num(2))).toEqual(numError)
    expect(getBuiltin('GEOMEAN')?.(num(-1), num(2))).toEqual(numError)
    expect(getBuiltin('HARMEAN')?.(num(0), num(2))).toEqual(numError)
    expect(getBuiltin('HARMEAN')?.(num(-1), num(2))).toEqual(numError)

    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => {
        if (address === 'A1') {
          return str('bad')
        }
        if (address === 'A2') {
          return num(2)
        }
        if (address === 'A3') {
          return { tag: ValueTag.Boolean, value: true }
        }
        if (address === 'A4') {
          return num(0)
        }
        return { tag: ValueTag.Empty }
      },
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        if (start === 'A1' && end === 'A3') {
          return [str('bad'), num(2), { tag: ValueTag.Boolean, value: true }]
        }
        if (start === 'A1' && end === 'A4') {
          return [str('bad'), num(2), { tag: ValueTag.Boolean, value: true }, num(0)]
        }
        return []
      },
    }

    expect(evaluateAst(parseFormula('GEOMEAN("2","8")'), context)).toEqual(num(4))
    expect(evaluateAst(parseFormula('GEOMEAN("bad",2)'), context)).toEqual(valueError)
    expect(evaluateAst(parseFormula('GEOMEAN(A1:A3)'), context)).toEqual(num(2))
    expect(evaluateAst(parseFormula('GEOMEAN(A1:A4)'), context)).toEqual(numError)

    expect(evaluateAst(parseFormula('HARMEAN("2","8")'), context)).toEqual(num(3.2))
    expect(evaluateAst(parseFormula('HARMEAN("bad",2)'), context)).toEqual(valueError)
    expect(evaluateAst(parseFormula('HARMEAN(A1:A3)'), context)).toEqual(num(2))
    expect(evaluateAst(parseFormula('HARMEAN(A1:A4)'), context)).toEqual(numError)
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

  it('keeps combinatoric cancellation finite and reports unrepresentable results as #NUM', () => {
    expect(getBuiltin('FACT')?.(num(171))).toEqual(numError)
    expect(getBuiltin('FACTDOUBLE')?.(num(301))).toEqual(numError)
    expect(getBuiltin('COMBIN')?.(num(171), num(1))).toEqual(num(171))
    expect(getBuiltin('COMBIN')?.(num(171), num(2))).toEqual(num(14535))
    expect(getBuiltin('COMBINA')?.(num(171), num(1))).toEqual(num(171))

    const repeatedCombination = getBuiltin('COMBINA')?.(num(100), num(100))
    expect(repeatedCombination?.tag).toBe(ValueTag.Number)
    expect(repeatedCombination?.tag === ValueTag.Number ? repeatedCombination.value / 4.5274257328e58 : Number.NaN).toBeCloseTo(1, 11)

    expect(getBuiltin('PERMUT')?.(num(200), num(170))).toEqual(numError)
    expect(getBuiltin('PERMUTATIONA')?.(num(200), num(200))).toEqual(numError)
    expect(getBuiltin('MULTINOMIAL')?.(num(171), num(1))).toEqual(num(172))
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

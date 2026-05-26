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

function date(year: number, month: number, day: number): CellValue {
  const result = getBuiltin('DATE')!(num(year), num(month), num(day))
  if (result.tag !== ValueTag.Number) {
    throw new Error(`DATE(${year},${month},${day}) did not produce a serial`)
  }
  return result
}

describe('fixed-income formula domain errors', () => {
  it('returns #NUM! for discounted security numeric-domain errors while preserving #VALUE! coercion errors', () => {
    const DISC = getBuiltin('DISC')!
    const INTRATE = getBuiltin('INTRATE')!
    const RECEIVED = getBuiltin('RECEIVED')!
    const PRICEDISC = getBuiltin('PRICEDISC')!
    const YIELDDISC = getBuiltin('YIELDDISC')!

    const settlement = date(2023, 1, 1)
    const maturity = date(2023, 4, 1)

    expect(DISC(maturity, settlement, num(97), num(100), num(2))).toEqual(numError)
    expect(DISC(settlement, maturity, num(0), num(100), num(2))).toEqual(numError)
    expect(DISC(settlement, maturity, num(97), num(100), num(5))).toEqual(numError)
    expect(DISC(settlement, maturity, text('bad'), num(100), num(2))).toEqual(valueError)

    expect(INTRATE(settlement, maturity, num(0), num(1030), num(2))).toEqual(numError)
    expect(INTRATE(settlement, maturity, num(1000), num(1030), num(-1))).toEqual(numError)
    expect(INTRATE(settlement, maturity, text('bad', 2), num(1030), num(2))).toEqual(valueError)

    expect(RECEIVED(settlement, maturity, num(1000), num(0), num(2))).toEqual(numError)
    expect(RECEIVED(maturity, settlement, num(1000), num(0.12), num(2))).toEqual(numError)

    expect(PRICEDISC(maturity, settlement, num(0.0525), num(100), num(2))).toEqual(numError)
    expect(PRICEDISC(settlement, maturity, num(0), num(100), num(2))).toEqual(numError)

    expect(YIELDDISC(maturity, settlement, num(99.795), num(100), num(2))).toEqual(numError)
    expect(YIELDDISC(settlement, maturity, num(0), num(100), num(2))).toEqual(numError)
  })

  it('returns #NUM! for Treasury bill numeric-domain errors while preserving #VALUE! coercion errors', () => {
    const TBILLPRICE = getBuiltin('TBILLPRICE')!
    const TBILLYIELD = getBuiltin('TBILLYIELD')!
    const TBILLEQ = getBuiltin('TBILLEQ')!

    const settlement = date(2008, 3, 31)
    const maturity = date(2008, 6, 1)
    const moreThanOneYear = date(2009, 6, 1)

    expect(TBILLPRICE(settlement, maturity, num(0))).toEqual(numError)
    expect(TBILLPRICE(maturity, settlement, num(0.09))).toEqual(numError)
    expect(TBILLPRICE(settlement, moreThanOneYear, num(0.09))).toEqual(numError)
    expect(TBILLPRICE(settlement, maturity, text('bad'))).toEqual(valueError)

    expect(TBILLYIELD(settlement, maturity, num(0))).toEqual(numError)
    expect(TBILLYIELD(settlement, moreThanOneYear, num(98.45))).toEqual(numError)
    expect(TBILLYIELD(settlement, maturity, text('bad'))).toEqual(valueError)

    expect(TBILLEQ(settlement, maturity, num(0))).toEqual(numError)
    expect(TBILLEQ(settlement, moreThanOneYear, num(0.0914))).toEqual(numError)
    expect(TBILLEQ(settlement, maturity, text('bad'))).toEqual(valueError)
  })

  it('preserves incoming fixed-income errors before scalar coercion', () => {
    const COUPDAYBS = getBuiltin('COUPDAYBS')!
    const DISC = getBuiltin('DISC')!
    const TBILLPRICE = getBuiltin('TBILLPRICE')!

    const settlement = date(2023, 1, 1)
    const maturity = date(2023, 4, 1)

    expect(DISC(settlement, maturity, err(ErrorCode.Name), num(100), num(2))).toEqual(err(ErrorCode.Name))
    expect(TBILLPRICE(settlement, maturity, err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(COUPDAYBS(err(ErrorCode.Ref), maturity, num(2), num(0))).toEqual(err(ErrorCode.Ref))
  })
})

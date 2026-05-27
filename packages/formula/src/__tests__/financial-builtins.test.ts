import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import {
  cumulativePeriodicPayment,
  dbDepreciation,
  ddbDepreciation,
  futureValue,
  interestPayment,
  periodicPayment,
  presentValue,
  principalPayment,
  solveRate,
  totalPeriods,
  vdbDepreciation,
} from '../builtins/financial.js'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const str = (value: string, stringId = 1): CellValue => ({ tag: ValueTag.String, value, stringId })
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const

describe('financial helpers', () => {
  it('computes time-value-of-money helpers', () => {
    expect(futureValue(0.1, 2, -100, -1000, 0)).toBeCloseTo(1420, 12)
    expect(presentValue(0.1, 2, -100, 1420, 0)).toBeCloseTo(-1000, 12)
    expect(periodicPayment(0.1, 2, 1000, 0, 0)).toBeCloseTo(-576.1904761904761, 12)
    expect(totalPeriods(0.1, -576.1904761904761, 1000, 0, 0)).toBeCloseTo(2, 12)
    expect(solveRate(48, -200, 8000, 0, 0, 0.1)).toBeCloseTo(0.007701472488246008, 12)
  })

  it('covers time-value validation and zero-rate branches', () => {
    expect(futureValue(0, 3, -10, -100, 0)).toBe(130)
    expect(presentValue(0, 3, -10, 130, 0)).toBe(-100)
    expect(periodicPayment(0, 4, 100, 20, 0)).toBe(-30)
    expect(periodicPayment(0.1, 0, 100, 20, 0)).toBeUndefined()
    expect(periodicPayment(-1, 2, 100, 20, 1)).toBeUndefined()
    expect(totalPeriods(0, -20, 100, 0, 0)).toBe(5)
    expect(totalPeriods(0, 0, 100, 0, 0)).toBeUndefined()
    expect(totalPeriods(0.1, 0, 100, 0, 0)).toBeUndefined()
    expect(solveRate(0, -200, 8000, 0, 0, 0.1)).toBeUndefined()
    expect(solveRate(2, -50, 100, 0, 0, 0)).toBe(0)
    expect(solveRate(48, -200, 8000, 0, 0, Number.NaN)).toBeCloseTo(0.007701472488246008, 12)
    expect(solveRate(48, -200, 8000, 0, 0, -5)).toBeUndefined()
  })

  it('computes depreciation helpers', () => {
    expect(dbDepreciation(10000, 1000, 5, 1, 12)).toBeCloseTo(3690, 12)
    expect(ddbDepreciation(2400, 300, 10, 2, 2)).toBeCloseTo(384, 12)
    expect(vdbDepreciation(2400, 300, 10, 1, 3, 2, false)).toBeCloseTo(691.2, 12)
  })

  it('covers depreciation validation and partial-period branches', () => {
    expect(dbDepreciation(10000, 1000, 5, 1, 6)).toBeCloseTo(1845, 12)
    expect(dbDepreciation(10000, 1000, 5, 6, 6)).toBeCloseTo(238.5271245878818, 12)
    expect(dbDepreciation(0, 1000, 5, 1, 12)).toBeUndefined()
    expect(dbDepreciation(10000, 1000, 5, 1, 13)).toBeUndefined()
    expect(ddbDepreciation(2400, 300, 10, 1.5, 2)).toBeCloseTo(192, 12)
    expect(ddbDepreciation(2400, 300, 10, 2, 0)).toBeUndefined()
    expect(vdbDepreciation(2400, 300, 10, 0, 1, 2, true)).toBeCloseTo(480, 12)
    expect(vdbDepreciation(2400, 300, 10, 2, 2, 2, false)).toBe(0)
    expect(vdbDepreciation(2400, 300, 10, 3, 2, 2, false)).toBeUndefined()
  })

  it('computes interest and principal helpers', () => {
    expect(interestPayment(0.1, 1, 2, 1000, 0, 0)).toBeCloseTo(-100, 12)
    expect(principalPayment(0.1, 1, 2, 1000, 0, 0)).toBeCloseTo(-476.19047619047615, 12)
    expect(cumulativePeriodicPayment(0.09 / 12, 30 * 12, 125000, 13, 24, 0, false)).toBeCloseTo(-11135.232130750845, 12)
    expect(cumulativePeriodicPayment(0.09 / 12, 30 * 12, 125000, 13, 24, 0, true)).toBeCloseTo(-934.1071234208765, 12)
  })

  it('covers interest, principal, and cumulative validation branches', () => {
    expect(interestPayment(0.1, 1, 2, 1000, 0, 1)).toBe(0)
    expect(interestPayment(0.1, 0, 2, 1000, 0, 0)).toBeUndefined()
    expect(interestPayment(-1, 1, 2, 1000, 0, 1)).toBeUndefined()
    expect(principalPayment(0.1, 0, 2, 1000, 0, 0)).toBeUndefined()
    expect(cumulativePeriodicPayment(0, 30 * 12, 125000, 13, 24, 0, false)).toBeUndefined()
    expect(cumulativePeriodicPayment(0.09 / 12, 30 * 12, 125000, 24, 13, 0, false)).toBeUndefined()
    expect(cumulativePeriodicPayment(0.09 / 12, 30 * 12, 125000, 13, 400, 0, false)).toBeUndefined()
    expect(cumulativePeriodicPayment(0.1, 2, 100, 1, 1, -10, false)).toBeUndefined()
  })
})

describe('financial builtins', () => {
  it('coerces direct numeric text for annuity functions', () => {
    expect(getBuiltin('FV')?.(str('0.1'), str('2'), str('-100'), str('-1000'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(1420, 10),
    })
    expect(getBuiltin('PV')?.(str('0.1'), str('2'), str('-576.1904761904761'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(1000, 10),
    })
    expect(getBuiltin('PMT')?.(str('0.1'), str('2'), str('1000'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(-576.1904761904758, 10),
    })
    expect(getBuiltin('PMT')?.(str('0.1'), str('2'), str('1000'), str('0'), str('1'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(-523.8095238095234, 10),
    })
    expect(getBuiltin('NPER')?.(str('0.1'), str('-576.1904761904761'), str('1000'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 10),
    })
    expect(getBuiltin('IPMT')?.(str('0.1'), str('1'), str('2'), str('1000'))).toEqual(num(-100))
    expect(getBuiltin('PPMT')?.(str('0.1'), str('1'), str('2'), str('1000'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(-476.1904761904758, 10),
    })
  })

  it('coerces direct numeric text for cumulative payment period arguments', () => {
    expect(getBuiltin('CUMIPMT')?.(str('0.0075'), str('360'), str('125000'), str('13'), str('24'), str('0'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(-11135.232130750845, 8),
    })
    expect(getBuiltin('CUMPRINC')?.(str('0.0075'), str('360'), str('125000'), str('13'), str('24'), str('0'))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(-934.1071234208765, 8),
    })
  })

  it('rejects invalid direct financial text instead of defaulting to zero', () => {
    expect(getBuiltin('PMT')?.(str('bad'), str('2'), str('1000'))).toEqual(valueError)
    expect(getBuiltin('CUMIPMT')?.(str('0.0075'), str('bad'), str('125000'), str('13'), str('24'), str('0'))).toEqual(valueError)
  })
})

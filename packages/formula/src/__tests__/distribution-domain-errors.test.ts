import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string, stringId = 1): CellValue => ({ tag: ValueTag.String, value, stringId })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })

const numError = { tag: ValueTag.Error, code: ErrorCode.Num } as const
const naError = { tag: ValueTag.Error, code: ErrorCode.NA } as const
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const

describe('distribution builtins domain errors', () => {
  it('separates chi-square numeric-domain errors from coercion errors', () => {
    expect(getBuiltin('CHIDIST')?.(num(-1), num(1))).toEqual(numError)
    expect(getBuiltin('CHISQ.DIST.RT')?.(num(1), num(0))).toEqual(numError)
    expect(getBuiltin('CHISQ.DIST')?.(num(1), num(0), bool(true))).toEqual(numError)
    expect(getBuiltin('CHISQ.INV.RT')?.(num(-0.1), num(1))).toEqual(numError)
    expect(getBuiltin('CHISQ.INV')?.(num(0.5), num(0))).toEqual(numError)

    expect(getBuiltin('CHIDIST')?.(text('bad'), num(1))).toEqual(valueError)
    expect(getBuiltin('CHISQ.DIST')?.(num(1), num(1), text('bad'))).toEqual(valueError)
    expect(getBuiltin('CHISQ.INV')?.(text('bad'), num(1))).toEqual(valueError)
  })

  it('separates beta numeric-domain errors from coercion errors', () => {
    expect(getBuiltin('BETA.DIST')?.(num(0.5), num(0), num(1), bool(true))).toEqual(numError)
    expect(getBuiltin('BETA.DIST')?.(num(-0.1), num(1), num(1), bool(true))).toEqual(numError)
    expect(getBuiltin('BETADIST')?.(num(0.5), num(1), num(1), num(1), num(1))).toEqual(numError)
    expect(getBuiltin('BETA.INV')?.(num(-0.1), num(1), num(1))).toEqual(numError)
    expect(getBuiltin('BETAINV')?.(num(0.5), num(0), num(1))).toEqual(numError)

    expect(getBuiltin('BETA.DIST')?.(text('bad'), num(1), num(1), bool(true))).toEqual(valueError)
    expect(getBuiltin('BETA.DIST')?.(num(0.5), num(1), num(1), text('bad'))).toEqual(valueError)
    expect(getBuiltin('BETA.INV')?.(text('bad'), num(1), num(1))).toEqual(valueError)
  })

  it('separates f-distribution numeric-domain errors from coercion errors', () => {
    expect(getBuiltin('F.DIST')?.(num(-1), num(1), num(1), bool(true))).toEqual(numError)
    expect(getBuiltin('F.DIST')?.(num(1), num(0), num(1), bool(true))).toEqual(numError)
    expect(getBuiltin('F.DIST.RT')?.(num(1), num(1), num(0))).toEqual(numError)
    expect(getBuiltin('FDIST')?.(num(-1), num(1), num(1))).toEqual(numError)
    expect(getBuiltin('F.INV')?.(num(-0.1), num(1), num(1))).toEqual(numError)
    expect(getBuiltin('F.INV.RT')?.(num(0.5), num(0), num(1))).toEqual(numError)
    expect(getBuiltin('FINV')?.(num(1.1), num(1), num(1))).toEqual(numError)

    expect(getBuiltin('F.DIST')?.(text('bad'), num(1), num(1), bool(true))).toEqual(valueError)
    expect(getBuiltin('F.DIST')?.(num(1), num(1), num(1), text('bad'))).toEqual(valueError)
    expect(getBuiltin('F.INV')?.(text('bad'), num(1), num(1))).toEqual(valueError)
  })

  it('separates gamma inverse numeric-domain errors from coercion errors', () => {
    expect(getBuiltin('GAMMA.INV')?.(num(-0.1), num(3), num(2))).toEqual(numError)
    expect(getBuiltin('GAMMA.INV')?.(num(1.1), num(3), num(2))).toEqual(numError)
    expect(getBuiltin('GAMMA.INV')?.(num(0.5), num(0), num(2))).toEqual(numError)
    expect(getBuiltin('GAMMAINV')?.(num(0.5), num(3), num(0))).toEqual(numError)

    expect(getBuiltin('GAMMA.INV')?.(text('bad'), num(3), num(2))).toEqual(valueError)
    expect(getBuiltin('GAMMA.INV')?.(num(0.5), text('bad'), num(2))).toEqual(valueError)
  })

  it('separates confidence numeric-domain errors from coercion errors', () => {
    expect(getBuiltin('CONFIDENCE.NORM')?.(num(0), num(1), num(10))).toEqual(numError)
    expect(getBuiltin('CONFIDENCE.NORM')?.(num(0.05), num(0), num(10))).toEqual(numError)
    expect(getBuiltin('CONFIDENCE.NORM')?.(num(0.05), num(1), num(0))).toEqual(numError)
    expect(getBuiltin('CONFIDENCE')?.(num(1), num(1), num(10))).toEqual(numError)
    expect(getBuiltin('CONFIDENCE.T')?.(num(0), num(1), num(10))).toEqual(numError)
    expect(getBuiltin('CONFIDENCE.T')?.(num(0.05), num(0), num(10))).toEqual(numError)
    expect(getBuiltin('CONFIDENCE.T')?.(num(0.05), num(1), num(0))).toEqual(numError)
    expect(getBuiltin('CONFIDENCE.T')?.(num(0.05), num(1), num(1))).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })

    expect(getBuiltin('CONFIDENCE.NORM')?.(text('bad'), num(1), num(10))).toEqual(valueError)
    expect(getBuiltin('CONFIDENCE.T')?.(num(0.05), text('bad'), num(10))).toEqual(valueError)
  })

  it('truncates confidence size arguments before applying numeric-domain checks', () => {
    expect(getBuiltin('CONFIDENCE.NORM')?.(num(0.05), num(1), num(1.9))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(1.959963984540054, 8),
    })
    expect(getBuiltin('CONFIDENCE')?.(num(0.05), num(1), num(1.9))).toEqual({
      tag: ValueTag.Number,
      value: expect.closeTo(1.959963984540054, 8),
    })
    expect(getBuiltin('CONFIDENCE.NORM')?.(num(0.05), num(1), num(0.9))).toEqual(numError)
    expect(getBuiltin('CONFIDENCE.T')?.(num(0.05), num(1), num(1.9))).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
  })

  it('separates student-t numeric-domain errors from coercion errors', () => {
    expect(getBuiltin('T.DIST')?.(num(1), num(0), bool(true))).toEqual(numError)
    expect(getBuiltin('T.DIST.RT')?.(num(1), num(0))).toEqual(numError)
    expect(getBuiltin('T.DIST.2T')?.(num(-1), num(1))).toEqual(numError)
    expect(getBuiltin('TDIST')?.(num(1), num(1), num(3))).toEqual(numError)
    expect(getBuiltin('T.INV')?.(num(-0.1), num(1))).toEqual(numError)
    expect(getBuiltin('T.INV')?.(num(1), num(1))).toEqual(numError)
    expect(getBuiltin('T.INV.2T')?.(num(0.5), num(0))).toEqual(numError)
    expect(getBuiltin('TINV')?.(num(1.1), num(1))).toEqual(numError)

    expect(getBuiltin('T.INV.2T')?.(num(1), num(1))).toEqual(num(0))

    expect(getBuiltin('T.DIST')?.(text('bad'), num(1), bool(true))).toEqual(valueError)
    expect(getBuiltin('T.DIST')?.(num(1), num(1), text('bad'))).toEqual(valueError)
    expect(getBuiltin('T.INV')?.(text('bad'), num(1))).toEqual(valueError)
  })

  it('separates discrete distribution numeric-domain errors from coercion errors', () => {
    expect(getBuiltin('BINOMDIST')?.(num(4), num(3), num(0.5), bool(true))).toEqual(numError)
    expect(getBuiltin('BINOM.DIST')?.(num(1), num(3), num(-0.1), bool(true))).toEqual(numError)
    expect(getBuiltin('BINOM.DIST.RANGE')?.(num(3), num(-0.1), num(1))).toEqual(numError)
    expect(getBuiltin('BINOM.DIST.RANGE')?.(num(3), num(0.5), num(2), num(1))).toEqual(numError)
    expect(getBuiltin('CRITBINOM')?.(num(-1), num(0.5), num(0.7))).toEqual(numError)
    expect(getBuiltin('BINOM.INV')?.(num(3), num(0.5), num(0))).toEqual(numError)
    expect(getBuiltin('HYPGEOMDIST')?.(num(5), num(4), num(3), num(10))).toEqual(numError)
    expect(getBuiltin('HYPGEOM.DIST')?.(num(1), num(11), num(3), num(10), bool(true))).toEqual(numError)
    expect(getBuiltin('HYPGEOM.DIST')?.(num(1), num(4), num(11), num(10), bool(true))).toEqual(numError)
    expect(getBuiltin('NEGBINOMDIST')?.(num(-1), num(3), num(0.5))).toEqual(numError)
    expect(getBuiltin('NEGBINOM.DIST')?.(num(1), num(0), num(0.5), bool(true))).toEqual(numError)
    expect(getBuiltin('NEGBINOM.DIST')?.(num(1), num(3), num(-0.1), bool(true))).toEqual(numError)

    expect(getBuiltin('BINOMDIST')?.(text('bad'), num(3), num(0.5), bool(true))).toEqual(valueError)
    expect(getBuiltin('BINOM.DIST.RANGE')?.(num(3), text('bad'), num(1))).toEqual(valueError)
    expect(getBuiltin('CRITBINOM')?.(text('bad'), num(0.5), num(0.7))).toEqual(valueError)
    expect(getBuiltin('HYPGEOM.DIST')?.(text('bad'), num(4), num(3), num(10), bool(true))).toEqual(valueError)
    expect(getBuiltin('NEGBINOM.DIST')?.(text('bad'), num(3), num(0.5), bool(true))).toEqual(valueError)
  })

  it('preserves incoming errors before distribution coercion and domain checks', () => {
    const cases = [
      getBuiltin('CONFIDENCE.NORM')?.(naError, num(2), num(30)),
      getBuiltin('CONFIDENCE.T')?.(num(0.05), num(2), naError),
      getBuiltin('ERF')?.(num(0), naError),
      getBuiltin('ERFC')?.(naError),
      getBuiltin('FISHER')?.(naError),
      getBuiltin('GAMMALN')?.(naError),
      getBuiltin('GAMMA')?.(naError),
      getBuiltin('BETA.INV')?.(num(0.5), naError, num(3)),
      getBuiltin('GAMMA.INV')?.(num(0.5), num(2), naError),
      getBuiltin('CHIDIST')?.(num(1), naError),
      getBuiltin('CHISQ.INV')?.(naError, num(2)),
      getBuiltin('F.INV')?.(num(0.5), naError, num(3)),
      getBuiltin('T.INV')?.(naError, num(2)),
      getBuiltin('CRITBINOM')?.(num(3), num(0.5), naError),
      getBuiltin('HYPGEOMDIST')?.(num(1), num(2), num(3), naError),
      getBuiltin('NEGBINOMDIST')?.(num(2), num(3), naError),
    ]

    for (const actual of cases) {
      expect(actual).toEqual(naError)
    }
  })
})

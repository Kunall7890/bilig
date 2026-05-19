import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'

import { getLookupBuiltin, normalizeExactLookupNumber, sameExactLookupNumber, type RangeBuiltinArgument } from '../index.js'

function num(value: number) {
  return { tag: ValueTag.Number, value } as const
}

function text(value: string) {
  return { tag: ValueTag.String, value } as const
}

function cellRange(values: RangeBuiltinArgument['values'], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows, cols }
}

describe('lookup exact numeric matching', () => {
  it('preserves ordinary integer lookup keys exactly', () => {
    expect(normalizeExactLookupNumber(9_999)).toBe(9_999)
    expect(normalizeExactLookupNumber(999_999_999_999_999)).toBe(999_999_999_999_999)
    expect(sameExactLookupNumber(9_999, 9_999)).toBe(true)
  })

  it('matches formula-roundoff decimals using spreadsheet numeric precision', () => {
    expect(normalizeExactLookupNumber(2374.2799999999997)).toBe(2374.28)
    expect(sameExactLookupNumber(2374.2799999999997, 2374.28)).toBe(true)

    const XLOOKUP = getLookupBuiltin('XLOOKUP')!

    expect(
      XLOOKUP(num(2374.2799999999997), cellRange([num(2374.28)], 1, 1), cellRange([text('2026-04-06')], 1, 1), text(''), num(0)),
    ).toEqual(text('2026-04-06'))
  })
})

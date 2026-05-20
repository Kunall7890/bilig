import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import {
  cellValueCriteriaString,
  cellValuesEqual,
  decodeErrorCode,
  decodeRuntimeTag,
  directCriteriaCacheValueKey,
  directErrorResult,
  directNumberResult,
  evaluationErrorMessage,
  numericLikeValueAt,
  offsetDirectAggregateResult,
  referenceReplacementKey,
  sameExactNumericValue,
  strictNumericAggregateCandidateAt,
} from '../engine/services/formula-evaluation-helpers.js'
import { rememberDirectCriteriaResult } from '../engine/services/formula-evaluation-direct-criteria-cache.js'
import type { RuntimeDirectAggregateDescriptor } from '../engine/runtime-state.js'
import type { RuntimeColumnSlice } from '../engine/services/runtime-column-store-service.js'

function slice(tags: number[], numbers: number[]): RuntimeColumnSlice {
  return {
    sheetName: 'Sheet1',
    rowStart: 0,
    rowEnd: tags.length - 1,
    col: 0,
    length: tags.length,
    columnVersion: 1,
    structureVersion: 1,
    sheetColumnVersions: new Uint32Array(),
    tags: Uint8Array.from(tags),
    numbers: Float64Array.from(numbers),
    stringIds: new Uint32Array(tags.length),
    errors: new Uint16Array(tags.length),
  }
}

describe('formula evaluation helpers', () => {
  it('decodes runtime tags and errors with stable defaults', () => {
    expect(decodeErrorCode(undefined)).toBe(ErrorCode.None)
    expect(decodeErrorCode(ErrorCode.Ref)).toBe(ErrorCode.Ref)
    expect(decodeRuntimeTag(undefined)).toBe(ValueTag.Empty)
    expect(decodeRuntimeTag(0)).toBe(ValueTag.Empty)
    expect(decodeRuntimeTag(1)).toBe(ValueTag.Number)
    expect(decodeRuntimeTag(2)).toBe(ValueTag.Boolean)
    expect(decodeRuntimeTag(3)).toBe(ValueTag.String)
    expect(decodeRuntimeTag(4)).toBe(ValueTag.Error)
    expect(decodeRuntimeTag(99)).toBe(ValueTag.Empty)
    expect(evaluationErrorMessage('fallback', new Error('explicit'))).toBe('explicit')
    expect(evaluationErrorMessage('fallback', new Error(''))).toBe('fallback')
    expect(evaluationErrorMessage('fallback', 'not-error')).toBe('fallback')
  })

  it('formats formula values for lookup criteria and cache keys', () => {
    const values: CellValue[] = [
      { tag: ValueTag.Empty },
      { tag: ValueTag.Number, value: -0 },
      { tag: ValueTag.Boolean, value: true },
      { tag: ValueTag.Boolean, value: false },
      { tag: ValueTag.String, value: 'North', stringId: 7 },
      { tag: ValueTag.Error, code: ErrorCode.Name },
    ]

    expect(values.map(cellValueCriteriaString)).toEqual(['', '0', 'TRUE', 'FALSE', 'North', String(ErrorCode.Name)])
    expect(values.map(directCriteriaCacheValueKey)).toEqual(['e:', 'n:0', 'b:1', 'b:0', 's:North', `r:${ErrorCode.Name}`])
    expect(referenceReplacementKey(' sheet1 ', ' a1 ')).toBe('SHEET1!A1')
  })

  it('evicts the oldest direct criteria aggregate cache entry at capacity', () => {
    const cache = new Map<string, CellValue>()
    for (let index = 0; index < 16_384; index += 1) {
      cache.set(`key-${index}`, { tag: ValueTag.Number, value: index })
    }

    const value = { tag: ValueTag.Number, value: 16_384 } satisfies CellValue
    expect(rememberDirectCriteriaResult(cache, 'newest', value)).toBe(value)
    expect(cache.has('key-0')).toBe(false)
    expect(cache.get('newest')).toBe(value)
  })

  it('compares exact cell values by tag and payload', () => {
    expect(cellValuesEqual({ tag: ValueTag.Empty }, { tag: ValueTag.Empty })).toBe(true)
    expect(cellValuesEqual({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toBe(true)
    expect(cellValuesEqual({ tag: ValueTag.Boolean, value: true }, { tag: ValueTag.Boolean, value: false })).toBe(false)
    expect(cellValuesEqual({ tag: ValueTag.String, value: 'x', stringId: 1 }, { tag: ValueTag.String, value: 'x', stringId: 2 })).toBe(true)
    expect(cellValuesEqual({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toBe(true)
    expect(cellValuesEqual({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: '1', stringId: 1 })).toBe(false)
    expect(sameExactNumericValue(0, -0)).toBe(true)
    expect(sameExactNumericValue(1, 2)).toBe(false)
  })

  it('builds direct results and applies numeric aggregate offsets only to numbers', () => {
    const directAggregate: RuntimeDirectAggregateDescriptor = {
      regionId: 1,
      aggregateKind: 'sum',
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 0,
      col: 0,
      colEnd: 0,
      length: 1,
      resultOffset: 3,
    }

    expect(directNumberResult(12)).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(directErrorResult(ErrorCode.Div0)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(offsetDirectAggregateResult(directAggregate, { tag: ValueTag.Number, value: 9 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(offsetDirectAggregateResult(directAggregate, { tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(offsetDirectAggregateResult(undefined, { tag: ValueTag.Number, value: 9 })).toEqual({ tag: ValueTag.Number, value: 9 })
  })

  it('reads numeric-like column values for aggregate candidates', () => {
    const runtimeSlice = slice([ValueTag.Number, ValueTag.Boolean, ValueTag.Empty, ValueTag.String, ValueTag.Error], [4, 1, 0, 99, 88])

    expect(numericLikeValueAt(runtimeSlice, 0)).toBe(4)
    expect(numericLikeValueAt(runtimeSlice, 1)).toBe(1)
    expect(numericLikeValueAt(runtimeSlice, 2)).toBe(0)
    expect(numericLikeValueAt(runtimeSlice, 3)).toBeUndefined()
    expect(numericLikeValueAt(runtimeSlice, 4)).toBeUndefined()
    expect(strictNumericAggregateCandidateAt(runtimeSlice, 0)).toBe(4)
    expect(strictNumericAggregateCandidateAt(runtimeSlice, 1)).toBeUndefined()
  })
})

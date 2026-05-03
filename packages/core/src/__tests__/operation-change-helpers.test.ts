import { describe, expect, it } from 'vitest'
import {
  assertNever,
  cellRange,
  makeCompactExistingNumericMutationResult,
  makeExistingNumericMutationResult,
  mergeChangedCellIndices,
  mutationErrorMessage,
  normalizeRange,
  rangesIntersect,
  reverseUint32Array,
  tagTrustedPhysicalTrackedChanges,
  throwProtectionBlocked,
} from '../engine/services/operation-change-helpers.js'

describe('operation change helpers', () => {
  it('normalizes and compares operation ranges', () => {
    expect(cellRange('Sheet1', 'B2')).toEqual({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' })
    expect(normalizeRange({ sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'A1' })).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C3',
      startRow: 0,
      endRow: 2,
      startCol: 0,
      endCol: 2,
    })
    expect(
      rangesIntersect(
        { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'A1' },
        { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'D4' },
      ),
    ).toBe(true)
    expect(
      rangesIntersect(
        { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
        { sheetName: 'Other', startAddress: 'A1', endAddress: 'A2' },
      ),
    ).toBe(false)
  })

  it('merges changed cell indices without duplicates', () => {
    expect(Array.from(mergeChangedCellIndices([], [1, 2]))).toEqual([1, 2])
    expect(Array.from(mergeChangedCellIndices([1, 2], []))).toEqual([1, 2])
    expect(Array.from(mergeChangedCellIndices([1], [1]))).toEqual([1])
    expect(Array.from(mergeChangedCellIndices([1], [2]))).toEqual([1, 2])
    expect(Array.from(mergeChangedCellIndices([1, 2], [2, 3]))).toEqual([1, 2, 3])
  })

  it('reverses changed cell index arrays without mutating the input', () => {
    const values = Uint32Array.of(3, 4, 5)

    expect(Array.from(reverseUint32Array(values))).toEqual([5, 4, 3])
    expect(Array.from(values)).toEqual([3, 4, 5])
  })

  it('builds tracked physical mutation results', () => {
    const tagged = Uint32Array.of(1, 2)
    tagTrustedPhysicalTrackedChanges(tagged, 9, 1)

    expect(Reflect.get(tagged, '__biligTrackedPhysicalSheetId')).toBe(9)
    expect(Reflect.get(tagged, '__biligTrackedPhysicalSortedSliceSplit')).toBe(1)
    expect(makeExistingNumericMutationResult(tagged, 1)).toEqual({ changedCellIndices: tagged, explicitChangedCount: 1 })
    expect(makeCompactExistingNumericMutationResult(1, undefined, 1)).toEqual({
      firstChangedCellIndex: 1,
      changedCellCount: 1,
      explicitChangedCount: 1,
    })
    expect(makeCompactExistingNumericMutationResult(1, 2, 1, 12, { row: 3, col: 4 })).toEqual({
      firstChangedCellIndex: 1,
      secondChangedCellIndex: 2,
      changedCellCount: 2,
      explicitChangedCount: 1,
      secondChangedNumericValue: 12,
      secondChangedRow: 3,
      secondChangedCol: 4,
    })
  })

  it('formats operation errors consistently', () => {
    expect(mutationErrorMessage('fallback', new Error('from cause'))).toBe('from cause')
    expect(mutationErrorMessage('fallback', new Error(''))).toBe('fallback')
    expect(mutationErrorMessage('fallback', 'plain cause')).toBe('fallback')
    expect(() => throwProtectionBlocked('locked')).toThrow('Workbook protection blocks this change: locked')
    expect(() => {
      Reflect.apply(assertNever, undefined, ['unexpected'])
    }).toThrow('Unexpected value: unexpected')
  })
})

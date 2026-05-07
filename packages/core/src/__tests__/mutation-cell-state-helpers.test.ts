import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { CellFlags } from '../cell-store.js'
import {
  hasMutationCellContent,
  hasStoredMutationCellState,
  readDesiredMutationCellState,
  readStoredMutationCellState,
  shouldApplyMutationCellState,
} from '../engine/services/mutation-cell-state-helpers.js'

const emptyCell = (flags = 0, format: string | null = null): CellSnapshot => ({
  value: { tag: ValueTag.Empty },
  formula: undefined,
  format,
  flags,
})

describe('mutation cell state helpers', () => {
  it('normalizes stored formula, scalar, and blank states', () => {
    expect(
      readStoredMutationCellState(
        {
          value: { tag: ValueTag.Empty },
          formula: 'A1+1',
          format: null,
          flags: 0,
        },
        'fmt-1',
        undefined,
      ),
    ).toEqual({ formula: 'A1+1', value: null, format: 'fmt-1' })
    expect(
      readStoredMutationCellState(
        {
          value: { tag: ValueTag.Number, value: 12 },
          formula: undefined,
          format: null,
          flags: 0,
        },
        null,
        undefined,
      ),
    ).toEqual({ value: 12, format: null })
    expect(readStoredMutationCellState(emptyCell(), null, CellFlags.AuthoredBlank)).toEqual({
      value: null,
      format: null,
      authoredBlank: true,
    })
  })

  it('normalizes desired state and translates same-sheet formulas', () => {
    expect(
      readDesiredMutationCellState({
        targetSheetName: 'Sheet1',
        targetAddress: 'C3',
        sourceSheetName: 'Sheet1',
        sourceAddress: 'A1',
        snapshot: {
          value: { tag: ValueTag.Empty },
          formula: '=A1',
          format: 'fmt-1',
          flags: 0,
        },
      }),
    ).toEqual({ formula: 'C3', value: null, format: 'fmt-1' })
    expect(
      readDesiredMutationCellState({
        targetSheetName: 'Sheet1',
        targetAddress: 'A1',
        formatOverride: 'override',
        snapshot: {
          value: { tag: ValueTag.String, value: 'x' },
          formula: undefined,
          format: 'fmt-1',
          flags: 0,
        },
      }),
    ).toEqual({ value: 'x', format: 'override' })
  })

  it('detects content and stored state separately', () => {
    expect(hasMutationCellContent(emptyCell())).toBe(false)
    expect(hasMutationCellContent(emptyCell(CellFlags.AuthoredBlank))).toBe(true)
    expect(hasStoredMutationCellState(emptyCell(), undefined, undefined)).toBe(false)
    expect(hasStoredMutationCellState(emptyCell(), 'fmt-1', undefined)).toBe(true)
  })

  it('preserves the authored blank no-op edge case', () => {
    expect(
      shouldApplyMutationCellState({ value: null, format: null, authoredBlank: false }, { value: null, format: null, authoredBlank: true }),
    ).toBe(false)
    expect(
      shouldApplyMutationCellState(
        { value: null, format: null, authoredBlank: false },
        { value: null, format: 'fmt-1', authoredBlank: true },
      ),
    ).toBe(true)
    expect(shouldApplyMutationCellState({ value: 'old', format: null }, { value: 'new', format: null })).toBe(true)
  })
})

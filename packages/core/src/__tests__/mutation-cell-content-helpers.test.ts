import { describe, expect, it } from 'vitest'
import type { EngineOp } from '@bilig/workbook'
import type { CellSnapshot } from '@bilig/protocol'
import {
  collectLiveCreatedSheetNames,
  getMutationMatrixCell,
  inverseMutationStructuralInsertOp,
  isMutationStructuralInsertOp,
  translateMutationFormulaForTarget,
} from '../engine/services/mutation-cell-content-helpers.js'

const cell = (value: string): CellSnapshot => ({ value, formula: null, format: null })

describe('mutation cell content helpers', () => {
  it('reads matrix cells with precise missing row and cell errors', () => {
    const matrix = [[cell('A1')]]

    expect(getMutationMatrixCell(matrix, 0, 0)).toEqual(cell('A1'))
    expect(() => getMutationMatrixCell(matrix, 1, 0)).toThrow('Missing source row at index 1')
    expect(() => getMutationMatrixCell(matrix, 0, 1)).toThrow('Missing source cell at row 0, column 1')
  })

  it('identifies structural insert operations and builds inverse deletes', () => {
    const insertRows: EngineOp = { kind: 'insertRows', sheetName: 'Sheet1', start: 2, count: 3 }
    const insertColumns: EngineOp = { kind: 'insertColumns', sheetName: 'Sheet1', start: 4, count: 5 }
    const deleteRows: EngineOp = { kind: 'deleteRows', sheetName: 'Sheet1', start: 2, count: 3 }

    expect(isMutationStructuralInsertOp(insertRows)).toBe(true)
    expect(isMutationStructuralInsertOp(insertColumns)).toBe(true)
    expect(isMutationStructuralInsertOp(deleteRows)).toBe(false)
    expect(inverseMutationStructuralInsertOp(insertRows)).toEqual({ kind: 'deleteRows', sheetName: 'Sheet1', start: 2, count: 3 })
    expect(inverseMutationStructuralInsertOp(insertColumns)).toEqual({
      kind: 'deleteColumns',
      sheetName: 'Sheet1',
      start: 4,
      count: 5,
    })
  })

  it('tracks newly created sheet names through same-batch renames', () => {
    const liveCreated = collectLiveCreatedSheetNames(
      ['Existing'],
      [
        { kind: 'upsertSheet', name: 'Temp' },
        { kind: 'renameSheet', oldName: 'Temp', newName: 'Final' },
        { kind: 'upsertSheet', name: 'Existing' },
        { kind: 'renameSheet', oldName: 'Existing', newName: 'RenamedExisting' },
      ],
    )

    expect([...liveCreated]).toEqual(['Final'])
  })

  it('translates formulas only when source and target sheets match', () => {
    expect(translateMutationFormulaForTarget('=A1+$B$2', 'Sheet1', 'A1', 'Sheet1', 'C3')).toBe('C3+$B$2')
    expect(translateMutationFormulaForTarget('=A1', 'Sheet1', 'A1', 'Other', 'C3')).toBe('=A1')
  })
})

import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import {
  dependencyTouchesStructuralDeleteSpan,
  structuralDeletedCellUndoRecordToOps,
  structuralFormulaUndoRecordToOp,
} from '../engine/services/mutation-structural-undo-records.js'

const SNAPSHOT: CellSnapshot = {
  sheetName: 'Sheet1',
  address: 'C4',
  value: { tag: ValueTag.String, value: 'kept' },
  flags: 0,
  version: 1,
}

describe('mutation structural undo records', () => {
  it('detects dependencies that need structural delete undo protection', () => {
    expect(dependencyTouchesStructuralDeleteSpan('A5', 'Sheet1', 'Sheet1', 'row', 4)).toBe(true)
    expect(dependencyTouchesStructuralDeleteSpan('A3', 'Sheet1', 'Sheet1', 'row', 4)).toBe(false)
    expect(dependencyTouchesStructuralDeleteSpan('A1:C10', 'Sheet1', 'Sheet1', 'column', 2)).toBe(true)
    expect(dependencyTouchesStructuralDeleteSpan('Other!A5', 'Sheet1', 'Sheet1', 'row', 4)).toBe(false)
  })

  it('converts formula undo records into engine ops', () => {
    expect(structuralFormulaUndoRecordToOp({ sheetName: 'Sheet1', row: 2, col: 3, formula: 'A1+B1' })).toEqual({
      kind: 'setCellFormula',
      sheetName: 'Sheet1',
      address: 'D3',
      formula: 'A1+B1',
    })
  })

  it('restores deleted formulas, values, blanks, and snapshots', () => {
    const toCellStateOps = (sheetName: string, address: string, snapshot: CellSnapshot) => [
      { kind: 'setCellValue' as const, sheetName, address, value: snapshot.value.tag === ValueTag.String ? snapshot.value.value : null },
    ]

    expect(
      structuralDeletedCellUndoRecordToOps(
        { kind: 'formula', sheetName: 'Sheet1', row: 0, col: 0, formula: 'B1', explicitFormat: '0.00' },
        toCellStateOps,
      ),
    ).toEqual([
      { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'A1', formula: 'B1' },
      { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'A1', format: '0.00' },
    ])
    expect(
      structuralDeletedCellUndoRecordToOps({ kind: 'value', sheetName: 'Sheet1', row: 1, col: 1, value: true }, toCellStateOps),
    ).toEqual([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'B2', value: true }])
    expect(
      structuralDeletedCellUndoRecordToOps(
        { kind: 'blank', sheetName: 'Sheet1', row: 2, col: 2, restoreExplicitBlank: true, explicitFormat: '@' },
        toCellStateOps,
      ),
    ).toEqual([
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'C3', value: null },
      { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'C3', format: '@' },
    ])
    expect(
      structuralDeletedCellUndoRecordToOps({ kind: 'snapshot', sheetName: 'Sheet1', row: 3, col: 2, snapshot: SNAPSHOT }, toCellStateOps),
    ).toEqual([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'C4', value: 'kept' }])
  })
})

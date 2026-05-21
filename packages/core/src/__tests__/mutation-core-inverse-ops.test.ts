import { describe, expect, it } from 'vitest'
import type { EngineOp } from '@bilig/workbook'
import { createMutationCoreInverseOps } from '../engine/services/mutation-core-inverse-ops.js'
import { WorkbookStore } from '../workbook-store.js'

function createWorkbookWithSheet(): WorkbookStore {
  const workbook = new WorkbookStore('core-inverse')
  workbook.createSheet('Sheet1')
  return workbook
}

describe('mutation core inverse ops', () => {
  it('captures deleted sheet objects before sheet cell state', () => {
    const workbook = createWorkbookWithSheet()
    workbook.setFreezePane('Sheet1', 1, 0)
    workbook.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B4',
      columnNames: ['Region', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    const cellStateOp: EngineOp = { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 12 }
    const inverse = createMutationCoreInverseOps({
      workbook,
      captureSheetCellState: (sheetName) => (sheetName === 'Sheet1' ? [cellStateOp] : []),
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => [],
      captureFormulaCellStateForStructuralUndo: () => [],
    })

    const ops = inverse.inverseOpsFor({ kind: 'deleteSheet', name: 'Sheet1' })

    expect(ops).toEqual([
      { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
      { kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1, cols: 0 },
      {
        kind: 'upsertTable',
        table: {
          name: 'Sales',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B4',
          columnNames: ['Region', 'Amount'],
          headerRow: true,
          totalsRow: false,
        },
      },
      cellStateOp,
    ])
  })

  it('captures structural row delete state with formula undo after range cells', () => {
    const workbook = createWorkbookWithSheet()
    workbook.setRowMetadata('Sheet1', 1, 1, 24, false)
    const rowCellOp: EngineOp = { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A2', value: 7 }
    const formulaOp: EngineOp = { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'B3', formula: 'SUM(A1:A3)' }
    const inverse = createMutationCoreInverseOps({
      workbook,
      captureSheetCellState: () => [],
      captureRowRangeCellState: (sheetName, start, count) => (sheetName === 'Sheet1' && start === 1 && count === 1 ? [rowCellOp] : []),
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => [],
      captureFormulaCellStateForStructuralUndo: (sheetName, axis, start, count) =>
        sheetName === 'Sheet1' && axis === 'row' && start === 1 && count === 1 ? [formulaOp] : [],
    })

    const ops = inverse.inverseOpsFor({ kind: 'deleteRows', sheetName: 'Sheet1', start: 1, count: 1 })

    expect(ops).toContainEqual({
      kind: 'insertRows',
      sheetName: 'Sheet1',
      start: 1,
      count: 1,
      entries: [{ id: 'row-1', index: 1, size: 24, hidden: false }],
    })
    expect(ops.indexOf(rowCellOp)).toBeLessThan(ops.indexOf(formulaOp))
  })

  it('restores cell ops in reverse transaction order and rejects unsupported ops', () => {
    const workbook = createWorkbookWithSheet()
    const inverse = createMutationCoreInverseOps({
      workbook,
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: (sheetName, address) => [{ kind: 'setCellValue', sheetName, address, value: `old-${address}` }],
      captureFormulaCellStateForStructuralUndo: () => [],
    })

    expect(
      inverse.buildInverseOps([
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 },
        { kind: 'clearCell', sheetName: 'Sheet1', address: 'B2' },
      ]),
    ).toEqual([
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B2', value: 'old-B2' },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'old-A1' },
    ])
    expect(() => {
      Reflect.apply(inverse.inverseOpsFor, inverse, [{ kind: 'unsupportedMutation' }])
    }).toThrow('Unhandled metadata inverse operation: unsupportedMutation')
  })

  it('preserves table-header rename suppression on restored value and clear inverses', () => {
    const workbook = createWorkbookWithSheet()
    const inverse = createMutationCoreInverseOps({
      workbook,
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: (sheetName, address) => [
        { kind: 'setCellValue', sheetName, address, value: `old-${address}` },
        { kind: 'setCellFormula', sheetName, address: 'C3', formula: 'A1' },
      ],
      captureFormulaCellStateForStructuralUndo: () => [],
    })

    expect(
      inverse.inverseOpsFor({ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1, skipTableHeaderRename: true }),
    ).toEqual([
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'old-A1', skipTableHeaderRename: true },
      { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'C3', formula: 'A1' },
    ])
    expect(inverse.inverseOpsFor({ kind: 'clearCell', sheetName: 'Sheet1', address: 'B2', skipTableHeaderRename: true })[0]).toEqual({
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: 'B2',
      value: 'old-B2',
      skipTableHeaderRename: true,
    })
    expect(inverse.inverseOpsFor({ kind: 'setCellFormula', sheetName: 'Sheet1', address: 'D4', formula: 'A1' })[0]).toEqual({
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: 'D4',
      value: 'old-D4',
    })
  })
})

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

  it('captures deleted sheet spill, pivot, chart, image, and shape metadata', () => {
    const workbook = createWorkbookWithSheet()
    workbook.createSheet('Data')
    workbook.setSpill('Sheet1', 'D4', 2, 3)
    workbook.setPivot({
      name: 'RevenuePivot',
      sheetName: 'Sheet1',
      address: 'G2',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B5' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales' }],
      rows: 4,
      cols: 3,
    })
    workbook.setChart({
      id: 'chart-1',
      sheetName: 'Sheet1',
      address: 'J2',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B5' },
      chartType: 'line',
      rows: 5,
      cols: 6,
    })
    workbook.setImage({
      id: 'image-1',
      sheetName: 'Sheet1',
      address: 'J9',
      sourceUrl: 'https://example.com/image.png',
      rows: 3,
      cols: 4,
      altText: 'Revenue image',
    })
    workbook.setShape({
      id: 'shape-1',
      sheetName: 'Sheet1',
      address: 'N9',
      shapeType: 'rectangle',
      rows: 2,
      cols: 5,
      text: 'Callout',
    })
    const cellStateOp: EngineOp = { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'cell-after-metadata' }
    const inverse = createMutationCoreInverseOps({
      workbook,
      captureSheetCellState: (sheetName) => (sheetName === 'Sheet1' ? [cellStateOp] : []),
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => [],
      captureFormulaCellStateForStructuralUndo: () => [],
    })

    const ops = inverse.inverseOpsFor({ kind: 'deleteSheet', name: 'Sheet1' })

    expect(ops).toContainEqual({ kind: 'upsertSpillRange', sheetName: 'Sheet1', address: 'D4', rows: 2, cols: 3 })
    expect(ops).toContainEqual({
      kind: 'upsertPivotTable',
      name: 'RevenuePivot',
      sheetName: 'Sheet1',
      address: 'G2',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B5' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales' }],
      rows: 4,
      cols: 3,
    })
    expect(ops).toContainEqual({
      kind: 'upsertChart',
      chart: {
        id: 'chart-1',
        sheetName: 'Sheet1',
        address: 'J2',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B5' },
        chartType: 'line',
        rows: 5,
        cols: 6,
      },
    })
    expect(ops).toContainEqual({
      kind: 'upsertImage',
      image: {
        id: 'image-1',
        sheetName: 'Sheet1',
        address: 'J9',
        sourceUrl: 'https://example.com/image.png',
        rows: 3,
        cols: 4,
        altText: 'Revenue image',
      },
    })
    expect(ops).toContainEqual({
      kind: 'upsertShape',
      shape: {
        id: 'shape-1',
        sheetName: 'Sheet1',
        address: 'N9',
        shapeType: 'rectangle',
        rows: 2,
        cols: 5,
        text: 'Callout',
      },
    })
    expect(ops.at(-1)).toBe(cellStateOp)
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

  it('restores inconsistent table header metadata before restoring the cell value', () => {
    const workbook = createWorkbookWithSheet()
    workbook.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Region', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    const inverse = createMutationCoreInverseOps({
      workbook,
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: (sheetName, address) => [{ kind: 'setCellValue', sheetName, address, value: 'East' }],
      captureFormulaCellStateForStructuralUndo: () => [],
    })

    expect(inverse.inverseOpsFor({ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' })).toEqual([
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'Region' },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'East', skipTableHeaderRename: true },
    ])
  })

  it('uses regular header restore when the restored value already matches table metadata', () => {
    const workbook = createWorkbookWithSheet()
    workbook.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Region', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    const inverse = createMutationCoreInverseOps({
      workbook,
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: (sheetName, address) => [{ kind: 'setCellValue', sheetName, address, value: 'Region' }],
      captureFormulaCellStateForStructuralUndo: () => [],
    })

    expect(inverse.inverseOpsFor({ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'North' })).toEqual([
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'Region' },
    ])
  })

  it('restores cleared table headers through metadata-safe clear inverses', () => {
    const workbook = createWorkbookWithSheet()
    workbook.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Region', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    const inverse = createMutationCoreInverseOps({
      workbook,
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: (sheetName, address) => [{ kind: 'clearCell', sheetName, address }],
      captureFormulaCellStateForStructuralUndo: () => [],
    })

    expect(inverse.inverseOpsFor({ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'North' })).toEqual([
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'Region' },
      { kind: 'clearCell', sheetName: 'Sheet1', address: 'A1', skipTableHeaderRename: true },
    ])
  })
})

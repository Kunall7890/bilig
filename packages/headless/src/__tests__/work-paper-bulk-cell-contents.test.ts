import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import {
  WorkPaper,
  type WorkPaperCellAddress,
  type WorkPaperCellValueUpdate,
  type WorkPaperSheetCellValueUpdate,
  type WorkPaperSheetRangeValues,
} from '../index.js'
import { hasDeferredTrackedIndexChanges } from '../tracked-cell-index-changes.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function buildTwoInputFormulaRows(rowCount: number): (number | string)[][] {
  return Array.from({ length: rowCount }, (_, row) => {
    const rowNumber = row + 1
    return [rowNumber, rowNumber * 2, `=A${rowNumber}+B${rowNumber}`, `=A${rowNumber}*B${rowNumber}`]
  })
}

function buildTwoColumnUpdates(sheetId: number, rowCount: number): WorkPaperCellValueUpdate[] {
  return Array.from({ length: rowCount * 2 }, (_value, index) => {
    const row = Math.floor(index / 2)
    const col = index % 2
    return {
      address: cell(sheetId, row, col),
      value: col === 0 ? row * 3 : row * 5,
    }
  })
}

function buildTwoColumnSheetUpdates(rowCount: number): WorkPaperSheetCellValueUpdate[] {
  return Array.from({ length: rowCount * 2 }, (_value, index) => {
    const row = Math.floor(index / 2)
    const col = index % 2
    return {
      row,
      col,
      value: col === 0 ? row * 3 : row * 5,
    }
  })
}

function buildTwoColumnRangeValues(rowCount: number): WorkPaperSheetRangeValues {
  return Array.from({ length: rowCount }, (_value, row) => [row * 3, row * 5])
}

function buildDenseNumericRange(rowCount: number, colCount: number, offset = 0): WorkPaperSheetRangeValues {
  return Array.from({ length: rowCount }, (_rowValue, row) =>
    Array.from({ length: colCount }, (_colValue, col) => (row + 1) * (col + 2) + offset),
  )
}

describe('bulk cell values', () => {
  it('applies sparse updates through one tracked engine mutation', () => {
    const rowCount = 160
    const workbook = WorkPaper.buildFromSheets({ Bench: buildTwoInputFormulaRows(rowCount) })
    const sheetId = workbook.getSheetId('Bench')!
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('bulk sparse updates should not rebuild visibility snapshots')
    })

    try {
      const changes = workbook.setCellValues(buildTwoColumnUpdates(sheetId, rowCount))

      expect(changes).toHaveLength(rowCount * 4)
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 2))).toEqual({
        tag: ValueTag.Number,
        value: (rowCount - 1) * 8,
      })
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 3))).toEqual({
        tag: ValueTag.Number,
        value: (rowCount - 1) * 3 * ((rowCount - 1) * 5),
      })
      expect(workbook.getPerformanceCounters()).toMatchObject({
        directScalarDeltaApplications: rowCount * 2,
        directScalarDeltaOnlyRecalcSkips: 1,
      })
    } finally {
      captureVisibilitySnapshot.mockRestore()
    }
  })

  it('applies same-sheet sparse updates without per-update sheet addresses', () => {
    const rowCount = 160
    const workbook = WorkPaper.buildFromSheets({ Bench: buildTwoInputFormulaRows(rowCount) })
    const sheetId = workbook.getSheetId('Bench')!
    const changes = workbook.setSheetCellValues(sheetId, buildTwoColumnSheetUpdates(rowCount))

    expect(changes).toHaveLength(rowCount * 4)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
    expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 2))).toEqual({
      tag: ValueTag.Number,
      value: (rowCount - 1) * 8,
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directScalarDeltaApplications: rowCount * 2,
      directScalarDeltaOnlyRecalcSkips: 1,
    })
  })

  it('applies dense same-sheet range values without caller-allocated update objects', () => {
    const rowCount = 160
    const workbook = WorkPaper.buildFromSheets({ Bench: buildTwoInputFormulaRows(rowCount) })
    const sheetId = workbook.getSheetId('Bench')!
    const changes = workbook.setSheetRangeValues(sheetId, 0, 0, buildTwoColumnRangeValues(rowCount))

    expect(changes).toHaveLength(rowCount * 4)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
    expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 2))).toEqual({
      tag: ValueTag.Number,
      value: (rowCount - 1) * 8,
    })
    expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 3))).toEqual({
      tag: ValueTag.Number,
      value: (rowCount - 1) * 3 * ((rowCount - 1) * 5),
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directScalarDeltaApplications: rowCount * 2,
      directScalarDeltaOnlyRecalcSkips: 1,
    })
  })

  it('applies fresh dense numeric range values through one rectangular kernel-sync path', () => {
    const rowCount = 80
    const colCount = 8
    const values: WorkPaperSheetRangeValues = Array.from({ length: rowCount }, (_rowValue, row) =>
      Array.from({ length: colCount }, (_colValue, col) => (row + 1) * (col + 2)),
    )
    const workbook = WorkPaper.buildFromSheets({ Bench: [] })
    const sheetId = workbook.getSheetId('Bench')!
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('fresh dense range values should not rebuild visibility snapshots')
    })

    try {
      workbook.resetPerformanceCounters()
      const changes = workbook.setSheetRangeValues(sheetId, 0, 0, values)

      expect(changes).toHaveLength(rowCount * colCount)
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, colCount - 1))).toEqual({
        tag: ValueTag.Number,
        value: rowCount * (colCount + 1),
      })
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ height: rowCount, width: colCount })
      expect(workbook.getPerformanceCounters()).toMatchObject({
        changedCellPayloadsBuilt: 0,
        kernelSyncOnlyRecalcSkips: 1,
        topoRebuilds: 0,
      })
    } finally {
      captureVisibilitySnapshot.mockRestore()
    }
  })

  it('overwrites existing dense numeric range values through one rectangular kernel-sync path', () => {
    const rowCount = 80
    const colCount = 8
    const workbook = WorkPaper.buildFromSheets({ Bench: buildDenseNumericRange(rowCount, colCount) })
    const sheetId = workbook.getSheetId('Bench')!
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('existing dense range values should not rebuild visibility snapshots')
    })

    try {
      workbook.resetPerformanceCounters()
      const changes = workbook.setSheetRangeValues(sheetId, 0, 0, buildDenseNumericRange(rowCount, colCount, 10_000))

      expect(changes).toHaveLength(rowCount * colCount)
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, colCount - 1))).toEqual({
        tag: ValueTag.Number,
        value: rowCount * (colCount + 1) + 10_000,
      })
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ height: rowCount, width: colCount })
      expect(workbook.getPerformanceCounters()).toMatchObject({
        changedCellPayloadsBuilt: 0,
        kernelSyncOnlyRecalcSkips: 1,
        topoRebuilds: 0,
      })

      const undoChanges = workbook.undo()
      expect(undoChanges).toHaveLength(rowCount * colCount)
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, colCount - 1))).toEqual({
        tag: ValueTag.Number,
        value: rowCount * (colCount + 1),
      })
    } finally {
      captureVisibilitySnapshot.mockRestore()
    }
  })

  it('applies null clears and ragged range value rows in one public call', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1'],
        [3, 4, '=A2+B2'],
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!

    workbook.setSheetRangeValues(sheetId, 0, 0, [[null, 5], [7]])

    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toEqual({ tag: ValueTag.Empty })
    expect(workbook.getCellValue(cell(sheetId, 0, 2))).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(workbook.getCellValue(cell(sheetId, 1, 1))).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(workbook.getCellValue(cell(sheetId, 1, 2))).toEqual({ tag: ValueTag.Number, value: 11 })
  })

  it('defers sparse updates while evaluation is suspended and flushes once on resume', () => {
    const rowCount = 96
    const workbook = WorkPaper.buildFromSheets({ Bench: buildTwoInputFormulaRows(rowCount) })
    const sheetId = workbook.getSheetId('Bench')!

    workbook.suspendEvaluation()
    expect(workbook.setCellValues(buildTwoColumnUpdates(sheetId, rowCount))).toEqual([])
    const changes = workbook.resumeEvaluation()

    expect(changes).toHaveLength(rowCount * 4)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
    expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 2))).toEqual({
      tag: ValueTag.Number,
      value: (rowCount - 1) * 8,
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directScalarDeltaApplications: rowCount * 2,
      directScalarDeltaOnlyRecalcSkips: 1,
    })
  })

  it('keeps last-write-wins sparse value patches in one public call', () => {
    const workbook = WorkPaper.buildFromSheets({ Bench: [[1, 2, '=A1+B1']] })
    const sheetId = workbook.getSheetId('Bench')!

    const changes = workbook.setCellValues([
      { address: cell(sheetId, 0, 0), value: 4 },
      { address: cell(sheetId, 0, 0), value: 5 },
      { address: cell(sheetId, 0, 1), value: 7 },
    ])

    expect(changes.map((change) => (change.kind === 'cell' ? change.a1 : change.name))).toEqual(['A1', 'B1', 'C1'])
    expect(workbook.getCellValue(cell(sheetId, 0, 2))).toEqual({ tag: ValueTag.Number, value: 12 })
  })

  it('applies addressed sparse value patches across sheets in one public call', () => {
    const workbook = WorkPaper.buildFromSheets({
      First: [[1, '=A1*2']],
      Second: [[2, '=A1*3']],
    })
    const firstSheet = workbook.getSheetId('First')!
    const secondSheet = workbook.getSheetId('Second')!

    const changes = workbook.setCellValues([
      { address: cell(firstSheet, 0, 0), value: 5 },
      { address: cell(secondSheet, 0, 0), value: 7 },
      { address: cell(firstSheet, 0, 0), value: 11 },
    ])

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : change.name))).toEqual([
      'First!A1',
      'First!B1',
      'Second!A1',
      'Second!B1',
    ])
    expect(workbook.getCellValue(cell(firstSheet, 0, 1))).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(workbook.getCellValue(cell(secondSheet, 0, 1))).toEqual({ tag: ValueTag.Number, value: 21 })
  })

  it('rejects formula strings so formula rewrites keep using the semantic single-cell path', () => {
    const workbook = WorkPaper.buildFromSheets({ Bench: [[1, 2, '=A1+B1']] })
    const sheetId = workbook.getSheetId('Bench')!

    expect(() => {
      workbook.setCellValues([{ address: cell(sheetId, 0, 2), value: '=A1*B1' }])
    }).toThrow('Bulk cell value updates require literal values')
    expect(() => {
      workbook.setSheetCellValues(sheetId, [{ row: 0, col: 2, value: '=A1*B1' }])
    }).toThrow('Bulk cell value updates require literal values')
    expect(() => {
      workbook.setSheetRangeValues(sheetId, 0, 0, [[1, '=A1*B1']])
    }).toThrow('Bulk cell value updates require literal values')
  })
})

import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import {
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  WorkPaper,
  type RawCellContent,
} from '../index.js'

const sheetName = 'Sheet1'

describe('WorkPaper/core parity regressions', () => {
  it('recalculates overlapping direct aggregates after column insert and text replacement across save/load', () => {
    const { workbook, sheetId } = createSeededWorkbook()

    workbook.addColumns(sheetId, 0, 1)
    setCell(workbook, sheetId, 0, 2, '=SUM(A1:B2)')

    expect(workbook.getCellValue({ sheet: sheetId, row: 5, col: 1 })).toEqual({ tag: ValueTag.Number, value: 105 })
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(workbook.getCellValue({ sheet: sheetId, row: 5, col: 2 })).toEqual({ tag: ValueTag.Number, value: 13 })

    setCell(workbook, sheetId, 1, 1, 'north')

    expect(workbook.getCellValue({ sheet: sheetId, row: 5, col: 1 })).toEqual({ tag: ValueTag.Number, value: 94 })
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(workbook.getCellValue({ sheet: sheetId, row: 5, col: 2 })).toEqual({ tag: ValueTag.Number, value: 2 })

    const restored = createWorkPaperFromDocument(
      parseWorkPaperDocument(serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))),
    )
    const restoredSheetId = restored.getSheetId(sheetName)
    expect(restoredSheetId).not.toBeUndefined()
    expect(restored.getCellValue({ sheet: restoredSheetId!, row: 5, col: 1 })).toEqual({ tag: ValueTag.Number, value: 94 })
    expect(restored.getCellValue({ sheet: restoredSheetId!, row: 0, col: 2 })).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(restored.getCellValue({ sheet: restoredSheetId!, row: 5, col: 2 })).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('recalculates expanded direct aggregates after row moves', () => {
    const { workbook, sheetId } = createSeededWorkbook()

    setCell(workbook, sheetId, 0, 2, '=SUM(A1:B2)')
    workbook.moveRows(sheetId, 2, 1, 1)

    expect(workbook.getCellSerialized({ sheet: sheetId, row: 0, col: 2 })).toBe('=SUM(A1:B3)')
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toEqual({ tag: ValueTag.Number, value: 69 })
  })

  it('recalculates dependents of precomputed aggregate values after row deletes', () => {
    const { workbook, sheetId } = createSeededWorkbook()

    workbook.moveRows(sheetId, 5, 1, 0)
    setCell(workbook, sheetId, 0, 1, '=IF(A1>0,A1,0)')
    workbook.removeRows(sheetId, 1, 1)

    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 1 })).toEqual({ tag: ValueTag.Number, value: 104 })
  })

  it('restores formula bindings when undoing column inserts', () => {
    const { workbook, sheetId } = createSeededWorkbook()

    setCell(workbook, sheetId, 0, 2, '=A1&"-"&B1')
    workbook.addColumns(sheetId, 0, 1)
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 3 })).toMatchObject({ tag: ValueTag.String, value: '1-2' })

    workbook.undo()
    expect(workbook.getCellSerialized({ sheet: sheetId, row: 0, col: 2 })).toBe('=A1&"-"&B1')
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toMatchObject({ tag: ValueTag.String, value: '1-2' })

    setCell(workbook, sheetId, 0, 0, 9)
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toMatchObject({ tag: ValueTag.String, value: '9-2' })
    setCell(workbook, sheetId, 0, 1, 8)
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toMatchObject({ tag: ValueTag.String, value: '9-8' })
  })

  it('recalculates scalar formulas when a shared aggregate input changes before column insert', () => {
    const { workbook, sheetId } = createSeededWorkbook()

    setCell(workbook, sheetId, 0, 0, 0)

    expect(workbook.getCellValue({ sheet: sheetId, row: 5, col: 0 })).toEqual({ tag: ValueTag.Number, value: 104 })
    expect(workbook.getCellValue({ sheet: sheetId, row: 5, col: 1 })).toEqual({ tag: ValueTag.Number, value: 2 })

    workbook.addColumns(sheetId, 0, 1)

    expect(workbook.getCellSerialized({ sheet: sheetId, row: 5, col: 1 })).toBe('=SUM(B1:B5)')
    expect(workbook.getCellValue({ sheet: sheetId, row: 5, col: 1 })).toEqual({ tag: ValueTag.Number, value: 104 })
    expect(workbook.getCellSerialized({ sheet: sheetId, row: 5, col: 2 })).toBe('=B1+C1')
    expect(workbook.getCellValue({ sheet: sheetId, row: 5, col: 2 })).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('recalculates scalar formulas when a shared aggregate input changes from number to boolean', () => {
    const { workbook, sheetId } = createSeededWorkbook()

    setCell(workbook, sheetId, 0, 2, '=SUM(A1:B2)')
    setCell(workbook, sheetId, 0, 0, false)

    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toEqual({ tag: ValueTag.Number, value: 25 })
    expect(workbook.getCellValue({ sheet: sheetId, row: 5, col: 0 })).toEqual({ tag: ValueTag.Number, value: 104 })
    expect(workbook.getCellSerialized({ sheet: sheetId, row: 5, col: 1 })).toBe('=A1+B1')
    expect(workbook.getCellValue({ sheet: sheetId, row: 5, col: 1 })).toEqual({ tag: ValueTag.Number, value: 2 })

    const restored = createWorkPaperFromDocument(
      parseWorkPaperDocument(serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))),
    )
    const restoredSheetId = restored.getSheetId(sheetName)
    expect(restoredSheetId).not.toBeUndefined()
    expect(restored.getCellSerialized({ sheet: restoredSheetId!, row: 5, col: 1 })).toBe('=A1+B1')
    expect(restored.getCellValue({ sheet: restoredSheetId!, row: 5, col: 1 })).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('preserves structural reference errors over induced cycles across save and load', () => {
    const { workbook, sheetId } = createSeededWorkbook()

    setCell(workbook, sheetId, 0, 1, '=A1+B1')
    workbook.removeColumns(sheetId, 0, 1)

    expect(workbook.getCellSerialized({ sheet: sheetId, row: 0, col: 0 })).toBe('=#REF!+A1')
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 0 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })

    const restored = createWorkPaperFromDocument(
      parseWorkPaperDocument(serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))),
    )
    const restoredSheetId = restored.getSheetId(sheetName)
    expect(restoredSheetId).not.toBeUndefined()
    expect(restored.getCellSerialized({ sheet: restoredSheetId!, row: 0, col: 0 })).toBe('=#REF!+A1')
    expect(restored.getCellValue({ sheet: restoredSheetId!, row: 0, col: 0 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })

  it('recalculates direct scalar rewrites before save/load when dependencies change', () => {
    const { workbook, sheetId } = createSeededWorkbook()

    workbook.removeRows(sheetId, 1, 1)
    setCell(workbook, sheetId, 5, 1, '=A1+B1')
    workbook.addRows(sheetId, 0, 1)

    expect(workbook.getCellSerialized({ sheet: sheetId, row: 6, col: 1 })).toBe('=A3+B3')
    expect(workbook.getCellValue({ sheet: sheetId, row: 6, col: 1 })).toEqual({ tag: ValueTag.Number, value: 43 })

    const restored = createWorkPaperFromDocument(
      parseWorkPaperDocument(serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))),
    )
    const restoredSheetId = restored.getSheetId(sheetName)
    expect(restoredSheetId).not.toBeUndefined()
    expect(restored.getCellSerialized({ sheet: restoredSheetId!, row: 6, col: 1 })).toBe('=A3+B3')
    expect(restored.getCellValue({ sheet: restoredSheetId!, row: 6, col: 1 })).toEqual({ tag: ValueTag.Number, value: 43 })
  })

  it('replays direct scalar insert rewrites through redo history', () => {
    const { workbook, sheetId } = createSeededWorkbook()

    workbook.removeRows(sheetId, 1, 1)
    setCell(workbook, sheetId, 5, 1, '=A1+B1')
    workbook.addRows(sheetId, 0, 1)

    expect(workbook.getCellSerialized({ sheet: sheetId, row: 6, col: 1 })).toBe('=A3+B3')
    expect(workbook.getCellValue({ sheet: sheetId, row: 6, col: 1 })).toEqual({ tag: ValueTag.Number, value: 43 })

    workbook.undo()
    expect(workbook.getCellSerialized({ sheet: sheetId, row: 5, col: 1 })).toBe('=A1+B1')
    expect(workbook.getCellValue({ sheet: sheetId, row: 5, col: 1 })).toEqual({ tag: ValueTag.Number, value: 3 })

    workbook.redo()
    expect(workbook.getCellSerialized({ sheet: sheetId, row: 6, col: 1 })).toBe('=A3+B3')
    expect(workbook.getCellValue({ sheet: sheetId, row: 6, col: 1 })).toEqual({ tag: ValueTag.Number, value: 43 })
  })
})

function createSeededWorkbook(): { workbook: WorkPaper; sheetId: number } {
  const workbook = WorkPaper.buildEmpty({
    parseDateTime: () => undefined,
    functionPlugins: [],
  })
  workbook.addSheet(sheetName)
  const sheetId = workbook.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error('Expected sheet to exist')
  }

  for (let row = 0; row < 5; row += 1) {
    for (let col = 0; col < 5; col += 1) {
      setCell(workbook, sheetId, row, col, row * 10 + col + 1)
    }
  }
  setCell(workbook, sheetId, 5, 0, '=SUM(A1:A5)')
  setCell(workbook, sheetId, 5, 1, '=A1+B1')
  return { workbook, sheetId }
}

function setCell(workbook: WorkPaper, sheetId: number, row: number, col: number, value: RawCellContent): void {
  workbook.setCellContents({ sheet: sheetId, row, col }, value)
}

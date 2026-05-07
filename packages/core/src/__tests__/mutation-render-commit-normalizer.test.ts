import { describe, expect, it } from 'vitest'
import { normalizeRenderCommitOps } from '../engine/services/mutation-render-commit-normalizer.js'

describe('mutation render commit normalizer', () => {
  it('normalizes workbook, sheet, and cell commit records into engine ops', () => {
    expect(
      normalizeRenderCommitOps([
        { kind: 'upsertWorkbook', name: 'Book' },
        { kind: 'upsertSheet', name: 'Sheet1' },
        { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'B3', value: 7, format: '0.00' },
        { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'C4', formula: 'B3*2' },
        { kind: 'deleteCell', sheetName: 'Sheet1', addr: 'D5' },
      ]),
    ).toEqual({
      engineOps: [
        { kind: 'upsertWorkbook', name: 'Book' },
        { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B3', value: 7 },
        { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'B3', format: '0.00' },
        { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'C4', formula: 'B3*2' },
        { kind: 'clearCell', sheetName: 'Sheet1', address: 'D5' },
        { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'D5', format: null },
      ],
      potentialNewCells: 2,
      preparedCellAddressesByOpIndex: [null, null, { row: 2, col: 1 }, null, { row: 3, col: 2 }, { row: 4, col: 3 }, null],
    })
  })

  it('drops malformed partial commit records', () => {
    expect(
      normalizeRenderCommitOps([
        { kind: 'upsertWorkbook' },
        { kind: 'renameSheet', oldName: 'OnlyOld' },
        { kind: 'deleteSheet' },
        { kind: 'upsertCell', sheetName: 'Sheet1' },
        { kind: 'deleteCell', addr: 'A1' },
      ]),
    ).toEqual({ engineOps: [], potentialNewCells: 0, preparedCellAddressesByOpIndex: [] })
  })
})

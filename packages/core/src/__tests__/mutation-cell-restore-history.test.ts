import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { writeLiteralToCellStore } from '../engine-value-utils.js'
import { createMutationCellRestoreHistoryHelpers, tryMutationCellRefsFromOps } from '../engine/services/mutation-cell-restore-history.js'
import { transactionRecordOps } from '../engine/services/mutation-transaction-records.js'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'

function createWorkbook() {
  const workbook = new WorkbookStore('restore-history')
  const sheet = workbook.createSheet('Sheet1')
  const strings = new StringPool()
  const getCellByIndex = (cellIndex: number): CellSnapshot => ({
    sheetName: workbook.getSheetNameById(workbook.cellStore.sheetIds[cellIndex]),
    address: workbook.getAddress(cellIndex),
    value: { tag: ValueTag.Empty },
    flags: workbook.cellStore.flags[cellIndex] ?? 0,
    version: workbook.cellStore.versions[cellIndex] ?? 0,
  })
  return { workbook, sheet, strings, getCellByIndex }
}

describe('mutation cell restore history helpers', () => {
  it('converts simple cell ops into mutation refs with existing cell indices', () => {
    const { workbook, sheet } = createWorkbook()
    const existingIndex = workbook.ensureCell('Sheet1', 'B2')

    expect(
      tryMutationCellRefsFromOps(workbook, [
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B2', value: 3 },
        { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'C3', formula: 'B2*2' },
        { kind: 'clearCell', sheetName: 'Sheet1', address: 'D4' },
      ]),
    ).toEqual([
      { sheetId: sheet.id, cellIndex: existingIndex, mutation: { kind: 'setCellValue', row: 1, col: 1, value: 3 } },
      { sheetId: sheet.id, mutation: { kind: 'setCellFormula', row: 2, col: 2, formula: 'B2*2' } },
      { sheetId: sheet.id, mutation: { kind: 'clearCell', row: 3, col: 3 } },
    ])
    expect(tryMutationCellRefsFromOps(workbook, [{ kind: 'upsertSheet', name: 'Sheet2', order: 1 }])).toBeNull()
    expect(tryMutationCellRefsFromOps(workbook, [{ kind: 'clearCell', sheetName: 'Missing', address: 'A1' }])).toBeNull()
  })

  it('captures inverse records for existing numeric cell mutations', () => {
    const { workbook, sheet, strings, getCellByIndex } = createWorkbook()
    const cellIndex = workbook.ensureCell('Sheet1', 'B2')
    writeLiteralToCellStore(workbook.cellStore, cellIndex, 7, strings)
    const helpers = createMutationCellRestoreHistoryHelpers({
      workbook,
      formulas: new Map(),
      getCellByIndex,
    })
    const ref = {
      sheetId: sheet.id,
      cellIndex,
      mutation: { kind: 'setCellValue' as const, row: 1, col: 1, value: 99 },
    }

    expect(helpers.restoreCellOpFromRef(ref)).toEqual({ kind: 'setCellValue', sheetName: 'Sheet1', address: 'B2', value: 7 })
    expect(helpers.tryRestoreSimpleCellOpFromStore('Sheet1', 'B2')).toEqual({
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: 'B2',
      value: 7,
    })

    const inverseRecord = helpers.tryCreateSingleExistingNumericInverseCellMutationRecord([ref])
    expect(inverseRecord ? transactionRecordOps(workbook, inverseRecord) : null).toEqual([
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B2', value: 7 },
    ])
    const history = helpers.buildFastMutationHistoryFromRefs([ref], 0)
    expect(history.undoOps).toEqual([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'B2', value: 7 }])
  })
})

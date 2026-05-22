import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook'
import { CellFlags } from '../cell-store.js'
import { tryBuildFastMutationHistory } from '../engine/services/mutation-history-fast-path.js'
import { WorkbookStore } from '../workbook-store.js'

function createWorkbookWithSheet(): WorkbookStore {
  const workbook = new WorkbookStore('fast-history')
  workbook.createSheet('Sheet1')
  return workbook
}

function numberSnapshot(sheetName: string, address: string, value: number): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Number, value },
    flags: 0,
    version: 1,
  }
}

function formulaSnapshot(sheetName: string, address: string, formula: string): CellSnapshot {
  return {
    sheetName,
    address,
    formula,
    value: { tag: ValueTag.Number, value: 42 },
    flags: 0,
    version: 1,
  }
}

function authoredBlankSnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: CellFlags.AuthoredBlank,
    version: 1,
  }
}

function errorSnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Error, code: ErrorCode.Value },
    flags: 0,
    version: 1,
  }
}

describe('mutation history fast path', () => {
  it('restores formulas, authored blanks, errors, and formats from indexed cell snapshots', () => {
    const workbook = createWorkbookWithSheet()
    const snapshots = new Map<number, CellSnapshot>()
    const a1 = workbook.ensureCell('Sheet1', 'A1')
    const b1 = workbook.ensureCell('Sheet1', 'B1')
    const c1 = workbook.ensureCell('Sheet1', 'C1')
    const d1 = workbook.ensureCell('Sheet1', 'D1')
    const e1 = workbook.ensureCell('Sheet1', 'E1')
    workbook.setCellFormat(d1, '#,##0')
    snapshots.set(a1, formulaSnapshot('Sheet1', 'A1', 'SUM(B1:B2)'))
    snapshots.set(b1, numberSnapshot('Sheet1', 'B1', 7))
    snapshots.set(c1, authoredBlankSnapshot('Sheet1', 'C1'))
    snapshots.set(e1, errorSnapshot('Sheet1', 'E1'))

    const history = tryBuildFastMutationHistory({
      workbook,
      getCellByIndex: (cellIndex) => snapshots.get(cellIndex)!,
      ops: [
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 99, skipTableHeaderRename: true },
        { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'B1', formula: 'A1+1' },
        { kind: 'clearCell', sheetName: 'Sheet1', address: 'C1', skipTableHeaderRename: true },
        { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'D1', format: '0.0' },
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'E1', value: 'replacement' },
      ],
      potentialNewCells: 5,
    })

    expect(history?.forward).toEqual({
      kind: 'ops',
      ops: [
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 99, skipTableHeaderRename: true },
        { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'B1', formula: 'A1+1' },
        { kind: 'clearCell', sheetName: 'Sheet1', address: 'C1', skipTableHeaderRename: true },
        { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'D1', format: '0.0' },
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'E1', value: 'replacement' },
      ],
      potentialNewCells: 5,
    })
    expect(history?.inverse).toEqual({
      kind: 'ops',
      ops: [
        { kind: 'clearCell', sheetName: 'Sheet1', address: 'E1' },
        { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'D1', format: '#,##0' },
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'C1', value: null, skipTableHeaderRename: true },
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 7 },
        { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'A1', formula: 'SUM(B1:B2)' },
      ],
      potentialNewCells: 5,
    })
    expect(history?.undoOps).toEqual(history?.inverse.ops)
  })

  it('skips cell inverses for sheets created inside the same fast-history transaction', () => {
    const workbook = createWorkbookWithSheet()
    const history = tryBuildFastMutationHistory({
      workbook,
      getCellByIndex: () => {
        throw new Error('created sheet cells should not be read')
      },
      ops: [
        { kind: 'upsertSheet', name: 'Created', order: 1 },
        { kind: 'setCellValue', sheetName: 'Created', address: 'A1', value: 1 },
      ],
    })

    expect(history?.inverse).toEqual({
      kind: 'ops',
      ops: [{ kind: 'deleteSheet', name: 'Created' }],
      potentialNewCells: 2,
    })
    expect(history?.undoOps).toEqual([{ kind: 'deleteSheet', name: 'Created' }])
  })

  it('supports workbook and rename inverses while preserving reused forward ops', () => {
    const workbook = createWorkbookWithSheet()
    workbook.workbookName = 'Original'
    const existing = workbook.createSheet('Existing', 3)
    const workbookOp = { kind: 'upsertWorkbook', name: 'Next' } satisfies EngineOp
    const renameOp = { kind: 'renameSheet', oldName: 'Previous', newName: existing.name } satisfies EngineOp
    const history = tryBuildFastMutationHistory({
      workbook,
      getCellByIndex: () => {
        throw new Error('metadata-only history should not read cells')
      },
      ops: [workbookOp, renameOp],
      cloneForwardOps: false,
      includeUndoOps: false,
    })

    expect(history?.forward).toEqual({ kind: 'ops', ops: [workbookOp, renameOp] })
    expect(history?.inverse).toEqual({
      kind: 'ops',
      ops: [
        { kind: 'renameSheet', oldName: 'Existing', newName: 'Previous' },
        { kind: 'upsertWorkbook', name: 'Original' },
      ],
      potentialNewCells: 2,
    })
    expect(history?.undoOps).toBeNull()
  })

  it('rejects unsupported, malformed, and misaligned fast-history inputs', () => {
    const workbook = createWorkbookWithSheet()

    expect(
      tryBuildFastMutationHistory({
        workbook,
        getCellByIndex: () => numberSnapshot('Sheet1', 'A1', 1),
        ops: [{ kind: 'deleteRows', sheetName: 'Sheet1', start: 0, count: 1 }],
      }),
    ).toBeNull()
    expect(
      tryBuildFastMutationHistory({
        workbook,
        getCellByIndex: () => numberSnapshot('Sheet1', 'A1', 1),
        ops: [{ kind: 'renameSheet', oldName: 'Sheet1', newName: 'Missing' }],
      }),
    ).toBeNull()
    expect(() =>
      tryBuildFastMutationHistory({
        workbook,
        getCellByIndex: () => numberSnapshot('Sheet1', 'A1', 1),
        ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }],
        preparedCellAddressesByOpIndex: [],
      }),
    ).toThrow('Prepared cell addresses must align with fast-history operations')

    const sparseOps: EngineOp[] = []
    sparseOps.length = 1
    expect(
      tryBuildFastMutationHistory({
        workbook,
        getCellByIndex: () => numberSnapshot('Sheet1', 'A1', 1),
        ops: sparseOps,
      }),
    ).toBeNull()
  })
})

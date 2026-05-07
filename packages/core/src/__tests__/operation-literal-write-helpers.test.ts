import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { CellFlags, CellStore } from '../cell-store.js'
import { StringPool } from '../string-pool.js'
import {
  canFastPathOperationLiteralOverwrite,
  canSkipOperationFormulaColumnVersion,
  cellsShareOperationVersionColumn,
  isOperationClearCellNoOp,
  isOperationNullLiteralWriteNoOp,
  writeFastPathOperationLiteralToExistingCell,
  writeOperationLiteralToExistingCell,
  writeOperationNumericLiteralToExistingCell,
  writeTrustedOperationExistingNumericLiteralToCell,
} from '../engine/services/operation-literal-write-helpers.js'

function createCellStoreWithCells(): CellStore {
  const cellStore = new CellStore()
  cellStore.allocate(1, 0, 0)
  cellStore.allocate(1, 1, 0)
  cellStore.allocate(1, 1, 1)
  return cellStore
}

describe('operation literal write helpers', () => {
  it('writes general literals and notifies when the cell store has no set-value hook', () => {
    const cellStore = createCellStoreWithCells()
    const strings = new StringPool()
    const notifications: number[] = []

    writeOperationLiteralToExistingCell({
      workbook: { cellStore, notifyCellValueWritten: (cellIndex) => notifications.push(cellIndex) },
      strings,
      cellIndex: 0,
      value: 'hello',
    })

    expect(cellStore.tags[0]).toBe(ValueTag.String)
    expect(strings.get(cellStore.stringIds[0])).toBe('hello')
    expect(notifications).toEqual([0])
  })

  it('writes numeric literals on the hot path and clears authored-blank flags', () => {
    const cellStore = createCellStoreWithCells()
    const notifications: number[] = []
    cellStore.flags[1] = CellFlags.AuthoredBlank

    writeOperationNumericLiteralToExistingCell({
      workbook: { cellStore, notifyCellValueWritten: (cellIndex) => notifications.push(cellIndex) },
      cellIndex: 1,
      value: 42,
    })

    expect(cellStore.tags[1]).toBe(ValueTag.Number)
    expect(cellStore.errors[1]).toBe(ErrorCode.None)
    expect(cellStore.stringIds[1]).toBe(0)
    expect(cellStore.numbers[1]).toBe(42)
    expect(cellStore.flags[1] & CellFlags.AuthoredBlank).toBe(0)
    expect(cellStore.versions[1]).toBe(1)
    expect(notifications).toEqual([1])
  })

  it('uses trusted numeric writes to bump column versions when no hook owns notification', () => {
    const cellStore = createCellStoreWithCells()
    const columnVersions = new Uint32Array(3)
    cellStore.flags[2] = CellFlags.AuthoredBlank

    writeTrustedOperationExistingNumericLiteralToCell({
      cellStore,
      cellIndex: 2,
      sheet: { columnVersions },
      col: 1,
      value: 7,
    })

    expect(cellStore.numbers[2]).toBe(7)
    expect(cellStore.flags[2] & CellFlags.AuthoredBlank).toBe(0)
    expect(columnVersions[1]).toBe(1)
  })

  it('routes fast-path literal writes by value shape and detects shared physical version columns', () => {
    const cellStore = createCellStoreWithCells()
    const strings = new StringPool()
    const notifications: number[] = []
    const workbook = {
      cellStore,
      notifyCellValueWritten: (cellIndex: number) => notifications.push(cellIndex),
      getSheetById: () => undefined,
    }

    writeFastPathOperationLiteralToExistingCell({ workbook, strings, cellIndex: 0, value: 5 })
    writeFastPathOperationLiteralToExistingCell({ workbook, strings, cellIndex: 1, value: true })

    expect(cellStore.numbers[0]).toBe(5)
    expect(cellStore.tags[1]).toBe(ValueTag.Boolean)
    expect(cellsShareOperationVersionColumn({ workbook, leftCellIndex: 0, rightCellIndex: 1 })).toBe(true)
    expect(cellsShareOperationVersionColumn({ workbook, leftCellIndex: 0, rightCellIndex: 2 })).toBe(false)
  })

  it('skips formula column version bumps only when the visible column has no tracked dependents', () => {
    const cellStore = createCellStoreWithCells()
    const trackedColumns = new Set<string>(['1:0'])
    const workbook = {
      cellStore,
      getSheetById: () => undefined,
    }

    const hasTrackedColumnDependents = (sheetId: number, col: number): boolean => trackedColumns.has(`${sheetId}:${col}`)

    expect(canSkipOperationFormulaColumnVersion({ workbook, cellIndex: 0, hasTrackedColumnDependents })).toBe(false)
    expect(canSkipOperationFormulaColumnVersion({ workbook, cellIndex: 2, hasTrackedColumnDependents })).toBe(true)
    expect(canSkipOperationFormulaColumnVersion({ workbook, cellIndex: 99, hasTrackedColumnDependents })).toBe(false)
  })

  it('uses logical visible columns for version-column dependency checks', () => {
    const cellStore = createCellStoreWithCells()
    cellStore.cols[0] = 8
    cellStore.cols[1] = 9
    const sheet = {
      structureVersion: 2,
      logical: {
        getCellVisiblePosition: (cellIndex: number) => (cellIndex === 0 ? { row: 0, col: 3 } : { row: 1, col: 3 }),
      },
    }
    const workbook = {
      cellStore,
      getSheetById: () => sheet,
    }

    expect(
      canSkipOperationFormulaColumnVersion({
        workbook,
        cellIndex: 0,
        hasTrackedColumnDependents: (_sheetId, col) => col === 3,
      }),
    ).toBe(false)
    expect(cellsShareOperationVersionColumn({ workbook, leftCellIndex: 0, rightCellIndex: 1 })).toBe(true)
  })

  it('classifies literal overwrites and clear no-ops from formula, format, flag, and value state', () => {
    const cellStore = createCellStoreWithCells()
    const strings = new StringPool()
    const formulas = new Map<number, unknown>()
    const state = {
      workbook: {
        cellStore,
        getCellFormat: (cellIndex: number) => (cellIndex === 1 ? '0.00' : undefined),
      },
      strings,
      formulas,
    }

    expect(canFastPathOperationLiteralOverwrite({ cellStore, formulas, cellIndex: 0 })).toBe(true)
    expect(isOperationNullLiteralWriteNoOp({ state, cellIndex: 0 })).toBe(true)
    expect(isOperationClearCellNoOp({ state, cellIndex: 0 })).toBe(true)

    cellStore.flags[0] = CellFlags.HasFormula
    expect(canFastPathOperationLiteralOverwrite({ cellStore, formulas, cellIndex: 0 })).toBe(false)
    expect(isOperationNullLiteralWriteNoOp({ state, cellIndex: 0 })).toBe(false)

    formulas.set(0, {})
    cellStore.flags[0] = 0
    expect(canFastPathOperationLiteralOverwrite({ cellStore, formulas, cellIndex: 0 })).toBe(false)
    expect(isOperationNullLiteralWriteNoOp({ state, cellIndex: 0 })).toBe(false)

    expect(isOperationNullLiteralWriteNoOp({ state, cellIndex: 1 })).toBe(false)
    cellStore.flags[2] = CellFlags.AuthoredBlank
    expect(isOperationNullLiteralWriteNoOp({ state, cellIndex: 2 })).toBe(true)
    expect(isOperationClearCellNoOp({ state, cellIndex: 2 })).toBe(false)
  })
})

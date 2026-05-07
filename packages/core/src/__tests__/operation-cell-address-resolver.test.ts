import { describe, expect, it, vi } from 'vitest'
import { makeCellKey } from '../workbook-store.js'
import type { EngineCellMutationRef } from '../cell-mutations-at.js'
import {
  createOperationPreparedCellTracker,
  createOperationSheetNameResolver,
  resolveOperationExistingMutationCellIndex,
} from '../engine/services/operation-cell-address-resolver.js'

function createPreparedWorkbook() {
  const sheets = new Map<string, { id: number }>([['Sheet1', { id: 1 }]])
  const sheetsById = new Map<number, { id: number }>([[1, { id: 1 }]])
  const cellKeyToIndex = new Map<number, number>([[makeCellKey(1, 2, 3), 23]])
  return {
    sheets,
    sheetsById,
    cellKeyToIndex,
    workbook: {
      cellKeyToIndex,
      getCellIndex: vi.fn((sheetName: string, address: string) => (sheetName === 'Sheet1' && address === 'D3' ? 11 : undefined)),
      getSheet: vi.fn((sheetName: string) => sheets.get(sheetName)),
      getSheetById: vi.fn((sheetId: number) => sheetsById.get(sheetId)),
      getOrCreateSheet: vi.fn((sheetName: string) => {
        const existing = sheets.get(sheetName)
        if (existing) {
          return existing
        }
        const sheet = { id: sheets.size + 1 }
        sheets.set(sheetName, sheet)
        sheetsById.set(sheet.id, sheet)
        return sheet
      }),
      ensureCellAt: vi.fn((sheetId: number, row: number, col: number) => ({ cellIndex: sheetId * 100 + row * 10 + col })),
    },
  }
}

function mutationRef(overrides: Partial<EngineCellMutationRef> = {}): EngineCellMutationRef {
  return {
    sheetId: 1,
    mutation: { kind: 'setCellValue', row: 2, col: 3, value: 42 },
    ...overrides,
  }
}

describe('operation cell address resolver', () => {
  it('uses normal workbook address lookup when no prepared cell address is available', () => {
    const { workbook } = createPreparedWorkbook()
    const tracker = createOperationPreparedCellTracker({
      workbook,
      ensureCellTracked: vi.fn(() => 99),
    })

    expect(tracker.getExistingCellIndex('Sheet1', 'D3', null)).toBe(11)
    expect(workbook.getCellIndex).toHaveBeenCalledWith('Sheet1', 'D3')
  })

  it('resolves prepared existing cell addresses through cached sheet ids', () => {
    const { workbook } = createPreparedWorkbook()
    const tracker = createOperationPreparedCellTracker({
      workbook,
      ensureCellTracked: vi.fn(() => 99),
    })

    expect(tracker.getExistingCellIndex('Sheet1', 'ignored', { row: 2, col: 3 })).toBe(23)
    expect(tracker.getExistingCellIndex('Sheet1', 'ignored', { row: 2, col: 3 })).toBe(23)
    expect(workbook.getSheet).toHaveBeenCalledTimes(1)
  })

  it('invalidates prepared sheet id cache after sheet identity changes', () => {
    const { workbook, sheets, sheetsById, cellKeyToIndex } = createPreparedWorkbook()
    const tracker = createOperationPreparedCellTracker({
      workbook,
      ensureCellTracked: vi.fn(() => 99),
    })

    expect(tracker.getExistingCellIndex('Sheet1', 'ignored', { row: 2, col: 3 })).toBe(23)
    sheets.set('Sheet1', { id: 2 })
    sheetsById.delete(1)
    sheetsById.set(2, { id: 2 })
    cellKeyToIndex.set(makeCellKey(2, 2, 3), 46)
    tracker.invalidateSheetName('Sheet1')

    expect(tracker.getExistingCellIndex('Sheet1', 'ignored', { row: 2, col: 3 })).toBe(46)
  })

  it('ensures prepared cells by sheet id and falls back to the address path without prepared coordinates', () => {
    const { workbook } = createPreparedWorkbook()
    const ensureCellTracked = vi.fn(() => 77)
    const tracker = createOperationPreparedCellTracker({ workbook, ensureCellTracked })

    expect(tracker.ensureCellTracked('Sheet1', 'D3', { row: 2, col: 3 })).toBe(123)
    expect(workbook.ensureCellAt).toHaveBeenCalledWith(1, 2, 3)
    expect(tracker.ensureCellTracked('Sheet1', 'D3', null)).toBe(77)
    expect(ensureCellTracked).toHaveBeenCalledWith('Sheet1', 'D3')
  })

  it('returns the supplied mutation cell index for current physical sheet coordinates', () => {
    const workbook = {
      cellStore: {
        sheetIds: [undefined, 1],
        rows: [undefined, 2],
        cols: [undefined, 3],
      },
      cellKeyToIndex: new Map<number, number>(),
      getSheetById: () => ({ structureVersion: 1 }),
    }

    expect(resolveOperationExistingMutationCellIndex(workbook, mutationRef({ cellIndex: 1 }))).toBe(1)
  })

  it('returns the supplied mutation cell index for current logical visible coordinates', () => {
    const workbook = {
      cellStore: {
        sheetIds: [undefined, 1],
        rows: [undefined, 20],
        cols: [undefined, 30],
      },
      cellKeyToIndex: new Map<number, number>(),
      getSheetById: () => ({
        structureVersion: 2,
        logical: {
          getCellVisiblePosition: () => ({ row: 2, col: 3 }),
        },
      }),
    }

    expect(resolveOperationExistingMutationCellIndex(workbook, mutationRef({ cellIndex: 1 }))).toBe(1)
  })

  it('falls back to cell-key lookup when a supplied mutation cell index is stale', () => {
    const workbook = {
      cellStore: {
        sheetIds: [undefined, 1],
        rows: [undefined, 9],
        cols: [undefined, 9],
      },
      cellKeyToIndex: new Map<number, number>([[makeCellKey(1, 2, 3), 88]]),
      getSheetById: () => ({ structureVersion: 1 }),
    }

    expect(resolveOperationExistingMutationCellIndex(workbook, mutationRef({ cellIndex: 1 }))).toBe(88)
  })

  it('resolves sheet names through a cached sheet-id lookup', () => {
    const getSheetById = vi.fn((sheetId: number) => (sheetId === 7 ? { name: 'Sheet7' } : undefined))
    const resolver = createOperationSheetNameResolver({ getSheetById })

    expect(resolver.resolve(7)).toBe('Sheet7')
    expect(resolver.resolve(7)).toBe('Sheet7')
    expect(getSheetById).toHaveBeenCalledTimes(1)
    expect(() => resolver.resolve(8)).toThrow('Unknown sheet id: 8')
  })
})

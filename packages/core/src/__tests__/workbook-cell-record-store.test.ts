import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { CellFlags, CellStore } from '../cell-store.js'
import { WorkbookCellRecordStore } from '../workbook-cell-record-store.js'
import { makeCellKey } from '../workbook-cell-key-index.js'
import { createWorkbookSheetRecord, type SheetRecord } from '../workbook-sheet-record.js'

function createStore() {
  let nextRowId = 1
  let nextColumnId = 1
  const cellStore = new CellStore()
  const cellKeyToIndex = new Map<number, number>()
  const cellFormats = new Map<number, string>()
  const sheetsByName = new Map<string, SheetRecord>()
  const sheetsById = new Map<number, SheetRecord>()
  const createSheet = (name: string, id = sheetsById.size + 1) => {
    const sheet = createWorkbookSheetRecord({ id, name, order: sheetsByName.size, counters: undefined })
    sheetsByName.set(name, sheet)
    sheetsById.set(id, sheet)
    return sheet
  }
  const getOrCreateSheet = (name: string) => sheetsByName.get(name) ?? createSheet(name)
  const records = new WorkbookCellRecordStore({
    cellStore,
    cellKeyToIndex,
    cellFormats,
    getSheet: (sheetName) => sheetsByName.get(sheetName),
    getOrCreateSheet,
    getSheetById: (sheetId) => sheetsById.get(sheetId),
    getSheetNameById: (sheetId) => sheetsById.get(sheetId)?.name ?? '',
    createLogicalAxisId: (axis) => (axis === 'row' ? `row-${nextRowId++}` : `col-${nextColumnId++}`),
  })
  return { records, cellStore, cellKeyToIndex, cellFormats, createSheet }
}

describe('WorkbookCellRecordStore', () => {
  it('creates and resolves physical and logical cell records', () => {
    const { records, cellStore, cellKeyToIndex, createSheet } = createStore()
    const sheet = createSheet('Sheet1', 7)

    expect(() => records.ensureCellAt(999, 0, 0)).toThrow('Unknown sheet id: 999')
    expect(() => records.attachAllocatedCell(999, 0, 0, 0)).toThrow('Unknown sheet id: 999')
    expect(() => records.ensureLogicalAxisId(999, 'row', 0)).toThrow('Unknown sheet id: 999')
    expect(() => records.createLogicalAxisIdEnsurer(999, 'column')).toThrow('Unknown sheet id: 999')
    expect(() => records.attachAllocatedCellWithLogicalAxisIds(999, 0, 0, 0, 'r', 'c')).toThrow('Unknown sheet id: 999')

    const created = records.ensureCellRecord('Sheet1', 'B2')
    expect(created.created).toBe(true)
    expect(records.ensureCellAt(sheet.id, 1, 1)).toEqual({ cellIndex: created.cellIndex, created: false })
    expect(records.ensureCell('Sheet1', 'B2')).toBe(created.cellIndex)
    expect(records.getCellIndex('Missing', 'A1')).toBeUndefined()
    expect(records.getCellIndex('Sheet1', 'B2')).toBe(created.cellIndex)
    expect(records.getCellIndexAt(sheet.id, 1, 1)).toBe(created.cellIndex)
    expect(records.getFreshCellIndexAt(sheet.id, 1, 1)).toBe(created.cellIndex)
    expect(records.getFreshCellIndexAt(999, 1, 1)).toBeUndefined()
    expect(records.getAddress(created.cellIndex)).toBe('B2')
    expect(records.getQualifiedAddress(created.cellIndex)).toBe('Sheet1!B2')
    expect(records.getCellAxisIndex(created.cellIndex, 'row')).toBe(1)
    expect(records.getCellAxisIndex(999, 'row')).toBeUndefined()

    const rowId = records.ensureLogicalAxisId(sheet.id, 'row', 5)
    const columnId = records.createLogicalAxisIdEnsurer(sheet.id, 'column')(3)
    const logicalIndex = cellStore.allocate(sheet.id, 5, 3)
    records.attachAllocatedCellWithLogicalAxisIds(sheet.id, 5, 3, logicalIndex, rowId, columnId)
    sheet.structureVersion = 2

    expect(records.getCellIndexAt(sheet.id, 5, 3)).toBe(logicalIndex)
    expect(records.getCellPosition(logicalIndex)).toEqual({ sheetId: sheet.id, row: 5, col: 3 })
    expect(records.getCellAxisIndex(logicalIndex, 'column')).toBe(3)

    const staleLogicalIndex = cellStore.allocate(sheet.id, 8, 4)
    records.attachAllocatedCellWithLogicalAxisIds(sheet.id, 8, 4, staleLogicalIndex, 'stale-row', 'stale-column')
    expect(records.getCellPosition(staleLogicalIndex)).toBeUndefined()
    expect(records.getCellAxisIndex(staleLogicalIndex, 'row')).toBeUndefined()

    const detachedLogicalIndex = cellStore.allocate(sheet.id, 8, 5)
    expect(records.getCellPosition(detachedLogicalIndex)).toEqual({ sheetId: sheet.id, row: 8, col: 5 })
    expect(records.getCellAxisIndex(detachedLogicalIndex, 'column')).toBe(5)

    expect(cellKeyToIndex.get(makeCellKey(sheet.id, 5, 3))).toBe(logicalIndex)
    expect(records.detachCellIndex(999)).toBe(false)
    expect(records.detachCellIndex(logicalIndex)).toBe(true)
    expect(cellKeyToIndex.has(makeCellKey(sheet.id, 5, 3))).toBe(false)
    expect(cellStore.flags[logicalIndex] & CellFlags.Materialized).toBe(0)
  })

  it('prunes only materialized empty cells with no authored state', () => {
    const { records, cellStore, cellFormats, createSheet } = createStore()
    const sheet = createSheet('Sheet1', 3)

    expect(records.pruneCellIfEmpty(123)).toBe(false)

    const formatted = records.ensureCellAt(sheet.id, 0, 0).cellIndex
    cellFormats.set(formatted, 'currency')
    expect(records.pruneCellIfEmpty(formatted)).toBe(false)

    const valued = records.ensureCellAt(sheet.id, 1, 0).cellIndex
    cellStore.setValue(valued, { tag: ValueTag.Number, value: 12 })
    expect(records.pruneCellIfEmpty(valued)).toBe(false)

    const formula = records.ensureCellAt(sheet.id, 2, 0).cellIndex
    cellStore.flags[formula] = (cellStore.flags[formula] ?? 0) | CellFlags.HasFormula
    expect(records.pruneCellIfEmpty(formula)).toBe(false)

    const empty = records.ensureCellAt(sheet.id, 3, 0).cellIndex
    expect(records.pruneCellIfEmpty(empty)).toBe(true)
    expect(records.getCellIndexAt(sheet.id, 3, 0)).toBeUndefined()
  })
})

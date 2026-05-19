import { describe, expect, it } from 'vitest'
import { makeLogicalCellKey } from '../workbook-store.js'
import { AxisResidentCellIndex } from '../storage/axis-resident-cell-index.js'
import { CellPageStore } from '../storage/cell-page-store.js'
import { CellAxisIdentityStore } from '../storage/cell-axis-identity-store.js'
import { LogicalSheetStore } from '../storage/logical-sheet-store.js'
import { SheetAxisMap } from '../storage/sheet-axis-map.js'

describe('LogicalSheetStore', () => {
  it('resolves stable visible row and column ids and keeps cell pages attached to those ids', () => {
    const cellIdentities = new CellAxisIdentityStore()
    const cellPages = new CellPageStore(
      new Map<string, number>(),
      (location) => makeLogicalCellKey(location.sheetId, location.rowId, location.colId),
      (callback) => {
        cellIdentities.forEach((identity, cellIndex) => {
          callback(identity, cellIndex)
        })
      },
    )
    const axisMap = new SheetAxisMap()
    const logical = new LogicalSheetStore(7, axisMap, cellPages, cellIdentities, new AxisResidentCellIndex())

    logical.setVisibleCell(1, 1, 42, {
      createRowId: () => 'row-b',
      createColumnId: () => 'column-b',
    })

    expect(logical.resolveVisibleCell(1, 1)).toEqual({
      sheetId: 7,
      row: 1,
      col: 1,
      rowRef: { index: 1, id: 'row-b' },
      colRef: { index: 1, id: 'column-b' },
    })
    expect(logical.getVisibleCell(1, 1)).toBe(42)

    axisMap.move('row', 1, 1, 0)
    axisMap.move('column', 1, 1, 0)

    expect(logical.resolveVisibleCell(0, 0)).toEqual({
      sheetId: 7,
      row: 0,
      col: 0,
      rowRef: { index: 0, id: 'row-b' },
      colRef: { index: 0, id: 'column-b' },
    })
    expect(logical.getVisibleCell(0, 0)).toBe(42)
    expect(logical.getCellVisiblePosition(42)).toEqual({ row: 0, col: 0 })
    expect(logical.listResidentCellIndices('row', ['row-b'])).toEqual([42])
    expect(logical.getVisibleCell(1, 1)).toBeUndefined()

    const rowMatches: Array<{ cellIndex: number; axisIndex: number; rowId: string; colId: string }> = []
    logical.forEachResidentCellInAxisEntries('row', [{ id: 'row-b', index: 0 }], (cellIndex, identity, axisIndex) => {
      rowMatches.push({ cellIndex, axisIndex, rowId: identity.rowId, colId: identity.colId })
    })
    expect(rowMatches).toEqual([{ cellIndex: 42, axisIndex: 0, rowId: 'row-b', colId: 'column-b' }])

    logical.setSheetId(9)
    expect(logical.resolveVisibleCell(0, 0).sheetId).toBe(9)
  })

  it('materializes deferred cell pages from cell identities on first logical lookup', () => {
    const cellIdentities = new CellAxisIdentityStore()
    const cellPages = new CellPageStore(
      new Map<string, number>(),
      (location) => makeLogicalCellKey(location.sheetId, location.rowId, location.colId),
      (callback) => {
        cellIdentities.forEach((identity, cellIndex) => {
          callback(identity, cellIndex)
        })
      },
    )
    const axisMap = new SheetAxisMap()
    const logical = new LogicalSheetStore(7, axisMap, cellPages, cellIdentities, new AxisResidentCellIndex())

    axisMap.ensureId('row', 2, () => 'row-c')
    axisMap.ensureId('column', 3, () => 'column-d')
    cellIdentities.set(42, { sheetId: 7, rowId: 'row-c', colId: 'column-d' })
    cellPages.setDeferred({ sheetId: 7, rowId: 'row-c', colId: 'column-d' }, 42)

    expect(logical.getVisibleCell(2, 3)).toBe(42)
    expect(logical.deleteVisibleCellByIds('row-c', 'column-d')).toBe(true)
    expect(logical.getVisibleCell(2, 3)).toBeUndefined()
  })

  it('iterates visible row cells through the resident row index', () => {
    const cellIdentities = new CellAxisIdentityStore()
    const cellPages = new CellPageStore(
      new Map<string, number>(),
      (location) => makeLogicalCellKey(location.sheetId, location.rowId, location.colId),
      (callback) => {
        cellIdentities.forEach((identity, cellIndex) => {
          callback(identity, cellIndex)
        })
      },
    )
    const axisMap = new SheetAxisMap()
    const logical = new LogicalSheetStore(7, axisMap, cellPages, cellIdentities, new AxisResidentCellIndex())
    let nextRowId = 0
    let nextColumnId = 0
    const factories = {
      createRowId: () => `row-${nextRowId++}`,
      createColumnId: () => `column-${nextColumnId++}`,
    }

    logical.setVisibleCell(4, 3, 43, factories)
    logical.setVisibleCell(4, 1, 41, factories)
    logical.setVisibleCell(5, 9, 59, factories)

    const entries: Array<{ cellIndex: number; col: number }> = []
    logical.forEachVisibleRowCellEntry(4, (cellIndex, col) => {
      entries.push({ cellIndex, col })
    })

    expect(entries).toEqual([
      { cellIndex: 41, col: 1 },
      { cellIndex: 43, col: 3 },
    ])
  })

  it('attaches fresh visible cells through primitive deferred indexes for initial load', () => {
    const cellIdentities = new CellAxisIdentityStore()
    let objectLocationKeyCalls = 0
    let primitiveLocationKeyCalls = 0
    const cellPages = new CellPageStore(
      new Map<string, number>(),
      (location) => {
        objectLocationKeyCalls += 1
        return makeLogicalCellKey(location.sheetId, location.rowId, location.colId)
      },
      (callback) => {
        cellIdentities.forEach((identity, cellIndex) => {
          callback(identity, cellIndex)
        })
      },
      (sheetId, rowId, colId) => {
        primitiveLocationKeyCalls += 1
        return makeLogicalCellKey(sheetId, rowId, colId)
      },
    )
    const axisMap = new SheetAxisMap()
    axisMap.ensureId('row', 2, () => 'row-c')
    axisMap.ensureId('column', 3, () => 'column-d')
    const logical = new LogicalSheetStore(7, axisMap, cellPages, cellIdentities, new AxisResidentCellIndex())

    logical.setFreshVisibleCellWithAxisIdsDeferred(2, 3, 42, 'row-c', 'column-d')

    expect(objectLocationKeyCalls).toBe(0)
    expect(primitiveLocationKeyCalls).toBe(1)
    expect(logical.getCellIdentity(42)).toEqual({ sheetId: 7, rowId: 'row-c', colId: 'column-d' })
    expect(logical.listResidentCellIndices('column', ['column-d'])).toEqual([42])
    expect(logical.getVisibleCell(2, 3)).toBe(42)
  })

  it('resolves dense row-major logical cells after axis edits without rebuilding cell pages', () => {
    const cellIdentities = new CellAxisIdentityStore()
    let rebuildCount = 0
    const cellPages = new CellPageStore(
      new Map<string, number>(),
      (location) => makeLogicalCellKey(location.sheetId, location.rowId, location.colId),
      (callback) => {
        rebuildCount += 1
        cellIdentities.forEach((identity, cellIndex) => {
          callback(identity, cellIndex)
        })
      },
    )
    const axisMap = new SheetAxisMap()
    const rowIds = ['row-1', 'row-2', 'row-3']
    const colIds = ['col-a', 'col-b', 'col-c']
    rowIds.forEach((id, index) => {
      axisMap.setId('row', index, id)
    })
    colIds.forEach((id, index) => {
      axisMap.setId('column', index, id)
    })
    const logical = new LogicalSheetStore(7, axisMap, cellPages, cellIdentities, new AxisResidentCellIndex())

    logical.deferVisibleCellPageRebuild()
    logical.setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred(100, rowIds, colIds)
    axisMap.splice('column', 1, 0, 1, [])

    expect(logical.getVisibleCell(2, 3)).toBe(108)
    expect(rebuildCount).toBe(0)
    expect(logical.getVisibleCell(2, 1)).toBeUndefined()
  })
})

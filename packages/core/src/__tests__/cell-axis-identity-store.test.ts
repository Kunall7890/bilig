import { describe, expect, it } from 'vitest'
import { AxisResidentCellIndex } from '../storage/axis-resident-cell-index.js'
import { CellAxisIdentityStore } from '../storage/cell-axis-identity-store.js'

describe('CellAxisIdentityStore', () => {
  it('tracks cell index to stable row and column ids', () => {
    const store = new CellAxisIdentityStore()

    store.set(7, { sheetId: 2, rowId: 'row-a', colId: 'column-c' })

    expect(store.get(7)).toEqual({ sheetId: 2, rowId: 'row-a', colId: 'column-c' })
    expect(store.entries()).toEqual([[7, { sheetId: 2, rowId: 'row-a', colId: 'column-c' }]])

    expect(store.delete(7)).toBe(true)
    expect(store.get(7)).toBeUndefined()
  })

  it('indexes resident cells by row and column ids for structural deletes', () => {
    const index = new AxisResidentCellIndex()

    index.set(3, { rowId: 'row-a', colId: 'column-a' })
    index.set(1, { rowId: 'row-a', colId: 'column-b' })
    index.set(2, { rowId: 'row-b', colId: 'column-b' })

    expect(index.cellsInRow('row-a')).toEqual([1, 3])
    expect(index.cellsInColumn('column-b')).toEqual([1, 2])
    expect(index.cellsInRows(['row-b', 'row-a'])).toEqual([1, 2, 3])
    expect(index.cellsInRowsUnordered(['row-b', 'row-a'])).toEqual([2, 3, 1])
    expect(index.cellsInColumnsUnordered(['column-b', 'column-a'])).toEqual([1, 2, 3])
    const rowCells: number[] = []
    const columnCells: number[] = []
    index.forEachCellInRow('row-a', (cellIndex) => rowCells.push(cellIndex))
    index.forEachCellInColumn('column-b', (cellIndex) => columnCells.push(cellIndex))
    expect(rowCells).toEqual([3, 1])
    expect(columnCells).toEqual([1, 2])

    index.set(1, { rowId: 'row-c', colId: 'column-c' })

    expect(index.cellsInRow('row-a')).toEqual([3])
    expect(index.cellsInColumn('column-b')).toEqual([2])
    expect(index.cellsInRows(['row-a', 'row-c'])).toEqual([1, 3])
    expect(index.delete(2)).toBe(true)
    expect(index.cellsInColumn('column-b')).toEqual([])
  })

  it('builds resident row and column indexes lazily after deferred bulk attachment', () => {
    const index = new AxisResidentCellIndex()

    index.addDeferred(3, { rowId: 'row-a', colId: 'column-a' })
    index.addDeferred(1, { rowId: 'row-a', colId: 'column-b' })
    index.addDeferred(2, { rowId: 'row-b', colId: 'column-b' })
    index.delete(2)
    index.set(1, { rowId: 'row-c', colId: 'column-c' })

    expect(index.get(3)).toEqual({ rowId: 'row-a', colId: 'column-a' })
    expect(index.cellsInRow('row-a')).toEqual([3])
    expect(index.cellsInColumn('column-b')).toEqual([])
    expect(index.cellsInRows(['row-c', 'row-a'])).toEqual([1, 3])
    expect(index.cellsInRowsUnordered(['row-c', 'row-a'])).toEqual([1, 3])
    const rowCells: number[] = []
    index.forEachCellInRow('row-c', (cellIndex) => rowCells.push(cellIndex))
    expect(rowCells).toEqual([1])
  })

  it('rebuilds resident cells from identity storage after a deferred primary rebuild', () => {
    const store = new CellAxisIdentityStore()
    const index = new AxisResidentCellIndex((callback) => {
      store.forEach((identity, cellIndex) => callback(cellIndex, identity))
    })

    index.deferRebuild()
    store.set(3, { sheetId: 1, rowId: 'row-a', colId: 'column-a' })
    store.set(1, { sheetId: 1, rowId: 'row-a', colId: 'column-b' })
    store.set(2, { sheetId: 1, rowId: 'row-b', colId: 'column-b' })

    expect(index.get(3)).toEqual({ sheetId: 1, rowId: 'row-a', colId: 'column-a' })
    expect(index.cellsInRow('row-a')).toEqual([1, 3])
    expect(index.cellsInColumn('column-b')).toEqual([1, 2])

    store.delete(2)
    index.delete(2)
    store.set(1, { sheetId: 1, rowId: 'row-c', colId: 'column-c' })
    index.set(1, { rowId: 'row-c', colId: 'column-c' })

    expect(index.cellsInRows(['row-a', 'row-c'])).toEqual([1, 3])
    expect(index.cellsInColumn('column-b')).toEqual([])
  })
})

import type { StructuralAxisTransform } from '@bilig/formula'
import { mapStructuralAxisIndex } from './engine-structural-utils.js'
import { buildStructuralTransaction, structuralScopeForTransform, type StructuralTransaction } from './engine/structural-transaction.js'
import { addEngineCounter, type EngineCounters } from './perf/engine-counters.js'
import type { CellStore } from './cell-store.js'
import type { SheetGridAxisRemapScope } from './sheet-grid.js'
import { makeCellKey } from './workbook-cell-key-index.js'
import type { SheetRecord } from './workbook-sheet-record.js'

type WorkbookAxis = 'row' | 'column'

export class WorkbookStructuralCellStore {
  constructor(
    private readonly options: {
      readonly counters: EngineCounters | undefined
      readonly cellStore: CellStore
      readonly cellKeyToIndex: Map<number, number>
      readonly getSheet: (sheetName: string) => SheetRecord | undefined
      readonly createLogicalAxisId: (axis: WorkbookAxis) => string
    },
  ) {}

  remapSheetCells(
    sheetName: string,
    axis: WorkbookAxis,
    remapIndex: (index: number) => number | undefined,
    scope?: SheetGridAxisRemapScope,
  ): { changedCellIndices: number[]; removedCellIndices: number[] } {
    const sheet = this.options.getSheet(sheetName)
    if (!sheet) {
      return { changedCellIndices: [], removedCellIndices: [] }
    }
    const changedEntries = sheet.grid.remapAxis(axis, remapIndex, scope)
    if (this.options.counters && changedEntries.length > 0) {
      addEngineCounter(this.options.counters, 'cellsRemapped', changedEntries.length)
    }
    changedEntries.forEach(({ row, col }) => {
      sheet.logical.deleteVisibleCell(row, col)
      this.options.cellKeyToIndex.delete(makeCellKey(sheet.id, row, col))
    })

    const changedCellIndices: number[] = []
    const removedCellIndices: number[] = []
    for (const { cellIndex, nextRow, nextCol } of changedEntries) {
      if (nextRow === undefined || nextCol === undefined) {
        removedCellIndices.push(cellIndex)
        continue
      }
      this.options.cellStore.rows[cellIndex] = nextRow
      this.options.cellStore.cols[cellIndex] = nextCol
      this.options.cellKeyToIndex.set(makeCellKey(sheet.id, nextRow, nextCol), cellIndex)
      sheet.logical.setVisibleCell(nextRow, nextCol, cellIndex, {
        createRowId: () => this.options.createLogicalAxisId('row'),
        createColumnId: () => this.options.createLogicalAxisId('column'),
      })
      changedCellIndices.push(cellIndex)
    }

    return { changedCellIndices, removedCellIndices }
  }

  planStructuralAxisTransform(sheetName: string, transform: StructuralAxisTransform): StructuralTransaction | undefined {
    const sheet = this.options.getSheet(sheetName)
    if (!sheet) {
      return undefined
    }
    if (this.options.counters) {
      addEngineCounter(this.options.counters, 'structuralTransactions')
    }
    const remappedCells: Array<StructuralTransaction['remappedCells'][number]> = []
    if (transform.kind === 'delete') {
      const deletedAxisEntries = sheet.logicalAxisMap.snapshot(transform.axis, transform.start, transform.count)
      sheet.logical.forEachResidentCellInAxisEntries(transform.axis, deletedAxisEntries, (cellIndex, identity, deletedAxisIndex) => {
        const otherAxis = transform.axis === 'row' ? 'column' : 'row'
        const otherAxisId = transform.axis === 'row' ? identity.colId : identity.rowId
        const otherAxisIndex = sheet.logicalAxisMap.indexOf(otherAxis, otherAxisId)
        if (otherAxisIndex < 0) {
          return
        }
        const fromRow = transform.axis === 'row' ? deletedAxisIndex : otherAxisIndex
        const fromCol = transform.axis === 'row' ? otherAxisIndex : deletedAxisIndex
        remappedCells.push({
          cellIndex,
          fromRow,
          fromCol,
          fromRowId: identity.rowId,
          fromColId: identity.colId,
          toRow: undefined,
          toCol: undefined,
        })
      })
    }
    if (this.options.counters && remappedCells.length > 0) {
      addEngineCounter(this.options.counters, 'structuralPlannedCells', remappedCells.length)
      addEngineCounter(this.options.counters, 'structuralRemovedCells', remappedCells.length)
    }
    return buildStructuralTransaction({
      sheetName,
      sheetId: sheet.id,
      transform,
      remappedCells,
    })
  }

  applyPlannedStructuralTransaction(transaction: StructuralTransaction): StructuralTransaction | undefined {
    const sheet = this.options.getSheet(transaction.sheetName)
    if (!sheet) {
      return undefined
    }
    let hasSurvivingRemap = false
    let survivingRemapCount = 0
    for (const entry of transaction.remappedCells) {
      if (entry.toRow !== undefined && entry.toCol !== undefined) {
        hasSurvivingRemap = true
        survivingRemapCount += 1
      }
    }
    if (this.options.counters && survivingRemapCount > 0) {
      addEngineCounter(this.options.counters, 'cellsRemapped', survivingRemapCount)
      addEngineCounter(this.options.counters, 'structuralSurvivorCellsRemapped', survivingRemapCount)
    }
    if (!hasSurvivingRemap) {
      return transaction
    }
    for (const entry of transaction.remappedCells) {
      this.options.cellKeyToIndex.delete(makeCellKey(sheet.id, entry.fromRow, entry.fromCol))
      if (sheet.grid.get(entry.fromRow, entry.fromCol) === entry.cellIndex) {
        sheet.grid.clear(entry.fromRow, entry.fromCol)
      }
      if (entry.toRow === undefined || entry.toCol === undefined) {
        if (entry.fromRowId && entry.fromColId) {
          const hasTombstonedIdentity =
            sheet.logical.getCellVisiblePosition(entry.cellIndex) === undefined &&
            sheet.logical.getCellIdentity(entry.cellIndex) !== undefined
          if (!hasTombstonedIdentity) {
            sheet.logical.deleteVisibleCellByIds(entry.fromRowId, entry.fromColId)
          }
        } else {
          sheet.logical.deleteVisibleCell(entry.fromRow, entry.fromCol)
        }
      } else {
        hasSurvivingRemap = true
      }
    }
    for (const entry of transaction.remappedCells) {
      if (entry.toRow === undefined || entry.toCol === undefined) {
        continue
      }
      this.options.cellStore.rows[entry.cellIndex] = entry.toRow
      this.options.cellStore.cols[entry.cellIndex] = entry.toCol
      this.options.cellKeyToIndex.set(makeCellKey(sheet.id, entry.toRow, entry.toCol), entry.cellIndex)
      sheet.grid.set(entry.toRow, entry.toCol, entry.cellIndex)
    }
    return transaction
  }

  applyStructuralAxisTransform(sheetName: string, transform: StructuralAxisTransform): StructuralTransaction | undefined {
    const sheet = this.options.getSheet(sheetName)
    if (!sheet) {
      return undefined
    }

    const scope = structuralScopeForTransform(transform)
    const remappedEntries = sheet.grid.remapAxis(transform.axis, (index) => mapStructuralAxisIndex(index, transform), scope)
    if (this.options.counters && remappedEntries.length > 0) {
      addEngineCounter(this.options.counters, 'cellsRemapped', remappedEntries.length)
    }
    remappedEntries.forEach(({ row, col }) => {
      sheet.logical.deleteVisibleCell(row, col)
      this.options.cellKeyToIndex.delete(makeCellKey(sheet.id, row, col))
    })

    const remappedCells = remappedEntries.map(({ cellIndex, row, col, nextRow, nextCol }) => {
      if (nextRow !== undefined && nextCol !== undefined) {
        this.options.cellStore.rows[cellIndex] = nextRow
        this.options.cellStore.cols[cellIndex] = nextCol
        this.options.cellKeyToIndex.set(makeCellKey(sheet.id, nextRow, nextCol), cellIndex)
        sheet.logical.setVisibleCell(nextRow, nextCol, cellIndex, {
          createRowId: () => this.options.createLogicalAxisId('row'),
          createColumnId: () => this.options.createLogicalAxisId('column'),
        })
      }
      return {
        cellIndex,
        fromRow: row,
        fromCol: col,
        toRow: nextRow,
        toCol: nextCol,
      }
    })

    return buildStructuralTransaction({
      sheetName,
      sheetId: sheet.id,
      transform,
      remappedCells,
    })
  }
}

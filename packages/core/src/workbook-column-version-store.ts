import { MAX_COLS } from '@bilig/protocol'
import type { CellStore } from './cell-store.js'
import type { SheetRecord } from './workbook-sheet-record.js'

export class WorkbookColumnVersionStore {
  private batchedColumnVersionUpdates: Map<number, Set<number>> | null = null

  constructor(
    private readonly options: {
      readonly cellStore: CellStore
      readonly getSheetById: (sheetId: number) => SheetRecord | undefined
    },
  ) {}

  withBatchedColumnVersionUpdates<T>(execute: () => T): T {
    if (this.batchedColumnVersionUpdates) {
      return execute()
    }
    const pending = new Map<number, Set<number>>()
    this.batchedColumnVersionUpdates = pending
    try {
      return execute()
    } finally {
      this.batchedColumnVersionUpdates = null
      pending.forEach((columns, sheetId) => {
        columns.forEach((col) => {
          this.bumpColumnVersion(sheetId, col)
        })
      })
    }
  }

  notifyCellValueWritten(cellIndex: number): void {
    this.bumpColumnVersionByCellIndex(cellIndex)
  }

  notifyColumnsWritten(sheetId: number, columns: readonly number[] | Uint32Array): void {
    const pending = this.batchedColumnVersionUpdates
    if (pending) {
      let pendingColumns = pending.get(sheetId)
      if (!pendingColumns) {
        pendingColumns = new Set<number>()
        pending.set(sheetId, pendingColumns)
      }
      for (let index = 0; index < columns.length; index += 1) {
        pendingColumns.add(columns[index]!)
      }
      return
    }
    for (let index = 0; index < columns.length; index += 1) {
      this.bumpColumnVersion(sheetId, columns[index]!)
    }
  }

  private bumpColumnVersionByCellIndex(cellIndex: number): void {
    const sheetId = this.options.cellStore.sheetIds[cellIndex]!
    const sheet = this.options.getSheetById(sheetId)
    if (!sheet) {
      return
    }
    const col =
      sheet.structureVersion === 1
        ? this.options.cellStore.cols[cellIndex]!
        : (sheet.logical.getCellVisiblePosition(cellIndex)?.col ?? this.options.cellStore.cols[cellIndex]!)
    const pending = this.batchedColumnVersionUpdates
    if (pending) {
      let columns = pending.get(sheetId)
      if (!columns) {
        columns = new Set<number>()
        pending.set(sheetId, columns)
      }
      columns.add(col)
      return
    }
    const columnVersions = ensureColumnVersionCapacity(sheet, col)
    columnVersions[col] = (columnVersions[col] ?? 0) + 1
  }

  private bumpColumnVersion(sheetId: number, col: number): void {
    const sheet = this.options.getSheetById(sheetId)
    if (!sheet) {
      return
    }
    const columnVersions = ensureColumnVersionCapacity(sheet, col)
    columnVersions[col] = (columnVersions[col] ?? 0) + 1
  }
}

function ensureColumnVersionCapacity(sheet: SheetRecord, col: number): Uint32Array {
  if (col < sheet.columnVersions.length) {
    return sheet.columnVersions
  }
  let nextLength = sheet.columnVersions.length === 0 ? 16 : sheet.columnVersions.length
  while (nextLength <= col && nextLength < MAX_COLS) {
    nextLength = Math.min(MAX_COLS, nextLength * 2)
  }
  const nextColumnVersions = new Uint32Array(nextLength)
  nextColumnVersions.set(sheet.columnVersions)
  sheet.columnVersions = nextColumnVersions
  return nextColumnVersions
}

import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import type { EngineRuntimeState } from '../runtime-state.js'

interface InitialWrittenColumnTracker {
  smallColumns: number
  columns?: Uint8Array
  count: number
}

export interface InitialFormulaValueWriter {
  readonly writeValue: (cellIndex: number, value: CellValue) => void
  readonly writeValueAt: (cellIndex: number, sheetId: number, col: number, value: CellValue) => void
  readonly writeNumber: (cellIndex: number, value: number) => void
  readonly writeNumberAt: (cellIndex: number, sheetId: number, col: number, value: number) => void
  readonly flush: () => void
}

function createInitialWrittenColumnTracker(): InitialWrittenColumnTracker {
  return {
    smallColumns: 0,
    count: 0,
  }
}

function markInitialWrittenColumn(tracker: InitialWrittenColumnTracker, col: number): void {
  if (col < 30) {
    const bit = 1 << col
    if ((tracker.smallColumns & bit) !== 0) {
      return
    }
    tracker.smallColumns |= bit
    tracker.count += 1
    return
  }
  let columns = tracker.columns
  if (!columns) {
    columns = new Uint8Array(Math.max(32, col + 1))
    tracker.columns = columns
  } else if (col >= columns.length) {
    let nextLength = columns.length
    while (nextLength <= col) {
      nextLength *= 2
    }
    const nextColumns = new Uint8Array(nextLength)
    nextColumns.set(columns)
    columns = nextColumns
    tracker.columns = columns
  }
  if (columns[col] !== 0) {
    return
  }
  columns[col] = 1
  tracker.count += 1
}

function materializeInitialWrittenColumns(tracker: InitialWrittenColumnTracker): Uint32Array {
  const columns = new Uint32Array(tracker.count)
  let writeIndex = 0
  let smallColumns = tracker.smallColumns
  while (smallColumns !== 0) {
    const bit = smallColumns & -smallColumns
    columns[writeIndex] = 31 - Math.clz32(bit)
    writeIndex += 1
    smallColumns &= smallColumns - 1
  }
  const largeColumns = tracker.columns
  if (largeColumns) {
    for (let col = 30; col < largeColumns.length; col += 1) {
      if (largeColumns[col] === 0) {
        continue
      }
      columns[writeIndex] = col
      writeIndex += 1
    }
  }
  return columns
}

export function createInitialFormulaValueWriter(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings'>
}): InitialFormulaValueWriter {
  let singleSheetId: number | undefined
  let singleSheetTracker: InitialWrittenColumnTracker | undefined
  let writtenColumnsBySheetId: Map<number, InitialWrittenColumnTracker> | undefined
  const promoteSingleSheetTracker = (): Map<number, InitialWrittenColumnTracker> => {
    writtenColumnsBySheetId = new Map()
    if (singleSheetId !== undefined && singleSheetTracker !== undefined) {
      writtenColumnsBySheetId.set(singleSheetId, singleSheetTracker)
    }
    singleSheetId = undefined
    singleSheetTracker = undefined
    return writtenColumnsBySheetId
  }
  const markKnownColumn = (sheetId: number, col: number): void => {
    if (!writtenColumnsBySheetId && (singleSheetId === undefined || singleSheetId === sheetId)) {
      singleSheetId = sheetId
      singleSheetTracker ??= createInitialWrittenColumnTracker()
      markInitialWrittenColumn(singleSheetTracker, col)
      return
    }
    const trackers = writtenColumnsBySheetId ?? promoteSingleSheetTracker()
    let tracker = trackers.get(sheetId)
    if (!tracker) {
      tracker = createInitialWrittenColumnTracker()
      trackers.set(sheetId, tracker)
    }
    markInitialWrittenColumn(tracker, col)
  }
  const markCellColumn = (cellIndex: number): void => {
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    const col = args.state.workbook.cellStore.cols[cellIndex]
    if (sheetId === undefined || col === undefined) {
      return
    }
    markKnownColumn(sheetId, col)
  }
  const clearDerivedFlags = (cellIndex: number): void => {
    args.state.workbook.cellStore.flags[cellIndex] =
      (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
  }
  const writeNumberCore = (cellIndex: number, value: number): void => {
    const cellStore = args.state.workbook.cellStore
    clearDerivedFlags(cellIndex)
    cellStore.tags[cellIndex] = ValueTag.Number
    cellStore.errors[cellIndex] = ErrorCode.None
    cellStore.stringIds[cellIndex] = 0
    cellStore.numbers[cellIndex] = value
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  }
  const writeValueCore = (cellIndex: number, value: CellValue): void => {
    const cellStore = args.state.workbook.cellStore
    clearDerivedFlags(cellIndex)
    cellStore.tags[cellIndex] = value.tag
    cellStore.errors[cellIndex] = value.tag === ValueTag.Error ? value.code : ErrorCode.None
    cellStore.stringIds[cellIndex] = value.tag === ValueTag.String ? args.state.strings.intern(value.value) : 0
    cellStore.numbers[cellIndex] = value.tag === ValueTag.Number ? value.value : value.tag === ValueTag.Boolean ? (value.value ? 1 : 0) : 0
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  }
  return {
    writeValue(cellIndex, value) {
      writeValueCore(cellIndex, value)
      markCellColumn(cellIndex)
    },
    writeValueAt(cellIndex, sheetId, col, value) {
      writeValueCore(cellIndex, value)
      markKnownColumn(sheetId, col)
    },
    writeNumber(cellIndex, value) {
      writeNumberCore(cellIndex, value)
      markCellColumn(cellIndex)
    },
    writeNumberAt(cellIndex, sheetId, col, value) {
      writeNumberCore(cellIndex, value)
      markKnownColumn(sheetId, col)
    },
    flush() {
      if (!writtenColumnsBySheetId) {
        if (singleSheetId !== undefined && singleSheetTracker !== undefined && singleSheetTracker.count > 0) {
          args.state.workbook.notifyColumnsWritten(singleSheetId, materializeInitialWrittenColumns(singleSheetTracker))
        }
        return
      }
      writtenColumnsBySheetId.forEach((tracker, sheetId) => {
        if (tracker.count > 0) {
          args.state.workbook.notifyColumnsWritten(sheetId, materializeInitialWrittenColumns(tracker))
        }
      })
    },
  }
}

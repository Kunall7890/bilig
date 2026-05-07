import type { U32 } from '../runtime-state.js'
import { formulaColumnCountKey } from './formula-binding-dependency-helpers.js'

export interface PendingInitialFormulaCellTracker {
  readonly withCells: <T>(cellIndices: readonly number[] | U32, callback: () => T) => T
  readonly isFormulaCell: (cellIndex: number) => boolean
  readonly hasColumnMembers: (sheetId: number, col: number) => boolean
}

export function createPendingInitialFormulaCellTracker(args: {
  readonly getCellCapacity: () => number
  readonly getSheetId: (cellIndex: number) => number | undefined
  readonly getCol: (cellIndex: number) => number | undefined
  readonly isBoundFormulaCell: (cellIndex: number) => boolean
  readonly hasBoundColumnMembers: (sheetId: number, col: number) => boolean
}): PendingInitialFormulaCellTracker {
  let pendingCells: Uint8Array | undefined
  let pendingColumnCounts: Map<number, number> | undefined

  const hasPendingCell = (cellIndex: number): boolean => (pendingCells?.[cellIndex] ?? 0) !== 0

  return {
    withCells<T>(cellIndices: readonly number[] | U32, callback: () => T): T {
      const previousCells = pendingCells
      const previousColumnCounts = pendingColumnCounts
      let capacity = args.getCellCapacity() + 1
      for (let index = 0; index < cellIndices.length; index += 1) {
        capacity = Math.max(capacity, (cellIndices[index] ?? 0) + 1)
      }
      if (previousCells) {
        capacity = Math.max(capacity, previousCells.length)
      }

      const nextCells = new Uint8Array(capacity)
      if (previousCells) {
        nextCells.set(previousCells)
      }
      const nextColumnCounts = new Map(previousColumnCounts)
      for (let index = 0; index < cellIndices.length; index += 1) {
        const cellIndex = cellIndices[index]!
        if (nextCells[cellIndex] !== 0) {
          continue
        }
        nextCells[cellIndex] = 1
        const sheetId = args.getSheetId(cellIndex)
        const col = args.getCol(cellIndex)
        if (sheetId === undefined || col === undefined) {
          continue
        }
        const key = formulaColumnCountKey(sheetId, col)
        nextColumnCounts.set(key, (nextColumnCounts.get(key) ?? 0) + 1)
      }

      pendingCells = nextCells
      pendingColumnCounts = nextColumnCounts
      try {
        return callback()
      } finally {
        pendingCells = previousCells
        pendingColumnCounts = previousColumnCounts
      }
    },
    isFormulaCell(cellIndex) {
      return args.isBoundFormulaCell(cellIndex) || hasPendingCell(cellIndex)
    },
    hasColumnMembers(sheetId, col) {
      return args.hasBoundColumnMembers(sheetId, col) || (pendingColumnCounts?.get(formulaColumnCountKey(sheetId, col)) ?? 0) !== 0
    },
  }
}

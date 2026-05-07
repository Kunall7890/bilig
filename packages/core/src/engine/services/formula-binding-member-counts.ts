import { formulaColumnCountKey } from './formula-binding-dependency-helpers.js'

export interface FormulaBindingMemberCounts {
  readonly clear: () => void
  readonly increment: (sheetId: number, col: number | undefined) => void
  readonly decrement: (sheetId: number, col: number | undefined) => void
  readonly countSheetMembers: (sheetId: number) => number
  readonly hasColumnMembers: (sheetId: number, col: number) => boolean
}

function decrementMapCount(map: Map<number, number>, key: number): void {
  const nextCount = (map.get(key) ?? 1) - 1
  if (nextCount <= 0) {
    map.delete(key)
    return
  }
  map.set(key, nextCount)
}

export function createFormulaBindingMemberCounts(): FormulaBindingMemberCounts {
  const columnCounts = new Map<number, number>()
  const sheetCounts = new Map<number, number>()

  return {
    clear() {
      columnCounts.clear()
      sheetCounts.clear()
    },
    increment(sheetId, col) {
      sheetCounts.set(sheetId, (sheetCounts.get(sheetId) ?? 0) + 1)
      if (col === undefined) {
        return
      }
      const columnKey = formulaColumnCountKey(sheetId, col)
      columnCounts.set(columnKey, (columnCounts.get(columnKey) ?? 0) + 1)
    },
    decrement(sheetId, col) {
      decrementMapCount(sheetCounts, sheetId)
      if (col === undefined) {
        return
      }
      decrementMapCount(columnCounts, formulaColumnCountKey(sheetId, col))
    },
    countSheetMembers(sheetId) {
      return sheetCounts.get(sheetId) ?? 0
    },
    hasColumnMembers(sheetId, col) {
      return (columnCounts.get(formulaColumnCountKey(sheetId, col)) ?? 0) !== 0
    },
  }
}

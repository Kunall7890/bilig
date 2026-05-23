import type { EngineRuntimeState } from '../runtime-state.js'

export interface RowVisibilityResolvers {
  readonly isRowHidden: (sheetName: string, rowIndex: number) => boolean
  readonly isRowFiltered: (sheetName: string, rowIndex: number) => boolean
}

export function createRowVisibilityResolvers(workbook: EngineRuntimeState['workbook']): RowVisibilityResolvers {
  const hiddenRowsBySheet = new Map<string, Set<number>>()
  const filteredRowsBySheet = new Map<string, Set<number>>()
  const isRowHidden = (sheetName: string, rowIndex: number): boolean => {
    if (!Number.isInteger(rowIndex) || rowIndex < 0) {
      return false
    }
    let hiddenRows = hiddenRowsBySheet.get(sheetName)
    if (hiddenRows === undefined) {
      hiddenRows = new Set<number>()
      for (const entry of workbook.listRowAxisEntries(sheetName)) {
        if (entry.hidden === true) {
          hiddenRows.add(entry.index)
        }
      }
      for (const record of workbook.listRowMetadata(sheetName)) {
        if (record.hidden !== true) {
          continue
        }
        for (let row = record.start; row < record.start + record.count; row += 1) {
          hiddenRows.add(row)
        }
      }
      hiddenRowsBySheet.set(sheetName, hiddenRows)
    }
    return hiddenRows.has(rowIndex)
  }
  const isRowFiltered = (sheetName: string, rowIndex: number): boolean => {
    if (!Number.isInteger(rowIndex) || rowIndex < 0) {
      return false
    }
    let filteredRows = filteredRowsBySheet.get(sheetName)
    if (filteredRows === undefined) {
      filteredRows = new Set<number>()
      for (const entry of workbook.listRowAxisEntries(sheetName)) {
        if (entry.filterHidden === true) {
          filteredRows.add(entry.index)
        }
      }
      for (const record of workbook.listRowMetadata(sheetName)) {
        if (record.filterHidden !== true) {
          continue
        }
        for (let row = record.start; row < record.start + record.count; row += 1) {
          filteredRows.add(row)
        }
      }
      filteredRowsBySheet.set(sheetName, filteredRows)
    }
    return filteredRows.has(rowIndex)
  }
  return { isRowHidden, isRowFiltered }
}

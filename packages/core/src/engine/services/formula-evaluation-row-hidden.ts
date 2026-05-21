import type { EngineRuntimeState } from '../runtime-state.js'

export function createRowHiddenResolver(workbook: EngineRuntimeState['workbook']): (sheetName: string, rowIndex: number) => boolean {
  const hiddenRowsBySheet = new Map<string, Set<number>>()
  return (sheetName, rowIndex) => {
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
}

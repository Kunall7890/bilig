import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'

export function releaseProjectedCellScanStorage(
  cellScan: ImportedWorksheetCellScan,
  options: {
    readonly releaseArenaAfterMaterialization: boolean | undefined
    readonly useLazyCells: boolean
  },
): void {
  if (options.releaseArenaAfterMaterialization !== true) {
    return
  }
  if (options.useLazyCells) {
    cellScan.arena.releaseMaterializationScratch()
  } else {
    cellScan.arena.release()
  }
  cellScan.styleIndexes.release()
}
